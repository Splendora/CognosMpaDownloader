#requires -Version 5.1
<#
.SYNOPSIS
    Download multiple IBM Cognos Analytics 11 report outputs using the Cognos RDS REST interface.

.DESCRIPTION
    - Stores the Cognos username/password in Windows Credential Manager.
    - Reads report definitions from a JSON configuration file (UTF-8 encoded).
    - Resolves dynamic date tokens (e.g. {Yesterday}, {Today}, {Today-7d}) in parameters and file paths.
    - Parameterizes file names (e.g. {Yesterday:yyyyMMdd}_{ReportName}.xlsx).
    - Authenticates through Cognos Mashup/RDS auth/logon.
    - Retrieves report output using outputFormat.
    - Supports multiple reports, multiple formats, prompt parameters, and output paths.

    First run:
        .\CognosReportDownloader.ps1 -SetupCredential

    Normal run:
        .\CognosReportDownloader.ps1

    Test connection:
        .\CognosReportDownloader.ps1 -TestConnection

    The config file defaults to:
        .\cognos-reports.json
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'cognos-reports.json'),
    [switch]$SetupCredential,
    [switch]$TestConnection
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load System.Net.Http assembly (required for Windows PowerShell 5.1)
Add-Type -AssemblyName System.Net.Http

# -----------------------------------------------------------------------------
# Windows Credential Manager native functions
# -----------------------------------------------------------------------------

if (-not ('CognosCredManager.NativeMethods' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

namespace CognosCredManager
{
    public static class NativeMethods
    {
        /*
         * Native Windows CREDENTIAL structure.
         * FILETIME is represented as Int64 to avoid PtrToStructure layout issues in PS 5.1.
         */
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct CREDENTIAL
        {
            public UInt32 Flags;
            public UInt32 Type;
            public IntPtr TargetName;
            public IntPtr Comment;
            public Int64 LastWritten;
            public UInt32 CredentialBlobSize;
            public IntPtr CredentialBlob;
            public UInt32 Persist;
            public UInt32 AttributeCount;
            public IntPtr Attributes;
            public IntPtr TargetAlias;
            public IntPtr UserName;
        }

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool CredRead(
            string TargetName,
            UInt32 Type,
            UInt32 Flags,
            out IntPtr Credential
        );

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool CredWrite(
            ref CREDENTIAL Credential,
            UInt32 Flags
        );

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern void CredFree(IntPtr Credential);
    }
}
"@
}

function Set-WindowsGenericCredential {
    param(
        [Parameter(Mandatory)]
        [string]$Target,

        [Parameter(Mandatory)]
        [string]$Username,

        [Parameter(Mandatory)]
        [Security.SecureString]$Password
    )

    $bstr = [IntPtr]::Zero
    $plain = $null
    $cred = New-Object CognosCredManager.NativeMethods+CREDENTIAL

    try {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)

        $cred.Flags = 0
        $cred.Type = 1                    # CRED_TYPE_GENERIC
        $cred.Persist = 2                 # CRED_PERSIST_LOCAL_MACHINE

        $cred.TargetName = [Runtime.InteropServices.Marshal]::StringToCoTaskMemUni($Target)
        $cred.UserName   = [Runtime.InteropServices.Marshal]::StringToCoTaskMemUni($Username)

        $blobBytes = [Text.Encoding]::Unicode.GetBytes($plain)
        $cred.CredentialBlobSize = [UInt32]$blobBytes.Length

        if ($blobBytes.Length -gt 0) {
            $cred.CredentialBlob = [Runtime.InteropServices.Marshal]::AllocCoTaskMem($blobBytes.Length)
            [Runtime.InteropServices.Marshal]::Copy($blobBytes, 0, $cred.CredentialBlob, $blobBytes.Length)
        }

        if (-not [CognosCredManager.NativeMethods]::CredWrite([ref]$cred, 0)) {
            $code = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
            throw "CredWrite failed. Windows error code: $code"
        }
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
        if ($cred.TargetName -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeCoTaskMem($cred.TargetName) }
        if ($cred.UserName -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeCoTaskMem($cred.UserName) }
        if ($cred.CredentialBlob -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeCoTaskMem($cred.CredentialBlob) }
        $plain = $null
    }
}

function Get-WindowsGenericCredential {
    param(
        [Parameter(Mandatory)]
        [string]$Target
    )

    $ptr = [IntPtr]::Zero

    try {
        $success = [CognosCredManager.NativeMethods]::CredRead($Target, 1, 0, [ref]$ptr)

        if (-not $success) {
            $code = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
            if ($code -eq 1168) { return $null } # ERROR_NOT_FOUND
            throw "CredRead failed. Windows error code: $code"
        }

        # Explicitly invoke PtrToStructure(IntPtr, Type) to avoid PowerShell 5.1 overload ambiguity
        $method = [Runtime.InteropServices.Marshal].GetMethod('PtrToStructure', [type[]]@([IntPtr], [Type]))
        $native = $method.Invoke($null, @($ptr, [CognosCredManager.NativeMethods+CREDENTIAL]))

        $username = $null
        $password = $null

        if ($native.UserName -ne [IntPtr]::Zero) {
            $username = [Runtime.InteropServices.Marshal]::PtrToStringUni($native.UserName)
        }

        if ($native.CredentialBlob -ne [IntPtr]::Zero -and $native.CredentialBlobSize -gt 0) {
            $blobBytes = New-Object byte[] $native.CredentialBlobSize
            [Runtime.InteropServices.Marshal]::Copy($native.CredentialBlob, $blobBytes, 0, [int]$native.CredentialBlobSize)
            $password = [Text.Encoding]::Unicode.GetString($blobBytes)
        }
        else {
            $password = ''
        }

        return [pscustomobject]@{
            Username = $username
            Password = (ConvertTo-SecureString $password -AsPlainText -Force)
        }
    }
    finally {
        if ($ptr -ne [IntPtr]::Zero) {
            [CognosCredManager.NativeMethods]::CredFree($ptr)
        }
    }
}

function Get-PlainTextFromSecureString {
    param(
        [Parameter(Mandatory)]
        [Security.SecureString]$SecureString
    )

    $bstr = [IntPtr]::Zero
    try {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

# -----------------------------------------------------------------------------
# Dynamic Token & Date Resolution
# -----------------------------------------------------------------------------

function Resolve-DynamicTokens {
    param(
        [Parameter(Mandatory)]
        [string]$Text,

        $Report = $null,

        [string]$Format = ''
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }

    $now = Get-Date
    $resolved = $Text

    # 1. Base relative dates
    $today = $now.Date
    $yesterday = $today.AddDays(-1)
    $monthStart = [DateTime]::new($today.Year, $today.Month, 1)
    $monthEnd = $monthStart.AddMonths(1).AddDays(-1)

    # 2. Named Tokens (Default ISO format: yyyy-MM-dd)
    $resolved = $resolved.Replace('{Yesterday}', $yesterday.ToString('yyyy-MM-dd'))
    $resolved = $resolved.Replace('{Today}', $today.ToString('yyyy-MM-dd'))
    $resolved = $resolved.Replace('{MonthStart}', $monthStart.ToString('yyyy-MM-dd'))
    $resolved = $resolved.Replace('{MonthEnd}', $monthEnd.ToString('yyyy-MM-dd'))

    # 3. Custom Date Format Tokens (e.g. {Yesterday:yyyyMMdd}, {Today:yyyyMMdd_HHmmss})
    $resolved = [regex]::Replace($resolved, '\{(Yesterday|Today)(:([^}]+))?\}', {
        param($match)
        $baseDate = if ($match.Groups[1].Value -eq 'Yesterday') { $yesterday } else { $today }
        $fmt = if ($match.Groups[3].Success) { $match.Groups[3].Value } else { 'yyyy-MM-dd' }
        return $baseDate.ToString($fmt)
    })

    # 4. Offset tokens (e.g. {Today-7d}, {Today-1d:yyyyMMdd})
    $resolved = [regex]::Replace($resolved, '\{Today([+-]\d+)d?(:([^}]+))?\}', {
        param($match)
        $days = [int]$match.Groups[1].Value
        $targetDate = $today.AddDays($days)
        $fmt = if ($match.Groups[3].Success) { $match.Groups[3].Value } else { 'yyyy-MM-dd' }
        return $targetDate.ToString($fmt)
    })

    # 5. Standard Date/Time Tokens
    $resolved = $resolved.Replace('{yyyy}', $now.ToString('yyyy'))
    $resolved = $resolved.Replace('{MM}', $now.ToString('MM'))
    $resolved = $resolved.Replace('{dd}', $now.ToString('dd'))
    $resolved = $resolved.Replace('{yyyyMMdd}', $now.ToString('yyyyMMdd'))
    $resolved = $resolved.Replace('{yyyy-MM-dd}', $now.ToString('yyyy-MM-dd'))
    $resolved = $resolved.Replace('{HHmmss}', $now.ToString('HHmmss'))

    # 6. Report Metadata & Prompt Parameter Tokens (if Report object provided)
    if ($null -ne $Report) {
        $reportName = if ($Report.PSObject.Properties['Name'] -and $Report.Name) { [string]$Report.Name } else { [string]$Report.Source }
        $resolved = $resolved.Replace('{ReportName}', $reportName)
        $resolved = $resolved.Replace('{Source}', [string]$Report.Source)
        if (-not [string]::IsNullOrWhiteSpace($Format)) {
            $resolved = $resolved.Replace('{Format}', $Format)
        }

        if ($Report.PSObject.Properties['Parameters'] -and $null -ne $Report.Parameters) {
            foreach ($prop in $Report.Parameters.PSObject.Properties) {
                $rawVal = if ($null -ne $prop.Value) { [string]$prop.Value } else { '' }
                $evaluatedVal = Resolve-DynamicTokens -Text $rawVal
                $cleanVal = $evaluatedVal -replace '[\\/:*?"<>|]', '-'
                $resolved = $resolved.Replace("{$($prop.Name)}", $cleanVal)
                if ($prop.Name.StartsWith('p_')) {
                    $resolved = $resolved.Replace("{$($prop.Name.Substring(2))}", $cleanVal)
                }
            }
        }
    }

    return $resolved
}

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $prefix = switch ($Level) {
        'OK'    { '[OK]   ' }
        'WARN'  { '[WARN] ' }
        'ERROR' { '[ERROR]' }
        default { '[INFO] ' }
    }

    Write-Host "$prefix $Message"
}

function Get-RequiredProperty {
    param(
        [Parameter(Mandatory)]
        $Object,

        [Parameter(Mandatory)]
        [string]$Name
    )

    $prop = $Object.PSObject.Properties[$Name]

    if ($null -eq $prop -or [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
        throw "Required configuration property '$Name' is missing."
    }

    return [string]$prop.Value
}

function Add-QueryParameter {
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[string]]$Parts,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        $Value
    )

    if ($Value -is [System.Array] -and -not ($Value -is [string])) {
        foreach ($item in $Value) {
            $Parts.Add("$([uri]::EscapeDataString($Name))=$([uri]::EscapeDataString([string]$item))")
        }
    }
    else {
        $Parts.Add("$([uri]::EscapeDataString($Name))=$([uri]::EscapeDataString([string]$Value))")
    }
}

function Get-CookieValue {
    param(
        [Parameter(Mandatory)]
        [System.Net.CookieContainer]$Container,

        [Parameter(Mandatory)]
        [uri]$Uri,

        [Parameter(Mandatory)]
        [string]$Name
    )

    foreach ($cookie in $Container.GetCookies($Uri)) {
        if ($cookie.Name -ieq $Name) {
            return $cookie.Value
        }
    }
    return $null
}

function New-CognosHttpClient {
    $cookieContainer = New-Object System.Net.CookieContainer

    $handler = New-Object System.Net.Http.HttpClientHandler -Property @{
        CookieContainer   = $cookieContainer
        AllowAutoRedirect = $false
        UseCookies        = $true
    }

    $client = New-Object System.Net.Http.HttpClient($handler)
    $client.Timeout = [TimeSpan]::FromMinutes(30)
    $client.DefaultRequestHeaders.Accept.Clear()
    $client.DefaultRequestHeaders.Accept.Add(
        [System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new('*/*')
    )

    return [pscustomobject]@{
        Client  = $client
        Handler = $handler
        Cookies = $cookieContainer
    }
}

function Invoke-CognosRequest {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] [ValidateSet('GET', 'POST')] [string]$Method,
        [Parameter(Mandatory)] [uri]$Uri,
        [hashtable]$Headers,
        [System.Net.Http.HttpContent]$Content,
        [int]$MaxRedirects = 10
    )

    $current = $Uri
    $redirectCount = 0

    while ($true) {
        $httpMethod = if ($Method -eq 'GET') { [System.Net.Http.HttpMethod]::Get } else { [System.Net.Http.HttpMethod]::Post }
        $request = New-Object System.Net.Http.HttpRequestMessage($httpMethod, $current)

        if ($Headers) {
            foreach ($key in $Headers.Keys) {
                [void]$request.Headers.TryAddWithoutValidation($key, [string]$Headers[$key])
            }
        }

        if ($null -ne $Content) {
            $request.Content = $Content
        }

        $response = $Context.Client.SendAsync($request).GetAwaiter().GetResult()
        $statusCode = [int]$response.StatusCode

        # Check HTTP 300..308 redirect codes directly
        if ($statusCode -in @(300, 301, 302, 303, 307, 308)) {
            if ($redirectCount -ge $MaxRedirects) {
                throw "Too many redirects while requesting $current"
            }

            $location = $response.Headers.Location
            if ($null -eq $location) {
                throw "Cognos returned HTTP $statusCode without a Location header."
            }

            $current = if ($location.IsAbsoluteUri) { $location } else { [uri]::new($current, $location) }
            $redirectCount++
            $response.Dispose()
            continue
        }

        return [pscustomobject]@{
            Response = $response
            FinalUri = $current
        }
    }
}

# -----------------------------------------------------------------------------
# Cognos Authentication
# -----------------------------------------------------------------------------

function Invoke-CognosLogin {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] [string]$CognosBaseUrl,
        [Parameter(Mandatory)] [string]$Namespace,
        [Parameter(Mandatory)] [string]$Username,
        [Parameter(Mandatory)] [Security.SecureString]$Password
    )

    $plainPassword = Get-PlainTextFromSecureString $Password

    try {
        $xml = @"
<credentials xmlns="http://developer.cognos.com/schemas/ccs/auth/types/1">
  <credentialElements>
    <name>CAMNamespace</name>
    <label>Namespace:</label>
    <value>
      <actualValue>$([System.Security.SecurityElement]::Escape($Namespace))</actualValue>
    </value>
  </credentialElements>
  <credentialElements>
    <name>CAMUsername</name>
    <label>User ID:</label>
    <value>
      <actualValue>$([System.Security.SecurityElement]::Escape($Username))</actualValue>
    </value>
  </credentialElements>
  <credentialElements>
    <name>CAMPassword</name>
    <label>Password:</label>
    <value>
      <actualValue>$([System.Security.SecurityElement]::Escape($plainPassword))</actualValue>
    </value>
  </credentialElements>
</credentials>
"@

        $dict = [System.Collections.Generic.Dictionary[string, string]]::new()
        $dict['xmlData'] = $xml
        $form = New-Object System.Net.Http.FormUrlEncodedContent($dict)

        $url = [uri]::new(($CognosBaseUrl.TrimEnd('/') + '/v1/disp/rds/auth/logon'))
        $result = Invoke-CognosRequest -Context $Context -Method POST -Uri $url -Content $form
        $body = $result.Response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $status = [int]$result.Response.StatusCode
        $result.Response.Dispose()

        if ($status -ne 200) {
            throw "Cognos logon failed with HTTP $status. Server response: $body"
        }

        if ($body -notmatch 'accountInfo') {
            throw "Cognos did not return accountInfo. Authentication was not successful. Server response: $body"
        }

        $xsrf = Get-CookieValue -Container $Context.Cookies -Uri $url -Name 'XSRF-TOKEN'

        if ([string]::IsNullOrWhiteSpace($xsrf)) {
            $probeUrl = [uri]::new(($CognosBaseUrl.TrimEnd('/') + '/v1/disp/rds/auth/logon?xmlData=%3Ccredentials%2F%3E'))
            $probe = Invoke-CognosRequest -Context $Context -Method GET -Uri $probeUrl
            $probe.Response.Dispose()
            $xsrf = Get-CookieValue -Container $Context.Cookies -Uri $probeUrl -Name 'XSRF-TOKEN'
        }

        if ([string]::IsNullOrWhiteSpace($xsrf)) {
            throw 'Cognos authenticated successfully but no XSRF-TOKEN cookie was received.'
        }

        return $xsrf
    }
    finally {
        $plainPassword = $null
    }
}

# -----------------------------------------------------------------------------
# Report URL Construction
# -----------------------------------------------------------------------------

function Get-ReportDefinitionUrl {
    param(
        [Parameter(Mandatory)] [string]$CognosBaseUrl,
        [Parameter(Mandatory)] $Report,
        [Parameter(Mandatory)] [string]$Format
    )

    $sourceType = if ($Report.PSObject.Properties['SourceType'] -and $Report.SourceType) {
        [string]$Report.SourceType
    } else {
        'report'
    }

    $source = Get-RequiredProperty -Object $Report -Name 'Source'
    $encodedSource = [uri]::EscapeDataString($source)
    $encodedFormat = [uri]::EscapeDataString($Format)

    $parts = New-Object 'System.Collections.Generic.List[string]'
    $parts.Add('v=3')

    # Resolve dynamic tokens (e.g. {Yesterday}) in prompt parameter values
    if ($Report.PSObject.Properties['Parameters'] -and $null -ne $Report.Parameters) {
        foreach ($property in $Report.Parameters.PSObject.Properties) {
            $rawVal = if ($null -ne $property.Value) { [string]$property.Value } else { '' }
            if (-not [string]::IsNullOrWhiteSpace($rawVal)) {
                $evaluatedVal = Resolve-DynamicTokens -Text $rawVal
                Add-QueryParameter -Parts $parts -Name $property.Name -Value $evaluatedVal
            }
        }
    }

    if ($Report.PSObject.Properties['Options'] -and $null -ne $Report.Options) {
        foreach ($property in $Report.Options.PSObject.Properties) {
            Add-QueryParameter -Parts $parts -Name $property.Name -Value $property.Value
        }
    }

    return (
        "$($CognosBaseUrl.TrimEnd('/'))" +
        "/v1/disp/rds/outputFormat/" +
        "$sourceType/" +
        "$encodedSource/" +
        "${encodedFormat}?" +
        ($parts -join '&')
    )
}

# -----------------------------------------------------------------------------
# Load Configuration
# -----------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    if ($SetupCredential) {
        throw "Config file not found: $ConfigPath. Create the config first."
    }
    throw "Config file not found: $ConfigPath. Run Manage-CognosConfig.ps1 to create it."
}

$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

$cognosBaseUrl = Get-RequiredProperty -Object $config -Name 'CognosBaseUrl'
$namespace     = Get-RequiredProperty -Object $config -Name 'Namespace'

$credentialTarget = if ($config.PSObject.Properties['CredentialTarget'] -and $config.CredentialTarget) {
    [string]$config.CredentialTarget
} else {
    'CognosReportAutomation:' + ([uri]$cognosBaseUrl).Host
}

# -----------------------------------------------------------------------------
# Credential Setup Mode
# -----------------------------------------------------------------------------

if ($SetupCredential) {
    $existing = Get-WindowsGenericCredential -Target $credentialTarget
    if ($existing) {
        Write-Log "A credential already exists for '$credentialTarget'. It will be replaced." 'WARN'
    }

    $credential = Get-Credential -Message "Enter the Cognos account to save in Windows Credential Manager"
    if ($null -eq $credential) { throw 'Credential setup was cancelled.' }

    Set-WindowsGenericCredential -Target $credentialTarget -Username $credential.UserName -Password $credential.Password
    Write-Log "Credential saved to Windows Credential Manager as '$credentialTarget'." 'OK'
    exit 0
}

# -----------------------------------------------------------------------------
# Retrieve Saved Credential
# -----------------------------------------------------------------------------

$stored = Get-WindowsGenericCredential -Target $credentialTarget

if ($null -eq $stored) {
    Write-Log "No saved credential found for '$credentialTarget'." 'WARN'
    $credential = Get-Credential -Message "Enter the Cognos account (it will be saved in Windows Credential Manager)"
    if ($null -eq $credential) { throw 'Credential entry was cancelled.' }

    Set-WindowsGenericCredential -Target $credentialTarget -Username $credential.UserName -Password $credential.Password
    $stored = [pscustomobject]@{ Username = $credential.UserName; Password = $credential.Password }
    Write-Log 'Credential saved.' 'OK'
}

# -----------------------------------------------------------------------------
# Connect / Authenticate
# -----------------------------------------------------------------------------

Write-Log "Connecting to Cognos at $cognosBaseUrl"
$context = New-CognosHttpClient

try {
    $xsrf = Invoke-CognosLogin `
        -Context $context `
        -CognosBaseUrl $cognosBaseUrl `
        -Namespace $namespace `
        -Username $stored.Username `
        -Password $stored.Password

    Write-Log "Authenticated as $($stored.Username) in namespace '$namespace'." 'OK'

    if ($TestConnection) {
        Write-Log 'Cognos connection/authentication test completed successfully.' 'OK'
        exit 0
    }

    # -------------------------------------------------------------------------
    # Validate Reports
    # -------------------------------------------------------------------------

    $reportsProp = $config.PSObject.Properties['Reports']
    if ($null -eq $reportsProp -or $null -eq $reportsProp.Value -or @($reportsProp.Value).Count -eq 0) {
        throw 'No reports are defined in the configuration file.'
    }

    $total = 0
    $success = 0
    $failed = 0

    # -------------------------------------------------------------------------
    # Download Reports
    # -------------------------------------------------------------------------

    foreach ($report in @($config.Reports)) {
        $reportName = if ($report.PSObject.Properties['Name'] -and $report.Name) { [string]$report.Name } else { [string]$report.Source }

        if ($report.PSObject.Properties['Enabled'] -and $report.Enabled -eq $false) {
            Write-Log "Skipping disabled report: $reportName" 'WARN'
            continue
        }

        $formatsProp = $report.PSObject.Properties['Formats']
        if ($null -eq $formatsProp -or $null -eq $formatsProp.Value -or @($formatsProp.Value).Count -eq 0) {
            Write-Log "Report '$reportName' has no formats configured." 'WARN'
            continue
        }

        foreach ($formatConfig in @($report.Formats)) {
            $total++

            $format = Get-RequiredProperty -Object $formatConfig -Name 'Format'
            $rawOutputPath = Get-RequiredProperty -Object $formatConfig -Name 'OutputPath'

            # -----------------------------------------------------------------
            # Resolve Dynamic Output Path (e.g. {Yesterday:yyyyMMdd}_{ReportName})
            # -----------------------------------------------------------------
            $outputPath = Resolve-DynamicTokens -Text $rawOutputPath -Report $report -Format $format

            try {
                # Create output directory if it doesn't exist
                $parent = Split-Path -Parent $outputPath
                if ($parent -and -not (Test-Path -LiteralPath $parent)) {
                    New-Item -ItemType Directory -Path $parent -Force | Out-Null
                }

                # Build Cognos REST URL
                $url = Get-ReportDefinitionUrl -CognosBaseUrl $cognosBaseUrl -Report $report -Format $format

                Write-Log "Downloading '$reportName' as $format..."

                # Request report execution
                $headers = @{ 'X-XSRF-TOKEN' = $xsrf }
                $result = Invoke-CognosRequest -Context $context -Method GET -Uri ([uri]$url) -Headers $headers
                $response = $result.Response
                $status = [int]$response.StatusCode

                if ($status -ne 200) {
                    $body = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                    $response.Dispose()
                    throw "Cognos returned HTTP $status. $body"
                }

                # Read binary output
                $bytes = $response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
                $contentType = if ($response.Content.Headers.ContentType) { $response.Content.Headers.ContentType.MediaType } else { '' }
                $response.Dispose()

                # Guard against HTML/XML error pages saved as reports
                if ($bytes.Length -lt 1000 -and $contentType -match 'text|xml|html') {
                    $text = [Text.Encoding]::UTF8.GetString($bytes)
                    if ($text -match 'RDS-ERR' -or $text -match '<rds:error') {
                        throw "Cognos returned an error response instead of report output: $text"
                    }
                }

                # Write output to disk
                [IO.File]::WriteAllBytes($outputPath, $bytes)
                $sizeMb = [Math]::Round($bytes.Length / 1MB, 2)
                Write-Log "Saved $outputPath ($sizeMb MB)" 'OK'
                $success++
            }
            catch {
                $failed++
                Write-Log "Failed '$reportName' / $format : $($_.Exception.Message)" 'ERROR'
            }
        }
    }

    # -------------------------------------------------------------------------
    # Summary
    # -------------------------------------------------------------------------

    Write-Host ''
    Write-Log "Download complete."
    Write-Log "Total: $total"
    Write-Log "Successful: $success" 'OK'

    if ($failed -gt 0) {
        Write-Log "Failed: $failed" 'ERROR'
    } else {
        Write-Log "Failed: 0" 'OK'
    }
}
finally {
    if ($null -ne $context) {
        if ($null -ne $context.Client)  { $context.Client.Dispose() }
        if ($null -ne $context.Handler) { $context.Handler.Dispose() }
    }
}