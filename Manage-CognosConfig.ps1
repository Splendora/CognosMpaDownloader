#requires -Version 5.1
<#
.SYNOPSIS
    Interactive one-stop-shop manager to create, view, add, edit, and sync cognos-reports.json.
    Includes smart date parameter detection, {Yesterday} presets, and live path evaluation.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'cognos-reports.json')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load System.Net.Http assembly (required for Windows PowerShell 5.1)
Add-Type -AssemblyName System.Net.Http

# -----------------------------------------------------------------------------
# Dynamic Token & Date Resolution (for live previewing)
# -----------------------------------------------------------------------------

function Resolve-DynamicTokens {
    param(
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Text = '',

        $Report = $null,

        [AllowEmptyString()]
        [AllowNull()]
        [string]$Format = ''
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }

    $now = Get-Date
    $resolved = $Text

    $today = $now.Date
    $yesterday = $today.AddDays(-1)
    $monthStart = [DateTime]::new($today.Year, $today.Month, 1)
    $monthEnd = $monthStart.AddMonths(1).AddDays(-1)

    $resolved = $resolved.Replace('{Yesterday}', $yesterday.ToString('yyyy-MM-dd'))
    $resolved = $resolved.Replace('{Today}', $today.ToString('yyyy-MM-dd'))
    $resolved = $resolved.Replace('{MonthStart}', $monthStart.ToString('yyyy-MM-dd'))
    $resolved = $resolved.Replace('{MonthEnd}', $monthEnd.ToString('yyyy-MM-dd'))

    $resolved = [regex]::Replace($resolved, '\{(Yesterday|Today)(:([^}]+))?\}', {
        param($match)
        $baseDate = if ($match.Groups[1].Value -eq 'Yesterday') { $yesterday } else { $today }
        $fmt = if ($match.Groups[3].Success) { $match.Groups[3].Value } else { 'yyyy-MM-dd' }
        return $baseDate.ToString($fmt)
    })

    $resolved = [regex]::Replace($resolved, '\{Today([+-]\d+)d?(:([^}]+))?\}', {
        param($match)
        $days = [int]$match.Groups[1].Value
        $targetDate = $today.AddDays($days)
        $fmt = if ($match.Groups[3].Success) { $match.Groups[3].Value } else { 'yyyy-MM-dd' }
        return $targetDate.ToString($fmt)
    })

    $resolved = $resolved.Replace('{yyyy}', $now.ToString('yyyy'))
    $resolved = $resolved.Replace('{MM}', $now.ToString('MM'))
    $resolved = $resolved.Replace('{dd}', $now.ToString('dd'))
    $resolved = $resolved.Replace('{yyyyMMdd}', $now.ToString('yyyyMMdd'))
    $resolved = $resolved.Replace('{yyyy-MM-dd}', $now.ToString('yyyy-MM-dd'))
    $resolved = $resolved.Replace('{HHmmss}', $now.ToString('HHmmss'))

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
# Windows Credential Manager Functions
# -----------------------------------------------------------------------------

if (-not ('CognosCredManager.NativeMethods' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

namespace CognosCredManager
{
    public static class NativeMethods
    {
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
        public static extern bool CredRead(string TargetName, UInt32 Type, UInt32 Flags, out IntPtr Credential);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool CredWrite(ref CREDENTIAL Credential, UInt32 Flags);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern void CredFree(IntPtr Credential);
    }
}
"@
}

function Set-StoredCredential {
    param([string]$Target, [string]$Username, [Security.SecureString]$Password)
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    $cred = New-Object CognosCredManager.NativeMethods+CREDENTIAL

    try {
        $cred.Flags = 0
        $cred.Type = 1
        $cred.Persist = 2
        $cred.TargetName = [Runtime.InteropServices.Marshal]::StringToCoTaskMemUni($Target)
        $cred.UserName = [Runtime.InteropServices.Marshal]::StringToCoTaskMemUni($Username)
        $blobBytes = [Text.Encoding]::Unicode.GetBytes($plain)
        $cred.CredentialBlobSize = [UInt32]$blobBytes.Length
        if ($blobBytes.Length -gt 0) {
            $cred.CredentialBlob = [Runtime.InteropServices.Marshal]::AllocCoTaskMem($blobBytes.Length)
            [Runtime.InteropServices.Marshal]::Copy($blobBytes, 0, $cred.CredentialBlob, $blobBytes.Length)
        }
        if (-not [CognosCredManager.NativeMethods]::CredWrite([ref]$cred, 0)) {
            throw "CredWrite failed: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())"
        }
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
        if ($cred.TargetName -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeCoTaskMem($cred.TargetName) }
        if ($cred.UserName -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeCoTaskMem($cred.UserName) }
        if ($cred.CredentialBlob -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeCoTaskMem($cred.CredentialBlob) }
    }
}

function Get-StoredCredential {
    param([string]$Target)
    $ptr = [IntPtr]::Zero
    try {
        $success = [CognosCredManager.NativeMethods]::CredRead($Target, 1, 0, [ref]$ptr)
        if (-not $success) { return $null }

        $method = [Runtime.InteropServices.Marshal].GetMethod('PtrToStructure', [type[]]@([IntPtr], [Type]))
        $native = $method.Invoke($null, @($ptr, [CognosCredManager.NativeMethods+CREDENTIAL]))

        $username = if ($native.UserName -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::PtrToStringUni($native.UserName) } else { '' }
        $password = ''
        if ($native.CredentialBlob -ne [IntPtr]::Zero -and $native.CredentialBlobSize -gt 0) {
            $blobBytes = New-Object byte[] $native.CredentialBlobSize
            [Runtime.InteropServices.Marshal]::Copy($native.CredentialBlob, $blobBytes, 0, [int]$native.CredentialBlobSize)
            $password = [Text.Encoding]::Unicode.GetString($blobBytes)
        }
        return [pscustomobject]@{ Username = $username; Password = (ConvertTo-SecureString $password -AsPlainText -Force) }
    }
    finally {
        if ($ptr -ne [IntPtr]::Zero) { [CognosCredManager.NativeMethods]::CredFree($ptr) }
    }
}

function Get-PlainText {
    param([Security.SecureString]$SecureString)
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try { [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

# -----------------------------------------------------------------------------
# HTTP & Login Functions
# -----------------------------------------------------------------------------

function New-HttpClient {
    $cookies = New-Object System.Net.CookieContainer
    $handler = New-Object System.Net.Http.HttpClientHandler -Property @{
        CookieContainer   = $cookies
        AllowAutoRedirect = $false
        UseCookies        = $true
    }
    $client = New-Object System.Net.Http.HttpClient($handler)
    $client.Timeout = [TimeSpan]::FromMinutes(5)
    return [pscustomobject]@{ Client = $client; Cookies = $cookies }
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
        if ($null -ne $Content) { $request.Content = $Content }

        $response = $Context.Client.SendAsync($request).GetAwaiter().GetResult()
        $statusCode = [int]$response.StatusCode

        if ($statusCode -in @(300, 301, 302, 303, 307, 308)) {
            if ($redirectCount -ge $MaxRedirects) { throw "Too many redirects requesting $current" }
            $location = $response.Headers.Location
            if ($null -ne $location) {
                $current = if ($location.IsAbsoluteUri) { $location } else { [uri]::new($current, $location) }
                $redirectCount++
                $response.Dispose()
                continue
            }
        }
        return [pscustomobject]@{ Response = $response; FinalUri = $current }
    }
}

function Invoke-Login {
    param($Context, [string]$BaseUrl, [string]$Namespace, [string]$Username, [Security.SecureString]$Password)
    $plain = Get-PlainText $Password
    $xml = @"
<credentials xmlns="http://developer.cognos.com/schemas/ccs/auth/types/1">
  <credentialElements><name>CAMNamespace</name><label>Namespace:</label><value><actualValue>$([System.Security.SecurityElement]::Escape($Namespace))</actualValue></value></credentialElements>
  <credentialElements><name>CAMUsername</name><label>User ID:</label><value><actualValue>$([System.Security.SecurityElement]::Escape($Username))</actualValue></value></credentialElements>
  <credentialElements><name>CAMPassword</name><label>Password:</label><value><actualValue>$([System.Security.SecurityElement]::Escape($plain))</actualValue></value></credentialElements>
</credentials>
"@
    $dict = [System.Collections.Generic.Dictionary[string, string]]::new()
    $dict['xmlData'] = $xml
    $form = New-Object System.Net.Http.FormUrlEncodedContent($dict)
    $url = [uri]::new(($BaseUrl.TrimEnd('/') + '/v1/disp/rds/auth/logon'))

    $result = Invoke-CognosRequest -Context $Context -Method POST -Uri $url -Content $form
    $body = $result.Response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
    $result.Response.Dispose()

    if ($body -notmatch 'accountInfo') {
        throw "Logon failed. Server response: $body"
    }

    foreach ($cookie in $Context.Cookies.GetCookies($url)) {
        if ($cookie.Name -ieq 'XSRF-TOKEN') { return $cookie.Value }
    }
    throw 'Login succeeded but XSRF-TOKEN cookie was not found.'
}

function Get-ReportParameters {
    param($Context, [string]$BaseUrl, [string]$SourceType, [string]$Source, [string]$Xsrf)

    $endpointType = if ($SourceType -ieq 'path') { 'path' } else { 'report' }
    $encodedSource = [uri]::EscapeDataString($Source)
    $url = [uri]::new(($BaseUrl.TrimEnd('/') + "/v1/disp/rds/reportPrompts/${endpointType}/${encodedSource}?v=3"))

    $headers = @{ 'X-XSRF-TOKEN' = $Xsrf }
    $result = Invoke-CognosRequest -Context $Context -Method GET -Uri $url -Headers $headers
    $xmlText = $result.Response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
    $currentUri = $result.FinalUri
    $result.Response.Dispose()

    $pollCount = 0
    while ($xmlText -match '<rds:url>(.*?)</rds:url>' -and $pollCount -lt 10) {
        $relUrl = $matches[1].Replace('&amp;', '&')
        $sessionUrl = [uri]::new($currentUri, $relUrl)
        Start-Sleep -Milliseconds 500
        $sessionResult = Invoke-CognosRequest -Context $Context -Method GET -Uri $sessionUrl -Headers $headers
        $xmlText = $sessionResult.Response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $currentUri = $sessionResult.FinalUri
        $sessionResult.Response.Dispose()
        $pollCount++
    }

    $params = [ordered]@{}
    try {
        [xml]$doc = $xmlText
        $pnodes = $doc.SelectNodes("//*[local-name()='pname' or local-name()='parameter']")
        if ($null -ne $pnodes) {
            foreach ($node in $pnodes) {
                $rawName = $node.InnerText.Trim()
                if (-not [string]::IsNullOrWhiteSpace($rawName)) {
                    $paramKey = if ($rawName.StartsWith('p_')) { $rawName } else { "p_$rawName" }
                    $params[$paramKey] = ""
                }
            }
        }
    }
    catch {
        Write-Warning "Error parsing prompt XML for '$Source': $($_.Exception.Message)"
    }
    return $params
}

# -----------------------------------------------------------------------------
# Configuration File Management
# -----------------------------------------------------------------------------

function Load-Config {
    if (-not (Test-Path -LiteralPath $ConfigPath)) { return $null }
    $raw = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8
    return ($raw | ConvertFrom-Json)
}

function Save-Config {
    param($Config)
    if (Test-Path -LiteralPath $ConfigPath) {
        Copy-Item -LiteralPath $ConfigPath -Destination "$ConfigPath.bak" -Force
    }
    $jsonText = $Config | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($ConfigPath, $jsonText, [System.Text.Encoding]::UTF8)
    Write-Host "[OK] Configuration saved to: $ConfigPath" -ForegroundColor Green
}

function Connect-CognosSession {
    param($Config)
    $credTarget = if ($Config.PSObject.Properties['CredentialTarget'] -and $Config.CredentialTarget) {
        [string]$Config.CredentialTarget
    } else {
        'CognosReportAutomation:' + ([uri]$Config.CognosBaseUrl).Host
    }

    $cred = Get-StoredCredential -Target $credTarget
    if ($null -eq $cred) {
        Write-Host "`n[AUTH] No saved credentials found for $credTarget. Please enter them:" -ForegroundColor Yellow
        $userCred = Get-Credential -Message "Cognos Credentials"
        if ($null -eq $userCred) { throw "Credential input cancelled." }
        Set-StoredCredential -Target $credTarget -Username $userCred.UserName -Password $userCred.Password
        $cred = [pscustomobject]@{ Username = $userCred.UserName; Password = $userCred.Password }
    }

    Write-Host "Connecting to Cognos at $($Config.CognosBaseUrl)..." -ForegroundColor Cyan
    $http = New-HttpClient
    $xsrf = Invoke-Login -Context $http -BaseUrl $Config.CognosBaseUrl -Namespace $Config.Namespace -Username $cred.Username -Password $cred.Password
    Write-Host "[OK] Authenticated as $($cred.Username)." -ForegroundColor Green

    return [pscustomobject]@{ Http = $http; Xsrf = $xsrf }
}

# Helper to intelligently prompt for parameter value (with dynamic date suggestions)
function Prompt-ForParameterValue {
    param([string]$ParamName, [AllowEmptyString()][string]$CurrentVal = '')

    $isDateParam = ($ParamName -match 'date|ngay|time|denhan|quahan')
    
    if ($isDateParam) {
        $defaultOption = if (-not [string]::IsNullOrWhiteSpace($CurrentVal)) { $CurrentVal } else { '{Yesterday}' }
        Write-Host "`n  Parameter '$ParamName' looks like a Date/Time parameter:" -ForegroundColor Cyan
        Write-Host "    [1] {Yesterday}   (Evaluates to: $((Get-Date).AddDays(-1).ToString('yyyy-MM-dd')))"
        Write-Host "    [2] {Today}       (Evaluates to: $((Get-Date).ToString('yyyy-MM-dd')))"
        Write-Host "    [3] {MonthStart}  (First day of month)"
        Write-Host "    [4] Custom Value"
        Write-Host "    [5] Leave Empty"
        $choice = Read-Host "  Select preset [1-5, Default: $defaultOption]"
        
        switch ($choice) {
            '1' { return '{Yesterday}' }
            '2' { return '{Today}' }
            '3' { return '{MonthStart}' }
            '4' { return (Read-Host "  Enter custom value") }
            '5' { return '' }
            default { return $defaultOption }
        }
    } else {
        $promptMsg = "  $ParamName" + (if (-not [string]::IsNullOrWhiteSpace($CurrentVal)) { " [Current: '$CurrentVal']" } else { " (press Enter to leave empty)" })
        $val = Read-Host $promptMsg
        if ([string]::IsNullOrWhiteSpace($val)) { return $CurrentVal }
        return $val
    }
}

# -----------------------------------------------------------------------------
# Interactive Actions
# -----------------------------------------------------------------------------

function Action-InitConfig {
    Write-Host "`n=== Initial Server Configuration ===" -ForegroundColor Cyan
    $url = Read-Host "Enter Cognos Base URL (e.g. http://cognos-server:9300/bi)"
    $ns  = Read-Host "Enter Cognos Namespace (e.g. LDAP or ActiveDirectory)"
    $target = "CognosReportAutomation:" + ([uri]$url).Host

    Write-Host "Enter Cognos login credentials to store securely in Windows Credential Manager:" -ForegroundColor Yellow
    $userCred = Get-Credential -Message "Cognos Account"
    if ($null -ne $userCred) {
        Set-StoredCredential -Target $target -Username $userCred.UserName -Password $userCred.Password
    }

    $newConfig = [ordered]@{
        CognosBaseUrl    = $url.TrimEnd('/')
        Namespace        = $ns.Trim()
        CredentialTarget = $target
        Reports          = @()
    }
    Save-Config -Config $newConfig
    return $newConfig
}

function Action-ListReports {
    param($Config)
    $reports = if ($Config.PSObject.Properties['Reports'] -and $Config.Reports) { @($Config.Reports) } else { @() }
    if (@($reports).Count -eq 0) {
        Write-Host "`nNo reports configured yet." -ForegroundColor Yellow
        return
    }

    Write-Host "`n=== Configured Reports ($(@($reports).Count)) ===" -ForegroundColor Cyan
    for ($i = 0; $i -lt @($reports).Count; $i++) {
        $rep = $reports[$i]
        $enabledStr = if ($rep.PSObject.Properties['Enabled'] -and $rep.Enabled -eq $false) { "[DISABLED]" } else { "[ACTIVE]" }
        $color = if ($enabledStr -eq '[ACTIVE]') { 'Green' } else { 'DarkGray' }
        Write-Host "[$($i + 1)] $enabledStr $($rep.Name) (Source: $($rep.Source))" -ForegroundColor $color

        # Display Parameters with live evaluation preview
        if ($rep.PSObject.Properties['Parameters'] -and $null -ne $rep.Parameters) {
            $paramProps = @($rep.Parameters.PSObject.Properties)
            if ($paramProps.Count -gt 0) {
                Write-Host "    Parameters:" -ForegroundColor Gray
                foreach ($p in $paramProps) {
                    $rawVal = if ([string]::IsNullOrWhiteSpace([string]$p.Value)) { "<EMPTY>" } else { [string]$p.Value }
                    $preview = if ($rawVal -ne '<EMPTY>') { Resolve-DynamicTokens -Text $rawVal } else { '' }
                    
                    if ($preview -and $preview -ne $rawVal) {
                        Write-Host "      - $($p.Name) = $rawVal -> (Evaluates to: $preview)" -ForegroundColor Yellow
                    } else {
                        Write-Host "      - $($p.Name) = $rawVal" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "    Parameters: None" -ForegroundColor Gray
            }
        } else {
            Write-Host "    Parameters: None" -ForegroundColor Gray
        }

        # Display Formats with live path preview
        if ($rep.PSObject.Properties['Formats'] -and $null -ne $rep.Formats) {
            $formatList = @($rep.Formats)
            if ($formatList.Count -gt 0) {
                Write-Host "    Formats:" -ForegroundColor Gray
                foreach ($f in $formatList) {
                    $rawPath = if ($f.PSObject.Properties['OutputPath'] -and $null -ne $f.OutputPath) { [string]$f.OutputPath } else { '' }
                    $rawFmt  = if ($f.PSObject.Properties['Format'] -and $null -ne $f.Format) { [string]$f.Format } else { '' }

                    Write-Host "      - $rawFmt -> $rawPath" -ForegroundColor White
                    if (-not [string]::IsNullOrWhiteSpace($rawPath)) {
                        $evaluatedPath = Resolve-DynamicTokens -Text $rawPath -Report $rep -Format $rawFmt
                        Write-Host "        Live File Preview: $evaluatedPath" -ForegroundColor DarkCyan
                    }
                }
            }
        }
    }
}

function Action-AddReport {
    param($Config)
    $session = Connect-CognosSession -Config $Config

    Write-Host "`n=== Add New Report ===" -ForegroundColor Cyan
    $source = Read-Host "Enter Report ID / StoreID (e.g. i54414D93B29A4D2289C4E88469871644)"
    if ([string]::IsNullOrWhiteSpace($source)) { return }

    $name = Read-Host "Enter friendly Report Name (optional, press Enter for default)"
    if ([string]::IsNullOrWhiteSpace($name)) { $name = "Report_$source" }

    Write-Host "Inspecting prompt parameters from Cognos..." -ForegroundColor Cyan
    $discovered = Get-ReportParameters -Context $session.Http -BaseUrl $Config.CognosBaseUrl -SourceType "report" -Source $source -Xsrf $session.Xsrf

    $params = [ordered]@{}
    if ($discovered.Count -gt 0) {
        Write-Host "`nFound $($discovered.Count) parameter(s):" -ForegroundColor Green
        foreach ($k in $discovered.Keys) {
            $params[$k] = Prompt-ForParameterValue -ParamName $k
        }
    } else {
        Write-Host "No prompt parameters required for this report." -ForegroundColor Yellow
    }

    # Choose format
    Write-Host "`nSelect Default Output Format:"
    Write-Host "  [1] xlsxData (Excel Data - simple tabular list)"
    Write-Host "  [2] spreadsheetML (Excel XML - best for formatted reports)"
    Write-Host "  [3] PDF"
    Write-Host "  [4] CSV"
    $fmtChoice = Read-Host "Choose format [1-4, default: 1]"
    $formatName = switch ($fmtChoice) {
        '2' { 'spreadsheetML' }
        '3' { 'PDF' }
        '4' { 'CSV' }
        default { 'xlsxData' }
    }
    $ext = switch ($formatName) {
        'PDF' { 'pdf' }
        'CSV' { 'csv' }
        'spreadsheetML' { 'xml' }
        default { 'xlsx' }
    }

    # Standardized Default Naming Template: D:\CognosReports\{Yesterday:yyyyMMdd}_{ReportName}.ext
    $defaultTemplate = "D:\CognosReports\{Yesterday:yyyyMMdd}_{ReportName}.${ext}"
    Write-Host "`nDefault File Path Template: $defaultTemplate" -ForegroundColor Cyan
    $pathInput = Read-Host "Enter Output Path (press Enter to accept default)"
    $outputPath = if (-not [string]::IsNullOrWhiteSpace($pathInput)) { $pathInput.Trim() } else { $defaultTemplate }

    $newRep = [ordered]@{
        Name       = $name
        Source     = $source
        SourceType = "report"
        Enabled    = $true
        Parameters = $params
        Formats    = @(
            [ordered]@{
                Format     = $formatName
                OutputPath = $outputPath
            }
        )
    }

    $list = [System.Collections.Generic.List[object]]::new()
    if ($Config.PSObject.Properties['Reports'] -and $Config.Reports) {
        foreach ($r in @($Config.Reports)) { $list.Add($r) }
    }
    $list.Add($newRep)

    $finalConfig = [ordered]@{
        CognosBaseUrl    = $Config.CognosBaseUrl
        Namespace        = $Config.Namespace
        CredentialTarget = $Config.CredentialTarget
        Reports          = $list
    }

    Save-Config -Config $finalConfig
    Write-Host "[SUCCESS] Added '$name' to configuration." -ForegroundColor Green
}

function Action-EditReport {
    param($Config)
    $reports = if ($Config.PSObject.Properties['Reports'] -and $Config.Reports) { @($Config.Reports) } else { @() }
    if (@($reports).Count -eq 0) {
        Write-Host "No reports available to edit." -ForegroundColor Yellow
        return
    }

    Action-ListReports -Config $Config
    $choice = Read-Host "`nEnter the number of the report you want to edit [1-$($reports.Count)]"
    $idx = [int]$choice - 1
    if ($idx -lt 0 -or $idx -ge @($reports).Count) { return }

    $rep = $reports[$idx]
    Write-Host "`nEditing: $($rep.Name)" -ForegroundColor Cyan
    Write-Host " [1] Edit Parameter Values (Set {Yesterday}, {Today}, or Custom)"
    Write-Host " [2] Toggle Enabled/Disabled (Current: $(if ($rep.PSObject.Properties['Enabled'] -and $rep.Enabled -eq $false) { 'DISABLED' } else { 'ENABLED' }))"
    Write-Host " [3] Edit Output Path Template"
    Write-Host " [4] Re-sync parameters from Cognos server"
    Write-Host " [0] Cancel"
    $action = Read-Host "Choose option [0-4]"

    switch ($action) {
        '1' {
            if ($rep.PSObject.Properties['Parameters'] -and $null -ne $rep.Parameters) {
                $paramProps = @($rep.Parameters.PSObject.Properties)
                if ($paramProps.Count -gt 0) {
                    Write-Host "`nEnter parameter values:" -ForegroundColor Cyan
                    foreach ($p in $paramProps) {
                        $p.Value = Prompt-ForParameterValue -ParamName ($p.Name) -CurrentVal ([string]$p.Value)
                    }
                } else {
                    Write-Host "This report has no parameters configured." -ForegroundColor Yellow
                }
            } else {
                Write-Host "This report has no parameters configured." -ForegroundColor Yellow
            }
        }
        '2' {
            $rep.Enabled = if ($rep.PSObject.Properties['Enabled']) { -not $rep.Enabled } else { $false }
            Write-Host "Report status updated to: $(if ($rep.Enabled) { 'ENABLED' } else { 'DISABLED' })" -ForegroundColor Green
        }
        '3' {
            if ($rep.PSObject.Properties['Formats'] -and $null -ne $rep.Formats) {
                foreach ($f in @($rep.Formats)) {
                    Write-Host "`nAvailable Tokens: {Yesterday:yyyyMMdd}, {Today:yyyyMMdd}, {ReportName}, {Format}" -ForegroundColor DarkGray
                    $newPath = Read-Host "Enter new path template for $($f.Format) [Current: $($f.OutputPath)]"
                    if (-not [string]::IsNullOrWhiteSpace($newPath)) { $f.OutputPath = $newPath }
                }
            }
        }
        '4' {
            $session = Connect-CognosSession -Config $Config
            Write-Host "Refreshing prompts from server..." -ForegroundColor Cyan
            $discovered = Get-ReportParameters -Context $session.Http -BaseUrl $Config.CognosBaseUrl -SourceType ($rep.SourceType) -Source ($rep.Source) -Xsrf $session.Xsrf
            
            $merged = [ordered]@{}
            if ($rep.PSObject.Properties['Parameters'] -and $rep.Parameters) {
                foreach ($p in $rep.Parameters.PSObject.Properties) { $merged[$p.Name] = $p.Value }
            }
            foreach ($k in $discovered.Keys) {
                if (-not $merged.Contains($k)) {
                    $merged[$k] = Prompt-ForParameterValue -ParamName $k
                    Write-Host "  + Discovered new parameter: $k = $($merged[$k])" -ForegroundColor Green
                }
            }
            $rep.Parameters = $merged
        }
        default { return }
    }

    Save-Config -Config $Config
}

function Action-RemoveReport {
    param($Config)
    $reports = if ($Config.PSObject.Properties['Reports'] -and $Config.Reports) { @($Config.Reports) } else { @() }
    if (@($reports).Count -eq 0) { return }

    Action-ListReports -Config $Config
    $choice = Read-Host "`nEnter number of report to DELETE [1-$($reports.Count), 0 to Cancel]"
    $idx = [int]$choice - 1
    if ($idx -lt 0 -or $idx -ge @($reports).Count) { return }

    $confirm = Read-Host "Are you sure you want to remove '$($reports[$idx].Name)'? (y/N)"
    if ($confirm -ieq 'y') {
        $list = [System.Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt @($reports).Count; $i++) {
            if ($i -ne $idx) { $list.Add($reports[$i]) }
        }
        $Config.Reports = $list
        Save-Config -Config $Config
        Write-Host "[OK] Report removed." -ForegroundColor Green
    }
}

function Action-TestConnection {
    param($Config)
    $session = Connect-CognosSession -Config $Config
    Write-Host "`n[SUCCESS] Successfully connected and authenticated with Cognos Analytics!" -ForegroundColor Green
}

# -----------------------------------------------------------------------------
# Main Menu Loop
# -----------------------------------------------------------------------------

$config = Load-Config

if ($null -eq $config) {
    Write-Host "No configuration file found at: $ConfigPath" -ForegroundColor Yellow
    $init = Read-Host "Would you like to initialize a new config now? (Y/n)"
    if ($init -ieq 'n') { exit 0 }
    $config = Action-InitConfig
}

while ($true) {
    Write-Host "`n=======================================================" -ForegroundColor Cyan
    Write-Host "           COGNOS CONFIGURATION MANAGER" -ForegroundColor Cyan
    Write-Host " Target File: $ConfigPath" -ForegroundColor DarkGray
    Write-Host " Server:      $($config.CognosBaseUrl) ($($config.Namespace))" -ForegroundColor DarkGray
    Write-Host "=======================================================" -ForegroundColor Cyan
    Write-Host " [1] List / View all configured reports (Live Previews)"
    Write-Host " [2] Add a new report (Auto-detect prompts & Date Presets)"
    Write-Host " [3] Edit an existing report (Parameters / Formats / Status)"
    Write-Host " [4] Remove a report"
    Write-Host " [5] Test Cognos connection & authentication"
    Write-Host " [6] Re-configure server URL / Credentials"
    Write-Host " [0] Exit"
    Write-Host "-------------------------------------------------------"

    $opt = Read-Host "Select an option [0-6]"
    switch ($opt) {
        '1' { Action-ListReports -Config $config }
        '2' { Action-AddReport -Config $config; $config = Load-Config }
        '3' { Action-EditReport -Config $config; $config = Load-Config }
        '4' { Action-RemoveReport -Config $config; $config = Load-Config }
        '5' { Action-TestConnection -Config $config }
        '6' { $config = Action-InitConfig }
        '0' { Write-Host "Goodbye!"; exit 0 }
        default { Write-Host "Invalid option." -ForegroundColor Yellow }
    }
}