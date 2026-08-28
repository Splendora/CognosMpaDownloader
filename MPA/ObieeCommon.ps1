#requires -Version 5.1
<#
.SYNOPSIS
    OBIEE / MPA SOAP Web Services Core Module.
.DESCRIPTION
    Shared library providing:
    - Win32 Windows Credential Manager integration (CredRead/CredWrite/CredFree for target 'MaCB')
    - Dynamic date and metadata token resolution ({Yesterday}, {Today}, {LastEOM}, {LastEOQ}, {LastEOY}, etc.)
    - Structured file logging, rolling monthly audit CSV, and summary reporting
    - SOAP 1.1 HTTP client dispatcher with XML Fault handling (namespace: urn://oracle.bi.webservices/v10)
    - Session management (Connect-Obiee / Disconnect-Obiee via nQSessionService)
    - Web Catalog operations (Get-ObieeItemInfo, Get-ObieeCatalogItems, Read-ObieeObject via webCatalogService)
    - Analysis export engine (Export-ObieeAnalysis, Save-ObieeExportData via analysisExportViewsService)
    - JSON configuration file management, validation, and automated backups
    - Robust retry execution policy with exponential backoff and session recovery
.NOTES
    Compatible with Windows PowerShell 5.1 and PowerShell Core 7+.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Net.Http

# -----------------------------------------------------------------------------
# 1. Win32 Windows Credential Manager P/Invoke
# -----------------------------------------------------------------------------
if (-not ([System.Management.Automation.PSTypeName]'ObieeCredManager.NativeMethods').Type) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace ObieeCredManager {
    public static class NativeMethods {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct CREDENTIAL {
            public uint Flags;
            public uint Type;
            public IntPtr TargetName;
            public IntPtr Comment;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
            public uint CredentialBlobSize;
            public IntPtr CredentialBlob;
            public uint Persist;
            public uint AttributeCount;
            public IntPtr Attributes;
            public IntPtr TargetAlias;
            public IntPtr UserName;
        }

        [DllImport("advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool CredRead(
            string target,
            uint type,
            int reservedFlag,
            out IntPtr credentialPtr
        );

        [DllImport("advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool CredWrite(
            [In] ref CREDENTIAL userCredential,
            uint flags
        );

        [DllImport("advapi32.dll", EntryPoint = "CredFree", SetLastError = false)]
        public static extern void CredFree([In] IntPtr buffer);
    }
}
"@
}

function Get-WindowsGenericCredential {
    param(
        [Parameter(Mandatory)]
        [string]$Target
    )

    $ptr = [IntPtr]::Zero
    try {
        $success = [ObieeCredManager.NativeMethods]::CredRead($Target, 1, 0, [ref]$ptr)
        if (-not $success) {
            $code = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
            if ($code -eq 1168) { return $null } # ERROR_NOT_FOUND
            throw "CredRead failed for target '$Target'. Windows error code: $code"
        }

        $method = [Runtime.InteropServices.Marshal].GetMethod('PtrToStructure', [type[]]@([IntPtr], [Type]))
        $native = $method.Invoke($null, @($ptr, [ObieeCredManager.NativeMethods+CREDENTIAL]))

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
            [ObieeCredManager.NativeMethods]::CredFree($ptr)
        }
    }
}

function Set-WindowsGenericCredential {
    param(
        [Parameter(Mandatory)] [string]$Target,
        [Parameter(Mandatory)] [string]$Username,
        [Parameter(Mandatory)] [Security.SecureString]$Password
    )

    $plain = Get-PlainTextFromSecureString -SecureString $Password
    $bytes = [Text.Encoding]::Unicode.GetBytes($plain)
    $blobPtr = [Runtime.InteropServices.Marshal]::AllocHGlobal($bytes.Length)
    $targetPtr = [Runtime.InteropServices.Marshal]::StringToCoTaskMemUni($Target)
    $userPtr = [Runtime.InteropServices.Marshal]::StringToCoTaskMemUni($Username)

    try {
        [Runtime.InteropServices.Marshal]::Copy($bytes, 0, $blobPtr, $bytes.Length)

        $cred = New-Object ObieeCredManager.NativeMethods+CREDENTIAL
        $cred.Type = 1 # CRED_TYPE_GENERIC
        $cred.TargetName = $targetPtr
        $cred.UserName = $userPtr
        $cred.CredentialBlob = $blobPtr
        $cred.CredentialBlobSize = [uint32]$bytes.Length
        $cred.Persist = 2 # CRED_PERSIST_LOCAL_MACHINE

        $success = [ObieeCredManager.NativeMethods]::CredWrite([ref]$cred, 0)
        if (-not $success) {
            $code = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
            throw "CredWrite failed for target '$Target'. Windows error code: $code"
        }
    }
    finally {
        if ($blobPtr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeHGlobal($blobPtr) }
        if ($targetPtr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeCoTaskMem($targetPtr) }
        if ($userPtr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeCoTaskMem($userPtr) }
        $plain = $null
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

function Get-ObieeCredential {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Target = 'MaCB'
    )

    $cred = Get-WindowsGenericCredential -Target $Target
    if ($null -eq $cred) { return $null }

    $plainPassword = Get-PlainTextFromSecureString -SecureString $cred.Password
    return [pscustomobject]@{
        Target   = $Target
        Username = $cred.Username
        Password = $plainPassword
    }
}

# -----------------------------------------------------------------------------
# 2. Dynamic Token & Date Resolution Engine
# -----------------------------------------------------------------------------

function Resolve-DynamicTokens {
    param(
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Text = '',

        $Report = $null,

        [AllowEmptyString()]
        [AllowNull()]
        [string]$Format = '',

        [string]$DefaultDateFormat = 'MM/dd/yyyy'
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }

    $now = Get-Date
    $resolved = $Text

    # 1. Base date calculations (Strictly End-of-Period focused)
    $today = $now.Date
    $yesterday = $today.AddDays(-1)
    $tomorrow = $today.AddDays(1)

    # Month End bounds
    $monthStart = [DateTime]::new($today.Year, $today.Month, 1)
    $monthEnd = $monthStart.AddMonths(1).AddDays(-1)
    $prevMonthEnd = $monthStart.AddDays(-1)
    $nextMonthEnd = $monthStart.AddMonths(2).AddDays(-1)

    # Quarter End bounds
    $curQ = [int][Math]::Floor(($today.Month - 1) / 3) + 1
    $quarterStart = [DateTime]::new($today.Year, (($curQ - 1) * 3 + 1), 1)
    $quarterEnd = $quarterStart.AddMonths(3).AddDays(-1)
    $prevQuarterEnd = $quarterStart.AddDays(-1)
    $nextQuarterEnd = $quarterStart.AddMonths(6).AddDays(-1)

    # Half-Year End bounds
    $halfYearEnd = if ($today.Month -le 6) { [DateTime]::new($today.Year, 6, 30) } else { [DateTime]::new($today.Year, 12, 31) }
    $prevHalfYearEnd = if ($today.Month -le 6) { [DateTime]::new($today.Year - 1, 12, 31) } else { [DateTime]::new($today.Year, 6, 30) }

    # Year End bounds
    $yearEnd = [DateTime]::new($today.Year, 12, 31)
    $prevYearEnd = [DateTime]::new($today.Year - 1, 12, 31)
    $nextYearEnd = [DateTime]::new($today.Year + 1, 12, 31)

    # Lookup mapping for strictly End-of-Period named tokens
    $dateTokenMap = @{
        'today'               = $today
        'yesterday'           = $yesterday
        'tomorrow'            = $tomorrow
        'eop'                 = $yesterday
        'lasteop'             = $prevMonthEnd
        'previouseop'         = $prevMonthEnd
        'monthend'            = $monthEnd
        'eom'                 = $monthEnd
        'previousmonthend'    = $prevMonthEnd
        'lastmonthend'        = $prevMonthEnd
        'prevmonthend'        = $prevMonthEnd
        'previouseom'         = $prevMonthEnd
        'lasteom'             = $prevMonthEnd
        'preveom'             = $prevMonthEnd
        'nextmonthend'        = $nextMonthEnd
        'nexteom'             = $nextMonthEnd
        'quarterend'          = $quarterEnd
        'eoq'                 = $quarterEnd
        'previousquarterend'  = $prevQuarterEnd
        'lastquarterend'      = $prevQuarterEnd
        'prevquarterend'      = $prevQuarterEnd
        'previouseoq'         = $prevQuarterEnd
        'lasteoq'             = $prevQuarterEnd
        'preveoq'             = $prevQuarterEnd
        'nextquarterend'      = $nextQuarterEnd
        'nexteoq'             = $nextQuarterEnd
        'halfyearend'         = $halfYearEnd
        'eoh'                 = $halfYearEnd
        'previoushalfyearend' = $prevHalfYearEnd
        'lasthalfyearend'     = $prevHalfYearEnd
        'prevhalfyearend'     = $prevHalfYearEnd
        'previouseoh'         = $prevHalfYearEnd
        'lasteoh'             = $prevHalfYearEnd
        'preveoh'             = $prevHalfYearEnd
        'yearend'             = $yearEnd
        'eoy'                 = $yearEnd
        'previousyearend'     = $prevYearEnd
        'lastyearend'         = $prevYearEnd
        'prevyearend'         = $prevYearEnd
        'previouseoy'         = $prevYearEnd
        'lasteoy'             = $prevYearEnd
        'preveoy'             = $prevYearEnd
        'nextyearend'         = $nextYearEnd
        'nexteoy'             = $nextYearEnd
    }

    # 2. Match and resolve named tokens with optional offsets and formatting:
    # Patterns: {Token}, {Token:Format}, {Token+3d}, {Token-7d:yyyyMMdd}
    $tokenNames = ($dateTokenMap.Keys | Sort-Object -Descending { $_.Length }) -join '|'
    $namedPattern = "(?i)\{($tokenNames)([+-]\d+d?)?(:([^}]+))?\}"

    $resolved = [regex]::Replace($resolved, $namedPattern, {
        param($match)
        $tokName = $match.Groups[1].Value.ToLowerInvariant()
        $baseDate = $dateTokenMap[$tokName]
        
        # Apply day offset if present (e.g. +1d, -7d)
        if ($match.Groups[2].Success) {
            $offsetStr = $match.Groups[2].Value.TrimEnd('d', 'D')
            $days = [int]$offsetStr
            $baseDate = $baseDate.AddDays($days)
        }

        # Apply custom format if specified, otherwise use DefaultDateFormat (MM/dd/yyyy)
        $fmt = if ($match.Groups[4].Success) { $match.Groups[4].Value } else { $DefaultDateFormat }
        return $baseDate.ToString($fmt, [System.Globalization.CultureInfo]::InvariantCulture)
    })

    # 3. Standard Timestamp Tokens
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $resolved = $resolved.Replace('{yyyy}', $now.ToString('yyyy', $inv))
    $resolved = $resolved.Replace('{yy}', $now.ToString('yy', $inv))
    $resolved = $resolved.Replace('{MM}', $now.ToString('MM', $inv))
    $resolved = $resolved.Replace('{dd}', $now.ToString('dd', $inv))
    $resolved = $resolved.Replace('{yyyyMM}', $now.ToString('yyyyMM', $inv))
    $resolved = $resolved.Replace('{yyyy-MM}', $now.ToString('yyyy-MM', $inv))
    $resolved = $resolved.Replace('{yyyyMMdd}', $now.ToString('yyyyMMdd', $inv))
    $resolved = $resolved.Replace('{yyyy-MM-dd}', $now.ToString('yyyy-MM-dd', $inv))
    $resolved = $resolved.Replace('{MM/dd/yyyy}', $now.ToString('MM/dd/yyyy', $inv))
    $resolved = $resolved.Replace('{MM-dd-yyyy}', $now.ToString('MM-dd-yyyy', $inv))
    $resolved = $resolved.Replace('{HHmmss}', $now.ToString('HHmmss', $inv))
    $resolved = $resolved.Replace('{HHmm}', $now.ToString('HHmm', $inv))
    $resolved = $resolved.Replace('{HH}', $now.ToString('HH', $inv))
    $resolved = $resolved.Replace('{Quarter}', "Q$curQ")
    $resolved = $resolved.Replace('{Q}', "$curQ")

    # 4. Report Metadata & Dynamic Parameter Tokens
    if ($null -ne $Report) {
        $repName = Get-PropOrKey -Object $Report -Name 'Name'
        $repPath = Get-PropOrKey -Object $Report -Name 'Path'
        $repInst = Get-PropOrKey -Object $Report -Name 'Instance'

        if (-not [string]::IsNullOrWhiteSpace($repName)) {
            $resolved = $resolved.Replace('{ReportName}', $repName)
        }
        if (-not [string]::IsNullOrWhiteSpace($repPath)) {
            $resolved = $resolved.Replace('{Path}', $repPath)
            $resolved = $resolved.Replace('{ReportPath}', $repPath)
        }
        if (-not [string]::IsNullOrWhiteSpace($repInst)) {
            $resolved = $resolved.Replace('{Instance}', $repInst)
        }

        # Parameter tokens ({p_<Name>})
        $paramsObj = Get-PropOrKey -Object $Report -Name 'Parameters'
        if ($null -ne $paramsObj) {
            if ($paramsObj -is [System.Collections.IDictionary]) {
                foreach ($pk in $paramsObj.Keys) {
                    $pv = [string]$paramsObj[$pk]
                    $pResolved = Resolve-DynamicTokens -Text $pv -DefaultDateFormat $DefaultDateFormat
                    $resolved = $resolved.Replace("{$pk}", $pResolved)
                }
            } else {
                foreach ($prop in $paramsObj.PSObject.Properties) {
                    $pv = [string]$prop.Value
                    $pResolved = Resolve-DynamicTokens -Text $pv -DefaultDateFormat $DefaultDateFormat
                    $resolved = $resolved.Replace("{$($prop.Name)}", $pResolved)
                }
            }
        }
    }

    # 5. Format Token
    if (-not [string]::IsNullOrWhiteSpace($Format)) {
        $resolved = $resolved.Replace('{Format}', $Format)
    }

    return $resolved
}

function Resolve-DynamicTokenArray {
    param(
        $Value,
        $Report = $null,
        [string]$DefaultDateFormat = 'MM/dd/yyyy'
    )

    if ($null -eq $Value) { return @() }

    # If already an array / collection (and not string)
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $results = New-Object System.Collections.Generic.List[string]
        foreach ($elem in $Value) {
            $sub = Resolve-DynamicTokenArray -Value $elem -Report $Report -DefaultDateFormat $DefaultDateFormat
            foreach ($s in $sub) { [void]$results.Add($s) }
        }
        return $results.ToArray()
    }

    $valStr = [string]$Value
    if ([string]::IsNullOrWhiteSpace($valStr)) { return @('') }

    $now = Get-Date
    $today = $now.Date
    $inv = [System.Globalization.CultureInfo]::InvariantCulture

    # Check for Multi-EOM / Multi-EOQ Range Tokens:
    $rangePattern = '(?i)^\{(AllEOM_YTD|EOM_YTD|AllEOM_CurrentYear|AllEOM_Year|EOM_Year|AllEOM_LastYear|AllEOM_PreviousYear|AllEOM_Last12M|EOM_Last12M|AllEOQ_YTD|EOQ_YTD|AllEOQ_CurrentYear|AllEOQ_Year|AllEOQ_LastYear|AllEOQ_PreviousYear|AllEOH_CurrentYear|AllEOH_Year|AllEOH_LastYear)(:([^}]+))?\}$'

    if ($valStr -match $rangePattern) {
        $tokType = $matches[1].ToLowerInvariant()
        $fmt = if ($matches[3]) { $matches[3] } else { $DefaultDateFormat }
        $dates = New-Object System.Collections.Generic.List[DateTime]

        # Check if custom year was passed in format (e.g. {AllEOM_Year:2025})
        $customYear = 0
        if ([int]::TryParse($fmt, [ref]$customYear) -and $customYear -ge 1900 -and $customYear -le 2100) {
            $fmt = $DefaultDateFormat
        } else {
            $customYear = 0
        }

        switch -Regex ($tokType) {
            '^(alleom_ytd|eom_ytd)$' {
                $maxMonth = [math]::Max(1, $today.Month - 1)
                for ($m = 1; $m -le $maxMonth; $m++) {
                    $start = [DateTime]::new($today.Year, $m, 1)
                    [void]$dates.Add($start.AddMonths(1).AddDays(-1))
                }
            }
            '^(alleom_currentyear|alleom_year|eom_year)$' {
                $targetYear = if ($customYear -gt 0) { $customYear } else { $today.Year }
                for ($m = 1; $m -le 12; $m++) {
                    $start = [DateTime]::new($targetYear, $m, 1)
                    [void]$dates.Add($start.AddMonths(1).AddDays(-1))
                }
            }
            '^(alleom_lastyear|alleom_previousyear)$' {
                $targetYear = $today.Year - 1
                for ($m = 1; $m -le 12; $m++) {
                    $start = [DateTime]::new($targetYear, $m, 1)
                    [void]$dates.Add($start.AddMonths(1).AddDays(-1))
                }
            }
            '^(alleom_last12m|eom_last12m)$' {
                $curMonthStart = [DateTime]::new($today.Year, $today.Month, 1)
                for ($i = 12; $i -ge 1; $i--) {
                    $mStart = $curMonthStart.AddMonths(-$i)
                    [void]$dates.Add($mStart.AddMonths(1).AddDays(-1))
                }
            }
            '^(alleoq_ytd|eoq_ytd)$' {
                $curQ = [int][Math]::Floor(($today.Month - 1) / 3) + 1
                for ($q = 1; $q -le $curQ; $q++) {
                    $qStart = [DateTime]::new($today.Year, (($q - 1) * 3 + 1), 1)
                    [void]$dates.Add($qStart.AddMonths(3).AddDays(-1))
                }
            }
            '^(alleoq_currentyear|alleoq_year)$' {
                $targetYear = if ($customYear -gt 0) { $customYear } else { $today.Year }
                for ($q = 1; $q -le 4; $q++) {
                    $qStart = [DateTime]::new($targetYear, (($q - 1) * 3 + 1), 1)
                    [void]$dates.Add($qStart.AddMonths(3).AddDays(-1))
                }
            }
            '^(alleoq_lastyear|alleoq_previousyear)$' {
                $targetYear = $today.Year - 1
                for ($q = 1; $q -le 4; $q++) {
                    $qStart = [DateTime]::new($targetYear, (($q - 1) * 3 + 1), 1)
                    [void]$dates.Add($qStart.AddMonths(3).AddDays(-1))
                }
            }
            '^(alleoh_currentyear|alleoh_year)$' {
                $targetYear = if ($customYear -gt 0) { $customYear } else { $today.Year }
                [void]$dates.Add([DateTime]::new($targetYear, 6, 30))
                [void]$dates.Add([DateTime]::new($targetYear, 12, 31))
            }
            '^(alleoh_lastyear|alleoh_previousyear)$' {
                $targetYear = $today.Year - 1
                [void]$dates.Add([DateTime]::new($targetYear, 6, 30))
                [void]$dates.Add([DateTime]::new($targetYear, 12, 31))
            }
        }

        $formattedList = New-Object System.Collections.Generic.List[string]
        foreach ($d in $dates) {
            [void]$formattedList.Add($d.ToString($fmt, $inv))
        }
        return $formattedList.ToArray()
    }

    # Otherwise, resolve standard dynamic tokens for single string
    $single = Resolve-DynamicTokens -Text $valStr -Report $Report -DefaultDateFormat $DefaultDateFormat
    return @($single)
}

# -----------------------------------------------------------------------------
# 3. Logging & Auditing Engine
# -----------------------------------------------------------------------------

$script:MpaLogConfig = $null
$script:MpaCurrentLogFile = $null
$script:MpaAuditCsvPath = $null
$script:ConsoleDebug = $false
$script:ExecutionRecords = New-Object System.Collections.Generic.List[PSCustomObject]
$script:BatchRunStartTime = $null

function Initialize-MpaLogging {
    param(
        $LoggingConfig,
        [string]$BaseDirectory = ''
    )

    $script:MpaLogConfig = $LoggingConfig
    if ($null -eq $LoggingConfig) { return }

    $enabled = Get-PropOrKey -Object $LoggingConfig -Name 'Enabled' -Default $true
    if (-not $enabled) { return }

    $logDir = Get-PropOrKey -Object $LoggingConfig -Name 'LogDirectory' -Default '.\Logs'
    if (-not [System.IO.Path]::IsPathRooted($logDir) -and -not [string]::IsNullOrWhiteSpace($BaseDirectory)) {
        $logDir = Join-Path $BaseDirectory $logDir
    }

    if (-not (Test-Path -LiteralPath $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    $logFileName = Get-PropOrKey -Object $LoggingConfig -Name 'LogFileName' -Default 'MpaDownloader_{yyyyMMdd}.log'
    $resolvedLogName = Resolve-DynamicTokens -Text $logFileName
    $script:MpaCurrentLogFile = Join-Path $logDir $resolvedLogName

    $auditEnabled = Get-PropOrKey -Object $LoggingConfig -Name 'AuditCsvEnabled' -Default $true
    if ($auditEnabled) {
        $auditPath = Get-PropOrKey -Object $LoggingConfig -Name 'AuditCsvPath' -Default '.\Logs\Audit_{yyyyMM}.csv'
        if (-not [System.IO.Path]::IsPathRooted($auditPath) -and -not [string]::IsNullOrWhiteSpace($BaseDirectory)) {
            $auditPath = Join-Path $BaseDirectory $auditPath
        }
        $resolvedAuditPath = Resolve-DynamicTokens -Text $auditPath
        $script:MpaAuditCsvPath = $resolvedAuditPath

        # Initialize Audit CSV Header if not present
        if (-not (Test-Path -LiteralPath $script:MpaAuditCsvPath)) {
            $auditDir = Split-Path -Parent $script:MpaAuditCsvPath
            if (-not (Test-Path -LiteralPath $auditDir)) {
                New-Item -ItemType Directory -Path $auditDir -Force | Out-Null
            }
            $header = '"Timestamp","ReportName","Path","Format","Status","DurationMs","FileSizeBytes","OutputPath","ErrorMessage"'
            [System.IO.File]::WriteAllLines($script:MpaAuditCsvPath, @($header), [System.Text.Encoding]::UTF8)
        }
    }

    $script:ConsoleDebug = Get-PropOrKey -Object $LoggingConfig -Name 'ConsoleDebug' -Default $false
    $script:BatchRunStartTime = Get-Date

    # Retention Cleanup
    $retentionDays = Get-PropOrKey -Object $LoggingConfig -Name 'RetentionDays' -Default 30
    if ($retentionDays -gt 0) {
        Invoke-LogRetentionCleanup -LogDirectory $logDir -RetentionDays $retentionDays
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)] [string]$Message,
        [ValidateSet('DEBUG', 'INFO', 'OK', 'WARN', 'ERROR')] [string]$Level = 'INFO'
    )

    $now = Get-Date
    $timestamp = $now.ToString('yyyy-MM-dd HH:mm:ss')
    $logLine = "[$timestamp] [$Level] $Message"

    $color = switch ($Level) {
        'DEBUG' { 'DarkGray' }
        'INFO'  { 'White' }
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
    }

    if ($Level -ne 'DEBUG' -or $script:ConsoleDebug) {
        Write-Host $logLine -ForegroundColor $color
    }

    if ($null -ne $script:MpaCurrentLogFile) {
        try {
            [System.IO.File]::AppendAllText($script:MpaCurrentLogFile, "$logLine`r`n", [System.Text.Encoding]::UTF8)
        }
        catch { }
    }
}

function Write-AuditLog {
    param(
        [Parameter(Mandatory)] [string]$ReportName,
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$Format,
        [Parameter(Mandatory)] [string]$Status,
        [int64]$DurationMs = 0,
        [int64]$FileSizeBytes = 0,
        [string]$OutputPath = '',
        [string]$ErrorMessage = ''
    )

    $record = [pscustomobject]@{
        Timestamp     = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        ReportName    = $ReportName
        Path          = $Path
        Format        = $Format
        Status        = $Status
        DurationMs    = $DurationMs
        FileSizeBytes = $FileSizeBytes
        OutputPath    = $OutputPath
        ErrorMessage  = $ErrorMessage
    }

    $script:ExecutionRecords.Add($record)

    if ($null -ne $script:MpaAuditCsvPath) {
        try {
            $csvLine = '"{0}","{1}","{2}","{3}","{4}",{5},{6},"{7}","{8}"' -f `
                $record.Timestamp,
                $record.ReportName.Replace('"', '""'),
                $record.Path.Replace('"', '""'),
                $record.Format,
                $record.Status,
                $record.DurationMs,
                $record.FileSizeBytes,
                $record.OutputPath.Replace('"', '""'),
                $record.ErrorMessage.Replace('"', '""')

            [System.IO.File]::AppendAllText($script:MpaAuditCsvPath, "$csvLine`r`n", [System.Text.Encoding]::UTF8)
        }
        catch { }
    }
}

function Write-ExecutionSummaryReport {
    param(
        [string]$BaseDirectory = ''
    )

    if ($script:ExecutionRecords.Count -eq 0) { return }

    $totalDuration = if ($null -ne $script:BatchRunStartTime) {
        [int]((Get-Date) - $script:BatchRunStartTime).TotalSeconds
    } else { 0 }

    $succeeded = @($script:ExecutionRecords | Where-Object { $_.Status -eq 'SUCCESS' }).Count
    $failed = @($script:ExecutionRecords | Where-Object { $_.Status -ne 'SUCCESS' }).Count
    $totalSize = ($script:ExecutionRecords | Measure-Object -Property FileSizeBytes -Sum).Sum
    $sizeMB = [math]::Round(($totalSize / 1MB), 2)

    Write-Host "`n" + ("=" * 80) -ForegroundColor Cyan
    Write-Host "                    BÁO CÁO TỔNG HỢP TIẾN ĐỘ TẢI MPA                    " -ForegroundColor Cyan
    Write-Host ("=" * 80) -ForegroundColor Cyan

    $summaryTable = foreach ($r in $script:ExecutionRecords) {
        [pscustomobject]@{
            'Báo cáo'    = $r.ReportName
            'Định dạng'  = $r.Format
            'Trạng thái' = $r.Status
            'Thời gian'  = "$([math]::Round($r.DurationMs / 1000, 1))s"
            'Dung lượng' = if ($r.FileSizeBytes -gt 0) { "$([math]::Round($r.FileSizeBytes / 1KB, 1)) KB" } else { '-' }
            'Tệp tin'    = if (-not [string]::IsNullOrWhiteSpace($r.OutputPath)) { [System.IO.Path]::GetFileName($r.OutputPath) } else { '-' }
        }
    }

    $summaryTable | Format-Table -AutoSize | Out-Host

    Write-Host ("-" * 80) -ForegroundColor Gray
    Write-Host "Tổng kết: $succeeded Thành công | $failed Thất bại | Thời gian: ${totalDuration}s | Tổng dung lượng: ${sizeMB} MB" `
        -ForegroundColor $(if ($failed -eq 0) { 'Green' } else { 'Yellow' })
    Write-Host ("=" * 80) + "`n" -ForegroundColor Cyan

    if ($null -ne $script:MpaLogConfig) {
        $summaryJsonEnabled = Get-PropOrKey -Object $script:MpaLogConfig -Name 'SummaryJsonEnabled' -Default $true
        if ($summaryJsonEnabled) {
            $summaryJsonPath = Get-PropOrKey -Object $script:MpaLogConfig -Name 'SummaryJsonPath' -Default '.\Logs\LatestRun.json'
            if (-not [System.IO.Path]::IsPathRooted($summaryJsonPath) -and -not [string]::IsNullOrWhiteSpace($BaseDirectory)) {
                $summaryJsonPath = Join-Path $BaseDirectory $summaryJsonPath
            }
            $runData = [pscustomobject]@{
                ExecutionTime   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                TotalReports    = $script:ExecutionRecords.Count
                Succeeded       = $succeeded
                Failed          = $failed
                TotalDurationSec= $totalDuration
                TotalSizeMB     = $sizeMB
                Records         = $script:ExecutionRecords
            }
            try {
                $jsonDir = Split-Path -Parent $summaryJsonPath
                if (-not (Test-Path -LiteralPath $jsonDir)) { New-Item -ItemType Directory -Path $jsonDir -Force | Out-Null }
                $jsonContent = $runData | ConvertTo-Json -Depth 5
                [System.IO.File]::WriteAllText($summaryJsonPath, $jsonContent, [System.Text.Encoding]::UTF8)
            }
            catch { }
        }
    }
}

function Invoke-LogRetentionCleanup {
    param(
        [Parameter(Mandatory)] [string]$LogDirectory,
        [int]$RetentionDays = 30
    )

    try {
        if (-not (Test-Path -LiteralPath $LogDirectory)) { return }
        $cutoff = (Get-Date).AddDays(-$RetentionDays)
        $files = Get-ChildItem -Path $LogDirectory -File | Where-Object { $_.LastWriteTime -lt $cutoff }
        foreach ($f in $files) {
            Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue
        }
    }
    catch { }
}

# -----------------------------------------------------------------------------
# 4. Low-Level SOAP Request Dispatcher
# -----------------------------------------------------------------------------

function Invoke-ObieeSoapRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ServiceUrl,

        [Parameter(Mandatory)]
        [string]$SoapBodyContent,

        [Parameter()]
        [int]$TimeoutSeconds = 120
    )

    $soapEnvelope = @"
<?xml version="1.0" encoding="utf-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:sawsoap="urn://oracle.bi.webservices/v10" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
    <soapenv:Header/>
    <soapenv:Body>
$SoapBodyContent
    </soapenv:Body>
</soapenv:Envelope>
"@

    $handler = New-Object System.Net.Http.HttpClientHandler
    $handler.UseCookies = $true
    $client = New-Object System.Net.Http.HttpClient($handler)
    $client.Timeout = [TimeSpan]::FromSeconds($TimeoutSeconds)

    $request = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Post, $ServiceUrl)
    $content = New-Object System.Net.Http.StringContent($soapEnvelope, [System.Text.Encoding]::UTF8, "text/xml")
    $request.Content = $content
    $request.Headers.Add("SOAPAction", '""')

    try {
        $response = $client.SendAsync($request).GetAwaiter().GetResult()
        $rawResponse = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()

        if (-not $response.IsSuccessStatusCode) {
            try {
                [xml]$faultDoc = $rawResponse
                $nsManager = New-Object System.Xml.XmlNamespaceManager($faultDoc.NameTable)
                $nsManager.AddNamespace("soapenv", "http://schemas.xmlsoap.org/soap/envelope/")
                $faultNode = $faultDoc.SelectSingleNode("//soapenv:Fault", $nsManager)
                if ($null -ne $faultNode) {
                    $faultCode = $faultNode.SelectSingleNode("faultcode")
                    $faultString = $faultNode.SelectSingleNode("faultstring")
                    $detail = $faultNode.SelectSingleNode("detail")
                    $fcText = if ($null -ne $faultCode) { $faultCode.InnerText } else { 'Unknown' }
                    $fsText = if ($null -ne $faultString) { $faultString.InnerText } else { 'SOAP Fault occurred' }
                    $dtText = if ($null -ne $detail) { $detail.InnerText } else { '' }
                    throw "OBIEE SOAP Fault [$fcText]: $fsText $($dtText.Trim())"
                }
            }
            catch [System.Management.Automation.RuntimeException] {
                if ($_.Exception.Message -match "OBIEE SOAP Fault") { throw }
            }

            $snippet = if ($rawResponse.Length -gt 200) { $rawResponse.Substring(0, 200) + '...' } else { $rawResponse }
            throw "HTTP $([int]$response.StatusCode) ($($response.StatusCode)) from $ServiceUrl. Response: $snippet"
        }

        try {
            [xml]$xmlDoc = $rawResponse
        }
        catch {
            $snippet = if ($rawResponse.Length -gt 200) { $rawResponse.Substring(0, 200) + '...' } else { $rawResponse }
            throw "Failed to parse OBIEE SOAP response as XML. Content: $snippet"
        }

        $nsManager = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
        $nsManager.AddNamespace("soapenv", "http://schemas.xmlsoap.org/soap/envelope/")
        $nsManager.AddNamespace("sawsoap", "urn://oracle.bi.webservices/v10")

        $faultNode = $xmlDoc.SelectSingleNode("//soapenv:Fault", $nsManager)
        if ($null -ne $faultNode) {
            $faultCode = $faultNode.SelectSingleNode("faultcode")
            $faultString = $faultNode.SelectSingleNode("faultstring")
            $detail = $faultNode.SelectSingleNode("detail")

            $fcText = if ($null -ne $faultCode) { $faultCode.InnerText } else { 'Unknown' }
            $fsText = if ($null -ne $faultString) { $faultString.InnerText } else { 'SOAP Fault occurred' }
            $dtText = if ($null -ne $detail) { $detail.InnerText } else { '' }

            throw "OBIEE SOAP Fault [$fcText]: $fsText $($dtText.Trim())"
        }

        return $xmlDoc
    }
    finally {
        if ($null -ne $request) { $request.Dispose() }
        if ($null -ne $client) { $client.Dispose() }
        if ($null -ne $handler) { $handler.Dispose() }
    }
}

# -----------------------------------------------------------------------------
# 5. Authentication & Session Management (nQSessionService)
# -----------------------------------------------------------------------------

function Connect-Obiee {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$BaseUrl,
        [Parameter(Mandatory)] [string]$Username,
        [Parameter(Mandatory)] [string]$Password
    )

    $endpoint = $BaseUrl.TrimEnd('/')
    if ($endpoint -notmatch '\?SoapImpl=') {
        $endpoint += "?SoapImpl=nQSessionService"
    }

    $body = @"
        <sawsoap:logon>
            <sawsoap:name>$Username</sawsoap:name>
            <sawsoap:password>$Password</sawsoap:password>
        </sawsoap:logon>
"@

    $xml = Invoke-ObieeSoapRequest -ServiceUrl $endpoint -SoapBodyContent $body
    $sessionNode = $xml.SelectSingleNode("//*[local-name()='sessionID' or local-name()='logonResult']")
    if ($null -eq $sessionNode -or [string]::IsNullOrWhiteSpace($sessionNode.InnerText)) {
        throw "Failed to extract sessionID from OBIEE logon response."
    }

    return $sessionNode.InnerText.Trim()
}

function Disconnect-Obiee {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$BaseUrl,
        [Parameter(Mandatory)] [string]$SessionId
    )

    $endpoint = $BaseUrl.TrimEnd('/')
    if ($endpoint -notmatch '\?SoapImpl=') {
        $endpoint += "?SoapImpl=nQSessionService"
    }

    $body = @"
        <sawsoap:logoff>
            <sawsoap:sessionID>$SessionId</sawsoap:sessionID>
        </sawsoap:logoff>
"@

    try {
        $null = Invoke-ObieeSoapRequest -ServiceUrl $endpoint -SoapBodyContent $body
        return $true
    }
    catch {
        Write-Log "OBIEE Disconnect warning: $_" 'WARN'
        return $false
    }
}

# -----------------------------------------------------------------------------
# 6. Catalog Operations (webCatalogService)
# -----------------------------------------------------------------------------

function Get-ObieeItemInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$BaseUrl,
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$SessionId,
        [Parameter()] [bool]$ResolveLinks = $true
    )

    $endpoint = $BaseUrl.TrimEnd('/')
    if ($endpoint -notmatch '\?SoapImpl=') {
        $endpoint += "?SoapImpl=webCatalogService"
    }

    $resolveStr = if ($ResolveLinks) { 'true' } else { 'false' }

    $body = @"
        <sawsoap:getItemInfo>
            <sawsoap:path>$Path</sawsoap:path>
            <sawsoap:resolveLinks>$resolveStr</sawsoap:resolveLinks>
            <sawsoap:sessionID>$SessionId</sawsoap:sessionID>
        </sawsoap:getItemInfo>
"@

    $xml = Invoke-ObieeSoapRequest -ServiceUrl $endpoint -SoapBodyContent $body
    $itemNode = $xml.SelectSingleNode("//*[local-name()='itemInfo' or local-name()='ItemInfo' or local-name()='getItemInfoResult']")
    if ($null -eq $itemNode) {
        throw "No ItemInfo returned for path: $Path"
    }

    return (Convert-ObieeXmlToItemInfo -Node $itemNode)
}

function Get-ObieeCatalogItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$BaseUrl,
        [Parameter(Mandatory)] [string]$Path,
        [Parameter()] [string]$Mask = '*',
        [Parameter(Mandatory)] [string]$SessionId,
        [Parameter()] [bool]$ResolveLinks = $true
    )

    $endpoint = $BaseUrl.TrimEnd('/')
    if ($endpoint -notmatch '\?SoapImpl=') {
        $endpoint += "?SoapImpl=webCatalogService"
    }

    $resolveStr = if ($ResolveLinks) { 'true' } else { 'false' }

    $body = @"
        <sawsoap:getSubItems>
            <sawsoap:path>$Path</sawsoap:path>
            <sawsoap:mask>$Mask</sawsoap:mask>
            <sawsoap:resolveLinks>$resolveStr</sawsoap:resolveLinks>
            <sawsoap:options>
                <sawsoap:includeACL>false</sawsoap:includeACL>
                <sawsoap:withPermission>0</sawsoap:withPermission>
                <sawsoap:withPermissionMask>0</sawsoap:withPermissionMask>
                <sawsoap:withAttributes>0</sawsoap:withAttributes>
                <sawsoap:withAttributesMask>0</sawsoap:withAttributesMask>
                <sawsoap:preserveOriginalLinkPath>false</sawsoap:preserveOriginalLinkPath>
            </sawsoap:options>
            <sawsoap:sessionID>$SessionId</sawsoap:sessionID>
        </sawsoap:getSubItems>
"@

    $xml = Invoke-ObieeSoapRequest -ServiceUrl $endpoint -SoapBodyContent $body
    $itemNodes = $xml.SelectNodes("//*[local-name()='itemInfo' or local-name()='ItemInfo']")
    $results = New-Object System.Collections.Generic.List[PSCustomObject]

    if ($null -ne $itemNodes) {
        foreach ($node in $itemNodes) {
            $parsed = Convert-ObieeXmlToItemInfo -Node $node
            $results.Add($parsed)
        }
    }

    return $results.ToArray()
}

function Read-ObieeObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$BaseUrl,
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$SessionId,
        [Parameter()] [bool]$ResolveLinks = $true
    )

    $endpoint = $BaseUrl.TrimEnd('/')
    if ($endpoint -notmatch '\?SoapImpl=') {
        $endpoint += "?SoapImpl=webCatalogService"
    }

    $resolveStr = if ($ResolveLinks) { 'true' } else { 'false' }

    $body = @"
        <sawsoap:readObject>
            <sawsoap:path>$Path</sawsoap:path>
            <sawsoap:resolveLinks>$resolveStr</sawsoap:resolveLinks>
            <sawsoap:outputEncodingMode>2</sawsoap:outputEncodingMode>
            <sawsoap:sessionID>$SessionId</sawsoap:sessionID>
        </sawsoap:readObject>
"@

    $xml = Invoke-ObieeSoapRequest -ServiceUrl $endpoint -SoapBodyContent $body
    $catalogObjectNode = $xml.SelectSingleNode("//*[local-name()='catalogObject' or local-name()='readObjectResult']")
    if ($null -eq $catalogObjectNode) {
        throw "Failed to read catalog object XML from path: $Path"
    }

    return $catalogObjectNode.InnerText
}

function Convert-ObieeXmlToItemInfo {
    param(
        [Parameter(Mandatory)]
        $Node
    )

    if ($null -eq $Node) { return $null }

    $targetNode = $Node
    if ($targetNode -is [System.Xml.XmlNode]) {
        $child = $targetNode.SelectSingleNode(".//*[local-name()='itemInfo' or local-name()='ItemInfo']")
        if ($null -ne $child) { $targetNode = $child }
    }

    $extract = {
        param([string]$fieldName)

        $lowerField = $fieldName.ToLowerInvariant()
        $upperField = $fieldName.ToUpperInvariant()

        if ($targetNode -is [System.Xml.XmlElement]) {
            $attrVal = $targetNode.GetAttribute($fieldName)
            if (-not [string]::IsNullOrWhiteSpace($attrVal)) { return $attrVal }
            $attrVal = $targetNode.GetAttribute($lowerField)
            if (-not [string]::IsNullOrWhiteSpace($attrVal)) { return $attrVal }
            $attrVal = $targetNode.GetAttribute($upperField)
            if (-not [string]::IsNullOrWhiteSpace($attrVal)) { return $attrVal }
        }

        if ($targetNode -is [System.Xml.XmlNode]) {
            $attrNode = $targetNode.SelectSingleNode("@*[translate(local-name(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='$lowerField']")
            if ($null -ne $attrNode -and -not [string]::IsNullOrWhiteSpace($attrNode.Value)) {
                return $attrNode.Value
            }

            $childNode = $targetNode.SelectSingleNode("*[translate(local-name(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='$lowerField']")
            if ($null -ne $childNode -and -not [string]::IsNullOrWhiteSpace($childNode.InnerText)) {
                return $childNode.InnerText
            }
        }

        return $null
    }

    $path      = & $extract 'path'
    $type      = & $extract 'type'
    $caption   = & $extract 'caption'
    $signature = & $extract 'signature'
    $lastMod   = & $extract 'lastModified'
    $created   = & $extract 'created'

    if ([string]::IsNullOrWhiteSpace($caption) -and -not [string]::IsNullOrWhiteSpace($path)) {
        $caption = [System.IO.Path]::GetFileName($path)
    }

    return [pscustomobject]@{
        Path         = $path
        Caption      = $caption
        Type         = $type
        Signature    = $signature
        LastModified = $lastMod
        Created      = $created
        RawNode      = $targetNode
    }
}

function Show-ObieeCatalogItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [array]$Items
    )

    if ($null -eq $Items -or $Items.Count -eq 0) {
        Write-Host "  (Không có mục nào trong danh mục)" -ForegroundColor DarkGray
        return
    }

    $tableData = foreach ($it in $Items) {
        [pscustomobject]@{
            'Type'      = if ($null -ne $it.Type) { $it.Type } else { 'N/A' }
            'Caption'   = if ($null -ne $it.Caption) { $it.Caption } else { 'N/A' }
            'Signature' = if ($null -ne $it.Signature) { $it.Signature } else { '' }
            'Path'      = if ($null -ne $it.Path) { $it.Path } else { 'N/A' }
        }
    }

    $tableData | Format-Table -AutoSize | Out-Host
}

function Get-ObjectKeyValuePairs {
    param($Object)
    if ($null -eq $Object) { return @() }
    if ($Object -is [System.Collections.IDictionary]) {
        return @($Object.Keys | ForEach-Object { [pscustomobject]@{ Name = $_; Value = $Object[$_] } })
    }
    return @($Object.PSObject.Properties | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Value = $_.Value } })
}

# -----------------------------------------------------------------------------
# 7. Analysis Export Service (analysisExportViewsService)
# -----------------------------------------------------------------------------

function Export-ObieeAnalysis {
    <#
    .SYNOPSIS
        Exports an OBIEE analysis report to CSV, PDF, EXCEL2007, or MHT format with optional parameters and filters.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$BaseUrl,
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [ValidateSet('CSV', 'PDF', 'EXCEL2007', 'MHT')] [string]$Format,
        [Parameter(Mandatory)] [string]$SessionId,
        $Parameters = $null,
        $FilterExpressions = $null,
        [int]$MaxRows = 100000,
        [bool]$Refresh = $true,
        [int]$TimeoutSeconds = 300
    )

    $endpoint = $BaseUrl.TrimEnd('/')
    if ($endpoint -notmatch '\?SoapImpl=') {
        $endpoint += "?SoapImpl=analysisExportViewsService"
    }

    $refreshStr = if ($Refresh) { 'true' } else { 'false' }

    # Construct reportParams XML (variables & filter expressions)
    $reportParamsXml = '<sawsoap:reportParams xsi:nil="true"/>'
    $hasParams = ($null -ne $Parameters -and (Get-ObjectKeyValuePairs -Object $Parameters).Count -gt 0)
    $hasFilters = ($null -ne $FilterExpressions -and @($FilterExpressions).Count -gt 0)

    if ($hasParams -or $hasFilters) {
        $sbParams = New-Object System.Text.StringBuilder
        [void]$sbParams.Append("<sawsoap:reportParams>")

        if ($hasFilters) {
            foreach ($flt in @($FilterExpressions)) {
                $resolvedFlt = Resolve-DynamicTokens -Text ([string]$flt)
                $escFlt = [System.Security.SecurityElement]::Escape($resolvedFlt)
                [void]$sbParams.Append("<sawsoap:filterExpressions>$escFlt</sawsoap:filterExpressions>")
            }
        }

        if ($hasParams) {
            $pairs = Get-ObjectKeyValuePairs -Object $Parameters
            foreach ($p in $pairs) {
                $pName = [System.Security.SecurityElement]::Escape([string]$p.Name)
                $pVal = $p.Value
                $resolvedArray = Resolve-DynamicTokenArray -Value $pVal -DefaultDateFormat 'MM/dd/yyyy'
                foreach ($item in $resolvedArray) {
                    $escItem = [System.Security.SecurityElement]::Escape([string]$item)
                    [void]$sbParams.Append("<sawsoap:variables><sawsoap:name>$pName</sawsoap:name><sawsoap:value>$escItem</sawsoap:value></sawsoap:variables>")
                }
            }
        }

        [void]$sbParams.Append("</sawsoap:reportParams>")
        $reportParamsXml = $sbParams.ToString()
    }

    $body = @"
        <sawsoap:initiateAnalysisExport>
            <sawsoap:report>
                <sawsoap:reportPath>$Path</sawsoap:reportPath>
            </sawsoap:report>
            <sawsoap:outputFormat>$Format</sawsoap:outputFormat>
            <sawsoap:executionOptions>
                <sawsoap:async>false</sawsoap:async>
                <sawsoap:maxRowsPerPage>$MaxRows</sawsoap:maxRowsPerPage>
                <sawsoap:refresh>$refreshStr</sawsoap:refresh>
                <sawsoap:useMtom>false</sawsoap:useMtom>
            </sawsoap:executionOptions>
            $reportParamsXml
            <sawsoap:reportViewName xsi:nil="true"/>
            <sawsoap:sessionID>$SessionId</sawsoap:sessionID>
        </sawsoap:initiateAnalysisExport>
"@

    $xml = Invoke-ObieeSoapRequest -ServiceUrl $endpoint -SoapBodyContent $body -TimeoutSeconds $TimeoutSeconds

    $statusNode = $xml.SelectSingleNode("//*[local-name()='exportStatus']")
    $mimeNode   = $xml.SelectSingleNode("//*[local-name()='mimeType']")
    $dataNode   = $xml.SelectSingleNode("//*[local-name()='viewData']")
    $queryNode  = $xml.SelectSingleNode("//*[local-name()='queryID']")

    $status = if ($null -ne $statusNode) { $statusNode.InnerText.Trim() } else { 'Done' }
    $mimeType = if ($null -ne $mimeNode) { $mimeNode.InnerText.Trim() } else { '' }
    $viewData = if ($null -ne $dataNode) { $dataNode.InnerText } else { '' }
    $queryId = if ($null -ne $queryNode) { $queryNode.InnerText.Trim() } else { '' }

    if ($status -eq 'Error') {
        throw "OBIEE Analysis Export returned status 'Error' for path: $Path"
    }

    # If viewData is not directly included, complete the export with queryID
    if ([string]::IsNullOrWhiteSpace($viewData) -and -not [string]::IsNullOrWhiteSpace($queryId)) {
        $completeBody = @"
        <sawsoap:completeAnalysisExport>
            <sawsoap:queryID>$queryId</sawsoap:queryID>
            <sawsoap:sessionID>$SessionId</sawsoap:sessionID>
        </sawsoap:completeAnalysisExport>
"@
        $xmlComplete = Invoke-ObieeSoapRequest -ServiceUrl $endpoint -SoapBodyContent $completeBody -TimeoutSeconds $TimeoutSeconds
        $compDataNode = $xmlComplete.SelectSingleNode("//*[local-name()='viewData']")
        if ($null -ne $compDataNode) {
            $viewData = $compDataNode.InnerText
        }
        $compMimeNode = $xmlComplete.SelectSingleNode("//*[local-name()='mimeType']")
        if ($null -ne $compMimeNode) {
            $mimeType = $compMimeNode.InnerText.Trim()
        }
    }

    return [pscustomobject]@{
        Path         = $Path
        Format       = $Format
        Status       = $status
        MimeType     = $mimeType
        ViewData     = $viewData
        QueryId      = $queryId
        XmlResponse  = $xml
    }
}

function Save-ObieeExportData {
    <#
    .SYNOPSIS
        Decodes and writes exported analysis payload to file on disk.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$ViewData,
        [Parameter(Mandatory)] [string]$Format,
        [Parameter(Mandatory)] [string]$OutputPath
    )

    $parentDir = Split-Path -Parent $OutputPath
    if (-not [string]::IsNullOrWhiteSpace($parentDir) -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $isBinary = ($Format -in @('PDF', 'EXCEL2007', 'PNG', 'GIF'))

    if ($isBinary) {
        try {
            $bytes = [System.Convert]::FromBase64String($ViewData.Trim())
            [System.IO.File]::WriteAllBytes($OutputPath, $bytes)
        }
        catch {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($ViewData)
            [System.IO.File]::WriteAllBytes($OutputPath, $bytes)
        }
    }
    else {
        $decodedText = $ViewData
        try {
            $testBytes = [System.Convert]::FromBase64String($ViewData.Trim())
            if ($testBytes.Length -gt 0) {
                $decodedText = [System.Text.Encoding]::UTF8.GetString($testBytes)
            }
        }
        catch {
            $decodedText = $ViewData
        }

        [System.IO.File]::WriteAllText($OutputPath, $decodedText, [System.Text.Encoding]::UTF8)
    }

    $fileInfo = Get-Item -LiteralPath $OutputPath
    return $fileInfo.Length
}

# -----------------------------------------------------------------------------
# 8. Configuration File Management
# -----------------------------------------------------------------------------

function Load-MpaConfig {
    param(
        [Parameter(Mandatory)] [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) { return $null }

    $json = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    return ($json | ConvertFrom-Json)
}

function Save-MpaConfig {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] $Config
    )

    if (Test-Path -LiteralPath $Path) {
        $backupPath = "$Path.bak"
        Copy-Item -LiteralPath $Path -Destination $backupPath -Force
    }

    $json = $Config | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($Path, $json, [System.Text.Encoding]::UTF8)
}

function Get-PropOrKey {
    param(
        $Object,
        [string]$Name,
        $Default = $null
    )

    if ($null -eq $Object) { return $Default }

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $Default
    }

    $prop = $Object.PSObject.Properties[$Name]
    if ($null -ne $prop) { return $prop.Value }

    return $Default
}

function Get-RequiredProperty {
    param(
        [Parameter(Mandatory)] $Object,
        [Parameter(Mandatory)] [string]$Name
    )

    $val = Get-PropOrKey -Object $Object -Name $Name
    if ($null -eq $val -or [string]::IsNullOrWhiteSpace([string]$val)) {
        throw "Cấu hình thiếu trường bắt buộc: '$Name'"
    }
    return $val
}

function Get-MpaInstances {
    param([Parameter(Mandatory)] $Config)
    $instObj = Get-PropOrKey -Object $Config -Name 'Instances'
    $dict = [ordered]@{}
    if ($null -ne $instObj) {
        if ($instObj -is [System.Collections.IDictionary]) {
            foreach ($k in $instObj.Keys) { $dict[$k] = $instObj[$k] }
        }
        else {
            foreach ($p in $instObj.PSObject.Properties) { $dict[$p.Name] = $p.Value }
        }
    }
    return $dict
}

function Get-MpaReportInstance {
    param(
        [Parameter(Mandatory)] $Config,
        [Parameter(Mandatory)] $Report
    )

    $defaultInstName = Get-PropOrKey -Object $Config -Name 'DefaultInstance' -Default 'MPA'
    $reportInstName = Get-PropOrKey -Object $Report -Name 'Instance' -Default $defaultInstName
    $instances = Get-MpaInstances -Config $Config

    if ($instances.Contains($reportInstName)) {
        return @{
            Name = $reportInstName
            Config = $instances[$reportInstName]
        }
    }

    if ($instances.Contains($defaultInstName)) {
        return @{
            Name = $defaultInstName
            Config = $instances[$defaultInstName]
        }
    }

    throw "Không tìm thấy cấu hình máy chủ cho instance: '$reportInstName'"
}

function Get-MpaHttpSettings {
    param([Parameter(Mandatory)] $Config)
    $http = Get-PropOrKey -Object $Config -Name 'HttpSettings'
    $timeoutMin = if ($null -ne $http) { Get-PropOrKey -Object $http -Name 'TimeoutMinutes' -Default 5 } else { 5 }
    return @{ TimeoutMinutes = [int]$timeoutMin }
}

function Get-MpaRetryPolicy {
    param([Parameter(Mandatory)] $Config)
    $rp = Get-PropOrKey -Object $Config -Name 'RetryPolicy'
    if ($null -eq $rp) {
        return @{
            Enabled             = $true
            MaxRetries          = 3
            InitialDelaySeconds = 5
            BackoffMultiplier   = 2
        }
    }

    return @{
        Enabled             = [bool](Get-PropOrKey -Object $rp -Name 'Enabled' -Default $true)
        MaxRetries          = [int](Get-PropOrKey -Object $rp -Name 'MaxRetries' -Default 3)
        InitialDelaySeconds = [int](Get-PropOrKey -Object $rp -Name 'InitialDelaySeconds' -Default 5)
        BackoffMultiplier   = [int](Get-PropOrKey -Object $rp -Name 'BackoffMultiplier' -Default 2)
    }
}

function Get-MpaSchedulingConfig {
    param([Parameter(Mandatory)] $Config)
    $sched = Get-PropOrKey -Object $Config -Name 'Scheduling'
    if ($null -eq $sched) {
        return @{
            TaskName     = 'MpaReportDownloader'
            ScheduleType = 'DAILY'
            StartTime    = '08:00'
            RunElevated  = $false
        }
    }

    return @{
        TaskName     = [string](Get-PropOrKey -Object $sched -Name 'TaskName' -Default 'MpaReportDownloader')
        ScheduleType = [string](Get-PropOrKey -Object $sched -Name 'ScheduleType' -Default 'DAILY')
        StartTime    = [string](Get-PropOrKey -Object $sched -Name 'StartTime' -Default '08:00')
        RunElevated  = [bool](Get-PropOrKey -Object $sched -Name 'RunElevated' -Default $false)
    }
}

function Get-MpaScheduledTask {
    param(
        [string]$TaskName = 'MpaReportDownloader'
    )

    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($null -eq $task) { return $null }

        $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction SilentlyContinue

        $triggers = if ($task.Triggers) { @($task.Triggers) } else { @() }
        $trigger = if ($triggers.Count -gt 0) { $triggers[0] } else { $null }
        $scheduleType = if ($trigger) {
            if ($trigger.CimClass.CimClassName -match 'Daily') { 'DAILY' }
            elseif ($trigger.CimClass.CimClassName -match 'Weekly') {
                if ($trigger.DaysOfWeek -eq 62 -or "$($trigger.DaysOfWeek)" -match 'Monday|Tuesday|Wednesday|Thursday|Friday') {
                    'WEEKDAY'
                } else {
                    'WEEKLY'
                }
            }
            elseif ($trigger.CimClass.CimClassName -match 'Time' -or $trigger.Repetition.Interval) { 'HOURLY' }
            else { 'CUSTOM' }
        } else { 'UNKNOWN' }

        $startTime = if ($trigger -and $trigger.StartBoundary) {
            try { [DateTime]::Parse($trigger.StartBoundary).ToString("HH:mm") } catch { $trigger.StartBoundary }
        } else { '' }

        return [pscustomobject]@{
            Exists         = $true
            TaskName       = $task.TaskName
            State          = [string]$task.State
            ScheduleType   = $scheduleType
            StartTime      = $startTime
            LastRunTime    = if ($taskInfo) { $taskInfo.LastRunTime } else { $null }
            NextRunTime    = if ($taskInfo) { $taskInfo.NextRunTime } else { $null }
            LastTaskResult = if ($taskInfo) { $taskInfo.LastTaskResult } else { 0 }
            Author         = $task.Author
            Description    = $task.Description
        }
    }
    catch {
        return $null
    }
}

function Set-MpaScheduledTask {
    param(
        [string]$TaskName = 'MpaReportDownloader',
        [string]$ScriptPath = '',
        [string]$ConfigPath = '',
        [ValidateSet('DAILY', 'WEEKDAY', 'WEEKLY', 'HOURLY')]
        [string]$ScheduleType = 'DAILY',
        [string]$StartTime = '08:00',
        [bool]$RunElevated = $false
    )

    if (-not $ScriptPath) {
        $ScriptPath = Join-Path -Path $PSScriptRoot -ChildPath 'MpaReportDownloader.ps1'
    }
    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "Không tìm thấy tệp thực thi tải báo cáo tại: $ScriptPath"
    }

    $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""
    if ($ConfigPath -and (Test-Path -LiteralPath $ConfigPath)) {
        $argList += " -ConfigPath `"$ConfigPath`""
    }

    $psExe = (Get-Command 'powershell.exe').Source
    if (-not $psExe) { $psExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" }

    $parsedTime = [DateTime]::Today.AddHours(8)
    if ($StartTime -match '^(\d{1,2}):(\d{2})$') {
        $hh = [int]$matches[1]
        $mm = [int]$matches[2]
        $parsedTime = [DateTime]::Today.AddHours($hh).AddMinutes($mm)
    }

    $action = New-ScheduledTaskAction -Execute $psExe -Argument $argList -WorkingDirectory (Split-Path -Parent $ScriptPath)
    
    $trigger = switch ($ScheduleType) {
        'DAILY'   { New-ScheduledTaskTrigger -Daily -At $parsedTime }
        'WEEKDAY' { New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At $parsedTime }
        'WEEKLY'  { New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At $parsedTime }
        'HOURLY'  { 
            $t = New-ScheduledTaskTrigger -Once -At $parsedTime
            $t.RepetitionDuration = [TimeSpan]::FromDays(3650)
            $t.RepetitionInterval = [TimeSpan]::FromHours(1)
            $t
        }
    }

    $principal = if ($RunElevated) {
        New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Highest
    } else {
        New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive
    }

    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew

    $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Tự động tải báo cáo OBIEE / MPA (MpaReportDownloader)"
    
    Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null
    Write-Log -Message "Đã tạo/cập nhật thành công lịch tự động '$TaskName' (Loại: $ScheduleType, Giờ chạy: $($parsedTime.ToString('HH:mm')))" -Level 'OK'
    return $true
}

function Remove-MpaScheduledTask {
    param([Parameter(Mandatory)] [string]$TaskName)
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
}

function Start-MpaScheduledTask {
    param([Parameter(Mandatory)] [string]$TaskName)
    Start-ScheduledTask -TaskName $TaskName -ErrorAction Stop
}

function Set-ObjectProperty {
    param(
        [Parameter(Mandatory)] $Object,
        [Parameter(Mandatory)] [string]$Name,
        $Value
    )

    if ($Object -is [System.Collections.IDictionary]) {
        $Object[$Name] = $Value
        return
    }

    if ($Object.PSObject.Properties[$Name]) {
        $Object.PSObject.Properties[$Name].Value = $Value
    }
    else {
        $Object | Add-Member -MemberType NoteProperty -Name $Name -Value $Value -Force
    }
}

# -----------------------------------------------------------------------------
# 9. Download Execution with Retry Policy
# -----------------------------------------------------------------------------

function Invoke-MpaReportDownloadWithRetry {
    param(
        [Parameter(Mandatory)] [string]$BaseUrl,
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$Format,
        [Parameter(Mandatory)] [string]$OutputPath,
        [Parameter(Mandatory)] [string]$SessionId,
        [Parameter(Mandatory)] [hashtable]$RetryPolicy,
        $Parameters = $null,
        $FilterExpressions = $null,
        [int]$TimeoutSeconds = 300,
        [scriptblock]$OnRenewSession = $null
    )

    $maxAttempts = if ($RetryPolicy.Enabled) { [math]::Max(1, $RetryPolicy.MaxRetries) } else { 1 }
    $delay = $RetryPolicy.InitialDelaySeconds
    $currentSessionId = $SessionId

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            Write-Log "Đang tải báo cáo: '$Path' (Định dạng: $Format, Lần thử: $attempt/$maxAttempts)..." 'INFO'
            
            $exportResult = Export-ObieeAnalysis `
                -BaseUrl $BaseUrl `
                -Path $Path `
                -Format $Format `
                -SessionId $currentSessionId `
                -Parameters $Parameters `
                -FilterExpressions $FilterExpressions `
                -TimeoutSeconds $TimeoutSeconds

            $fileSize = Save-ObieeExportData `
                -ViewData $exportResult.ViewData `
                -Format $Format `
                -OutputPath $OutputPath

            $sw.Stop()
            $sizeKB = [math]::Round($fileSize / 1KB, 2)
            Write-Log "Tải thành công: '$OutputPath' ($sizeKB KB, $($sw.ElapsedMilliseconds)ms)" 'OK'

            return [pscustomobject]@{
                Success       = $true
                FileSizeBytes = $fileSize
                DurationMs    = $sw.ElapsedMilliseconds
                OutputPath    = $OutputPath
                SessionId     = $currentSessionId
                ErrorMessage  = ''
            }
        }
        catch {
            $sw.Stop()
            $errMsg = $_.Exception.Message

            if ($attempt -lt $maxAttempts) {
                Write-Log "Lỗi khi tải báo cáo '$Path': $errMsg. Thử lại sau ${delay}s..." 'WARN'
                Start-Sleep -Seconds $delay
                $delay = $delay * $RetryPolicy.BackoffMultiplier

                if ($null -ne $OnRenewSession) {
                    try {
                        Write-Log "Đang làm mới phiên kết nối OBIEE..." 'INFO'
                        $currentSessionId = & $OnRenewSession
                    }
                    catch {
                        Write-Log "Không thể làm mới phiên: $_" 'WARN'
                    }
                }
            }
            else {
                Write-Log "Thất bại khi tải báo cáo '$Path' sau $maxAttempts lần thử: $errMsg" 'ERROR'
                return [pscustomobject]@{
                    Success       = $false
                    FileSizeBytes = 0
                    DurationMs    = $sw.ElapsedMilliseconds
                    OutputPath    = $OutputPath
                    SessionId     = $currentSessionId
                    ErrorMessage  = $errMsg
                }
            }
        }
    }
}
