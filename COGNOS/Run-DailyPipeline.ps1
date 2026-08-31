<#
.SYNOPSIS
    Automated Cognos -> Excel VBA daily pipeline.

.DESCRIPTION
    Existing Cognos downloader and existing VBA processing macros are reused.

    Pipeline:
        1. Download TONG QUAN TTKH
        2. Run handle_file_tong_quan
        3. Verify fresh Input_cif.txt
        4. Download TONG HOA LOI ICH ODS
        5. Download SPDV
        6. Run handle_file_tong_hoa_loi_ich
        7. Run handle_file_spdv

    The original cognos-reports.json is NOT modified.
    Temporary configuration files are created for each Cognos step.

.NOTES
    Designed for Windows PowerShell 5.1.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================================
# PATHS
# ============================================================================

$DownloaderPath = 'D:\Repo\CognosMpaDownloader\COGNOS\CognosReportDownloader.ps1'

$ConfigPath = 'D:\Repo\CognosMpaDownloader\COGNOS\cognos-reports.json'

$MacroWorkbook = '\\10.26.136.31\DuLieu121\DATAHUB\Xu ly du lieu_update.xlsm'

$WorkingDirectory = '\\10.26.136.31\DuLieu121\DATAHUB'

$InputCifPath = Join-Path $WorkingDirectory 'Input_cif.txt'

# ============================================================================
# VALIDATE PATHS
# ============================================================================

if (-not (Test-Path -LiteralPath $DownloaderPath)) {
    throw "Cognos downloader not found: $DownloaderPath"
}

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "Cognos configuration not found: $ConfigPath"
}

if (-not (Test-Path -LiteralPath $MacroWorkbook)) {
    throw "Macro workbook not found: $MacroWorkbook"
}

if (-not (Test-Path -LiteralPath $WorkingDirectory)) {
    throw "Working directory not found: $WorkingDirectory"
}

# ============================================================================
# REPORT DATE
# ============================================================================

$ReportDate = (Get-Date).Date.AddDays(-1)

# ============================================================================
# LOGGING
# ============================================================================

$PipelineStart = Get-Date

function Write-PipelineLog {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $Time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    switch ($Level) {
        'INFO' {
            Write-Host "[$Time] $Message"
        }

        'OK' {
            Write-Host "[$Time] [OK] $Message"
        }

        'WARN' {
            Write-Host "[$Time] [WARN] $Message"
        }

        'ERROR' {
            Write-Host "[$Time] [ERROR] $Message"
        }
    }
}

# ============================================================================
# LOAD ORIGINAL CONFIGURATION
# ============================================================================

Write-PipelineLog 'Loading Cognos configuration...'

$config = Get-Content `
    -LiteralPath $ConfigPath `
    -Raw |
    ConvertFrom-Json

if ($null -eq $config.Reports) {
    throw "No Reports section found in $ConfigPath"
}

# ============================================================================
# CONFIG HELPERS
# ============================================================================

function Get-ReportByName {
    param(
        [Parameter(Mandatory)]
        $Config,

        [Parameter(Mandatory)]
        [string]$ReportName
    )

    foreach ($Report in @($Config.Reports)) {
        if ([string]$Report.Name -eq $ReportName) {
            return $Report
        }
    }

    throw "Report '$ReportName' not found in $ConfigPath"
}

function Get-ReportOutputPath {
    param(
        [Parameter(Mandatory)]
        $Report
    )

    if ($null -eq $Report.Formats) {
        throw "Report '$($Report.Name)' has no Formats configuration."
    }

    $Formats = @($Report.Formats)

    if ($Formats.Count -eq 0) {
        throw "Report '$($Report.Name)' has no Formats configuration."
    }

    $FormatConfig = $Formats[0]

    if ($null -eq $FormatConfig.OutputPath) {
        throw "Report '$($Report.Name)' has no OutputPath."
    }

    $OutputPath = [string]$FormatConfig.OutputPath

    $OutputPath = $OutputPath.Replace(
        '{ReportName}',
        [string]$Report.Name
    )

    return $OutputPath
}

# ============================================================================
# CREATE TEMPORARY CONFIG
# ============================================================================

function New-StepConfig {
    param(
        [Parameter(Mandatory)]
        [string]$ReportName
    )

    # Reload original config every time.
    # The permanent config is therefore never modified.
    $StepConfig = Get-Content `
        -LiteralPath $ConfigPath `
        -Raw |
        ConvertFrom-Json

    foreach ($Report in @($StepConfig.Reports)) {

        if ([string]$Report.Name -eq $ReportName) {
            $Report.Enabled = $true
        }
        else {
            $Report.Enabled = $false
        }
    }

    $TempFileName = ".pipeline-$([Guid]::NewGuid().ToString('N')).json"

    $TempPath = Join-Path `
        (Split-Path -Parent $ConfigPath) `
        $TempFileName

    $Json = $StepConfig | ConvertTo-Json -Depth 20

    # UTF-8 without BOM
    $Utf8 = New-Object System.Text.UTF8Encoding($false)

    [System.IO.File]::WriteAllText(
        $TempPath,
        $Json,
        $Utf8
    )

    return $TempPath
}

# ============================================================================
# RUN ONE COGNOS REPORT
# ============================================================================

function Invoke-CognosReport {
    param(
        [Parameter(Mandatory)]
        [string]$ReportName
    )

    $Report = Get-ReportByName `
        -Config $config `
        -ReportName $ReportName

    $OutputPath = Get-ReportOutputPath -Report $Report

    Write-PipelineLog '------------------------------------------------------------'
    Write-PipelineLog "Cognos report: $ReportName"
    Write-PipelineLog "Output: $OutputPath"
    Write-PipelineLog '------------------------------------------------------------'

    # Remove old output to avoid accidentally using stale data.
    if (Test-Path -LiteralPath $OutputPath) {

        Write-PipelineLog `
            "Removing previous output file." `
            'WARN'

        Remove-Item `
            -LiteralPath $OutputPath `
            -Force
    }

    $TempConfig = New-StepConfig -ReportName $ReportName

    try {

        Write-PipelineLog 'Starting Cognos downloader...'

        & "$PSHOME\powershell.exe" `
            -NoProfile `
            -ExecutionPolicy Bypass `
            -File $DownloaderPath `
            -ConfigPath $TempConfig

        $ExitCode = $LASTEXITCODE

        if ($ExitCode -ne 0) {
            throw `
                "Cognos downloader failed for '$ReportName'. Exit code: $ExitCode"
        }

        if (-not (Test-Path -LiteralPath $OutputPath)) {
            throw `
                "Cognos downloader reported success, but output was not found: $OutputPath"
        }

        $OutputFile = Get-Item -LiteralPath $OutputPath

        if ($OutputFile.Length -le 0) {
            throw `
                "Downloaded output is empty: $OutputPath"
        }

        Write-PipelineLog `
            "Cognos download completed: $ReportName" `
            'OK'

        Write-PipelineLog `
            "File size: $([math]::Round($OutputFile.Length / 1MB, 2)) MB"
    }
    finally {

        if (Test-Path -LiteralPath $TempConfig) {

            Remove-Item `
                -LiteralPath $TempConfig `
                -Force `
                -ErrorAction SilentlyContinue
        }
    }
}

# ============================================================================
# RUN EXISTING EXCEL VBA MACRO
# ============================================================================

function Invoke-ExcelMacro {
    param(
        [Parameter(Mandatory)]
        [string]$MacroName
    )

    Write-PipelineLog '------------------------------------------------------------'
    Write-PipelineLog "Excel VBA: $MacroName"
    Write-PipelineLog '------------------------------------------------------------'

    $Excel = $null
    $Workbook = $null

    try {

        Write-PipelineLog 'Starting Excel...'

        $Excel = New-Object -ComObject Excel.Application

        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false

        # Allow VBA to execute when workbook is opened by automation.
        $Excel.AutomationSecurity = 1

        Write-PipelineLog "Opening: $MacroWorkbook"

        $Workbook = $Excel.Workbooks.Open(
            $MacroWorkbook,
            $false,
            $false
        )

        Write-PipelineLog "Running macro: $MacroName"

        # Explicitly qualify the workbook so Excel doesn't run a macro
        # with the same name from another open workbook.
        $Excel.Run(
            "'$($Workbook.Name)'!$MacroName"
        )

        Write-PipelineLog `
            "Macro completed: $MacroName" `
            'OK'

        # The VBA itself saves the processed external workbook.
        # Save the macro workbook as well in case any workbook state changed.
        $Workbook.Save()
    }
    finally {

        if ($Workbook) {
            try {
                $Workbook.Close($true)
            }
            catch {}
        }

        if ($Excel) {
            try {
                $Excel.Quit()
            }
            catch {}
        }

        if ($Workbook) {
            try {
                [Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                    $Workbook
                ) | Out-Null
            }
            catch {}
        }

        if ($Excel) {
            try {
                [Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                    $Excel
                ) | Out-Null
            }
            catch {}
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# ============================================================================
# VERIFY INPUT_CIF.TXT
# ============================================================================

function Assert-FreshInputCif {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [datetime]$StepStarted
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Input_cif.txt was not created: $Path"
    }

    $File = Get-Item -LiteralPath $Path

    if ($File.Length -le 0) {
        throw "Input_cif.txt is empty: $Path"
    }

    if ($File.LastWriteTime -lt $StepStarted) {
        throw @"
Input_cif.txt was not refreshed during this run.

File: $Path
Last write: $($File.LastWriteTime)
Step started: $StepStarted
"@
    }

    $Content = (
        Get-Content `
            -LiteralPath $Path `
            -Raw
    ).Trim()

    if ([string]::IsNullOrWhiteSpace($Content)) {
        throw "Input_cif.txt contains no CIF values."
    }

    $Cifs = @(
        $Content -split ',' |
        ForEach-Object {
            $_.Trim()
        } |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace($_)
        }
    )

    if ($Cifs.Count -eq 0) {
        throw "No usable CIF values found in Input_cif.txt."
    }

    Write-PipelineLog `
        "Input_cif.txt verified: $($Cifs.Count) CIF value(s)." `
        'OK'
}

# ============================================================================
# MAIN PIPELINE
# ============================================================================

try {

    Write-PipelineLog '============================================================'
    Write-PipelineLog 'STARTING DAILY PIPELINE'
    Write-PipelineLog "Report date: $($ReportDate.ToString('yyyy-MM-dd'))"
    Write-PipelineLog '============================================================'

    # ------------------------------------------------------------------------
    # STEP 1
    # ------------------------------------------------------------------------

    Write-PipelineLog 'STEP 1: TONG QUAN TTKH'

    # Delete stale parameter file.
    if (Test-Path -LiteralPath $InputCifPath) {

        Write-PipelineLog `
            'Removing previous Input_cif.txt.' `
            'WARN'

        Remove-Item `
            -LiteralPath $InputCifPath `
            -Force
    }

    $Step1Started = Get-Date

    # Download TONG QUAN
    Invoke-CognosReport `
        -ReportName 'TONG QUAN TTKH'

    # Process TONG QUAN using existing VBA
    Invoke-ExcelMacro `
        -MacroName 'handle_file_tong_quan'

    # Verify that Step 1 generated a new CIF list.
    Assert-FreshInputCif `
        -Path $InputCifPath `
        -StepStarted $Step1Started

    # ------------------------------------------------------------------------
    # STEP 2
    # ------------------------------------------------------------------------

    Write-PipelineLog 'STEP 2: TONG HOA LOI ICH ODS - DOWNLOAD'

    Invoke-CognosReport `
        -ReportName 'TONG HOA LOI ICH ODS'

    # ------------------------------------------------------------------------
    # STEP 3
    # ------------------------------------------------------------------------

    Write-PipelineLog 'STEP 3: SPDV - DOWNLOAD'

    Invoke-CognosReport `
        -ReportName 'SPDV'

    # ------------------------------------------------------------------------
    # STEP 4
    # ------------------------------------------------------------------------

    Write-PipelineLog 'STEP 4: TONG HOA LOI ICH ODS - PROCESS'

    Invoke-ExcelMacro `
        -MacroName 'handle_file_tong_hoa_loi_ich'

    # ------------------------------------------------------------------------
    # STEP 5
    # ------------------------------------------------------------------------

    Write-PipelineLog 'STEP 5: SPDV - PROCESS'

    Invoke-ExcelMacro `
        -MacroName 'handle_file_spdv'

    # ------------------------------------------------------------------------
    # COMPLETE
    # ------------------------------------------------------------------------

    $Elapsed = (Get-Date) - $PipelineStart

    Write-PipelineLog '============================================================'
    Write-PipelineLog 'PIPELINE COMPLETED SUCCESSFULLY' 'OK'
    Write-PipelineLog "Report date: $($ReportDate.ToString('yyyy-MM-dd'))" 'OK'
    Write-PipelineLog "Elapsed: $($Elapsed.ToString('hh\:mm\:ss'))" 'OK'
    Write-PipelineLog '============================================================'

    exit 0
}
catch {

    $Elapsed = (Get-Date) - $PipelineStart

    Write-PipelineLog '============================================================' 'ERROR'
    Write-PipelineLog 'PIPELINE FAILED' 'ERROR'
    Write-PipelineLog $_.Exception.Message 'ERROR'
    Write-PipelineLog "Elapsed: $($Elapsed.ToString('hh\:mm\:ss'))" 'ERROR'
    Write-PipelineLog '============================================================' 'ERROR'

    exit 1
}