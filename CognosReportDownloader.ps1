<#
.SYNOPSIS
    Tự động tải hàng loạt báo cáo IBM Cognos Analytics 11 qua giao diện REST RDS.

.DESCRIPTION
    - Lưu trữ tên đăng nhập/mật khẩu an toàn trong Windows Credential Manager.
    - Đọc cấu hình báo cáo từ tệp JSON (chuẩn UTF-8).
    - Tự động tính toán các token ngày động ({Yesterday}, {Today}, {Today-7d}) trong tham số và đường dẫn tệp.
    - Hỗ trợ nhiều máy chủ Cognos, nhiều định dạng, nhiều tham số prompt.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = '.\cognos-reports.json',
    [switch]$SetupCredential,
    [switch]$TestConnection
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($ConfigPath) -or $ConfigPath -eq '.\cognos-reports.json') {
    $ConfigPath = Join-Path $scriptDir 'cognos-reports.json'
}

# Nạp module dùng chung
. (Join-Path $scriptDir 'CognosCommon.ps1')

# -----------------------------------------------------------------------------
# Nạp Cấu hình
# -----------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    if ($SetupCredential) {
        throw "Không tìm thấy tệp cấu hình: $ConfigPath. Vui lòng tạo cấu hình trước."
    }
    throw "Không tìm thấy tệp cấu hình: $ConfigPath. Vui lòng chạy Manage-CognosConfig.ps1 hoặc CognosConfigGui.ps1 để tạo."
}

$config = Load-CognosConfig -Path $ConfigPath
if ($null -eq $config) {
    throw "Không thể phân tích định dạng tệp cấu hình: $ConfigPath"
}

# Khởi tạo hệ thống ghi log
$loggingConfig = Get-PropOrKey -Object $config -Name 'Logging'
Initialize-CognosLogging -LoggingConfig $loggingConfig -BaseDirectory $PSScriptRoot

$credentialTarget = Get-RequiredProperty -Object $config -Name 'CredentialTarget'

# -----------------------------------------------------------------------------
# Chế độ Cài đặt Tài khoản Xác thực
# -----------------------------------------------------------------------------

if ($SetupCredential) {
    $existing = Get-WindowsGenericCredential -Target $credentialTarget
    if ($existing) {
        Write-Log "Đã tồn tại tài khoản cho '$credentialTarget'. Thông tin mới sẽ ghi đè lên tài khoản cũ." 'WARN'
    }

    $credential = Get-Credential -Message "Nhập thông tin tài khoản Cognos để lưu vào Windows Credential Manager"
    if ($null -eq $credential) { throw 'Đã hủy thao tác cài đặt tài khoản.' }

    Set-WindowsGenericCredential -Target $credentialTarget -Username $credential.UserName -Password $credential.Password
    Write-Log "Đã lưu tài khoản bảo mật vào Windows Credential Manager với target '$credentialTarget'." 'OK'
    exit 0
}

# -----------------------------------------------------------------------------
# Đọc Tài khoản từ Windows Credential Manager
# -----------------------------------------------------------------------------

$stored = Get-WindowsGenericCredential -Target $credentialTarget

if ($null -eq $stored) {
    Write-Log "Chưa tìm thấy thông tin tài khoản cho '$credentialTarget'." 'WARN'
    $credential = Get-Credential -Message "Nhập thông tin tài khoản Cognos (sẽ được lưu vào Windows Credential Manager)"
    if ($null -eq $credential) { throw 'Đã hủy thao tác nhập tài khoản.' }

    Set-WindowsGenericCredential -Target $credentialTarget -Username $credential.UserName -Password $credential.Password
    $stored = [pscustomobject]@{ Username = $credential.UserName; Password = $credential.Password }
    Write-Log 'Đã lưu thông tin tài khoản thành công.' 'OK'
}

$httpSettings = Get-CognosHttpSettings -Config $config
$retryPolicy = Get-CognosRetryPolicy -Config $config

$allInstances = Get-CognosInstances -Config $config
if (@($allInstances.Keys).Count -eq 0) {
    throw "Không có máy chủ Cognos nào được cấu hình trong $ConfigPath."
}

$sessions = @{}

function Get-OrCreateCognosSession {
    param(
        [Parameter(Mandatory)] $Instance,
        [Parameter(Mandatory)] $StoredCredential
    )

    $key = "$($Instance.CognosBaseUrl)|$($Instance.Namespace)"
    if ($sessions.ContainsKey($key)) {
        return $sessions[$key]
    }

    Write-Log "Đang kết nối tới máy chủ Cognos [$($Instance.Name)] tại $($Instance.CognosBaseUrl)..."
    $ctx = New-CognosHttpClient -TimeoutMinutes $httpSettings.TimeoutMinutes
    $xsrf = Invoke-CognosLogin `
        -Context $ctx `
        -CognosBaseUrl $Instance.CognosBaseUrl `
        -Namespace $Instance.Namespace `
        -Username $StoredCredential.Username `
        -Password $StoredCredential.Password

    Write-Log "Đăng nhập thành công với tài khoản $($StoredCredential.Username) trên [$($Instance.Name)] (Namespace: '$($Instance.Namespace)')." 'OK'
    
    $session = [pscustomobject]@{
        Context  = $ctx
        Xsrf     = $xsrf
        Instance = $Instance
    }
    $sessions[$key] = $session
    return $session
}

try {
    # -------------------------------------------------------------------------
    # Chế độ Kiểm tra Kết nối
    # -------------------------------------------------------------------------

    if ($TestConnection) {
        Write-Log "Bắt đầu kiểm tra kết nối tới tất cả ($(@($allInstances.Keys).Count)) máy chủ Cognos đã cấu hình..."
        foreach ($instName in $allInstances.Keys) {
            $inst = $allInstances[$instName]
            $null = Get-OrCreateCognosSession -Instance $inst -StoredCredential $stored
        }
        Write-Log 'Tất cả các kiểm tra kết nối và xác thực Cognos đã hoàn tất thành công.' 'OK'
        exit 0
    }

    # -------------------------------------------------------------------------
    # Kiểm tra Tính hợp lệ của Danh sách Báo cáo
    # -------------------------------------------------------------------------

    $reportsProp = Get-PropOrKey -Object $config -Name 'Reports'
    if ($null -eq $reportsProp -or @($reportsProp).Count -eq 0) {
        throw 'Không có báo cáo nào được định nghĩa trong tệp cấu hình.'
    }

    $executionResults = [System.Collections.Generic.List[object]]::new()

    # -------------------------------------------------------------------------
    # Tiến trình Tải Báo cáo Hàng loạt
    # -------------------------------------------------------------------------

    foreach ($report in @($reportsProp)) {
        $repName = Get-PropOrKey -Object $report -Name 'Name'
        $repSource = Get-PropOrKey -Object $report -Name 'Source'
        $reportName = if (-not [string]::IsNullOrWhiteSpace($repName)) { [string]$repName } else { [string]$repSource }

        $repEnabled = Get-PropOrKey -Object $report -Name 'Enabled'
        if ($null -ne $repEnabled -and $repEnabled -eq $false) {
            Write-Log "Bỏ qua báo cáo đang tắt: $reportName" 'WARN'
            continue
        }

        $formatsProp = Get-PropOrKey -Object $report -Name 'Formats'
        if ($null -eq $formatsProp -or @($formatsProp).Count -eq 0) {
            Write-Log "Báo cáo '$reportName' chưa cấu hình định dạng xuất." 'WARN'
            continue
        }

        # Tìm máy chủ tương ứng của báo cáo
        $reportInstance = Get-CognosReportInstance -Config $config -Report $report

        foreach ($formatConfig in @($formatsProp)) {
            $format = Get-RequiredProperty -Object $formatConfig -Name 'Format'
            $rawOutputPath = Get-RequiredProperty -Object $formatConfig -Name 'OutputPath'

            # -----------------------------------------------------------------
            # Tính toán Đường dẫn Tệp Động (Dynamic Token Resolution)
            # -----------------------------------------------------------------
            $outputPath = Resolve-DynamicTokens -Text $rawOutputPath -Report $report -Format $format

            # Tải báo cáo với cơ chế tự động thử lại (Retry with Backoff theo cấu hình JSON)
            $res = Invoke-CognosReportDownloadWithRetry `
                -GetSessionScript { Get-OrCreateCognosSession -Instance $reportInstance -StoredCredential $stored } `
                -BaseUrl $reportInstance.CognosBaseUrl `
                -Report $report `
                -Format $format `
                -OutputPath $outputPath `
                -InstanceName $reportInstance.Name `
                -MaxRetries $retryPolicy.MaxRetries `
                -InitialDelaySeconds $retryPolicy.InitialDelaySeconds `
                -BackoffMultiplier $retryPolicy.BackoffMultiplier

            $executionResults.Add($res)
        }
    }

    # -------------------------------------------------------------------------
    # Tổng kết Thực thi (Execution Summary Table & LatestRun.json)
    # -------------------------------------------------------------------------

    $logDir = if ($config.Logging -and $config.Logging.LogDirectory) { $config.Logging.LogDirectory } else { '.\Logs' }
    Write-ExecutionSummaryReport -Results @($executionResults) -LogDirectory $logDir

    $failedCount = @($executionResults | Where-Object { $_.Status -ne 'SUCCESS' }).Count
    if ($failedCount -gt 0) {
        exit 1
    }
}
finally {
    if ($null -ne $sessions) {
        foreach ($s in $sessions.Values) {
            if ($null -ne $s.Context) {
                if ($null -ne $s.Context.Client)  { $s.Context.Client.Dispose() }
                if ($null -ne $s.Context.Handler) { $s.Context.Handler.Dispose() }
            }
        }
    }
}