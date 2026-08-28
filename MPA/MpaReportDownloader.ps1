#requires -Version 5.1
<#
.SYNOPSIS
    Tự động tải hàng loạt báo cáo phân tích OBIEE / MPA qua giao diện SOAP Web Services.

.DESCRIPTION
    - Quản lý tài khoản bảo mật trong Windows Credential Manager (mặc định target 'MaCB').
    - Nạp cấu hình báo cáo từ tệp JSON (chuẩn UTF-8).
    - Tự động phân giải các token động ({Yesterday}, {Today}, {MonthStart}, {ReportName}, {Username}) trong đường dẫn catalog và tệp xuất.
    - Hỗ trợ xuất nhiều định dạng (CSV, EXCEL2007, PDF, MHT) với cơ chế thử lại tự động (retry policy).
    - Ghi log chi tiết, tạo bảng tổng hợp tiến độ và tệp audit rolling hàng tháng.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = '.\mpa-reports.json',
    [switch]$SetupCredential,
    [switch]$TestConnection
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($ConfigPath) -or $ConfigPath -eq '.\mpa-reports.json') {
    $ConfigPath = Join-Path $scriptDir 'mpa-reports.json'
}

# Dot-source core module
. (Join-Path $scriptDir 'ObieeCommon.ps1')

# -----------------------------------------------------------------------------
# Nạp Cấu hình
# -----------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    if ($SetupCredential) {
        throw "Không tìm thấy tệp cấu hình: $ConfigPath. Vui lòng tạo cấu hình trước."
    }
    throw "Không tìm thấy tệp cấu hình: $ConfigPath. Vui lòng chạy Manage-MpaConfig.ps1 để tạo."
}

$config = Load-MpaConfig -Path $ConfigPath
if ($null -eq $config) {
    throw "Không thể phân tích định dạng tệp cấu hình: $ConfigPath"
}

# Khởi tạo hệ thống ghi log
$loggingConfig = Get-PropOrKey -Object $config -Name 'Logging'
Initialize-MpaLogging -LoggingConfig $loggingConfig -BaseDirectory $scriptDir

$credentialTarget = Get-RequiredProperty -Object $config -Name 'CredentialTarget'

# -----------------------------------------------------------------------------
# Chế độ Cài đặt Tài khoản Xác thực
# -----------------------------------------------------------------------------

if ($SetupCredential) {
    $existing = Get-WindowsGenericCredential -Target $credentialTarget
    if ($existing) {
        Write-Log "Đã tồn tại tài khoản cho '$credentialTarget'. Thông tin mới sẽ ghi đè lên tài khoản cũ." 'WARN'
    }

    $credential = Get-Credential -Message "Nhập thông tin tài khoản MPA / OBIEE để lưu vào Windows Credential Manager"
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
    $credential = Get-Credential -Message "Nhập thông tin tài khoản MPA / OBIEE (sẽ được lưu vào Windows Credential Manager)"
    if ($null -eq $credential) { throw 'Đã hủy thao tác nhập tài khoản.' }

    Set-WindowsGenericCredential -Target $credentialTarget -Username $credential.UserName -Password $credential.Password
    $stored = [pscustomobject]@{ Username = $credential.UserName; Password = $credential.Password }
    Write-Log 'Đã lưu thông tin tài khoản thành công.' 'OK'
}

$plainPassword = Get-PlainTextFromSecureString -SecureString $stored.Password
$httpSettings = Get-MpaHttpSettings -Config $config
$retryPolicy = Get-MpaRetryPolicy -Config $config

$allInstances = Get-MpaInstances -Config $config
if (@($allInstances.Keys).Count -eq 0) {
    throw "Không có máy chủ MPA nào được cấu hình trong $ConfigPath."
}

$sessions = @{}

function Get-OrCreateMpaSession {
    param(
        [Parameter(Mandatory)] [string]$InstanceName,
        [Parameter(Mandatory)] $InstanceConfig
    )

    if ($sessions.ContainsKey($InstanceName) -and -not [string]::IsNullOrWhiteSpace($sessions[$InstanceName])) {
        return $sessions[$InstanceName]
    }

    $baseUrl = Get-RequiredProperty -Object $InstanceConfig -Name 'MpaBaseUrl'
    Write-Log "Đang kết nối và khởi tạo phiên OBIEE SOAP trên máy chủ '$InstanceName' ($baseUrl)..." 'INFO'
    $sid = Connect-Obiee -BaseUrl $baseUrl -Username $stored.Username -Password $plainPassword
    $sessions[$InstanceName] = $sid
    Write-Log "Đã đăng nhập thành công vào '$InstanceName'. Session ID: $($sid.Substring(0, [math]::Min(12, $sid.Length)))..." 'OK'
    return $sid
}

# -----------------------------------------------------------------------------
# Kiểm tra Kết nối (-TestConnection)
# -----------------------------------------------------------------------------

if ($TestConnection) {
    Write-Log "Bắt đầu kiểm tra kết nối tới tất cả máy chủ MPA..." 'INFO'
    $allOk = $true

    foreach ($instName in $allInstances.Keys) {
        $instConfig = $allInstances[$instName]
        $baseUrl = Get-RequiredProperty -Object $instConfig -Name 'MpaBaseUrl'
        try {
            Write-Log "Kiểm tra xác thực trên '$instName' ($baseUrl)..." 'INFO'
            $sid = Connect-Obiee -BaseUrl $baseUrl -Username $stored.Username -Password $plainPassword
            Write-Log "Đăng nhập thành công '$instName'. Kiểm tra ngắt phiên..." 'OK'
            Disconnect-Obiee -BaseUrl $baseUrl -SessionId $sid | Out-Null
            Write-Log "Ngắt phiên thành công trên '$instName'." 'OK'
        }
        catch {
            Write-Log "Kiểm tra kết nối thất bại cho '$instName': $_" 'ERROR'
            $allOk = $false
        }
    }

    if ($allOk) {
        Write-Log "Tất cả máy chủ MPA hoạt động bình thường!" 'OK'
        exit 0
    }
    else {
        Write-Log "Một hoặc nhiều máy chủ MPA gặp sự cố kết nối." 'ERROR'
        exit 1
    }
}

# -----------------------------------------------------------------------------
# Tiến trình Tải Báo cáo Hàng loạt
# -----------------------------------------------------------------------------

$reports = Get-PropOrKey -Object $config -Name 'Reports'
if ($null -eq $reports -or @($reports).Count -eq 0) {
    Write-Log "Không tìm thấy danh sách báo cáo nào trong cấu hình: $ConfigPath" 'WARN'
    exit 0
}

$enabledReports = @($reports | Where-Object {
    $en = Get-PropOrKey -Object $_ -Name 'Enabled' -Default $true
    [bool]$en -eq $true
})

Write-Log "Tìm thấy $($enabledReports.Count)/$(@($reports).Count) báo cáo đang được kích hoạt." 'INFO'
if ($enabledReports.Count -eq 0) {
    Write-Log "Không có báo cáo nào được kích hoạt để tải." 'WARN'
    exit 0
}

try {
    foreach ($rep in $enabledReports) {
        $repName = Get-RequiredProperty -Object $rep -Name 'Name'
        $repPath = Get-RequiredProperty -Object $rep -Name 'Path'
        $instInfo = Get-MpaReportInstance -Config $config -Report $rep
        $instName = $instInfo.Name
        $instConfig = $instInfo.Config
        $baseUrl = Get-RequiredProperty -Object $instConfig -Name 'MpaBaseUrl'

        $formats = Get-PropOrKey -Object $rep -Name 'Formats'
        if ($null -eq $formats -or @($formats).Count -eq 0) {
            Write-Log "Báo cáo '$repName' không có định dạng tải nào được cấu hình." 'WARN'
            continue
        }

        # Resolve tokens in report path (e.g. /users/{Username}/_portal/)
        $resolvedCatalogPath = Resolve-DynamicTokens -Text $repPath -Report $rep
        $resolvedCatalogPath = $resolvedCatalogPath.Replace('{Username}', $stored.Username)

        foreach ($fmtConfig in $formats) {
            $format = Get-RequiredProperty -Object $fmtConfig -Name 'Format'
            $rawOutputPath = Get-RequiredProperty -Object $fmtConfig -Name 'OutputPath'

            $resolvedOutputPath = Resolve-DynamicTokens -Text $rawOutputPath -Report $rep -Format $format
            $resolvedOutputPath = $resolvedOutputPath.Replace('{Username}', $stored.Username)
            if (-not [System.IO.Path]::IsPathRooted($resolvedOutputPath)) {
                $resolvedOutputPath = Join-Path $scriptDir $resolvedOutputPath
            }

            try {
                $sessionId = Get-OrCreateMpaSession -InstanceName $instName -InstanceConfig $instConfig

                $renewCallback = {
                    $sessions.Remove($instName)
                    return (Get-OrCreateMpaSession -InstanceName $instName -InstanceConfig $instConfig)
                }

                $downloadResult = Invoke-MpaReportDownloadWithRetry `
                    -BaseUrl $baseUrl `
                    -Path $resolvedCatalogPath `
                    -Format $format `
                    -OutputPath $resolvedOutputPath `
                    -SessionId $sessionId `
                    -RetryPolicy $retryPolicy `
                    -Parameters (Get-PropOrKey -Object $rep -Name 'Parameters') `
                    -FilterExpressions (Get-PropOrKey -Object $rep -Name 'FilterExpressions') `
                    -TimeoutSeconds ($httpSettings.TimeoutMinutes * 60) `
                    -OnRenewSession $renewCallback

                if ($downloadResult.Success) {
                    Write-AuditLog `
                        -ReportName $repName `
                        -Path $resolvedCatalogPath `
                        -Format $format `
                        -Status 'SUCCESS' `
                        -DurationMs $downloadResult.DurationMs `
                        -FileSizeBytes $downloadResult.FileSizeBytes `
                        -OutputPath $resolvedOutputPath
                }
                else {
                    Write-AuditLog `
                        -ReportName $repName `
                        -Path $resolvedCatalogPath `
                        -Format $format `
                        -Status 'FAILED' `
                        -DurationMs $downloadResult.DurationMs `
                        -FileSizeBytes 0 `
                        -OutputPath $resolvedOutputPath `
                        -ErrorMessage $downloadResult.ErrorMessage
                }
            }
            catch {
                Write-Log "Lỗi nghiêm trọng khi xử lý báo cáo '$repName' ($format): $_" 'ERROR'
                Write-AuditLog `
                    -ReportName $repName `
                    -Path $resolvedCatalogPath `
                    -Format $format `
                    -Status 'ERROR' `
                    -DurationMs 0 `
                    -FileSizeBytes 0 `
                    -OutputPath $resolvedOutputPath `
                    -ErrorMessage $_.ToString()
            }
        }
    }
}
finally {
    # Clean up all active sessions on the server
    foreach ($instKey in @($sessions.Keys)) {
        $sid = $sessions[$instKey]
        if (-not [string]::IsNullOrWhiteSpace($sid)) {
            try {
                $instConfig = $allInstances[$instKey]
                $baseUrl = Get-RequiredProperty -Object $instConfig -Name 'MpaBaseUrl'
                Write-Log "Đang giải phóng phiên đăng nhập cho '$instKey'..." 'INFO'
                Disconnect-Obiee -BaseUrl $baseUrl -SessionId $sid | Out-Null
            }
            catch { }
        }
    }
    $sessions.Clear()

    # Render summary table
    Write-ExecutionSummaryReport -BaseDirectory $scriptDir
}
