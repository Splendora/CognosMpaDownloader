<#
.SYNOPSIS
    Công cụ dòng lệnh tương tác (CLI) quản lý cấu hình cognos-reports.json.
    Hỗ trợ môi trường đa máy chủ Cognos, dùng chung tài khoản, gợi ý tham số ngày thông minh và xem trước đường dẫn tệp.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = '.\cognos-reports.json',
    [switch]$Gui
)

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($ConfigPath) -or $ConfigPath -eq '.\cognos-reports.json') {
    $ConfigPath = Join-Path $scriptDir 'cognos-reports.json'
}

if ($Gui) {
    & (Join-Path $scriptDir 'CognosConfigGui.ps1') -ConfigPath $ConfigPath
    exit 0
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Nạp module dùng chung
. (Join-Path $scriptDir 'CognosCommon.ps1')

# -----------------------------------------------------------------------------
# Hàm Trợ giúp Phiên Kết nối (Session Helper)
# -----------------------------------------------------------------------------

function Connect-CognosSession {
    param(
        [Parameter(Mandatory)] $Config,
        $Instance = $null
    )

    $allInstances = Get-CognosInstances -Config $Config
    if (@($allInstances.Keys).Count -eq 0) {
        throw "Không có máy chủ Cognos nào được cấu hình trong $ConfigPath."
    }

    $targetInstance = $Instance
    if ($null -eq $targetInstance) {
        if (@($allInstances.Keys).Count -eq 1) {
            foreach ($k in $allInstances.Keys) { $targetInstance = $allInstances[$k] }
        } else {
            Write-Host "`nChọn máy chủ Cognos để kết nối:" -ForegroundColor Cyan
            $instKeys = @($allInstances.Keys)
            for ($i = 0; $i -lt @($instKeys).Count; $i++) {
                $k = $instKeys[$i]
                $inst = $allInstances[$k]
                Write-Host "  [$($i + 1)] $k ($($inst.CognosBaseUrl) - $($inst.Namespace))"
            }
            $choice = Read-Host "Chọn máy chủ [1-$(@($instKeys).Count), mặc định: 1]"
            $idx = if ([int]::TryParse($choice, [ref]$null) -and [int]$choice -ge 1 -and [int]$choice -le @($instKeys).Count) { [int]$choice - 1 } else { 0 }
            $targetInstance = $allInstances[$instKeys[$idx]]
        }
    }

    $credTarget = Get-RequiredProperty -Object $Config -Name 'CredentialTarget'

    $cred = Get-WindowsGenericCredential -Target $credTarget
    if ($null -eq $cred) {
        Write-Host "`n[XÁC THỰC] Chưa tìm thấy thông tin đăng nhập cho '$credTarget'. Vui lòng nhập:" -ForegroundColor Yellow
        $userCred = Get-Credential -Message "Tài khoản Cognos ($credTarget)"
        if ($null -eq $userCred) { throw "Đã hủy nhập thông tin tài khoản." }
        Set-WindowsGenericCredential -Target $credTarget -Username $userCred.UserName -Password $userCred.Password
        $cred = [pscustomobject]@{ Username = $userCred.UserName; Password = $userCred.Password }
    }

    Write-Host "Đang kết nối tới [$($targetInstance.Name)] tại $($targetInstance.CognosBaseUrl)..." -ForegroundColor Cyan
    $http = New-CognosHttpClient -TimeoutMinutes 5
    $xsrf = Invoke-CognosLogin -Context $http -CognosBaseUrl $targetInstance.CognosBaseUrl -Namespace $targetInstance.Namespace -Username $cred.Username -Password $cred.Password
    Write-Host "[OK] Đăng nhập thành công với tài khoản $($cred.Username) trên [$($targetInstance.Name)]." -ForegroundColor Green

    return [pscustomobject]@{ Http = $http; Xsrf = $xsrf; Instance = $targetInstance }
}

# -----------------------------------------------------------------------------
# Hàm Trợ giúp Nhập Tham số (Parameter Prompting Helper)
# -----------------------------------------------------------------------------

function Prompt-ForParameterValue {
    param(
        [string]$ParamName,
        [AllowEmptyString()][string]$CurrentVal = '',
        $ParamInfo = $null
    )

    $reqLabel = if ($null -ne $ParamInfo -and $ParamInfo.IsRequired) { "[BẮT BUỘC *]" } else { "[Tùy chọn]" }

    # 1. Nếu có danh sách lựa chọn từ máy chủ Cognos
    if ($null -ne $ParamInfo -and $ParamInfo.Choices -and @($ParamInfo.Choices).Count -gt 0) {
        $choices = @($ParamInfo.Choices)
        Write-Host "`n  Tham số '$ParamName' $reqLabel có $(@($choices).Count) lựa chọn từ máy chủ Cognos:" -ForegroundColor $(if ($null -ne $ParamInfo -and $ParamInfo.IsRequired) { 'Yellow' } else { 'Cyan' })
        $maxShow = [Math]::Min(@($choices).Count, 25)
        for ($i = 0; $i -lt $maxShow; $i++) {
            $c = $choices[$i]
            $disp = if ($c.Display -and $c.Display -ne $c.Use) { "$($c.Use) - $($c.Display)" } else { "$($c.Use)" }
            Write-Host "    [$($i + 1)] $disp"
        }
        if (@($choices).Count -gt 25) {
            Write-Host "    ... (còn $(@($choices).Count - 25) lựa chọn khác)" -ForegroundColor DarkGray
        }

        $defChoice = if ($ParamInfo.DefaultValue) {
            if ($ParamInfo.DefaultValue -is [System.Collections.IEnumerable] -and -not ($ParamInfo.DefaultValue -is [string])) {
                ($ParamInfo.DefaultValue | ForEach-Object { [string]$_ }) -join ', '
            } else {
                [string]$ParamInfo.DefaultValue
            }
        } elseif (-not [string]::IsNullOrWhiteSpace($CurrentVal)) {
            $CurrentVal
        } else { '' }

        Write-Host "    [T] Nhập giá trị tùy chỉnh / token"
        Write-Host "    [0] Để trống"

        $multiTag = if ($ParamInfo.IsMultiSelect) { " (Có thể nhập nhiều số cách nhau dấu phẩy VD: 1, 2)" } else { "" }
        $promptMsg = "  Chọn mục [1-$maxShow$multiTag" + (if ($defChoice) { ", Mặc định: $defChoice" } else { "" }) + "]"
        $ans = Read-Host $promptMsg

        if ([string]::IsNullOrWhiteSpace($ans)) { return $defChoice }
        if ($ans -ieq '0') { return '' }
        if ($ans -ieq 't') { return (Read-Host "  Nhập giá trị tùy chỉnh") }

        $tokens = @($ans -split '[,;\s]+' | Where-Object { $_ })
        $picked = [System.Collections.Generic.List[string]]::new()
        foreach ($tok in $tokens) {
            $idx = 0
            if ([int]::TryParse($tok, [ref]$idx) -and $idx -ge 1 -and $idx -le @($choices).Count) {
                $picked.Add($choices[$idx - 1].Use)
            } else {
                $picked.Add($tok)
            }
        }
        if (@($picked).Count -gt 0) {
            return ($picked -join ', ')
        }
        return $defChoice
    }

    # 2. Nếu là kiểu ngày tháng
    $isDateParam = ($ParamName -match 'date|ngay|time|denhan|quahan')
    
    if ($isDateParam) {
        $defaultOption = if (-not [string]::IsNullOrWhiteSpace($CurrentVal)) { $CurrentVal } else { '{Yesterday}' }
        Write-Host "`n  Tham số '$ParamName' $reqLabel là kiểu Ngày/Giờ:" -ForegroundColor $(if ($null -ne $ParamInfo -and $ParamInfo.IsRequired) { 'Yellow' } else { 'Cyan' })
        Write-Host "    [1] {Yesterday}   (Tính toán: $((Get-Date).AddDays(-1).ToString('yyyy-MM-dd')) - Ngày hôm qua)"
        Write-Host "    [2] {Today}       (Tính toán: $((Get-Date).ToString('yyyy-MM-dd')) - Ngày hôm nay)"
        Write-Host "    [3] {MonthStart}  (Ngày đầu tiên của tháng)"
        Write-Host "    [4] Nhập giá trị tùy chỉnh"
        Write-Host "    [5] Để trống"
        $choice = Read-Host "  Chọn định dạng mẫu [1-5, Mặc định: $defaultOption]"
        
        switch ($choice) {
            '1' { return '{Yesterday}' }
            '2' { return '{Today}' }
            '3' { return '{MonthStart}' }
            '4' { return (Read-Host "  Nhập giá trị tùy chỉnh") }
            '5' { return '' }
            default { return $defaultOption }
        }
    } else {
        $promptMsg = "  $ParamName $reqLabel" + (if (-not [string]::IsNullOrWhiteSpace($CurrentVal)) { " [Hiện tại: '$CurrentVal']" } else { " (nhấn Enter để để trống)" })
        $val = Read-Host $promptMsg
        if ([string]::IsNullOrWhiteSpace($val)) { return $CurrentVal }
        return $val
    }
}

# -----------------------------------------------------------------------------
# Các Chức năng Tương tác (Interactive Actions)
# -----------------------------------------------------------------------------

function Action-InitConfig {
    Write-Host "`n=== Khởi tạo Cấu hình Máy chủ Mới ===" -ForegroundColor Cyan
    $instName = Read-Host "Nhập tên máy chủ chính (VD: ODS, BIDV_Core) [mặc định: ODS]"
    if ([string]::IsNullOrWhiteSpace($instName)) { $instName = 'ODS' }

    $url = Read-Host "Nhập Cognos Base URL cho $instName (VD: http://10.53.153.173/ibmcognos/bi)"
    $ns  = Read-Host "Nhập Cognos Namespace (VD: BIDV)"
    $target = Read-Host "Nhập tên đối tượng xác thực Windows (Credential Target) [mặc định: COGNOS]"
    if ([string]::IsNullOrWhiteSpace($target)) { $target = 'COGNOS' }

    Write-Host "Nhập thông tin tài khoản đăng nhập Cognos để lưu bảo mật trong Windows Credential Manager:" -ForegroundColor Yellow
    $userCred = Get-Credential -Message "Tài khoản Cognos ($target)"
    if ($null -ne $userCred) {
        Set-WindowsGenericCredential -Target $target -Username $userCred.UserName -Password $userCred.Password
    }

    $instances = [ordered]@{}
    $instances[$instName] = [ordered]@{
        CognosBaseUrl = $url.TrimEnd('/')
        Namespace     = $ns.Trim()
    }

    $newConfig = [ordered]@{
        CredentialTarget = $target
        DefaultInstance  = $instName
        Instances        = $instances
        Logging          = [ordered]@{
            Enabled         = $true
            LogDirectory    = ".\\Logs"
            LogFileName     = "CognosDownloader_{yyyyMMdd}.log"
            LogLevel        = "INFO"
            RetentionDays   = 30
            AuditCsvEnabled = $true
            AuditCsvPath    = ".\\Logs\\Audit_{yyyyMM}.csv"
        }
        Reports          = @()
    }
    Save-CognosConfig -Path $ConfigPath -Config $newConfig
    Write-Host "[OK] Cấu hình đã được lưu vào: $ConfigPath" -ForegroundColor Green
    return $newConfig
}

function Action-ListReports {
    param($Config)
    $rawReports = Get-PropOrKey -Object $Config -Name 'Reports'
    $reports = if ($null -ne $rawReports) { @($rawReports) } else { @() }
    if (@($reports).Count -eq 0) {
        Write-Host "`nChưa có báo cáo nào được cấu hình." -ForegroundColor Yellow
        return
    }

    Write-Host "`n=== Danh sách Báo cáo đã Cấu hình ($(@($reports).Count)) ===" -ForegroundColor Cyan
    for ($i = 0; $i -lt @($reports).Count; $i++) {
        $rep = $reports[$i]
        $enabled = Get-PropOrKey -Object $rep -Name 'Enabled'
        $enabledStr = if ($null -ne $enabled -and $enabled -eq $false) { "[ĐÃ TẮT]" } else { "[ĐANG BẬT]" }
        $color = if ($enabledStr -eq '[ĐANG BẬT]') { 'Green' } else { 'DarkGray' }
        
        $instObj = try { Get-CognosReportInstance -Config $Config -Report $rep } catch { $null }
        $instTag = if ($instObj) { "[Máy chủ: $($instObj.Name)]" } else { "" }

        $repName = Get-PropOrKey -Object $rep -Name 'Name'
        $repSource = Get-PropOrKey -Object $rep -Name 'Source'
        Write-Host "[$($i + 1)] $enabledStr $instTag $repName (StoreID: $repSource)" -ForegroundColor $color

        # Hiển thị Tham số với giá trị tính toán thực tế
        $paramsObj = Get-PropOrKey -Object $rep -Name 'Parameters'
        if ($null -ne $paramsObj) {
            $paramProps = @(Get-ObjectKeyValuePairs -Object $paramsObj)
            if (@($paramProps).Count -gt 0) {
                Write-Host "    Tham số Prompt:" -ForegroundColor Gray
                foreach ($p in @($paramProps)) {
                    $rawVal = if ($null -ne $p.Value) {
                        if ($p.Value -is [System.Collections.IEnumerable] -and -not ($p.Value -is [string])) {
                            ($p.Value | ForEach-Object { [string]$_ }) -join ', '
                        } else {
                            [string]$p.Value
                        }
                    } else { '<TRỐNG>' }

                    try {
                        $evalArray = @(Resolve-DynamicTokenArray -Value $p.Value -Report $rep)
                        if ($evalArray.Length -gt 1) {
                            $sampleStr = ($evalArray | Select-Object -First 5) -join ', '
                            $moreStr = if ($evalArray.Length -gt 5) { "... (+$(($evalArray.Length - 5)) giá trị)" } else { '' }
                            Write-Host "      - $($p.Name) = $rawVal -> [Đã nạp $(($evalArray.Length)) giá trị: $sampleStr $moreStr]" -ForegroundColor Yellow
                        } elseif ($evalArray.Length -eq 1 -and $evalArray[0] -ne $rawVal -and $rawVal -ne '<TRỐNG>') {
                            Write-Host "      - $($p.Name) = $rawVal -> (Tính toán: $($evalArray[0]))" -ForegroundColor Yellow
                        } else {
                            Write-Host "      - $($p.Name) = $rawVal" -ForegroundColor Yellow
                        }
                    } catch {
                        Write-Host "      - $($p.Name) = $rawVal -> [Lỗi đọc tệp: $($_.Exception.Message)]" -ForegroundColor Red
                    }
                }
            } else {
                Write-Host "    Tham số Prompt: Không có" -ForegroundColor Gray
            }
        } else {
            Write-Host "    Tham số Prompt: Không có" -ForegroundColor Gray
        }

        # Hiển thị Định dạng & Xem trước Đường dẫn
        $fmtsObj = Get-PropOrKey -Object $rep -Name 'Formats'
        if ($null -ne $fmtsObj) {
            $formatList = @($fmtsObj)
            if (@($formatList).Count -gt 0) {
                Write-Host "    Định dạng:" -ForegroundColor Gray
                foreach ($f in $formatList) {
                    $rawPath = [string](Get-PropOrKey -Object $f -Name 'OutputPath')
                    $rawFmt  = [string](Get-PropOrKey -Object $f -Name 'Format')

                    Write-Host "      - $rawFmt -> $rawPath" -ForegroundColor White
                    if (-not [string]::IsNullOrWhiteSpace($rawPath)) {
                        $evaluatedPath = Resolve-DynamicTokens -Text $rawPath -Report $rep -Format $rawFmt
                        Write-Host "        Đường dẫn thực tế: $evaluatedPath" -ForegroundColor DarkCyan
                    }
                }
            }
        }
    }
}

function Action-AddReport {
    param($Config)
    $allInstances = Get-CognosInstances -Config $Config
    if (@($allInstances.Keys).Count -eq 0) {
        Write-Host "Không có máy chủ Cognos nào được cấu hình." -ForegroundColor Red
        return
    }

    # Chọn máy chủ nếu có nhiều máy chủ
    $selectedInstance = $null
    if (@($allInstances.Keys).Count -eq 1) {
        foreach ($k in $allInstances.Keys) { $selectedInstance = $allInstances[$k] }
    } else {
        Write-Host "`nChọn máy chủ Cognos cho báo cáo này:" -ForegroundColor Cyan
        $instKeys = @($allInstances.Keys)
        for ($i = 0; $i -lt @($instKeys).Count; $i++) {
            $k = $instKeys[$i]
            $inst = $allInstances[$k]
            Write-Host "  [$($i + 1)] $k ($($inst.CognosBaseUrl))"
        }
        $choice = Read-Host "Chọn máy chủ [1-$(@($instKeys).Count), mặc định: 1]"
        $idx = if ([int]::TryParse($choice, [ref]$null) -and [int]$choice -ge 1 -and [int]$choice -le @($instKeys).Count) { [int]$choice - 1 } else { 0 }
        $selectedInstance = $allInstances[$instKeys[$idx]]
    }

    $session = Connect-CognosSession -Config $Config -Instance $selectedInstance

    Write-Host "`n=== Thêm Báo cáo Mới vào [$($selectedInstance.Name)] ===" -ForegroundColor Cyan
    $source = Read-Host "Nhập Mã Báo cáo / StoreID (VD: i54414D93B29A4D2289C4E88469871644)"
    if ([string]::IsNullOrWhiteSpace($source)) { return }

    $name = Read-Host "Nhập Tên Báo cáo (tùy chọn, nhấn Enter để lấy mặc định)"
    if ([string]::IsNullOrWhiteSpace($name)) { $name = "Report_$source" }

    Write-Host "Đang dò tìm tham số prompt từ máy chủ Cognos [$($selectedInstance.Name)]..." -ForegroundColor Cyan
    $discovered = Get-CognosReportParameters -Context $session.Http -BaseUrl $selectedInstance.CognosBaseUrl -SourceType "report" -Source $source -Xsrf $session.Xsrf

    $params = [ordered]@{}
    if (@($discovered.Keys).Count -gt 0) {
        Write-Host "`nTìm thấy $(@($discovered.Keys).Count) tham số từ máy chủ Cognos:" -ForegroundColor Green
        foreach ($k in $discovered.Keys) {
            $pInfo = $discovered[$k]
            $params[$k] = Prompt-ForParameterValue -ParamName $k -ParamInfo $pInfo
        }
    } else {
        Write-Host "Báo cáo này không yêu cầu tham số prompt." -ForegroundColor Yellow
    }

    # Chọn định dạng
    Write-Host "`nChọn Định dạng Xuất Mặc định:"
    Write-Host "  [1] xlsxData (Excel Dữ liệu - dạng bảng thuần túy)"
    Write-Host "  [2] spreadsheetML (Excel XML - giữ nguyên định dạng mẫu biểu)"
    Write-Host "  [3] PDF"
    Write-Host "  [4] CSV"
    $fmtChoice = Read-Host "Chọn định dạng [1-4, mặc định: 1]"
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

    $defaultTemplate = "D:\CognosReports\{Yesterday:yyyyMMdd}_{ReportName}.${ext}"
    Write-Host "`nMẫu Đường dẫn Tệp Mặc định: $defaultTemplate" -ForegroundColor Cyan
    $pathInput = Read-Host "Nhập Đường dẫn Lưu (nhấn Enter để dùng mặc định)"
    $outputPath = if (-not [string]::IsNullOrWhiteSpace($pathInput)) { $pathInput.Trim() } else { $defaultTemplate }

    $newRep = [ordered]@{
        Name       = $name
        Instance   = $selectedInstance.Name
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

    $finalConfig = [ordered]@{}
    if ($Config.PSObject.Properties['CognosBaseUrl']) { $finalConfig['CognosBaseUrl'] = $Config.CognosBaseUrl }
    if ($Config.PSObject.Properties['Namespace']) { $finalConfig['Namespace'] = $Config.Namespace }
    if ($Config.PSObject.Properties['CredentialTarget']) { $finalConfig['CredentialTarget'] = $Config.CredentialTarget }
    if ($Config.PSObject.Properties['DefaultInstance']) { $finalConfig['DefaultInstance'] = $Config.DefaultInstance }
    if ($Config.PSObject.Properties['Instances']) { $finalConfig['Instances'] = $Config.Instances }
    if ($Config.PSObject.Properties['Logging']) { $finalConfig['Logging'] = $Config.Logging }
    $finalConfig['Reports'] = $list

    Save-CognosConfig -Path $ConfigPath -Config $finalConfig
    Write-Host "[OK] Cấu hình đã được lưu vào: $ConfigPath" -ForegroundColor Green
    Write-Host "[THÀNH CÔNG] Đã thêm báo cáo '$name' vào máy chủ [$($selectedInstance.Name)]." -ForegroundColor Green
}

function Action-EditReport {
    param($Config)
    $rawReports = Get-PropOrKey -Object $Config -Name 'Reports'
    $reports = if ($null -ne $rawReports) { @($rawReports) } else { @() }
    if (@($reports).Count -eq 0) {
        Write-Host "Không có báo cáo nào để sửa." -ForegroundColor Yellow
        return
    }

    Action-ListReports -Config $Config
    $choice = Read-Host "`nNhập số thứ tự báo cáo cần sửa [1-$(@($reports).Count)]"
    $idx = [int]$choice - 1
    if ($idx -lt 0 -or $idx -ge @($reports).Count) { return }

    $rep = $reports[$idx]
    $instObj = try { Get-CognosReportInstance -Config $Config -Report $rep } catch { $null }
    $instName = if ($instObj) { $instObj.Name } else { 'Mặc định' }

    $repName = Get-PropOrKey -Object $rep -Name 'Name'
    if ([string]::IsNullOrWhiteSpace($repName)) { $repName = Get-PropOrKey -Object $rep -Name 'Source' }
    $repEnabled = Get-PropOrKey -Object $rep -Name 'Enabled'
    $isCurrentlyEnabled = if ($null -ne $repEnabled -and $repEnabled -eq $false) { $false } else { $true }

    Write-Host "`nĐang sửa: $repName [Máy chủ: $instName]" -ForegroundColor Cyan
    Write-Host " [1] Sửa giá trị Tham số (Thiết lập {Yesterday}, {Today}, hoặc Tùy chỉnh)"
    Write-Host " [2] Bật/Tắt Trạng thái (Hiện tại: $(if ($isCurrentlyEnabled) { 'ĐANG BẬT' } else { 'ĐÃ TẮT' }))"
    Write-Host " [3] Sửa Mẫu Đường dẫn Tệp"
    Write-Host " [4] Đồng bộ lại Tham số từ máy chủ Cognos"
    Write-Host " [5] Thay đổi Máy chủ Cognos liên kết"
    Write-Host " [0] Hủy bỏ"
    $action = Read-Host "Chọn chức năng [0-5]"

    switch ($action) {
        '1' {
            $paramsObj = Get-PropOrKey -Object $rep -Name 'Parameters'
            if ($null -ne $paramsObj) {
                $paramProps = @(Get-ObjectKeyValuePairs -Object $paramsObj)
                if (@($paramProps).Count -gt 0) {
                    Write-Host "`nNhập giá trị tham số:" -ForegroundColor Cyan
                    foreach ($p in @($paramProps)) {
                        $newVal = Prompt-ForParameterValue -ParamName ($p.Name) -CurrentVal ([string]$p.Value)
                        Set-ObjectProperty -Object $paramsObj -Name $p.Name -Value $newVal
                    }
                } else {
                    Write-Host "Báo cáo này chưa cấu hình tham số nào." -ForegroundColor Yellow
                }
            } else {
                Write-Host "Báo cáo này chưa cấu hình tham số nào." -ForegroundColor Yellow
            }
        }
        '2' {
            $newStatus = -not $isCurrentlyEnabled
            Set-ObjectProperty -Object $rep -Name 'Enabled' -Value $newStatus
            Write-Host "Trạng thái báo cáo đã đổi thành: $(if ($newStatus) { 'ĐANG BẬT' } else { 'ĐÃ TẮT' })" -ForegroundColor Green
        }
        '3' {
            $fmtsObj = Get-PropOrKey -Object $rep -Name 'Formats'
            if ($null -ne $fmtsObj) {
                foreach ($f in @($fmtsObj)) {
                    $fFormat = Get-PropOrKey -Object $f -Name 'Format'
                    $fPath = Get-PropOrKey -Object $f -Name 'OutputPath'
                    Write-Host "`nCác Token hỗ trợ: {Yesterday:yyyyMMdd}, {Today:yyyyMMdd}, {ReportName}, {Instance}, {Format}" -ForegroundColor DarkGray
                    $newPath = Read-Host "Nhập mẫu đường dẫn mới cho $fFormat [Hiện tại: $fPath]"
                    if (-not [string]::IsNullOrWhiteSpace($newPath)) {
                        Set-ObjectProperty -Object $f -Name 'OutputPath' -Value $newPath
                    }
                }
            }
        }
        '4' {
            $session = Connect-CognosSession -Config $Config -Instance $instObj
            Write-Host "Đang làm mới tham số từ máy chủ [$($instObj.Name)]..." -ForegroundColor Cyan
            $srcType = Get-PropOrKey -Object $rep -Name 'SourceType'
            if ([string]::IsNullOrWhiteSpace($srcType)) { $srcType = 'report' }
            $srcVal = Get-RequiredProperty -Object $rep -Name 'Source'
            $discovered = Get-CognosReportParameters -Context $session.Http -BaseUrl $instObj.CognosBaseUrl -SourceType $srcType -Source $srcVal -Xsrf $session.Xsrf
            
            $merged = [ordered]@{}
            $paramsObj = Get-PropOrKey -Object $rep -Name 'Parameters'
            if ($null -ne $paramsObj) {
                foreach ($p in @(Get-ObjectKeyValuePairs -Object $paramsObj)) { $merged[$p.Name] = $p.Value }
            }
            foreach ($k in $discovered.Keys) {
                $pInfo = $discovered[$k]
                if (-not $merged.Contains($k)) {
                    $merged[$k] = Prompt-ForParameterValue -ParamName $k -ParamInfo $pInfo
                    Write-Host "  + Phát hiện tham số mới: $k = $($merged[$k])" -ForegroundColor Green
                }
            }
            Set-ObjectProperty -Object $rep -Name 'Parameters' -Value $merged
        }
        '5' {
            $allInstances = Get-CognosInstances -Config $Config
            Write-Host "`nChọn máy chủ Cognos mới cho báo cáo này:" -ForegroundColor Cyan
            $instKeys = @($allInstances.Keys)
            for ($i = 0; $i -lt @($instKeys).Count; $i++) {
                $k = $instKeys[$i]
                $inst = $allInstances[$k]
                Write-Host "  [$($i + 1)] $k ($($inst.CognosBaseUrl))"
            }
            $choice = Read-Host "Chọn máy chủ [1-$(@($instKeys).Count)]"
            if ([int]::TryParse($choice, [ref]$null) -and [int]$choice -ge 1 -and [int]$choice -le @($instKeys).Count) {
                $selectedKey = $instKeys[[int]$choice - 1]
                Set-ObjectProperty -Object $rep -Name 'Instance' -Value $selectedKey
                Write-Host "Đã liên kết báo cáo với máy chủ: $selectedKey" -ForegroundColor Green
            }
        }
        default { return }
    }

    Save-CognosConfig -Path $ConfigPath -Config $Config
    Write-Host "[OK] Cấu hình đã được lưu vào: $ConfigPath" -ForegroundColor Green
}

function Action-RemoveReport {
    param($Config)
    $rawReports = Get-PropOrKey -Object $Config -Name 'Reports'
    $reports = if ($null -ne $rawReports) { @($rawReports) } else { @() }
    if (@($reports).Count -eq 0) { return }

    Action-ListReports -Config $Config
    $choice = Read-Host "`nNhập số thứ tự báo cáo cần XÓA [1-$(@($reports).Count), 0 để Hủy]"
    $idx = [int]$choice - 1
    if ($idx -lt 0 -or $idx -ge @($reports).Count) { return }

    $targetRep = $reports[$idx]
    $delName = Get-PropOrKey -Object $targetRep -Name 'Name'
    if ([string]::IsNullOrWhiteSpace($delName)) { $delName = Get-PropOrKey -Object $targetRep -Name 'Source' }
    $confirm = Read-Host "Bạn có chắc chắn muốn xóa '$delName' không? (y/N)"
    if ($confirm -ieq 'y') {
        $list = [System.Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt @($reports).Count; $i++) {
            if ($i -ne $idx) { $list.Add($reports[$i]) }
        }
        Set-ObjectProperty -Object $Config -Name 'Reports' -Value $list
        Save-CognosConfig -Path $ConfigPath -Config $Config
        Write-Host "[OK] Cấu hình đã được lưu vào: $ConfigPath" -ForegroundColor Green
        Write-Host "[OK] Đã xóa báo cáo thành công." -ForegroundColor Green
    }
}

function Action-TestConnection {
    param($Config)
    $allInstances = Get-CognosInstances -Config $Config
    Write-Host "`nĐang kiểm tra kết nối tới $(@($allInstances.Keys).Count) máy chủ Cognos..." -ForegroundColor Cyan

    $credTarget = Get-RequiredProperty -Object $Config -Name 'CredentialTarget'
    $stored = Get-WindowsGenericCredential -Target $credTarget
    if ($null -eq $stored) {
        Write-Host "Chưa tìm thấy thông tin tài khoản trong Windows Credential Manager. Vui lòng nhập..." -ForegroundColor Yellow
        $userCred = Get-Credential -Message "Tài khoản Cognos ($credTarget)"
        if ($null -eq $userCred) { return }
        Set-WindowsGenericCredential -Target $credTarget -Username $userCred.UserName -Password $userCred.Password
        $stored = [pscustomobject]@{ Username = $userCred.UserName; Password = $userCred.Password }
    }

    foreach ($k in $allInstances.Keys) {
        $inst = $allInstances[$k]
        try {
            Write-Host "  Kiểm tra [$k] ($($inst.CognosBaseUrl))..." -NoNewline
            $http = New-CognosHttpClient -TimeoutMinutes 1
            $xsrf = Invoke-CognosLogin -Context $http -CognosBaseUrl $inst.CognosBaseUrl -Namespace $inst.Namespace -Username $stored.Username -Password $stored.Password
            $http.Client.Dispose()
            $http.Handler.Dispose()
            Write-Host " [THÀNH CÔNG]" -ForegroundColor Green
        }
        catch {
            Write-Host " [THẤT BẠI: $($_.Exception.Message)]" -ForegroundColor Red
        }
    }
}

function Action-ManageSchedule {
    param($Config)
    
    $taskName = 'CognosReportDownloader'
    $tInfo = Get-CognosScheduledTask -TaskName $taskName

    Write-Host "`n--- QUẢN LÝ LỊCH TỰ ĐỘNG (WINDOWS TASK SCHEDULER) ---" -ForegroundColor Cyan
    if ($null -ne $tInfo -and $tInfo.Exists) {
        Write-Host "  Tên tác vụ:       $($tInfo.TaskName)" -ForegroundColor Green
        Write-Host "  Trạng thái:       $($tInfo.State) (Loại: $($tInfo.ScheduleType), Giờ: $($tInfo.StartTime))" -ForegroundColor Green
        $lastRunStr = if ($tInfo.LastRunTime) { $tInfo.LastRunTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Chưa chạy' }
        Write-Host "  Lần chạy gần nhất: $lastRunStr (Mã kết quả: $($tInfo.LastTaskResult))"
        $nextRunStr = if ($tInfo.NextRunTime) { $tInfo.NextRunTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Không xác định' }
        Write-Host "  Lần chạy tiếp:    $nextRunStr"
    } else {
        Write-Host "  Chưa có lịch tự động cho '$taskName'." -ForegroundColor Yellow
    }

    Write-Host "`n [1] Tạo / Cập nhật lịch tự động (Daily / Weekday / Weekly / Hourly)"
    Write-Host " [2] Chạy thử tác vụ ngay (Trigger Run Now)"
    Write-Host " [3] Xóa lịch tự động (Unregister Task)"
    Write-Host " [4] Mở giao diện Windows Task Scheduler (taskschd.msc)"
    Write-Host " [0] Quay lại Menu chính"

    $subOpt = Read-Host "`nChọn thao tác [0-4]"
    switch ($subOpt) {
        '1' {
            $freqChoice = Read-Host "Chọn tần suất: [1] Hàng ngày (Daily), [2] Các ngày làm việc (Thứ 2-6), [3] Hàng tuần (Weekly), [4] Hàng giờ (Hourly) [Mặc định: 1]"
            $freq = switch ($freqChoice) {
                '2' { 'WEEKDAY' }
                '3' { 'WEEKLY' }
                '4' { 'HOURLY' }
                default { 'DAILY' }
            }
            $timeInput = Read-Host "Nhập giờ bắt đầu chạy [HH:mm, mặc định: 06:00]"
            if (-not $timeInput) { $timeInput = '06:00' }
            if ($timeInput -notmatch '^\d{1,2}:\d{2}$') {
                Write-Host "[LỖI] Định dạng giờ không hợp lệ (cần dạng HH:mm, VD: 06:00, 07:30)." -ForegroundColor Red
                return
            }

            $elevatedChoice = Read-Host "Chạy với quyền Administrator cao nhất (Highest Privileges)? (y/N)"
            $isElevated = ($elevatedChoice -ieq 'y')

            $scriptPath = Join-Path -Path $scriptDir -ChildPath 'CognosReportDownloader.ps1'
            try {
                Set-CognosScheduledTask -TaskName $taskName -ScriptPath $scriptPath -ConfigPath $ConfigPath -ScheduleType $freq -StartTime $timeInput -RunElevated $isElevated | Out-Null
                
                Set-ObjectProperty -Object $Config -Name 'Scheduling' -Value ([ordered]@{
                    TaskName     = $taskName
                    ScheduleType = $freq
                    StartTime    = $timeInput
                    RunElevated  = $isElevated
                })
                Save-CognosConfig -Path $ConfigPath -Config $Config
                Write-Host "[OK] Đã thiết lập thành công lịch tự động '$taskName' ($freq lúc $timeInput) và lưu vào JSON!" -ForegroundColor Green
            }
            catch {
                Write-Host "[LỖI] Thiết lập lịch thất bại: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        '2' {
            try {
                Start-CognosScheduledTask -TaskName $taskName | Out-Null
                Write-Host "[OK] Đã kích hoạt chạy tác vụ '$taskName' trong nền bởi Windows Task Scheduler." -ForegroundColor Green
            }
            catch {
                Write-Host "[LỖI] Kích hoạt tác vụ thất bại: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        '3' {
            $confirm = Read-Host "Bạn có chắc chắn muốn xóa lịch '$taskName' không? (y/N)"
            if ($confirm -ieq 'y') {
                Remove-CognosScheduledTask -TaskName $taskName | Out-Null
                Write-Host "[OK] Đã xóa lịch tự động '$taskName'." -ForegroundColor Green
            }
        }
        '4' {
            try { Start-Process "taskschd.msc" } catch { Write-Host "[LỖI] $($_.Exception.Message)" -ForegroundColor Red }
        }
        default { return }
    }
}

# -----------------------------------------------------------------------------
# Vòng Lặp Menu Chính (Main Menu Loop)
# -----------------------------------------------------------------------------

$config = Load-CognosConfig -Path $ConfigPath

if ($null -eq $config) {
    Write-Host "Không tìm thấy tệp cấu hình tại: $ConfigPath" -ForegroundColor Yellow
    $init = Read-Host "Bạn có muốn khởi tạo cấu hình mới ngay bây giờ không? (Y/n)"
    if ($init -ieq 'n') { exit 0 }
    $config = Action-InitConfig
}

$logCfg = if ($config.PSObject.Properties['Logging']) { $config.Logging } else { $null }
Initialize-CognosLogging -LoggingConfig $logCfg -BaseDirectory $scriptDir

while ($true) {
    $instMap = Get-CognosInstances -Config $config
    $instSummary = ($instMap.Keys | ForEach-Object { "$_ ($($instMap[$_].CognosBaseUrl))" }) -join ', '

    Write-Host "`n=======================================================" -ForegroundColor Cyan
    Write-Host "          TRÌNH QUẢN LÝ CẤU HÌNH COGNOS" -ForegroundColor Cyan
    Write-Host " Tệp cấu hình: $ConfigPath" -ForegroundColor DarkGray
    Write-Host " Target xác thực: $($config.CredentialTarget)" -ForegroundColor DarkGray
    Write-Host " Máy chủ:      $instSummary" -ForegroundColor DarkGray
    Write-Host "=======================================================" -ForegroundColor Cyan
    Write-Host " [1] Xem danh sách báo cáo & xem trước đường dẫn (Live Preview)"
    Write-Host " [2] Thêm báo cáo mới (Tự động dò tìm tham số & Chọn máy chủ)"
    Write-Host " [3] Sửa báo cáo hiện có (Tham số / Định dạng / Máy chủ)"
    Write-Host " [4] Xóa báo cáo"
    Write-Host " [5] Kiểm tra kết nối & Đăng nhập Cognos (Tất cả máy chủ)"
    Write-Host " [6] Cấu hình lại Máy chủ / Tài khoản xác thực"
    Write-Host " [7] Mở Giao diện Đồ họa (Windows Forms GUI)"
    Write-Host " [8] Chạy tải toàn bộ báo cáo ngay (Batch Downloader)"
    Write-Host " [9] Cấu hình Lịch Tự Động (Windows Task Scheduler)"
    Write-Host " [0] Thoát"
    Write-Host "-------------------------------------------------------"

    $opt = Read-Host "Chọn chức năng [0-9]"
    switch ($opt) {
        '1' { Action-ListReports -Config $config }
        '2' { Action-AddReport -Config $config; $config = Load-CognosConfig -Path $ConfigPath }
        '3' { Action-EditReport -Config $config; $config = Load-CognosConfig -Path $ConfigPath }
        '4' { Action-RemoveReport -Config $config; $config = Load-CognosConfig -Path $ConfigPath }
        '5' { Action-TestConnection -Config $config }
        '6' { $config = Action-InitConfig }
        '7' { & (Join-Path $scriptDir 'CognosConfigGui.ps1') -ConfigPath $ConfigPath; $config = Load-CognosConfig -Path $ConfigPath }
        '8' { & (Join-Path $scriptDir 'CognosReportDownloader.ps1') -ConfigPath $ConfigPath }
        '9' { Action-ManageSchedule -Config $config }
        '0' { Write-Host "Tạm biệt!"; exit 0 }
        default { Write-Host "Lựa chọn không hợp lệ." -ForegroundColor Yellow }
    }
}