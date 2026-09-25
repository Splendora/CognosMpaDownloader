#requires -Version 5.1
<#
.SYNOPSIS
    Công cụ dòng lệnh tương tác (CLI) quản lý cấu hình mpa-reports.json và duyệt Catalog OBIEE.
.DESCRIPTION
    - Quản lý danh sách báo cáo phân tích OBIEE / MPA, máy chủ, định dạng xuất và token đường dẫn.
    - Trình duyệt Catalog tương tác (Browse Web Catalog) theo cấu trúc thư mục người dùng và thư mục dùng chung (/shared).
    - Tải thử nghiệm báo cáo trực tiếp và xem trước đường dẫn tệp sau khi phân giải token.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = '.\mpa-reports.json',
    [switch]$Gui
)

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($ConfigPath) -or $ConfigPath -eq '.\mpa-reports.json') {
    $ConfigPath = Join-Path $scriptDir 'mpa-reports.json'
}

if ($Gui) {
    & (Join-Path $scriptDir 'MpaConfigGui.ps1') -ConfigPath $ConfigPath
    exit 0
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Nạp module dùng chung
. (Join-Path $scriptDir 'ObieeCommon.ps1')

# -----------------------------------------------------------------------------
# Trợ giúp Phiên Kết nối Tương tác
# -----------------------------------------------------------------------------

function Connect-MpaSessionInteractive {
    param(
        [Parameter(Mandatory)] $Config,
        $Instance = $null
    )

    $allInstances = Get-MpaInstances -Config $Config
    if (@($allInstances.Keys).Count -eq 0) {
        throw "Không có máy chủ MPA nào được cấu hình trong $ConfigPath."
    }

    $targetInstanceName = $null
    $targetInstanceConfig = $null

    if ($null -ne $Instance) {
        $targetInstanceName = $Instance.Name
        $targetInstanceConfig = $Instance.Config
    }
    else {
        if (@($allInstances.Keys).Count -eq 1) {
            foreach ($k in $allInstances.Keys) {
                $targetInstanceName = $k
                $targetInstanceConfig = $allInstances[$k]
            }
        } else {
            Write-Host "`nChọn máy chủ MPA để kết nối:" -ForegroundColor Cyan
            $instKeys = @($allInstances.Keys)
            for ($i = 0; $i -lt @($instKeys).Count; $i++) {
                $k = $instKeys[$i]
                $inst = $allInstances[$k]
                Write-Host "  [$($i + 1)] $k ($($inst.MpaBaseUrl))"
            }
            $choice = Read-Host "Chọn máy chủ [1-$(@($instKeys).Count), mặc định: 1]"
            $idx = if ([int]::TryParse($choice, [ref]$null) -and [int]$choice -ge 1 -and [int]$choice -le @($instKeys).Count) { [int]$choice - 1 } else { 0 }
            $targetInstanceName = $instKeys[$idx]
            $targetInstanceConfig = $allInstances[$targetInstanceName]
        }
    }

    $credTarget = Get-RequiredProperty -Object $Config -Name 'CredentialTarget'
    $cred = Get-WindowsGenericCredential -Target $credTarget
    if ($null -eq $cred) {
        Write-Host "`n[XÁC THỰC] Chưa tìm thấy thông tin đăng nhập cho '$credTarget'. Vui lòng nhập:" -ForegroundColor Yellow
        $userCred = Get-Credential -Message "Tài khoản MPA ($credTarget)"
        if ($null -eq $userCred) { throw "Đã hủy nhập thông tin tài khoản." }
        Set-WindowsGenericCredential -Target $credTarget -Username $userCred.UserName -Password $userCred.Password
        $cred = [pscustomobject]@{ Username = $userCred.UserName; Password = $userCred.Password }
    }

    $plainPassword = Get-PlainTextFromSecureString -SecureString $cred.Password
    $baseUrl = Get-RequiredProperty -Object $targetInstanceConfig -Name 'MpaBaseUrl'

    Write-Host "Đang kết nối tới [$targetInstanceName] tại $baseUrl..." -ForegroundColor Cyan
    $sid = Connect-Obiee -BaseUrl $baseUrl -Username $cred.Username -Password $plainPassword
    Write-Host "[OK] Đăng nhập thành công với tài khoản $($cred.Username) trên [$targetInstanceName]." -ForegroundColor Green

    return [pscustomobject]@{
        BaseUrl        = $baseUrl
        SessionId      = $sid
        InstanceName   = $targetInstanceName
        InstanceConfig = $targetInstanceConfig
        Username       = $cred.Username
    }
}

# -----------------------------------------------------------------------------
# Các Hành động Tương tác (Interactive Actions)
# -----------------------------------------------------------------------------

function Action-TestConnection {
    param($Config)
    Write-Host "`n=== [1] Kiểm tra Kết nối & Đăng nhập OBIEE SOAP ===" -ForegroundColor Cyan
    $sess = $null
    try {
        $sess = Connect-MpaSessionInteractive -Config $Config
        Write-Host "[+] Kết nối và lấy Session ID thành công: $($sess.SessionId)" -ForegroundColor Green
    }
    catch {
        Write-Host "[-] Lỗi kết nối: $_" -ForegroundColor Red
    }
    finally {
        if ($null -ne $sess -and -not [string]::IsNullOrWhiteSpace($sess.SessionId)) {
            Write-Host "Đang đăng xuất..." -ForegroundColor Gray
            Disconnect-Obiee -BaseUrl $sess.BaseUrl -SessionId $sess.SessionId | Out-Null
            Write-Host "[+] Đăng xuất thành công." -ForegroundColor Green
        }
    }
}

function Action-BrowseCatalog {
    param($Config)
    Write-Host "`n=== [2] Duyệt Cây Thư mục Web Catalog OBIEE ===" -ForegroundColor Cyan

    $sess = $null
    try {
        $sess = Connect-MpaSessionInteractive -Config $Config
        $currentPath = "/users/$($sess.Username)"

        while ($true) {
            Write-Host "`nĐang ở thư mục: $currentPath" -ForegroundColor Yellow
            $items = Get-ObieeCatalogItems -BaseUrl $sess.BaseUrl -Path $currentPath -SessionId $sess.SessionId
            
            if ($null -eq $items -or @($items).Count -eq 0) {
                Write-Host "  (Thư mục trống)" -ForegroundColor DarkGray
            }
            else {
                Show-ObieeCatalogItems -Items $items
            }

            Write-Host "`nTùy chọn duyệt:" -ForegroundColor Cyan
            Write-Host "  [1] Đi tới thư mục con / Chọn đối tượng"
            Write-Host "  [2] Lên một cấp thư mục (..)"
            Write-Host "  [3] Chuyển tới thư mục cá nhân (/users/$($sess.Username))"
            Write-Host "  [4] Chuyển tới thư mục dùng chung (/shared)"
            Write-Host "  [5] Nhập đường dẫn trực tiếp"
            Write-Host "  [0] Quay lại Menu chính"

            $choice = Read-Host "`nLựa chọn [0-5]"
            switch ($choice) {
                '1' {
                    $subName = Read-Host "Nhập tên hoặc đường dẫn thư mục/báo cáo muốn mở"
                    if (-not [string]::IsNullOrWhiteSpace($subName)) {
                        if ($subName.StartsWith('/')) {
                            $currentPath = $subName
                        } else {
                            $currentPath = "$currentPath/$subName"
                        }
                    }
                }
                '2' {
                    $parent = Split-Path -Parent $currentPath
                    if (-not [string]::IsNullOrWhiteSpace($parent) -and $parent -ne '/') {
                        $currentPath = $parent.Replace('\', '/')
                    } else {
                        $currentPath = '/'
                    }
                }
                '3' { $currentPath = "/users/$($sess.Username)" }
                '4' { $currentPath = "/shared" }
                '5' {
                    $custom = Read-Host "Nhập đường dẫn Catalog đầy đủ"
                    if (-not [string]::IsNullOrWhiteSpace($custom)) {
                        $currentPath = $custom
                    }
                }
                default { return }
            }
        }
    }
    catch {
        Write-Host "[-] Lỗi khi duyệt Catalog: $_" -ForegroundColor Red
    }
    finally {
        if ($null -ne $sess -and -not [string]::IsNullOrWhiteSpace($sess.SessionId)) {
            Disconnect-Obiee -BaseUrl $sess.BaseUrl -SessionId $sess.SessionId | Out-Null
        }
    }
}

function Action-InspectItem {
    param($Config)
    Write-Host "`n=== [3] Tra cứu Thông tin Chi tiết Đối tượng Catalog ===" -ForegroundColor Cyan
    $path = Read-Host "Nhập đường dẫn Catalog (VD: /users/{Username}/_portal/ hoặc /shared/...)"
    if ([string]::IsNullOrWhiteSpace($path)) { return }

    $sess = $null
    try {
        $sess = Connect-MpaSessionInteractive -Config $Config
        $resolvedPath = $path.Replace('{Username}', $sess.Username)

        $info = Get-ObieeItemInfo -BaseUrl $sess.BaseUrl -Path $resolvedPath -SessionId $sess.SessionId
        Write-Host "`n[+] Thông tin đối tượng:" -ForegroundColor Green
        Write-Host "    Caption   : $($info.Caption)" -ForegroundColor Cyan
        Write-Host "    Type      : $($info.Type)" -ForegroundColor Cyan
        Write-Host "    Signature : $($info.Signature)" -ForegroundColor Cyan
        Write-Host "    Path      : $($info.Path)" -ForegroundColor Gray
        Write-Host "    LastMod   : $($info.LastModified)" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "[-] Lỗi tra cứu: $_" -ForegroundColor Red
    }
    finally {
        if ($null -ne $sess -and -not [string]::IsNullOrWhiteSpace($sess.SessionId)) {
            Disconnect-Obiee -BaseUrl $sess.BaseUrl -SessionId $sess.SessionId | Out-Null
        }
    }
}

function Action-ListReports {
    param($Config)
    Write-Host "`n=== [4] Danh sách Báo cáo Đã Cấu hình ===" -ForegroundColor Cyan
    $reports = Get-PropOrKey -Object $Config -Name 'Reports'
    if ($null -eq $reports -or @($reports).Count -eq 0) {
        Write-Host "  (Chưa có báo cáo nào được cấu hình)" -ForegroundColor Yellow
        return
    }

    $idx = 1
    foreach ($r in $reports) {
        $enLabel = if ($r.Enabled) { "[BẬT]" } else { "[TẮT]" }
        $enColor = if ($r.Enabled) { 'Green' } else { 'DarkGray' }
        Write-Host "  [$idx] $enLabel $($r.Name) - Inst: $($r.Instance)" -ForegroundColor $enColor
        Write-Host "        Path: $($r.Path)" -ForegroundColor Gray

        $params = Get-PropOrKey -Object $r -Name 'Parameters'
        if ($null -ne $params) {
            $pairs = Get-ObjectKeyValuePairs -Object $params
            if ($pairs.Count -gt 0) {
                Write-Host "        Tham số Prompt:" -ForegroundColor Gray
                foreach ($p in $pairs) {
                    $rawVal = [string]$p.Value
                    try {
                        $evalArray = @(Resolve-DynamicTokenArray -Value $p.Value -Report $r -DefaultDateFormat 'MM/dd/yyyy')
                        if ($evalArray.Length -gt 1) {
                            $sampleStr = ($evalArray | Select-Object -First 5) -join ', '
                            Write-Host "          - $($p.Name) = $rawVal -> [Đã nạp $(($evalArray.Length)) giá trị: $sampleStr]" -ForegroundColor Yellow
                        } elseif ($evalArray.Length -eq 1 -and $evalArray[0] -ne $rawVal) {
                            Write-Host "          - $($p.Name) = $rawVal -> (Tính toán: $($evalArray[0]))" -ForegroundColor Yellow
                        } else {
                            Write-Host "          - $($p.Name) = $rawVal" -ForegroundColor Yellow
                        }
                    } catch {
                        Write-Host "          - $($p.Name) = $rawVal" -ForegroundColor Yellow
                    }
                }
            }
        }

        $filters = Get-PropOrKey -Object $r -Name 'FilterExpressions'
        if ($null -ne $filters -and @($filters).Count -gt 0) {
            Write-Host "        Điều kiện lọc (FilterExpressions):" -ForegroundColor Gray
            foreach ($flt in @($filters)) {
                Write-Host "          - $flt" -ForegroundColor DarkYellow
            }
        }

        foreach ($f in $r.Formats) {
            Write-Host "        - Định dạng: $($f.Format) => $($f.OutputPath)" -ForegroundColor DarkCyan
        }
        $idx++
    }
}

function Action-AddReport {
    param($Config)
    Write-Host "`n=== [5] Thêm Báo cáo Mới vào Cấu hình ===" -ForegroundColor Cyan

    $name = Read-Host "Nhập tên gợi nhớ của báo cáo"
    if ([string]::IsNullOrWhiteSpace($name)) { return }

    $allInst = Get-MpaInstances -Config $Config
    $instName = Get-PropOrKey -Object $Config -Name 'DefaultInstance' -Default 'MPA'
    if (@($allInst.Keys).Count -gt 1) {
        Write-Host "Chọn máy chủ:"
        $keys = @($allInst.Keys)
        for ($i = 0; $i -lt @($keys).Count; $i++) { Write-Host "  [$($i+1)] $($keys[$i])" }
        $c = Read-Host "Chọn [1-$(@($keys).Count), mặc định: 1]"
        $instName = if ([int]::TryParse($c, [ref]$null) -and [int]$c -ge 1 -and [int]$c -le @($keys).Count) { $keys[[int]$c - 1] } else { $keys[0] }
    }

    $path = Read-Host "Nhập đường dẫn Catalog (VD: /users/{Username}/_portal/ hoặc /shared/Finance/Report1)"
    if ([string]::IsNullOrWhiteSpace($path)) { return }

    $params = [ordered]@{}
    $askInspect = Read-Host "`nBạn có muốn tự động dò tìm tham số từ máy chủ không? [y/N]"
    if ($askInspect -eq 'y' -or $askInspect -eq 'Y') {
        $sess = $null
        try {
            $sess = Connect-MpaSessionInteractive -Config $Config -Instance (@{ Name = $instName; Config = $allInst[$instName] })
            $resolvedPath = Resolve-DynamicTokens -Text $path
            $resolvedPath = $resolvedPath.Replace('{Username}', $sess.Username)
            Write-Host "Đang quét tham số và biến lọc từ báo cáo '$resolvedPath'..." -ForegroundColor Cyan
            $detected = Get-ObieeReportParameters -BaseUrl $sess.BaseUrl -Path $resolvedPath -SessionId $sess.SessionId
            if ($null -ne $detected -and @($detected).Count -gt 0) {
                Write-Host "Đã phát hiện $(@($detected).Count) tham số:" -ForegroundColor Green
                foreach ($dp in $detected) {
                    $val = Read-Host "  Giá trị cho '$($dp.Name)' [Mặc định: $($dp.DefaultValue)]"
                    if ([string]::IsNullOrWhiteSpace($val)) { $val = $dp.DefaultValue }
                    $params[$dp.Name] = $val
                }
            } else {
                Write-Host "Không phát hiện tham số prompted nào trong báo cáo." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "[-] Không thể dò tìm tham số từ server: $_" -ForegroundColor Red
        } finally {
            if ($null -ne $sess -and -not [string]::IsNullOrWhiteSpace($sess.SessionId)) {
                Disconnect-Obiee -BaseUrl $sess.BaseUrl -SessionId $sess.SessionId | Out-Null
            }
        }
    }

    Write-Host "`nNhập thêm tham số thủ công (Enter để bỏ qua):" -ForegroundColor Gray
    while ($true) {
        $pInput = Read-Host "Nhập tham số dạng 'TenThamSo=GiaTri' (hoặc nhấn Enter để tiếp tục)"
        if ([string]::IsNullOrWhiteSpace($pInput)) { break }
        if ($pInput -match '^([^=]+)=(.*)$') {
            $params[$matches[1].Trim()] = $matches[2].Trim()
        } else {
            Write-Host "Định dạng không hợp lệ, vui lòng nhập dạng Key=Value." -ForegroundColor Yellow
        }
    }

    $filtersList = @()
    $filterInput = Read-Host "`nNhập điều kiện lọc FilterExpressions nếu có (phân tách dấu chấm phẩy, hoặc Enter để bỏ qua)"
    if (-not [string]::IsNullOrWhiteSpace($filterInput)) {
        $filtersList = @($filterInput.Split(';') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    Write-Host "`nChọn định dạng xuất chính:" -ForegroundColor Cyan
    Write-Host "  [1] CSV (Bảng dữ liệu phân tách dấu phẩy)"
    Write-Host "  [2] EXCEL2007 (Tệp Excel .xlsx)"
    Write-Host "  [3] PDF (Tệp tài liệu PDF)"
    Write-Host "  [4] MHT (Web Archive HTML)"
    $fmtChoice = Read-Host "Chọn định dạng [1-4, Mặc định: 1]"
    $fmt = switch ($fmtChoice) {
        '2' { 'EXCEL2007' }
        '3' { 'PDF' }
        '4' { 'MHT' }
        default { 'CSV' }
    }

    $ext = switch ($fmt) {
        'EXCEL2007' { 'xlsx' }
        'PDF' { 'pdf' }
        'MHT' { 'mht' }
        default { 'csv' }
    }

    $defaultOut = ".\\Reports\\{Yesterday:yyyyMMdd}_{ReportName}.$ext"
    $outPath = Read-Host "Nhập đường dẫn lưu tệp [Mặc định: $defaultOut]"
    if ([string]::IsNullOrWhiteSpace($outPath)) { $outPath = $defaultOut }

    $repObj = [ordered]@{
        Name       = $name
        Instance   = $instName
        Path       = $path
        Enabled    = $true
        Parameters = $params
        Formats    = @(
            [pscustomobject]@{
                Format     = $fmt
                OutputPath = $outPath
            }
        )
    }
    if ($filtersList.Count -gt 0) {
        $repObj['FilterExpressions'] = $filtersList
    }

    $newReport = [pscustomobject]$repObj

    $currentReports = New-Object System.Collections.Generic.List[object]
    $existing = Get-PropOrKey -Object $Config -Name 'Reports'
    if ($null -ne $existing) {
        foreach ($r in $existing) { $currentReports.Add($r) }
    }
    $currentReports.Add($newReport)

    $Config.Reports = $currentReports.ToArray()
    Save-MpaConfig -Path $ConfigPath -Config $Config
    Write-Host "`n[+] Đã thêm báo cáo '$name' thành công và lưu cấu hình." -ForegroundColor Green
}

function Action-ToggleOrEditReport {
    param($Config)
    Write-Host "`n=== [6] Bật / Tắt hoặc Chỉnh sửa Báo cáo ===" -ForegroundColor Cyan
    $reports = Get-PropOrKey -Object $Config -Name 'Reports'
    if ($null -eq $reports -or @($reports).Count -eq 0) {
        Write-Host "Chưa có báo cáo nào." -ForegroundColor Yellow
        return
    }

    Action-ListReports -Config $Config
    $sel = Read-Host "`nChọn số thứ tự báo cáo cần chỉnh sửa [1-$(@($reports).Count)]"
    if (-not [int]::TryParse($sel, [ref]$null) -or [int]$sel -lt 1 -or [int]$sel -gt @($reports).Count) { return }

    $rep = $reports[[int]$sel - 1]
    Write-Host "`nChỉnh sửa: $($rep.Name)" -ForegroundColor Yellow
    Write-Host "  [1] Bật / Tắt kích hoạt (Hiện tại: $(if ($rep.Enabled) { 'BẬT' } else { 'TẮT' }))"
    Write-Host "  [2] Sửa tên hoặc đường dẫn Catalog"
    Write-Host "  [3] Sửa đường dẫn lưu tệp đầu ra"
    Write-Host "  [4] Chỉnh sửa / Thêm Tham số (Parameters)"
    Write-Host "  [5] Tự động dò tìm tham số từ máy chủ (Get-ObieeReportParameters)"
    Write-Host "  [6] Chỉnh sửa Điều kiện lọc (FilterExpressions)"
    Write-Host "  [7] Xóa báo cáo khỏi cấu hình"
    Write-Host "  [0] Hủy bỏ"

    $subC = Read-Host "Lựa chọn [0-7]"
    switch ($subC) {
        '1' {
            $rep.Enabled = -not $rep.Enabled
            Save-MpaConfig -Path $ConfigPath -Config $Config
            Write-Host "[+] Đã đổi trạng thái sang: $(if ($rep.Enabled) { 'BẬT' } else { 'TẮT' })" -ForegroundColor Green
        }
        '2' {
            $newName = Read-Host "Tên mới [Hiện tại: $($rep.Name)]"
            if (-not [string]::IsNullOrWhiteSpace($newName)) { Set-ObjectProperty -Object $rep -Name 'Name' -Value $newName }
            $newPath = Read-Host "Đường dẫn Catalog mới [Hiện tại: $($rep.Path)]"
            if (-not [string]::IsNullOrWhiteSpace($newPath)) { Set-ObjectProperty -Object $rep -Name 'Path' -Value $newPath }
            Save-MpaConfig -Path $ConfigPath -Config $Config
            Write-Host "[+] Đã cập nhật thông tin báo cáo." -ForegroundColor Green
        }
        '3' {
            if ($rep.Formats -and @($rep.Formats).Count -gt 0) {
                $f = $rep.Formats[0]
                $newOut = Read-Host "Đường dẫn lưu tệp mới [Hiện tại: $($f.OutputPath)]"
                if (-not [string]::IsNullOrWhiteSpace($newOut)) {
                    $f.OutputPath = $newOut
                    Save-MpaConfig -Path $ConfigPath -Config $Config
                    Write-Host "[+] Đã cập nhật đường dẫn lưu tệp." -ForegroundColor Green
                }
            }
        }
        '4' {
            $pObj = Get-PropOrKey -Object $rep -Name 'Parameters'
            $pDict = [ordered]@{}
            if ($null -ne $pObj) {
                $pairs = Get-ObjectKeyValuePairs -Object $pObj
                foreach ($pair in $pairs) { $pDict[$pair.Name] = $pair.Value }
            }
            Write-Host "`nCác tham số hiện tại:" -ForegroundColor Cyan
            foreach ($k in @($pDict.Keys)) {
                $newV = Read-Host "  $k [Hiện tại: $($pDict[$k])] (Nhấn Enter giữ nguyên, gõ 'DELETE' để xóa)"
                if ($newV -eq 'DELETE') {
                    $pDict.Remove($k)
                } elseif (-not [string]::IsNullOrWhiteSpace($newV)) {
                    $pDict[$k] = $newV
                }
            }
            Write-Host "Nhập thêm tham số mới (hoặc Enter để kết thúc):" -ForegroundColor Gray
            while ($true) {
                $pInput = Read-Host "Nhập dạng 'TenThamSo=GiaTri'"
                if ([string]::IsNullOrWhiteSpace($pInput)) { break }
                if ($pInput -match '^([^=]+)=(.*)$') {
                    $pDict[$matches[1].Trim()] = $matches[2].Trim()
                }
            }
            Set-ObjectProperty -Object $rep -Name 'Parameters' -Value $pDict
            Save-MpaConfig -Path $ConfigPath -Config $Config
            Write-Host "[+] Đã cập nhật tham số thành công." -ForegroundColor Green
        }
        '5' {
            $instInfo = Get-MpaReportInstance -Config $Config -Report $rep
            $sess = $null
            try {
                $sess = Connect-MpaSessionInteractive -Config $Config -Instance $instInfo
                $resolvedPath = Resolve-DynamicTokens -Text $rep.Path -Report $rep
                $resolvedPath = $resolvedPath.Replace('{Username}', $sess.Username)
                Write-Host "Đang dò tìm tham số từ máy chủ..." -ForegroundColor Cyan
                $detected = Get-ObieeReportParameters -BaseUrl $sess.BaseUrl -Path $resolvedPath -SessionId $sess.SessionId
                if ($null -ne $detected -and @($detected).Count -gt 0) {
                    $pObj = Get-PropOrKey -Object $rep -Name 'Parameters'
                    $pDict = [ordered]@{}
                    if ($null -ne $pObj) {
                        foreach ($pair in (Get-ObjectKeyValuePairs -Object $pObj)) { $pDict[$pair.Name] = $pair.Value }
                    }
                    $added = 0
                    foreach ($param in $detected) {
                        if (-not $pDict.Contains($param.Name)) {
                            $val = Read-Host "Tìm thấy tham số '$($param.Name)'. Giá trị [Mặc định: $($param.DefaultValue)]"
                            if ([string]::IsNullOrWhiteSpace($val)) { $val = $param.DefaultValue }
                            $pDict[$param.Name] = $val
                            $added++
                        }
                    }
                    Set-ObjectProperty -Object $rep -Name 'Parameters' -Value $pDict
                    Save-MpaConfig -Path $ConfigPath -Config $Config
                    Write-Host "[+] Đã thêm $added tham số mới từ server." -ForegroundColor Green
                } else {
                    Write-Host "Không tìm thấy tham số nào trên server." -ForegroundColor Yellow
                }
            } catch {
                Write-Host "[-] Lỗi dò tìm tham số: $_" -ForegroundColor Red
            } finally {
                if ($null -ne $sess -and -not [string]::IsNullOrWhiteSpace($sess.SessionId)) {
                    Disconnect-Obiee -BaseUrl $sess.BaseUrl -SessionId $sess.SessionId | Out-Null
                }
            }
        }
        '6' {
            $curFlts = Get-PropOrKey -Object $rep -Name 'FilterExpressions'
            $curStr = if ($null -ne $curFlts) { (@($curFlts) -join '; ') } else { '' }
            Write-Host "Điều kiện lọc hiện tại: $curStr" -ForegroundColor Cyan
            $newFltsInput = Read-Host "Nhập điều kiện mới (cách nhau dấu ';', gõ 'CLEAR' để xóa)"
            if ($newFltsInput -eq 'CLEAR') {
                if ($rep -is [System.Collections.IDictionary]) { [void]$rep.Remove('FilterExpressions') }
                elseif ($rep.PSObject.Properties['FilterExpressions']) { $rep.PSObject.Properties.Remove('FilterExpressions') }
            } elseif (-not [string]::IsNullOrWhiteSpace($newFltsInput)) {
                $flts = @($newFltsInput.Split(';') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                Set-ObjectProperty -Object $rep -Name 'FilterExpressions' -Value $flts
            }
            Save-MpaConfig -Path $ConfigPath -Config $Config
            Write-Host "[+] Đã cập nhật điều kiện lọc." -ForegroundColor Green
        }
        '7' {
            $confirm = Read-Host "Xác nhận xóa '$($rep.Name)'? [y/N]"
            if ($confirm -eq 'y' -or $confirm -eq 'Y') {
                $list = New-Object System.Collections.Generic.List[object]
                foreach ($r in $reports) {
                    if ($r -ne $rep) { $list.Add($r) }
                }
                $Config.Reports = $list.ToArray()
                Save-MpaConfig -Path $ConfigPath -Config $Config
                Write-Host "[+] Đã xóa báo cáo khỏi cấu hình." -ForegroundColor Green
            }
        }
    }
}

function Action-TestDownloadSingle {
    param($Config)
    Write-Host "`n=== [7] Tải Thử nghiệm 1 Báo cáo ===" -ForegroundColor Cyan
    $reports = Get-PropOrKey -Object $Config -Name 'Reports'
    if ($null -eq $reports -or @($reports).Count -eq 0) {
        Write-Host "Chưa có báo cáo nào." -ForegroundColor Yellow
        return
    }

    Action-ListReports -Config $Config
    $sel = Read-Host "`nChọn số thứ tự báo cáo để tải thử [1-$(@($reports).Count)]"
    if (-not [int]::TryParse($sel, [ref]$null) -or [int]$sel -lt 1 -or [int]$sel -gt @($reports).Count) { return }

    $rep = $reports[[int]$sel - 1]
    $instInfo = Get-MpaReportInstance -Config $Config -Report $rep
    $baseUrl = Get-RequiredProperty -Object $instInfo.Config -Name 'MpaBaseUrl'
    $retryPolicy = Get-MpaRetryPolicy -Config $Config
    $httpSettings = Get-MpaHttpSettings -Config $Config

    $sess = $null
    try {
        $sess = Connect-MpaSessionInteractive -Config $Config -Instance $instInfo

        $resolvedPath = Resolve-DynamicTokens -Text $rep.Path -Report $rep
        $resolvedPath = $resolvedPath.Replace('{Username}', $sess.Username)

        foreach ($fmtConfig in $rep.Formats) {
            $format = $fmtConfig.Format
            $rawOut = $fmtConfig.OutputPath
            $resolvedOut = Resolve-DynamicTokens -Text $rawOut -Report $rep -Format $format
            $resolvedOut = $resolvedOut.Replace('{Username}', $sess.Username)
            if (-not [System.IO.Path]::IsPathRooted($resolvedOut)) {
                $resolvedOut = Join-Path $scriptDir $resolvedOut
            }

            Write-Host "`nĐường dẫn xuất: $resolvedOut" -ForegroundColor Cyan
            $res = Invoke-MpaReportDownloadWithRetry `
                -BaseUrl $baseUrl `
                -Path $resolvedPath `
                -Format $format `
                -OutputPath $resolvedOut `
                -SessionId $sess.SessionId `
                -RetryPolicy $retryPolicy `
                -Parameters (Get-PropOrKey -Object $rep -Name 'Parameters') `
                -FilterExpressions (Get-PropOrKey -Object $rep -Name 'FilterExpressions') `
                -TimeoutSeconds ($httpSettings.TimeoutMinutes * 60)

            if ($res.Success) {
                Write-Host "[+] Tải hoàn tất! Tệp: $resolvedOut ($([math]::Round($res.FileSizeBytes / 1KB, 1)) KB)" -ForegroundColor Green
            } else {
                Write-Host "[-] Tải thất bại: $($res.ErrorMessage)" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host "[-] Lỗi trong quá trình tải: $_" -ForegroundColor Red
    }
    finally {
        if ($null -ne $sess -and -not [string]::IsNullOrWhiteSpace($sess.SessionId)) {
            Disconnect-Obiee -BaseUrl $sess.BaseUrl -SessionId $sess.SessionId | Out-Null
        }
    }
}

function Action-RunBatch {
    Write-Host "`n=== [8] Chạy Tải Báo cáo Hàng loạt (Batch Runner) ===" -ForegroundColor Cyan
    $downloaderScript = Join-Path $scriptDir 'MpaReportDownloader.ps1'
    if (Test-Path -LiteralPath $downloaderScript) {
        & $downloaderScript -ConfigPath $ConfigPath
    } else {
        Write-Host "Không tìm thấy $downloaderScript" -ForegroundColor Red
    }
}

function Action-ManageCredentials {
    param($Config)
    Write-Host "`n=== [9] Cài đặt Tài khoản Xác thực (Windows Credential Manager) ===" -ForegroundColor Cyan
    $credTarget = Get-RequiredProperty -Object $Config -Name 'CredentialTarget'
    $existing = Get-WindowsGenericCredential -Target $credTarget
    if ($null -ne $existing) {
        Write-Host "Tài khoản hiện tại cho '$credTarget': $($existing.Username)" -ForegroundColor Yellow
    }

    $cred = Get-Credential -Message "Nhập thông tin tài khoản MPA ($credTarget)"
    if ($null -ne $cred) {
        Set-WindowsGenericCredential -Target $credTarget -Username $cred.UserName -Password $cred.Password
        Write-Host "[+] Đã lưu thông tin tài khoản thành công cho target '$credTarget'." -ForegroundColor Green
    }
}

# -----------------------------------------------------------------------------
# Vòng lặp Menu Chính
# -----------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    Write-Host "Không tìm thấy $ConfigPath. Đang tạo cấu hình mặc định..." -ForegroundColor Yellow
    # Create minimal config
}

$config = Load-MpaConfig -Path $ConfigPath

while ($true) {
    Write-Host "`n========================================================" -ForegroundColor Cyan
    Write-Host "          TRÌNH QUẢN LÝ CẤU HÌNH BÁO CÁO MPA / OBIEE    " -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "  [1] Kiểm tra Kết nối & Đăng nhập OBIEE SOAP"
    Write-Host "  [2] Duyệt Cây Thư mục Web Catalog (Browse Catalog)"
    Write-Host "  [3] Tra cứu Thông tin Chi tiết Đối tượng Catalog"
    Write-Host "  [4] Xem Danh sách Báo cáo Đã Cấu hình"
    Write-Host "  [5] Thêm Báo cáo Mới vào Cấu hình"
    Write-Host "  [6] Bật / Tắt hoặc Chỉnh sửa Báo cáo"
    Write-Host "  [7] Tải Thử nghiệm 1 Báo cáo (Test Download)"
    Write-Host "  [8] Chạy Tải Báo cáo Hàng loạt (Batch Downloader)"
    Write-Host "  [9] Cài đặt / Đổi Mật khẩu trong Windows Credential Manager"
    Write-Host "  [0] Thoát"
    Write-Host "--------------------------------------------------------" -ForegroundColor Gray

    $mainChoice = Read-Host "Nhập lựa chọn của bạn [0-9]"
    switch ($mainChoice) {
        '1' { Action-TestConnection -Config $config }
        '2' { Action-BrowseCatalog -Config $config }
        '3' { Action-InspectItem -Config $config }
        '4' { Action-ListReports -Config $config }
        '5' { Action-AddReport -Config $config; $config = Load-MpaConfig -Path $ConfigPath }
        '6' { Action-ToggleOrEditReport -Config $config; $config = Load-MpaConfig -Path $ConfigPath }
        '7' { Action-TestDownloadSingle -Config $config }
        '8' { Action-RunBatch }
        '9' { Action-ManageCredentials -Config $config }
        '0' { Write-Host "`nTạm biệt!`n" -ForegroundColor Cyan; exit 0 }
        default { Write-Host "Lựa chọn không hợp lệ." -ForegroundColor Yellow }
    }
}
