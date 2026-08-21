<#
.SYNOPSIS
    Giao diện đồ họa (Windows Forms) Quản lý Cấu hình CognosDownloader.

.DESCRIPTION
    Cung cấp giao diện quản lý trực quan:
    - Quản lý danh sách báo cáo Cognos, tham số prompt và định dạng đầu ra.
    - Tự động dò tìm tham số prompt trực tiếp từ máy chủ Cognos qua REST API.
    - Xem trước (Live Preview) giá trị dynamic token và đường dẫn lưu tệp.
    - Cấu hình nhiều máy chủ Cognos dùng chung tài khoản xác thực.
    - Cài đặt hệ thống ghi log và audit CSV theo dõi hiệu năng.
    - Quản lý tài khoản đăng nhập bảo mật qua Windows Credential Manager.
    - Kiểm tra kết nối và chạy thử nghiệm tải báo cáo ngay trên giao diện.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = '.\cognos-reports.json'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($ConfigPath) -or $ConfigPath -eq '.\cognos-reports.json') {
    $ConfigPath = Join-Path $scriptDir 'cognos-reports.json'
}

# Nạp thư viện đồ họa Windows Forms và Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Nạp module dùng chung
. (Join-Path $scriptDir 'CognosCommon.ps1')

# -----------------------------------------------------------------------------
# Trạng thái Toàn cục & Đọc Cấu hình
# -----------------------------------------------------------------------------

$script:CurrentConfig = $null

function Load-AppConfig {
    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        throw "Không tìm thấy tệp cấu hình: $ConfigPath. Vui lòng kiểm tra lại đường dẫn."
    }
    $script:CurrentConfig = Load-CognosConfig -Path $ConfigPath
    if ($null -eq $script:CurrentConfig) {
        throw "Không thể phân tích định dạng JSON của tệp: $ConfigPath"
    }
}

Load-AppConfig

# -----------------------------------------------------------------------------
# Khởi tạo Cửa sổ Chính
# -----------------------------------------------------------------------------

$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Cognos Downloader - Trình Quản lý Cấu hình"
$mainForm.Size = New-Object System.Drawing.Size(1100, 760)
$mainForm.MinimumSize = New-Object System.Drawing.Size(940, 640)
$mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
$mainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$mainForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font

# -----------------------------------------------------------------------------
# Thanh Công cụ Phía trên (Header Toolbar)
# -----------------------------------------------------------------------------

$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$topPanel.Height = 56
$topPanel.BackColor = [System.Drawing.Color]::FromArgb(243, 245, 249)
$topPanel.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)
$mainForm.Controls.Add($topPanel)

$pnlTopButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlTopButtons.Dock = [System.Windows.Forms.DockStyle]::Right
$pnlTopButtons.AutoSize = $true
$pnlTopButtons.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
$pnlTopButtons.WrapContents = $false
$topPanel.Controls.Add($pnlTopButtons)

$btnSetCred = New-Object System.Windows.Forms.Button
$btnSetCred.Text = "Tài khoản Đăng nhập"
$btnSetCred.Size = New-Object System.Drawing.Size(145, 32)
$btnSetCred.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$btnSetCred.BackColor = [System.Drawing.Color]::FromArgb(240, 242, 245)
$btnSetCred.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSetCred.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(215, 218, 224)
$pnlTopButtons.Controls.Add($btnSetCred)

$btnTestAll = New-Object System.Windows.Forms.Button
$btnTestAll.Text = "Kiểm tra Kết nối"
$btnTestAll.Size = New-Object System.Drawing.Size(130, 32)
$btnTestAll.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$btnTestAll.BackColor = [System.Drawing.Color]::FromArgb(232, 240, 254)
$btnTestAll.ForeColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
$btnTestAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnTestAll.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(190, 215, 250)
$pnlTopButtons.Controls.Add($btnTestAll)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Lưu Cấu hình"
$btnSave.Size = New-Object System.Drawing.Size(110, 32)
$btnSave.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$btnSave.BackColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
$btnSave.ForeColor = [System.Drawing.Color]::White
$btnSave.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSave.FlatAppearance.BorderSize = 0
$pnlTopButtons.Controls.Add($btnSave)

$lblConfigPath = New-Object System.Windows.Forms.Label
$lblConfigPath.Text = "Tệp cấu hình: $ConfigPath"
$lblConfigPath.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblConfigPath.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblConfigPath.ForeColor = [System.Drawing.Color]::FromArgb(60, 64, 67)
$lblConfigPath.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$topPanel.Controls.Add($lblConfigPath)

# -----------------------------------------------------------------------------
# Khung Hiển thị Nhật ký Phía dưới (Bottom Log Panel)
# -----------------------------------------------------------------------------

$bottomPanel = New-Object System.Windows.Forms.Panel
$bottomPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$bottomPanel.Height = 140
$bottomPanel.Padding = New-Object System.Windows.Forms.Padding(16, 8, 16, 12)

$lblLogHeader = New-Object System.Windows.Forms.Label
$lblLogHeader.Text = "Nhật ký Hoạt động & Thực thi:"
$lblLogHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$lblLogHeader.Height = 22
$lblLogHeader.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblLogHeader.ForeColor = [System.Drawing.Color]::FromArgb(90, 95, 100)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.Multiline = $true
$txtLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(32, 33, 36)
$txtLog.ForeColor = [System.Drawing.Color]::FromArgb(232, 234, 237)
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)

$bottomPanel.Controls.Add($lblLogHeader)
$bottomPanel.Controls.Add($txtLog)
$txtLog.BringToFront()
$mainForm.Controls.Add($bottomPanel)

function Gui-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $time = [DateTime]::Now.ToString("HH:mm:ss")
    $line = "[$time] [$Level] $Message" + [Environment]::NewLine
    $txtLog.AppendText($line)
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
}

# -----------------------------------------------------------------------------
# Hệ thống Tab Quản lý (TabControl)
# -----------------------------------------------------------------------------

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabControl.Padding = New-Object System.Drawing.Point(14, 8)

$tabReports   = New-Object System.Windows.Forms.TabPage -Property @{ Text = "Quản lý Báo cáo"; Padding = New-Object System.Windows.Forms.Padding(12) }
$tabInstances = New-Object System.Windows.Forms.TabPage -Property @{ Text = "Máy chủ Cognos"; Padding = New-Object System.Windows.Forms.Padding(12) }
$tabSettings  = New-Object System.Windows.Forms.TabPage -Property @{ Text = "Nhật ký & Cài đặt"; Padding = New-Object System.Windows.Forms.Padding(16) }

$tabControl.Controls.Add($tabReports)
$tabControl.Controls.Add($tabInstances)
$tabControl.Controls.Add($tabSettings)

$mainForm.Controls.Add($tabControl)
$tabControl.BringToFront()

# =============================================================================
# TAB 1: QUẢN LÝ BÁO CÁO
# =============================================================================

$splitReports = New-Object System.Windows.Forms.SplitContainer
$splitReports.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitReports.Orientation = [System.Windows.Forms.Orientation]::Vertical
$splitReports.SplitterWidth = 6
$tabReports.Controls.Add($splitReports)

# Phía trái: Danh sách Báo cáo & Thanh nút chức năng
$panelReportsLeft = New-Object System.Windows.Forms.Panel
$panelReportsLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitReports.Panel1.Controls.Add($panelReportsLeft)

$pnlReportButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlReportButtons.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlReportButtons.Height = 42
$pnlReportButtons.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$pnlReportButtons.WrapContents = $false
$panelReportsLeft.Controls.Add($pnlReportButtons)

$btnAddReport = New-Object System.Windows.Forms.Button -Property @{ Text = "+ Thêm Báo cáo"; Size = New-Object System.Drawing.Size(105, 32); Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0) }
$btnEditReport = New-Object System.Windows.Forms.Button -Property @{ Text = "Sửa Báo cáo"; Size = New-Object System.Drawing.Size(95, 32); Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0) }
$btnDeleteReport = New-Object System.Windows.Forms.Button -Property @{ Text = "Xóa"; Size = New-Object System.Drawing.Size(60, 32); Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0) }
$btnRunReport = New-Object System.Windows.Forms.Button -Property @{ Text = "Tải Mục Chọn"; Size = New-Object System.Drawing.Size(100, 32); Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0) }
$btnRunAllReports = New-Object System.Windows.Forms.Button -Property @{ Text = "Tải Tất Cả Báo Cáo"; Size = New-Object System.Drawing.Size(140, 32); Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0); BackColor = [System.Drawing.Color]::FromArgb(232, 240, 254); ForeColor = [System.Drawing.Color]::FromArgb(26, 115, 232); FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
$btnRunAllReports.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(190, 215, 250)

$pnlReportButtons.Controls.AddRange(@($btnAddReport, $btnEditReport, $btnDeleteReport, $btnRunReport, $btnRunAllReports))

$gridReports = New-Object System.Windows.Forms.DataGridView
$gridReports.Dock = [System.Windows.Forms.DockStyle]::Fill
$gridReports.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$gridReports.MultiSelect = $false
$gridReports.ReadOnly = $true
$gridReports.AllowUserToAddRows = $false
$gridReports.AllowUserToDeleteRows = $false
$gridReports.RowHeadersVisible = $false
$gridReports.BackgroundColor = [System.Drawing.Color]::White
$gridReports.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$gridReports.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D

$colName   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Name"; HeaderText = "Tên Báo cáo"; FillWeight = 160 }
$colInst   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Instance"; HeaderText = "Máy chủ"; FillWeight = 70 }
$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Status"; HeaderText = "Trạng thái"; FillWeight = 55 }
$colFmt    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Formats"; HeaderText = "Định dạng"; FillWeight = 65 }

[void]$gridReports.Columns.Add($colName)
[void]$gridReports.Columns.Add($colInst)
[void]$gridReports.Columns.Add($colStatus)
[void]$gridReports.Columns.Add($colFmt)
$panelReportsLeft.Controls.Add($gridReports)
$gridReports.BringToFront()

# Phía phải: Xem trước Chi tiết & Token Động
$panelReportsRight = New-Object System.Windows.Forms.Panel
$panelReportsRight.Dock = [System.Windows.Forms.DockStyle]::Fill
$panelReportsRight.Padding = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$splitReports.Panel2.Controls.Add($panelReportsRight)

$lblDetailsHeader = New-Object System.Windows.Forms.Label
$lblDetailsHeader.Text = "Chi tiết Báo cáo & Xem trước Dynamic Token"
$lblDetailsHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$lblDetailsHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$lblDetailsHeader.Height = 28

$txtDetails = New-Object System.Windows.Forms.TextBox
$txtDetails.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtDetails.Multiline = $true
$txtDetails.ReadOnly = $true
$txtDetails.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtDetails.BackColor = [System.Drawing.Color]::FromArgb(250, 252, 255)
$txtDetails.Font = New-Object System.Drawing.Font("Consolas", 9.5)

$panelReportsRight.Controls.Add($lblDetailsHeader)
$panelReportsRight.Controls.Add($txtDetails)
$txtDetails.BringToFront()

function Refresh-ReportsGrid {
    $gridReports.Rows.Clear()
    $rawReports = Get-PropOrKey -Object $script:CurrentConfig -Name 'Reports'
    $reports = if ($null -ne $rawReports) { @($rawReports) } else { @() }
    
    foreach ($rep in $reports) {
        $name = Get-PropOrKey -Object $rep -Name 'Name'
        $source = Get-PropOrKey -Object $rep -Name 'Source'
        if ([string]::IsNullOrWhiteSpace($name)) { $name = $source }
        $instObj = try { Get-CognosReportInstance -Config $script:CurrentConfig -Report $rep } catch { $null }
        $inst = if ($instObj) { $instObj.Name } else { "Chưa xác định" }
        $enabled = Get-PropOrKey -Object $rep -Name 'Enabled'
        $status = if ($null -ne $enabled -and $enabled -eq $false) { "Đã tắt" } else { "Đang bật" }
        $rawFmts = Get-PropOrKey -Object $rep -Name 'Formats'
        $fmts = if ($null -ne $rawFmts -and @($rawFmts).Count -gt 0) { (@($rawFmts) | ForEach-Object { Get-PropOrKey -Object $_ -Name 'Format' }) -join ', ' } else { "Không có" }

        $rowIdx = $gridReports.Rows.Add($name, $inst, $status, $fmts)
        $gridReports.Rows[$rowIdx].Tag = $rep
    }

    if ($gridReports.Rows.Count -gt 0) {
        $gridReports.Rows[0].Selected = $true
        Update-ReportDetailsPreview
    } else {
        $txtDetails.Text = "Chưa có báo cáo nào được cấu hình hoặc lựa chọn."
    }
}

function Update-ReportDetailsPreview {
    if ($gridReports.SelectedRows.Count -eq 0) {
        $txtDetails.Text = "Chưa chọn báo cáo."
        return
    }

    $rep = $gridReports.SelectedRows[0].Tag
    if ($null -eq $rep) { return }

    $instObj = try { Get-CognosReportInstance -Config $script:CurrentConfig -Report $rep } catch { $null }
    $instName = if ($instObj) { "$($instObj.Name) ($($instObj.CognosBaseUrl))" } else { "<Lỗi: Máy chủ không xác định>" }

    $repName = Get-PropOrKey -Object $rep -Name 'Name'
    $repSource = Get-PropOrKey -Object $rep -Name 'Source'
    $repSourceType = Get-PropOrKey -Object $rep -Name 'SourceType'
    if ([string]::IsNullOrWhiteSpace($repSourceType)) { $repSourceType = 'report' }
    $repEnabled = Get-PropOrKey -Object $rep -Name 'Enabled'

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("TÊN BÁO CÁO:       $repName")
    [void]$sb.AppendLine("MÃ BÁO CÁO (ID):   $repSource (Loại: $repSourceType)")
    [void]$sb.AppendLine("MÁY CHỦ COGNOS:    $instName")
    [void]$sb.AppendLine("TRẠNG THÁI:        $(if ($null -ne $repEnabled -and $repEnabled -eq $false) { 'ĐÃ TẮT' } else { 'ĐANG BẬT' })")
    [void]$sb.AppendLine("`n-----------------------------------------------------------")
    [void]$sb.AppendLine("THAM SỐ BÁO CÁO (PROMPTS) & GIÁ TRỊ TÍNH TOÁN:")
    [void]$sb.AppendLine("-----------------------------------------------------------")
    
    $paramsObj = Get-PropOrKey -Object $rep -Name 'Parameters'
    if ($null -ne $paramsObj) {
        $paramPairs = @(Get-ObjectKeyValuePairs -Object $paramsObj)
        if (@($paramPairs).Count -gt 0) {
            foreach ($prop in @($paramPairs)) {
                if ($prop.Value -is [System.Collections.IEnumerable] -and -not ($prop.Value -is [string])) {
                    $arrStr = ($prop.Value | ForEach-Object { [string]$_ }) -join ', '
                    [void]$sb.AppendLine("  * $($prop.Name) = [$arrStr] (Mảng $(@($prop.Value).Count) phần tử)")
                } else {
                    $raw = if ($null -ne $prop.Value) { [string]$prop.Value } else { '' }
                    $eval = if ($raw) { Resolve-DynamicTokens -Text $raw -Report $rep } else { '<trống>' }
                    if ($raw -and $raw -ne $eval) {
                        [void]$sb.AppendLine("  * $($prop.Name) = `"$raw`" -> [Giá trị thực tế: $eval]")
                    } else {
                        [void]$sb.AppendLine("  * $($prop.Name) = `"$raw`"")
                    }
                }
            }
        } else {
            [void]$sb.AppendLine("  (Không có tham số)")
        }
    } else {
        [void]$sb.AppendLine("  (Không có tham số)")
    }

    [void]$sb.AppendLine("`n-----------------------------------------------------------")
    [void]$sb.AppendLine("ĐỊNH DẠNG ĐẦU RA & ĐƯỜNG DẪN TỆP LƯU:")
    [void]$sb.AppendLine("-----------------------------------------------------------")
    $fmtsObj = Get-PropOrKey -Object $rep -Name 'Formats'
    if ($null -ne $fmtsObj -and @($fmtsObj).Count -gt 0) {
        foreach ($f in @($fmtsObj)) {
            $fFormat = Get-PropOrKey -Object $f -Name 'Format'
            $rawPath = [string](Get-PropOrKey -Object $f -Name 'OutputPath')
            $evalPath = Resolve-DynamicTokens -Text $rawPath -Report $rep -Format ([string]$fFormat)
            [void]$sb.AppendLine("  * Định dạng:  $fFormat")
            [void]$sb.AppendLine("    Mẫu tệp:    $rawPath")
            [void]$sb.AppendLine("    Đường dẫn:  $evalPath")
        }
    } else {
        [void]$sb.AppendLine("  (Không có định dạng)")
    }

    $txtDetails.Text = $sb.ToString()
}

$gridReports.add_SelectionChanged({ Update-ReportDetailsPreview })

# =============================================================================
# TAB 2: CẤU HÌNH MÁY CHỦ COGNOS
# =============================================================================

$panelInst = New-Object System.Windows.Forms.Panel
$panelInst.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabInstances.Controls.Add($panelInst)

$pnlInstButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlInstButtons.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlInstButtons.Height = 42
$pnlInstButtons.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$pnlInstButtons.WrapContents = $false
$panelInst.Controls.Add($pnlInstButtons)

$btnAddInst = New-Object System.Windows.Forms.Button -Property @{ Text = "+ Thêm Máy chủ"; Size = New-Object System.Drawing.Size(115, 32); Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0) }
$btnEditInst = New-Object System.Windows.Forms.Button -Property @{ Text = "Sửa Máy chủ"; Size = New-Object System.Drawing.Size(105, 32); Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0) }
$btnDeleteInst = New-Object System.Windows.Forms.Button -Property @{ Text = "Xóa"; Size = New-Object System.Drawing.Size(70, 32); Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0) }
$btnSetDefaultInst = New-Object System.Windows.Forms.Button -Property @{ Text = "Đặt làm Mặc định"; Size = New-Object System.Drawing.Size(130, 32); Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0) }

$pnlInstButtons.Controls.AddRange(@($btnAddInst, $btnEditInst, $btnDeleteInst, $btnSetDefaultInst))

$gridInstances = New-Object System.Windows.Forms.DataGridView
$gridInstances.Dock = [System.Windows.Forms.DockStyle]::Fill
$gridInstances.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$gridInstances.MultiSelect = $false
$gridInstances.ReadOnly = $true
$gridInstances.AllowUserToAddRows = $false
$gridInstances.AllowUserToDeleteRows = $false
$gridInstances.RowHeadersVisible = $false
$gridInstances.BackgroundColor = [System.Drawing.Color]::White
$gridInstances.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$gridInstances.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D

$colInstKey = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Key"; HeaderText = "Tên Máy chủ"; FillWeight = 70 }
$colInstUrl = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "BaseUrl"; HeaderText = "Đường dẫn Base URL"; FillWeight = 160 }
$colInstNs  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Namespace"; HeaderText = "Namespace (Không gian tên)"; FillWeight = 70 }
$colInstDef = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "IsDefault"; HeaderText = "Mặc định"; FillWeight = 50 }

[void]$gridInstances.Columns.Add($colInstKey)
[void]$gridInstances.Columns.Add($colInstUrl)
[void]$gridInstances.Columns.Add($colInstNs)
[void]$gridInstances.Columns.Add($colInstDef)
$panelInst.Controls.Add($gridInstances)
$gridInstances.BringToFront()

function Refresh-InstancesGrid {
    $gridInstances.Rows.Clear()
    $instMap = Get-CognosInstances -Config $script:CurrentConfig
    $defaultName = if ($script:CurrentConfig.PSObject.Properties['DefaultInstance'] -and $script:CurrentConfig.DefaultInstance) { [string]$script:CurrentConfig.DefaultInstance } else { '' }

    foreach ($k in $instMap.Keys) {
        $inst = $instMap[$k]
        $isDef = if ($defaultName -and $k -ieq $defaultName) { "MẶC ĐỊNH" } else { "" }
        [void]$gridInstances.Rows.Add($k, $inst.CognosBaseUrl, $inst.Namespace, $isDef)
    }
}

# =============================================================================
# TAB 3: NHẬT KÝ & CÀI ĐẶT
# =============================================================================

$panelSettings = New-Object System.Windows.Forms.Panel
$panelSettings.Dock = [System.Windows.Forms.DockStyle]::Fill
$panelSettings.AutoScroll = $true
$tabSettings.Controls.Add($panelSettings)

# Hàm trợ giúp tạo dòng cấu hình chuẩn
function New-SettingRowPanel {
    param([string]$LabelText, [System.Windows.Forms.Control]$InputControl, [System.Windows.Forms.Control]$ExtraControl = $null)
    $row = New-Object System.Windows.Forms.Panel
    $row.Dock = [System.Windows.Forms.DockStyle]::Top
    $row.Height = 38
    $row.Padding = New-Object System.Windows.Forms.Padding(0, 4, 0, 4)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $LabelText
    $lbl.Dock = [System.Windows.Forms.DockStyle]::Left
    $lbl.Width = 200
    $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $row.Controls.Add($lbl)

    if ($null -ne $ExtraControl) {
        $ExtraControl.Dock = [System.Windows.Forms.DockStyle]::Right
        $row.Controls.Add($ExtraControl)
    }

    $InputControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $row.Controls.Add($InputControl)
    $InputControl.BringToFront()

    return $row
}

# Thẻ 1: Cấu hình Tài khoản Xác thực
$pnlCredCard = New-Object System.Windows.Forms.Panel
$pnlCredCard.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlCredCard.Height = 110
$pnlCredCard.BackColor = [System.Drawing.Color]::FromArgb(250, 252, 255)
$pnlCredCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlCredCard.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)
$panelSettings.Controls.Add($pnlCredCard)

$lblCredCardHeader = New-Object System.Windows.Forms.Label
$lblCredCardHeader.Text = "Cấu hình Tài khoản Xác thực (Windows Credential)"
$lblCredCardHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$lblCredCardHeader.Height = 24
$lblCredCardHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$lblCredCardHeader.ForeColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
$pnlCredCard.Controls.Add($lblCredCardHeader)

$txtCredTarget = New-Object System.Windows.Forms.TextBox
$btnSaveSettingsOnly = New-Object System.Windows.Forms.Button -Property @{ Text = "Lưu Cài đặt"; Width = 110; Height = 28; BackColor = [System.Drawing.Color]::FromArgb(26, 115, 232); ForeColor = [System.Drawing.Color]::White; FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
$btnSaveSettingsOnly.FlatAppearance.BorderSize = 0
$rowCred = New-SettingRowPanel -LabelText "Tên Target xác thực:" -InputControl $txtCredTarget -ExtraControl $btnSaveSettingsOnly
$pnlCredCard.Controls.Add($rowCred)
$rowCred.BringToFront()

# Khoảng đệm giữa các thẻ
$pnlSpacer = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 16 }
$panelSettings.Controls.Add($pnlSpacer)
$pnlSpacer.BringToFront()

# Thẻ 2: Cấu hình Nhật ký & Audit
$pnlLogCard = New-Object System.Windows.Forms.Panel
$pnlLogCard.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlLogCard.Height = 250
$pnlLogCard.BackColor = [System.Drawing.Color]::FromArgb(250, 252, 255)
$pnlLogCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlLogCard.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)
$panelSettings.Controls.Add($pnlLogCard)
$pnlLogCard.BringToFront()

$lblLogCardHeader = New-Object System.Windows.Forms.Label
$lblLogCardHeader.Text = "Cấu hình Nhật ký & Giám sát Thực thi"
$lblLogCardHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$lblLogCardHeader.Height = 24
$lblLogCardHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$lblLogCardHeader.ForeColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
$pnlLogCard.Controls.Add($lblLogCardHeader)

# Dòng: Đường dẫn Audit CSV
$txtAuditPath = New-Object System.Windows.Forms.TextBox
$rowAuditPath = New-SettingRowPanel -LabelText "Đường dẫn Audit CSV:" -InputControl $txtAuditPath
$pnlLogCard.Controls.Add($rowAuditPath)
$rowAuditPath.BringToFront()

# Dòng: Bật Audit CSV
$pnlAuditChk = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 32 }
$chkAuditCsv = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Kích hoạt ghi nhận chỉ số thực thi hàng tháng (Audit CSV)"; Dock = [System.Windows.Forms.DockStyle]::Fill }
$pnlAuditChk.Controls.Add($chkAuditCsv)
$pnlLogCard.Controls.Add($pnlAuditChk)
$pnlAuditChk.BringToFront()

# Dòng: Mức Log & Thời gian lưu trữ
$pnlLevelRet = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Fill; Margin = New-Object System.Windows.Forms.Padding(0) }
$cmbLogLevel = New-Object System.Windows.Forms.ComboBox -Property @{ Width = 120; DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList; Margin = New-Object System.Windows.Forms.Padding(0, 3, 20, 3) }
[void]$cmbLogLevel.Items.AddRange(@('DEBUG', 'INFO', 'WARN', 'ERROR'))
$lblRetention = New-Object System.Windows.Forms.Label -Property @{ Text = "Lưu trữ (Ngày):"; AutoSize = $true; Margin = New-Object System.Windows.Forms.Padding(0, 6, 8, 0) }
$numRetention = New-Object System.Windows.Forms.NumericUpDown -Property @{ Width = 70; Minimum = 1; Maximum = 365; Value = 30; Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 3) }
$pnlLevelRet.Controls.AddRange(@($cmbLogLevel, $lblRetention, $numRetention))
$rowLogLevel = New-SettingRowPanel -LabelText "Mức Log & Lưu trữ:" -InputControl $pnlLevelRet
$pnlLogCard.Controls.Add($rowLogLevel)
$rowLogLevel.BringToFront()

# Dòng: Thư mục chứa Log
$txtLogDir = New-Object System.Windows.Forms.TextBox
$btnBrowseLogDir = New-Object System.Windows.Forms.Button -Property @{ Text = "Duyệt..."; Width = 80; Height = 26; Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0) }
$rowLogDir = New-SettingRowPanel -LabelText "Thư mục Nhật ký:" -InputControl $txtLogDir -ExtraControl $btnBrowseLogDir
$pnlLogCard.Controls.Add($rowLogDir)
$rowLogDir.BringToFront()

# Dòng: Bật File Log
$pnlLogChk = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 32 }
$chkLogEnabled = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Kích hoạt ghi nhật ký ra tệp văn bản (.log)"; Dock = [System.Windows.Forms.DockStyle]::Fill }
$pnlLogChk.Controls.Add($chkLogEnabled)
$pnlLogCard.Controls.Add($pnlLogChk)
$pnlLogChk.BringToFront()

function Load-SettingsTab {
    $logCfg = if ($script:CurrentConfig.PSObject.Properties['Logging']) { $script:CurrentConfig.Logging } else { $null }
    if ($logCfg) {
        $chkLogEnabled.Checked = if ($logCfg.PSObject.Properties['Enabled']) { [bool]$logCfg.Enabled } else { $true }
        $txtLogDir.Text = if ($logCfg.PSObject.Properties['LogDirectory']) { [string]$logCfg.LogDirectory } else { ".\Logs" }
        $cmbLogLevel.SelectedItem = if ($logCfg.PSObject.Properties['LogLevel']) { [string]$logCfg.LogLevel } else { "INFO" }
        $numRetention.Value = if ($logCfg.PSObject.Properties['RetentionDays']) { [int]$logCfg.RetentionDays } else { 30 }
        $chkAuditCsv.Checked = if ($logCfg.PSObject.Properties['AuditCsvEnabled']) { [bool]$logCfg.AuditCsvEnabled } else { $true }
        $txtAuditPath.Text = if ($logCfg.PSObject.Properties['AuditCsvPath']) { [string]$logCfg.AuditCsvPath } else { ".\Logs\Audit_{yyyyMM}.csv" }
    }
    $txtCredTarget.Text = if ($script:CurrentConfig.PSObject.Properties['CredentialTarget'] -and $script:CurrentConfig.CredentialTarget) { [string]$script:CurrentConfig.CredentialTarget } else { '' }
}

$btnBrowseLogDir.add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Chọn Thư mục Lưu trữ Nhật ký"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtLogDir.Text = $dlg.SelectedPath
    }
})

function Sync-SettingsFromUI {
    $targetName = $txtCredTarget.Text.Trim()
    if ($targetName) {
        Set-ObjectProperty -Object $script:CurrentConfig -Name 'CredentialTarget' -Value $targetName
    }
    Set-ObjectProperty -Object $script:CurrentConfig -Name 'Logging' -Value ([ordered]@{
        Enabled         = $chkLogEnabled.Checked
        LogDirectory    = $txtLogDir.Text.Trim()
        LogFileName     = "CognosDownloader_{yyyyMMdd}.log"
        LogLevel        = [string]$cmbLogLevel.SelectedItem
        RetentionDays   = [int]$numRetention.Value
        AuditCsvEnabled = $chkAuditCsv.Checked
        AuditCsvPath    = $txtAuditPath.Text.Trim()
    })
}

$btnSaveSettingsOnly.add_Click({
    $targetName = $txtCredTarget.Text.Trim()
    if (-not $targetName) {
        [System.Windows.Forms.MessageBox]::Show("Tên Target xác thực không được để trống.", "Lỗi Xác thực", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    Sync-SettingsFromUI
    Save-CognosConfig -Path $ConfigPath -Config $script:CurrentConfig
    Gui-Log "Cài đặt đã được cập nhật và lưu vào: $ConfigPath" 'OK'
    [System.Windows.Forms.MessageBox]::Show("Cài đặt đã được lưu thành công.", "Đã lưu Cài đặt", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# =============================================================================
# CÁC HỘP THOẠI & XỬ LÝ SỰ KIỆN
# =============================================================================

# Xử lý: Lưu toàn bộ Cấu hình
$btnSave.add_Click({
    Sync-SettingsFromUI
    Save-CognosConfig -Path $ConfigPath -Config $script:CurrentConfig
    Gui-Log "Cấu hình đã được lưu thành công vào: $ConfigPath" 'OK'
    [System.Windows.Forms.MessageBox]::Show("Cấu hình đã được lưu thành công vào:`n$ConfigPath", "Đã lưu Cấu hình", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Xử lý: Kiểm tra Kết nối Tất cả Máy chủ
$btnTestAll.add_Click({
    $btnTestAll.Enabled = $false
    try {
        $instMap = Get-CognosInstances -Config $script:CurrentConfig
        $credTarget = Get-RequiredProperty -Object $script:CurrentConfig -Name 'CredentialTarget'
        $stored = Get-WindowsGenericCredential -Target $credTarget
        
        if ($null -eq $stored) {
            Gui-Log "Không tìm thấy thông tin tài khoản cho '$credTarget' trong Windows Credential Manager." 'WARN'
            [System.Windows.Forms.MessageBox]::Show("Vui lòng bấm 'Tài khoản Đăng nhập' để lưu tài khoản Cognos trước.", "Thiếu Thông tin Đăng nhập", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        Gui-Log "Bắt đầu kiểm tra kết nối tới $($instMap.Count) máy chủ với target '$credTarget'..."
        foreach ($k in $instMap.Keys) {
            $inst = $instMap[$k]
            try {
                Gui-Log "Đang kết nối tới [$k] tại $($inst.CognosBaseUrl)..."
                $http = New-CognosHttpClient -TimeoutMinutes 1
                $xsrf = Invoke-CognosLogin -Context $http -CognosBaseUrl $inst.CognosBaseUrl -Namespace $inst.Namespace -Username $stored.Username -Password $stored.Password
                $http.Client.Dispose()
                $http.Handler.Dispose()
                Gui-Log "Xác thực THÀNH CÔNG trên máy chủ [$k] (Namespace: $($inst.Namespace))." 'OK'
            }
            catch {
                Gui-Log "Xác thực THẤT BẠI trên máy chủ [$k]: $($_.Exception.Message)" 'ERROR'
            }
        }
    }
    finally {
        $btnTestAll.Enabled = $true
    }
})

# Xử lý: Hộp thoại Nhập Tài khoản Xác thực
$btnSetCred.add_Click({
    $credTarget = Get-RequiredProperty -Object $script:CurrentConfig -Name 'CredentialTarget'
    $stored = Get-WindowsGenericCredential -Target $credTarget

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Cài đặt Tài khoản Xác thực Windows ($credTarget)"
    $dlg.Size = New-Object System.Drawing.Size(460, 240)
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dlg.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $dlg.Padding = New-Object System.Windows.Forms.Padding(20)

    $pnlDlgContent = New-Object System.Windows.Forms.Panel
    $pnlDlgContent.Dock = [System.Windows.Forms.DockStyle]::Fill
    $dlg.Controls.Add($pnlDlgContent)

    $txtU = New-Object System.Windows.Forms.TextBox -Property @{ Text = if ($stored) { $stored.Username } else { '' } }
    $rowU = New-SettingRowPanel -LabelText "Tên đăng nhập:" -InputControl $txtU
    $pnlDlgContent.Controls.Add($rowU)
    $rowU.SendToBack()

    $txtP = New-Object System.Windows.Forms.TextBox -Property @{ PasswordChar = '*' }
    $rowP = New-SettingRowPanel -LabelText "Mật khẩu:" -InputControl $txtP
    $pnlDlgContent.Controls.Add($rowP)
    $rowP.SendToBack()

    $pnlDlgBtns = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Bottom; Height = 36; FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft }
    $btnCancel = New-Object System.Windows.Forms.Button -Property @{ Text = "Hủy bỏ"; Size = New-Object System.Drawing.Size(85, 30); DialogResult = [System.Windows.Forms.DialogResult]::Cancel }
    $btnOk = New-Object System.Windows.Forms.Button -Property @{ Text = "Lưu Tài khoản"; Size = New-Object System.Drawing.Size(130, 30); DialogResult = [System.Windows.Forms.DialogResult]::OK; BackColor = [System.Drawing.Color]::FromArgb(26, 115, 232); ForeColor = [System.Drawing.Color]::White; FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
    $btnOk.FlatAppearance.BorderSize = 0
    $pnlDlgBtns.Controls.AddRange(@($btnCancel, $btnOk))
    $dlg.Controls.Add($pnlDlgBtns)
    $pnlDlgBtns.BringToFront()
    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel

    if ($dlg.ShowDialog($mainForm) -eq [System.Windows.Forms.DialogResult]::OK) {
        if (-not [string]::IsNullOrWhiteSpace($txtU.Text) -and -not [string]::IsNullOrWhiteSpace($txtP.Text)) {
            $secPass = ConvertTo-SecureString $txtP.Text -AsPlainText -Force
            Set-WindowsGenericCredential -Target $credTarget -Username $txtU.Text.Trim() -Password $secPass
            Gui-Log "Đã lưu tài khoản xác thực cho '$credTarget' vào Windows Credential Manager." 'OK'
            [System.Windows.Forms.MessageBox]::Show("Tài khoản đã được lưu bảo mật trong Windows Credential Manager.", "Đã lưu Tài khoản", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
})

# -----------------------------------------------------------------------------
# Hộp thoại Lựa chọn Tham số từ Danh mục Server
# -----------------------------------------------------------------------------

function Show-ChoiceSelectionDialog {
    param(
        [Parameter(Mandatory)] [string]$ParamName,
        [Parameter(Mandatory)] $Choices,
        [bool]$IsMultiSelect = $false,
        [string]$CurrentValue = ''
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Lựa chọn giá trị cho: $ParamName"
    $dlg.Size = New-Object System.Drawing.Size(560, 480)
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dlg.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $dlg.Padding = New-Object System.Windows.Forms.Padding(14)

    $pnlTop = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 58 }
    $lblInfo = New-Object System.Windows.Forms.Label -Property @{
        Text = "Tìm kiếm danh mục ($(@($Choices).Count) lựa chọn từ Cognos server):"
        Dock = [System.Windows.Forms.DockStyle]::Top
        Height = 22
        Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    }
    $txtSearch = New-Object System.Windows.Forms.TextBox -Property @{ Dock = [System.Windows.Forms.DockStyle]::Fill }
    $pnlTop.Controls.Add($txtSearch)
    $pnlTop.Controls.Add($lblInfo)
    $lblInfo.BringToFront()

    $pnlBottom = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Bottom; Height = 44 }
    $btnCancel = New-Object System.Windows.Forms.Button -Property @{ Text = "Hủy"; Size = New-Object System.Drawing.Size(80, 30); DialogResult = [System.Windows.Forms.DialogResult]::Cancel }
    $btnOk = New-Object System.Windows.Forms.Button -Property @{ Text = "Chọn"; Size = New-Object System.Drawing.Size(100, 30); DialogResult = [System.Windows.Forms.DialogResult]::OK; BackColor = [System.Drawing.Color]::FromArgb(26, 115, 232); ForeColor = [System.Drawing.Color]::White; FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
    $btnOk.FlatAppearance.BorderSize = 0

    $pnlBtnsRight = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Right; AutoSize = $true; FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft }
    $pnlBtnsRight.Controls.AddRange(@($btnCancel, $btnOk))

    $pnlQuick = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Left; AutoSize = $true }
    $btnSelectAll = New-Object System.Windows.Forms.Button -Property @{ Text = "Chọn tất cả"; Height = 28; Width = 95 }
    $btnDeselectAll = New-Object System.Windows.Forms.Button -Property @{ Text = "Bỏ chọn"; Height = 28; Width = 80 }
    $pnlQuick.Controls.AddRange(@($btnSelectAll, $btnDeselectAll))

    $pnlBottom.Controls.Add($pnlBtnsRight)
    $pnlBottom.Controls.Add($pnlQuick)

    $chkList = New-Object System.Windows.Forms.CheckedListBox
    $chkList.Dock = [System.Windows.Forms.DockStyle]::Fill
    $chkList.CheckOnClick = $true

    $currentSelectedTokens = @($CurrentValue -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

    $script:FilteredChoices = [System.Collections.Generic.List[object]]::new()

    $populateList = {
        param([string]$Filter = '')
        $chkList.Items.Clear()
        $script:FilteredChoices.Clear()
        foreach ($c in @($Choices)) {
            $displayText = if ($c.Display -and $c.Display -ne $c.Use) { "$($c.Use) - $($c.Display)" } else { "$($c.Use)" }
            if ([string]::IsNullOrWhiteSpace($Filter) -or $displayText -like "*$Filter*") {
                $script:FilteredChoices.Add($c)
                $idx = $chkList.Items.Add($displayText)
                if ($currentSelectedTokens -contains $c.Use -or $currentSelectedTokens -contains $displayText) {
                    $chkList.SetItemChecked($idx, $true)
                }
            }
        }
    }

    & $populateList
    $txtSearch.add_TextChanged({ & $populateList -Filter $txtSearch.Text.Trim() })

    $btnSelectAll.add_Click({
        for ($i = 0; $i -lt $chkList.Items.Count; $i++) { $chkList.SetItemChecked($i, $true) }
    })
    $btnDeselectAll.add_Click({
        for ($i = 0; $i -lt $chkList.Items.Count; $i++) { $chkList.SetItemChecked($i, $false) }
    })

    $dlg.Controls.Add($chkList)
    $dlg.Controls.Add($pnlTop)
    $dlg.Controls.Add($pnlBottom)
    $pnlTop.SendToBack()
    $pnlBottom.BringToFront()
    $chkList.BringToFront()
    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedUseValues = [System.Collections.Generic.List[string]]::new()
        for ($i = 0; $i -lt $chkList.CheckedIndices.Count; $i++) {
            $cIdx = $chkList.CheckedIndices[$i]
            if ($cIdx -ge 0 -and $cIdx -lt $script:FilteredChoices.Count) {
                $selectedUseValues.Add($script:FilteredChoices[$cIdx].Use)
            }
        }
        return ($selectedUseValues -join ', ')
    }
    return $null
}

# -----------------------------------------------------------------------------
# Hộp thoại Thêm / Sửa Báo cáo
# -----------------------------------------------------------------------------

function Show-ReportEditDialog {
    param($ExistingReport = $null)

    $isEdit = ($null -ne $ExistingReport)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = if ($isEdit) { "Sửa Báo cáo: $($ExistingReport.Name)" } else { "Thêm Báo cáo Cognos Mới" }
    $dlg.Size = New-Object System.Drawing.Size(780, 720)
    $dlg.MinimumSize = New-Object System.Drawing.Size(700, 620)
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dlg.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $dlg.Padding = New-Object System.Windows.Forms.Padding(16)

    # Thanh nút hành động phía dưới
    $pnlDlgBottom = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Bottom; Height = 50; Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0) }
    $btnCancelReport = New-Object System.Windows.Forms.Button -Property @{ Text = "Hủy bỏ"; Size = New-Object System.Drawing.Size(95, 32); DialogResult = [System.Windows.Forms.DialogResult]::Cancel }
    $btnSaveReport = New-Object System.Windows.Forms.Button -Property @{ Text = "Lưu Báo cáo"; Size = New-Object System.Drawing.Size(120, 32); BackColor = [System.Drawing.Color]::FromArgb(26, 115, 232); ForeColor = [System.Drawing.Color]::White; FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0) }
    $btnSaveReport.FlatAppearance.BorderSize = 0

    $pnlDlgButtonsRight = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Right; AutoSize = $true; FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft; WrapContents = $false }
    $pnlDlgButtonsRight.Controls.AddRange(@($btnCancelReport, $btnSaveReport))
    $pnlDlgBottom.Controls.Add($pnlDlgButtonsRight)

    $btnSaveReport.add_Click({
        [void]$gridParams.EndEdit()
        $src = $txtS.Text.Trim()
        if (-not $src) {
            [System.Windows.Forms.MessageBox]::Show("Vui lòng nhập mã StoreID báo cáo.", "Lỗi Nhập liệu", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $outPath = $txtPath.Text.Trim()
        if (-not $outPath) {
            [System.Windows.Forms.MessageBox]::Show("Vui lòng nhập mẫu đường dẫn tệp xuất.", "Lỗi Nhập liệu", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dlg.Close()
    })

    # Thân Form nhập liệu
    $pnlFormBody = New-Object System.Windows.Forms.Panel
    $pnlFormBody.Dock = [System.Windows.Forms.DockStyle]::Fill
    $pnlFormBody.AutoScroll = $true

    $dlg.Controls.Add($pnlFormBody)
    $dlg.Controls.Add($pnlDlgBottom)
    $pnlDlgBottom.BringToFront()
    $dlg.AcceptButton = $btnSaveReport
    $dlg.CancelButton = $btnCancelReport

    # 1. Tên Báo cáo
    $txtN = New-Object System.Windows.Forms.TextBox -Property @{ Text = if ($isEdit) { [string](Get-PropOrKey -Object $ExistingReport -Name 'Name') } else { '' } }
    $rowName = New-SettingRowPanel -LabelText "Tên Báo cáo:" -InputControl $txtN
    $pnlFormBody.Controls.Add($rowName)
    $rowName.SendToBack()

    # 2. Máy chủ & Trạng thái Bật/Tắt
    $pnlInstRow = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Fill; Margin = New-Object System.Windows.Forms.Padding(0) }
    $cmbI = New-Object System.Windows.Forms.ComboBox -Property @{ Width = 200; DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList; Margin = New-Object System.Windows.Forms.Padding(0, 2, 20, 2) }
    $instMap = Get-CognosInstances -Config $script:CurrentConfig
    foreach ($k in $instMap.Keys) { [void]$cmbI.Items.Add($k) }
    $existingInst = if ($isEdit) { Get-PropOrKey -Object $ExistingReport -Name 'Instance' } else { $null }
    $defaultInst = Get-PropOrKey -Object $script:CurrentConfig -Name 'DefaultInstance'
    $currentInstName = if ($isEdit -and -not [string]::IsNullOrWhiteSpace($existingInst)) {
        [string]$existingInst
    } elseif (-not [string]::IsNullOrWhiteSpace($defaultInst)) {
        [string]$defaultInst
    } else { '' }
    if ($currentInstName -and $cmbI.Items.Contains($currentInstName)) { $cmbI.SelectedItem = $currentInstName } elseif ($cmbI.Items.Count -gt 0) { $cmbI.SelectedIndex = 0 }

    $existingEn = if ($isEdit) { Get-PropOrKey -Object $ExistingReport -Name 'Enabled' } else { $null }
    $chkEn = New-Object System.Windows.Forms.CheckBox -Property @{
        Text = "Kích hoạt tải báo cáo"
        AutoSize = $true
        Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
        Checked = if ($isEdit -and $null -ne $existingEn -and $existingEn -eq $false) { $false } else { $true }
    }
    $pnlInstRow.Controls.AddRange(@($cmbI, $chkEn))
    $rowInst = New-SettingRowPanel -LabelText "Máy chủ Cognos:" -InputControl $pnlInstRow
    $pnlFormBody.Controls.Add($rowInst)
    $rowInst.SendToBack()

    # 3. Mã StoreID & Nút Dò tìm Tham số
    $txtS = New-Object System.Windows.Forms.TextBox -Property @{ Text = if ($isEdit) { [string](Get-PropOrKey -Object $ExistingReport -Name 'Source') } else { '' } }
    $btnInspect = New-Object System.Windows.Forms.Button -Property @{ Text = "Dò tìm Tham số & Danh mục"; Width = 190; Height = 28; Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0) }
    $rowSource = New-SettingRowPanel -LabelText "StoreID / Đường dẫn:" -InputControl $txtS -ExtraControl $btnInspect
    $pnlFormBody.Controls.Add($rowSource)
    $rowSource.SendToBack()

    # 4. Bảng Tham số (Parameters Grid)
    $pnlParamSection = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 205; Padding = New-Object System.Windows.Forms.Padding(0, 6, 0, 6) }
    
    $pnlParamHeaderBar = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 26 }
    $lblParamHeader = New-Object System.Windows.Forms.Label -Property @{ Text = "Danh sách Tham số (Sử dụng token động như {Yesterday}, {Today}):"; Dock = [System.Windows.Forms.DockStyle]::Left; AutoSize = $true; Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold) }
    $btnPickChoice = New-Object System.Windows.Forms.Button -Property @{ Text = "Chọn từ Danh mục Server..."; Dock = [System.Windows.Forms.DockStyle]::Right; Width = 180; Height = 24; Margin = New-Object System.Windows.Forms.Padding(0) }
    $pnlParamHeaderBar.Controls.Add($btnPickChoice)
    $pnlParamHeaderBar.Controls.Add($lblParamHeader)

    $gridParams = New-Object System.Windows.Forms.DataGridView
    $gridParams.Dock = [System.Windows.Forms.DockStyle]::Fill
    $gridParams.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $gridParams.RowHeadersVisible = $false
    $gridParams.BackgroundColor = [System.Drawing.Color]::White

    $colPK = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "ParamName"; HeaderText = "Tên Tham số (ParamName)"; FillWeight = 85 }
    $colPR = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Required"; HeaderText = "Yêu cầu"; FillWeight = 45; ReadOnly = $true }
    $colPV = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "ParamValue"; HeaderText = "Giá trị / Token Động (VD: {Yesterday})"; FillWeight = 130 }
    [void]$gridParams.Columns.Add($colPK)
    [void]$gridParams.Columns.Add($colPR)
    [void]$gridParams.Columns.Add($colPV)

    $existingParams = if ($isEdit) { Get-PropOrKey -Object $ExistingReport -Name 'Parameters' } else { $null }
    if ($null -ne $existingParams) {
        foreach ($prop in @(Get-ObjectKeyValuePairs -Object $existingParams)) {
            $valDisplay = if ($prop.Value -is [System.Collections.IEnumerable] -and -not ($prop.Value -is [string])) {
                ($prop.Value | ForEach-Object { [string]$_ }) -join ', '
            } else {
                [string]$prop.Value
            }
            [void]$gridParams.Rows.Add($prop.Name, "", $valDisplay)
        }
    }
    $pnlParamSection.Controls.Add($gridParams)
    $pnlParamSection.Controls.Add($pnlParamHeaderBar)
    $pnlParamHeaderBar.BringToFront()
    $pnlFormBody.Controls.Add($pnlParamSection)
    $pnlParamSection.SendToBack()

    # Xử lý: Nút Chọn từ Danh mục
    $handlePickChoice = {
        if ($gridParams.SelectedCells.Count -eq 0 -and $gridParams.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Vui lòng chọn dòng tham số cần thiết lập danh mục.", "Chưa Chọn Dòng", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $rowIdx = if ($gridParams.SelectedCells.Count -gt 0) { $gridParams.SelectedCells[0].RowIndex } else { $gridParams.SelectedRows[0].Index }
        $targetRow = $gridParams.Rows[$rowIdx]
        $pName = [string]$targetRow.Cells[0].Value
        $pVal = [string]$targetRow.Cells[2].Value
        $pInfo = $targetRow.Tag

        if ($null -ne $pInfo -and $pInfo.Choices -and @($pInfo.Choices).Count -gt 0) {
            $res = Show-ChoiceSelectionDialog -ParamName $pName -Choices $pInfo.Choices -IsMultiSelect ([bool]$pInfo.IsMultiSelect) -CurrentValue $pVal
            if ($null -ne $res) {
                $targetRow.Cells[2].Value = $res
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Tham số '$pName' không có danh mục lựa chọn cố định từ server hoặc chưa bấm 'Dò tìm Tham số & Danh mục'.", "Không Có Danh Mục", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }

    $btnPickChoice.add_Click($handlePickChoice)
    $gridParams.add_CellDoubleClick({ & $handlePickChoice })

    # 5. Lựa chọn Định dạng Đầu ra
    $cmbFmt = New-Object System.Windows.Forms.ComboBox -Property @{ Width = 200; DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList }
    [void]$cmbFmt.Items.AddRange(@('xlsxData', 'spreadsheetML', 'PDF', 'CSV'))
    $existingFormats = if ($isEdit) { Get-PropOrKey -Object $ExistingReport -Name 'Formats' } else { $null }
    $firstFmt = if ($null -ne $existingFormats -and @($existingFormats).Count -gt 0) { @($existingFormats)[0] } else { $null }
    $existingFmt = if ($firstFmt) { [string](Get-PropOrKey -Object $firstFmt -Name 'Format') } else { 'xlsxData' }
    $cmbFmt.SelectedItem = $existingFmt
    $pnlFmtContainer = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Fill }
    $pnlFmtContainer.Controls.Add($cmbFmt)
    $rowFmt = New-SettingRowPanel -LabelText "Định dạng Xuất:" -InputControl $pnlFmtContainer
    $pnlFormBody.Controls.Add($rowFmt)
    $rowFmt.SendToBack()

    # 6. Mẫu Đường dẫn Tệp Xuất
    $existingPath = if ($firstFmt) { [string](Get-PropOrKey -Object $firstFmt -Name 'OutputPath') } else { '' }
    $txtPath = New-Object System.Windows.Forms.TextBox -Property @{ Text = $existingPath }
    $rowPath = New-SettingRowPanel -LabelText "Mẫu Đường dẫn Tệp:" -InputControl $txtPath
    $pnlFormBody.Controls.Add($rowPath)
    $rowPath.SendToBack()

    # 7. Các nút Chèn Token Nhanh
    $pnlTokens = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Fill; FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight; WrapContents = $false; Margin = New-Object System.Windows.Forms.Padding(0) }
    $tokens = @('{Yesterday:yyyyMMdd}', '{Today:yyyyMMdd}', '{ReportName}', '{Instance}')
    foreach ($t in $tokens) {
        $btnT = New-Object System.Windows.Forms.Button -Property @{ Text = $t; AutoSize = $true; Height = 26; Margin = New-Object System.Windows.Forms.Padding(0, 2, 6, 2) }
        $btnT.add_Click({ $txtPath.Paste($this.Text) })
        $pnlTokens.Controls.Add($btnT)
    }
    $rowTokens = New-SettingRowPanel -LabelText "Chèn Token Động:" -InputControl $pnlTokens
    $pnlFormBody.Controls.Add($rowTokens)
    $rowTokens.SendToBack()

    # Xử lý: Dò tìm tham số từ máy chủ
    $btnInspect.add_Click({
        $src = $txtS.Text.Trim()
        if (-not $src) {
            [System.Windows.Forms.MessageBox]::Show("Vui lòng nhập mã StoreID báo cáo trước.", "Cần Nhập Thông tin", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $instKey = [string]$cmbI.SelectedItem
        $instObj = $instMap[$instKey]

        $credTarget = Get-RequiredProperty -Object $script:CurrentConfig -Name 'CredentialTarget'
        $stored = Get-WindowsGenericCredential -Target $credTarget
        if ($null -eq $stored) {
            [System.Windows.Forms.MessageBox]::Show("Chưa có thông tin đăng nhập cho '$credTarget'. Vui lòng cài đặt tài khoản trước.", "Thiếu Tài khoản", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $btnInspect.Enabled = $false
        try {
            Gui-Log "Đang kết nối tới [$instKey] để dò tìm tham số & danh mục cho '$src'..."
            $http = New-CognosHttpClient -TimeoutMinutes 2
            $xsrf = Invoke-CognosLogin -Context $http -CognosBaseUrl $instObj.CognosBaseUrl -Namespace $instObj.Namespace -Username $stored.Username -Password $stored.Password
            $params = Get-CognosReportParameters -Context $http -BaseUrl $instObj.CognosBaseUrl -SourceType 'report' -Source $src -Xsrf $xsrf
            $http.Client.Dispose()
            $http.Handler.Dispose()

            $existingKeys = @{}
            for ($r = 0; $r -lt $gridParams.Rows.Count; $r++) {
                $k = [string]$gridParams.Rows[$r].Cells[0].Value
                if ($k) { $existingKeys[$k] = $gridParams.Rows[$r] }
            }

            $addedCount = 0
            $choiceCount = 0
            $reqCount = 0
            foreach ($k in $params.Keys) {
                $pInfo = $params[$k]
                $hasChoices = ($null -ne $pInfo.Choices -and @($pInfo.Choices).Count -gt 0)
                if ($hasChoices) { $choiceCount++ }
                if ($pInfo.IsRequired) { $reqCount++ }

                $reqText = if ($pInfo.IsRequired) { "BẮT BUỘC (*)" } else { "Tùy chọn" }

                $defaultVal = if ($pInfo.DefaultValue) {
                    if ($pInfo.DefaultValue -is [System.Collections.IEnumerable] -and -not ($pInfo.DefaultValue -is [string])) {
                        ($pInfo.DefaultValue | ForEach-Object { [string]$_ }) -join ', '
                    } else {
                        [string]$pInfo.DefaultValue
                    }
                } elseif ($k -match 'date|ngay|time|denhan|quahan') {
                    '{Yesterday}'
                } else {
                    ''
                }

                if ($existingKeys.ContainsKey($k)) {
                    $row = $existingKeys[$k]
                    $row.Tag = $pInfo
                    $row.Cells[1].Value = $reqText
                    if ($pInfo.IsRequired) {
                        $row.Cells[1].Style.ForeColor = [System.Drawing.Color]::FromArgb(217, 48, 37)
                    } else {
                        $row.Cells[1].Style.ForeColor = [System.Drawing.Color]::FromArgb(95, 99, 104)
                    }
                    if ([string]::IsNullOrWhiteSpace([string]$row.Cells[2].Value) -and $defaultVal) {
                        $row.Cells[2].Value = $defaultVal
                    }
                } else {
                    $rowIdx = $gridParams.Rows.Add($k, $reqText, $defaultVal)
                    $row = $gridParams.Rows[$rowIdx]
                    $row.Tag = $pInfo
                    if ($pInfo.IsRequired) {
                        $row.Cells[1].Style.ForeColor = [System.Drawing.Color]::FromArgb(217, 48, 37)
                    } else {
                        $row.Cells[1].Style.ForeColor = [System.Drawing.Color]::FromArgb(95, 99, 104)
                    }
                    $addedCount++
                }
            }

            Gui-Log "Đã phát hiện $($params.Count) tham số trên máy chủ [$instKey] ($reqCount bắt buộc, $choiceCount có danh mục)." 'OK'
            [System.Windows.Forms.MessageBox]::Show("Đã phát hiện $($params.Count) tham số:`n- $reqCount tham số bắt buộc (*)`n- $choiceCount tham số có danh mục lựa chọn từ Cognos server.`n`nBạn có thể nhấp đúp vào dòng tham số hoặc bấm 'Chọn từ Danh mục Server' để xem danh sách lựa chọn.", "Đã Tìm Thấy Tham Số & Danh Mục", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            Gui-Log "Dò tìm tham số thất bại: $($_.Exception.Message)" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show("Dò tìm tham số thất bại: $($_.Exception.Message)", "Lỗi", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
        finally {
            $btnInspect.Enabled = $true
        }
    })

    if ($dlg.ShowDialog($mainForm) -eq [System.Windows.Forms.DialogResult]::OK) {
        $repName = $txtN.Text.Trim()
        if (-not $repName) { $repName = "Report_" + $txtS.Text.Trim() }

        $paramDict = [ordered]@{}
        for ($r = 0; $r -lt $gridParams.Rows.Count; $r++) {
            $pk = [string]$gridParams.Rows[$r].Cells[0].Value
            $pv = [string]$gridParams.Rows[$r].Cells[2].Value
            if (-not [string]::IsNullOrWhiteSpace($pk)) {
                $paramDict[$pk.Trim()] = if ($pv) { $pv.Trim() } else { '' }
            }
        }

        $repObj = [ordered]@{
            Name       = $repName
            Instance   = [string]$cmbI.SelectedItem
            Source     = $txtS.Text.Trim()
            SourceType = "report"
            Enabled    = $chkEn.Checked
            Parameters = $paramDict
            Formats    = @(
                [ordered]@{
                    Format     = [string]$cmbFmt.SelectedItem
                    OutputPath = $txtPath.Text.Trim()
                }
            )
        }

        $list = [System.Collections.Generic.List[object]]::new()
        $reports = if ($script:CurrentConfig.PSObject.Properties['Reports']) { @($script:CurrentConfig.Reports) } else { @() }
        
        if ($isEdit) {
            foreach ($r in $reports) {
                if ($r -eq $ExistingReport) {
                    $list.Add($repObj)
                } else {
                    $list.Add($r)
                }
            }
        } else {
            foreach ($r in $reports) { $list.Add($r) }
            $list.Add($repObj)
        }

        Set-ObjectProperty -Object $script:CurrentConfig -Name 'Reports' -Value $list
        Save-CognosConfig -Path $ConfigPath -Config $script:CurrentConfig
        $script:CurrentConfig = Load-CognosConfig -Path $ConfigPath
        Refresh-ReportsGrid
        Gui-Log "Đã lưu báo cáo '$repName' vào cấu hình." 'OK'
    }
}

$btnAddReport.add_Click({ Show-ReportEditDialog })
$btnEditReport.add_Click({
    if ($gridReports.SelectedRows.Count -gt 0) {
        $rep = $gridReports.SelectedRows[0].Tag
        Show-ReportEditDialog -ExistingReport $rep
    }
})

$btnDeleteReport.add_Click({
    if ($gridReports.SelectedRows.Count -gt 0) {
        $rep = $gridReports.SelectedRows[0].Tag
        $delRepName = Get-PropOrKey -Object $rep -Name 'Name'
        if ([string]::IsNullOrWhiteSpace($delRepName)) { $delRepName = Get-PropOrKey -Object $rep -Name 'Source' }
        $confirm = [System.Windows.Forms.MessageBox]::Show("Bạn có chắc chắn muốn xóa báo cáo '$delRepName' không?", "Xác nhận Xóa", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            $list = [System.Collections.Generic.List[object]]::new()
            $currReports = Get-PropOrKey -Object $script:CurrentConfig -Name 'Reports'
            foreach ($r in @($currReports)) {
                if ($r -ne $rep) { $list.Add($r) }
            }
            Set-ObjectProperty -Object $script:CurrentConfig -Name 'Reports' -Value $list
            Save-CognosConfig -Path $ConfigPath -Config $script:CurrentConfig
            $script:CurrentConfig = Load-CognosConfig -Path $ConfigPath
            Refresh-ReportsGrid
            Gui-Log "Đã xóa báo cáo '$delRepName'." 'OK'
        }
    }
})

# Xử lý: Chạy Thử nghiệm Tải Báo cáo Ngay
$btnRunReport.add_Click({
    if ($gridReports.SelectedRows.Count -eq 0) { return }
    $rep = $gridReports.SelectedRows[0].Tag
    $btnRunReport.Enabled = $false

    try {
        $instObj = Get-CognosReportInstance -Config $script:CurrentConfig -Report $rep
        $credTarget = Get-RequiredProperty -Object $script:CurrentConfig -Name 'CredentialTarget'
        $stored = Get-WindowsGenericCredential -Target $credTarget
        if ($null -eq $stored) {
            [System.Windows.Forms.MessageBox]::Show("Vui lòng cấu hình tài khoản đăng nhập trước.", "Thiếu Tài khoản", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $repName = Get-PropOrKey -Object $rep -Name 'Name'
        $repSource = Get-PropOrKey -Object $rep -Name 'Source'
        if ([string]::IsNullOrWhiteSpace($repName)) { $repName = $repSource }

        Gui-Log "Bắt đầu tải báo cáo '$repName' trên máy chủ [$($instObj.Name)]..."
        $http = New-CognosHttpClient -TimeoutMinutes 10
        $xsrf = Invoke-CognosLogin -Context $http -CognosBaseUrl $instObj.CognosBaseUrl -Namespace $instObj.Namespace -Username $stored.Username -Password $stored.Password

        $repFormats = Get-PropOrKey -Object $rep -Name 'Formats'
        foreach ($fmtConfig in @($repFormats)) {
            $fmt = [string](Get-RequiredProperty -Object $fmtConfig -Name 'Format')
            $rawPath = [string](Get-RequiredProperty -Object $fmtConfig -Name 'OutputPath')
            $outPath = Resolve-DynamicTokens -Text $rawPath -Report $rep -Format $fmt

            $parent = Split-Path -Parent $outPath
            if ($parent -and -not (Test-Path -LiteralPath $parent)) {
                New-Item -ItemType Directory -Path $parent -Force | Out-Null
            }

            $url = Get-ReportDefinitionUrl -CognosBaseUrl $instObj.CognosBaseUrl -Report $rep -Format $fmt
            Gui-Log "Đang tải định dạng $fmt về $outPath..."

            $downloadResult = Invoke-CognosReportDownload `
                -Context $http `
                -Url $url `
                -Xsrf $xsrf `
                -Format $fmt

            $bytes = $downloadResult.Bytes
            $sizeMb = [Math]::Round($bytes.Length / 1MB, 2)
            [IO.File]::WriteAllBytes($outPath, $bytes)
            Gui-Log "THÀNH CÔNG: Đã lưu $outPath ($sizeMb MB)" 'OK'
            [System.Windows.Forms.MessageBox]::Show("Đã tải và lưu báo cáo thành công ($sizeMb MB):`n$outPath", "Tải Thành Công", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        $http.Client.Dispose()
        $http.Handler.Dispose()
    }
    catch {
        Gui-Log "Tải báo cáo thất bại: $($_.Exception.Message)" 'ERROR'
        [System.Windows.Forms.MessageBox]::Show("Tải báo cáo thất bại: $($_.Exception.Message)", "Lỗi Tải Báo Cáo", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $btnRunReport.Enabled = $true
    }
})

# Xử lý: Tải Tất cả Báo cáo Đang Kích hoạt
$btnRunAllReports.add_Click({
    $rawReports = Get-PropOrKey -Object $script:CurrentConfig -Name 'Reports'
    $reports = if ($null -ne $rawReports) { @($rawReports) } else { @() }
    $enabledReports = @($reports | Where-Object { $null -eq $_.Enabled -or $_.Enabled -ne $false })

    if ($enabledReports.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Không có báo cáo nào đang ở trạng thái kích hoạt (ĐANG BẬT).", "Không Có Báo Cáo", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $credTarget = Get-RequiredProperty -Object $script:CurrentConfig -Name 'CredentialTarget'
    $stored = Get-WindowsGenericCredential -Target $credTarget
    if ($null -eq $stored) {
        [System.Windows.Forms.MessageBox]::Show("Vui lòng cấu hình tài khoản đăng nhập trước khi tải báo cáo.", "Thiếu Tài khoản", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show("Bạn có muốn bắt đầu tải toàn bộ $($enabledReports.Count) báo cáo đang kích hoạt không?", "Xác nhận Tải Tất Cả", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $btnRunAllReports.Enabled = $false
    $btnRunReport.Enabled = $false
    $btnAddReport.Enabled = $false
    $btnEditReport.Enabled = $false
    $btnDeleteReport.Enabled = $false

    $sessions = @{}
    $total = 0
    $success = 0
    $failed = 0

    try {
        Gui-Log "=== BẮT ĐẦU TIẾN TRÌNH TẢI TẤT CẢ ($($enabledReports.Count)) BÁO CÁO ===" 'INFO'
        [System.Windows.Forms.Application]::DoEvents()

        foreach ($rep in $enabledReports) {
            $repName = Get-PropOrKey -Object $rep -Name 'Name'
            $repSource = Get-PropOrKey -Object $rep -Name 'Source'
            if ([string]::IsNullOrWhiteSpace($repName)) { $repName = $repSource }

            $instObj = try { Get-CognosReportInstance -Config $script:CurrentConfig -Report $rep } catch { $null }
            if ($null -eq $instObj) {
                Gui-Log "Bỏ qua '$repName': Không tìm thấy máy chủ liên kết." 'ERROR'
                $failed++
                continue
            }

            # Lấy hoặc tạo phiên kết nối máy chủ
            $sessionKey = "$($instObj.CognosBaseUrl)|$($instObj.Namespace)"
            if (-not $sessions.ContainsKey($sessionKey)) {
                try {
                    Gui-Log "Đang đăng nhập vào [$($instObj.Name)] ($($instObj.CognosBaseUrl))..." 'INFO'
                    [System.Windows.Forms.Application]::DoEvents()
                    $http = New-CognosHttpClient -TimeoutMinutes 15
                    $xsrf = Invoke-CognosLogin -Context $http -CognosBaseUrl $instObj.CognosBaseUrl -Namespace $instObj.Namespace -Username $stored.Username -Password $stored.Password
                    $sessions[$sessionKey] = [pscustomobject]@{ Http = $http; Xsrf = $xsrf; Instance = $instObj }
                    Gui-Log "Đăng nhập thành công [$($instObj.Name)]." 'OK'
                }
                catch {
                    Gui-Log "Đăng nhập thất bại vào [$($instObj.Name)]: $($_.Exception.Message)" 'ERROR'
                    $failed++
                    continue
                }
            }

            $currentSession = $sessions[$sessionKey]
            $repFormats = Get-PropOrKey -Object $rep -Name 'Formats'
            $formatList = if ($null -ne $repFormats) { @($repFormats) } else { @() }

            if ($formatList.Count -eq 0) {
                Gui-Log "Báo cáo '$repName' chưa cấu hình định dạng xuất." 'WARN'
                continue
            }

            foreach ($fmtConfig in $formatList) {
                $total++
                $fmt = [string](Get-RequiredProperty -Object $fmtConfig -Name 'Format')
                $rawPath = [string](Get-RequiredProperty -Object $fmtConfig -Name 'OutputPath')
                $outPath = Resolve-DynamicTokens -Text $rawPath -Report $rep -Format $fmt

                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                $httpStatus = 0
                $fileBytes = 0

                try {
                    $parent = Split-Path -Parent $outPath
                    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
                        New-Item -ItemType Directory -Path $parent -Force | Out-Null
                    }

                    $url = Get-ReportDefinitionUrl -CognosBaseUrl $instObj.CognosBaseUrl -Report $rep -Format $fmt
                    Gui-Log "Đang tải '$repName' [$($instObj.Name)] định dạng $fmt..." 'INFO'
                    [System.Windows.Forms.Application]::DoEvents()

                    $downloadResult = Invoke-CognosReportDownload `
                        -Context $currentSession.Http `
                        -Url $url `
                        -Xsrf $currentSession.Xsrf `
                        -Format $fmt

                    $bytes = $downloadResult.Bytes
                    $httpStatus = $downloadResult.HttpStatus
                    [IO.File]::WriteAllBytes($outPath, $bytes)
                    $fileBytes = $bytes.Length
                    $sw.Stop()

                    $sizeMb = [Math]::Round($fileBytes / 1MB, 2)
                    $durSec = [Math]::Round($sw.ElapsedMilliseconds / 1000, 2)
                    Gui-Log "Đã lưu $outPath ($sizeMb MB) trong ${durSec}s" 'OK'
                    Write-AuditLog -ReportName $repName -Source ([string]$rep.Source) -Format $fmt -Status 'SUCCESS' -HttpStatusCode $httpStatus -FileSizeBytes $fileBytes -DurationMs $sw.ElapsedMilliseconds -OutputPath $outPath

                    $success++
                }
                catch {
                    $sw.Stop()
                    $failed++
                    Gui-Log "Tải thất bại '$repName' / $fmt : $($_.Exception.Message)" 'ERROR'
                    Write-AuditLog -ReportName $repName -Source ([string]$rep.Source) -Format $fmt -Status 'FAILED' -HttpStatusCode $httpStatus -FileSizeBytes $fileBytes -DurationMs $sw.ElapsedMilliseconds -OutputPath $outPath -ErrorMessage $_.Exception.Message
                }

                [System.Windows.Forms.Application]::DoEvents()
            }
        }

        Gui-Log "=== HOÀN THÀNH: $success thành công, $failed thất bại (Tổng số định dạng: $total) ===" $(if ($failed -eq 0) { 'OK' } else { 'WARN' })
        [System.Windows.Forms.MessageBox]::Show("Hoàn thành tiến trình tải tất cả báo cáo:`n`n- Thành công: $success`n- Thất bại: $failed`n- Tổng định dạng: $total", "Kết Quả Tải Báo Cáo", [System.Windows.Forms.MessageBoxButtons]::OK, $(if ($failed -eq 0) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }))
    }
    finally {
        foreach ($s in $sessions.Values) {
            if ($s.Http) {
                if ($s.Http.Client) { $s.Http.Client.Dispose() }
                if ($s.Http.Handler) { $s.Http.Handler.Dispose() }
            }
        }
        $btnRunAllReports.Enabled = $true
        $btnRunReport.Enabled = $true
        $btnAddReport.Enabled = $true
        $btnEditReport.Enabled = $true
        $btnDeleteReport.Enabled = $true
    }
})

# -----------------------------------------------------------------------------
# Hộp thoại Thêm / Sửa Máy chủ Cognos
# -----------------------------------------------------------------------------

function Show-InstanceDialog {
    param([string]$Key = '', $ExistingInstance = $null)

    $isEdit = ($null -ne $ExistingInstance)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = if ($isEdit) { "Sửa Máy chủ Cognos: $Key" } else { "Thêm Máy chủ Cognos Mới" }
    $dlg.Size = New-Object System.Drawing.Size(520, 270)
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dlg.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $dlg.Padding = New-Object System.Windows.Forms.Padding(20)

    $pnlDlgContent = New-Object System.Windows.Forms.Panel
    $pnlDlgContent.Dock = [System.Windows.Forms.DockStyle]::Fill
    $dlg.Controls.Add($pnlDlgContent)

    $txtK = New-Object System.Windows.Forms.TextBox -Property @{ Text = $Key; ReadOnly = $isEdit }
    $rowK = New-SettingRowPanel -LabelText "Tên Máy chủ:" -InputControl $txtK
    $pnlDlgContent.Controls.Add($rowK)
    $rowK.SendToBack()

    $txtU = New-Object System.Windows.Forms.TextBox -Property @{ Text = if ($isEdit) { [string]$ExistingInstance.CognosBaseUrl } else { '' } }
    $rowU = New-SettingRowPanel -LabelText "Đường dẫn Base URL:" -InputControl $txtU
    $pnlDlgContent.Controls.Add($rowU)
    $rowU.SendToBack()

    $txtN = New-Object System.Windows.Forms.TextBox -Property @{ Text = if ($isEdit) { [string]$ExistingInstance.Namespace } else { '' } }
    $rowN = New-SettingRowPanel -LabelText "Không gian tên (Namespace):" -InputControl $txtN
    $pnlDlgContent.Controls.Add($rowN)
    $rowN.SendToBack()

    $pnlBtns = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Bottom; Height = 36; FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft }
    $btnCancel = New-Object System.Windows.Forms.Button -Property @{ Text = "Hủy bỏ"; Size = New-Object System.Drawing.Size(85, 30); DialogResult = [System.Windows.Forms.DialogResult]::Cancel }
    $btnOk = New-Object System.Windows.Forms.Button -Property @{ Text = "Lưu Máy chủ"; Size = New-Object System.Drawing.Size(120, 30); BackColor = [System.Drawing.Color]::FromArgb(26, 115, 232); ForeColor = [System.Drawing.Color]::White; FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
    $btnOk.FlatAppearance.BorderSize = 0

    $btnOk.add_Click({
        $instName = $txtK.Text.Trim()
        $instUrl  = $txtU.Text.Trim()
        $instNs   = $txtN.Text.Trim()

        if (-not $instName -or -not $instUrl -or -not $instNs) {
            [System.Windows.Forms.MessageBox]::Show("Tên máy chủ, URL và Namespace đều là bắt buộc.", "Lỗi Nhập liệu", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dlg.Close()
    })

    $pnlBtns.Controls.AddRange(@($btnCancel, $btnOk))
    $dlg.Controls.Add($pnlBtns)
    $pnlBtns.BringToFront()
    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel

    if ($dlg.ShowDialog($mainForm) -eq [System.Windows.Forms.DialogResult]::OK) {
        $instName = $txtK.Text.Trim()
        $instUrl  = $txtU.Text.Trim()
        $instNs   = $txtN.Text.Trim()

        if (-not $instName -or -not $instUrl -or -not $instNs) {
            [System.Windows.Forms.MessageBox]::Show("Tên máy chủ, URL và Namespace đều là bắt buộc.", "Lỗi Nhập liệu", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $instObj = [ordered]@{
            CognosBaseUrl = $instUrl.TrimEnd('/')
            Namespace     = $instNs
        }

        $existingMap = try { Get-CognosInstances -Config $script:CurrentConfig } catch { [ordered]@{} }
        $instancesDict = [ordered]@{}
        foreach ($k in $existingMap.Keys) {
            $inst = $existingMap[$k]
            $instancesDict[$k] = [ordered]@{
                CognosBaseUrl = $inst.CognosBaseUrl
                Namespace     = $inst.Namespace
            }
        }
        $instancesDict[$instName] = $instObj
        Set-ObjectProperty -Object $script:CurrentConfig -Name 'Instances' -Value $instancesDict
        Save-CognosConfig -Path $ConfigPath -Config $script:CurrentConfig
        Refresh-InstancesGrid
        Refresh-ReportsGrid
        Gui-Log "Đã lưu máy chủ Cognos [$instName]." 'OK'
    }
}

$btnAddInst.add_Click({ Show-InstanceDialog })
$btnEditInst.add_Click({
    if ($gridInstances.SelectedRows.Count -gt 0) {
        $k = [string]$gridInstances.SelectedRows[0].Cells[0].Value
        $instMap = Get-CognosInstances -Config $script:CurrentConfig
        if ($instMap.Contains($k)) {
            Show-InstanceDialog -Key $k -ExistingInstance $instMap[$k]
        }
    }
})

$btnSetDefaultInst.add_Click({
    if ($gridInstances.SelectedRows.Count -gt 0) {
        $k = [string]$gridInstances.SelectedRows[0].Cells[0].Value
        Set-ObjectProperty -Object $script:CurrentConfig -Name 'DefaultInstance' -Value $k
        Save-CognosConfig -Path $ConfigPath -Config $script:CurrentConfig
        Refresh-InstancesGrid
        Gui-Log "Đã đặt [$k] làm máy chủ Cognos mặc định." 'OK'
    }
})

$btnDeleteInst.add_Click({
    if ($gridInstances.SelectedRows.Count -gt 0) {
        $k = [string]$gridInstances.SelectedRows[0].Cells[0].Value
        $confirm = [System.Windows.Forms.MessageBox]::Show("Bạn có chắc chắn muốn xóa máy chủ [$k] không?", "Xác nhận Xóa", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            $existingMap = try { Get-CognosInstances -Config $script:CurrentConfig } catch { [ordered]@{} }
            $instancesDict = [ordered]@{}
            foreach ($existingKey in $existingMap.Keys) {
                if ($existingKey -ne $k) {
                    $inst = $existingMap[$existingKey]
                    $instancesDict[$existingKey] = [ordered]@{
                        CognosBaseUrl = $inst.CognosBaseUrl
                        Namespace     = $inst.Namespace
                    }
                }
            }
            Set-ObjectProperty -Object $script:CurrentConfig -Name 'Instances' -Value $instancesDict
            Save-CognosConfig -Path $ConfigPath -Config $script:CurrentConfig
            Refresh-InstancesGrid
            Refresh-ReportsGrid
            Gui-Log "Đã xóa máy chủ [$k]." 'OK'
        }
    }
})

# -----------------------------------------------------------------------------
# Khởi động & Khởi tạo Dữ liệu
# -----------------------------------------------------------------------------

function Initialize-Gui {
    Refresh-ReportsGrid
    Refresh-InstancesGrid
    Load-SettingsTab
    Gui-Log "Đã nạp cấu hình thành công từ: $ConfigPath" 'OK'
    $instMap = Get-CognosInstances -Config $script:CurrentConfig
    Gui-Log "Danh sách máy chủ đang hoạt động: $($instMap.Keys -join ', ')" 'INFO'
}

$mainForm.add_Shown({
    try {
        if ($splitReports.Width -gt 500) {
            $splitReports.SplitterDistance = [int]($splitReports.Width * 0.52)
        }
    } catch { }
})

Initialize-Gui

[void]$mainForm.ShowDialog()
