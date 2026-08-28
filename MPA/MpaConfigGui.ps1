#requires -Version 5.1
<#
.SYNOPSIS
    Giao diện đồ họa (Windows Forms) Quản lý Cấu hình & Tải Báo cáo MPA / OBIEE SOAP.
.DESCRIPTION
    Cung cấp giao diện desktop chuyên nghiệp:
    - Quản lý danh sách báo cáo phân tích OBIEE / MPA, đường dẫn catalog và định dạng xuất (CSV, EXCEL2007, PDF, MHT).
    - Trình duyệt Web Catalog trực quan (Catalog Explorer) với tính năng duyệt thư mục, tra cứu caption/type và chọn báo cáo.
    - Xem trước (Live Preview) giá trị dynamic token và đường dẫn lưu tệp trên đĩa.
    - Cấu hình đa máy chủ MPA dùng chung tài khoản Windows Credential Manager ('MaCB').
    - Tải thử nghiệm từng báo cáo hoặc tải hàng loạt (Batch Runner) với thanh tiến độ và nhật ký thực thi trực tiếp.
    - Quản lý tích hợp tác vụ lập lịch tự động Windows Task Scheduler.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = '.\mpa-reports.json'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($ConfigPath) -or $ConfigPath -eq '.\mpa-reports.json') {
    $ConfigPath = Join-Path $scriptDir 'mpa-reports.json'
}

# Nạp thư viện đồ họa Windows Forms và Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Nạp module dùng chung
. (Join-Path $scriptDir 'ObieeCommon.ps1')

# -----------------------------------------------------------------------------
# Trạng thái Toàn cục & Đọc Cấu hình
# -----------------------------------------------------------------------------

$script:CurrentConfig = $null

function Load-AppConfig {
    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        throw "Không tìm thấy tệp cấu hình: $ConfigPath. Vui lòng kiểm tra lại đường dẫn."
    }
    $script:CurrentConfig = Load-MpaConfig -Path $ConfigPath
    if ($null -eq $script:CurrentConfig) {
        throw "Không thể phân tích định dạng JSON của tệp: $ConfigPath"
    }
    $logCfg = if ($script:CurrentConfig.PSObject.Properties['Logging']) { $script:CurrentConfig.Logging } else { $null }
    Initialize-MpaLogging -LoggingConfig $logCfg -BaseDirectory $scriptDir
}

Load-AppConfig

# -----------------------------------------------------------------------------
# Khởi tạo Cửa sổ Chính
# -----------------------------------------------------------------------------

$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "MPA / OBIEE Report Downloader - Trình Quản lý Cấu hình"
$mainForm.Size = New-Object System.Drawing.Size(1120, 780)
$mainForm.MinimumSize = New-Object System.Drawing.Size(960, 660)
$mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$mainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$mainForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font

# -----------------------------------------------------------------------------
# 1. Thanh Công cụ Phía trên (Header Toolbar)
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

$btnSetCred = New-Object System.Windows.Forms.Button -Property @{
    Text = "Tài khoản Đăng nhập"
    Size = New-Object System.Drawing.Size(160, 32)
    Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
    BackColor = [System.Drawing.Color]::White
    ForeColor = [System.Drawing.Color]::FromArgb(60, 64, 67)
    FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
$btnSetCred.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(218, 220, 224)
$pnlTopButtons.Controls.Add($btnSetCred)

$btnTestAll = New-Object System.Windows.Forms.Button -Property @{
    Text = "Kiểm tra Kết nối"
    Size = New-Object System.Drawing.Size(130, 32)
    Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
    BackColor = [System.Drawing.Color]::FromArgb(232, 240, 254)
    ForeColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
    FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
$btnTestAll.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(190, 215, 250)
$pnlTopButtons.Controls.Add($btnTestAll)

$btnSave = New-Object System.Windows.Forms.Button -Property @{
    Text = "Lưu Cấu hình"
    Size = New-Object System.Drawing.Size(115, 32)
    Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
    BackColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
    ForeColor = [System.Drawing.Color]::White
    FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
$btnSave.FlatAppearance.BorderSize = 0
$pnlTopButtons.Controls.Add($btnSave)

$lblConfigPath = New-Object System.Windows.Forms.Label
$lblConfigPath.Text = "Tệp cấu hình: $ConfigPath"
$lblConfigPath.Dock = [System.Windows.Forms.DockStyle]::Fill
$lblConfigPath.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblConfigPath.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblConfigPath.ForeColor = [System.Drawing.Color]::FromArgb(60, 64, 67)
$lblConfigPath.UseMnemonic = $false
$topPanel.Controls.Add($lblConfigPath)

# -----------------------------------------------------------------------------
# 2. Khung Hiển thị Nhật ký Phía dưới (Bottom Log Panel)
# -----------------------------------------------------------------------------

$bottomPanel = New-Object System.Windows.Forms.Panel
$bottomPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$bottomPanel.Height = 140
$bottomPanel.Padding = New-Object System.Windows.Forms.Padding(16, 8, 16, 12)

$lblLogHeader = New-Object System.Windows.Forms.Label
$lblLogHeader.Text = "Nhật ký Hoạt động & Thực thi:"
$lblLogHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$lblLogHeader.Height = 24
$lblLogHeader.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
$lblLogHeader.ForeColor = [System.Drawing.Color]::FromArgb(90, 95, 100)
$lblLogHeader.UseMnemonic = $false

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
# 3. Hệ thống Tab Quản lý Chính (TabControl)
# -----------------------------------------------------------------------------

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabControl.Padding = New-Object System.Drawing.Point(14, 8)

$tabReports   = New-Object System.Windows.Forms.TabPage -Property @{ Text = "  Quản lý Báo cáo  "; Padding = New-Object System.Windows.Forms.Padding(12) }
$tabInstances = New-Object System.Windows.Forms.TabPage -Property @{ Text = "  Máy chủ MPA  "; Padding = New-Object System.Windows.Forms.Padding(12) }
$tabSchedule  = New-Object System.Windows.Forms.TabPage -Property @{ Text = "  Lập Lịch Tự Động  "; Padding = New-Object System.Windows.Forms.Padding(16) }
$tabSettings  = New-Object System.Windows.Forms.TabPage -Property @{ Text = "  Nhật ký & Cài đặt  "; Padding = New-Object System.Windows.Forms.Padding(16) }

$tabControl.Controls.Add($tabReports)
$tabControl.Controls.Add($tabInstances)
$tabControl.Controls.Add($tabSchedule)
$tabControl.Controls.Add($tabSettings)

$mainForm.Controls.Add($tabControl)
$tabControl.BringToFront()

# =============================================================================
# TAB 1: QUẢN LÝ BÁO CÁO (REPORTS TAB)
# =============================================================================

$splitReports = New-Object System.Windows.Forms.SplitContainer
$splitReports.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitReports.Orientation = [System.Windows.Forms.Orientation]::Vertical
$splitReports.SplitterDistance = 600
$splitReports.SplitterWidth = 6
$tabReports.Controls.Add($splitReports)

# Phía trái: Toolbar & Bảng Danh sách Báo cáo
$panelReportsLeft = New-Object System.Windows.Forms.Panel
$panelReportsLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitReports.Panel1.Controls.Add($panelReportsLeft)

$pnlReportButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlReportButtons.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlReportButtons.AutoSize = $true
$pnlReportButtons.MinimumSize = New-Object System.Drawing.Size(0, 42)
$pnlReportButtons.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$pnlReportButtons.WrapContents = $true
$panelReportsLeft.Controls.Add($pnlReportButtons)

$btnAddReport = New-Object System.Windows.Forms.Button -Property @{
    Text = "+ Thêm Báo cáo"
    Size = New-Object System.Drawing.Size(120, 32)
    Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
    BackColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
    ForeColor = [System.Drawing.Color]::White
    FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
$btnAddReport.FlatAppearance.BorderSize = 0

$btnBrowseCatalog = New-Object System.Windows.Forms.Button -Property @{
    Text = "Duyệt Catalog..."
    Size = New-Object System.Drawing.Size(120, 32)
    Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
    BackColor = [System.Drawing.Color]::FromArgb(241, 243, 244)
    ForeColor = [System.Drawing.Color]::FromArgb(60, 64, 67)
    FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
$btnBrowseCatalog.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(218, 220, 224)

$btnEditReport = New-Object System.Windows.Forms.Button -Property @{
    Text = "Sửa Báo cáo"
    Size = New-Object System.Drawing.Size(100, 32)
    Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
    BackColor = [System.Drawing.Color]::White
    ForeColor = [System.Drawing.Color]::FromArgb(60, 64, 67)
    FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
$btnEditReport.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(218, 220, 224)

$btnDeleteReport = New-Object System.Windows.Forms.Button -Property @{
    Text = "Xóa"
    Size = New-Object System.Drawing.Size(65, 32)
    Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
    BackColor = [System.Drawing.Color]::White
    ForeColor = [System.Drawing.Color]::FromArgb(217, 48, 37)
    FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
$btnDeleteReport.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(245, 198, 203)

$btnRunReport = New-Object System.Windows.Forms.Button -Property @{
    Text = "Tải Mục Chọn"
    Size = New-Object System.Drawing.Size(110, 32)
    Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
    BackColor = [System.Drawing.Color]::White
    ForeColor = [System.Drawing.Color]::FromArgb(19, 115, 51)
    FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
$btnRunReport.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(206, 234, 214)

$btnRunAllReports = New-Object System.Windows.Forms.Button -Property @{
    Text = "Tải Tất Cả Báo Cáo"
    Size = New-Object System.Drawing.Size(150, 32)
    Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
    BackColor = [System.Drawing.Color]::FromArgb(232, 240, 254)
    ForeColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
    FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
$btnRunAllReports.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(190, 215, 250)

$pnlReportButtons.Controls.AddRange(@($btnAddReport, $btnBrowseCatalog, $btnEditReport, $btnDeleteReport, $btnRunReport, $btnRunAllReports))

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
$gridReports.RowTemplate.Height = 32

$colName   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Name"; HeaderText = "Tên Báo cáo"; FillWeight = 140 }
$colInst   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Instance"; HeaderText = "Máy chủ"; FillWeight = 65 }
$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Status"; HeaderText = "Trạng thái"; FillWeight = 55 }
$colFmt    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Formats"; HeaderText = "Định dạng"; FillWeight = 60 }

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
$lblDetailsHeader.Height = 32
$lblDetailsHeader.ForeColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
$lblDetailsHeader.UseMnemonic = $false

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
        $path = Get-PropOrKey -Object $rep -Name 'Path'
        if ([string]::IsNullOrWhiteSpace($name)) { $name = $path }
        $instObj = try { Get-MpaReportInstance -Config $script:CurrentConfig -Report $rep } catch { $null }
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
        $txtDetails.Text = "Chưa có báo cáo nào được cấu hình trong tệp."
    }
}

function Update-ReportDetailsPreview {
    if ($gridReports.SelectedRows.Count -eq 0) {
        $txtDetails.Text = "Chưa chọn báo cáo."
        return
    }

    $rep = $gridReports.SelectedRows[0].Tag
    if ($null -eq $rep) { return }

    $instObj = try { Get-MpaReportInstance -Config $script:CurrentConfig -Report $rep } catch { $null }
    $instName = if ($instObj) { "$($instObj.Name) ($($instObj.Config.MpaBaseUrl))" } else { "<Lỗi: Máy chủ không xác định>" }

    $repName = Get-PropOrKey -Object $rep -Name 'Name'
    $repPath = Get-PropOrKey -Object $rep -Name 'Path'
    $repEnabled = Get-PropOrKey -Object $rep -Name 'Enabled'

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("TÊN BÁO CÁO:       $repName")
    [void]$sb.AppendLine("ĐƯỜNG DẪN CATALOG: $repPath")
    [void]$sb.AppendLine("MÁY CHỦ MPA:       $instName")
    [void]$sb.AppendLine("TRẠNG THÁI:        $(if ($null -ne $repEnabled -and $repEnabled -eq $false) { 'ĐÃ TẮT' } else { 'ĐANG BẬT' })")
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
            [void]$sb.AppendLine("    Đường dẫn:  $evalPath`n")
        }
    } else {
        [void]$sb.AppendLine("  (Không có định dạng nào)")
    }

    $paramsObj = Get-PropOrKey -Object $rep -Name 'Parameters'
    if ($null -ne $paramsObj) {
        [void]$sb.AppendLine("-----------------------------------------------------------")
        [void]$sb.AppendLine("THAM SỐ & BỘ LỌC BÁO CÁO (VARIABLES / PARAMETERS):")
        [void]$sb.AppendLine("-----------------------------------------------------------")
        $pairs = Get-ObjectKeyValuePairs -Object $paramsObj
        if ($pairs.Count -gt 0) {
            foreach ($p in $pairs) {
                $evalArray = Resolve-DynamicTokenArray -Value $p.Value -DefaultDateFormat 'MM/dd/yyyy'
                if ($evalArray.Count -gt 1) {
                    [void]$sb.AppendLine("  * $($p.Name) = $($p.Value)  -->  [$($evalArray.Count) mốc ngày: $($evalArray[0]) ... $($evalArray[-1])]")
                    foreach ($item in $evalArray) {
                        [void]$sb.AppendLine("      - $item")
                    }
                } else {
                    $singleVal = if ($evalArray.Count -eq 1) { $evalArray[0] } else { '' }
                    [void]$sb.AppendLine("  * $($p.Name) = $($p.Value)  -->  $singleVal")
                }
            }
        } else {
            [void]$sb.AppendLine("  (Không có tham số)")
        }
    }

    $txtDetails.Text = $sb.ToString()
}

$gridReports.Add_SelectionChanged({ Update-ReportDetailsPreview })

# =============================================================================
# TAB 2: CẤU HÌNH MÁY CHỦ MPA (INSTANCES TAB)
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

$btnAddInst = New-Object System.Windows.Forms.Button -Property @{
    Text = "+ Thêm Máy chủ"
    Size = New-Object System.Drawing.Size(125, 32)
    Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
    BackColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
    ForeColor = [System.Drawing.Color]::White
    FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
$btnAddInst.FlatAppearance.BorderSize = 0

$btnEditInst = New-Object System.Windows.Forms.Button -Property @{
    Text = "Sửa Máy chủ"
    Size = New-Object System.Drawing.Size(105, 32)
    Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
    BackColor = [System.Drawing.Color]::White
    ForeColor = [System.Drawing.Color]::FromArgb(60, 64, 67)
    FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
$btnEditInst.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(218, 220, 224)

$btnDeleteInst = New-Object System.Windows.Forms.Button -Property @{
    Text = "Xóa"
    Size = New-Object System.Drawing.Size(65, 32)
    Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
    BackColor = [System.Drawing.Color]::White
    ForeColor = [System.Drawing.Color]::FromArgb(217, 48, 37)
    FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
$btnDeleteInst.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(245, 198, 203)

$btnSetDefaultInst = New-Object System.Windows.Forms.Button -Property @{
    Text = "Đặt làm Mặc định"
    Size = New-Object System.Drawing.Size(140, 32)
    Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
    BackColor = [System.Drawing.Color]::FromArgb(232, 240, 254)
    ForeColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
    FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}
$btnSetDefaultInst.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(190, 215, 250)

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
$gridInstances.RowTemplate.Height = 32

$colInstKey = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "Key"; HeaderText = "Tên Máy chủ"; FillWeight = 70 }
$colInstUrl = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "BaseUrl"; HeaderText = "Đường dẫn SOAP URL"; FillWeight = 160 }
$colInstDef = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "IsDefault"; HeaderText = "Mặc định"; FillWeight = 50 }

[void]$gridInstances.Columns.Add($colInstKey)
[void]$gridInstances.Columns.Add($colInstUrl)
[void]$gridInstances.Columns.Add($colInstDef)
$panelInst.Controls.Add($gridInstances)
$gridInstances.BringToFront()

function Refresh-InstancesGrid {
    $gridInstances.Rows.Clear()
    $instMap = Get-MpaInstances -Config $script:CurrentConfig
    $defaultName = if ($script:CurrentConfig.PSObject.Properties['DefaultInstance'] -and $script:CurrentConfig.DefaultInstance) { [string]$script:CurrentConfig.DefaultInstance } else { '' }

    foreach ($k in $instMap.Keys) {
        $inst = $instMap[$k]
        $isDef = if ($defaultName -and $k -ieq $defaultName) { "MẶC ĐỊNH" } else { "" }
        [void]$gridInstances.Rows.Add($k, $inst.MpaBaseUrl, $isDef)
    }
}

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
    $lbl.Width = 220
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

# =============================================================================
# TAB 3: LẬP LỊCH TỰ ĐỘNG (WINDOWS TASK SCHEDULER)
# =============================================================================

$panelSchedule = New-Object System.Windows.Forms.Panel
$panelSchedule.Dock = [System.Windows.Forms.DockStyle]::Fill
$panelSchedule.AutoScroll = $true
$tabSchedule.Controls.Add($panelSchedule)

# Thẻ 1: Trạng thái Lịch hiện tại
$pnlSchedStatusCard = New-Object System.Windows.Forms.Panel
$pnlSchedStatusCard.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSchedStatusCard.Height = 160
$pnlSchedStatusCard.BackColor = [System.Drawing.Color]::FromArgb(250, 252, 255)
$pnlSchedStatusCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlSchedStatusCard.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)
$panelSchedule.Controls.Add($pnlSchedStatusCard)

$lblSchedStatusHeader = New-Object System.Windows.Forms.Label
$lblSchedStatusHeader.Text = "Trạng thái Lập Lịch Tự Động (Windows Task Scheduler)"
$lblSchedStatusHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$lblSchedStatusHeader.Height = 26
$lblSchedStatusHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$lblSchedStatusHeader.ForeColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
$lblSchedStatusHeader.UseMnemonic = $false
$pnlSchedStatusCard.Controls.Add($lblSchedStatusHeader)

$lblSchedState = New-Object System.Windows.Forms.Label -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 24; Font = New-Object System.Drawing.Font("Segoe UI", 9) }
$lblSchedLastRun = New-Object System.Windows.Forms.Label -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 24; Font = New-Object System.Drawing.Font("Segoe UI", 9) }
$lblSchedNextRun = New-Object System.Windows.Forms.Label -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 24; Font = New-Object System.Drawing.Font("Segoe UI", 9) }
$lblSchedAction = New-Object System.Windows.Forms.Label -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 24; Font = New-Object System.Drawing.Font("Segoe UI", 9); ForeColor = [System.Drawing.Color]::FromArgb(90, 95, 100) }

$pnlSchedStatusCard.Controls.Add($lblSchedAction)
$pnlSchedStatusCard.Controls.Add($lblSchedNextRun)
$pnlSchedStatusCard.Controls.Add($lblSchedLastRun)
$pnlSchedStatusCard.Controls.Add($lblSchedState)
$lblSchedState.BringToFront()
$lblSchedLastRun.BringToFront()
$lblSchedNextRun.BringToFront()
$lblSchedAction.BringToFront()

# Khoảng đệm
$pnlSchedSpacer = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 16 }
$panelSchedule.Controls.Add($pnlSchedSpacer)
$pnlSchedSpacer.BringToFront()

# Thẻ 2: Cài đặt Lịch mới
$pnlSchedConfigCard = New-Object System.Windows.Forms.Panel
$pnlSchedConfigCard.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSchedConfigCard.Height = 220
$pnlSchedConfigCard.BackColor = [System.Drawing.Color]::FromArgb(250, 252, 255)
$pnlSchedConfigCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlSchedConfigCard.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)
$panelSchedule.Controls.Add($pnlSchedConfigCard)
$pnlSchedConfigCard.BringToFront()

$lblSchedConfigHeader = New-Object System.Windows.Forms.Label
$lblSchedConfigHeader.Text = "Cài đặt & Lập Lịch Chạy Định Kỳ"
$lblSchedConfigHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$lblSchedConfigHeader.Height = 26
$lblSchedConfigHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$lblSchedConfigHeader.ForeColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
$lblSchedConfigHeader.UseMnemonic = $false
$pnlSchedConfigCard.Controls.Add($lblSchedConfigHeader)

# Dòng 1: Tên tác vụ
$txtTaskName = New-Object System.Windows.Forms.TextBox -Property @{ Text = "MpaReportDownloader" }
$rowTaskName = New-SettingRowPanel -LabelText "Tên Tác vụ (Task Name):" -InputControl $txtTaskName
$pnlSchedConfigCard.Controls.Add($rowTaskName)
$rowTaskName.BringToFront()

# Dòng 2: Tần suất chạy
$cmbSchedFreq = New-Object System.Windows.Forms.ComboBox -Property @{ DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList }
[void]$cmbSchedFreq.Items.AddRange(@(
    'Hàng ngày (Daily)',
    'Các ngày làm việc (Thứ 2 - Thứ 6)',
    'Hàng tuần (Weekly - Thứ 2)',
    'Hàng giờ (Hourly)'
))
$cmbSchedFreq.SelectedIndex = 0
$rowSchedFreq = New-SettingRowPanel -LabelText "Tần suất thực thi:" -InputControl $cmbSchedFreq
$pnlSchedConfigCard.Controls.Add($rowSchedFreq)
$rowSchedFreq.BringToFront()

# Dòng 3: Giờ bắt đầu
$txtSchedTime = New-Object System.Windows.Forms.TextBox -Property @{ Text = "08:00" }
$rowSchedTime = New-SettingRowPanel -LabelText "Giờ bắt đầu chạy (HH:mm):" -InputControl $txtSchedTime
$pnlSchedConfigCard.Controls.Add($rowSchedTime)
$rowSchedTime.BringToFront()

# Dòng 4: Quyền thực thi
$pnlSchedElevated = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 32 }
$chkSchedElevated = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Chạy với quyền Administrator cao nhất (Highest Privileges)"; Dock = [System.Windows.Forms.DockStyle]::Fill; Checked = $false }
$pnlSchedElevated.Controls.Add($chkSchedElevated)
$pnlSchedConfigCard.Controls.Add($pnlSchedElevated)
$pnlSchedElevated.BringToFront()

# Khoảng đệm
$pnlSchedSpacer2 = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 16 }
$panelSchedule.Controls.Add($pnlSchedSpacer2)
$pnlSchedSpacer2.BringToFront()

# Thanh nút điều khiển Lịch
$pnlSchedButtons = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 46; FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight; WrapContents = $false }
$btnSetSchedule = New-Object System.Windows.Forms.Button -Property @{ Text = "Lưu & Kích hoạt Lịch"; Size = New-Object System.Drawing.Size(160, 34); BackColor = [System.Drawing.Color]::FromArgb(26, 115, 232); ForeColor = [System.Drawing.Color]::White; FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0) }
$btnSetSchedule.FlatAppearance.BorderSize = 0

$btnRunScheduleNow = New-Object System.Windows.Forms.Button -Property @{ Text = "Chạy Thử Tác Vụ Ngay"; Size = New-Object System.Drawing.Size(160, 34); BackColor = [System.Drawing.Color]::FromArgb(230, 244, 234); ForeColor = [System.Drawing.Color]::FromArgb(19, 115, 51); FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0) }
$btnRunScheduleNow.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(206, 234, 214)

$btnOpenTaskScheduler = New-Object System.Windows.Forms.Button -Property @{ Text = "Mở Task Scheduler GUI"; Size = New-Object System.Drawing.Size(160, 34); BackColor = [System.Drawing.Color]::White; ForeColor = [System.Drawing.Color]::FromArgb(60, 64, 67); FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0) }
$btnOpenTaskScheduler.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(218, 220, 224)

$btnDeleteSchedule = New-Object System.Windows.Forms.Button -Property @{ Text = "Xóa Lịch Tác Vụ"; Size = New-Object System.Drawing.Size(130, 34); BackColor = [System.Drawing.Color]::White; ForeColor = [System.Drawing.Color]::FromArgb(217, 48, 37); FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0) }
$btnDeleteSchedule.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(245, 198, 203)

$btnRefreshSchedule = New-Object System.Windows.Forms.Button -Property @{ Text = "Làm mới"; Size = New-Object System.Drawing.Size(85, 34); BackColor = [System.Drawing.Color]::White; ForeColor = [System.Drawing.Color]::FromArgb(60, 64, 67); FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
$btnRefreshSchedule.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(218, 220, 224)

$pnlSchedButtons.Controls.AddRange(@($btnSetSchedule, $btnRunScheduleNow, $btnOpenTaskScheduler, $btnDeleteSchedule, $btnRefreshSchedule))
$panelSchedule.Controls.Add($pnlSchedButtons)
$pnlSchedButtons.BringToFront()

function Refresh-ScheduleTab {
    $schedCfg = Get-MpaSchedulingConfig -Config $script:CurrentConfig
    $tName = $txtTaskName.Text.Trim()
    if (-not $tName) {
        $tName = $schedCfg.TaskName
        $txtTaskName.Text = $tName
    }
    $tInfo = Get-MpaScheduledTask -TaskName $tName

    if ($null -ne $tInfo -and $tInfo.Exists) {
        $lblSchedState.Text = "Trạng thái: " + $tInfo.State + " (Loại lịch: " + $tInfo.ScheduleType + ", Giờ chạy: " + $tInfo.StartTime + ")"
        $lblSchedState.ForeColor = [System.Drawing.Color]::FromArgb(19, 115, 51)
        
        $lastRunStr = if ($tInfo.LastRunTime) { $tInfo.LastRunTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "Chưa chạy" }
        $lblSchedLastRun.Text = "Lần chạy gần nhất: $lastRunStr (Mã kết quả: $($tInfo.LastTaskResult))"
        
        $nextRunStr = if ($tInfo.NextRunTime) { $tInfo.NextRunTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "Không xác định" }
        $lblSchedNextRun.Text = "Lần chạy tiếp theo: $nextRunStr"

        $lblSchedAction.Text = "Đường dẫn thực thi: powershell.exe -> MpaReportDownloader.ps1"
        $btnDeleteSchedule.Enabled = $true
        $btnRunScheduleNow.Enabled = $true
    } else {
        $lblSchedState.Text = "Trạng thái: Chưa thiết lập lịch tự động cho '$tName'"
        $lblSchedState.ForeColor = [System.Drawing.Color]::FromArgb(128, 134, 139)
        $lblSchedLastRun.Text = "Lần chạy gần nhất: --"
        $lblSchedNextRun.Text = "Lần chạy tiếp theo: --"
        $lblSchedAction.Text = "Bấm 'Lưu & Kích hoạt Lịch' để tự động tạo Task trong Windows Task Scheduler."
        $btnDeleteSchedule.Enabled = $false
        $btnRunScheduleNow.Enabled = $false
    }
}

$btnSetSchedule.Add_Click({
    $tName = $txtTaskName.Text.Trim()
    if (-not $tName) {
        [System.Windows.Forms.MessageBox]::Show("Vui lòng nhập tên tác vụ.", "Thiếu Tên Tác Vụ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $timeStr = $txtSchedTime.Text.Trim()
    if ($timeStr -notmatch '^\d{1,2}:\d{2}$') {
        [System.Windows.Forms.MessageBox]::Show("Giờ chạy không đúng định dạng HH:mm (VD: 08:00, 09:30).", "Sai Định Dạng Giờ", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $freq = switch ($cmbSchedFreq.SelectedIndex) {
        0 { 'DAILY' }
        1 { 'WEEKDAY' }
        2 { 'WEEKLY' }
        3 { 'HOURLY' }
        default { 'DAILY' }
    }

    try {
        $scriptPath = Join-Path -Path $scriptDir -ChildPath 'MpaReportDownloader.ps1'
        Set-MpaScheduledTask `
            -TaskName $tName `
            -ScriptPath $scriptPath `
            -ConfigPath $ConfigPath `
            -ScheduleType $freq `
            -StartTime $timeStr `
            -RunElevated $chkSchedElevated.Checked | Out-Null

        Set-ObjectProperty -Object $script:CurrentConfig -Name 'Scheduling' -Value ([ordered]@{
            TaskName     = $tName
            ScheduleType = $freq
            StartTime    = $timeStr
            RunElevated  = $chkSchedElevated.Checked
        })
        Save-MpaConfig -Path $ConfigPath -Config $script:CurrentConfig

        Gui-Log "Đã thiết lập thành công lịch tự động '$tName' ($freq lúc $timeStr)." 'OK'
        Refresh-ScheduleTab
        [System.Windows.Forms.MessageBox]::Show("Đã cài đặt thành công lịch tác vụ Windows Task Scheduler!`n`n- Tên: $tName`n- Tần suất: $freq`n- Giờ chạy: $timeStr`n`nCấu hình đã được lưu vào tệp JSON.", "Lập Lịch Thành Công", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        Gui-Log "Lập lịch thất bại: $($_.Exception.Message)" 'ERROR'
        [System.Windows.Forms.MessageBox]::Show("Lập lịch thất bại: $($_.Exception.Message)", "Lỗi Lập Lịch", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnRunScheduleNow.Add_Click({
    $tName = $txtTaskName.Text.Trim()
    try {
        Start-MpaScheduledTask -TaskName $tName | Out-Null
        Gui-Log "Đã gửi tín hiệu chạy tác vụ '$tName'." 'OK'
        [System.Windows.Forms.MessageBox]::Show("Tác vụ '$tName' đã được kích hoạt chạy trong nền bởi Windows Task Scheduler.", "Đã Kích Hoạt", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        Refresh-ScheduleTab
    }
    catch {
        Gui-Log "Kích hoạt tác vụ thất bại: $($_.Exception.Message)" 'ERROR'
        [System.Windows.Forms.MessageBox]::Show("Kích hoạt tác vụ thất bại: $($_.Exception.Message)", "Lỗi", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnOpenTaskScheduler.Add_Click({
    try { Start-Process "taskschd.msc" } catch { [System.Windows.Forms.MessageBox]::Show("Không thể mở Task Scheduler: $($_.Exception.Message)", "Lỗi", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
})

$btnDeleteSchedule.Add_Click({
    $tName = $txtTaskName.Text.Trim()
    $confirm = [System.Windows.Forms.MessageBox]::Show("Bạn có chắc chắn muốn xóa lịch tác vụ '$tName' không?", "Xác nhận Xóa Lịch", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Remove-MpaScheduledTask -TaskName $tName | Out-Null
            Gui-Log "Đã xóa lịch tác vụ '$tName'." 'OK'
            Refresh-ScheduleTab
            [System.Windows.Forms.MessageBox]::Show("Đã xóa lịch tác vụ '$tName' khỏi Windows Task Scheduler.", "Đã Xóa", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            Gui-Log "Xóa lịch thất bại: $($_.Exception.Message)" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show("Xóa lịch thất bại: $($_.Exception.Message)", "Lỗi", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

$btnRefreshSchedule.Add_Click({ Refresh-ScheduleTab })

# =============================================================================
# TAB 4: NHẬT KÝ & CÀI ĐẶT (SETTINGS TAB)
# =============================================================================

$panelSettings = New-Object System.Windows.Forms.Panel
$panelSettings.Dock = [System.Windows.Forms.DockStyle]::Fill
$panelSettings.AutoScroll = $true
$tabSettings.Controls.Add($panelSettings)

# Thẻ 1: Cấu hình Tài khoản & Mạng HTTP
$pnlCredCard = New-Object System.Windows.Forms.Panel
$pnlCredCard.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlCredCard.Height = 150
$pnlCredCard.BackColor = [System.Drawing.Color]::FromArgb(250, 252, 255)
$pnlCredCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlCredCard.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)
$panelSettings.Controls.Add($pnlCredCard)

$lblCredCardHeader = New-Object System.Windows.Forms.Label
$lblCredCardHeader.Text = "Cấu hình Tài khoản Xác thực & Mạng HTTP"
$lblCredCardHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$lblCredCardHeader.Height = 24
$lblCredCardHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$lblCredCardHeader.ForeColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
$lblCredCardHeader.UseMnemonic = $false
$pnlCredCard.Controls.Add($lblCredCardHeader)

$txtCredTarget = New-Object System.Windows.Forms.TextBox -Property @{ Text = Get-PropOrKey -Object $script:CurrentConfig -Name 'CredentialTarget' -Default 'MaCB' }
$rowCred = New-SettingRowPanel -LabelText "Tên Target xác thực (Cred Manager):" -InputControl $txtCredTarget
$pnlCredCard.Controls.Add($rowCred)
$rowCred.BringToFront()

$txtDefInst = New-Object System.Windows.Forms.TextBox -Property @{ Text = Get-PropOrKey -Object $script:CurrentConfig -Name 'DefaultInstance' -Default 'MPA' }
$rowDefInst = New-SettingRowPanel -LabelText "Máy chủ mặc định:" -InputControl $txtDefInst
$pnlCredCard.Controls.Add($rowDefInst)
$rowDefInst.BringToFront()

$numHttpTimeout = New-Object System.Windows.Forms.NumericUpDown -Property @{ Width = 80; Minimum = 1; Maximum = 120; Value = 5 }
$httpCfg = Get-PropOrKey -Object $script:CurrentConfig -Name 'HttpSettings'
if ($null -ne $httpCfg) { $numHttpTimeout.Value = [int](Get-PropOrKey -Object $httpCfg -Name 'TimeoutMinutes' -Default 5) }
$pnlTimeoutCont = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Fill }
$pnlTimeoutCont.Controls.Add($numHttpTimeout)
$rowTimeout = New-SettingRowPanel -LabelText "Thời gian chờ HTTP (Phút):" -InputControl $pnlTimeoutCont
$pnlCredCard.Controls.Add($rowTimeout)
$rowTimeout.BringToFront()

# Khoảng đệm
$pnlSetSpacer = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 16 }
$panelSettings.Controls.Add($pnlSetSpacer)
$pnlSetSpacer.BringToFront()

# Thẻ 2: Hệ thống Nhật ký & Báo cáo Audit
$pnlLogCard = New-Object System.Windows.Forms.Panel
$pnlLogCard.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlLogCard.Height = 220
$pnlLogCard.BackColor = [System.Drawing.Color]::FromArgb(250, 252, 255)
$pnlLogCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlLogCard.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)
$panelSettings.Controls.Add($pnlLogCard)
$pnlLogCard.BringToFront()

$lblLogCardHeader = New-Object System.Windows.Forms.Label
$lblLogCardHeader.Text = "Cài đặt Nhật ký (Logging) & Báo cáo Thực thi"
$lblLogCardHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$lblLogCardHeader.Height = 24
$lblLogCardHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$lblLogCardHeader.ForeColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
$lblLogCardHeader.UseMnemonic = $false
$pnlLogCard.Controls.Add($lblLogCardHeader)

$logCfg = Get-PropOrKey -Object $script:CurrentConfig -Name 'Logging'

$chkLogEnabled = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Kích hoạt ghi nhật ký ra tệp (File Logging)"; Checked = [bool](Get-PropOrKey -Object $logCfg -Name 'Enabled' -Default $true); Dock = [System.Windows.Forms.DockStyle]::Fill }
$pnlLogEnCont = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Fill }
$pnlLogEnCont.Controls.Add($chkLogEnabled)
$rowLogEn = New-SettingRowPanel -LabelText "Trạng thái ghi log:" -InputControl $pnlLogEnCont
$pnlLogCard.Controls.Add($rowLogEn)
$rowLogEn.BringToFront()

$txtLogDir = New-Object System.Windows.Forms.TextBox -Property @{ Text = [string](Get-PropOrKey -Object $logCfg -Name 'LogDirectory' -Default '.\Logs') }
$rowLogDir = New-SettingRowPanel -LabelText "Thư mục lưu log:" -InputControl $txtLogDir
$pnlLogCard.Controls.Add($rowLogDir)
$rowLogDir.BringToFront()

$numRetention = New-Object System.Windows.Forms.NumericUpDown -Property @{ Width = 80; Minimum = 1; Maximum = 365; Value = [int](Get-PropOrKey -Object $logCfg -Name 'RetentionDays' -Default 30) }
$pnlRetCont = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Fill }
$pnlRetCont.Controls.Add($numRetention)
$rowRetention = New-SettingRowPanel -LabelText "Thời gian lưu log (Ngày):" -InputControl $pnlRetCont
$pnlLogCard.Controls.Add($rowRetention)
$rowRetention.BringToFront()

$chkAuditCsv = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Tạo tệp bảng tính CSV tổng kết hàng tháng (Audit CSV)"; Checked = [bool](Get-PropOrKey -Object $logCfg -Name 'AuditCsvEnabled' -Default $true); Dock = [System.Windows.Forms.DockStyle]::Fill }
$pnlAuditCont = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Fill }
$pnlAuditCont.Controls.Add($chkAuditCsv)
$rowAudit = New-SettingRowPanel -LabelText "Nhật ký kiểm toán (Audit):" -InputControl $pnlAuditCont
$pnlLogCard.Controls.Add($rowAudit)
$rowAudit.BringToFront()

# Khoảng đệm
$pnlSetSpacer3 = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 16 }
$panelSettings.Controls.Add($pnlSetSpacer3)
$pnlSetSpacer3.BringToFront()

# Thẻ 3: Chính sách Thử lại (Retry Policy)
$pnlRetryCard = New-Object System.Windows.Forms.Panel
$pnlRetryCard.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlRetryCard.Height = 180
$pnlRetryCard.BackColor = [System.Drawing.Color]::FromArgb(250, 252, 255)
$pnlRetryCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlRetryCard.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)
$panelSettings.Controls.Add($pnlRetryCard)
$pnlRetryCard.BringToFront()

$lblRetryCardHeader = New-Object System.Windows.Forms.Label
$lblRetryCardHeader.Text = "Chính sách Thử lại Tự động khi Gặp sự cố Mạng (Retry Policy)"
$lblRetryCardHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$lblRetryCardHeader.Height = 24
$lblRetryCardHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$lblRetryCardHeader.ForeColor = [System.Drawing.Color]::FromArgb(26, 115, 232)
$lblRetryCardHeader.UseMnemonic = $false
$pnlRetryCard.Controls.Add($lblRetryCardHeader)

$rpCfg = Get-MpaRetryPolicy -Config $script:CurrentConfig

$chkRetryEn = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Tự động thử lại khi tải thất bại hoặc ngắt kết nối"; Checked = $rpCfg.Enabled; Dock = [System.Windows.Forms.DockStyle]::Fill }
$pnlRetryEnCont = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Fill }
$pnlRetryEnCont.Controls.Add($chkRetryEn)
$rowRetryEn = New-SettingRowPanel -LabelText "Kích hoạt thử lại:" -InputControl $pnlRetryEnCont
$pnlRetryCard.Controls.Add($rowRetryEn)
$rowRetryEn.BringToFront()

$numMaxRetries = New-Object System.Windows.Forms.NumericUpDown -Property @{ Width = 80; Minimum = 1; Maximum = 10; Value = $rpCfg.MaxRetries }
$pnlMaxRCont = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Fill }
$pnlMaxRCont.Controls.Add($numMaxRetries)
$rowMaxR = New-SettingRowPanel -LabelText "Số lần thử lại tối đa:" -InputControl $pnlMaxRCont
$pnlRetryCard.Controls.Add($rowMaxR)
$rowMaxR.BringToFront()

$numInitDelay = New-Object System.Windows.Forms.NumericUpDown -Property @{ Width = 80; Minimum = 1; Maximum = 60; Value = $rpCfg.InitialDelaySeconds }
$pnlDelayCont = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Fill }
$pnlDelayCont.Controls.Add($numInitDelay)
$rowDelay = New-SettingRowPanel -LabelText "Thời gian chờ lần đầu (Giây):" -InputControl $pnlDelayCont
$pnlRetryCard.Controls.Add($rowDelay)
$rowDelay.BringToFront()

# -----------------------------------------------------------------------------
# TRÌNH DUYỆT CATALOG ĐỒ HỌA (NULL-SAFE CATALOG EXPLORER DIALOG)
# -----------------------------------------------------------------------------

function Show-CatalogBrowserDialog {
    param([string]$InstanceKey)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Duyệt Web Catalog OBIEE [$InstanceKey]"
    $dlg.Size = New-Object System.Drawing.Size(760, 560)
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $pnlTop = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Top; Height = 46; Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 8) }
    $dlg.Controls.Add($pnlTop)

    $btnUp = New-Object System.Windows.Forms.Button -Property @{ Text = "Lên (..)"; Dock = [System.Windows.Forms.DockStyle]::Right; Width = 80; Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0) }
    $btnGo = New-Object System.Windows.Forms.Button -Property @{ Text = "Đi tới"; Dock = [System.Windows.Forms.DockStyle]::Right; Width = 70; Margin = New-Object System.Windows.Forms.Padding(6, 0, 6, 0) }
    $txtCurPath = New-Object System.Windows.Forms.TextBox -Property @{ Dock = [System.Windows.Forms.DockStyle]::Fill }

    $pnlTop.Controls.Add($txtCurPath)
    $pnlTop.Controls.Add($btnGo)
    $pnlTop.Controls.Add($btnUp)
    $txtCurPath.BringToFront()

    $lstItems = New-Object System.Windows.Forms.ListView -Property @{
        Dock = [System.Windows.Forms.DockStyle]::Fill
        View = [System.Windows.Forms.View]::Details
        FullRowSelect = $true
        GridLines = $true
        MultiSelect = $false
    }
    [void]$lstItems.Columns.Add("Tên / Caption", 220)
    [void]$lstItems.Columns.Add("Kiểu", 90)
    [void]$lstItems.Columns.Add("Chữ ký (Signature)", 120)
    [void]$lstItems.Columns.Add("Đường dẫn Catalog", 290)
    $dlg.Controls.Add($lstItems)

    $pnlBottom = New-Object System.Windows.Forms.Panel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Bottom; Height = 48; Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 8) }
    $dlg.Controls.Add($pnlBottom)

    $btnCls = New-Object System.Windows.Forms.Button -Property @{ Text = "Đóng"; Dock = [System.Windows.Forms.DockStyle]::Right; Width = 85; DialogResult = [System.Windows.Forms.DialogResult]::Cancel }
    $btnSelect = New-Object System.Windows.Forms.Button -Property @{ Text = "Chọn Mục Này"; Dock = [System.Windows.Forms.DockStyle]::Right; Width = 120; BackColor = [System.Drawing.Color]::FromArgb(26, 115, 232); ForeColor = [System.Drawing.Color]::White; FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0) }
    $btnSelect.FlatAppearance.BorderSize = 0

    $pnlBottom.Controls.Add($btnSelect)
    $pnlBottom.Controls.Add($btnCls)

    $instDict = Get-MpaInstances -Config $script:CurrentConfig
    if (-not $instDict.Contains($InstanceKey)) {
        if (@($instDict.Keys).Count -gt 0) { $InstanceKey = @($instDict.Keys)[0] }
        else {
            [System.Windows.Forms.MessageBox]::Show("Chưa cấu hình máy chủ MPA nào.", "Lỗi", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return $null
        }
    }
    $instConfig = $instDict[$InstanceKey]
    $baseUrl = $instConfig.MpaBaseUrl
    $credTarget = Get-RequiredProperty -Object $script:CurrentConfig -Name 'CredentialTarget'
    $cred = Get-WindowsGenericCredential -Target $credTarget

    if ($null -eq $cred) {
        [System.Windows.Forms.MessageBox]::Show("Chưa có tài khoản đăng nhập cho '$credTarget'.", "Lỗi", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }

    $plainPassword = Get-PlainTextFromSecureString -SecureString $cred.Password
    $sessionId = $null

    $loadFolder = {
        param([string]$path)
        $lstItems.Items.Clear()
        $txtCurPath.Text = $path
        try {
            $items = Get-ObieeCatalogItems -BaseUrl $baseUrl -Path $path -SessionId $sessionId
            if ($null -ne $items) {
                foreach ($it in $items) {
                    $cap = if (-not [string]::IsNullOrWhiteSpace($it.Caption)) { [string]$it.Caption } else { [string][System.IO.Path]::GetFileName($it.Path) }
                    if ([string]::IsNullOrWhiteSpace($cap)) { $cap = "Mục không tên" }
                    
                    $typeStr = if ($null -ne $it.Type) { [string]$it.Type } else { 'N/A' }
                    $sigStr  = if ($null -ne $it.Signature) { [string]$it.Signature } else { '' }
                    $pathStr = if ($null -ne $it.Path) { [string]$it.Path } else { '' }

                    $lv = New-Object System.Windows.Forms.ListViewItem($cap)
                    [void]$lv.SubItems.Add($typeStr)
                    [void]$lv.SubItems.Add($sigStr)
                    [void]$lv.SubItems.Add($pathStr)
                    $lv.Tag = $it
                    [void]$lstItems.Items.Add($lv)
                }
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Lỗi tải danh mục từ '$path': $($_.Exception.Message)", "Lỗi", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    try {
        $sessionId = Connect-Obiee -BaseUrl $baseUrl -Username $cred.Username -Password $plainPassword
        $currentPath = "/users/$($cred.Username)"
        & $loadFolder $currentPath

        $lstItems.Add_SelectedIndexChanged({
            if ($lstItems.SelectedItems.Count -gt 0) {
                $sel = $lstItems.SelectedItems[0].Tag
                if ($null -ne $sel -and -not [string]::IsNullOrWhiteSpace($sel.Path)) {
                    $txtCurPath.Text = $sel.Path
                }
            }
        })

        $txtCurPath.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $_.SuppressKeyPress = $true
                try { & $loadFolder $txtCurPath.Text.Trim() } catch { }
            }
        })

        $btnGo.Add_Click({
            try {
                $target = $txtCurPath.Text.Trim()
                if (-not [string]::IsNullOrWhiteSpace($target)) {
                    & $loadFolder $target
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Lỗi: $($_.Exception.Message)", "Lỗi", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })

        $btnUp.Add_Click({
            try {
                $p = Split-Path -Parent $txtCurPath.Text.Trim()
                if (-not [string]::IsNullOrWhiteSpace($p)) { & $loadFolder ($p.Replace('\', '/')) }
                else { & $loadFolder '/' }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Lỗi: $($_.Exception.Message)", "Lỗi", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })

        $lstItems.Add_DoubleClick({
            try {
                if ($lstItems.SelectedItems.Count -gt 0) {
                    $selected = $lstItems.SelectedItems[0].Tag
                    if ($selected.Type -match '(?i)folder|dir|portal|link' -or [string]::IsNullOrWhiteSpace($selected.Signature)) {
                        & $loadFolder $selected.Path
                    } else {
                        $script:ChosenItemPath = $selected.Path
                        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
                        $dlg.Close()
                    }
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Lỗi: $($_.Exception.Message)", "Lỗi", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })

        $btnSelect.Add_Click({
            try {
                if ($lstItems.SelectedItems.Count -gt 0) {
                    $selected = $lstItems.SelectedItems[0].Tag
                    $script:ChosenItemPath = $selected.Path
                    $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $dlg.Close()
                } elseif (-not [string]::IsNullOrWhiteSpace($txtCurPath.Text)) {
                    $script:ChosenItemPath = $txtCurPath.Text.Trim()
                    $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
                    $dlg.Close()
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Lỗi: $($_.Exception.Message)", "Lỗi", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })

        $script:ChosenItemPath = $null
        if ($dlg.ShowDialog($mainForm) -eq [System.Windows.Forms.DialogResult]::OK) {
            return $script:ChosenItemPath
        }
    }
    finally {
        if (-not [string]::IsNullOrWhiteSpace($sessionId)) {
            Disconnect-Obiee -BaseUrl $baseUrl -SessionId $sessionId | Out-Null
        }
    }

    return $null
}

# -----------------------------------------------------------------------------
# HỘP THOẠI SOẠN THẢO BÁO CÁO (ADD / EDIT REPORT DIALOG)
# -----------------------------------------------------------------------------

function Show-ReportEditDialog {
    param($ReportToEdit = $null)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = if ($null -eq $ReportToEdit) { "Thêm Báo Cáo MPA Mới" } else { "Chỉnh Sửa Báo Cáo MPA" }
    $dlg.Size = New-Object System.Drawing.Size(680, 560)
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $y = 20

    # 1. Tên Báo cáo
    $lblN = New-Object System.Windows.Forms.Label -Property @{ Text = "Tên Báo Cáo:"; Location = New-Object System.Drawing.Point(24, $y); Size = New-Object System.Drawing.Size(140, 24) }
    $dlg.Controls.Add($lblN)

    $txtN = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(170, $y); Size = New-Object System.Drawing.Size(460, 24) }
    if ($null -ne $ReportToEdit) { $txtN.Text = $ReportToEdit.Name }
    $dlg.Controls.Add($txtN)

    $y += 40

    # 2. Máy chủ
    $lblI = New-Object System.Windows.Forms.Label -Property @{ Text = "Máy Chủ:"; Location = New-Object System.Drawing.Point(24, $y); Size = New-Object System.Drawing.Size(140, 24) }
    $dlg.Controls.Add($lblI)

    $cboI = New-Object System.Windows.Forms.ComboBox -Property @{ Location = New-Object System.Drawing.Point(170, $y); Size = New-Object System.Drawing.Size(240, 24); DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList }
    $instDict = Get-MpaInstances -Config $script:CurrentConfig
    foreach ($k in $instDict.Keys) { [void]$cboI.Items.Add($k) }
    if ($cboI.Items.Count -gt 0) {
        if ($null -ne $ReportToEdit -and $cboI.Items.Contains($ReportToEdit.Instance)) {
            $cboI.SelectedItem = $ReportToEdit.Instance
        } else {
            $cboI.SelectedIndex = 0
        }
    }
    $dlg.Controls.Add($cboI)

    $y += 40

    # 3. Đường dẫn Catalog
    $lblP = New-Object System.Windows.Forms.Label -Property @{ Text = "Đường dẫn Catalog:"; Location = New-Object System.Drawing.Point(24, $y); Size = New-Object System.Drawing.Size(140, 24) }
    $dlg.Controls.Add($lblP)

    $txtP = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(170, $y); Size = New-Object System.Drawing.Size(340, 24) }
    if ($null -ne $ReportToEdit) { $txtP.Text = $ReportToEdit.Path } else { $txtP.Text = "/users/{Username}/_portal/" }
    $dlg.Controls.Add($txtP)

    $btnDuyet = New-Object System.Windows.Forms.Button -Property @{ Text = "Duyệt..."; Location = New-Object System.Drawing.Point(520, ($y - 1)); Size = New-Object System.Drawing.Size(110, 26) }
    $dlg.Controls.Add($btnDuyet)

    $y += 40

    # 4. Định dạng Xuất
    $lblF = New-Object System.Windows.Forms.Label -Property @{ Text = "Định Dạng:"; Location = New-Object System.Drawing.Point(24, $y); Size = New-Object System.Drawing.Size(140, 24) }
    $dlg.Controls.Add($lblF)

    $cboF = New-Object System.Windows.Forms.ComboBox -Property @{ Location = New-Object System.Drawing.Point(170, $y); Size = New-Object System.Drawing.Size(240, 24); DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList }
    @('CSV', 'EXCEL2007', 'PDF', 'MHT') | ForEach-Object { [void]$cboF.Items.Add($_) }
    $selectedFmt = if ($null -ne $ReportToEdit -and $ReportToEdit.Formats -and @($ReportToEdit.Formats).Count -gt 0) {
        $ReportToEdit.Formats[0].Format
    } else { 'CSV' }
    $cboF.SelectedItem = $selectedFmt
    $dlg.Controls.Add($cboF)

    $y += 40

    # 5. Đường dẫn Xuất Tệp
    $lblO = New-Object System.Windows.Forms.Label -Property @{ Text = "Đường Dẫn Lưu Tệp:"; Location = New-Object System.Drawing.Point(24, $y); Size = New-Object System.Drawing.Size(140, 24) }
    $dlg.Controls.Add($lblO)

    $txtO = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(170, $y); Size = New-Object System.Drawing.Size(460, 24) }
    $existingOut = if ($null -ne $ReportToEdit -and $ReportToEdit.Formats -and @($ReportToEdit.Formats).Count -gt 0) {
        $ReportToEdit.Formats[0].OutputPath
    } else { ".\\Reports\\{Yesterday:yyyyMMdd}_{ReportName}.csv" }
    $txtO.Text = $existingOut
    $dlg.Controls.Add($txtO)

    # 6. Tham số / Biến Báo cáo (Parameters / Variables)
    $lblParams = New-Object System.Windows.Forms.Label -Property @{ Text = "Tham số (Key=Value):"; Location = New-Object System.Drawing.Point(24, $y); Size = New-Object System.Drawing.Size(140, 24) }
    $dlg.Controls.Add($lblParams)

    $txtParams = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(170, $y); Size = New-Object System.Drawing.Size(460, 60); Multiline = $true; ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical }
    $initParamsLines = @()
    if ($null -ne $ReportToEdit) {
        $existParams = Get-PropOrKey -Object $ReportToEdit -Name 'Parameters'
        if ($null -ne $existParams) {
            $pairs = Get-ObjectKeyValuePairs -Object $existParams
            foreach ($p in $pairs) {
                $initParamsLines += "$($p.Name)=$($p.Value)"
            }
        }
    }
    $txtParams.Text = $initParamsLines -join "`r`n"
    $dlg.Controls.Add($txtParams)

    $y += 68

    # 7. Live Preview
    $lblPrev = New-Object System.Windows.Forms.Label -Property @{ Text = "Xem trước: "; Location = New-Object System.Drawing.Point(170, $y); Size = New-Object System.Drawing.Size(460, 48); ForeColor = [System.Drawing.Color]::FromArgb(24, 134, 75) }
    $dlg.Controls.Add($lblPrev)

    $updatePreview = {
        $fmt = [string]$cboF.SelectedItem
        $raw = $txtO.Text
        $fakeRep = [pscustomobject]@{ Name = if (-not [string]::IsNullOrWhiteSpace($txtN.Text)) { $txtN.Text } else { "ReportName" }; Path = $txtP.Text; Instance = [string]$cboI.SelectedItem }
        $resolved = Resolve-DynamicTokens -Text $raw -Report $fakeRep -Format $fmt
        $lblPrev.Text = "Xem trước: $resolved"
    }

    $txtO.Add_TextChanged($updatePreview)
    $txtN.Add_TextChanged($updatePreview)
    $cboF.Add_SelectedIndexChanged($updatePreview)
    & $updatePreview

    $y += 56

    # Nút Lưu / Hủy
    $btnOk = New-Object System.Windows.Forms.Button -Property @{ Text = "Xác Nhận"; Location = New-Object System.Drawing.Point(420, $y); Size = New-Object System.Drawing.Size(100, 32); BackColor = [System.Drawing.Color]::FromArgb(26, 115, 232); ForeColor = [System.Drawing.Color]::White; FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; DialogResult = [System.Windows.Forms.DialogResult]::OK }
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button -Property @{ Text = "Hủy Bỏ"; Location = New-Object System.Drawing.Point(530, $y); Size = New-Object System.Drawing.Size(100, 32); DialogResult = [System.Windows.Forms.DialogResult]::Cancel }
    $dlg.Controls.Add($btnCancel)

    $btnDuyet.Add_Click({
        $chosenPath = Show-CatalogBrowserDialog -InstanceKey ([string]$cboI.SelectedItem)
        if (-not [string]::IsNullOrWhiteSpace($chosenPath)) {
            $txtP.Text = $chosenPath
            if ([string]::IsNullOrWhiteSpace($txtN.Text)) {
                $txtN.Text = [System.IO.Path]::GetFileName($chosenPath)
            }
        }
    })

    if ($dlg.ShowDialog($mainForm) -eq [System.Windows.Forms.DialogResult]::OK) {
        if ([string]::IsNullOrWhiteSpace($txtN.Text) -or [string]::IsNullOrWhiteSpace($txtP.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Tên báo cáo và đường dẫn catalog không được để trống.", "Thông báo", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return $null
        }

        $parsedParams = [ordered]@{}
        if (-not [string]::IsNullOrWhiteSpace($txtParams.Text)) {
            $lines = $txtParams.Text.Split("`n")
            foreach ($l in $lines) {
                $trimmed = $l.Trim()
                if ($trimmed -match '^([^=]+)=(.*)$') {
                    $k = $matches[1].Trim()
                    $v = $matches[2].Trim()
                    if (-not [string]::IsNullOrWhiteSpace($k)) {
                        $parsedParams[$k] = $v
                    }
                }
            }
        }

        return [pscustomobject]@{
            Name       = $txtN.Text.Trim()
            Instance   = [string]$cboI.SelectedItem
            Path       = $txtP.Text.Trim()
            Enabled    = if ($null -ne $ReportToEdit) { $ReportToEdit.Enabled } else { $true }
            Parameters = $parsedParams
            Formats    = @(
                [pscustomobject]@{
                    Format     = [string]$cboF.SelectedItem
                    OutputPath = $txtO.Text.Trim()
                }
            )
        }
    }
    return $null
}

# -----------------------------------------------------------------------------
# SỰ KIỆN NÚT BẤM VÀ ĐIỀU KHIỂN (BUTTON & ACTION EVENT HANDLERS)
# -----------------------------------------------------------------------------

$btnAddReport.Add_Click({
    $newRep = Show-ReportEditDialog
    if ($null -ne $newRep) {
        $list = New-Object System.Collections.Generic.List[object]
        $existing = Get-PropOrKey -Object $script:CurrentConfig -Name 'Reports'
        if ($null -ne $existing) { foreach ($r in $existing) { [void]$list.Add($r) } }
        [void]$list.Add($newRep)
        $script:CurrentConfig.Reports = $list.ToArray()
        Refresh-ReportsGrid
        Gui-Log "Đã thêm báo cáo mới: '$($newRep.Name)'. Nhớ bấm 'Lưu Cấu hình'." 'OK'
    }
})

$btnEditReport.Add_Click({
    if ($gridReports.SelectedRows.Count -eq 0) { return }
    $selRep = $gridReports.SelectedRows[0].Tag
    $edited = Show-ReportEditDialog -ReportToEdit $selRep
    if ($null -ne $edited) {
        $selRep.Name = $edited.Name
        $selRep.Instance = $edited.Instance
        $selRep.Path = $edited.Path
        $selRep.Formats = $edited.Formats
        Refresh-ReportsGrid
        Gui-Log "Đã cập nhật thông tin báo cáo: '$($edited.Name)'." 'OK'
    }
})

$btnDeleteReport.Add_Click({
    if ($gridReports.SelectedRows.Count -eq 0) { return }
    $selRep = $gridReports.SelectedRows[0].Tag
    $ans = [System.Windows.Forms.MessageBox]::Show("Bạn có chắc chắn muốn xóa báo cáo '$($selRep.Name)' khỏi cấu hình?", "Xác nhận Xóa", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
        $list = New-Object System.Collections.Generic.List[object]
        foreach ($r in $script:CurrentConfig.Reports) { if ($r -ne $selRep) { [void]$list.Add($r) } }
        $script:CurrentConfig.Reports = $list.ToArray()
        Refresh-ReportsGrid
        Gui-Log "Đã xóa báo cáo '$($selRep.Name)'." 'WARN'
    }
})

$btnBrowseCatalog.Add_Click({
    $instDict = Get-MpaInstances -Config $script:CurrentConfig
    if (@($instDict.Keys).Count -gt 0) {
        $firstKey = @($instDict.Keys)[0]
        $p = Show-CatalogBrowserDialog -InstanceKey $firstKey
        if (-not [string]::IsNullOrWhiteSpace($p)) {
            Gui-Log "Đã chọn đường dẫn từ Catalog: $p" 'OK'
        }
    }
})

$btnSave.Add_Click({
    try {
        $script:CurrentConfig.CredentialTarget = $txtCredTarget.Text.Trim()
        $script:CurrentConfig.DefaultInstance = $txtDefInst.Text.Trim()

        # Update HttpSettings
        Set-ObjectProperty -Object $script:CurrentConfig -Name 'HttpSettings' -Value ([ordered]@{
            TimeoutMinutes = [int]$numHttpTimeout.Value
        })

        # Update Logging
        Set-ObjectProperty -Object $script:CurrentConfig -Name 'Logging' -Value ([ordered]@{
            Enabled             = $chkLogEnabled.Checked
            LogDirectory        = $txtLogDir.Text.Trim()
            LogFileName         = "MpaDownloader_{yyyyMMdd}.log"
            LogLevel            = "INFO"
            ConsoleDebug        = $false
            RetentionDays       = [int]$numRetention.Value
            AuditCsvEnabled     = $chkAuditCsv.Checked
            AuditCsvPath        = ".\Logs\Audit_{yyyyMM}.csv"
            SummaryJsonEnabled  = $true
            SummaryJsonPath     = ".\Logs\LatestRun.json"
        })

        # Update RetryPolicy
        Set-ObjectProperty -Object $script:CurrentConfig -Name 'RetryPolicy' -Value ([ordered]@{
            Enabled             = $chkRetryEn.Checked
            MaxRetries          = [int]$numMaxRetries.Value
            InitialDelaySeconds = [int]$numInitDelay.Value
            BackoffMultiplier   = 2
        })

        Save-MpaConfig -Path $ConfigPath -Config $script:CurrentConfig
        Gui-Log "Đã lưu cấu hình thành công vào $ConfigPath." 'OK'
        [System.Windows.Forms.MessageBox]::Show("Đã lưu cấu hình thành công!", "Thông báo", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        Gui-Log "Lỗi khi lưu cấu hình: $($_.Exception.Message)" 'ERROR'
        [System.Windows.Forms.MessageBox]::Show("Lỗi khi lưu cấu hình: $($_.Exception.Message)", "Lỗi", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnSetCred.Add_Click({
    $target = $txtCredTarget.Text.Trim()
    $existing = Get-WindowsGenericCredential -Target $target
    $userPrompt = if ($null -ne $existing) { "Tài khoản hiện tại: $($existing.Username)`nNhập tài khoản mới:" } else { "Nhập tài khoản MPA ($target):" }
    
    $cred = Get-Credential -Message $userPrompt
    if ($null -ne $cred) {
        Set-WindowsGenericCredential -Target $target -Username $cred.UserName -Password $cred.Password
        Gui-Log "Đã cập nhật thông tin tài khoản Windows Credential Manager cho target '$target'." 'OK'
        [System.Windows.Forms.MessageBox]::Show("Đã lưu tài khoản vào Windows Credential Manager.", "Thành công", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$btnTestAll.Add_Click({
    Gui-Log "Bắt đầu kiểm tra kết nối tới tất cả máy chủ MPA..." 'INFO'
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $target = $txtCredTarget.Text.Trim()
        $cred = Get-WindowsGenericCredential -Target $target
        if ($null -eq $cred) { throw "Chưa thiết lập tài khoản cho '$target'." }
        $plain = Get-PlainTextFromSecureString -SecureString $cred.Password

        $instDict = Get-MpaInstances -Config $script:CurrentConfig
        $msg = ""
        foreach ($k in $instDict.Keys) {
            $inst = $instDict[$k]
            try {
                $sid = Connect-Obiee -BaseUrl $inst.MpaBaseUrl -Username $cred.Username -Password $plain
                Disconnect-Obiee -BaseUrl $inst.MpaBaseUrl -SessionId $sid | Out-Null
                $msg += "[OK] $k ($($inst.MpaBaseUrl)): Đăng nhập & ngắt phiên thành công!`n"
                Gui-Log "Kết nối thành công tới $k ($($inst.MpaBaseUrl))." 'OK'
            }
            catch {
                $msg += "[LỖI] $k ($($inst.MpaBaseUrl)): $($_.Exception.Message)`n"
                Gui-Log "Kết nối thất bại tới $($k): $($_.Exception.Message)" 'ERROR'
            }
        }
        [System.Windows.Forms.MessageBox]::Show($msg, "Kết quả Kiểm tra Kết nối", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        Gui-Log "Lỗi kiểm tra kết nối: $($_.Exception.Message)" 'ERROR'
        [System.Windows.Forms.MessageBox]::Show("Lỗi: $($_.Exception.Message)", "Lỗi", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnRunReport.Add_Click({
    if ($gridReports.SelectedRows.Count -eq 0) { return }
    $rep = $gridReports.SelectedRows[0].Tag
    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Gui-Log "Đang tải thử nghiệm báo cáo '$($rep.Name)'..." 'INFO'

    try {
        $instInfo = Get-MpaReportInstance -Config $script:CurrentConfig -Report $rep
        $baseUrl = $instInfo.Config.MpaBaseUrl
        $target = $txtCredTarget.Text.Trim()
        $cred = Get-WindowsGenericCredential -Target $target
        if ($null -eq $cred) { throw "Chưa thiết lập tài khoản cho '$target'." }
        $plain = Get-PlainTextFromSecureString -SecureString $cred.Password

        $sid = Connect-Obiee -BaseUrl $baseUrl -Username $cred.Username -Password $plain
        try {
            $fmtConfig = $rep.Formats[0]
            $format = $fmtConfig.Format
            $rawOut = $fmtConfig.OutputPath

            $resolvedPath = Resolve-DynamicTokens -Text $rep.Path -Report $rep
            $resolvedPath = $resolvedPath.Replace('{Username}', $cred.Username)

            $resolvedOut = Resolve-DynamicTokens -Text $rawOut -Report $rep -Format $format
            $resolvedOut = $resolvedOut.Replace('{Username}', $cred.Username)
            if (-not [System.IO.Path]::IsPathRooted($resolvedOut)) {
                $resolvedOut = Join-Path $scriptDir $resolvedOut
            }

            $repParams = Get-PropOrKey -Object $rep -Name 'Parameters'
            $repFilters = Get-PropOrKey -Object $rep -Name 'FilterExpressions'
            $exportResult = Export-ObieeAnalysis -BaseUrl $baseUrl -Path $resolvedPath -Format $format -SessionId $sid -Parameters $repParams -FilterExpressions $repFilters
            $size = Save-ObieeExportData -ViewData $exportResult.ViewData -Format $format -OutputPath $resolvedOut

            $sizeKb = [math]::Round($size / 1KB, 1)
            Gui-Log "Tải báo cáo '$($rep.Name)' thành công! File: $resolvedOut ($sizeKb KB)" 'OK'
            [System.Windows.Forms.MessageBox]::Show("Tải báo cáo '$($rep.Name)' thành công!`n`nĐường dẫn: $resolvedOut`nDung lượng: $sizeKb KB", "Thành công", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        finally {
            Disconnect-Obiee -BaseUrl $baseUrl -SessionId $sid | Out-Null
        }
    }
    catch {
        Gui-Log "Lỗi khi tải báo cáo '$($rep.Name)': $($_.Exception.Message)" 'ERROR'
        [System.Windows.Forms.MessageBox]::Show("Lỗi: $($_.Exception.Message)", "Lỗi", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnRunAllReports.Add_Click({
    $ans = [System.Windows.Forms.MessageBox]::Show("Bắt đầu tiến trình tải tất cả các báo cáo đang được kích hoạt?", "Xác nhận", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
        $downloader = Join-Path $scriptDir "MpaReportDownloader.ps1"
        Start-Process powershell.exe -ArgumentList "-NoExit -File `"$downloader`" -ConfigPath `"$ConfigPath`""
    }
})

# Nạp dữ liệu ban đầu
Refresh-ReportsGrid
Refresh-InstancesGrid
Refresh-ScheduleTab
Gui-Log "Ứng dụng MPA Report Downloader đã sẵn sàng." 'INFO'

# Hiển thị cửa sổ chính
[System.Windows.Forms.Application]::Run($mainForm)
