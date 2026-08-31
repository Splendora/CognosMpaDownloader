# BỘ CÔNG CỤ TỰ ĐỘNG HÓA BÁO CÁO COGNOS & MPA (OBIEE)

Hệ thống tự động hóa toàn diện quy trình kết nối, phân tích tham số và tải dữ liệu báo cáo định kỳ từ hai nền tảng phân tích doanh nghiệp lớn: **IBM Cognos Analytics 11** (qua REST Mashup API) và **Oracle Business Intelligence / MPA** (qua SOAP Web Services) sử dụng **PowerShell 5.1+**, bảo mật tài khoản bằng **Windows Credential Manager** và hỗ trợ đầy đủ giao diện **Windows Forms GUI** cùng **CLI**.

---

## 📂 BẢN ĐỒ TÀI LIỆU CHI TIẾT

| Phân Hệ | Tài Liệu Hướng Dẫn Chi Tiết | Nền Tảng Kỹ Thuật | Giao Diện Hỗ Trợ |
| :--- | :--- | :--- | :--- |
| **IBM Cognos** | 📘 [**COGNOS/COGNOS.md**](COGNOS/COGNOS.md) | Cognos RDS (Mashup REST API), CAM Auth, XML POST Payload | GUI (`CognosConfigGui.ps1`), CLI (`Manage-CognosConfig.ps1`), Batch (`CognosReportDownloader.ps1`), Pipeline (`Run-DailyPipeline.ps1`) |
| **MPA / OBIEE** | 📕 [**MPA/MPA.md**](MPA/MPA.md) | Oracle BI Web Services (SOAP v10), Web Catalog, Export View | GUI (`MpaConfigGui.ps1`), CLI (`Manage-MpaConfig.ps1`), Batch (`MpaReportDownloader.ps1`), Diagnostic (`Test-ObieeSoap.ps1`) |

---

## 🌟 TÍNH NĂNG NỔI BẬT

1. **Bảo Mật An Toàn Tuyệt Đối:**
   - Sử dụng Windows Credential Manager native (`advapi32.dll` P/Invoke) để lưu trữ tài khoản (Target `MaCB`).
   - Tuyệt đối không lưu mật khẩu dạng plaintext trong mã nguồn hay các tệp JSON cấu hình.
2. **Khám Phá & Lựa Chọn Tham Số Tự Động (Smart Prompt Discovery):**
   - Tự động truy vấn server để lấy danh sách prompt, trạng thái bắt buộc/tùy chọn, giá trị lựa chọn (`useValue`, `displayValue`) và giá trị mặc định.
   - Hộp thoại tìm kiếm và tích chọn trực quan trên giao diện GUI.
   - Hỗ trợ nạp hàng nghìn mã CIF / chi nhánh từ tệp văn bản ngoài (`@path` hoặc `{file:...}`).
3. **Vượt Qua Giới Hạn URL 64KB (Cognos XML POST Payload):**
   - Đóng gói toàn bộ tham số vào cấu trúc chuẩn `<promptAnswers>` XML và truyền qua HTTP POST `ByteArrayContent`, giải quyết triệt để lỗi `HTTP 414 Request-URI Too Large` hoặc `Invalid URI`.
4. **Duyệt Cây Thư Mục Catalog Trực Quan (MPA / OBIEE):**
   - Giao diện cây thư mục giúp duyệt trực tiếp các báo cáo trên `/shared` và `/users` của Oracle BI.
5. **Hệ Thống Phân Giải Token Động (Dynamic Token Engine):**
   - Hỗ trợ thay thế linh hoạt các biến ngày tháng và metadata: `{Yesterday}`, `{Today}`, `{MonthStart}`, `{MonthEnd}`, `{Today-7d}`, `{Yesterday:yyyyMMdd}`, `{ReportName}`, `{Username}`, `{p_...}` trong cả tham số và đường dẫn xuất tệp.
6. **Đa Định Dạng Đầu Ra:**
   - Xuất dữ liệu sang `xlsxData` (Excel Data), `spreadsheetML` (Excel XML), `EXCEL2007` (`.xlsx`), `CSV`, `PDF`, `MHT`.
7. **Độ Tin Cậy & Tự Động Thử Lại (Retry Policy):**
   - Cơ chế Exponential Backoff (tự động thử lại 3 lần kèm giãn cách thời gian) và tự động cấp mới session khi phiên bị hết hạn.
8. **Nhật Ký & Báo Cáo Thực Thi:**
   - Ghi log theo cấp độ (`INFO`, `DEBUG`, `WARN`, `ERROR`), tệp rolling audit CSV hàng tháng, xuất `LatestRun.json` và tự động dọn dẹp log cũ theo số ngày lưu trữ (`RetentionDays`).
9. **Tích Hợp Sẵn Lịch Chạy Nền:**
   - Hỗ trợ đăng ký trực tiếp vào Windows Task Scheduler từ giao diện GUI hoặc dòng lệnh.

---

## 🏗️ CẤU TRÚC THƯ MỤC DỰ ÁN

```
CognosMpaDownloader/
├── README.md                    # Tài liệu tổng quan toàn bộ dự án (Tiếng Việt)
├── AGENTS.md                    # Quy chuẩn và hướng dẫn kỹ thuật dành cho AI Agents / Developers
│
├── COGNOS/                      # Phân hệ tải báo cáo IBM Cognos Analytics 11
│   ├── COGNOS.md                # Tài liệu hướng dẫn chi tiết phân hệ Cognos
│   ├── CognosCommon.ps1         # Module lõi dùng chung (Auth, HTTP, Token, XML Prompt, Log)
│   ├── CognosConfigGui.ps1      # Giao diện đồ họa Windows Forms quản trị Cognos
│   ├── Manage-CognosConfig.ps1  # Giao diện dòng lệnh tương tác CLI Cognos
│   ├── CognosReportDownloader.ps1 # Script chạy tự động hàng loạt không giao diện
│   ├── Run-DailyPipeline.ps1    # Pipeline tự động Cognos kết hợp Macro VBA Excel
│   ├── cognos-reports.json      # Tệp cấu hình các báo cáo và máy chủ Cognos
│   └── Logs/                    # Nhật ký và audit của phân hệ Cognos
│
└── MPA/                         # Phân hệ tải báo cáo Oracle BI / MPA (OBIEE)
    ├── MPA.md                   # Tài liệu hướng dẫn chi tiết phân hệ MPA / OBIEE
    ├── ObieeCommon.ps1          # Module lõi SOAP, Catalog, Token, Export & Log
    ├── MpaConfigGui.ps1         # Giao diện đồ họa Windows Forms quản trị MPA
    ├── Manage-MpaConfig.ps1     # Giao diện dòng lệnh tương tác CLI MPA
    ├── MpaReportDownloader.ps1  # Script chạy tự động hàng loạt không giao diện
    ├── Test-ObieeSoap.ps1       # Công cụ kiểm tra & chẩn đoán dịch vụ SOAP
    ├── mpa-reports.json         # Tệp cấu hình các báo cáo và máy chủ MPA
    ├── Reports/                 # Thư mục chứa báo cáo tải về (mặc định)
    └── Logs/                    # Nhật ký và audit của phân hệ MPA
```

---

## ⚡ BẮT ĐẦU NHANH (QUICK START)

### 1. Yêu Cầu Môi Trường
- Hệ điều hành: **Windows 10 / 11 / Windows Server**
- **Windows PowerShell 5.1** hoặc **PowerShell 7+**

Kích hoạt quyền thực thi kịch bản (mở PowerShell):
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

### 2. Thiết Lập Tài Khoản Xác Thực (Chung cho cả 2 hệ thống)
Hệ thống sử dụng chung Target Name `MaCB` trong Windows Credential Manager.

```powershell
# Thiết lập cho Cognos
powershell -ExecutionPolicy RemoteSigned -File .\COGNOS\CognosReportDownloader.ps1 -SetupCredential

# Hoặc thiết lập cho MPA
powershell -ExecutionPolicy RemoteSigned -File .\MPA\MpaReportDownloader.ps1 -SetupCredential
```

---

### 3. Khởi Chạy Giao Diện Quản Trị Đồ Họa (GUI)

```powershell
# Mở giao diện Cognos GUI
powershell -ExecutionPolicy RemoteSigned -File .\COGNOS\CognosConfigGui.ps1

# Mở giao diện MPA / OBIEE GUI
powershell -ExecutionPolicy RemoteSigned -File .\MPA\MpaConfigGui.ps1
```

---

### 4. Chạy Tải Báo Cáo Tự Động Hàng Loạt (Batch Headless)

```powershell
# Tải toàn bộ báo cáo Cognos
powershell -NoProfile -ExecutionPolicy RemoteSigned -File .\COGNOS\CognosReportDownloader.ps1

# Tải toàn bộ báo cáo MPA / OBIEE
powershell -NoProfile -ExecutionPolicy RemoteSigned -File .\MPA\MpaReportDownloader.ps1
```

---

## 📅 THIẾT LẬP LỊCH TỰ ĐỘNG (WINDOWS TASK SCHEDULER)

Bạn có thể thiết lập lịch chạy tự động hàng ngày thông qua tab **"Lịch Tự Động"** trên giao diện GUI của từng phân hệ, hoặc đăng ký bằng dòng lệnh PowerShell với quyền Administrator:

```powershell
# Lịch tải Cognos lúc 10:00 hàng ngày
$actionCognos = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy RemoteSigned -File `"D:\Repo\CognosMpaDownloader\COGNOS\CognosReportDownloader.ps1`""
$triggerCognos = New-ScheduledTaskTrigger -Daily -At "10:00AM"
Register-ScheduledTask -TaskName "CognosReportDownloader" -Action $actionCognos -Trigger $triggerCognos -Description "Tải báo cáo Cognos hàng ngày" -User "SYSTEM"

# Lịch tải MPA lúc 08:00 hàng ngày
$actionMpa = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy RemoteSigned -File `"D:\Repo\CognosMpaDownloader\MPA\MpaReportDownloader.ps1`""
$triggerMpa = New-ScheduledTaskTrigger -Daily -At "08:00AM"
Register-ScheduledTask -TaskName "MpaReportDownloader" -Action $actionMpa -Trigger $triggerMpa -Description "Tải báo cáo MPA OBIEE hàng ngày" -User "SYSTEM"
```

---

## 📖 XEM CHI TIẾT TỪNG PHÂN HỆ
- [📘 Hướng dẫn đầy đủ tính năng và cấu hình COGNOS](COGNOS/COGNOS.md)
- [📕 Hướng dẫn đầy đủ tính năng và cấu hình MPA / OBIEE](MPA/MPA.md)
