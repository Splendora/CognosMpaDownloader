# TÀI LIỆU HƯỚNG DẪN HỆ THỐNG MPA / OBIEE REPORT DOWNLOADER

Bộ công cụ tự động hóa kết nối, duyệt cây danh mục (Catalog) và tải báo cáo phân tích từ hệ thống **Oracle Business Intelligence Enterprise Edition (OBIEE / MPA)** qua chuẩn **SOAP Web Services** sử dụng **PowerShell 5.1+**, tích hợp quản lý tài khoản an toàn qua **Windows Credential Manager** và giao diện quản trị đồ họa **Windows Forms GUI** / dòng lệnh **CLI**.

---

## MỤC LỤC

1. [Tổng Quan & Kiến Trúc](#1-tổng-quan--kiến-trúc)
2. [Cấu Trúc Thư Mục & Thành Phần](#2-cấu-trúc-thư-mục--thành-phần)
3. [Yêu Cầu Môi Trường & Cài Đặt Ban Đầu](#3-yêu-cầu-môi-trường--cài-đặt-ban-đầu)
4. [Quản Lý Thông Tin Xác Thực (Credentials)](#4-quản-lý-thông-tin-xác-thực-credentials)
5. [Dịch Vụ SOAP Web Services của OBIEE](#5-dịch-vụ-soap-web-services-của-obiee)
6. [Hướng Dẫn Sử Dụng Chi Tiết](#6-hướng-dẫn-sử-dụng-chi-tiết)
   - [6.1. Giao Diện Đồ Họa (MpaConfigGui.ps1)](#61-giao-diện-đồ-họa-mpaconfigguips1)
   - [6.2. Giao Diện Dòng Lệnh Tương Tác (Manage-MpaConfig.ps1)](#62-giao-diện-dòng-lệnh-tương-tác-manage-mpaconfigps1)
   - [6.3. Chạy Tự Động Hàng Loạt (MpaReportDownloader.ps1)](#63-chạy-tự-động-hàng-loạt-mpareportdownloaderps1)
   - [6.4. Công Cụ Kiểm Tra & Chẩn Đoán SOAP (Test-ObieeSoap.ps1)](#64-công-cụ-kiểm-tra--chẩn-đoán-soap-test-obieesoapps1)
7. [Cấu Trúc Tệp Cấu Hình (mpa-reports.json)](#7-cấu-trúc-tệp-cấu-hình-mpa-reportsjson)
8. [Hệ Thống Phân Giải Token Động (Dynamic Tokens)](#8-hệ-thống-phân-giải-token-động-dynamic-tokens)
9. [Định Dạng Xuất Dữ Liệu Hỗ Trợ](#9-định-dạng-xuất-dữ-liệu-hỗ-trợ)
10. [Chính Sách Thử Lại & Quản Lý Phiên (Session Management)](#10-chính-sách-thử-lại--quản-lý-phiên-session-management)
11. [Hệ Thống Ghi Log, Audit & Tự Động Dọn Dẹp](#11-hệ-thống-ghi-log-audit--tự-động-dọn-dẹp)
12. [Thiết Lập Lịch Tự Động (Windows Task Scheduler)](#12-thiết-lập-lịch-tự-động-windows-task-scheduler)
13. [Khắc Phục Sự Cố Thường Gặp (Troubleshooting)](#13-khắc-phục-sự-cố-thường-gặp-troubleshooting)

---

## 1. Tổng Quan & Kiến Trúc

Hệ thống MPA Downloader được xây dựng chuyên biệt để giao tiếp trực tiếp với hạ tầng Oracle BI (OBIEE / MPA) thông qua giao thức SOAP XML chuẩn công nghiệp, cho phép tự động hóa hoàn toàn quy trình truy xuất báo cáo phân tích mà không cần mở trình duyệt web.

```
                    ┌────────────────────────────┐
                    │ Windows Credential Manager │
                    │   (advapi32.dll P/Invoke)  │
                    └─────────────┬──────────────┘
                                  │ Target: 'MaCB'
                                  ▼
┌────────────────────────────────────────────────────────────────────────┐
│                           ObieeCommon.ps1                              │
│  - SOAP Envelopes & Serialization (urn://oracle.bi.webservices/v10)    │
│  - Session Management (logon, logoff, keep-alive via nQSessionService) │
│  - Web Catalog Traversal (getItemInfo, getSubItems)                    │
│  - Analysis Export View Service (CSV, EXCEL2007, PDF, MHT)             │
│  - Dynamic Date & Metadata Token Engine                                │
│  - Retry Policy, File Logging, CSV Audit & LatestRun.json Summary      │
└──────┬──────────────────────────┬───────────────────────────────┬──────┘
       │                          │                               │
       ▼                          ▼                               ▼
┌──────────────┐          ┌──────────────┐          ┌───────────────────────┐
│ GUI Manager  │          │ CLI Manager  │          │ Batch Downloader /    │
│ (Catalog Tree│          │(Interactive  │          │ SOAP Diagnostic Tool  │
│  & Downloader│          │ Menu App)    │          │ (Headless Execution)  │
└──────────────┘          └──────────────┘          └───────────┬───────────┘
                                                                │
                                                                ▼
                                                    ┌───────────────────────┐
                                                    │   Oracle BI (MPA)     │
                                                    │ /analytics-ws/saw.dll │
                                                    └───────────────────────┘
```

### Các tính năng cốt lõi:
- **Tương tác SOAP nguyên bản:** Giao tiếp trực tiếp với các Web Service của OBIEE (`nQSessionService`, `webCatalogService`, `xmlViewService`, `analysisExportViewsService`) với chuẩn namespace `urn://oracle.bi.webservices/v10`.
- **Duyệt Catalog trực quan:** Cho phép duyệt cây thư mục `/shared`, `/users` của OBIEE trực tiếp từ giao diện đồ họa để chọn báo cáo cần xuất.
- **Bảo mật tuyệt đối:** Sử dụng Win32 API tương tác trực tiếp với Windows Credential Manager, hỗ trợ mật khẩu chứa ký tự đặc biệt phức tạp.
- **Hỗ trợ định dạng phong phú:** Xuất báo cáo sang `CSV`, `EXCEL2007` (bảng tính Excel), `PDF`, `MHT`.
- **Đóng phiên an toàn (Safe Logoff):** Đảm bảo hàm `logoff` luôn được thực thi trong khối `finally`, giải phóng session trên máy chủ ứng dụng Oracle WebLogic.
- **Tự động thử lại thông minh (Smart Retry):** Tự động cấp mới session và thử lại khi gặp lỗi gián đoạn đường truyền hoặc server bận.

---

## 2. Cấu Trúc Thư Mục & Thành Phần

```
MPA/
├── ObieeCommon.ps1             # Module lõi SOAP, Auth, Catalog, Export, Token & Logging
├── MpaReportDownloader.ps1     # Script tải báo cáo hàng loạt chạy ngầm (Scheduled Task)
├── Manage-MpaConfig.ps1        # Giao diện dòng lệnh (CLI) duyệt catalog & cấu hình
├── MpaConfigGui.ps1            # Giao diện đồ họa (Windows Forms GUI) quản trị & tải dữ liệu
├── Test-ObieeSoap.ps1          # Công cụ kiểm tra chẩn đoán kết nối & dịch vụ SOAP
├── mpa-reports.json            # Tệp cấu hình máy chủ MPA, danh sách báo cáo & lịch chạy
├── Reports/                    # Thư mục chứa các tệp báo cáo tải về (mặc định)
└── Logs/                       # Thư mục lưu trữ nhật ký thực thi và tệp audit
```

### Vai trò từng tệp tin:
- **`ObieeCommon.ps1`**: Thư viện chứa toàn bộ hàm giao tiếp SOAP XML, tạo envelope, xử lý mã hóa UTF-8, gọi API `advapi32.dll`, phân giải token ngày tháng, quản lý nhật ký và dọn dẹp log cũ.
- **`MpaReportDownloader.ps1`**: Thực thi tải tự động các báo cáo kích hoạt trong `mpa-reports.json`, hỗ trợ đa định dạng, ghi nhận tiến độ và xuất file audit.
- **`MpaConfigGui.ps1`**: Ứng dụng desktop Windows Forms với cây thư mục Catalog OBIEE, cho phép click chọn báo cáo, xem trước đường dẫn xuất động, chạy thử nghiệm tải báo cáo và cấu hình Windows Task Scheduler.
- **`Manage-MpaConfig.ps1`**: Console tương tác giúp quản trị viên cấu hình máy chủ, duyệt cây báo cáo trên terminal và thực hiện kiểm thử nhanh.
- **`Test-ObieeSoap.ps1`**: Kịch bản kiểm thử độc lập giúp chẩn đoán chi tiết từng bước: Đăng nhập -> Kiểm tra thông tin báo cáo (`getItemInfo`) -> Liệt kê thư mục con (`getSubItems`) -> Xuất thử dữ liệu -> Đăng xuất (`logoff`).

---

## 3. Yêu Cầu Môi Trường & Cài Đặt Ban Đầu

### Yêu cầu hệ thống:
- Hệ điều hành: **Windows 10 / Windows 11 / Windows Server 2016+**
- Môi trường thực thi: **Windows PowerShell 5.1** hoặc **PowerShell 7+**
- Kết nối mạng: Truy cập được tới máy chủ Oracle BI MPA qua cổng HTTP/SOAP (ví dụ: `http://10.53.44.180:9704/analytics-ws/saw.dll`).

### Thiết lập quyền thực thi PowerShell:
Mở PowerShell và cấp quyền thực thi kịch bản:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

---

## 4. Quản Lý Thông Tin Xác Thực (Credentials)

Thông tin tài khoản được lưu an toàn trong Windows Credential Manager dưới mục Generic Credential (mặc định Target Name: `MaCB`).

### Cách 1: Thiết lập qua Giao diện Đồ họa GUI
1. Khởi chạy `MpaConfigGui.ps1`.
2. Nhấn nút **"Cài Đặt Tài Khoản"** (hoặc phím tắt `Ctrl+K`).
3. Nhập Tên đăng nhập và Mật khẩu, nhấn **"Lưu Tài Khoản"**.

### Cách 2: Thiết lập qua Dòng Lệnh CLI
```powershell
.\MpaReportDownloader.ps1 -SetupCredential
```

### Kiểm tra kết nối xác thực:
```powershell
.\MpaReportDownloader.ps1 -TestConnection
```

---

## 5. Dịch Vụ SOAP Web Services của OBIEE

Hệ thống giao tiếp với Oracle BI thông qua các endpoint sau:

| Dịch Vụ (Service) | Endpoint URL | Chức Năng Chính |
| :--- | :--- | :--- |
| **`nQSessionService`** | `?SoapImpl=nQSessionService` | Xác thực đăng nhập (`logon`), duy trì kết nối (`keepAlive`), kết thúc phiên (`logoff`) |
| **`webCatalogService`** | `?SoapImpl=webCatalogService` | Lấy thông tin đối tượng (`getItemInfo`), duyệt thư mục con (`getSubItems`), đọc metadata |
| **`xmlViewService`** | `?SoapImpl=xmlViewService` | Thực thi truy vấn phân tích và lấy kết quả định dạng XML/HTML |
| **`analysisExportViewsService`** | `?SoapImpl=analysisExportViewsService` | Xuất dữ liệu báo cáo sang định dạng nhị phân/văn bản (`CSV`, `EXCEL2007`, `PDF`, `MHT`) |

*Ghi chú: Toàn bộ thông điệp SOAP sử dụng chuẩn XML Namespace:*
```xml
xmlns:sawsoap="urn://oracle.bi.webservices/v10"
```

---

## 6. Hướng Dẫn Sử Dụng Chi Tiết

### 6.1. Giao Diện Đồ Họa (`MpaConfigGui.ps1`)

Khởi chạy ứng dụng GUI:
```powershell
powershell -ExecutionPolicy RemoteSigned -File .\MpaConfigGui.ps1
```

#### Các tính năng chính:
1. **Duyệt Cây Thư Mục Catalog Trực Quan (Catalog Browser):**
   - Tự động kết nối máy chủ và hiển thị cây thư mục phân cấp `/shared` và `/users`.
   - Nhấp chọn báo cáo để tự động điền đường dẫn Catalog vào biểu mẫu cấu hình.
2. **Quản Lý Danh Sách Báo Cáo:**
   - Thêm mới, chỉnh sửa, bật/tắt kích hoạt (`Enabled`), đổi định dạng xuất (`CSV`, `EXCEL2007`, `PDF`).
3. **Xem Trước Đường Dẫn Xuất (Live Path Preview):**
   - Tự động giải mã token động theo thời gian thực (ví dụ: `Reports\{Yesterday:yyyyMMdd}_{ReportName}.csv` -> `Reports\20260830_ROA.csv`).
4. **Tải Thử Nghiệm & Tải Hàng Loạt:**
   - Nút **"Tải Thử Báo Cáo"**: Thực hiện tải ngay 1 báo cáo để kiểm tra dữ liệu.
   - Nút **"Tải Tất Cả Báo Cáo (Batch Run)"**: Tải toàn bộ danh sách báo cáo đang bật với bảng tiến độ chi tiết.
5. **Cấu Hình Lịch Chạy Tự Động (Scheduling):**
   - Tab "Lịch Tự Động" hỗ trợ đăng ký task chạy nền hàng ngày vào Windows Task Scheduler.

---

### 6.2. Giao Diện Dòng Lệnh Tương Tác (`Manage-MpaConfig.ps1`)

Khởi động console CLI:
```powershell
.\Manage-MpaConfig.ps1
```

```text
======================================================================
         QUẢN LÝ CẤU HÌNH TẢI BÁO CÁO MPA / OBIEE (CLI)
======================================================================
  [1] Cài đặt / Cập nhật tài khoản trong Windows Credential Manager
  [2] Kiểm tra kết nối & đăng nhập thử nghiệm
  [3] Xem danh sách báo cáo đã cấu hình
  [4] Thêm báo cáo mới (Duyệt cây thư mục Catalog)
  [5] Chỉnh sửa cấu hình báo cáo
  [6] Bật / Tắt trạng thái tải của báo cáo
  [7] Tải thử nghiệm một báo cáo
  [8] TẢI TOÀN BỘ BÁO CÁO (Chạy Batch)
  [9] Cấu hình Máy chủ & Cài đặt chung
  [10] Mở Giao diện Đồ họa (GUI)
  [0] Thoát
======================================================================
```

---

### 6.3. Chạy Tự Động Hàng Loạt (`MpaReportDownloader.ps1`)

Sử dụng cho các tác vụ định kỳ tự động chạy không cần người dùng thao tác:

```powershell
# Chạy với tệp cấu hình mặc định (.\mpa-reports.json)
powershell -NoProfile -ExecutionPolicy RemoteSigned -File .\MpaReportDownloader.ps1

# Chạy với tệp cấu hình khác
powershell -NoProfile -ExecutionPolicy RemoteSigned -File .\MpaReportDownloader.ps1 -ConfigPath "D:\Config\mpa-custom.json"
```

---

### 6.4. Công Cụ Kiểm Tra & Chẩn Đoán SOAP (`Test-ObieeSoap.ps1`)

Dùng để kiểm tra nhanh khả năng kết nối và độ phản hồi của từng dịch vụ SOAP trên máy chủ MPA:

```powershell
powershell -ExecutionPolicy RemoteSigned -File .\Test-ObieeSoap.ps1
```

Script sẽ thực hiện tuần tự:
1. Đăng nhập SOAP và lấy `sessionID`.
2. Kiểm tra thông tin phân tích `ROA` (`/users/user/_portal/ROA` hoặc đường dẫn tùy chọn).
3. Liệt kê các đối tượng trong thư mục catalog (`getSubItems`).
4. Xuất thử dữ liệu báo cáo sang file tạm.
5. Đăng xuất an toàn và in báo cáo chẩn đoán chi tiết.

---

## 7. Cấu Trúc Tệp Cấu Hình (`mpa-reports.json`)

```json
{
  "CredentialTarget": "MaCB",
  "DefaultInstance": "MPA",
  "HttpSettings": {
    "TimeoutMinutes": 5
  },
  "RetryPolicy": {
    "Enabled": true,
    "MaxRetries": 3,
    "InitialDelaySeconds": 5,
    "BackoffMultiplier": 2
  },
  "Scheduling": {
    "TaskName": "MpaReportDownloader",
    "ScheduleType": "DAILY",
    "StartTime": "08:00",
    "RunElevated": false
  },
  "Instances": {
    "MPA": {
      "MpaBaseUrl": "http://10.53.44.180:9704/analytics-ws/saw.dll"
    }
  },
  "Logging": {
    "Enabled": true,
    "LogDirectory": ".\\Logs",
    "LogFileName": "MpaDownloader_{yyyyMMdd}.log",
    "LogLevel": "INFO",
    "ConsoleDebug": false,
    "RetentionDays": 18,
    "AuditCsvEnabled": true,
    "AuditCsvPath": ".\\Logs\\Audit_{yyyyMM}.csv",
    "SummaryJsonEnabled": true,
    "SummaryJsonPath": ".\\Logs\\LatestRun.json"
  },
  "Reports": [
    {
      "Name": "Báo cáo ROA Chi Nhánh",
      "Instance": "MPA",
      "CatalogPath": "/users/{Username}/_portal/ROA",
      "Enabled": true,
      "Formats": [
        {
          "Format": "CSV",
          "OutputPath": "D:\\MpaReports\\{Yesterday:yyyyMMdd}_{ReportName}.csv"
        }
      ]
    }
  ]
}
```

---

## 8. Hệ Thống Phân Giải Token Động (Dynamic Tokens)

Hệ thống hỗ trợ thay thế linh hoạt các biến ngày tháng và metadata trong cả đường dẫn Catalog (`CatalogPath`) lẫn đường dẫn xuất tệp (`OutputPath`):

| Token | Ví Dụ Giá Trị Đầu Ra | Ý Nghĩa / Mục Đích |
| :--- | :--- | :--- |
| `{Yesterday}` | `2026-08-30` | Ngày hôm qua (định dạng ISO `yyyy-MM-dd`) |
| `{Today}` | `2026-08-31` | Ngày hôm nay (định dạng ISO `yyyy-MM-dd`) |
| `{MonthStart}` | `2026-08-01` | Ngày đầu tiên của tháng hiện tại |
| `{MonthEnd}` | `2026-08-31` | Ngày cuối cùng của tháng hiện tại |
| `{Today-7d}` / `{Today+1d}` | `2026-08-24` | Ngày dịch chuyển N ngày so với hiện tại |
| `{Yesterday:yyyyMMdd}` | `20260830` | Ngày hôm qua kèm định dạng tùy chỉnh |
| `{Today:dd/MM/yyyy}` | `31/08/2026` | Ngày hôm nay kèm định dạng tùy chỉnh |
| `{yyyyMMdd}`, `{yyyy-MM-dd}` | `20260831` | Dấu thời gian năm tháng ngày |
| `{HHmmss}` | `083000` | Dấu thời gian giờ phút giây |
| `{ReportName}` | `Bao_cao_ROA` | Tên báo cáo (đã chuẩn hóa an toàn ký tự) |
| `{Username}` | `user01` | Tên đăng nhập từ Windows Credential Manager |
| `{Instance}` | `MPA` | Tên cấu hình máy chủ đang kết nối |
| `{Format}` | `CSV` | Định dạng tệp xuất |

---

## 9. Định Dạng Xuất Dữ Liệu Hỗ Trợ

Thông qua dịch vụ `analysisExportViewsService`, MPA Downloader hỗ trợ xuất các định dạng:

1. **`CSV`**: Dữ liệu dạng văn bản phân tách dấu phẩy (Comma-Separated Values), tối ưu cho xử lý ETL, nạp cơ sở dữ liệu và phân tích tự động.
2. **`EXCEL2007`**: Tệp bảng tính Microsoft Excel (`.xlsx`), giữ nguyên cấu trúc bảng và kiểu dữ liệu.
3. **`PDF`**: Tài liệu Portable Document Format, phù hợp lưu trữ và gửi báo cáo trình ký.
4. **`MHT`**: Lưu trữ dạng Web Archive hoàn chỉnh (bao gồm cả định dạng giao diện HTML).

---

## 10. Chính Sách Thử Lại & Quản Lý Phiên (Session Management)

- **Quản lý vòng đời phiên (Session Lifecycle):**
  - Đăng nhập lấy `sessionID` 48 ký tự qua `nQSessionService.logon`.
  - Thực hiện toàn bộ tác vụ duyệt catalog và xuất báo cáo với `sessionID` này.
  - Luôn luôn gọi `nQSessionService.logoff` trong khối `finally` để giải phóng tài nguyên trên máy chủ Oracle WebLogic.
- **Cơ chế Thử Lại Tự Động (Exponential Backoff):**
  - Khi gặp sự cố mạng tạm thời hoặc phiên làm việc bị timeout, hàm `Invoke-MpaReportDownloadWithRetry` sẽ tự động tạm dừng (5s, 10s, 20s), kết nối lại để xin cấp `sessionID` mới và thực hiện lại tác vụ tải báo cáo.
  - Tối đa thử lại 3 lần trước khi đánh dấu thất bại.

---

## 11. Hệ Thống Ghi Log, Audit & Tự Động Dọn Dẹp

- **Ghi log chi tiết (`Write-Log`):** Ghi nhận đầy đủ thông tin các bước SOAP envelope, HTTP status code và nội dung phản hồi.
- **Audit CSV hàng tháng (`Audit_{yyyyMM}.csv`):** Lưu trữ lịch sử tải chi tiết phục vụ kiểm tra và đối soát vận hành:
  - *Cột dữ liệu:* `Timestamp`, `ReportName`, `CatalogPath`, `Format`, `Status`, `HttpStatusCode`, `FileSizeBytes`, `DurationMs`, `OutputPath`, `ErrorMessage`.
- **Tổng kết báo cáo (`LatestRun.json`):** Cung cấp dữ liệu thống kê tổng thể số lượt thành công/thất bại, tổng dung lượng và thời gian thực thi cho các hệ thống giám sát.
- **Tự động dọn dẹp log:** Tự động xóa các file nhật ký vượt quá số ngày cấu hình (`RetentionDays`, mặc định 18-30 ngày).

---

## 12. Thiết Lập Lịch Tự Động (Windows Task Scheduler)

Bạn có thể thiết lập lịch chạy tự động hàng ngày qua giao diện GUI hoặc sử dụng lệnh PowerShell sau:

```powershell
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy RemoteSigned -File `"$PSScriptRoot\MpaReportDownloader.ps1`""
$trigger = New-ScheduledTaskTrigger -Daily -At "08:00AM"
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
Register-ScheduledTask -TaskName "MpaReportDownloader" -Action $action -Trigger $trigger -Settings $settings -Description "Tự động tải báo cáo MPA OBIEE hàng ngày" -User "SYSTEM"
```

---

## 13. Khắc Phục Sự Cố Thường Gặp (Troubleshooting)

### 1. Lỗi đăng nhập SOAP "Invalid User Name or Password"
- **Nguyên nhân:** Tên đăng nhập hoặc mật khẩu trong Windows Credential Manager (Target `MaCB`) không chính xác.
- **Cách xử lý:** Chạy `.\MpaReportDownloader.ps1 -SetupCredential` để cập nhật lại thông tin tài khoản.

### 2. Lỗi "Catalog object not found: /users/..."
- **Nguyên nhân:** Đường dẫn báo cáo trên Catalog bị sai hoặc tài khoản không có quyền truy cập vào thư mục của người dùng khác.
- **Cách xử lý:** Mở `MpaConfigGui.ps1`, sử dụng cây thư mục Catalog để duyệt và click trực tiếp vào báo cáo cần xuất.

### 3. Lỗi "SOAP Fault: Error while executing export"
- **Nguyên nhân:** Báo cáo trên OBIEE chứa prompt bắt buộc chưa được gán giá trị mặc định, hoặc truy vấn database phía backend bị lỗi.
- **Cách xử lý:** Chạy `.\Test-ObieeSoap.ps1` để kiểm tra thông điệp SOAP Fault chi tiết từ server.
