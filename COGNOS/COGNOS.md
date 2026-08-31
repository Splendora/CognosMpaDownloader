# TÀI LIỆU HƯỚNG DẪN HỆ THỐNG COGNOS REPORT DOWNLOADER

Bộ công cụ tự động hóa trích xuất và tải báo cáo từ **IBM Cognos Analytics 11** qua giao diện **REST API (Report Data Service - RDS / Mashup)** sử dụng **PowerShell 5.1+**, tích hợp quản lý thông tin xác thực an toàn qua **Windows Credential Manager** và giao diện quản trị đồ họa **Windows Forms GUI** / dòng lệnh **CLI**.

---

## MỤC LỤC

1. [Tổng Quan & Kiến Trúc](#1-tổng-quan--kiến-trúc)
2. [Cấu Trúc Thư Mục & Thành Phần](#2-cấu-trúc-thư-mục--thành-phần)
3. [Yêu Cầu Môi Trường & Cài Đặt Ban Đầu](#3-yêu-cầu-môi-trường--cài-đặt-ban-đầu)
4. [Quản Lý Thông Tin Xác Thực (Credentials)](#4-quản-lý-thông-tin-xác-thực-credentials)
5. [Hướng Dẫn Sử Dụng Chi Tiết](#5-hướng-dẫn-sử-dụng-chi-tiết)
   - [5.1. Giao Diện Đồ Họa (CognosConfigGui.ps1)](#51-giao-diện-đồ-họa-cognosconfigguips1)
   - [5.2. Giao Diện Dòng Lệnh Tương Tác (Manage-CognosConfig.ps1)](#52-giao-diện-dòng-lệnh-tương-tác-manage-cognosconfigps1)
   - [5.3. Chạy Tự Động Hàng Loạt (CognosReportDownloader.ps1)](#53-chạy-tự-động-hàng-loạt-cognosreportdownloaderps1)
   - [5.4. Pipeline Xử Lý Hàng Ngày & Excel Macro (Run-DailyPipeline.ps1)](#54-pipeline-xử-lý-hàng-ngày--excel-macro-run-dailypipelineps1)
6. [Cấu Trúc Tệp Cấu Hình (cognos-reports.json)](#6-cấu-trúc-tệp-cấu-hình-cognos-reportsjson)
7. [Hệ Thống Phân Giải Token Động (Dynamic Tokens)](#7-hệ-thống-phân-giải-token-động-dynamic-tokens)
8. [Cơ Chế Gửi Tham Số XML Payload (Bypass Giới Hạn 64KB)](#8-cơ-chế-gửi-tham-số-xml-payload-bypass-giới-hạn-64kb)
9. [Chính Sách Thử Lại & Độ Tin Cậy (Retry Policy)](#9-chính-sách-thử-lại--độ-tin-cậy-retry-policy)
10. [Hệ Thống Ghi Log, Audit & Tự Động Dọn Dẹp](#10-hệ-thống-ghi-log-audit--tự-động-dọn-dẹp)
11. [Thiết Lập Lịch Tự Động (Windows Task Scheduler)](#11-thiết-lập-lịch-tự-động-windows-task-scheduler)
12. [Khắc Phục Sự Cố Thường Gặp (Troubleshooting)](#12-khắc-phục-sự-cố-thường-gặp-troubleshooting)

---

## 1. Tổng Quan & Kiến Trúc

Hệ thống được thiết kế để thay thế hoàn toàn các thao tác thủ công khi đăng nhập, chọn tham số và tải dữ liệu từ IBM Cognos Analytics 11.

```
                    ┌────────────────────────────┐
                    │ Windows Credential Manager │
                    │   (advapi32.dll P/Invoke)  │
                    └─────────────┬──────────────┘
                                  │ Target: 'MaCB'
                                  ▼
┌────────────────────────────────────────────────────────────────────────┐
│                          CognosCommon.ps1                              │
│  - CAM XML Authentication (/v1/disp/rds/auth/logon)                    │
│  - Parameter Discovery & Choice Lists (/v1/disp/rds/sessionParameters) │
│  - Dynamic Date & Metadata Token Engine                                │
│  - XML Prompt Serialization & POST ByteArray Payload Delivery          │
│  - Logging, Retention, CSV Audit & Summary Reporting                   │
└──────┬──────────────────────────┬───────────────────────────────┬──────┘
       │                          │                               │
       ▼                          ▼                               ▼
┌──────────────┐          ┌──────────────┐          ┌───────────────────────┐
│ GUI Manager  │          │ CLI Manager  │          │ Batch Downloader /    │
│  (Forms UI)  │          │ (Interactive)│          │ Daily Pipeline Runner │
└──────────────┘          └──────────────┘          └───────────┬───────────┘
                                                                │
                                                                ▼
                                                    ┌───────────────────────┐
                                                    │ IBM Cognos Analytics  │
                                                    │ (ODS / COGNOS11 / ...)│
                                                    └───────────────────────┘
```

### Các tính năng cốt lõi:
- **Tương thích đa instance:** Hỗ trợ kết nối song song nhiều máy chủ Cognos (ví dụ: `ODS`, `COGNOS11`, `BIDV_Core`) trong cùng một cấu hình.
- **Bảo mật chuẩn doanh nghiệp:** Mật khẩu được mã hóa và lưu trữ trong Windows Credential Manager, tuyệt đối không lưu plaintext trong mã nguồn hay JSON.
- **Khám phá tham số tự động (Prompt Discovery):** Tự động truy vấn server để lấy danh sách prompt, loại prompt, trạng thái bắt buộc/tùy chọn, danh sách giá trị lựa chọn (`useValue`, `displayValue`) và giá trị mặc định.
- **Hỗ trợ tải danh sách lớn từ tệp:** Tự động đọc hàng nghìn mã chi nhánh, CIF từ tệp ngoài (`@path` hoặc `{file:...}`).
- **Xử lý URL quá tải:** Tự động chuyển đổi toàn bộ tham số thành định dạng chuẩn `<promptAnswers>` XML và gửi qua HTTP POST body `ByteArrayContent` nhằm vượt qua giới hạn 64KB của .NET URI.
- **Đa định dạng đầu ra:** Xuất dữ liệu sang `xlsxData` (Excel Data thuần), `spreadsheetML` (Excel XML định dạng), `PDF`, `CSV`.

---

## 2. Cấu Trúc Thư Mục & Thành Phần

```
COGNOS/
├── CognosCommon.ps1             # Thư viện dùng chung (Auth, HTTP, Token, Logging, XML Prompt)
├── CognosReportDownloader.ps1   # Script tải tự động hàng loạt chạy ngầm (Scheduled Task/CI/CD)
├── Manage-CognosConfig.ps1      # Giao diện dòng lệnh (CLI) quản lý cấu hình & test tải
├── CognosConfigGui.ps1          # Giao diện đồ họa (Windows Forms GUI) quản trị toàn diện
├── Run-DailyPipeline.ps1        # Pipeline tự động tích hợp tải Cognos và chạy VBA Macro Excel
├── cognos-reports.json          # Tệp cấu hình máy chủ, tham số báo cáo & đường dẫn xuất
└── Logs/                        # Thư mục lưu nhật ký (Log, Audit CSV, LatestRun.json)
```

### Vai trò từng tệp tin:
- **`CognosCommon.ps1`**: Module lõi chứa toàn bộ logic xử lý Windows API, CAM Login, HTTP Client, xử lý cookie phiên, phân giải token, xây dựng payload XML, ghi log và tổng kết.
- **`CognosReportDownloader.ps1`**: Chạy không giao diện, đọc `cognos-reports.json`, tự động lặp qua các báo cáo đang bật (`Enabled: true`) và tải về thư mục chỉ định.
- **`CognosConfigGui.ps1`**: Giao diện cửa sổ Windows Forms trực quan, cho phép duyệt danh sách báo cáo, xem cây tham số, chọn giá trị từ danh sách máy chủ trả về, kiểm tra kết nối và tải thử nghiệm.
- **`Manage-CognosConfig.ps1`**: Trình đơn console (CLI) với đầy đủ chức năng quản lý, phân tích tham số từ server và chạy tải trực tiếp.
- **`Run-DailyPipeline.ps1`**: Kịch bản chạy hàng ngày quy trình phối hợp: Tải báo cáo tổng quan -> chạy macro trích xuất CIF -> tải báo cáo theo CIF -> chạy macro xử lý số liệu cuối cùng.

---

## 3. Yêu Cầu Môi Trường & Cài Đặt Ban Đầu

### Yêu cầu hệ thống:
- Hệ điều hành: **Windows 10 / Windows 11 / Windows Server 2016+**
- Môi trường thực thi: **Windows PowerShell 5.1** hoặc **PowerShell 7+**
- Kết nối mạng nội bộ: Truy cập được tới các địa chỉ IP của máy chủ Cognos (ví dụ: `10.53.153.173`, `10.53.6.164`).

### Thiết lập quyền thực thi PowerShell:
Trước khi chạy script lần đầu, mở PowerShell (Run as Administrator nếu cần) và thực hiện lệnh:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

---

## 4. Quản Lý Thông Tin Xác Thực (Credentials)

Hệ thống sử dụng Win32 API (`advapi32.dll`) để lưu thông tin đăng nhập vào **Windows Credential Manager** dạng Generic Credential, đảm bảo an toàn tuyệt đối.

### Cách 1: Thiết lập qua Giao diện Đồ họa GUI
1. Chạy file `CognosConfigGui.ps1`.
2. Trên thanh công cụ, nhấn nút **"Cài Đặt Tài Khoản"** (hoặc `Ctrl+K`).
3. Nhập Tên đăng nhập và Mật khẩu, nhấn **"Lưu Tài Khoản"**.

### Cách 2: Thiết lập qua Dòng Lệnh CLI
Chạy lệnh sau trong thư mục `COGNOS`:

```powershell
.\CognosReportDownloader.ps1 -SetupCredential
```
Hoặc qua trình đơn `Manage-CognosConfig.ps1` -> chọn mục `[1] Cài đặt / Cập nhật tài khoản trong Windows Credential Manager`.

### Kiểm tra kết nối xác thực:
```powershell
.\CognosReportDownloader.ps1 -TestConnection
```

---

## 5. Hướng Dẫn Sử Dụng Chi Tiết

### 5.1. Giao Diện Đồ Họa (`CognosConfigGui.ps1`)

Khởi động giao diện GUI:
```powershell
powershell -ExecutionPolicy RemoteSigned -File .\CognosConfigGui.ps1
```

#### Các chức năng nổi bật trên GUI:
1. **Quản lý danh sách Báo Cáo:**
   - Xem toàn bộ báo cáo, trạng thái Kích hoạt (`Enabled`), Instance nguồn (`ODS`, `COGNOS11`), Mã Source ID và Định dạng xuất.
   - Thêm mới, chỉnh sửa, nhân bản hoặc xóa báo cáo.
2. **Khám Phá & Chọn Tham Số Thông Minh:**
   - Nhấn nút **"Lấy Tham Số Từ Server"**: Ứng dụng tự động kết nối máy chủ Cognos, phân tích định nghĩa báo cáo và liệt kê toàn bộ Prompt.
   - Hiển thị nhãn **`BẮT BUỘC (*)`** (màu đỏ) hoặc **`Tùy chọn`** (màu xám).
   - Với các prompt dạng danh sách (`selectValue`), nhấn nút **"Chọn Giá Trị..."** để mở hộp thoại tìm kiếm và tích chọn nhiều mục trực quan.
3. **Xem Trước Đường Dẫn Xuất (Live Path Preview):**
   - Tự động thay thế các token động (`{Yesterday}`, `{ReportName}`) theo thời gian thực để người dùng kiểm tra chính xác tệp tin sẽ được lưu ở đâu.
4. **Tải Thử Nghiệm & Tải Hàng Loạt:**
   - Nút **"Tải Thử Nghiệm Báo Cáo Này"**: Kiểm tra việc tải 1 báo cáo đơn lẻ.
   - Nút **"Tải Tất Cả Báo Cáo (Batch Run)"**: Tải toàn bộ báo cáo có `Enabled: true` với cơ chế tái sử dụng phiên đăng nhập (Session Caching).
5. **Cấu Hình Task Scheduler:**
   - Tab **"Lịch Tự Động"**: Thiết lập giờ chạy định kỳ hàng ngày chỉ với 1 click.

---

### 5.2. Giao Diện Dòng Lệnh Tương Tác (`Manage-CognosConfig.ps1`)

Khởi động console CLI:
```powershell
.\Manage-CognosConfig.ps1
```

```text
======================================================================
         QUẢN LÝ CẤU HÌNH TẢI BÁO CÁO COGNOS (CLI)
======================================================================
  [1] Cài đặt / Cập nhật tài khoản trong Windows Credential Manager
  [2] Kiểm tra kết nối tới các máy chủ Cognos
  [3] Xem danh sách báo cáo đã cấu hình
  [4] Thêm báo cáo mới (Khám phá tham số tự động từ Server)
  [5] Chỉnh sửa tham số / định dạng báo cáo
  [6] Bật / Tắt trạng thái tải của báo cáo
  [7] Tải thử nghiệm một báo cáo
  [8] TẢI TOÀN BỘ BÁO CÁO (Chạy Batch)
  [9] Cấu hình Máy chủ & Cài đặt chung
  [10] Mở Giao diện Đồ họa (GUI)
  [0] Thoát
======================================================================
```

---

### 5.3. Chạy Tự Động Hàng Loạt (`CognosReportDownloader.ps1`)

Dùng để cấu hình trong Windows Task Scheduler, kịch bản CI/CD hoặc chạy nền:

```powershell
# Chạy với tệp cấu hình mặc định (.\cognos-reports.json)
powershell -NoProfile -ExecutionPolicy RemoteSigned -File .\CognosReportDownloader.ps1

# Chạy với tệp cấu hình tùy chọn
powershell -NoProfile -ExecutionPolicy RemoteSigned -File .\CognosReportDownloader.ps1 -ConfigPath "D:\Config\my-reports.json"
```

---

### 5.4. Pipeline Xử Lý Hàng Ngày & Excel Macro (`Run-DailyPipeline.ps1`)

Kịch bản chuyên biệt phối hợp tải dữ liệu Cognos đa bước và tự động kích hoạt macro VBA Excel:

```powershell
powershell -NoProfile -ExecutionPolicy RemoteSigned -File .\Run-DailyPipeline.ps1
```

**Quy trình thực hiện:**
1. Tải báo cáo `TONG QUAN TTKH` từ Cognos ODS.
2. Mở file Excel Macro `Xu ly du lieu_update.xlsm`, thực thi hàm VBA `handle_file_tong_quan`.
3. Kiểm tra tệp trung gian `Input_cif.txt` vừa được macro sinh ra.
4. Sử dụng danh sách CIF trong `Input_cif.txt` để tải 2 báo cáo chi tiết: `TONG HOA LOI ICH ODS` và `SPDV`.
5. Gọi lần lượt các hàm VBA `handle_file_tong_hoa_loi_ich` và `handle_file_spdv` để tổng hợp số liệu báo cáo cuối ngày.

---

## 6. Cấu Trúc Tệp Cấu Hình (`cognos-reports.json`)

```json
{
  "CredentialTarget": "MaCB",
  "DefaultInstance": "ODS",
  "HttpSettings": {
    "TimeoutMinutes": 10
  },
  "RetryPolicy": {
    "Enabled": true,
    "MaxRetries": 3,
    "InitialDelaySeconds": 5,
    "BackoffMultiplier": 2
  },
  "Scheduling": {
    "TaskName": "CognosReportDownloader",
    "ScheduleType": "DAILY",
    "StartTime": "10:00",
    "RunElevated": false
  },
  "Instances": {
    "ODS": {
      "CognosBaseUrl": "http://10.53.153.173/ibmcognos/bi",
      "Namespace": "BIDV"
    },
    "COGNOS11": {
      "CognosBaseUrl": "http://10.53.6.164/ibmcognos/bi",
      "Namespace": "BIDV"
    }
  },
  "Logging": {
    "Enabled": true,
    "LogDirectory": ".\\Logs",
    "LogFileName": "CognosDownloader_{yyyyMMdd}.log",
    "LogLevel": "INFO",
    "ConsoleDebug": false,
    "RetentionDays": 30,
    "AuditCsvEnabled": true,
    "AuditCsvPath": ".\\Logs\\Audit_{yyyyMM}.csv",
    "SummaryJsonEnabled": true,
    "SummaryJsonPath": ".\\Logs\\LatestRun.json"
  },
  "Reports": [
    {
      "Name": "Báo cáo nợ đến hạn, quá hạn hàng ngày",
      "Instance": "ODS",
      "Source": "i54414D93B29A4D2289C4E88469871644",
      "SourceType": "report",
      "Enabled": true,
      "Parameters": {
        "p_P_ngay": "{Yesterday}",
        "p_p_cn_xu_ly": "121000, 121150"
      },
      "Formats": [
        {
          "Format": "xlsxData",
          "OutputPath": "\\\\10.26.136.32\\Dulieu\\NoQuaHan\\{Yesterday:yyyyMMdd}_{ReportName}.xlsx"
        }
      ]
    }
  ]
}
```

---

## 7. Hệ Thống Phân Giải Token Động (Dynamic Tokens)

Hệ thống hỗ trợ cơ chế thay thế chuỗi linh hoạt cho cả giá trị tham số (`Parameters`) và đường dẫn lưu tệp (`OutputPath`):

| Token | Ví Dụ Giá Trị Đầu Ra | Ý Nghĩa / Mục Đích |
| :--- | :--- | :--- |
| `{Yesterday}` | `2026-08-30` | Ngày hôm qua (định dạng ISO `yyyy-MM-dd`) |
| `{Today}` | `2026-08-31` | Ngày hiện tại (định dạng ISO `yyyy-MM-dd`) |
| `{MonthStart}` | `2026-08-01` | Ngày đầu tiên của tháng hiện tại |
| `{MonthEnd}` | `2026-08-31` | Ngày cuối cùng của tháng hiện tại |
| `{Today-7d}` / `{Today+1d}` | `2026-08-24` | Ngày lùi/tiến N ngày so với ngày hiện tại |
| `{Yesterday:yyyyMMdd}` | `20260830` | Ngày hôm qua kèm định dạng tùy biến |
| `{Today:dd/MM/yyyy}` | `31/08/2026` | Ngày hôm nay kèm định dạng tùy biến |
| `{yyyyMMdd}`, `{yyyy-MM-dd}` | `20260831` | Dấu thời gian năm tháng ngày |
| `{HHmmss}` | `143000` | Dấu thời gian giờ phút giây |
| `{ReportName}` | `Bao_cao_no_den_han` | Tên báo cáo (đã loại bỏ ký tự đặc biệt) |
| `{Source}` | `i54414D93B29A4...` | Mã định danh StoreID của báo cáo |
| `{Instance}` | `ODS` | Tên máy chủ Cognos đang kết nối |
| `{Format}` | `xlsxData` | Định dạng xuất tệp |
| `{p_ParameterName}` | *giá trị tham số* | Lấy giá trị của một tham số khác gắn vào đường dẫn |
| `@D:\path\list.txt` | *mảng chuỗi* | Nạp danh sách nhiều giá trị từ tệp ngoài |
| `{file-single:D:\path\list.txt}` | `"121000", "121150"` | Nạp toàn bộ danh sách thành 1 chuỗi gom cách nhau dấu phẩy |

---

## 8. Cơ Chế Gửi Tham Số XML Payload (Bypass Giới Hạn 64KB)

### Vấn đề kỹ thuật:
Khi tải báo cáo có danh sách lọc lớn (ví dụ: lọc hàng trăm chi nhánh hoặc hàng nghìn CIF khách hàng), nếu truyền qua URL Query Parameters (`?p_CIF=...`), .NET WebClient / HttpClient sẽ báo lỗi:
- `HTTP 414 Request-URI Too Large`
- `System.UriFormatException: Invalid URI: The Uri string is too long` (giới hạn 65.520 ký tự).

### Giải pháp kỹ thuật trong `CognosCommon.ps1`:
1. **Tuần tự hóa cấu trúc XML Prompt:** Toàn bộ tham số được chuyển đổi thành tài liệu XML tiêu chuẩn của Cognos:
   ```xml
   <promptAnswers>
     <promptValues>
       <name>p_pCIF</name>
       <values><item><simplePVal><useValue>100001</useValue></simplePVal></item></values>
     </promptValues>
   </promptAnswers>
   ```
2. **Gửi qua HTTP POST ByteArrayContent:**
   - Dữ liệu XML được mã hóa thành UTF-8 bytes và gửi trong phần thân HTTP POST dạng `xmlData=<urlencoded_xml>`.
   - **Content-Type bắt buộc:** `application/x-www-form-urlencoded` (**Tuyệt đối không gắn thêm `; charset=utf-8`** do máy chủ WebSphere/Tomcat của Cognos sẽ bỏ qua request body nếu có header charset).
   - Cơ chế này cho phép gửi danh sách tham số dung lượng không giới hạn một cách mượt mà và ổn định 100%.

---

## 9. Chính Sách Thử Lại & Độ Tin Cậy (Retry Policy)

Hàm `Invoke-CognosReportDownloadWithRetry` đảm bảo tính sẵn sàng cao khi gặp sự cố mạng chập chờn hoặc máy chủ Cognos bị nghẽn:
- **Exponential Backoff:** Khi gặp lỗi mạng, HTTP 500, HTTP 503 hoặc phản hồi lỗi từ server, script sẽ dừng và đợi theo cấp số nhân (mặc định: lần 1 đợi 5 giây, lần 2 đợi 10 giây, lần 3 đợi 20 giây).
- **Tự động gia hạn phiên (Session Renewal):** Nếu phiên làm việc (`XSRF-TOKEN`) hết hạn giữa chừng, script sẽ tự động đăng nhập lại và lấy token mới trước khi thử tải lại báo cáo.
- **Kiểm tra tính toàn vẹn nhị phân:** Trước khi ghi tệp ra đĩa, script tự động quét nội dung byte đầu để phát hiện các thông báo lỗi dạng XML/HTML (`RDS-ERR`, `<soapenv:Fault>`, `<rds:error>`), tránh việc lưu nhầm tệp lỗi thành file Excel/PDF giả mạo.

---

## 10. Hệ Thống Ghi Log, Audit & Tự Động Dọn Dẹp

- **Ghi log nhiều cấp độ (`Write-Log`):** Hỗ trợ các mức `DEBUG`, `INFO`, `OK`, `WARN`, `ERROR`. Mặc định log `DEBUG` chỉ ghi vào tệp để giữ màn hình console luôn gọn gàng.
- **Tệp Audit hàng tháng (`Audit_{yyyyMM}.csv`):** Tự động ghi nhận thông số chi tiết của từng lượt tải: Thời gian, Tên báo cáo, Source ID, Định dạng, Trạng thái (SUCCESS/FAILED), Mã HTTP, Dung lượng (Bytes), Thời gian thực thi (ms), Đường dẫn lưu, Chi tiết lỗi.
- **Bảng tổng kết cuối phiên (`Write-ExecutionSummaryReport`):** Xuất bảng tổng hợp trực quan ra console và lưu tệp `LatestRun.json` phục vụ các hệ thống giám sát.
- **Tự động dọn dẹp log (`Invoke-LogRetentionCleanup`):** Tự động quét và xóa các tệp log/audit đã quá hạn lưu trữ (`RetentionDays`, mặc định 30 ngày) để giải phóng dung lượng ổ đĩa.

---

## 11. Thiết Lập Lịch Tự Động (Windows Task Scheduler)

Bạn có thể tạo lịch chạy tự động hàng ngày thông qua giao diện GUI (Tab "Lịch Tự Động") hoặc chạy lệnh PowerShell sau với quyền Administrator:

```powershell
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy RemoteSigned -File `"$PSScriptRoot\CognosReportDownloader.ps1`""
$trigger = New-ScheduledTaskTrigger -Daily -At "08:00AM"
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
Register-ScheduledTask -TaskName "CognosReportDownloader" -Action $action -Trigger $trigger -Settings $settings -Description "Tự động tải báo cáo Cognos hàng ngày" -User "SYSTEM"
```

---

## 12. Khắc Phục Sự Cố Thường Gặp (Troubleshooting)

### 1. Lỗi xác thực "CAM Authentication Failed" / "Invalid Credentials"
- **Nguyên nhân:** Tên đăng nhập hoặc mật khẩu trong Windows Credential Manager bị sai, hoặc tài khoản bị khóa trên Active Directory.
- **Cách xử lý:** Chạy `.\CognosReportDownloader.ps1 -SetupCredential` để nhập lại mật khẩu mới.

### 2. Tải về file 0 KB hoặc file Excel thông báo lỗi khi mở
- **Nguyên nhân:** Máy chủ trả về thông báo lỗi XML thay vì file dữ liệu.
- **Cách xử lý:** Mở tệp log trong thư mục `Logs/` để kiểm tra thông điệp lỗi chi tiết từ Cognos RDS (thường do tham số ngày tháng không hợp lệ hoặc không có quyền truy cập báo cáo).

### 3. Lỗi kết nối mạng "Unable to connect to remote server"
- **Nguyên nhân:** Chưa kết nối VPN / mạng nội bộ, hoặc máy chủ Cognos đang bảo trì.
- **Cách xử lý:** Chạy `.\CognosReportDownloader.ps1 -TestConnection` để kiểm tra kết nối tới từng Instance.
