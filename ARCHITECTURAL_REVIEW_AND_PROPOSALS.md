# Repository Review, Findings & Proposed Features

This document provides a comprehensive code scan of **CognosDownloader**, highlighting redundant code patterns, proposed UX and configuration improvements, and an architectural specification for a dedicated **Logging Feature**.

---

## 1. Redundant & Duplicated Code Scan

A comparison between [CognosReportDownloader.ps1](file:///d:/Repo/CognosDownloader/CognosReportDownloader.ps1) and [Manage-CognosConfig.ps1](file:///d:/Repo/CognosDownloader/Manage-CognosConfig.ps1) reveals approximately **~350–400 lines** of duplicate logic.

### Identified Duplications

1. **Win32 Credential Manager P/Invoke & Wrapper Functions (~150 lines)**
   - `CognosReportDownloader.ps1` (lines 45–210) & `Manage-CognosConfig.ps1` (lines 101–200)
   - *Duplicate logic*: Identical C# `CognosCredManager.NativeMethods` P/Invoke struct, `advapi32.dll` function imports (`CredRead`, `CredWrite`, `CredFree`), and helper functions (`Set-WindowsGenericCredential`/`Set-StoredCredential`, `Get-WindowsGenericCredential`/`Get-StoredCredential`, `Get-PlainTextFromSecureString`/`Get-PlainText`).

2. **Dynamic Token Engine (`Resolve-DynamicTokens`) (~75 lines)**
   - `CognosReportDownloader.ps1` (lines 216–291) & `Manage-CognosConfig.ps1` (lines 23–95)
   - *Duplicate logic*: Date math, token replacement for `{Yesterday}`, `{Today}`, `{MonthStart}`, `{MonthEnd}`, `{Today±Nd}`, custom date format strings, report placeholders, and parameter injections.
   - *Micro-redundancy inside function*: Lines replacing `{Yesterday}` and `{Today}` literally are immediately re-processed by the regular expression `\{(Yesterday|Today)(:([^}]+))?\}`, which already covers the non-formatted case.

3. **HTTP Client Initialization & Redirect Loop Handler (~60 lines)**
   - `CognosReportDownloader.ps1` (lines 376–450) & `Manage-CognosConfig.ps1` (lines 205–256)
   - *Duplicate logic*: `New-CognosHttpClient`/`New-HttpClient` (setting up `CookieContainer` and `HttpClientHandler`) and `Invoke-CognosRequest` (handling manual HTTP 300–308 redirect traversal).

4. **CAM XML Authentication (`Invoke-CognosLogin`) (~50 lines)**
   - `CognosReportDownloader.ps1` (lines 457–531) & `Manage-CognosConfig.ps1` (lines 258–285)
   - *Duplicate logic*: CAM XML credential payload construction, form URL encoding, dispatching POST to `/v1/disp/rds/auth/logon`, and extracting the `XSRF-TOKEN` cookie.

### Refactoring Recommendation
If eliminating duplication is preferred over keeping each script 100% self-contained:
* Extract shared logic into a common script module (e.g., `CognosCommon.ps1` or a private `.psm1` module) that is dot-sourced by both `CognosReportDownloader.ps1` and `Manage-CognosConfig.ps1`:
  ```powershell
  . (Join-Path $PSScriptRoot 'CognosCommon.ps1')
  ```

---

## 2. Proposed UX & CLI Enhancements

### Interactive Manager (`Manage-CognosConfig.ps1`)
1. **Direct Test Run Action**:
   - Add menu action `[7] Run / Test Report Download Now` to allow on-the-spot test downloading of an active report directly from the manager menu without needing to open a separate terminal.
2. **Directory & Path Validation**:
   - When users enter or preview an `OutputPath`, check if the target folder exists and prompt to create it if missing.
3. **Batch Operations**:
   - Enable/Disable all reports in a single action, or toggle multiple reports without re-opening the menu repeatedly.

### Headless Downloader (`CognosReportDownloader.ps1`)
1. **Targeted CLI Filtering Flags**:
   - Add command-line arguments to allow downloading specific reports on demand:
     ```powershell
     .\CognosReportDownloader.ps1 -ReportName "Báo cáo nợ đến hạn, quá hạn hàng ngày"
     .\CognosReportDownloader.ps1 -Source "i54414D93B29A4D2289C4E88469871644" -Format "xlsxData"
     ```
2. **Tabular Execution Summary**:
   - Output an execution report table at the end of the batch run:
     ```
     Report Name       Format    Status   Size (MB)   Duration   Output Path
     ----------------- --------  -------  ---------   --------   -------------------------------------
     Báo cáo nợ...     xlsxData  SUCCESS  2.41 MB     3.2s       D:\CognosReports\20260820_Report.xlsx
     ```

---

## 3. Proposed Configuration Schema Enhancements (`cognos-reports.json`)

1. **Global Default Output Directory**:
   - Allow a top-level `"DefaultOutputDir": "D:\\CognosReports"` so individual report paths can be specified as relative templates (`"{Yesterday:yyyyMMdd}_{ReportName}.xlsx"`).
2. **Environment Variable Tokens**:
   - Expand `Resolve-DynamicTokens` to support environment variable substitutions like `{env:USERPROFILE}`, `{env:APPDATA}`, or `{env:TEMP}`.
3. **Per-Report Retries & Timeouts**:
   - Add optional `"MaxRetries": 3` and `"TimeoutMinutes": 30` under each report definition to handle transient Cognos network glitches or long-running database queries.

---

## 4. Proposed Logging Feature Architecture

A robust logging framework is critical for headless scripts scheduled via Windows Task Scheduler or CI/CD pipelines.

### Key Objectives
* Provide real-time, colored console feedback for interactive sessions.
* Maintain persistent, structured log files for post-execution auditing.
* Support machine-readable metric records for downstream log aggregators (e.g., Splunk, Elastic, Datadog).
* Automatically rotate and clean up old logs to prevent disk exhaustion.

---

### Specification & Design

#### 1. Configuration in `cognos-reports.json`
Add an optional `"Logging"` configuration block:

```json
{
  "CognosBaseUrl": "http://10.53.153.173/ibmcognos/bi",
  "Namespace": "BIDV",
  "CredentialTarget": "CognosReportAutomation:cognos.example.com",
  "Logging": {
    "Enabled": true,
    "LogDirectory": ".\\Logs",
    "LogFileName": "CognosDownloader_{yyyyMMdd}.log",
    "LogLevel": "INFO",
    "RetentionDays": 30,
    "AuditCsvEnabled": true,
    "AuditCsvPath": ".\\Logs\\Audit_{yyyyMM}.csv"
  },
  "Reports": [ ... ]
}
```

#### 2. Log Levels & Formatting
* **Supported Levels**: `DEBUG` (10), `INFO` (20), `OK` (20), `WARN` (30), `ERROR` (40).
* **Text Log Entry Format**:
  ```
  [2026-08-21 11:35:00] [INFO ] Connecting to Cognos at http://10.53.153.173/ibmcognos/bi...
  [2026-08-21 11:35:01] [OK   ] Authenticated as automation_user in namespace 'BIDV'.
  [2026-08-21 11:35:04] [INFO ] Downloading 'Báo cáo nợ đến hạn, quá hạn hàng ngày' (xlsxData)...
  [2026-08-21 11:35:07] [OK   ] Saved D:\CognosReports\20260820_Báo cáo nợ đến hạn, quá hạn hàng ngày.xlsx (2.41 MB) in 3.12s.
  [2026-08-21 11:35:07] [OK   ] Batch completed. Total: 1, Successful: 1, Failed: 0.
  ```

#### 3. Audit CSV / Metrics Log (Machine-Readable)
For tracking download history, durations, file sizes, and errors in a format easily queried by Excel or reporting dashboards:

| Field | Description | Example |
| :--- | :--- | :--- |
| `Timestamp` | ISO 8601 execution timestamp | `2026-08-21T11:35:07+07:00` |
| `ReportName` | Report Name | `Báo cáo nợ đến hạn, quá hạn hàng ngày` |
| `Source` | Cognos StoreID / Path | `i54414D93B29A4D2289C4E88469871644` |
| `Format` | Output Format | `xlsxData` |
| `Status` | `SUCCESS` or `FAILED` | `SUCCESS` |
| `HttpStatusCode` | HTTP response code from Cognos | `200` |
| `FileSizeBytes` | Output file size | `2527064` |
| `DurationMs` | Elapsed time in milliseconds | `3120` |
| `OutputPath` | Full resolved output destination | `D:\CognosReports\20260820_Report.xlsx` |
| `ErrorMessage` | Exception text if failed | *(empty or error string)* |

#### 4. Automatic Log Retention & Rotation
* At startup, the logging engine scans the `LogDirectory` for `.log` and `.csv` files exceeding `RetentionDays` (e.g., 30 days) and deletes them automatically.

#### 5. Windows Event Log Integration (Optional Flag)
* An optional switch (`-LogToEventLog`) to write batch summary events (`EventID 1000 = Success`, `EventID 1001 = Failure`) to the Windows Application log for automated infrastructure monitoring.
