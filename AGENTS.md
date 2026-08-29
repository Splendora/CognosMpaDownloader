# AGENTS.md

## Repository Overview

**CognosDownloader** is a PowerShell-based automation tool and configuration manager for IBM Cognos Analytics 11. It automates the extraction and downloading of Cognos reports across multiple formats (`xlsxData`, `spreadsheetML`, `PDF`, `CSV`) via the Cognos RDS (Report Data Service / Mashup) REST interface, resolving dynamic date tokens and securely managing credentials using the native Windows Credential Manager.
Note: User should enable PowerShell before running.
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned


---

## File Structure & Component Map

```
CognosDownloader/
├── CognosCommon.ps1             # Shared core module (Credentials, HTTP client, dynamic tokens, CAM login, prompt inspector)
├── CognosReportDownloader.ps1   # Headless batch execution script for automated report downloads
├── Manage-CognosConfig.ps1      # Interactive CLI console manager for configuration, prompt discovery & test runner
├── CognosConfigGui.ps1          # Windows Forms GUI configuration manager, choice picker & batch runner
├── cognos-reports.json          # Configuration file defining server endpoints and report definitions
└── AGENTS.md                    # Repository documentation and instructions for AI agents
```

---

## Core Components & Architecture

### 1. `CognosCommon.ps1`
Shared module dot-sourced by `CognosReportDownloader.ps1`, `Manage-CognosConfig.ps1`, and `CognosConfigGui.ps1`.

* **Responsibilities**:
  * Win32 Windows Credential Manager P/Invoke (`advapi32.dll` `CredRead`, `CredWrite`, `CredFree`).
  * Dynamic token resolution engine (`Resolve-DynamicTokens`) for dates and metadata in paths and parameters.
  * HTTP Client setup, redirect loop handling (`Invoke-CognosRequest`), and cookie management.
  * CAM XML credentials payload generation and login authentication (`Invoke-CognosLogin`).
  * Structured Cognos prompt inspection (`Get-CognosReportParameters`):
    * Extracts prompt types (`selectValue`, `selectDate`, `textBox`).
    * Extracts multi-select & cardinality flags.
    * Extracts mandatory (`IsRequired`) vs. optional prompt indicators.
    * Extracts full choice lists (`useValue` and `displayValue`).
    * Extracts server-configured default values.
  * Multi-value query parameter serialization (`Add-QueryParameter`): serializes JSON array prompts as repeated query parameters (`&param=val1&param=val2`).
  * Multi-instance endpoint resolution (`Get-CognosInstances`, `Get-CognosReportInstance`).
  * UTF-8 JSON configuration file reading, writing, and backup management (`Load-CognosConfig`, `Save-CognosConfig`).
  * Unified log and console output formatting (`Write-Log`, `Write-AuditLog`).

### 2. `CognosReportDownloader.ps1`
Headless execution script designed for scheduled tasks (e.g., Windows Task Scheduler) or CI/CD pipelines.

* **Parameters**:
  * `-ConfigPath <string>`: Path to JSON configuration file (default: `.\cognos-reports.json`).
  * `-SetupCredential`: Interactive prompt to store credentials in Windows Credential Manager and exit.
  * `-TestConnection`: Verifies network connectivity and authentication against Cognos without running reports.
* **Execution Flow**:
  1. Loads configuration from `cognos-reports.json`.
  2. Retrieves credentials securely from Windows Credential Manager via Win32 `advapi32.dll` P/Invoke (`CredRead`).
  3. Authenticates via POST to `/v1/disp/rds/auth/logon` with CAM XML payload and extracts the `XSRF-TOKEN` session cookie.
  4. Iterates through all enabled reports and output formats across configured instances.
  5. Resolves dynamic date and prompt tokens in parameters and file output paths.
  6. Sends GET request to `/v1/disp/rds/outputFormat/{sourceType}/{source}/{format}?v=3` with `X-XSRF-TOKEN` header.
  7. Validates binary payload (checking against Cognos `RDS-ERR` XML responses, SOAP faults, and XML error envelopes) and writes files to disk.

### 3. `CognosConfigGui.ps1` & `Manage-CognosConfig.ps1`
Configuration management interfaces available in both graphical desktop and terminal interactive modes.

* **Graphical UI (`CognosConfigGui.ps1` / `Manage-CognosConfig.ps1 -Gui`)**:
  * Windows Forms desktop manager for managing reports, multi-instance endpoints, prompt discovery, live path previews, and on-demand test downloads.
  * **Batch Runner**: Main tab includes "Tải Tất Cả Báo Cáo" with multi-instance session caching.
  * **Visual Choice Picker**: Searchable checklist dialog (`Show-ChoiceSelectionDialog`) for browsing and selecting server-provided choice lists.
  * **Prompt Requirement Indicators**: Color-coded `BẮT BUỘC (*)` vs `Tùy chọn` prompt status in the parameters table.
* **Console UI (`Manage-CognosConfig.ps1`)**:
  * Terminal menu application with prompt parameter discovery, numbered server choices selection, live previews, and batch download triggering (Menu `[8]`).

### 4. `cognos-reports.json`
Schema structure for configuring multi-instance server endpoints, shared credentials, logging, and report definitions:

```json
{
  "CredentialTarget": "COGNOS",
  "DefaultInstance": "ODS",
  "Instances": {
    "ODS": {
      "CognosBaseUrl": "http://10.53.153.173/ibmcognos/bi",
      "Namespace": "BIDV"
    },
    "BIDV_Core": {
      "CognosBaseUrl": "http://10.53.153.174/ibmcognos/bi",
      "Namespace": "BIDV"
    }
  },
  "Logging": {
    "Enabled": true,
    "LogDirectory": ".\\Logs",
    "LogFileName": "CognosDownloader_{yyyyMMdd}.log",
    "LogLevel": "INFO",
    "RetentionDays": 30,
    "AuditCsvEnabled": true,
    "AuditCsvPath": ".\\Logs\\Audit_{yyyyMM}.csv"
  },
  "Reports": [
    {
      "Name": "Báo cáo nợ đến hạn, quá hạn hàng ngày",
      "Instance": "ODS",
      "Source": "i54414D93B29A4D2289C4E88469871644",
      "SourceType": "report",
      "Enabled": true,
      "Parameters": {
        "p_ReportDate": "{Yesterday}",
        "p_Branch": ["121000", "121150"]
      },
      "Formats": [
        {
          "Format": "xlsxData",
          "OutputPath": "D:\\CognosReports\\{Yesterday:yyyyMMdd}_{ReportName}.xlsx"
        }
      ]
    }
  ]
}
```

---

## Logging & Auditing System

The logging engine in `CognosCommon.ps1` provides unified console feedback, persistent file logs, machine-readable execution audits, and structured pipeline reports:

* **Console & File Logging (`Write-Log`)**:
  * Levels: `DEBUG`, `INFO`, `OK`, `WARN`, `ERROR` with configurable threshold filtering (`LogLevel`).
  * **Minimal Console Logging**: `DEBUG` statements (raw HTTP request URLs, headers, redirect hops, cookie internals) are recorded strictly to file and omitted from console output unless `$script:ConsoleDebug = $true`.
  * File format: `[yyyy-MM-dd HH:mm:ss] [LEVEL] Message`.
* **Execution Metrics & Audit Tracking (`Write-AuditLog`)**:
  * Appends execution records to a rolling monthly CSV (`Audit_{yyyyMM}.csv`).
  * Columns: `Timestamp`, `ReportName`, `Source`, `Format`, `Status`, `HttpStatusCode`, `FileSizeBytes`, `DurationMs`, `OutputPath`, `ErrorMessage`.
* **End-of-Run Execution Summary Table (`Write-ExecutionSummaryReport`)**:
  * Renders a clean ASCII summary table to console and appends it to the daily log upon completing a batch run:
    * Columns: `Report Name`, `Instance`, `Format`, `Status`, `Duration`, `Size`, `Output File`.
    * Bottom summary line: Succeeded/Failed count, Total Duration, and Total Download Size in MB.
  * Writes a structured `Logs/LatestRun.json` report containing complete batch metrics for ingestion by downstream ETL or monitoring tools.
* **Automatic Retention (`Invoke-LogRetentionCleanup`)**:
  * Automatically scans the log directory and purges log/audit files exceeding `RetentionDays` (default: 30 days).

---

## Automatic Retry Policy & Resilience

Report downloads across all scripts implement `Invoke-CognosReportDownloadWithRetry`:
* **Exponential Backoff**: When an intermittent network drop, HTTP 500/503, or Cognos queue congestion occurs, the downloader pauses (default: 5s initial delay, 2.0x multiplier) and retries up to 3 times before declaring failure.
* **Session Renewal**: Retries re-authenticate and acquire a fresh XSRF session token if the existing session was terminated on the server.

---

## Dynamic Token Engine

Both `CognosReportDownloader.ps1` and `Manage-CognosConfig.ps1` implement `Resolve-DynamicTokens` to dynamically substitute date and metadata tokens inside parameter values and output file paths:

| Token | Example Output | Description |
| :--- | :--- | :--- |
| `{Yesterday}` | `2026-08-20` | Previous calendar day (ISO `yyyy-MM-dd`) |
| `{Today}` | `2026-08-21` | Current calendar day (ISO `yyyy-MM-dd`) |
| `{MonthStart}` | `2026-08-01` | First day of current month |
| `{MonthEnd}` | `2026-08-31` | Last day of current month |
| `{Today-7d}` / `{Today+1d}` | `2026-08-14` | Day offset relative to current date |
| `{Yesterday:yyyyMMdd}` | `20260820` | Custom format specifier on named token |
| `{yyyyMMdd}`, `{yyyy-MM-dd}` | `20260821` | Standard timestamp tokens |
| `{HHmmss}` | `113000` | Standard time token |
| `{ReportName}` | `Daily Branch Summary` | Cleaned friendly report name |
| `{Source}` | `i54414D93B29A4D2289C4E88469871644` | Report ID / StoreID |
| `{Instance}` | `ODS` | Linked Cognos instance name |
| `{Format}` | `xlsxData` | Output format name |
| `{p_ParameterName}` | *evaluated value* | Dynamically injects another parameter's value into the path |
| `@D:\path\list.txt` / `{file:list.txt}` | *array of values* | Loads parameter values from external text file (supports `#`, `//`, `--`, `;` comments & comma/semicolon delimiters) |

---

## Supported Output Formats

* `xlsxData`: Excel Data format (clean tabular data suitable for ETL and programmatic ingestion).
* `spreadsheetML`: Excel XML format (retains formatting and report layout).
* `PDF`: Portable Document Format.
* `CSV`: Comma-Separated Values format.

---

## Security & Credential Management

* **No Plaintext Passwords**: Credentials are never stored in `cognos-reports.json` or source scripts.
* **Native Windows Credential Manager**:
  * Implemented via P/Invoke to Win32 APIs in `advapi32.dll` (`CredRead`, `CredWrite`, `CredFree`).
  * Type: `CRED_TYPE_GENERIC` (1).
  * Persistence: `CRED_PERSIST_LOCAL_MACHINE` (2).
  * Memory is explicitly freed using `Marshal::ZeroFreeBSTR` and `Marshal::FreeCoTaskMem` to avoid memory leakage.

---

## Agent Coding Guidelines

1. **PowerShell 5.1 & Core Compatibility**:
   * Always maintain backwards compatibility with Windows PowerShell 5.1 (`#requires -Version 5.1`).
   * Explicitly load `System.Net.Http` via `Add-Type -AssemblyName System.Net.Http`.
   * Enforce `Set-StrictMode -Version Latest` and `$ErrorActionPreference = 'Stop'`.
   * Avoid PowerShell 7+ specific operators (like ternary `? :` or null-coalescing `??`) inside `.ps1` scripts unless guarded or implemented using standard `if/else`.
2. **File Encoding Standard**:
   * **Mandatory UTF-8 with BOM (`utf-8-sig`)**: Always save all `.ps1` files with a UTF-8 BOM. Windows PowerShell 5.1 interprets UTF-8 files without BOM using the system ANSI code page (e.g. CP1252), causing Unicode string corruption and syntax parse errors.
3. **Error Handling & Resource Management**:
   * Always dispose `HttpClientHandler`, `HttpClient`, `HttpRequestMessage`, and `HttpResponseMessage` in `finally` blocks or via proper disposal routines.
   * Check HTTP status codes and inspect payload contents for `<rds:error`, `RDS-ERR`, `<soapenv:Fault>`, or XML error root tags on binary formats to avoid saving error HTML/XML pages as valid report binaries.
4. **Configuration Preservation**:
   * When modifying `cognos-reports.json` programmatically or via scripts, retain UTF-8 encoding and back up the previous config as `cognos-reports.json.bak`.
5. **Git Workflow & Commits**:
   * Do not commit or push every single change automatically. Only commit or push when explicitly requested by the user or when completing a defined milestone.