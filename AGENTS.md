# AGENTS.md

## Repository Overview

**CognosMpaDownloader** is a unified PowerShell-based automation suite and configuration manager for both **IBM Cognos Analytics 11** (via REST RDS / Mashup API) and **Oracle Business Intelligence Enterprise Edition / MPA** (via SOAP Web Services). It automates report extraction across multiple formats (`xlsxData`, `spreadsheetML`, `EXCEL2007`, `PDF`, `CSV`, `MHT`), resolves dynamic date tokens, and securely manages credentials using the native Windows Credential Manager.

```powershell
# Prerequisites: Enable PowerShell script execution
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

---

## File Structure & Component Map

```
CognosMpaDownloader/
├── README.md                    # Comprehensive repository documentation (Vietnamese)
├── AGENTS.md                    # Technical documentation and rules for AI agents
│
├── COGNOS/                      # IBM Cognos Analytics 11 Automation Suite
│   ├── COGNOS.md                # Comprehensive Cognos documentation (Vietnamese)
│   ├── CognosCommon.ps1         # Shared core module (Credentials, HTTP REST client, dynamic tokens, CAM login, prompt inspector, XML prompt builder)
│   ├── CognosReportDownloader.ps1 # Headless batch execution script for scheduled tasks
│   ├── Manage-CognosConfig.ps1  # Interactive CLI console manager for configuration & test runner
│   ├── CognosConfigGui.ps1      # Windows Forms GUI configuration manager, choice picker & batch runner
│   ├── Run-DailyPipeline.ps1    # Automated pipeline coordinating Cognos downloads and Excel VBA macros
│   ├── cognos-reports.json      # Configuration file defining server endpoints and report definitions
│   └── Logs/                    # Execution logs, audit CSV files, and LatestRun.json
│
└── MPA/                         # Oracle BI / MPA (OBIEE) Automation Suite
    ├── MPA.md                   # Comprehensive MPA / OBIEE documentation (Vietnamese)
    ├── ObieeCommon.ps1          # Shared core module (SOAP v10 envelopes, session management, catalog traversal, analysis export, tokens, logging)
    ├── MpaReportDownloader.ps1  # Headless batch execution script for MPA reports
    ├── Manage-MpaConfig.ps1     # Interactive CLI console manager for MPA catalog & config
    ├── MpaConfigGui.ps1         # Windows Forms GUI configuration manager, catalog browser & batch runner
    ├── Test-ObieeSoap.ps1       # Diagnostic & testing script for OBIEE SOAP Web Services
    ├── mpa-reports.json         # Configuration file defining MPA endpoints and report definitions
    ├── Reports/                 # Default directory for downloaded MPA reports
    └── Logs/                    # Execution logs, audit CSV files, and LatestRun.json
```

---

## Core Components & Architecture

### 1. COGNOS Suite (`COGNOS/`)
* **`CognosCommon.ps1`**:
  * Win32 Windows Credential Manager P/Invoke (`advapi32.dll` `CredRead`, `CredWrite`, `CredFree`).
  * Dynamic token resolution engine (`Resolve-DynamicTokens`) for dates and metadata.
  * CAM XML credentials payload generation and login authentication (`Invoke-CognosLogin`).
  * Structured Cognos prompt inspection (`Get-CognosReportParameters`): cardinality, requirement status, choice lists, defaults.
  * Unified XML Prompt Serialization (`New-CognosPromptAnswersXml`): serializes prompt parameters to `<promptAnswers><promptValues>`.
  * Parameter Submission (`Get-CognosReportRequest`): Pushes prompt XML via HTTP POST body `xmlData=` using `ByteArrayContent` to bypass .NET 64KB URI limit.
  * Exponential backoff retry policy (`Invoke-CognosReportDownloadWithRetry`).
  * Logging, monthly rolling audit CSV (`Audit_{yyyyMM}.csv`), execution summary (`LatestRun.json`), and log retention cleanup.
* **`CognosReportDownloader.ps1`**: Headless batch downloader for Task Scheduler / CI/CD.
* **`CognosConfigGui.ps1` & `Manage-CognosConfig.ps1`**: GUI and CLI management tools.
* **`Run-DailyPipeline.ps1`**: Coordinates multi-step Cognos report extraction and Excel VBA macro automation.

### 2. MPA / OBIEE Suite (`MPA/`)
* **`ObieeCommon.ps1`**:
  * SOAP Web Services client targeting `urn://oracle.bi.webservices/v10` (`sawsoap`).
  * Session lifecycle management via `nQSessionService` (`logon`, `keepAlive`, `logoff` in `finally`).
  * Web Catalog traversal via `webCatalogService` (`getItemInfo`, `getSubItems`).
  * Report data extraction via `analysisExportViewsService` (`CSV`, `EXCEL2007`, `PDF`, `MHT`).
  * Dynamic token resolution (`{Yesterday}`, `{Today}`, `{MonthStart}`, `{Username}`, etc.).
  * Exponential backoff retry and session recovery (`Invoke-MpaReportDownloadWithRetry`).
  * Logging, monthly audit CSV, `LatestRun.json`, and automatic log retention.
* **`MpaReportDownloader.ps1`**: Headless batch execution script for MPA.
* **`MpaConfigGui.ps1` & `Manage-MpaConfig.ps1`**: GUI (with visual Catalog tree browser) and CLI management tools.
* **`Test-ObieeSoap.ps1`**: Diagnostic script validating logon, catalog info, sub-items, and export.

---

## Agent Coding Guidelines

1. **PowerShell 5.1 & Core Compatibility**:
   * Always maintain backwards compatibility with Windows PowerShell 5.1 (`#requires -Version 5.1`).
   * Explicitly load `System.Net.Http` via `Add-Type -AssemblyName System.Net.Http`.
   * Enforce `Set-StrictMode -Version Latest` and `$ErrorActionPreference = 'Stop'`.
   * Avoid PowerShell 7+ specific operators (like ternary `? :` or null-coalescing `??`) inside `.ps1` scripts.
2. **File Encoding Standard**:
   * **Mandatory UTF-8 with BOM (`utf-8-sig`)**: Always save all `.ps1` files with a UTF-8 BOM. Windows PowerShell 5.1 interprets UTF-8 files without BOM using the system ANSI code page (e.g. CP1252), causing Unicode string corruption and syntax parse errors.
3. **Error Handling & Resource Management**:
   * Always dispose `HttpClientHandler`, `HttpClient`, `HttpRequestMessage`, and `HttpResponseMessage` in `finally` blocks.
   * For OBIEE SOAP, always invoke `logoff` in `finally` blocks to release WebLogic server sessions.
   * Check HTTP status codes and inspect payload contents for `<rds:error`, `RDS-ERR`, `<soapenv:Fault>`, or XML error root tags on binary formats to avoid saving error HTML/XML pages as valid report binaries.
4. **Cognos POST Payload Strictness**:
   * When pushing `xmlData` payloads to Cognos RDS Mashup endpoints via HTTP POST, the `Content-Type` header MUST be exactly `application/x-www-form-urlencoded`.
   * **Do NOT append `; charset=utf-8`**. IBM WebSphere/Tomcat servers often silently drop the POST body if the charset is appended, resulting in reports executing with no parameters (producing 0 MB data outputs).
5. **OBIEE SOAP WSDL Strictness**:
   * Namespace MUST be `urn://oracle.bi.webservices/v10` (`sawsoap`).
   * Do not change or invent other namespaces.