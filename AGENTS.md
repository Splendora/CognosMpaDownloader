# AGENTS.md

## Repository Overview

**CognosDownloader** is a PowerShell-based automation tool and configuration manager for IBM Cognos Analytics 11. It automates the extraction and downloading of Cognos reports across multiple formats (`xlsxData`, `spreadsheetML`, `PDF`, `CSV`) via the Cognos RDS (Report Data Service / Mashup) REST interface, resolving dynamic date tokens and securely managing credentials using the native Windows Credential Manager.

---

## File Structure & Component Map

```
CognosDownloader/
├── CognosReportDownloader.ps1   # Headless batch execution script for automated report downloads
├── Manage-CognosConfig.ps1      # Interactive CLI console manager for configuration & prompt inspection
├── cognos-reports.json          # Configuration file defining server endpoints and report definitions
└── AGENTS.md                    # Repository documentation and instructions for AI agents
```

---

## Core Components & Architecture

### 1. `CognosReportDownloader.ps1`
Headless execution script designed for scheduled tasks (e.g., Windows Task Scheduler) or CI/CD pipelines.

* **Parameters**:
  * `-ConfigPath <string>`: Path to JSON configuration file (default: `.\cognos-reports.json`).
  * `-SetupCredential`: Interactive prompt to store credentials in Windows Credential Manager and exit.
  * `-TestConnection`: Verifies network connectivity and authentication against Cognos without running reports.
* **Execution Flow**:
  1. Loads configuration from `cognos-reports.json`.
  2. Retrieves credentials securely from Windows Credential Manager via Win32 `advapi32.dll` P/Invoke (`CredRead`).
  3. Authenticates via POST to `/v1/disp/rds/auth/logon` with CAM XML payload and extracts the `XSRF-TOKEN` session cookie.
  4. Iterates through all enabled reports and output formats in parallel/sequential batches.
  5. Resolves dynamic date and prompt tokens in parameters and file output paths.
  6. Sends GET request to `/v1/disp/rds/outputFormat/{sourceType}/{source}/{format}?v=3` with `X-XSRF-TOKEN` header.
  7. Validates binary payload (checking against Cognos `RDS-ERR` XML responses) and writes files to disk.

### 2. `Manage-CognosConfig.ps1`
Interactive terminal application for managing report configurations without manual JSON editing.

* **Menu Actions**:
  1. `[1] List / View all configured reports`: Displays all configured reports with live evaluation preview of parameters and output paths.
  2. `[2] Add a new report`: Prompts for StoreID/ReportID, connects to Cognos RDS (`/v1/disp/rds/reportPrompts/`), auto-detects prompt parameters, offers date presets (`{Yesterday}`, `{Today}`, `{MonthStart}`), and sets up default output path templates.
  3. `[3] Edit an existing report`: Modifies parameter values, toggles enabled/disabled state, updates output path templates, or re-syncs prompt parameters from the server.
  4. `[4] Remove a report`: Deletes a report definition from the config.
  5. `[5] Test Cognos connection & authentication`: Validates current credentials and server connectivity.
  6. `[6] Re-configure server URL / Credentials`: Re-initializes server URL, namespace, and stored Windows generic credential.

### 3. `cognos-reports.json`
Schema structure for configuring server endpoints and report definitions:

```json
{
  "CognosBaseUrl": "http://<cognos-host>/ibmcognos/bi",
  "Namespace": "<Cognos-CAM-Namespace>",
  "CredentialTarget": "CognosReportAutomation:<cognos-host>",
  "Reports": [
    {
      "Name": "Descriptive Report Name",
      "Source": "<StoreID-or-Path>",
      "SourceType": "report",
      "Enabled": true,
      "Parameters": {
        "p_ReportDate": "{Yesterday}",
        "p_Branch": ""
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
| `{Format}` | `xlsxData` | Output format name |
| `{p_ParameterName}` | *evaluated value* | Dynamically injects another parameter's value into the path |

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
2. **Error Handling & Resource Management**:
   * Always dispose `HttpClientHandler`, `HttpClient`, `HttpRequestMessage`, and `HttpResponseMessage` in `finally` blocks or via proper disposal routines.
   * Check HTTP status codes and inspect payload contents for `<rds:error` or `RDS-ERR` tokens to avoid saving error HTML/XML pages as valid report binaries.
3. **Configuration Preservation**:
   * When modifying `cognos-reports.json` programmatically or via scripts, retain UTF-8 encoding and back up the previous config as `cognos-reports.json.bak`.
