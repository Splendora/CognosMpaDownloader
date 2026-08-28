#requires -Version 5.1
<#
.SYNOPSIS
    Interactive test runner for OBIEE SOAP integration (MPA).
.DESCRIPTION
    Tests the modular OBIEE SOAP functions:
    1. Credentials retrieval from Windows Credential Manager ('MaCB').
    2. Connection & session acquisition (nQSessionService).
    3. Item info lookup on '/users/user/_portal/' (webCatalogService).
    4. Catalog listing on '/users/user/_portal' (webCatalogService).
    5. Table display (Show-ObieeCatalogItems).
    6. Safe session disconnect in finally block.
.EXAMPLE
    .\Test-ObieeSoap.ps1 -BaseUrl "http://10.53.153.175/analytics-ws/saw.dll"
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$BaseUrl,

    [Parameter()]
    [string]$CredentialTarget = 'MaCB',

    [Parameter()]
    [string]$ItemPath = '/users/user/_portal/',

    [Parameter()]
    [string]$FolderPath = '/users/user/_portal'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Dot-source ObieeCommon.ps1
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir "ObieeCommon.ps1")

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "     OBIEE SOAP Integration Test Runner (MPA)           " -ForegroundColor Cyan
Write-Host "========================================================`n" -ForegroundColor Cyan

# 1. Base URL Prompt if not provided
if ([string]::IsNullOrWhiteSpace($BaseUrl)) {
    $defaultUrl = "http://10.53.44.180:9704/analytics-ws/saw.dll"
    $inputUrl = Read-Host "Enter OBIEE SOAP Base URL [Default: $defaultUrl]"
    $BaseUrl = if ([string]::IsNullOrWhiteSpace($inputUrl)) { $defaultUrl } else { $inputUrl }
}

Write-Host "[1/5] Retrieving Credentials for Target '$CredentialTarget'..." -ForegroundColor Yellow
$cred = Get-ObieeCredential -Target $CredentialTarget
if ($null -eq $cred -or [string]::IsNullOrWhiteSpace($cred.Username)) {
    Write-Host "[-] No credential found in Windows Credential Manager for target '$CredentialTarget'." -ForegroundColor Red
    Write-Host "    Please ensure target '$CredentialTarget' is configured in Windows Credential Manager." -ForegroundColor Yellow
    exit 1
}

Write-Host "[+] Found credentials for user: $($cred.Username)" -ForegroundColor Green

$sessionId = $null
try {
    # 2. Connect
    Write-Host "`n[2/5] Connecting to OBIEE ($BaseUrl)..." -ForegroundColor Yellow
    $sessionId = Connect-Obiee -BaseUrl $BaseUrl -Username $cred.Username -Password $cred.Password
    Write-Host "[+] OBIEE SOAP LOGIN: SUCCESS" -ForegroundColor Green
    Write-Host "    Session ID length: $($sessionId.Length) chars" -ForegroundColor Gray

    # Dynamic username resolution for personal folders
    $resolvedItemPath = $ItemPath.Replace('/users/user/', "/users/$($cred.Username)/").Replace('{Username}', $cred.Username)
    $resolvedFolderPath = $FolderPath.Replace('/users/user/', "/users/$($cred.Username)/").Replace('{Username}', $cred.Username)

    # 3. Item Info
    Write-Host "`n[3/5] Inspecting Item Info for '$resolvedItemPath'..." -ForegroundColor Yellow
    try {
        $item = Get-ObieeItemInfo -BaseUrl $BaseUrl -Path $resolvedItemPath -SessionId $sessionId
        Write-Host "[+] Item found successfully:" -ForegroundColor Green
        Write-Host "    Caption   : $($item.Caption)" -ForegroundColor Cyan
        Write-Host "    Type      : $($item.Type)" -ForegroundColor Cyan
        Write-Host "    Signature : $($item.Signature)" -ForegroundColor Cyan
        Write-Host "    Path      : $($item.Path)" -ForegroundColor Gray
    }
    catch {
        Write-Host "[-] Failed to retrieve item info for '$resolvedItemPath': $_" -ForegroundColor Red
    }

    # 4. Catalog SubItems
    Write-Host "`n[4/5] Listing Catalog Items in '$resolvedFolderPath'..." -ForegroundColor Yellow
    try {
        $subItems = Get-ObieeCatalogItems -BaseUrl $BaseUrl -Path $resolvedFolderPath -SessionId $sessionId
        Write-Host "[+] Retrieved $($subItems.Count) item(s):" -ForegroundColor Green
        Show-ObieeCatalogItems -Items $subItems
    }
    catch {
        Write-Host "[-] Failed to list catalog items in '$resolvedFolderPath': $_" -ForegroundColor Red
        
        # Exploratory fallback if personal portal folder isn't found
        $fallbackFolders = @("/users/$($cred.Username)", "/shared", "/users")
        foreach ($fb in $fallbackFolders) {
            if ($fb -ne $resolvedFolderPath) {
                Write-Host "    Trying fallback folder '$fb'..." -ForegroundColor DarkGray
                try {
                    $fbItems = Get-ObieeCatalogItems -BaseUrl $BaseUrl -Path $fb -SessionId $sessionId
                    if ($null -ne $fbItems -and $fbItems.Count -gt 0) {
                        Write-Host "[+] Retrieved $($fbItems.Count) item(s) from '$fb':" -ForegroundColor Green
                        Show-ObieeCatalogItems -Items $fbItems
                        break
                    }
                }
                catch {
                    # Suppress secondary trial errors
                }
            }
        }
    }
}
finally {
    # 5. Disconnect
    if (-not [string]::IsNullOrWhiteSpace($sessionId)) {
        Write-Host "`n[5/5] Disconnecting OBIEE session..." -ForegroundColor Yellow
        $disconnected = Disconnect-Obiee -BaseUrl $BaseUrl -SessionId $sessionId
        if ($disconnected) {
            Write-Host "[+] OBIEE Logoff: SUCCESS (HTTP 200)" -ForegroundColor Green
        }
        else {
            Write-Host "[-] OBIEE Logoff completed with warning." -ForegroundColor Yellow
        }
    }
}

Write-Host "`n[+] Test execution completed." -ForegroundColor Cyan
