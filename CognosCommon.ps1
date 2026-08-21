<#
.SYNOPSIS
    Shared core module and helper functions for CognosDownloader automation suite.

.DESCRIPTION
    Provides unified implementations for:
    - Win32 Windows Credential Manager P/Invoke (advapi32.dll)
    - Dynamic Token & Date Resolution Engine
    - HTTP Client & Session Management (CookieContainer & redirect loops)
    - Cognos Mashup / RDS CAM XML Authentication & Prompt Parameter Discovery
    - Configuration File I/O with UTF-8 & Backup protection
    - Common Logging & CLI formatting
#>

Set-StrictMode -Version Latest

# Load System.Net.Http assembly (required for Windows PowerShell 5.1 compatibility)
Add-Type -AssemblyName System.Net.Http

# -----------------------------------------------------------------------------
# 1. Windows Credential Manager Native P/Invoke
# -----------------------------------------------------------------------------

if (-not ('CognosCredManager.NativeMethods' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

namespace CognosCredManager
{
    public static class NativeMethods
    {
        /*
         * Native Windows CREDENTIAL structure.
         * FILETIME is represented as Int64 to avoid PtrToStructure layout issues in PS 5.1.
         */
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct CREDENTIAL
        {
            public UInt32 Flags;
            public UInt32 Type;
            public IntPtr TargetName;
            public IntPtr Comment;
            public Int64 LastWritten;
            public UInt32 CredentialBlobSize;
            public IntPtr CredentialBlob;
            public UInt32 Persist;
            public UInt32 AttributeCount;
            public IntPtr Attributes;
            public IntPtr TargetAlias;
            public IntPtr UserName;
        }

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool CredRead(
            string TargetName,
            UInt32 Type,
            UInt32 Flags,
            out IntPtr Credential
        );

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool CredWrite(
            ref CREDENTIAL Credential,
            UInt32 Flags
        );

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern void CredFree(IntPtr Credential);
    }
}
"@
}

function Set-WindowsGenericCredential {
    param(
        [Parameter(Mandatory)]
        [string]$Target,

        [Parameter(Mandatory)]
        [string]$Username,

        [Parameter(Mandatory)]
        [Security.SecureString]$Password
    )

    $bstr = [IntPtr]::Zero
    $plain = $null
    $cred = New-Object CognosCredManager.NativeMethods+CREDENTIAL

    try {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)

        $cred.Flags = 0
        $cred.Type = 1                    # CRED_TYPE_GENERIC
        $cred.Persist = 2                 # CRED_PERSIST_LOCAL_MACHINE

        $cred.TargetName = [Runtime.InteropServices.Marshal]::StringToCoTaskMemUni($Target)
        $cred.UserName   = [Runtime.InteropServices.Marshal]::StringToCoTaskMemUni($Username)

        $blobBytes = [Text.Encoding]::Unicode.GetBytes($plain)
        $cred.CredentialBlobSize = [UInt32]$blobBytes.Length

        if ($blobBytes.Length -gt 0) {
            $cred.CredentialBlob = [Runtime.InteropServices.Marshal]::AllocCoTaskMem($blobBytes.Length)
            [Runtime.InteropServices.Marshal]::Copy($blobBytes, 0, $cred.CredentialBlob, $blobBytes.Length)
        }

        if (-not [CognosCredManager.NativeMethods]::CredWrite([ref]$cred, 0)) {
            $code = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
            throw "CredWrite failed. Windows error code: $code"
        }
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
        if ($cred.TargetName -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeCoTaskMem($cred.TargetName) }
        if ($cred.UserName -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeCoTaskMem($cred.UserName) }
        if ($cred.CredentialBlob -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeCoTaskMem($cred.CredentialBlob) }
        $plain = $null
    }
}

function Get-WindowsGenericCredential {
    param(
        [Parameter(Mandatory)]
        [string]$Target
    )

    $ptr = [IntPtr]::Zero

    try {
        $success = [CognosCredManager.NativeMethods]::CredRead($Target, 1, 0, [ref]$ptr)

        if (-not $success) {
            $code = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
            if ($code -eq 1168) { return $null } # ERROR_NOT_FOUND
            throw "CredRead failed. Windows error code: $code"
        }

        # Explicitly invoke PtrToStructure(IntPtr, Type) to avoid PowerShell 5.1 overload ambiguity
        $method = [Runtime.InteropServices.Marshal].GetMethod('PtrToStructure', [type[]]@([IntPtr], [Type]))
        $native = $method.Invoke($null, @($ptr, [CognosCredManager.NativeMethods+CREDENTIAL]))

        $username = $null
        $password = $null

        if ($native.UserName -ne [IntPtr]::Zero) {
            $username = [Runtime.InteropServices.Marshal]::PtrToStringUni($native.UserName)
        }

        if ($native.CredentialBlob -ne [IntPtr]::Zero -and $native.CredentialBlobSize -gt 0) {
            $blobBytes = New-Object byte[] $native.CredentialBlobSize
            [Runtime.InteropServices.Marshal]::Copy($native.CredentialBlob, $blobBytes, 0, [int]$native.CredentialBlobSize)
            $password = [Text.Encoding]::Unicode.GetString($blobBytes)
        }
        else {
            $password = ''
        }

        return [pscustomobject]@{
            Username = $username
            Password = (ConvertTo-SecureString $password -AsPlainText -Force)
        }
    }
    finally {
        if ($ptr -ne [IntPtr]::Zero) {
            [CognosCredManager.NativeMethods]::CredFree($ptr)
        }
    }
}

function Get-PlainTextFromSecureString {
    param(
        [Parameter(Mandatory)]
        [Security.SecureString]$SecureString
    )

    $bstr = [IntPtr]::Zero
    try {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

# -----------------------------------------------------------------------------
# 2. Dynamic Token & Date Resolution Engine
# -----------------------------------------------------------------------------

function Resolve-DynamicTokens {
    param(
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Text = '',

        $Report = $null,

        [AllowEmptyString()]
        [AllowNull()]
        [string]$Format = ''
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }

    $now = Get-Date
    $resolved = $Text

    # 1. Base relative dates
    $today = $now.Date
    $yesterday = $today.AddDays(-1)
    $monthStart = [DateTime]::new($today.Year, $today.Month, 1)
    $monthEnd = $monthStart.AddMonths(1).AddDays(-1)

    # 2. Base named tokens (ISO format: yyyy-MM-dd)
    $resolved = $resolved.Replace('{Yesterday}', $yesterday.ToString('yyyy-MM-dd'))
    $resolved = $resolved.Replace('{Today}', $today.ToString('yyyy-MM-dd'))
    $resolved = $resolved.Replace('{MonthStart}', $monthStart.ToString('yyyy-MM-dd'))
    $resolved = $resolved.Replace('{MonthEnd}', $monthEnd.ToString('yyyy-MM-dd'))

    # 3. Custom Date Format Tokens (e.g. {Yesterday:yyyyMMdd}, {Today:yyyyMMdd_HHmmss})
    $resolved = [regex]::Replace($resolved, '\{(Yesterday|Today)(:([^}]+))?\}', {
        param($match)
        $baseDate = if ($match.Groups[1].Value -eq 'Yesterday') { $yesterday } else { $today }
        $fmt = if ($match.Groups[3].Success) { $match.Groups[3].Value } else { 'yyyy-MM-dd' }
        return $baseDate.ToString($fmt)
    })

    # 4. Offset tokens (e.g. {Today-7d}, {Today-1d:yyyyMMdd})
    $resolved = [regex]::Replace($resolved, '\{Today([+-]\d+)d?(:([^}]+))?\}', {
        param($match)
        $days = [int]$match.Groups[1].Value
        $targetDate = $today.AddDays($days)
        $fmt = if ($match.Groups[3].Success) { $match.Groups[3].Value } else { 'yyyy-MM-dd' }
        return $targetDate.ToString($fmt)
    })

    # 5. Standard Date/Time Tokens
    $resolved = $resolved.Replace('{yyyy}', $now.ToString('yyyy'))
    $resolved = $resolved.Replace('{MM}', $now.ToString('MM'))
    $resolved = $resolved.Replace('{dd}', $now.ToString('dd'))
    $resolved = $resolved.Replace('{yyyyMM}', $now.ToString('yyyyMM'))
    $resolved = $resolved.Replace('{yyyy-MM}', $now.ToString('yyyy-MM'))
    $resolved = $resolved.Replace('{yyyyMMdd}', $now.ToString('yyyyMMdd'))
    $resolved = $resolved.Replace('{yyyy-MM-dd}', $now.ToString('yyyy-MM-dd'))
    $resolved = $resolved.Replace('{HHmmss}', $now.ToString('HHmmss'))

    # 6. Report Metadata & Prompt Parameter Tokens (if Report object provided)
    if ($null -ne $Report) {
        $reportName = Get-PropOrKey -Object $Report -Name 'Name'
        $reportSource = Get-PropOrKey -Object $Report -Name 'Source'
        $reportInst = Get-PropOrKey -Object $Report -Name 'Instance'

        if ([string]::IsNullOrWhiteSpace($reportName)) {
            $reportName = $reportSource
        }
        $resolved = $resolved.Replace('{ReportName}', [string]$reportName)
        $resolved = $resolved.Replace('{Source}', [string]$reportSource)
        if (-not [string]::IsNullOrWhiteSpace($reportInst)) {
            $resolved = $resolved.Replace('{Instance}', [string]$reportInst)
        }
        if (-not [string]::IsNullOrWhiteSpace($Format)) {
            $resolved = $resolved.Replace('{Format}', $Format)
        }

        $paramsObj = Get-PropOrKey -Object $Report -Name 'Parameters'
        if ($null -ne $paramsObj) {
            foreach ($prop in @(Get-ObjectKeyValuePairs -Object $paramsObj)) {
                $rawVal = ''
                if ($null -ne $prop.Value) {
                    if ($prop.Value -is [System.Collections.IEnumerable] -and -not ($prop.Value -is [string])) {
                        $evalItems = @()
                        foreach ($it in $prop.Value) {
                            if ($null -ne $it) {
                                $evalItems += (Resolve-DynamicTokens -Text ([string]$it))
                            }
                        }
                        $rawVal = $evalItems -join '_'
                    } else {
                        $rawVal = Resolve-DynamicTokens -Text ([string]$prop.Value)
                    }
                }
                $cleanVal = $rawVal -replace '[\\/:*?"<>|]', '-'
                $resolved = $resolved.Replace("{$($prop.Name)}", $cleanVal)
                if ($prop.Name.StartsWith('p_')) {
                    $resolved = $resolved.Replace("{$($prop.Name.Substring(2))}", $cleanVal)
                }
            }
        }
    }

    return $resolved
}

# -----------------------------------------------------------------------------
# 3. HTTP Client & Redirect Handler
# -----------------------------------------------------------------------------

function New-CognosHttpClient {
    param(
        [int]$TimeoutMinutes = 30
    )

    $cookieContainer = New-Object System.Net.CookieContainer

    $handler = New-Object System.Net.Http.HttpClientHandler -Property @{
        CookieContainer   = $cookieContainer
        AllowAutoRedirect = $false
        UseCookies        = $true
    }

    $client = New-Object System.Net.Http.HttpClient($handler)
    $client.Timeout = [TimeSpan]::FromMinutes($TimeoutMinutes)
    $client.DefaultRequestHeaders.Accept.Clear()
    $client.DefaultRequestHeaders.Accept.Add(
        [System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new('*/*')
    )

    return [pscustomobject]@{
        Client  = $client
        Handler = $handler
        Cookies = $cookieContainer
    }
}

function Invoke-CognosRequest {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] [ValidateSet('GET', 'POST')] [string]$Method,
        [Parameter(Mandatory)] [uri]$Uri,
        [hashtable]$Headers,
        [System.Net.Http.HttpContent]$Content,
        [int]$MaxRedirects = 10
    )

    $current = $Uri
    $redirectCount = 0

    while ($true) {
        $httpMethod = if ($Method -eq 'GET') { [System.Net.Http.HttpMethod]::Get } else { [System.Net.Http.HttpMethod]::Post }
        $request = New-Object System.Net.Http.HttpRequestMessage($httpMethod, $current)

        if ($Headers) {
            foreach ($key in $Headers.Keys) {
                [void]$request.Headers.TryAddWithoutValidation($key, [string]$Headers[$key])
            }
        }

        if ($null -ne $Content) {
            $request.Content = $Content
        }

        Write-Log "HTTP $Method -> $current" 'DEBUG'

        $response = $Context.Client.SendAsync($request).GetAwaiter().GetResult()
        $statusCode = [int]$response.StatusCode

        Write-Log "HTTP Status $statusCode từ $current" 'DEBUG'

        # Check HTTP 300..308 redirect codes directly
        if ($statusCode -in @(300, 301, 302, 303, 307, 308)) {
            if ($redirectCount -ge $MaxRedirects) {
                throw "Too many redirects while requesting $current"
            }

            $location = $response.Headers.Location
            if ($null -eq $location) {
                throw "Cognos returned HTTP $statusCode without a Location header."
            }

            $current = if ($location.IsAbsoluteUri) { $location } else { [uri]::new($current, $location) }
            $redirectCount++
            Write-Log "Đang chuyển hướng (#$redirectCount) tới $current" 'DEBUG'
            $response.Dispose()
            continue
        }

        return [pscustomobject]@{
            Response = $response
            FinalUri = $current
        }
    }
}

function Get-CookieValue {
    param(
        [Parameter(Mandatory)]
        [System.Net.CookieContainer]$Container,

        [Parameter(Mandatory)]
        [uri]$Uri,

        [Parameter(Mandatory)]
        [string]$Name
    )

    foreach ($cookie in $Container.GetCookies($Uri)) {
        if ($cookie.Name -ieq $Name) {
            return $cookie.Value
        }
    }
    return $null
}

# -----------------------------------------------------------------------------
# 4. Cognos Authentication & Prompt Discovery
# -----------------------------------------------------------------------------

function Invoke-CognosLogin {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] [string]$CognosBaseUrl,
        [Parameter(Mandatory)] [string]$Namespace,
        [Parameter(Mandatory)] [string]$Username,
        [Parameter(Mandatory)] [Security.SecureString]$Password
    )

    $plainPassword = Get-PlainTextFromSecureString $Password

    try {
        $xml = @"
<credentials xmlns="http://developer.cognos.com/schemas/ccs/auth/types/1">
  <credentialElements>
    <name>CAMNamespace</name>
    <label>Namespace:</label>
    <value>
      <actualValue>$([System.Security.SecurityElement]::Escape($Namespace))</actualValue>
    </value>
  </credentialElements>
  <credentialElements>
    <name>CAMUsername</name>
    <label>User ID:</label>
    <value>
      <actualValue>$([System.Security.SecurityElement]::Escape($Username))</actualValue>
    </value>
  </credentialElements>
  <credentialElements>
    <name>CAMPassword</name>
    <label>Password:</label>
    <value>
      <actualValue>$([System.Security.SecurityElement]::Escape($plainPassword))</actualValue>
    </value>
  </credentialElements>
</credentials>
"@

        $dict = [System.Collections.Generic.Dictionary[string, string]]::new()
        $dict['xmlData'] = $xml
        $form = New-Object System.Net.Http.FormUrlEncodedContent($dict)

        $url = [uri]::new(($CognosBaseUrl.TrimEnd('/') + '/v1/disp/rds/auth/logon'))
        Write-Log "Gửi thông tin xác thực CAM tới $url (Namespace: '$Namespace', User: '$Username')" 'DEBUG'

        $result = Invoke-CognosRequest -Context $Context -Method POST -Uri $url -Content $form
        $body = $result.Response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $status = [int]$result.Response.StatusCode
        $result.Response.Dispose()

        if ($status -ne 200) {
            throw "Cognos logon failed with HTTP $status. Server response: $body"
        }

        if ($body -notmatch 'accountInfo') {
            throw "Cognos did not return accountInfo. Authentication was not successful. Server response: $body"
        }

        $xsrf = Get-CookieValue -Container $Context.Cookies -Uri $url -Name 'XSRF-TOKEN'

        if ([string]::IsNullOrWhiteSpace($xsrf)) {
            Write-Log "Chưa thấy cookie XSRF-TOKEN từ logon POST, đang gửi probe GET..." 'DEBUG'
            $probeUrl = [uri]::new(($CognosBaseUrl.TrimEnd('/') + '/v1/disp/rds/auth/logon?xmlData=%3Ccredentials%2F%3E'))
            $probe = Invoke-CognosRequest -Context $Context -Method GET -Uri $probeUrl
            $probe.Response.Dispose()
            $xsrf = Get-CookieValue -Container $Context.Cookies -Uri $probeUrl -Name 'XSRF-TOKEN'
        }

        if ([string]::IsNullOrWhiteSpace($xsrf)) {
            throw 'Cognos authenticated successfully but no XSRF-TOKEN cookie was received.'
        }

        Write-Log "Nhận thành công XSRF-TOKEN: $xsrf" 'DEBUG'
        return $xsrf
    }
    finally {
        $plainPassword = $null
    }
}

function Invoke-CognosReportDownload {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] [string]$Url,
        [Parameter(Mandatory)] [string]$Xsrf,
        [Parameter(Mandatory)] [string]$Format,
        [int]$MaxAsyncPolls = 60,
        [int]$PollIntervalMs = 1500
    )

    $headers = @{ 'X-XSRF-TOKEN' = $Xsrf }
    $currentUri = [uri]$Url
    $pollCount = 0

    while ($true) {
        $result = Invoke-CognosRequest -Context $Context -Method GET -Uri $currentUri -Headers $headers
        $response = $result.Response
        $httpStatus = [int]$response.StatusCode
        $currentUri = $result.FinalUri

        if ($httpStatus -ne 200) {
            $body = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            $response.Dispose()
            throw "Máy chủ Cognos phản hồi mã lỗi HTTP $httpStatus. $body"
        }

        $bytes = $response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
        $contentType = if ($response.Content.Headers.ContentType) { $response.Content.Headers.ContentType.MediaType } else { '' }
        $response.Dispose()

        # Kiểm tra phản hồi dạng văn bản / XML (có thể là phiên bất đồng bộ, lỗi hoặc trang prompt)
        if ($contentType -match 'text|xml|html' -or ($bytes.Length -lt 20000 -and $bytes.Length -gt 0)) {
            $prefixLen = [Math]::Min($bytes.Length, 4096)
            $sample = [Text.Encoding]::UTF8.GetString($bytes, 0, $prefixLen)

            # 1. Kiểm tra lỗi nghiệp vụ từ Cognos
            if ($sample -match 'RDS-ERR' -or $sample -match '<rds:error' -or $sample -match '<soapenv:Fault>') {
                $fullText = [Text.Encoding]::UTF8.GetString($bytes)
                throw "Máy chủ Cognos trả về nội dung lỗi: $fullText"
            }

            # 2. Kiểm tra nếu Cognos trả về trang bắt nhập tham số (Prompt Page XML)
            if ($sample -match '<document[^>]*layoutData' -or $sample -match '<promptPages>' -or $sample -match '<p_date' -or $sample -match '<p_value') {
                throw "Cognos yêu cầu bổ sung hoặc chỉnh sửa tham số prompt (máy chủ trả về trang prompt thay vì dữ liệu báo cáo). Vui lòng kiểm tra lại giá trị các tham số bắt buộc trong cấu hình."
            }

            # 3. Kiểm tra nếu là phiên xử lý bất đồng bộ cần theo dõi (Async Session Polling)
            if ($sample -match '<rds:url>(.*?)</rds:url>' -or $sample -match '<rds:asyncSession') {
                if ($pollCount -ge $MaxAsyncPolls) {
                    throw "Quá thời gian chờ thực thi báo cáo bất đồng bộ ($MaxAsyncPolls lần thăm dò)."
                }

                $relUrl = if ($sample -match '<rds:url>(.*?)</rds:url>') { $matches[1].Replace('&amp;', '&') } else { $null }
                if (-not [string]::IsNullOrWhiteSpace($relUrl)) {
                    $nextUri = [uri]::new($currentUri, $relUrl)
                    $pollCount++
                    Write-Log "Đang theo dõi phiên tải báo cáo bất đồng bộ (#$pollCount): $nextUri" 'DEBUG'
                    Start-Sleep -Milliseconds $PollIntervalMs
                    $currentUri = $nextUri
                    continue
                }
            }

            # 4. Nếu là tệp XML nhưng định dạng yêu cầu không phải XML/spreadsheetML
            if ($sample -match '<\?xml' -and $format -in @('xlsxData', 'PDF', 'CSV')) {
                $fullText = [Text.Encoding]::UTF8.GetString($bytes)
                throw "Máy chủ Cognos trả về XML thay vì tệp nhị phân $format. Phản hồi: $fullText"
            }
        }

        return [pscustomobject]@{
            Bytes       = $bytes
            ContentType = $contentType
            HttpStatus  = $httpStatus
        }
    }
}

function Get-CognosReportParameters {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] [string]$BaseUrl,
        [Parameter(Mandatory)] [string]$SourceType,
        [Parameter(Mandatory)] [string]$Source,
        [Parameter(Mandatory)] [string]$Xsrf
    )

    $endpointType = if ($SourceType -ieq 'path') { 'path' } else { 'report' }
    $encodedSource = [uri]::EscapeDataString($Source)
    $url = [uri]::new(($BaseUrl.TrimEnd('/') + "/v1/disp/rds/reportPrompts/${endpointType}/${encodedSource}?v=3"))

    Write-Log "Yêu cầu schema tham số prompt từ: $url" 'DEBUG'

    $headers = @{ 'X-XSRF-TOKEN' = $Xsrf }
    $result = Invoke-CognosRequest -Context $Context -Method GET -Uri $url -Headers $headers
    $xmlText = $result.Response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
    $currentUri = $result.FinalUri
    $result.Response.Dispose()

    $pollCount = 0
    while ($xmlText -match '<rds:url>(.*?)</rds:url>' -and $pollCount -lt 10) {
        $relUrl = $matches[1].Replace('&amp;', '&')
        $sessionUrl = [uri]::new($currentUri, $relUrl)
        Write-Log "Đang theo dõi phiên prompt bất đồng bộ (#$($pollCount + 1)): $sessionUrl" 'DEBUG'
        Start-Sleep -Milliseconds 500
        $sessionResult = Invoke-CognosRequest -Context $Context -Method GET -Uri $sessionUrl -Headers $headers
        $xmlText = $sessionResult.Response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $currentUri = $sessionResult.FinalUri
        $sessionResult.Response.Dispose()
        $pollCount++
    }

    $params = [ordered]@{}
    try {
        [xml]$doc = $xmlText

        # 1. Quét các nút prompt có cấu trúc trong Cognos 11 LDX và RDS XML chuẩn
        $promptNodes = $doc.SelectNodes("//*[local-name()='p_date' or local-name()='p_value' or local-name()='p_txtbox' or local-name()='p_tree' or local-name()='selectDate' or local-name()='selectValue' or local-name()='textBox' or local-name()='prompt' or local-name()='promptControl' or local-name()='parameter' or local-name()='parameterControl']")
        
        # Nếu không tìm thấy bằng danh sách thẻ, tìm tất cả nút cha của <pname> hoặc <parameterName>
        if ($null -eq $promptNodes -or $promptNodes.Count -eq 0) {
            $promptNodes = $doc.SelectNodes("//*[local-name()='pname' or local-name()='parameterName']/..")
        }

        if ($null -ne $promptNodes -and $promptNodes.Count -gt 0) {
            foreach ($pnode in $promptNodes) {
                # Tránh xử lý nút con nếu đã lấy nút cha
                if ($pnode.LocalName -in @('pname', 'parameterName', 'name')) { continue }

                # Tên tham số
                $rawName = $null
                $nameNode = $pnode.SelectSingleNode(".//*[local-name()='pname' or local-name()='parameterName' or local-name()='name']")
                if ($null -ne $nameNode -and -not [string]::IsNullOrWhiteSpace($nameNode.InnerText)) {
                    $rawName = $nameNode.InnerText.Trim()
                } elseif ($pnode.Attributes['pname']) {
                    $rawName = $pnode.Attributes['pname'].Value.Trim()
                } elseif ($pnode.Attributes['parameterName']) {
                    $rawName = $pnode.Attributes['parameterName'].Value.Trim()
                } elseif ($pnode.Attributes['parameter']) {
                    $rawName = $pnode.Attributes['parameter'].Value.Trim()
                } elseif ($pnode.Attributes['name']) {
                    $rawName = $pnode.Attributes['name'].Value.Trim()
                }

                if ([string]::IsNullOrWhiteSpace($rawName)) { continue }
                $paramKey = if ($rawName.StartsWith('p_')) { $rawName } else { "p_$rawName" }

                # Bỏ qua nếu đã tồn tại và đã có choices
                if ($params.Contains($paramKey) -and $params[$paramKey].Choices.Length -gt 0) {
                    continue
                }

                # Kiểu prompt (type)
                $pType = 'textBox'
                $nodeLocal = $pnode.LocalName.ToLowerInvariant()
                if ($nodeLocal -in @('p_date', 'selectdate')) {
                    $pType = 'selectDate'
                } elseif ($nodeLocal -in @('p_value', 'selectvalue')) {
                    $pType = 'selectValue'
                } elseif ($nodeLocal -in @('p_txtbox', 'textbox')) {
                    $pType = 'textBox'
                } elseif ($nodeLocal -in @('p_tree', 'tree')) {
                    $pType = 'tree'
                } else {
                    $typeNode = $pnode.SelectSingleNode(".//*[local-name()='type' or local-name()='uiType' or local-name()='selectUI' or local-name()='dateUI']")
                    if ($null -ne $typeNode -and -not [string]::IsNullOrWhiteSpace($typeNode.InnerText)) {
                        $pType = $typeNode.InnerText.Trim()
                    } elseif ($pnode.Attributes['type']) {
                        $pType = $pnode.Attributes['type'].Value.Trim()
                    }
                }

                # Kiểm tra cờ DateTime
                $isDateTime = $false
                $dtNode = $pnode.SelectSingleNode(".//*[local-name()='DateTime']")
                if ($null -ne $dtNode -and ($dtNode.InnerText -match 'true|1')) {
                    $isDateTime = $true
                }

                # Multi-select / cardinality
                $isMulti = $false
                $multiNode = $pnode.SelectSingleNode(".//*[local-name()='multi' or local-name()='multiSelect' or local-name()='cardinality']")
                if ($null -ne $multiNode) {
                    $isMulti = ($multiNode.InnerText -match 'multi|zeroOrMore|oneOrMore|true|1')
                } elseif ($pnode.Attributes['multi']) {
                    $isMulti = ($pnode.Attributes['multi'].Value -match 'true|1')
                } elseif ($pnode.Attributes['multiSelect']) {
                    $isMulti = ($pnode.Attributes['multiSelect'].Value -match 'true|1')
                } elseif ($pnode.Attributes['cardinality']) {
                    $isMulti = ($pnode.Attributes['cardinality'].Value -match 'multi|zeroOrMore|oneOrMore')
                }

                # Bắt buộc hay Tùy chọn (Required vs Optional)
                $isRequired = $true
                $reqNode = $pnode.SelectSingleNode(".//*[local-name()='req' or local-name()='required' or local-name()='optional' or local-name()='usage']")
                if ($null -ne $reqNode) {
                    if ($reqNode.LocalName -in @('req', 'required')) {
                        $isRequired = ($reqNode.InnerText -notmatch 'false|0')
                    } elseif ($reqNode.LocalName -eq 'optional') {
                        $isRequired = ($reqNode.InnerText -match 'false|0')
                    } else {
                        $isRequired = ($reqNode.InnerText -match 'required|true|1')
                    }
                } elseif ($pnode.Attributes['req']) {
                    $isRequired = ($pnode.Attributes['req'].Value -notmatch 'false|0')
                } elseif ($pnode.Attributes['required']) {
                    $isRequired = ($pnode.Attributes['required'].Value -notmatch 'false|0')
                } elseif ($pnode.Attributes['optional']) {
                    $isRequired = ($pnode.Attributes['optional'].Value -match 'false|0')
                } elseif ($pnode.Attributes['usage']) {
                    $isRequired = ($pnode.Attributes['usage'].Value -match 'required')
                }

                # Lựa chọn danh mục (Choices / Options trong LDX: <selOptions><sval> hoặc <choices><option>)
                $choices = [System.Collections.Generic.List[object]]::new()
                $optionNodes = $pnode.SelectNodes(".//*[local-name()='sval' or local-name()='option' or local-name()='choice' or local-name()='item' or local-name()='val']")
                if ($null -ne $optionNodes) {
                    foreach ($opt in $optionNodes) {
                        $useVal = $null
                        $dispVal = $null
                        
                        $useNode = $opt.SelectSingleNode(".//*[local-name()='use' or local-name()='useValue' or local-name()='val']")
                        if ($null -ne $useNode -and -not [string]::IsNullOrWhiteSpace($useNode.InnerText)) {
                            $useVal = $useNode.InnerText.Trim()
                        } elseif ($opt.Attributes['use']) {
                            $useVal = $opt.Attributes['use'].Value.Trim()
                        } elseif ($opt.Attributes['useValue']) {
                            $useVal = $opt.Attributes['useValue'].Value.Trim()
                        } elseif ($opt.Attributes['value']) {
                            $useVal = $opt.Attributes['value'].Value.Trim()
                        }

                        $dispNode = $opt.SelectSingleNode(".//*[local-name()='disp' or local-name()='display' or local-name()='displayValue' or local-name()='caption']")
                        if ($null -ne $dispNode -and -not [string]::IsNullOrWhiteSpace($dispNode.InnerText)) {
                            $dispVal = $dispNode.InnerText.Trim()
                        } elseif ($opt.Attributes['disp']) {
                            $dispVal = $opt.Attributes['disp'].Value.Trim()
                        } elseif ($opt.Attributes['display']) {
                            $dispVal = $opt.Attributes['display'].Value.Trim()
                        } elseif ($opt.Attributes['displayValue']) {
                            $dispVal = $opt.Attributes['displayValue'].Value.Trim()
                        } elseif ($opt.Attributes['caption']) {
                            $dispVal = $opt.Attributes['caption'].Value.Trim()
                        }

                        if ($null -eq $dispVal -or [string]::IsNullOrWhiteSpace($dispVal)) {
                            $dispVal = $useVal
                        }

                        if (-not [string]::IsNullOrWhiteSpace($useVal)) {
                            $choices.Add([pscustomobject]@{
                                Use     = $useVal
                                Display = $dispVal
                            })
                        }
                    }
                }

                # Giá trị mặc định (Default Values trong LDX: <defOptions><dval> hoặc <default>)
                $defaultVal = $null
                $defaultNode = $pnode.SelectSingleNode(".//*[local-name()='defOptions' or local-name()='default' or local-name()='defaultValue' or local-name()='defaultValues']")
                if ($null -ne $defaultNode) {
                    $defItems = $defaultNode.SelectNodes(".//*[local-name()='use' or local-name()='useValue' or local-name()='dval' or local-name()='sval' or local-name()='item' or local-name()='val']")
                    if ($null -ne $defItems -and $defItems.Count -gt 0) {
                        $defList = @()
                        foreach ($di in $defItems) {
                            if (-not [string]::IsNullOrWhiteSpace($di.InnerText)) {
                                $defList += $di.InnerText.Trim()
                            }
                        }
                        if ($defList.Count -gt 0) {
                            $defaultVal = if ($isMulti) { $defList } else { $defList[0] }
                        }
                    } elseif (-not [string]::IsNullOrWhiteSpace($defaultNode.InnerText)) {
                        $defaultVal = $defaultNode.InnerText.Trim()
                    }
                }

                if (-not $defaultVal) {
                    if ($isDateTime) {
                        $defaultVal = '{Yesterday}T00:00:00'
                    } elseif ($pType -match 'date' -or $paramKey -match 'date|ngay|time|denhan|quahan') {
                        $defaultVal = '{Yesterday}'
                    }
                }

                $params[$paramKey] = [pscustomobject]@{
                    Name          = $paramKey
                    Type          = $pType
                    IsMultiSelect = $isMulti
                    IsRequired    = $isRequired
                    IsDateTime    = $isDateTime
                    DefaultValue  = $defaultVal
                    Choices       = $choices.ToArray()
                }
            }
        }

        # 2. Fallback an toàn (chỉ tìm thẻ <pname> hoặc <parameterName>, tuyệt đối không quét thẻ <name> chung chung)
        if ($params.Count -eq 0) {
            $pnodes = $doc.SelectNodes("//*[local-name()='pname' or local-name()='parameterName']")
            if ($null -ne $pnodes) {
                foreach ($node in $pnodes) {
                    $rawName = $node.InnerText.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($rawName)) {
                        $paramKey = if ($rawName.StartsWith('p_')) { $rawName } else { "p_$rawName" }
                        if (-not $params.Contains($paramKey)) {
                            $defaultVal = if ($paramKey -match 'date|ngay|time|denhan|quahan') { '{Yesterday}' } else { '' }
                            $params[$paramKey] = [pscustomobject]@{
                                Name          = $paramKey
                                Type          = 'textBox'
                                IsMultiSelect = $false
                                IsRequired    = $false
                                IsDateTime    = $false
                                DefaultValue  = $defaultVal
                                Choices       = @()
                            }
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-Log "Lỗi phân tích cú pháp XML prompt cho '$Source': $($_.Exception.Message)" 'WARN'
    }

    Write-Log "Đã trích xuất ($($params.Count)) tham số từ schema: $(($params.Keys) -join ', ')" 'DEBUG'
    return $params
}

# -----------------------------------------------------------------------------
# 5. Configuration File I/O
# -----------------------------------------------------------------------------

function Load-CognosConfig {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $resolvedPath = if ([System.IO.Path]::IsPathRooted($Path)) { $Path } else { [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $Path)) }
    if (-not (Test-Path -LiteralPath $resolvedPath)) { return $null }
    $raw = Get-Content -LiteralPath $resolvedPath -Raw -Encoding UTF8
    return ($raw | ConvertFrom-Json)
}

function Save-CognosConfig {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        $Config
    )

    $resolvedPath = if ([System.IO.Path]::IsPathRooted($Path)) { $Path } else { [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $Path)) }
    if (Test-Path -LiteralPath $resolvedPath) {
        Copy-Item -LiteralPath $resolvedPath -Destination "$resolvedPath.bak" -Force
    }
    $jsonText = $Config | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($resolvedPath, $jsonText, [System.Text.Encoding]::UTF8)
}

function Get-PropOrKey {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    if ($Object.PSObject -and $Object.PSObject.Properties[$Name]) {
        return $Object.PSObject.Properties[$Name].Value
    }
    try {
        if ($null -ne $Object.$Name) {
            return $Object.$Name
        }
    } catch { }
    return $null
}

function Get-RequiredProperty {
    param(
        [Parameter(Mandatory)]
        $Object,

        [Parameter(Mandatory)]
        [string]$Name
    )

    $val = Get-PropOrKey -Object $Object -Name $Name

    if ($null -eq $val -or [string]::IsNullOrWhiteSpace([string]$val)) {
        throw "Lỗi cấu hình: Thuộc tính bắt buộc '$Name' không tồn tại hoặc bị để trống."
    }

    return [string]$val
}

function Get-ObjectKeyValuePairs {
    param($Object)
    if ($null -eq $Object) { return @() }
    $pairs = [System.Collections.Generic.List[object]]::new()
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($k in $Object.Keys) {
            $pairs.Add([pscustomobject]@{
                Name  = [string]$k
                Value = $Object[$k]
            })
        }
        return $pairs
    }
    if ($Object.PSObject -and $Object.PSObject.Properties) {
        foreach ($prop in $Object.PSObject.Properties) {
            $pairs.Add([pscustomobject]@{
                Name  = [string]$prop.Name
                Value = $prop.Value
            })
        }
        return $pairs
    }
    return @()
}

function Set-ObjectProperty {
    param(
        [Parameter(Mandatory)]
        $Object,

        [Parameter(Mandatory)]
        [string]$Name,

        $Value
    )

    if ($Object -is [System.Collections.IDictionary]) {
        $Object[$Name] = $Value
    }
    elseif ($Object.PSObject -and $Object.PSObject.Properties[$Name]) {
        $Object.PSObject.Properties[$Name].Value = $Value
    }
    else {
        $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
    }
}

function Add-QueryParameter {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]]$Parts,

        [Parameter(Mandatory)]
        [string]$Name,

        $Value
    )

    if ($null -eq $Value) { return }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        foreach ($item in $Value) {
            if ($null -ne $item) {
                $strItem = ([string]$item).Trim()
                if (-not [string]::IsNullOrWhiteSpace($strItem)) {
                    $Parts.Add("$([uri]::EscapeDataString($Name))=$([uri]::EscapeDataString($strItem))")
                }
            }
        }
    }
    else {
        $str = ([string]$Value).Trim()
        if (-not [string]::IsNullOrWhiteSpace($str)) {
            # Tự động tách chuỗi danh sách phân tách bằng dấu phẩy thành các query param lặp lại (trừ chuỗi ISO date)
            if ($str -match ',' -and $str -notmatch '^\d{4}-\d{2}-\d{2}') {
                $subItems = @($str -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                foreach ($sub in $subItems) {
                    $Parts.Add("$([uri]::EscapeDataString($Name))=$([uri]::EscapeDataString($sub))")
                }
            } else {
                $Parts.Add("$([uri]::EscapeDataString($Name))=$([uri]::EscapeDataString($str))")
            }
        }
    }
}

function Get-ReportDefinitionUrl {
    param(
        [Parameter(Mandatory)] [string]$CognosBaseUrl,
        [Parameter(Mandatory)] $Report,
        [Parameter(Mandatory)] [string]$Format
    )

    $sourceType = Get-PropOrKey -Object $Report -Name 'SourceType'
    if ([string]::IsNullOrWhiteSpace($sourceType)) {
        $sourceType = 'report'
    }

    $source = Get-RequiredProperty -Object $Report -Name 'Source'
    $encodedSource = [uri]::EscapeDataString($source)
    $encodedFormat = [uri]::EscapeDataString($Format)

    $parts = New-Object 'System.Collections.Generic.List[string]'
    $parts.Add('v=3')
    $parts.Add('prompt=false')
    $parts.Add('async=off')

    # Tính toán giá trị dynamic token (VD: {Yesterday}) trong tham số prompt
    $paramsObj = Get-PropOrKey -Object $Report -Name 'Parameters'
    if ($null -ne $paramsObj) {
        foreach ($pair in @(Get-ObjectKeyValuePairs -Object $paramsObj)) {
            $paramName = if ($pair.Name.StartsWith('p_')) { $pair.Name } else { "p_$($pair.Name)" }
            $val = $pair.Value
            if ($null -ne $val) {
                if ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string])) {
                    $evalList = New-Object 'System.Collections.Generic.List[string]'
                    foreach ($item in $val) {
                        if ($null -ne $item) {
                            $evalItem = Resolve-DynamicTokens -Text ([string]$item) -Report $Report -Format $Format
                            if (-not [string]::IsNullOrWhiteSpace($evalItem)) {
                                $evalList.Add($evalItem)
                            }
                        }
                    }
                    if ($evalList.Count -gt 0) {
                        Add-QueryParameter -Parts $parts -Name $paramName -Value $evalList
                    }
                }
                else {
                    $rawVal = [string]$val
                    if (-not [string]::IsNullOrWhiteSpace($rawVal)) {
                        $evaluatedVal = Resolve-DynamicTokens -Text $rawVal -Report $Report -Format $Format
                        Add-QueryParameter -Parts $parts -Name $paramName -Value $evaluatedVal
                    }
                }
            }
        }
    }

    $optionsObj = Get-PropOrKey -Object $Report -Name 'Options'
    if ($null -ne $optionsObj) {
        foreach ($pair in @(Get-ObjectKeyValuePairs -Object $optionsObj)) {
            Add-QueryParameter -Parts $parts -Name $pair.Name -Value $pair.Value
        }
    }

    $fullUrl = (
        "$($CognosBaseUrl.TrimEnd('/'))" +
        "/v1/disp/rds/outputFormat/" +
        "$sourceType/" +
        "$encodedSource/" +
        "${encodedFormat}?" +
        ($parts -join '&')
    )

    Write-Log "Đã tạo REST URL tải báo cáo: $fullUrl" 'DEBUG'
    return $fullUrl
}

function Get-CognosInstances {
    param(
        [Parameter(Mandatory)]
        $Config
    )

    $instances = [ordered]@{}
    $rawInstances = Get-PropOrKey -Object $Config -Name 'Instances'

    if ($null -ne $rawInstances) {
        if ($rawInstances -is [System.Collections.IDictionary]) {
            foreach ($k in $rawInstances.Keys) {
                $instObj = $rawInstances[$k]
                $baseUrl = [string](Get-PropOrKey -Object $instObj -Name 'CognosBaseUrl')
                $ns = [string](Get-PropOrKey -Object $instObj -Name 'Namespace')

                if ([string]::IsNullOrWhiteSpace($baseUrl)) {
                    throw "Lỗi cấu hình: Máy chủ '$k' thiếu thuộc tính 'CognosBaseUrl'."
                }
                if ([string]::IsNullOrWhiteSpace($ns)) {
                    throw "Lỗi cấu hình: Máy chủ '$k' thiếu thuộc tính 'Namespace'."
                }

                $instances[$k] = [pscustomobject]@{
                    Name          = $k
                    CognosBaseUrl = $baseUrl.TrimEnd('/')
                    Namespace     = $ns.Trim()
                }
            }
        }
        else {
            foreach ($prop in @(Get-ObjectKeyValuePairs -Object $rawInstances)) {
                $instObj = $prop.Value
                $baseUrl = [string](Get-PropOrKey -Object $instObj -Name 'CognosBaseUrl')
                $ns = [string](Get-PropOrKey -Object $instObj -Name 'Namespace')

                if ([string]::IsNullOrWhiteSpace($baseUrl)) {
                    throw "Lỗi cấu hình: Máy chủ '$($prop.Name)' thiếu thuộc tính 'CognosBaseUrl'."
                }
                if ([string]::IsNullOrWhiteSpace($ns)) {
                    throw "Lỗi cấu hình: Máy chủ '$($prop.Name)' thiếu thuộc tính 'Namespace'."
                }

                $instances[$prop.Name] = [pscustomobject]@{
                    Name          = $prop.Name
                    CognosBaseUrl = $baseUrl.TrimEnd('/')
                    Namespace     = $ns.Trim()
                }
            }
        }
    }

    if ($instances.Count -eq 0) {
        throw "Lỗi cấu hình: Không có danh sách 'Instances' hợp lệ trong tệp cấu hình."
    }

    return $instances
}
function Get-CognosReportInstance {
    param(
        [Parameter(Mandatory)] $Config,
        $Report = $null
    )

    $allInstances = Get-CognosInstances -Config $Config

    # 1. Trường hợp báo cáo chỉ định rõ tên máy chủ
    $repInstance = if ($Report) { Get-PropOrKey -Object $Report -Name 'Instance' } else { $null }
    if ($null -ne $repInstance -and -not [string]::IsNullOrWhiteSpace([string]$repInstance)) {
        $targetName = [string]$repInstance
        if ($allInstances.Contains($targetName)) {
            return $allInstances[$targetName]
        }
        $available = ($allInstances.Keys) -join ', '
        $repName = if ($Report) { Get-PropOrKey -Object $Report -Name 'Name' } else { $null }
        if ([string]::IsNullOrWhiteSpace($repName)) {
            $repName = if ($Report) { Get-PropOrKey -Object $Report -Name 'Source' } else { 'Không xác định' }
        }
        throw "Lỗi cấu hình: Báo cáo '$repName' liên kết với máy chủ không tồn tại '$targetName'. Các máy chủ hiện có: $available."
    }

    # 2. Kiểm tra DefaultInstance trong tệp cấu hình
    $defInstance = Get-PropOrKey -Object $Config -Name 'DefaultInstance'
    if ($null -ne $defInstance -and -not [string]::IsNullOrWhiteSpace([string]$defInstance)) {
        $defName = [string]$defInstance
        if ($allInstances.Contains($defName)) {
            return $allInstances[$defName]
        }
        $available = ($allInstances.Keys) -join ', '
        throw "Lỗi cấu hình: Máy chủ mặc định 'DefaultInstance' ('$defName') không tồn tại trong danh sách 'Instances' ($available)."
    }

    # 3. Nếu chỉ có đúng 1 máy chủ được cấu hình, tự động sử dụng máy chủ đó
    if ($allInstances.Count -eq 1) {
        foreach ($k in $allInstances.Keys) {
            return $allInstances[$k]
        }
    }

    $repName = if ($Report) { Get-PropOrKey -Object $Report -Name 'Name' } else { $null }
    if ([string]::IsNullOrWhiteSpace($repName)) { $repName = 'Không xác định' }
    throw "Lỗi cấu hình: Báo cáo '$repName' không chỉ định 'Instance', và cấu hình cũng không có 'DefaultInstance'."
}

# -----------------------------------------------------------------------------
# 6. Logging Engine
# -----------------------------------------------------------------------------

# Internal state for logging configuration
$script:LoggingState = @{
    Enabled         = $false
    LogFilePath     = $null
    LogLevel        = 'INFO'
    AuditCsvEnabled = $false
    AuditCsvPath    = $null
}

$script:LogLevelWeights = @{
    'DEBUG' = 10
    'INFO'  = 20
    'OK'    = 20
    'WARN'  = 30
    'ERROR' = 40
}

function Initialize-CognosLogging {
    param(
        $LoggingConfig,
        [string]$BaseDirectory = (Get-Location).Path
    )

    if ($null -eq $LoggingConfig) {
        $LoggingConfig = [pscustomobject]@{
            Enabled         = $true
            LogDirectory    = (Join-Path $BaseDirectory 'Logs')
            LogFileName     = 'CognosDownloader_{yyyyMMdd}.log'
            LogLevel        = 'INFO'
            RetentionDays   = 30
            AuditCsvEnabled = $true
            AuditCsvPath    = (Join-Path (Join-Path $BaseDirectory 'Logs') 'Audit_{yyyyMM}.csv')
        }
    }

    $logEnabled = Get-PropOrKey -Object $LoggingConfig -Name 'Enabled'
    $enabled = if ($null -ne $logEnabled) { [bool]$logEnabled } else { $true }

    if (-not $enabled) {
        $script:LoggingState.Enabled = $false
        return
    }

    $logDirProp = Get-PropOrKey -Object $LoggingConfig -Name 'LogDirectory'
    $logDirRaw = if ($null -ne $logDirProp -and -not [string]::IsNullOrWhiteSpace([string]$logDirProp)) {
        [string]$logDirProp
    } else {
        (Join-Path $BaseDirectory 'Logs')
    }

    $logDir = if ([System.IO.Path]::IsPathRooted($logDirRaw)) { $logDirRaw } else { (Join-Path $BaseDirectory $logDirRaw) }
    $logDir = Resolve-DynamicTokens -Text $logDir

    if (-not (Test-Path -LiteralPath $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    $logFileProp = Get-PropOrKey -Object $LoggingConfig -Name 'LogFileName'
    $logFileNameRaw = if ($null -ne $logFileProp -and -not [string]::IsNullOrWhiteSpace([string]$logFileProp)) {
        [string]$logFileProp
    } else {
        'CognosDownloader_{yyyyMMdd}.log'
    }

    $logFileName = Resolve-DynamicTokens -Text $logFileNameRaw
    $logFilePath = Join-Path $logDir $logFileName

    $logLevelProp = Get-PropOrKey -Object $LoggingConfig -Name 'LogLevel'
    $logLevel = if ($null -ne $logLevelProp -and -not [string]::IsNullOrWhiteSpace([string]$logLevelProp)) {
        ([string]$logLevelProp).ToUpperInvariant()
    } else {
        'INFO'
    }

    $auditEnabledProp = Get-PropOrKey -Object $LoggingConfig -Name 'AuditCsvEnabled'
    $auditEnabled = if ($null -ne $auditEnabledProp) { [bool]$auditEnabledProp } else { $true }

    $auditPath = $null
    if ($auditEnabled) {
        $auditProp = Get-PropOrKey -Object $LoggingConfig -Name 'AuditCsvPath'
        $auditRaw = if ($null -ne $auditProp -and -not [string]::IsNullOrWhiteSpace([string]$auditProp)) {
            [string]$auditProp
        } else {
            (Join-Path $logDir 'Audit_{yyyyMM}.csv')
        }
        $auditResolved = Resolve-DynamicTokens -Text $auditRaw
        $auditPath = if ([System.IO.Path]::IsPathRooted($auditResolved)) { $auditResolved } else { (Join-Path $BaseDirectory $auditResolved) }
        $auditParent = Split-Path -Parent $auditPath
        if ($auditParent -and -not (Test-Path -LiteralPath $auditParent)) {
            New-Item -ItemType Directory -Path $auditParent -Force | Out-Null
        }
    }

    $retentionProp = Get-PropOrKey -Object $LoggingConfig -Name 'RetentionDays'
    $retentionDays = if ($null -ne $retentionProp -and [int]$retentionProp -gt 0) {
        [int]$retentionProp
    } else {
        30
    }

    $script:LoggingState.Enabled         = $true
    $script:LoggingState.LogFilePath     = $logFilePath
    $script:LoggingState.LogLevel        = $logLevel
    $script:LoggingState.AuditCsvEnabled = $auditEnabled
    $script:LoggingState.AuditCsvPath    = $auditPath

    Invoke-LogRetentionCleanup -LogDirectory $logDir -RetentionDays $retentionDays
}

function Invoke-LogRetentionCleanup {
    param(
        [Parameter(Mandatory)]
        [string]$LogDirectory,

        [int]$RetentionDays = 30
    )

    if ($RetentionDays -le 0 -or -not (Test-Path -LiteralPath $LogDirectory)) { return }

    try {
        $cutoff = (Get-Date).AddDays(-$RetentionDays)
        $oldFiles = Get-ChildItem -LiteralPath $LogDirectory -File -ErrorAction SilentlyContinue |
            Where-Object { ($_.Extension -ieq '.log' -or $_.Extension -ieq '.csv') -and $_.LastWriteTime -lt $cutoff }

        foreach ($file in $oldFiles) {
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        # Silently continue on cleanup errors
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('DEBUG', 'INFO', 'OK', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $prefix = switch ($Level) {
        'DEBUG' { '[DEBUG]' }
        'OK'    { '[OK]   ' }
        'WARN'  { '[WARN] ' }
        'ERROR' { '[ERROR]' }
        default { '[INFO] ' }
    }

    $color = switch ($Level) {
        'DEBUG' { 'DarkGray' }
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'White' }
    }

    # Console output
    Write-Host "$prefix $Message" -ForegroundColor $color

    # File output
    if ($script:LoggingState.Enabled -and $script:LoggingState.LogFilePath) {
        $msgWeight = if ($script:LogLevelWeights.ContainsKey($Level)) { $script:LogLevelWeights[$Level] } else { 20 }
        $minWeight = if ($script:LogLevelWeights.ContainsKey($script:LoggingState.LogLevel)) { $script:LogLevelWeights[$script:LoggingState.LogLevel] } else { 20 }

        if ($msgWeight -ge $minWeight) {
            $timestamp = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
            $logLine = "[$timestamp] $prefix $Message" + [Environment]::NewLine
            try {
                [System.IO.File]::AppendAllText($script:LoggingState.LogFilePath, $logLine, [System.Text.Encoding]::UTF8)
            }
            catch {
                # Fallback to avoid breaking execution
            }
        }
    }
}

function Write-AuditLog {
    param(
        [Parameter(Mandatory)] [string]$ReportName,
        [Parameter(Mandatory)] [string]$Source,
        [Parameter(Mandatory)] [string]$Format,
        [Parameter(Mandatory)] [ValidateSet('SUCCESS', 'FAILED')] [string]$Status,
        [int]$HttpStatusCode = 0,
        [int64]$FileSizeBytes = 0,
        [int64]$DurationMs = 0,
        [string]$OutputPath = '',
        [string]$ErrorMessage = ''
    )

    if (-not $script:LoggingState.AuditCsvEnabled -or [string]::IsNullOrWhiteSpace($script:LoggingState.AuditCsvPath)) {
        return
    }

    try {
        $csvPath = $script:LoggingState.AuditCsvPath
        $parent = Split-Path -Parent $csvPath
        if ($parent -and -not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }

        # Write CSV header if file is new or empty
        if (-not (Test-Path -LiteralPath $csvPath) -or (Get-Item -LiteralPath $csvPath).Length -eq 0) {
            $header = "Timestamp,ReportName,Source,Format,Status,HttpStatusCode,FileSizeBytes,DurationMs,OutputPath,ErrorMessage" + [Environment]::NewLine
            [System.IO.File]::WriteAllText($csvPath, $header, [System.Text.Encoding]::UTF8)
        }

        $nowIso = [DateTime]::Now.ToString('yyyy-MM-ddTHH:mm:sszzz')
        
        $escapeCsv = {
            param([string]$val)
            if ($val -match '[",\r\n]') {
                return '"' + $val.Replace('"', '""') + '"'
            }
            return $val
        }

        $row = @(
            (& $escapeCsv $nowIso),
            (& $escapeCsv $ReportName),
            (& $escapeCsv $Source),
            (& $escapeCsv $Format),
            (& $escapeCsv $Status),
            $HttpStatusCode,
            $FileSizeBytes,
            $DurationMs,
            (& $escapeCsv $OutputPath),
            (& $escapeCsv $ErrorMessage)
        ) -join ','

        [System.IO.File]::AppendAllText($csvPath, $row + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
    }
    catch {
        # Fallback to avoid interrupting core execution
    }
}
