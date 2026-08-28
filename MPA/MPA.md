Understood. You wanted a **technical handoff/specification for the next LLM**, not a project summary.

# LLM Handoff Specification — OBIEE SOAP

## Objective

Continue building a PowerShell integration with legacy OBIEE SOAP Web Services.

Immediate task:

> Refactor the working script into independently testable functions, then stitch them together.

Eventually retrieve the OBIEE `ROA` analysis as CSV.

---

## Environment

* Windows PowerShell
* Legacy Oracle BI / OBIEE
* WSDL: `obiee-v10.wsdl`
* SOAP endpoint base:

```text
http://IP/analytics-ws/saw.dll
```

Services:

```text
?SoapImpl=nQSessionService
?SoapImpl=webCatalogService
?SoapImpl=analysisExportViewsService
```

WSDL namespace:

```xml
urn://oracle.bi.webservices/v10
```

Use:

```xml
xmlns:sawsoap="urn://oracle.bi.webservices/v10"
```

Do not invent or substitute another namespace.

---

## Authentication

Credentials come from Windows Credential Manager.

Credential target:

```text
COGNOS
```

The existing authentication implementation is **working** and must be preserved during refactoring.

Known successful output:

```text
OBIEE SOAP LOGIN: SUCCESS
Session ID received: YES
Session ID length: 48
```

The password can contain `@` and other special characters. It must be treated as credential data, not interpolated into URLs or otherwise escaped unnecessarily.

Do not print the actual password.

---

## Session

Successful login returns a session ID.

All subsequent SOAP operations should use that session ID.

Logoff is also working:

```text
Logoff HTTP Status: 200
```

Use `try/finally` so logoff occurs even when a later operation fails.

---

## Proven Catalog Target

The known analysis is:

```text
/users/user/_portal/ROA
```

Catalog information already confirmed:

```text
type      : Object
caption   : ROA
signature : queryitem1
```

Do **not** implement ROA searching. The path is already known.

---

## Proven `getItemInfo`

`webCatalogService.getItemInfo` has already been successfully tested against:

```text
/users/user/_portal/ROA
```

WSDL request:

```xml
<getItemInfo>
    <path>...</path>
    <resolveLinks>...</resolveLinks>
    <sessionID>...</sessionID>
</getItemInfo>
```

Definition:

```xml
<xsd:element name="getItemInfo">
    <xsd:complexType>
        <xsd:sequence>
            <xsd:element name="path" type="xsd:string"/>
            <xsd:element name="resolveLinks" type="xsd:boolean"/>
            <xsd:element name="sessionID"
                         nillable="true"
                         type="xsd:string"/>
        </xsd:sequence>
    </xsd:complexType>
</xsd:element>
```

Test result:

```text
ROA RETRIEVAL TEST: SUCCESS
Returned path matches requested path.
```

---

## Next Component: `getSubItems`

WSDL definition:

```xml
<xsd:complexType name="GetSubItemsParams">
    <xsd:sequence>
        <xsd:element name="filter"
                     nillable="true"
                     type="sawsoap:GetSubItemsFilter"/>
        <xsd:element name="includeACL"
                     type="xsd:boolean"/>
        <xsd:element name="withPermission"
                     type="xsd:int"/>
        <xsd:element name="withPermissionMask"
                     type="xsd:int"/>
        <xsd:element name="withAttributes"
                     type="xsd:int"/>
        <xsd:element name="withAttributesMask"
                     type="xsd:int"/>
        <xsd:element name="preserveOriginalLinkPath"
                     type="xsd:boolean"/>
    </xsd:sequence>
</xsd:complexType>
```

Request:

```xml
<xsd:element name="getSubItems">
    <xsd:complexType>
        <xsd:sequence>
            <xsd:element name="path" type="xsd:string"/>
            <xsd:element name="mask" type="xsd:string"/>
            <xsd:element name="resolveLinks" type="xsd:boolean"/>
            <xsd:element name="options"
                         nillable="true"
                         type="sawsoap:GetSubItemsParams"/>
            <xsd:element name="sessionID"
                         nillable="true"
                         type="xsd:string"/>
        </xsd:sequence>
    </xsd:complexType>
</xsd:element>
```

Response contains:

```xml
<itemInfo .../>
```

Test path:

```text
/users/user/_portal
```

The immediate goal is simply to display the returned catalog objects.

**Do not add recursive traversal yet.**

---

## Refactoring Target

Build these functions:

```powershell
Get-ObieeCredential
Connect-Obiee
Get-ObieeItemInfo
Get-ObieeCatalogItems
Show-ObieeCatalogItems
Disconnect-Obiee
```

Later:

```powershell
Read-ObieeObject
Export-ObieeAnalysis
Save-ObieeCsv
```

Each function should have one responsibility.

---

## Architecture

```text
Credential Manager
       ↓
Connect-Obiee
       ↓
Session ID
       ↓
Catalog functions
       ↓
Report/object retrieval
       ↓
Analysis export
       ↓
CSV
       ↓
Logoff
```

Keep authentication, SOAP requests, XML parsing, catalog logic, and file output separate.

---

## Eventual CSV Export

The WSDL confirms `analysisExportViewsService` supports:

```text
PDF
MHT
EXCEL2007
CSV
FLASH
SVG
GIF
PNG
JPEG
SVGFOP
VML
```

`AnalysisExportExecutionOptions`:

```text
async
useMtom
refresh
```

`AnalysisExportResult` contains:

```text
viewData
mimeType
queryID
exportStatus
```

Statuses:

```text
InProgress
Error
Done
```

Do not implement the export portion until the catalog/object retrieval layer is stable.

---

## Critical Constraints

1. **Preserve the currently working authentication code.**
2. Do not change the credential mechanism.
3. Do not hard-code passwords.
4. Do not search for ROA.
5. Do not add recursion yet.
6. Use the exact WSDL namespace:
   `urn://oracle.bi.webservices/v10`
7. Use `webCatalogService` for catalog operations.
8. Use the existing session ID.
9. Always log off in `finally`.
10. Build and test each function independently before stitching them together.

### Current starting point

**Authentication: proven.**
**Logoff: proven.**
**`getItemInfo(ROA)`: proven.**
**Next: implement/test `getSubItems(/users/user/_portal)`.**
