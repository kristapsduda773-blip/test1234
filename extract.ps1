<#
.SYNOPSIS
    Classify the provided Azure AD groups as on-premises synchronized or cloud-only.

.DESCRIPTION
    Fetches group metadata from Microsoft Graph for the supplied display names and determines
    whether each group originates from on-premises AD (DirSync) or is created directly in Azure AD.
    Uses the onPremisesSyncEnabled/onPremisesSecurityIdentifier flags to categorize each group.

.PARAMETER GroupNames
    Display names of the groups that need to be evaluated. Defaults to the list provided by the user.

.PARAMETER OutputPath
    Optional path to a CSV file that will contain the classification results. Defaults to
    ./graph-group-origin.csv in the current directory.

.PARAMETER ExcelOutputPath
    Optional path to an XLSX file. When provided, the script also writes a native Excel workbook
    without requiring Excel or third-party modules.

.PARAMETER PassThru
    Write the resulting objects to the pipeline in addition to exporting them to disk.

.PARAMETER SkipLogin
    Assume that Connect-MgGraph has already been called. Use this when running inside an
    automation context that handles authentication separately.

.EXAMPLE
    # Connect to Graph and classify the default group list
    Connect-MgGraph -Scopes "Group.Read.All"
    ./extract.ps1 -PassThru

.EXAMPLE
    # Use a custom list and capture the output CSV path
    ./extract.ps1 -GroupNames "DEV-ATD","SQL-00-PBI-Sync" -OutputPath ./custom.csv -PassThru

.EXAMPLE
    # Export only to Excel (skipping CSV) and emit objects to the pipeline
    ./extract.ps1 -OutputPath '' -ExcelOutputPath ./group-origin.xlsx -PassThru
#>

[CmdletBinding()]
param(
    [string[]]$GroupNames = @(
        "DEV-ATD",
        "DEV-AVIS-cloud",
        "DEV-BDAS-cloud",
        "DEV-CAKEHR-cloud",
        "DEV-DELTAPV-cloud",
        "DEV-ESTAPIKS2-cloud",
        "DEV-Evo-Roads",
        "DEV-FITS-cloud",
        "DEV-INTRANET-cloud",
        "DEV-LAMBDAPV-cloud",
        "DEV-MANAGEMENT-cloud",
        "DEV-NILDA2-cloud",
        "DEV-OPVS-CargoRail",
        "DEV-PRESERVICA-cloud",
        "DEV-VADDVS",
        "Dots-Sales",
        "Product Group",
        "BW-DEV-ATD",
        "BW-DEV-Common",
        "BW-DEV-EvoRoads",
        "BW-DEV-Kappa",
        "BW-DEV-LDz-OPVS",
        "BW-DEV-SAGE-MAGS",
        "DEV-AIHEN-cloud",
        "DEV-AIROS",
        "DEV-Digitalizacija",
        "DEV-EXT-AKKA-LAA",
        "DEV-External-ATD-SMARTIN",
        "DEV-External-SMARTIN-DESIGN",
        "DEV-EXT-Estapiks2",
        "DEV-EXT-Fits",
        "DEV-EXT-KAMIS",
        "DEV-EXT-Peruza",
        "DEV-EXT-Preservica",
        "DEV-FITS",
        "DEV-IC-FITS",
        "SQL-00-PBI-Sync"
    ),
    [AllowEmptyString()]
    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath "graph-group-origin.csv"),
    [string]$ExcelOutputPath,
    [switch]$PassThru,
    [switch]$SkipLogin
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-GraphModuleLoaded {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft.Graph PowerShell SDK is required. Install it via 'Install-Module Microsoft.Graph -Scope CurrentUser'."
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop | Out-Null
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop | Out-Null
}

function Ensure-GraphConnection {
    param([string[]]$Scopes = @("Group.Read.All"))

    if ($SkipLogin) {
        return
    }

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context -or -not $context.Account) {
        Write-Verbose "Connecting to Microsoft Graph..."
        Connect-MgGraph -Scopes $Scopes | Out-Null
    }
}

function Resolve-OutputPath {
    param([Parameter(Mandatory = $true)][string]$PathValue)

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return $PathValue
    }

    return (Join-Path -Path (Get-Location) -ChildPath $PathValue)
}

$script:ZipAssembliesLoaded = $false
function Ensure-ZipAssembliesLoaded {
    if (-not $script:ZipAssembliesLoaded) {
        Add-Type -AssemblyName System.IO.Compression | Out-Null
        Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
        $script:ZipAssembliesLoaded = $true
    }
}

function Add-ZipEntry {
    param(
        [Parameter(Mandatory = $true)]$Archive,
        [Parameter(Mandatory = $true)][string]$EntryName,
        [Parameter(Mandatory = $true)][string]$Content
    )

    $entry = $Archive.CreateEntry($EntryName)
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $writer = New-Object System.IO.StreamWriter($entry.Open(), $utf8)
    try {
        $writer.Write($Content)
    }
    finally {
        $writer.Dispose()
    }
}

function Get-ExcelColumnName {
    param([Parameter(Mandatory = $true)][int]$Index)

    if ($Index -lt 1) {
        return "A"
    }

    $name = ""
    $remaining = $Index
    while ($remaining -gt 0) {
        $remaining--
        $name = [char](65 + ($remaining % 26)) + $name
        $remaining = [math]::Floor($remaining / 26)
    }

    return $name
}

function Escape-WorksheetValue {
    param([string]$Value)

    if ($null -eq $Value) {
        return ""
    }

    $escaped = [System.Security.SecurityElement]::Escape($Value)
    $escaped = $escaped -replace "`r", "&#13;"
    $escaped = $escaped -replace "`n", "&#10;"
    return $escaped
}

function New-WorksheetXml {
    param(
        [System.Collections.IEnumerable]$Rows,
        [string[]]$Columns
    )

    $rowArray = @($Rows)
    if (-not $Columns -or $Columns.Length -eq 0) {
        if ($rowArray.Length -gt 0) {
            $Columns = $rowArray[0].PSObject.Properties.Name
        }
        else {
            $Columns = @("RequestedName", "Origin")
        }
    }

    $columnCount = $Columns.Length
    if ($columnCount -lt 1) {
        $Columns = @("RequestedName")
        $columnCount = 1
    }

    $dataRowCount = $rowArray.Length
    $maxRow = $dataRowCount + 1
    $lastColumnLetter = Get-ExcelColumnName -Index $columnCount
    $dimensionRef = "A1:$lastColumnLetter$maxRow"

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
    [void]$sb.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.AppendLine("  <dimension ref=""$dimensionRef""/>")
    [void]$sb.AppendLine('  <sheetViews>')
    [void]$sb.AppendLine('    <sheetView workbookViewId="0"/>')
    [void]$sb.AppendLine('  </sheetViews>')
    [void]$sb.AppendLine('  <sheetFormatPr defaultRowHeight="15"/>')
    [void]$sb.AppendLine('  <sheetData>')

    # Header row
    [void]$sb.AppendLine('    <row r="1">')
    for ($i = 0; $i -lt $columnCount; $i++) {
        $columnName = $Columns[$i]
        $cellRef = (Get-ExcelColumnName -Index ($i + 1)) + "1"
        $value = Escape-WorksheetValue -Value $columnName
        [void]$sb.AppendLine("      <c r=""$cellRef"" t=""inlineStr""><is><t xml:space=""preserve"">$value</t></is></c>")
    }
    [void]$sb.AppendLine('    </row>')

    $rowNumber = 1
    foreach ($row in $rowArray) {
        $rowNumber++
        [void]$sb.AppendLine("    <row r=""$rowNumber"">")
        for ($i = 0; $i -lt $columnCount; $i++) {
            $columnName = $Columns[$i]
            $cellRef = (Get-ExcelColumnName -Index ($i + 1)) + $rowNumber
            $cellValue = $null
            if ($null -ne $row -and $row.PSObject.Properties[$columnName]) {
                $cellValue = $row.$columnName
            }
            if ($null -eq $cellValue) {
                $cellValue = ""
            }
            $textValue = Escape-WorksheetValue -Value ([string]$cellValue)
            [void]$sb.AppendLine("      <c r=""$cellRef"" t=""inlineStr""><is><t xml:space=""preserve"">$textValue</t></is></c>")
        }
        [void]$sb.AppendLine('    </row>')
    }

    [void]$sb.AppendLine('  </sheetData>')
    [void]$sb.AppendLine('</worksheet>')

    return $sb.ToString()
}

function Write-XlsxFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][System.Collections.IEnumerable]$Rows,
        [Parameter(Mandatory = $true)][string[]]$Columns,
        [string]$WorksheetName = "Groups"
    )

    Ensure-ZipAssembliesLoaded

    $safeWorksheetName = if ([string]::IsNullOrWhiteSpace($WorksheetName)) { "Sheet1" } else { $WorksheetName }
    $safeWorksheetName = ($safeWorksheetName -replace "[\[\]\:\*\?\/\\]", "_")
    if ($safeWorksheetName.Length -gt 31) {
        $safeWorksheetName = $safeWorksheetName.Substring(0, 31)
    }

    $worksheetXml = New-WorksheetXml -Rows $Rows -Columns $Columns

    $contentTypesXml = @"
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"@

    $rootRelsXml = @"
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"@

    $workbookXml = @"
<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="$safeWorksheetName" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"@

    $workbookRelsXml = @"
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"@

    $stylesXml = @"
<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="2">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
"@

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }

    $fileStream = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
    try {
        $archive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Create, $false)
        try {
            Add-ZipEntry -Archive $archive -EntryName "[Content_Types].xml" -Content $contentTypesXml
            Add-ZipEntry -Archive $archive -EntryName "_rels/.rels" -Content $rootRelsXml
            Add-ZipEntry -Archive $archive -EntryName "xl/workbook.xml" -Content $workbookXml
            Add-ZipEntry -Archive $archive -EntryName "xl/_rels/workbook.xml.rels" -Content $workbookRelsXml
            Add-ZipEntry -Archive $archive -EntryName "xl/styles.xml" -Content $stylesXml
            Add-ZipEntry -Archive $archive -EntryName "xl/worksheets/sheet1.xml" -Content $worksheetXml
        }
        finally {
            $archive.Dispose()
        }
    }
    finally {
        $fileStream.Dispose()
    }
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory = $true)]$InputObject,
        [Parameter(Mandatory = $true)][string[]]$PropertyName
    )

    foreach ($name in $PropertyName) {
        $prop = $InputObject.PSObject.Properties[$name]
        if ($prop) {
            return $prop.Value
        }

        if ($InputObject.PSObject.Properties['AdditionalProperties']) {
            $additional = $InputObject.AdditionalProperties
            if ($additional -and $additional.ContainsKey($name)) {
                return $additional[$name]
            }
        }
    }

    return $null
}

function Get-GroupMatches {
    param([Parameter(Mandatory = $true)][string]$DisplayName)

    $escaped = $DisplayName -replace "'", "''"
    $filter = "displayName eq '$escaped'"

    $requestParams = @{
        Filter = $filter
        Property = @(
            "id",
            "displayName",
            "createdDateTime",
            "groupTypes",
            "onPremisesSyncEnabled",
            "onPremisesSecurityIdentifier",
            "securityEnabled",
            "mailEnabled"
        )
        ConsistencyLevel = "eventual"
        ErrorAction = "Stop"
    }

    $result = Get-MgGroup @requestParams
    if ($null -eq $result) {
        return @()
    }

    return @($result)
}

function Classify-Group {
    param(
        [Parameter(Mandatory = $true)][string]$RequestedName,
        [Parameter(Mandatory = $true)]$Group,
        [int]$MatchIndex = 1,
        [int]$TotalMatches = 1
    )

    $onPremSync = Get-ObjectPropertyValue -InputObject $Group -PropertyName @("OnPremisesSyncEnabled", "onPremisesSyncEnabled")
    $onPremSid = Get-ObjectPropertyValue -InputObject $Group -PropertyName @("OnPremisesSecurityIdentifier", "onPremisesSecurityIdentifier")
    $groupTypes = Get-ObjectPropertyValue -InputObject $Group -PropertyName @("GroupTypes", "groupTypes")

    $origin = "Cloud"
    if (($onPremSync -eq $true) -or (-not [string]::IsNullOrWhiteSpace($onPremSid))) {
        $origin = "OnPrem"
    }

    $groupTypeString = $null
    if ($groupTypes) {
        $groupTypeString = ($groupTypes -join '|')
    }

    $matchIndexValue = 1
    if ($TotalMatches -gt 1) {
        $matchIndexValue = $MatchIndex
    }

    [PSCustomObject]@{
        RequestedName = $RequestedName
        ResolvedName = $Group.DisplayName
        GroupId = $Group.Id
        Origin = $origin
        OnPremisesSyncEnabled = $onPremSync
        OnPremisesSecurityIdentifier = $onPremSid
        GroupTypes = $groupTypeString
        MatchIndex = $matchIndexValue
        TotalMatches = $TotalMatches
        Notes = $null
    }
}

Assert-GraphModuleLoaded
Ensure-GraphConnection

$groupNameCount = @($GroupNames).Length
if ($groupNameCount -eq 0) {
    throw "At least one group name is required."
}

$normalizedNames = @()
foreach ($name in $GroupNames) {
    if ($null -eq $name) {
        continue
    }

    $trimmed = $name.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        continue
    }

    if ($normalizedNames -notcontains $trimmed) {
        $normalizedNames += $trimmed
    }
}

$normalizedCount = @($normalizedNames).Length
if ($normalizedCount -eq 0) {
    throw "All supplied group names were empty."
}

$classification = @()
foreach ($groupName in $normalizedNames) {
    try {
        $matches = @(Get-GroupMatches -DisplayName $groupName)
    }
    catch {
        $classification += [PSCustomObject]@{
            RequestedName = $groupName
            ResolvedName = $null
            GroupId = $null
            Origin = "Error"
            OnPremisesSyncEnabled = $null
            OnPremisesSecurityIdentifier = $null
            GroupTypes = $null
            MatchIndex = 0
            TotalMatches = 0
            Notes = $_.Exception.Message
        }
        continue
    }

    if ($matches.Length -eq 0) {
        $classification += [PSCustomObject]@{
            RequestedName = $groupName
            ResolvedName = $null
            GroupId = $null
            Origin = "NotFound"
            OnPremisesSyncEnabled = $null
            OnPremisesSecurityIdentifier = $null
            GroupTypes = $null
            MatchIndex = 0
            TotalMatches = 0
            Notes = "Group not found in Microsoft Graph."
        }
        continue
    }

    $index = 0
    foreach ($match in $matches) {
        $index++
        $classification += Classify-Group -RequestedName $groupName -Group $match -MatchIndex $index -TotalMatches $matches.Length
    }
}

$classificationList = @($classification)

$columnOrder = @(
    "RequestedName",
    "ResolvedName",
    "GroupId",
    "Origin",
    "OnPremisesSyncEnabled",
    "OnPremisesSecurityIdentifier",
    "GroupTypes",
    "MatchIndex",
    "TotalMatches",
    "Notes"
)

$shouldWriteCsv = -not [string]::IsNullOrWhiteSpace($OutputPath)
$shouldWriteExcel = -not [string]::IsNullOrWhiteSpace($ExcelOutputPath)

if ($shouldWriteCsv) {
    $resolvedOutput = Resolve-OutputPath -PathValue $OutputPath
    $classificationList | Export-Csv -Path $resolvedOutput -NoTypeInformation -Encoding UTF8
    Write-Host "Saved classification for $($classificationList.Length) entries to '$resolvedOutput'." -ForegroundColor Green
}

if ($shouldWriteExcel) {
    $resolvedExcelOutput = Resolve-OutputPath -PathValue $ExcelOutputPath
    Write-XlsxFile -Path $resolvedExcelOutput -Rows $classificationList -Columns $columnOrder -WorksheetName "Groups"

    Write-Host "Saved Excel classification for $($classificationList.Length) entries to '$resolvedExcelOutput'." -ForegroundColor Green
}

if ($PassThru -or -not $shouldWriteCsv) {
    $classificationList
}
