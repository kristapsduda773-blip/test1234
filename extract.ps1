<#
.SYNOPSIS
    Retrieve Microsoft Entra ID (Azure AD) groups and label them as OnPrem or Cloud.

.DESCRIPTION
    Looks up each provided display name in Microsoft Graph, determines whether the
    group originates from on-premises AD (DirSync) or is cloud-only, and then writes
    a minimal two-column Excel workbook (GroupName, Origin).

.PARAMETER GroupNames
    List of display names to evaluate. Defaults to the list supplied by the user.

.PARAMETER ExcelOutputPath
    Target .xlsx file path. Defaults to ./graph-group-origin.xlsx in the current directory.

.PARAMETER SkipLogin
    Use this switch when the current session is already connected to Microsoft Graph.

.EXAMPLE
    # Standard run
    ./extract.ps1

.EXAMPLE
    # Custom list and output path, skipping login because Connect-MgGraph was already called
    ./extract.ps1 -GroupNames "DEV-ATD","SQL-00-PBI-Sync" -ExcelOutputPath ./custom.xlsx -SkipLogin
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
    [string]$ExcelOutputPath = (Join-Path -Path (Get-Location) -ChildPath "graph-group-origin.xlsx"),
    [switch]$SkipLogin
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Import-RequiredModule {
    param([Parameter(Mandatory = $true)][string]$ModuleName)

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        throw "$ModuleName module is required. Install it via 'Install-Module $ModuleName -Scope CurrentUser'."
    }

    Import-Module $ModuleName -ErrorAction Stop | Out-Null
}

Import-RequiredModule -ModuleName Microsoft.Graph.Authentication
Import-RequiredModule -ModuleName Microsoft.Graph.Groups
Import-RequiredModule -ModuleName ImportExcel

function Ensure-GraphConnection {
    if ($SkipLogin) {
        return
    }

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context -or -not $context.Account) {
        Connect-MgGraph -Scopes "Group.Read.All" | Out-Null
    }
}

Ensure-GraphConnection

$cleanNames = @()
foreach ($name in $GroupNames) {
    if ([string]::IsNullOrWhiteSpace($name)) {
        continue
    }

    $trimmed = $name.Trim()
    if ($cleanNames -notcontains $trimmed) {
        $cleanNames += $trimmed
    }
}

if ($cleanNames.Count -eq 0) {
    throw "Provide at least one group name."
}

function Get-GroupOrigin {
    param([Parameter(Mandatory = $true)][string]$DisplayName)

    $escaped = $DisplayName -replace "'", "''"
    $filter = "displayName eq '$escaped'"

    $group = Get-MgGroup -Filter $filter -Property @(
        "displayName",
        "onPremisesSyncEnabled",
        "onPremisesSecurityIdentifier"
    ) -ConsistencyLevel "eventual" -ErrorAction Stop | Select-Object -First 1

    if ($null -eq $group) {
        return [PSCustomObject]@{ GroupName = $DisplayName; Origin = "NotFound" }
    }

    $origin = if (($group.OnPremisesSyncEnabled -eq $true) -or (-not [string]::IsNullOrWhiteSpace($group.OnPremisesSecurityIdentifier))) {
        "OnPrem"
    }
    else {
        "Cloud"
    }

    return [PSCustomObject]@{ GroupName = $group.DisplayName; Origin = $origin }
}

$results = foreach ($groupName in $cleanNames) {
    try {
        Get-GroupOrigin -DisplayName $groupName
    }
    catch {
        [PSCustomObject]@{
            GroupName = $groupName
            Origin = "Error: $($_.Exception.Message)"
        }
    }
}

if ($results.Count -eq 0) {
    Write-Warning "No results to export."
    return
}

$resolvedExcelPath = if ([System.IO.Path]::IsPathRooted($ExcelOutputPath)) {
    $ExcelOutputPath
}
else {
    Join-Path -Path (Get-Location) -ChildPath $ExcelOutputPath
}

$results | Export-Excel -Path $resolvedExcelPath -WorksheetName "Groups" -TableName "GroupOrigin" -BoldTopRow -AutoSize -FreezeTopRow
Write-Host "Saved $($results.Count) rows to '$resolvedExcelPath'." -ForegroundColor Green

$results
