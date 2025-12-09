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
    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath "graph-group-origin.csv"),
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

if ($OutputPath) {
    $resolvedOutput = $OutputPath
    if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
        $resolvedOutput = Join-Path -Path (Get-Location) -ChildPath $OutputPath
    }

    $classificationList | Export-Csv -Path $resolvedOutput -NoTypeInformation -Encoding UTF8
    Write-Host "Saved classification for $($classificationList.Length) entries to '$resolvedOutput'." -ForegroundColor Green
}

if ($PassThru -or -not $OutputPath) {
    $classificationList
}
