#Requires -Version 5.1
<#
.SYNOPSIS
    Bulk changes Active Directory User Principal Names from a CSV file.

.DESCRIPTION
    Reads a CSV containing current UPNs (UserPrincipalName) and the desired new
    UPNs (EmailAddress) and updates each AD user object. Requires the Active
    Directory module. Skips users not found in AD and reports failures.

.PARAMETER CsvPath
    Path to the CSV file. Must contain 'UserPrincipalName' (current) and
    'EmailAddress' (new UPN) columns. REQUIRED.

.PARAMETER WhatIf
    Preview changes without applying them.

.EXAMPLE
    .\UPNChange.ps1 -CsvPath C:\Data\users.csv

.NOTES
    Improved: replaced hardcoded path, added parameters, AD module check,
    per-user error handling, not-found detection, result summary, and logging.
    Original: hardcoded C:\FiestaLogonCSVTest.csv with no error handling.
    Author: Antonio Vargas (original v1.0, 06/30/2015)
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$CsvPath
)

# Import shared utilities if available
$utilitiesPath = Join-Path $PSScriptRoot "PSO365-Utilities.psm1"
if (Test-Path $utilitiesPath) {
    Import-Module $utilitiesPath -Force
    Initialize-Log -ScriptName "UPNChange"
}
else {
    function Write-Log { param([string]$Message, [string]$Level = 'INFO') Write-Host "[$Level] $Message" }
}

# Validate AD module
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Log "ActiveDirectory module not found. Install RSAT or run on a domain controller." -Level ERROR
    throw "Missing module: ActiveDirectory"
}
Import-Module ActiveDirectory -ErrorAction Stop

# Load and validate CSV
if (-not (Test-Path $CsvPath)) {
    Write-Log "CSV not found: $CsvPath" -Level ERROR
    throw "CSV not found: $CsvPath"
}

$users = Import-Csv -Path $CsvPath
if (-not $users) {
    Write-Log "CSV is empty: $CsvPath" -Level ERROR
    throw "CSV is empty."
}

$required = @('UserPrincipalName', 'EmailAddress')
foreach ($col in $required) {
    if ($col -notin $users[0].PSObject.Properties.Name) {
        Write-Log "CSV missing required column: '$col'" -Level ERROR
        throw "Missing column: $col"
    }
}

Write-Log "CSV loaded: $CsvPath ($($users.Count) rows)" -Level INFO

$successCount  = 0
$failCount     = 0
$notFoundCount = 0

foreach ($user in $users) {
    $currentUPN = $user.UserPrincipalName
    $newUPN     = $user.EmailAddress

    if (-not $currentUPN -or -not $newUPN) {
        Write-Log "Skipping row with missing UPN or EmailAddress." -Level WARN
        $failCount++
        continue
    }

    Write-Log "Processing: $currentUPN -> $newUPN" -Level INFO

    try {
        $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$currentUPN'" -ErrorAction Stop
    }
    catch {
        Write-Log "AD query failed for $currentUPN: $_" -Level ERROR
        $failCount++
        continue
    }

    if (-not $adUser) {
        Write-Log "User not found in AD: $currentUPN" -Level WARN
        $notFoundCount++
        continue
    }

    if ($PSCmdlet.ShouldProcess($currentUPN, "Set UPN to $newUPN")) {
        try {
            Set-ADUser -Identity $adUser -UserPrincipalName $newUPN -ErrorAction Stop
            Write-Log "  UPN updated: $currentUPN -> $newUPN" -Level SUCCESS
            $successCount++
        }
        catch {
            Write-Log "  FAILED: $currentUPN - $_" -Level ERROR
            $failCount++
        }
    }
}

Write-Log "Complete. Updated: $successCount  |  Not Found: $notFoundCount  |  Failed: $failCount" -Level INFO
