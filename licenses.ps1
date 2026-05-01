#Requires -Version 5.1
<#
.SYNOPSIS
    Assigns an Office 365 license to users listed in a CSV file.

.DESCRIPTION
    Reads a CSV of user UPNs, sets the usage location on each account, then
    applies the specified license SKU. Displays available SKUs in a grid view
    to help identify the correct AccountSkuId.

.PARAMETER CsvPath
    Path to the CSV file. Must contain a 'UserPrincipalName' column.
    If omitted, a file-picker dialog will open.

.PARAMETER LicenseSku
    The AccountSkuId of the license to assign (e.g. "contoso:ENTERPRISEPACK").
    If omitted, available SKUs are shown and you will be prompted.

.PARAMETER UsageLocation
    Two-letter ISO country code for the usage location (e.g. "US", "GB", "CA").
    Defaults to "US". REQUIRED by M365 before assigning a license.

.EXAMPLE
    .\licenses.ps1 -CsvPath C:\users.csv -LicenseSku "contoso:ENTERPRISEPACK" -UsageLocation US

.EXAMPLE
    .\licenses.ps1
    # Interactive: opens file picker and shows SKU grid

.NOTES
    Improved: replaced hardcoded usage location and path, added parameters,
    CSV validation, per-user error handling, result summary, and logging.
    Original: hardcoded 'US' usage location, no error handling, no logging.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$CsvPath,

    [string]$LicenseSku,

    [ValidateLength(2, 2)]
    [string]$UsageLocation = "US"
)

# Import shared utilities if available
$utilitiesPath = Join-Path $PSScriptRoot "PSO365-Utilities.psm1"
if (Test-Path $utilitiesPath) {
    Import-Module $utilitiesPath -Force
    Initialize-Log -ScriptName "AssignLicenses"
}
else {
    function Write-Log { param([string]$Message, [string]$Level = 'INFO') Write-Host "[$Level] $Message" }
}

# ---- Connect to MSOnline ----
Write-Log "Connecting to Microsoft Online Service..." -Level INFO
try {
    Connect-MsolService -ErrorAction Stop
    Write-Log "Connected." -Level SUCCESS
}
catch {
    Write-Log "Failed to connect to MSOnline: $_" -Level ERROR
    throw
}

# ---- Select CSV ----
if (-not $CsvPath) {
    Write-Log "No CSV path provided - opening file picker." -Level INFO
    try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title            = "Select user CSV file"
        $dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        $dialog.Filter           = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Log "No file selected. Exiting." -Level WARN
            return
        }
        $CsvPath = $dialog.FileName
    }
    catch {
        Write-Log "File picker failed: $_" -Level ERROR
        $CsvPath = Read-Host "Enter the full path to the CSV file"
    }
}

if (-not (Test-Path $CsvPath)) {
    Write-Log "CSV file not found: $CsvPath" -Level ERROR
    throw "CSV file not found: $CsvPath"
}

$users = Import-Csv -Path $CsvPath
if (-not $users) {
    Write-Log "CSV file is empty: $CsvPath" -Level ERROR
    throw "CSV file is empty."
}
if ('UserPrincipalName' -notin $users[0].PSObject.Properties.Name) {
    Write-Log "CSV must contain a 'UserPrincipalName' column." -Level ERROR
    throw "Missing required CSV column: UserPrincipalName"
}
Write-Log "CSV loaded: $CsvPath ($($users.Count) rows)" -Level INFO

# ---- Select License SKU ----
if (-not $LicenseSku) {
    Write-Log "Available license SKUs:" -Level INFO
    $skus = Get-MsolAccountSku
    $selected = $skus | Out-GridView -Title "Select a license SKU" -PassThru
    if (-not $selected) {
        Write-Log "No SKU selected. Exiting." -Level WARN
        return
    }
    $LicenseSku = $selected.AccountSkuId
}
Write-Log "License SKU: $LicenseSku | Usage Location: $UsageLocation" -Level INFO

# ---- Assign licenses ----
$successCount = 0
$failCount    = 0
$skipCount    = 0

foreach ($user in $users) {
    $upn = $user.UserPrincipalName
    if (-not $upn) {
        Write-Log "Skipping row with empty UserPrincipalName." -Level WARN
        $skipCount++
        continue
    }

    if ($PSCmdlet.ShouldProcess($upn, "Set usage location '$UsageLocation' and assign license '$LicenseSku'")) {
        try {
            Set-MsolUser -UserPrincipalName $upn -UsageLocation $UsageLocation -ErrorAction Stop
            Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $LicenseSku -ErrorAction Stop
            Write-Log "Licensed: $upn" -Level SUCCESS
            $successCount++
        }
        catch {
            Write-Log "FAILED: $upn - $_" -Level ERROR
            $failCount++
        }
    }
}

Write-Log "Complete. Success: $successCount  |  Failed: $failCount  |  Skipped: $skipCount" -Level INFO

# Show results grid
Write-Log "Loading result view..." -Level INFO
Import-Csv -Path $CsvPath | ForEach-Object {
    Get-MsolUser -UserPrincipalName $_.UserPrincipalName -ErrorAction SilentlyContinue
} | Out-GridView -Title "License Assignment Results"
