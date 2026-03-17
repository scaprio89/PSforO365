#Requires -Version 5.1
<#
.SYNOPSIS
    Bulk imports external mail contacts from a CSV file.

.DESCRIPTION
    Reads a CSV and creates Exchange mail contacts for each row. Skips contacts
    that already exist. Reports successes, failures, and skips.

.PARAMETER CsvPath
    Path to the CSV file. Required columns: Name, ExternalEmailAddress,
    FirstName, LastName. Defaults to .\ExternalContacts.csv

.PARAMETER WhatIf
    Preview changes without applying them.

.EXAMPLE
    .\BulkAddExternalContacts.ps1 -CsvPath C:\Data\contacts.csv

.NOTES
    Improved: added parameters, CSV validation, duplicate detection, per-contact
    error handling, result summary, and logging.
    Original: one-liner with no validation, error handling, or logging.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$CsvPath = ".\ExternalContacts.csv"
)

# Import shared utilities if available
$utilitiesPath = Join-Path $PSScriptRoot "PSO365-Utilities.psm1"
if (Test-Path $utilitiesPath) {
    Import-Module $utilitiesPath -Force
    Initialize-Log -ScriptName "BulkAddExternalContacts"
}
else {
    function Write-Log { param([string]$Message, [string]$Level = 'INFO') Write-Host "[$Level] $Message" }
}

if (-not (Test-Path $CsvPath)) {
    Write-Log "CSV file not found: $CsvPath" -Level ERROR
    throw "CSV not found: $CsvPath"
}

$contacts = Import-Csv -Path $CsvPath
if (-not $contacts) {
    Write-Log "CSV is empty: $CsvPath" -Level ERROR
    throw "CSV is empty."
}

$required = @('Name', 'ExternalEmailAddress', 'FirstName', 'LastName')
foreach ($col in $required) {
    if ($col -notin $contacts[0].PSObject.Properties.Name) {
        Write-Log "CSV missing required column: '$col'" -Level ERROR
        throw "Missing column: $col"
    }
}

Write-Log "CSV loaded: $CsvPath ($($contacts.Count) rows)" -Level INFO

$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($contact in $contacts) {
    $name  = $contact.Name
    $email = $contact.ExternalEmailAddress

    if (-not $name -or -not $email) {
        Write-Log "Skipping row with missing Name or ExternalEmailAddress." -Level WARN
        $skipCount++
        continue
    }

    # Skip if contact already exists
    $existing = Get-MailContact -Identity $email -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Log "Already exists, skipping: $email" -Level WARN
        $skipCount++
        continue
    }

    if ($PSCmdlet.ShouldProcess($email, "New-MailContact '$name'")) {
        try {
            New-MailContact `
                -Name $name `
                -DisplayName $name `
                -ExternalEmailAddress $email `
                -FirstName $contact.FirstName `
                -LastName $contact.LastName `
                -ErrorAction Stop | Out-Null
            Write-Log "Created contact: $name <$email>" -Level SUCCESS
            $successCount++
        }
        catch {
            Write-Log "FAILED: $name <$email> - $_" -Level ERROR
            $failCount++
        }
    }
}

Write-Log "Complete. Created: $successCount  |  Skipped: $skipCount  |  Failed: $failCount" -Level INFO
