#Requires -Version 5.1
<#
.SYNOPSIS
    Adds a proxy email address to Exchange mailboxes from a CSV file.

.DESCRIPTION
    Reads a CSV containing PrimarySmtpAddress and Alias columns, then adds
    a proxy address in the format <Alias>@<ProxyDomain> to each mailbox.

.PARAMETER CsvPath
    Path to the CSV file. Must contain 'PrimarySmtpAddress' and 'Alias' columns.
    Defaults to .\UsersNoProxy.csv

.PARAMETER ProxyDomain
    The mail domain to use when building the proxy address
    (e.g. "contoso.mail.onmicrosoft.com"). REQUIRED.

.PARAMETER WhatIf
    Preview changes without applying them.

.EXAMPLE
    .\Exchange-SetUserProxy.ps1 `
        -CsvPath C:\Data\UsersNoProxy.csv `
        -ProxyDomain "contoso.mail.onmicrosoft.com"

.NOTES
    Improved: removed hardcoded CSV path and domain, added parameters, validation,
    per-mailbox error handling, and logging.
    Original: hardcoded C:\csv\UsersNoProxy.csv and contoso.mail.onmicrosoft.com.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$CsvPath = ".\UsersNoProxy.csv",

    [Parameter(Mandatory)]
    [string]$ProxyDomain
)

# Import shared utilities if available
$utilitiesPath = Join-Path $PSScriptRoot "PSO365-Utilities.psm1"
if (Test-Path $utilitiesPath) {
    Import-Module $utilitiesPath -Force
    Initialize-Log -ScriptName "Exchange-SetUserProxy"
}
else {
    function Write-Log { param([string]$Message, [string]$Level = 'INFO') Write-Host "[$Level] $Message" }
}

if (-not (Test-Path $CsvPath)) {
    Write-Log "CSV file not found: $CsvPath" -Level ERROR
    throw "CSV not found: $CsvPath"
}

$users = Import-Csv -Path $CsvPath
if (-not $users) {
    Write-Log "CSV is empty: $CsvPath" -Level ERROR
    throw "CSV is empty."
}

$required = @('PrimarySmtpAddress', 'Alias')
foreach ($col in $required) {
    if ($col -notin $users[0].PSObject.Properties.Name) {
        Write-Log "CSV missing required column: '$col'" -Level ERROR
        throw "Missing column: $col"
    }
}

Write-Log "CSV loaded: $CsvPath ($($users.Count) rows). Proxy domain: $ProxyDomain" -Level INFO

$successCount = 0
$failCount    = 0

foreach ($user in $users) {
    $id       = $user.PrimarySmtpAddress
    $alias    = $user.Alias
    $newProxy = "$alias@$ProxyDomain"

    if (-not $id -or -not $alias) {
        Write-Log "Skipping row with missing PrimarySmtpAddress or Alias." -Level WARN
        continue
    }

    if ($PSCmdlet.ShouldProcess($id, "Add proxy address $newProxy")) {
        try {
            Set-Mailbox -Identity $id `
                -EmailAddresses @{ add = $newProxy } `
                -ErrorAction Stop
            Write-Log "Added $newProxy to: $id" -Level SUCCESS
            $successCount++
        }
        catch {
            Write-Log "FAILED on $id: $_" -Level ERROR
            $failCount++
        }
    }
}

Write-Log "Complete. Success: $successCount  |  Failed: $failCount" -Level INFO
