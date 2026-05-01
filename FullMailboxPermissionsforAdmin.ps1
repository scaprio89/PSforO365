#Requires -Version 5.1
<#
.SYNOPSIS
    Grants Full Access mailbox permissions to a specified admin account.

.DESCRIPTION
    Iterates all user mailboxes and adds Full Access for the given admin UPN.
    Supports -WhatIf for dry-run previews and requires explicit confirmation
    before making changes.

.PARAMETER AdminUPN
    The admin account UPN to grant Full Access to. REQUIRED - no default.

.PARAMETER WhatIf
    Preview the changes without applying them.

.PARAMETER Force
    Skip the confirmation prompt (use in automation pipelines).

.EXAMPLE
    .\FullMailboxPermissionsforAdmin.ps1 -AdminUPN admin@contoso.onmicrosoft.com

.EXAMPLE
    .\FullMailboxPermissionsforAdmin.ps1 -AdminUPN admin@contoso.onmicrosoft.com -WhatIf

.NOTES
    Improved: removed hardcoded UPN, added -WhatIf support, confirmation prompt,
    error handling per mailbox, and logging.
    Original: single one-liner with hardcoded org-specific UPN.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$AdminUPN,

    [switch]$Force
)

# Import shared utilities if available
$utilitiesPath = Join-Path $PSScriptRoot "PSO365-Utilities.psm1"
if (Test-Path $utilitiesPath) {
    Import-Module $utilitiesPath -Force
    Initialize-Log -ScriptName "FullMailboxPermissionsforAdmin"
}
else {
    function Write-Log { param([string]$Message, [string]$Level = 'INFO') Write-Host "[$Level] $Message" }
    function Confirm-DestructiveAction {
        param([string]$ActionDescription, [switch]$Force)
        if ($Force) { return $true }
        Write-Host "`n  *** WARNING *** $ActionDescription" -ForegroundColor Yellow
        return (Read-Host "Type 'YES' to confirm") -eq 'YES'
    }
}

# Confirm before proceeding
$actionDesc = "Grant '$AdminUPN' Full Access to ALL user mailboxes in the organization."
if (-not $WhatIfPreference) {
    if (-not (Confirm-DestructiveAction -ActionDescription $actionDesc -Force:$Force)) {
        Write-Log "Operation cancelled by user." -Level INFO
        return
    }
}

Write-Log "Retrieving all user mailboxes..." -Level INFO
try {
    $mailboxes = Get-Mailbox -ResultSize Unlimited `
        -Filter "RecipientTypeDetails -eq 'UserMailbox'" `
        -ErrorAction Stop
}
catch {
    Write-Log "Failed to retrieve mailboxes: $_" -Level ERROR
    throw
}
Write-Log "Found $($mailboxes.Count) user mailbox(es)." -Level INFO

$successCount = 0
$failCount    = 0

foreach ($mbx in $mailboxes) {
    if ($PSCmdlet.ShouldProcess($mbx.UserPrincipalName, "Add-MailboxPermission FullAccess for $AdminUPN")) {
        try {
            Add-MailboxPermission `
                -Identity $mbx.UserPrincipalName `
                -User $AdminUPN `
                -AccessRights FullAccess `
                -InheritanceType All `
                -ErrorAction Stop | Out-Null
            Write-Log "Granted FullAccess on: $($mbx.UserPrincipalName)" -Level SUCCESS
            $successCount++
        }
        catch {
            Write-Log "FAILED on $($mbx.UserPrincipalName): $_" -Level ERROR
            $failCount++
        }
    }
}

Write-Log "Complete. Success: $successCount  |  Failed: $failCount" -Level INFO
