#Requires -Version 5.1
<#
.SYNOPSIS
    Disables IMAP and POP3 access on Exchange/EXO mailboxes.

.DESCRIPTION
    Sets ImapEnabled and PopEnabled to $false on all user mailboxes (default)
    or a scoped subset. Supports -WhatIf and requires confirmation before
    bulk changes.

.PARAMETER Identity
    Specific mailbox identity to target. If omitted, all mailboxes are targeted.

.PARAMETER WhatIf
    Preview the changes without applying them.

.PARAMETER Force
    Bypass the confirmation prompt.

.EXAMPLE
    .\DisableIMAPandPOP.ps1
    # Disables on ALL mailboxes (prompts for confirmation)

.EXAMPLE
    .\DisableIMAPandPOP.ps1 -Identity user@contoso.com

.EXAMPLE
    .\DisableIMAPandPOP.ps1 -WhatIf
    # Preview without changes

.NOTES
    Improved: added parameters, bulk confirmation prompt, per-mailbox error
    handling, result summary, and logging.
    Original: one-liner with no error handling, no confirmation, no logging.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Identity,
    [switch]$Force
)

# Import shared utilities if available
$utilitiesPath = Join-Path $PSScriptRoot "PSO365-Utilities.psm1"
if (Test-Path $utilitiesPath) {
    Import-Module $utilitiesPath -Force
    Initialize-Log -ScriptName "DisableIMAPandPOP"
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

# Retrieve target mailboxes
Write-Log "Retrieving mailboxes..." -Level INFO
try {
    if ($Identity) {
        $mailboxes = @(Get-Mailbox -Identity $Identity -ErrorAction Stop)
    }
    else {
        $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
    }
}
catch {
    Write-Log "Failed to retrieve mailboxes: $_" -Level ERROR
    throw
}
Write-Log "Found $($mailboxes.Count) mailbox(es)." -Level INFO

# Confirm bulk operation
if (-not $Identity -and -not $WhatIfPreference) {
    $actionDesc = "Disable IMAP and POP3 on ALL $($mailboxes.Count) mailbox(es)."
    if (-not (Confirm-DestructiveAction -ActionDescription $actionDesc -Force:$Force)) {
        Write-Log "Operation cancelled by user." -Level INFO
        return
    }
}

$successCount = 0
$failCount    = 0

foreach ($mbx in $mailboxes) {
    $id = $mbx.UserPrincipalName ?? $mbx.Identity
    if ($PSCmdlet.ShouldProcess($id, "Set-CASMailbox -ImapEnabled `$false -PopEnabled `$false")) {
        try {
            Set-CASMailbox -Identity $mbx.Identity `
                -ImapEnabled $false `
                -PopEnabled $false `
                -ErrorAction Stop
            Write-Log "IMAP/POP disabled: $id" -Level SUCCESS
            $successCount++
        }
        catch {
            Write-Log "FAILED on $id: $_" -Level ERROR
            $failCount++
        }
    }
}

Write-Log "Complete. Success: $successCount  |  Failed: $failCount" -Level INFO
