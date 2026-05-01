#Requires -Version 5.1
<#
.SYNOPSIS
    Sets up mailbox forwarding to external contacts for users in a specified OU.

.DESCRIPTION
    Loops through mailboxes in the given Organizational Unit. For each mailbox,
    creates a hidden mail contact at the forwarding email address (if one does
    not already exist) and configures ForwardingAddress on the mailbox.

.PARAMETER OrganizationalUnit
    LDAP/canonical OU path to scope mailbox retrieval. REQUIRED.

.PARAMETER ForwardingDomain
    External domain appended to each user's SamAccountName to form the
    forwarding email address. REQUIRED.

.PARAMETER ContactsOU
    OU where the mail contacts will be created. Defaults to a "Contacts"
    sub-OU beneath OrganizationalUnit.

.PARAMETER DeliverToMailboxAndForward
    If $true (default), mail is delivered locally AND forwarded.
    Set to $false to forward-only (no local copy).

.PARAMETER WhatIf
    Preview changes without applying them.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\Forwarding.ps1 `
        -OrganizationalUnit "contoso.local/SITES/Branch" `
        -ForwardingDomain "@contoso365.org" `
        -ContactsOU "contoso.local/SITES/Branch/Contacts"

.NOTES
    Improved: removed hardcoded OU and domain, added parameters, error handling
    per-mailbox, WhatIf support, confirmation prompt, and logging.
    Original: hardcoded elant.local OU and @elantcare.org domain with no error handling.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$OrganizationalUnit,

    [Parameter(Mandatory)]
    [ValidatePattern('^@.+\..+$')]
    [string]$ForwardingDomain,

    [string]$ContactsOU,

    [bool]$DeliverToMailboxAndForward = $true,

    [switch]$Force
)

# Import shared utilities if available
$utilitiesPath = Join-Path $PSScriptRoot "PSO365-Utilities.psm1"
if (Test-Path $utilitiesPath) {
    Import-Module $utilitiesPath -Force
    Initialize-Log -ScriptName "Forwarding"
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

# Default contacts OU to a sub-OU of the source OU
if (-not $ContactsOU) {
    $ContactsOU = "$OrganizationalUnit/Office365Contacts"
}

$actionDesc = "Set forwarding to '$ForwardingDomain' for all mailboxes in OU '$OrganizationalUnit'."
if (-not $WhatIfPreference) {
    if (-not (Confirm-DestructiveAction -ActionDescription $actionDesc -Force:$Force)) {
        Write-Log "Operation cancelled by user." -Level INFO
        return
    }
}

Write-Log "Retrieving mailboxes from OU: $OrganizationalUnit" -Level INFO
try {
    $mailboxes = Get-Mailbox -ResultSize Unlimited `
        -OrganizationalUnit $OrganizationalUnit `
        -ErrorAction Stop
}
catch {
    Write-Log "Failed to retrieve mailboxes: $_" -Level ERROR
    throw
}
Write-Log "Found $($mailboxes.Count) mailbox(es)." -Level INFO

$successCount = 0
$failCount    = 0

foreach ($mailbox in $mailboxes) {
    $forwardingAddress = $mailbox.SamAccountName + $ForwardingDomain
    Write-Log "Processing: $($mailbox.DisplayName) -> $forwardingAddress" -Level INFO

    if ($PSCmdlet.ShouldProcess($mailbox.DisplayName, "Set forwarding to $forwardingAddress")) {
        # Create contact if it does not exist
        $existingContact = Get-MailContact $forwardingAddress -ErrorAction SilentlyContinue
        if (-not $existingContact) {
            try {
                New-MailContact `
                    -Name $mailbox.DisplayName `
                    -ExternalEmailAddress $forwardingAddress `
                    -OrganizationalUnit $ContactsOU `
                    -ErrorAction Stop |
                    Set-MailContact -HiddenFromAddressListsEnabled $true -ErrorAction Stop
                Write-Log "  Created hidden contact: $forwardingAddress" -Level SUCCESS
            }
            catch {
                Write-Log "  Failed to create contact for $($mailbox.DisplayName): $_" -Level ERROR
                $failCount++
                continue
            }
        }
        else {
            Write-Log "  Contact already exists: $forwardingAddress" -Level INFO
        }

        # Set forwarding on the mailbox
        try {
            Set-Mailbox `
                -Identity $mailbox.SamAccountName `
                -ForwardingAddress $forwardingAddress `
                -DeliverToMailboxAndForward $DeliverToMailboxAndForward `
                -ErrorAction Stop
            Write-Log "  Forwarding set on: $($mailbox.SamAccountName)" -Level SUCCESS
            $successCount++
        }
        catch {
            Write-Log "  Failed to set forwarding on $($mailbox.SamAccountName): $_" -Level ERROR
            $failCount++
        }
    }
}

Write-Log "Complete. Success: $successCount  |  Failed: $failCount" -Level INFO
