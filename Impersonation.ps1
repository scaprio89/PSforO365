#Requires -Version 5.1
<#
.SYNOPSIS
    Assigns (or removes) the ApplicationImpersonation management role to a user.

.DESCRIPTION
    Grants the ApplicationImpersonation RBAC role assignment, which allows a
    service account to impersonate any mailbox. Requires confirmation before
    applying because this is a high-privilege operation.

.PARAMETER UserUPN
    UPN of the account to assign the role to. REQUIRED - no default.

.PARAMETER Remove
    Remove the role assignment instead of adding it.

.PARAMETER Force
    Bypass the confirmation prompt (for automation pipelines).

.EXAMPLE
    .\Impersonation.ps1 -UserUPN serviceaccount@contoso.onmicrosoft.com

.EXAMPLE
    .\Impersonation.ps1 -UserUPN serviceaccount@contoso.onmicrosoft.com -Remove

.NOTES
    Improved: removed hardcoded UPN, added -Remove switch, confirmation prompt,
    error handling, and logging.
    Original: single New-ManagementRoleAssignment one-liner with hardcoded UPN.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$UserUPN,

    [switch]$Remove,
    [switch]$Force
)

# Import shared utilities if available
$utilitiesPath = Join-Path $PSScriptRoot "PSO365-Utilities.psm1"
if (Test-Path $utilitiesPath) {
    Import-Module $utilitiesPath -Force
    Initialize-Log -ScriptName "Impersonation"
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

$roleName = "ApplicationImpersonation"
$operation = if ($Remove) { "REMOVE" } else { "ASSIGN" }
$actionDesc = "$operation the '$roleName' role to/from '$UserUPN'. This is a high-privilege operation."

if (-not $WhatIfPreference) {
    if (-not (Confirm-DestructiveAction -ActionDescription $actionDesc -Force:$Force)) {
        Write-Log "Operation cancelled by user." -Level INFO
        return
    }
}

if ($Remove) {
    Write-Log "Removing '$roleName' from '$UserUPN'..." -Level INFO
    try {
        if ($PSCmdlet.ShouldProcess($UserUPN, "Remove-ManagementRoleAssignment $roleName")) {
            Get-ManagementRoleAssignment -Role $roleName -RoleAssignee $UserUPN -ErrorAction Stop |
                Remove-ManagementRoleAssignment -Confirm:$false -ErrorAction Stop
            Write-Log "Role '$roleName' removed from '$UserUPN'." -Level SUCCESS
        }
    }
    catch {
        Write-Log "Failed to remove role: $_" -Level ERROR
        throw
    }
}
else {
    Write-Log "Assigning '$roleName' to '$UserUPN'..." -Level INFO
    try {
        if ($PSCmdlet.ShouldProcess($UserUPN, "New-ManagementRoleAssignment $roleName")) {
            New-ManagementRoleAssignment -Role $roleName -User $UserUPN -ErrorAction Stop | Out-Null
            Write-Log "Role '$roleName' assigned to '$UserUPN'." -Level SUCCESS
        }
    }
    catch {
        Write-Log "Failed to assign role: $_" -Level ERROR
        throw
    }
}
