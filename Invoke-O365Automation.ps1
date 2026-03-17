#Requires -Version 5.1
<#
.SYNOPSIS
    Combined Office 365 / Exchange automation menu.

.DESCRIPTION
    Interactive (or parameter-driven) automation hub that wraps the most common
    PSforO365 administrative workflows into a single entry point:

    1  User Provisioning      - Bulk create AD/M365 users and assign licenses
    2  UPN Migration          - Bulk update User Principal Names from CSV
    3  License Assignment     - Assign O365 licenses to users from CSV
    4  Mailbox Permission Audit - Export Full Access / Send As / Send On Behalf
    5  Proxy Address Repair   - Find and fix mailboxes missing proxy addresses
    6  Bulk External Contacts - Import external mail contacts from CSV
    7  Disable IMAP/POP       - Harden mailboxes by disabling legacy protocols
    8  Mailbox Forwarding     - Set forwarding to external contacts by OU
    9  Admin Full Access Grant - Grant admin Full Access to all user mailboxes
    10 Exchange Health Check  - Run Test-ExchangeServerHealth report
    11 Exchange Environment   - Run Get-ExchangeEnvironmentReport
    0  Exit

    Pass -Task <number> to run non-interactively (e.g. from a scheduled task).

.PARAMETER Task
    Task number to run directly (skips the menu). Use 0 for all tasks.

.PARAMETER AdminUPN
    Admin UPN used for connection and privileged tasks.

.PARAMETER CsvPath
    Default CSV path passed to tasks that require one.

.PARAMETER OutputDirectory
    Directory for logs and exported reports. Defaults to .\O365Reports.

.PARAMETER LegacyMode
    Use legacy PSSession connection instead of the ExchangeOnlineManagement module.

.PARAMETER Force
    Bypass confirmation prompts on destructive operations.

.EXAMPLE
    .\Invoke-O365Automation.ps1
    # Interactive menu

.EXAMPLE
    .\Invoke-O365Automation.ps1 -Task 7 -Force
    # Disable IMAP/POP on all mailboxes without prompting

.EXAMPLE
    .\Invoke-O365Automation.ps1 -Task 4 -OutputDirectory C:\AuditReports
    # Run mailbox permission audit and save to specified directory

.NOTES
    Version: 1.0
    Requires: PSO365-Utilities.psm1 in the same directory as this script.
    Individual scripts for each workflow must also be present in the same directory.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateRange(0, 11)]
    [int]$Task = -1,

    [string]$AdminUPN,

    [string]$CsvPath,

    [string]$OutputDirectory = ".\O365Reports",

    [switch]$LegacyMode,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────
# Bootstrap shared utilities
# ─────────────────────────────────────────────
$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) { $scriptRoot = Split-Path $MyInvocation.MyCommand.Path }

$utilitiesPath = Join-Path $scriptRoot "PSO365-Utilities.psm1"
if (-not (Test-Path $utilitiesPath)) {
    Write-Error "PSO365-Utilities.psm1 not found at $utilitiesPath. Cannot continue."
    exit 1
}
Import-Module $utilitiesPath -Force

if (-not (Test-Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}
Initialize-Log -LogDirectory $OutputDirectory -ScriptName "Invoke-O365Automation"

# ─────────────────────────────────────────────
# Helper: resolve a path to a child script
# ─────────────────────────────────────────────
function Resolve-ChildScript {
    param([string]$FileName)
    $path = Join-Path $scriptRoot $FileName
    if (-not (Test-Path $path)) {
        Write-Log "Script not found: $path" -Level WARN
        return $null
    }
    return $path
}

# ─────────────────────────────────────────────
# Connection State
# ─────────────────────────────────────────────
$script:EXOConnected  = $false
$script:MSOLConnected = $false

function Ensure-EXOConnected {
    if (-not $script:EXOConnected) {
        Write-Log "Establishing Exchange Online connection..." -Level INFO
        $connectParams = @{ LegacyMode = $LegacyMode }
        if ($AdminUPN) { $connectParams['UserPrincipalName'] = $AdminUPN }
        Connect-O365 @connectParams
        $script:EXOConnected = $true
    }
}

function Ensure-MSOLConnected {
    if (-not $script:MSOLConnected) {
        Write-Log "Establishing MSOnline connection..." -Level INFO
        Connect-MSOnline
        $script:MSOLConnected = $true
    }
}

# ─────────────────────────────────────────────
# Task Implementations
# ─────────────────────────────────────────────

function Invoke-UserProvisioning {
    Write-Log "=== TASK: Bulk User Provisioning ===" -Level INFO
    $csvPath = if ($CsvPath) { $CsvPath } else {
        try { Get-CsvFilePath -Title "Select user provisioning CSV" }
        catch { Write-Log "No CSV selected." -Level WARN; return }
    }
    $script1 = Resolve-ChildScript "Bulk User Creation\NewUserPS.ps1"
    if (-not $script1) { Write-Log "Bulk user creation script not found." -Level ERROR; return }
    Ensure-MSOLConnected
    Write-Log "Running bulk user creation..." -Level INFO
    & $script1 -CsvPath $csvPath
    Write-Log "User provisioning complete." -Level SUCCESS
}

function Invoke-UPNMigration {
    Write-Log "=== TASK: UPN Migration ===" -Level INFO
    $csvPath = if ($CsvPath) { $CsvPath } else {
        try { Get-CsvFilePath -Title "Select UPN migration CSV (columns: UserPrincipalName, EmailAddress)" }
        catch { Write-Log "No CSV selected." -Level WARN; return }
    }
    $script1 = Resolve-ChildScript "UPNChange.PS1"
    if (-not $script1) { Write-Log "UPNChange.PS1 not found." -Level ERROR; return }
    & $script1 -CsvPath $csvPath
}

function Invoke-LicenseAssignment {
    Write-Log "=== TASK: License Assignment ===" -Level INFO
    Ensure-MSOLConnected
    $script1 = Resolve-ChildScript "licenses.ps1"
    if (-not $script1) { Write-Log "licenses.ps1 not found." -Level ERROR; return }
    $params = @{}
    if ($CsvPath)  { $params['CsvPath'] = $CsvPath }
    if ($Force)    { $params['Force'] = $true }
    & $script1 @params
}

function Invoke-MailboxPermissionAudit {
    Write-Log "=== TASK: Mailbox Permission Audit ===" -Level INFO
    Ensure-EXOConnected

    $outFile = Join-Path $OutputDirectory ("MailboxPermissions_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv")
    Write-Log "Output: $outFile" -Level INFO

    $header = "DisplayName,EmailAddress,FullAccess,SendAs,SendOnBehalfOf"
    $header | Out-File $outFile -Encoding UTF8 -Force

    Write-Log "Retrieving mailboxes..." -Level INFO
    $mailboxes = Get-Mailbox -ResultSize Unlimited |
        Select-Object Identity, Alias, DisplayName, DistinguishedName, WindowsEmailAddress

    $total = $mailboxes.Count
    $i     = 0
    foreach ($mbx in $mailboxes) {
        $i++
        Write-Progress -Activity "Auditing mailbox permissions" `
                       -Status "$i of $total: $($mbx.DisplayName)" `
                       -PercentComplete (($i / $total) * 100)
        try {
            $sendOnBehalf = (Get-Mailbox $mbx.Identity -ErrorAction Stop).GrantSendOnBehalfTo -join ";"
            $sendAs = (Get-ADPermission $mbx.Identity -ErrorAction SilentlyContinue |
                Where-Object { ($_.ExtendedRights -like "*Send-As*") -and
                               ($_.User -notlike "NT AUTHORITY\SELF") -and
                               ($_.User -notlike "S-1-5-21*") }).User -join ";"
            $fullAccess = (Get-MailboxPermission $mbx.Identity -ErrorAction Stop |
                Where-Object { (-not $_.IsInherited) -and
                               ($_.User -notmatch "NT AUTHORITY") -and
                               ($_.AccessRights -contains "FullAccess") }).User -join ";"

            "$($mbx.DisplayName),$($mbx.WindowsEmailAddress),$fullAccess,$sendAs,$sendOnBehalf" |
                Out-File $outFile -Append -Encoding UTF8
        }
        catch {
            Write-Log "Error auditing $($mbx.DisplayName): $_" -Level WARN
        }
    }
    Write-Progress -Activity "Auditing mailbox permissions" -Completed
    Write-Log "Audit complete. Report saved: $outFile" -Level SUCCESS
}

function Invoke-ProxyAddressRepair {
    Write-Log "=== TASK: Proxy Address Repair ===" -Level INFO
    Ensure-EXOConnected

    $domain = Read-Host "Enter the proxy domain (e.g. contoso.mail.onmicrosoft.com)"
    if (-not $domain) { Write-Log "No domain entered." -Level WARN; return }

    # Find mailboxes missing the proxy address
    $noProxyFile = Join-Path $OutputDirectory ("UsersNoProxy_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv")
    Write-Log "Scanning for mailboxes missing @$domain proxy address..." -Level INFO

    $noProxyUsers = Get-Mailbox -ResultSize Unlimited | Where-Object {
        $_.EmailAddresses -notmatch [regex]::Escape("@$domain")
    } | Select-Object DisplayName, Alias, PrimarySmtpAddress

    $noProxyUsers | Export-Csv $noProxyFile -NoTypeInformation
    Write-Log "Found $($noProxyUsers.Count) mailbox(es) without @$domain. Saved: $noProxyFile" -Level INFO

    if ($noProxyUsers.Count -eq 0) {
        Write-Log "No proxy address repairs needed." -Level SUCCESS
        return
    }

    $fix = Read-Host "Apply proxy addresses now? (YES to confirm)"
    if ($fix -ne 'YES') { Write-Log "Repair skipped." -Level INFO; return }

    $script1 = Resolve-ChildScript "Exchange-SetUserProxy.ps1"
    if (-not $script1) { Write-Log "Exchange-SetUserProxy.ps1 not found." -Level ERROR; return }
    & $script1 -CsvPath $noProxyFile -ProxyDomain $domain
}

function Invoke-BulkExternalContacts {
    Write-Log "=== TASK: Bulk External Contacts Import ===" -Level INFO
    Ensure-EXOConnected
    $csvPath = if ($CsvPath) { $CsvPath } else {
        try { Get-CsvFilePath -Title "Select external contacts CSV (columns: Name, ExternalEmailAddress, FirstName, LastName)" }
        catch { Write-Log "No CSV selected." -Level WARN; return }
    }
    $script1 = Resolve-ChildScript "BulkAddExternalContacts.ps1"
    if (-not $script1) { Write-Log "BulkAddExternalContacts.ps1 not found." -Level ERROR; return }
    & $script1 -CsvPath $csvPath
}

function Invoke-DisableIMAPPOP {
    Write-Log "=== TASK: Disable IMAP and POP3 ===" -Level INFO
    Ensure-EXOConnected
    $script1 = Resolve-ChildScript "DisableIMAPandPOP.ps1"
    if (-not $script1) { Write-Log "DisableIMAPandPOP.ps1 not found." -Level ERROR; return }
    $params = @{}
    if ($Force) { $params['Force'] = $true }
    & $script1 @params
}

function Invoke-MailboxForwarding {
    Write-Log "=== TASK: Mailbox Forwarding Setup ===" -Level INFO
    Ensure-EXOConnected
    $ou     = Read-Host "Organizational Unit (e.g. contoso.local/SITES/Branch)"
    $domain = Read-Host "Forwarding domain with @ prefix (e.g. @contoso365.org)"
    if (-not $ou -or -not $domain) { Write-Log "OU and domain are required." -Level WARN; return }

    $script1 = Resolve-ChildScript "Forwarding.ps1"
    if (-not $script1) { Write-Log "Forwarding.ps1 not found." -Level ERROR; return }
    $params = @{ OrganizationalUnit = $ou; ForwardingDomain = $domain }
    if ($Force) { $params['Force'] = $true }
    & $script1 @params
}

function Invoke-AdminFullAccess {
    Write-Log "=== TASK: Grant Admin Full Access to All Mailboxes ===" -Level INFO
    Ensure-EXOConnected
    $admin = if ($AdminUPN) { $AdminUPN } else { Read-Host "Enter admin UPN" }
    if (-not $admin) { Write-Log "Admin UPN required." -Level WARN; return }

    $script1 = Resolve-ChildScript "FullMailboxPermissionsforAdmin.ps1"
    if (-not $script1) { Write-Log "FullMailboxPermissionsforAdmin.ps1 not found." -Level ERROR; return }
    $params = @{ AdminUPN = $admin }
    if ($Force) { $params['Force'] = $true }
    & $script1 @params
}

function Invoke-ExchangeHealthCheck {
    Write-Log "=== TASK: Exchange Server Health Check ===" -Level INFO
    $script1 = Resolve-ChildScript "Test-ExchangeServerHealth.ps1"
    if (-not $script1) { Write-Log "Test-ExchangeServerHealth.ps1 not found." -Level ERROR; return }
    & $script1
}

function Invoke-ExchangeEnvironmentReport {
    Write-Log "=== TASK: Exchange Environment Report ===" -Level INFO
    $script1 = Resolve-ChildScript "Get-ExchangeEnvironmentReport.ps1"
    if (-not $script1) { Write-Log "Get-ExchangeEnvironmentReport.ps1 not found." -Level ERROR; return }
    & $script1
}

# ─────────────────────────────────────────────
# Task Dispatch Table
# ─────────────────────────────────────────────
$tasks = @{
    1  = @{ Name = "User Provisioning";            Fn = { Invoke-UserProvisioning } }
    2  = @{ Name = "UPN Migration";                Fn = { Invoke-UPNMigration } }
    3  = @{ Name = "License Assignment";           Fn = { Invoke-LicenseAssignment } }
    4  = @{ Name = "Mailbox Permission Audit";     Fn = { Invoke-MailboxPermissionAudit } }
    5  = @{ Name = "Proxy Address Repair";         Fn = { Invoke-ProxyAddressRepair } }
    6  = @{ Name = "Bulk External Contacts";       Fn = { Invoke-BulkExternalContacts } }
    7  = @{ Name = "Disable IMAP/POP";             Fn = { Invoke-DisableIMAPPOP } }
    8  = @{ Name = "Mailbox Forwarding Setup";     Fn = { Invoke-MailboxForwarding } }
    9  = @{ Name = "Admin Full Access Grant";      Fn = { Invoke-AdminFullAccess } }
    10 = @{ Name = "Exchange Health Check";        Fn = { Invoke-ExchangeHealthCheck } }
    11 = @{ Name = "Exchange Environment Report";  Fn = { Invoke-ExchangeEnvironmentReport } }
}

function Show-Menu {
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║       PSforO365 Automation Hub v1.0          ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    foreach ($key in ($tasks.Keys | Sort-Object)) {
        Write-Host ("   {0,2}  {1}" -f $key, $tasks[$key].Name)
    }
    Write-Host "    0  Exit"
    Write-Host ""
}

function Invoke-Task {
    param([int]$TaskNumber)
    if ($tasks.ContainsKey($TaskNumber)) {
        Write-Log "Starting task $TaskNumber : $($tasks[$TaskNumber].Name)" -Level INFO
        try {
            & $tasks[$TaskNumber].Fn
        }
        catch {
            Write-Log "Task $TaskNumber failed: $_" -Level ERROR
        }
    }
    else {
        Write-Log "Unknown task number: $TaskNumber" -Level WARN
    }
}

# ─────────────────────────────────────────────
# Entry Point
# ─────────────────────────────────────────────
if ($Task -ge 1) {
    # Non-interactive: run the specified task directly
    Invoke-Task -TaskNumber $Task
}
elseif ($Task -eq 0) {
    # Run ALL tasks in sequence
    Write-Log "Running all tasks in sequence..." -Level WARN
    foreach ($key in ($tasks.Keys | Sort-Object)) {
        Invoke-Task -TaskNumber $key
    }
}
else {
    # Interactive menu loop
    do {
        Show-Menu
        $selection = Read-Host "  Select a task (0 to exit)"
        if ($selection -match '^\d+$') {
            $num = [int]$selection
            if ($num -eq 0) { break }
            Invoke-Task -TaskNumber $num
        }
        else {
            Write-Host "  Invalid selection. Enter a number from the menu." -ForegroundColor Yellow
        }
        Write-Host ""
        Read-Host "  Press Enter to return to the menu..."
    } while ($true)

    Write-Log "Exiting PSforO365 Automation Hub." -Level INFO
}
