#Requires -Version 5.1
<#
.SYNOPSIS
    Connects to Exchange Online using modern or legacy authentication.

.DESCRIPTION
    Establishes an Exchange Online PowerShell session. Prefers the modern
    ExchangeOnlineManagement module (V2/V3). Use -LegacyMode for older
    tenants that still rely on basic-auth remote PowerShell sessions.

.PARAMETER UserPrincipalName
    Admin UPN used for authentication. If omitted, you will be prompted.

.PARAMETER LegacyMode
    Use the old New-PSSession / basic-auth method instead of the EXO V2/V3 module.

.EXAMPLE
    .\Connectto365.ps1 -UserPrincipalName admin@contoso.onmicrosoft.com

.EXAMPLE
    .\Connectto365.ps1 -LegacyMode

.NOTES
    Improved: added parameter support, module validation, and error handling.
    Original: single New-PSSession call with no error handling.
#>
[CmdletBinding()]
param(
    [string]$UserPrincipalName,
    [switch]$LegacyMode
)

# Import shared utilities if available
$utilitiesPath = Join-Path $PSScriptRoot "PSO365-Utilities.psm1"
if (Test-Path $utilitiesPath) {
    Import-Module $utilitiesPath -Force
    Initialize-Log -ScriptName "Connectto365"
}
else {
    function Write-Log { param([string]$Message, [string]$Level = 'INFO') Write-Host "[$Level] $Message" }
}

if ($LegacyMode) {
    Write-Log "Using legacy PSSession mode." -Level WARN
    try {
        $cred = Get-Credential -Message "Enter Office 365 admin credentials" -UserName $UserPrincipalName
        $session = New-PSSession `
            -ConfigurationName Microsoft.Exchange `
            -ConnectionUri https://ps.outlook.com/powershell/ `
            -Credential $cred `
            -Authentication Basic `
            -AllowRedirection `
            -ErrorAction Stop
        Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
        Write-Log "Exchange Online session established (legacy)." -Level SUCCESS
    }
    catch {
        Write-Log "Connection failed: $_" -Level ERROR
        throw
    }
    return
}

# Modern path: ExchangeOnlineManagement module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Log "ExchangeOnlineManagement module not found." -Level ERROR
    Write-Log "Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser" -Level INFO
    throw "Missing required module: ExchangeOnlineManagement"
}

try {
    $connectParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
    if ($UserPrincipalName) { $connectParams['UserPrincipalName'] = $UserPrincipalName }
    Connect-ExchangeOnline @connectParams
    Write-Log "Connected to Exchange Online successfully." -Level SUCCESS
}
catch {
    Write-Log "Failed to connect to Exchange Online: $_" -Level ERROR
    throw
}
