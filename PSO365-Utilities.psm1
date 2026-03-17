#Requires -Version 5.1
<#
.SYNOPSIS
    Shared utility functions for PSforO365 scripts.

.DESCRIPTION
    Common functions used across multiple Office 365 / Exchange management scripts:
    - Logging
    - O365 session connection
    - CSV file selection
    - Module/prerequisite validation
    - Confirmation prompts for destructive operations

.NOTES
    Version: 1.0
    Repository: PSforO365
#>

#region Logging

$script:LogFile = $null

function Initialize-Log {
    <#
    .SYNOPSIS
        Initializes a log file for the current session.
    .PARAMETER LogDirectory
        Directory to write log files. Defaults to .\Logs.
    .PARAMETER ScriptName
        Name to use in the log filename.
    #>
    [CmdletBinding()]
    param(
        [string]$LogDirectory = ".\Logs",
        [string]$ScriptName = "PSO365"
    )
    if (-not (Test-Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $script:LogFile = Join-Path $LogDirectory "${ScriptName}_${timestamp}.log"
    Write-Log "Log initialized. Script: $ScriptName" -Level INFO
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped message to the console and optionally to a log file.
    .PARAMETER Message
        The message to log.
    .PARAMETER Level
        Log level: INFO, WARN, ERROR, SUCCESS. Defaults to INFO.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet('INFO','WARN','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        'INFO'    { Write-Host $entry -ForegroundColor Cyan }
        'WARN'    { Write-Host $entry -ForegroundColor Yellow }
        'ERROR'   { Write-Host $entry -ForegroundColor Red }
        'SUCCESS' { Write-Host $entry -ForegroundColor Green }
    }

    if ($script:LogFile) {
        $entry | Out-File -FilePath $script:LogFile -Append -Encoding UTF8
    }
}

#endregion

#region Connection Management

function Connect-O365 {
    <#
    .SYNOPSIS
        Connects to Exchange Online using modern authentication (EXO V2/V3 module).
    .DESCRIPTION
        Checks for the ExchangeOnlineManagement module and connects. Falls back to
        basic remote PowerShell if the module is not available (legacy environments).
    .PARAMETER UserPrincipalName
        Admin UPN to authenticate with (optional; omit to be prompted).
    .PARAMETER LegacyMode
        Use the old New-PSSession method (Exchange on-premises or EXO V1 tenants).
    #>
    [CmdletBinding()]
    param(
        [string]$UserPrincipalName,
        [switch]$LegacyMode
    )

    if ($LegacyMode) {
        Write-Log "Connecting to Exchange Online (legacy PSSession mode)..." -Level INFO
        try {
            $cred = Get-Credential -Message "Enter your Office 365 admin credentials" `
                                   -UserName $UserPrincipalName
            $session = New-PSSession `
                -ConfigurationName Microsoft.Exchange `
                -ConnectionUri https://ps.outlook.com/powershell/ `
                -Credential $cred `
                -Authentication Basic `
                -AllowRedirection `
                -ErrorAction Stop
            Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
            Write-Log "Legacy PSSession connected successfully." -Level SUCCESS
        }
        catch {
            Write-Log "Failed to create legacy PSSession: $_" -Level ERROR
            throw
        }
        return
    }

    # Modern EXO V2/V3
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log "ExchangeOnlineManagement module not found. Install with: Install-Module ExchangeOnlineManagement" -Level ERROR
        throw "Missing module: ExchangeOnlineManagement"
    }

    Write-Log "Connecting to Exchange Online (ExchangeOnlineManagement module)..." -Level INFO
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
}

function Connect-MSOnline {
    <#
    .SYNOPSIS
        Connects to Microsoft Online Service (MSOnline / MSOL).
    .DESCRIPTION
        Validates the MSOnline module is present then connects.
    #>
    [CmdletBinding()]
    param()

    Assert-ModuleAvailable -ModuleName MSOnline
    Write-Log "Connecting to Microsoft Online Service..." -Level INFO
    try {
        Connect-MsolService -ErrorAction Stop
        Write-Log "Connected to MSOnline successfully." -Level SUCCESS
    }
    catch {
        Write-Log "Failed to connect to MSOnline: $_" -Level ERROR
        throw
    }
}

#endregion

#region Module / Prerequisite Validation

function Assert-ModuleAvailable {
    <#
    .SYNOPSIS
        Throws an error if a required PowerShell module is not installed.
    .PARAMETER ModuleName
        Name of the module to check.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        $msg = "Required module '$ModuleName' is not installed. Run: Install-Module $ModuleName"
        Write-Log $msg -Level ERROR
        throw $msg
    }
    Write-Log "Module '$ModuleName' is available." -Level INFO
}

function Assert-RunningAsAdmin {
    <#
    .SYNOPSIS
        Throws if the script is not running with administrator privileges.
    #>
    $current = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($current)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $msg = "This script must be run as Administrator."
        Write-Log $msg -Level ERROR
        throw $msg
    }
}

#endregion

#region File Helpers

function Get-CsvFilePath {
    <#
    .SYNOPSIS
        Opens a GUI file-picker dialog and returns the selected CSV path.
    .PARAMETER InitialDirectory
        Starting directory for the dialog. Defaults to user's Desktop.
    .PARAMETER Title
        Dialog window title.
    #>
    [CmdletBinding()]
    param(
        [string]$InitialDirectory = [Environment]::GetFolderPath('Desktop'),
        [string]$Title = "Select CSV File"
    )
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title            = $Title
    $dialog.InitialDirectory = $InitialDirectory
    $dialog.Filter           = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }
    else {
        throw "No file selected."
    }
}

function Import-CsvSafe {
    <#
    .SYNOPSIS
        Imports a CSV with validation. Throws if the file is missing or empty.
    .PARAMETER Path
        Path to the CSV file.
    .PARAMETER RequiredColumns
        Optional array of column names that must exist in the CSV.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [string[]]$RequiredColumns
    )
    if (-not (Test-Path $Path)) {
        $msg = "CSV file not found: $Path"
        Write-Log $msg -Level ERROR
        throw $msg
    }
    $data = Import-Csv -Path $Path
    if (-not $data) {
        $msg = "CSV file is empty: $Path"
        Write-Log $msg -Level ERROR
        throw $msg
    }
    if ($RequiredColumns) {
        $headers = $data[0].PSObject.Properties.Name
        foreach ($col in $RequiredColumns) {
            if ($col -notin $headers) {
                $msg = "CSV is missing required column: '$col'"
                Write-Log $msg -Level ERROR
                throw $msg
            }
        }
    }
    Write-Log "CSV loaded: $Path ($($data.Count) rows)" -Level INFO
    return $data
}

#endregion

#region Safety Helpers

function Confirm-DestructiveAction {
    <#
    .SYNOPSIS
        Prompts the user to confirm a potentially destructive action.
    .PARAMETER ActionDescription
        Human-readable description of what the action will do.
    .PARAMETER Force
        Bypasses the prompt (for automation / -Force scenarios).
    .RETURNS
        $true if confirmed, $false if cancelled.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ActionDescription,
        [switch]$Force
    )
    if ($Force) { return $true }

    Write-Host ""
    Write-Host "  *** DESTRUCTIVE ACTION WARNING ***" -ForegroundColor Red
    Write-Host "  $ActionDescription" -ForegroundColor Yellow
    Write-Host ""
    $response = Read-Host "  Type 'YES' to confirm, anything else to cancel"
    if ($response -eq 'YES') {
        Write-Log "User confirmed: $ActionDescription" -Level WARN
        return $true
    }
    Write-Log "User cancelled: $ActionDescription" -Level INFO
    return $false
}

#endregion

Export-ModuleMember -Function `
    Initialize-Log, Write-Log, `
    Connect-O365, Connect-MSOnline, `
    Assert-ModuleAvailable, Assert-RunningAsAdmin, `
    Get-CsvFilePath, Import-CsvSafe, `
    Confirm-DestructiveAction
