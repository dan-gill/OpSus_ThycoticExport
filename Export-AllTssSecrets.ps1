#Requires -Version 5.1
#Requires -Modules @{ ModuleName='Thycotic.SecretServer'; ModuleVersion='0.60.7' }

# This script will export all secrets from Secret Server to a CSV file.
# The CSV file will contain the following columns:
# Secret Folder Path, Secret Name, Secret Password

$ErrorActionPreference = 'Stop'

# Button Values
enum Buttons {
    OK = 0
    OKCancel = 1
    AbortRetryIngnore = 2
    YesNoCancel = 3
    YesNo = 4
    RetryCancel = 5
    CancelTryAgainContinue = 6
}

# Icon Values
enum Icon {
    Stop = 16
    Question = 32
    Exclamation = 48
    Information = 64
}

# Return Values
enum Selection {
    None = -1
    OK = 1
    Cancel = 2
    Abort = 3
    Retry = 4
    Ignore = 5
    Yes = 6
    No = 7
    TryAgain = 10
    Continue = 11
}

# Prompt user for JSON file with settings
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = $PSScriptRoot
$OpenFileDialog.title = 'Select JSON file with settings'
$OpenFileDialog.filter = 'JavaScript Object Notation files (*.json)|*.json'
if ($OpenFileDialog.ShowDialog() -eq 'Cancel') {
    $wshell = New-Object -ComObject Wscript.Shell
    $null = $wshell.Popup('User canceled file selection. Exiting script.', 0, 'Exiting', `
            [Buttons]::OK + [Icon]::Exclamation)

    Exit 1223
}

$Settings = Get-Content "$($OpenFileDialog.filename)" -Raw | ConvertFrom-Json

# Load the Thycotic.SecretServer module
Import-Module Thycotic.SecretServer

# Prompt for Thycotic credentials
$ThycoticCreds = $null
$Session = $null
[Selection]$ButtonClicked = [Selection]::None

while (($null -eq $Session) -and ($ButtonClicked -ne [Selection]::Cancel)) {
    $ThycoticCreds = Get-Credential -Message 'Please enter your Thycotic credentials.'

    if ($ThycoticCreds) {
        $Error.Clear()
        try {
            # Create session on TSS
            $Session = New-TssSession -SecretServer $Settings.ssUri -Credential $ThycoticCreds `
                -ErrorAction $ErrorActionPreference
        } catch {
            $wshell = New-Object -ComObject Wscript.Shell
            $ButtonClicked = $wshell.Popup("Login to $($Settings.ssUri) failed. Retry?", 0, 'Failed login', `
                    [Buttons]::RetryCancel + [Icon]::Exclamation)
        } finally {
            $Error.Clear()
        }
    }
}

if ($ButtonClicked -eq [Selection]::Cancel) {
    if ($Session) { $null = Close-TssSession -TssSession $Session }
} else {
    try {
        # Get all secrets
        $Secrets = Search-TssSecret -TssSession $Session -IncludeSubFolders -IncludeInactive
        $Secrets | Select-Object SecretName, Active, @{Name = 'FolderPath'; Expression = {
            (Get-TssFolder -TssSession $Session -Id $_.FolderId).FolderPath }
        }, SecretTemplateName, @{Name = 'SecretPassword'; Expression = {
            (Get-TssSecret -TssSession $Session -Id $_.SecretId).GetFieldValue('Password') }
        } | Export-Csv -Path "$PSScriptRoot\secrets.csv" -NoTypeInformation
    } catch {
        Write-Error $_.Exception.Message
    } finally {
        if ($Session) { $null = Close-TssSession -TssSession $Session }
    }
}