#Requires -Version 5.1
#Requires -Modules @{ ModuleName='Thycotic.SecretServer'; ModuleVersion='0.60.7' }

# This script will export all secrets from Secret Server to a CSV file.
# The CSV file will contain the following columns:
# Secret Folder Path, Secret Name, Secret Password

$ErrorActionPreference = 'Stop'
$CsvPath = "$PSScriptRoot\secrets.csv"

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

# Prompt user for JSON file with settings; Allow script to run on non-Windows systems
try {
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
} catch {
    $Settings = Get-Content "$PSScriptRoot\.settings\settings.json" -Raw | ConvertFrom-Json
}


# Load the Thycotic.SecretServer module
Import-Module Thycotic.SecretServer

# Prompt for Thycotic credentials
$ThycoticCreds = $null
$Session = $null
[Selection]$ButtonClicked = [Selection]::None
$logInfoParam = @{
    LogFilePath = "$PSScriptRoot/.settings/Export-AllTssSecrets.log"
}

while (($null -eq $Session) -and ($ButtonClicked -ne [Selection]::Cancel)) {
    $ThycoticCreds = Get-Credential -Message 'Please enter your Thycotic credentials.'

    if ($ThycoticCreds) {
        $Error.Clear()
        try {

            # Create TSS log
            Start-TssLog @logInfoParam
            # Create session on TSS
            $Session = New-TssSession -SecretServer $Settings.ssUri -Credential $ThycoticCreds `
                -ErrorAction $ErrorActionPreference
            Write-TssLog @logInfoParam -Message "Token Time of Death: $($Session.TimeOfDeath)"
        } catch {
            $wshell = New-Object -ComObject Wscript.Shell
            $ButtonClicked = $wshell.Popup("Login to $($Settings.ssUri) failed. Retry?", 0, 'Failed login', `
                    [Buttons]::RetryCancel + [Icon]::Exclamation)
        } finally {
            $Error.Clear()
        }
    }
}

# Delete CSV file if it exists
if (Test-Path $CsvPath -PathType Leaf) {
    Remove-Item $CsvPath -Force
}

if ($ButtonClicked -eq [Selection]::Cancel) {
    if ($Session) { $null = Close-TssSession -TssSession $Session }
} else {
    try {
        Write-Host "$(Get-Date -Format G): Starting export."
        # Get all folders
        $Folders = Search-TssFolder -TssSession $Session
        Write-Host "$(Get-Date -Format G): Found $($Folders.Count) folders."
        # Loop through folders and get secrets and passwords
        $WriteProgressParams = @{
            Activity        = 'Exporting secrets'
            Status          = 'Exporting secrets'
            PercentComplete = 0
        }
        Write-Progress @WriteProgressParams
        $index = 0
        foreach ($Folder in $Folders) {
            # Check if sessions is within three minutes of timeout
            if ($Session.CheckTokenTtl(5)) {
                Write-TssLog @logInfoParam -Message 'Token nearing expiration, attempting to renew'
                try {
                    $null = Close-TssSession -TssSession $Session
                    $Session = New-TssSession -SecretServer $Settings.ssUri -Credential $ThycoticCreds `
                        -ErrorAction $ErrorActionPreference
                } catch {
                    Write-TssLog @logInfoParam -Message 'Token renewal failed, exiting script.'
                    break
                }
                Write-TssLog @logInfoParam -Message "Token renewal successful. New token Time of Death: $($Session.TimeOfDeath)"
            }
            Write-Host "$(Get-Date -Format G): Getting secrets from $($Folder.FolderPath)."
            $FolderSecrets = Search-TssSecret -TssSession $Session -FolderId $Folder.FolderId -IncludeInactive
            $FolderSecrets | Select-Object SecretName, Active, @{Name = 'FolderPath'; Expression = {
                    $Folder.FolderPath }
            }, SecretTemplateName, @{Name = 'SecretPassword'; Expression = {
                         (Get-TssSecret -TssSession $Session -Id $_.SecretId).GetFieldValue('Password') }

            } | Export-Csv -Path $CsvPath -NoTypeInformation -Append
            $index++
            $WriteProgressParams = @{
                Activity        = 'Exporting secrets'
                Status          = "$(Get-Date -Format G): Found $($FolderSecrets.Count) secrets in $($Folder.FolderPath)."
                PercentComplete = [Math]::Round(($index / $Folders.Count) * 100, 0)
            }
            Write-Progress @WriteProgressParams
        }
        Write-Progress @WriteProgressParams -Completed
    } catch {
        Write-Error $_.Exception.Message
    } finally {
        if ($Session) { $null = Close-TssSession -TssSession $Session }
        Stop-TssLog @logInfoParam
    }
}