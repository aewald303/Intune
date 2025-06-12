function Uninstall-ProgramByName {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string] $ProgramName

    )
    # Get the program from the list of installed programs
    $app = Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -Like "*$ProgramName*"}

    # Check if the program was found
    if ($app) {
        "[INFORMATION] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] Attempting to uninstall $ProgramName" | Out-File -FilePath "$Logfile" -Append
        # Uninstall the program
        try {
            $UninstallAppExitCode = (Start-Process "msiexec.exe" -ArgumentList "/x $($app.IdentifyingNumber) /norestart /qn" -NoNewWindow -Wait -PassThru -ErrorAction Stop).ExitCode
            "[INFORMATION] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] Successfully uninstalled $ProgramName. Exit Code: $UninstallAppExitCode" | Out-File -FilePath "$Logfile" -Append
            return 0
        }
        catch {
            "[ERROR] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] $($ProgramName) Exit Code: $UninstallAppExitCode" | Out-File -FilePath "$Logfile" -Append
            "[ERROR] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] Unable to uninstall $($ProgramName): $_" | Out-File -FilePath "$Logfile" -Append
            return 1
        } 
    } else {
        "[INFORMATION] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] No $ProgramName application found to uninstall" | Out-File -FilePath "$Logfile" -Append
        return 0
    }   
}
# Get the directory where the script is located
$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Path -Parent

# Find all .msi files in the script's directory
$MSIFiles = Get-ChildItem -Path $ScriptDirectory -Filter "*.msi"

# Get the base name of the file (without the .msi extension)
$BaseNameStart = [System.IO.Path]::GetFileNameWithoutExtension($MSIFiles.Name)

# Split the base name by either '-' or '_'
$NamePartsStart = $BaseNameStart -split '[-_]', 2 # Split at most into two parts

$Logfile = "$ENV:Programdata\Microsoft\IntuneManagementExtension\Logs\win32-$($NamePartsStart[0])_InstallScript-$(Get-Date -Format "yyyy-MM-dd_HHmmss").log"
"[INFORMATION] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] Script called" | Out-File -FilePath "$Logfile" -Append

"[INFORMATION] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] Attempting to stop WINWORD.EXE" | Out-File -FilePath "$Logfile" -Append
Stop-Process -Name "WINWORD.EXE" -Force -ErrorAction SilentlyContinue
"[INFORMATION] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] Attempting to stop EXCEL.EXE" | Out-File -FilePath "$Logfile" -Append
Stop-Process -Name "EXCEL.EXE" -Force -ErrorAction SilentlyContinue
"[INFORMATION] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] Attempting to stop POWERPNT.EXE" | Out-File -FilePath "$Logfile" -Append
Stop-Process -Name "POWERPNT.EXE" -Force -ErrorAction SilentlyContinue
"[INFORMATION] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] Attempting to stop ACRORD32.EXE" | Out-File -FilePath "$Logfile" -Append
Stop-Process -Name "ACRORD32.EXE" -Force -ErrorAction SilentlyContinue


# Loop through each found .msi file
foreach ($MSIFile in $MSIFiles) {
    # Get the base name of the file (without the .msi extension)
    $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($MSIFile.Name)

    # Split the base name by either '-' or '_'
    $NameParts = $BaseName -split '[-_]', 2 # Split at most into two parts

    # Output the first part of the split name
    if ($NameParts.Count -gt 0) {
        try {
            "[INFORMATION] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] Attempting to stop $($NameParts[0]).exe" | Out-File -FilePath "$Logfile" -Append
            Stop-Process -Name "$($NameParts[0]).exe" -Force -ErrorAction SilentlyContinue 
            Uninstall-ProgramByName -ProgramName $NameParts[0] -ErrorAction Stop
        }
        catch {
            Exit 1
        }

    }
    else {
        "[INFORMATION] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] Nothing to Uninstall" | Out-File -FilePath "$Logfile" -Append
    }
}
#
try {
    "[INFORMATION] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] Starting to install $($NameParts[0])" | Out-File -FilePath "$Logfile" -Append
    $AdobeExitCode = (Start-Process "$ScriptDirectory\setup.exe" -ArgumentList "--silent" -NoNewWindow -Wait -PassThru -ErrorAction Stop).ExitCode
    "[INFORMATION] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] $($NameParts[0]) installer ran successfully. Exit code: $AdobeExitCode" | Out-File -FilePath "$Logfile" -Append 
    "[INFORMATION] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] Script completed successfully." | Out-File -FilePath "$Logfile" -Append
    Exit 0
}
catch {
    "[ERROR] [$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] Something went wrong installing $($NameParts[0]): $_" | Out-File -FilePath "$Logfile" -Append
    Exit 1
}

