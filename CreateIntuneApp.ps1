#Ensure that IntuneWinAppUtil.exe is in the same directory as this script. 
#Change these variable to your desired default locations.
$OutputFolder = 'I:\Output'
$initialDirectory='I:\AppPrep'
Function Get-Folder($initialDirectory)

{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null
    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $initialDirectory

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}
Function Get-FileName($initialDirectory)
{  
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
    Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

$SetupFolder = Get-Folder
$SetupFile = Get-FileName -initialDirectory $SetupFolder

& $PSScriptRoot\IntuneWinAppUtil.exe -c $SetupFolder -s $SetupFile -o $OutputFolder