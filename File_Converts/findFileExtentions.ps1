Function Get-Folder($initialDirectory="")

{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

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

$extentionType = Read-Host "What extention do you want to search for?"
$extentionSearch = "*$extentionType"

Write-Host "What directory do you want to search for $extentionSearch?"
$selectedFolder = Get-Folder

$foundFiles = Get-ChildItem -Path $selectedFolder -Include $extentionSearch -Recurse

Write-Host "Where do you want the CSV Output Placed?"
$outputFolder = Get-Folder
$currentDay = Get-Date -Format "MM.dd"

$outputFile = "$outputFolder\$extentionType.$currentDay.csv"

$outputFile

$foundFiles | Select Directory, Name | Export-CSV -Path $outputFile

$countFiles = ($foundFiles | measure-object).count

Write-Host "We found $countFoundFiles $extentionType files in the $selectedFolder directory."