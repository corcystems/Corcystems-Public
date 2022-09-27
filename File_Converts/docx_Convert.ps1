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

$folderpath = Get-Folder

Add-Type -AssemblyName Microsoft.Office.Interop.Word
$docFixedFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLDocument
write-host $docFixedFormat
$word = New-Object -ComObject Word.Application
$word.visible = $true
$filetype ="*doc"

$allFiles = Get-ChildItem -Path $folderpath -Include $filetype -recurse

$allFiles | ForEach-Object -Begin {
	# Set the $i counter variable to zero.
	$i = 0

} -Process {

	$i = $i+1
	$Completed = ($i/$allFiles.count*100)
	Write-Progress -Activity "Searching Files" -Status "$Completed% Complete:" -PercentComplete $Completed

	$path = ($_.fullname).substring(0, ($_.FullName).lastindexOf("."))
    
	"Converting $path"
	$document = $word.Documents.open($_.fullname)
 
	$path += ".docx"
	$document.saveas($path, $docFixedFormat)
	$document.close()
    
	$oldFolder = $path.substring(0, $path.lastIndexOf("\")) + "\old"
    
	write-host $oldFolder
	if(-not (test-path $oldFolder))
	{
		new-item $oldFolder -type directory
	}

	move-item $_.fullname $oldFolder
    
}
$word.Quit()
$word = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()