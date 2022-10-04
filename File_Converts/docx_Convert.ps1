# Function to pull the folder path to convert. Includes all subfolders.
Function Get-Folder($initialDirectory=""){
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

	$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
	$foldername.Description = "Select a folder"
	$foldername.rootfolder = "MyComputer"
	$foldername.SelectedPath = $initialDirectory

	if($foldername.ShowDialog() -eq "OK"){
	$folder += $foldername.SelectedPath
	}
	return $folder
} # End of the Get-Folder function

# Call the Get-Folder 
$folderpath = Get-Folder

# Create the Word object
Add-Type -AssemblyName Microsoft.Office.Interop.Word
$docFixedFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLDocument
write-host $docFixedFormat
$word = New-Object -ComObject Word.Application
$word.visible = $true
$filetype ="*doc"

# Grab all files in the path chosen above.
$allFiles = Get-ChildItem -Path $folderpath -Include $filetype -recurse

# Start of the Convertion Loop.
# Loops through all files found above.
# Begin is just setting a counter for the progress bar.
$allFiles | ForEach-Object -Begin {
	# Set the $i counter variable to zero.
	$i = 0

	# Actual loop that opens each file found matching the extention, saves, closes, then moved the old file.
	} -Process {

		# Raise the counter by 1 for the progress bar.
		$i = $i+1
		# Progress bar lines
		$Completed = ($i/$allFiles.count*100)
		Write-Progress -Activity "Searching Files" -Status "$Completed% Complete:" -PercentComplete $Completed

		# Set the path variable.
		$path = ($_.fullname).substring(0, ($_.FullName).lastindexOf("."))
	
		# Converting File
		"Converting $path"
		$document = $word.Documents.open($_.fullname)
 
		$path += ".docx"
		$document.saveas($path, $docFixedFormat)
		$document.close()
	
		# Move the old file to another directory
		$oldFolder = $path.substring(0, $path.lastIndexOf("\")) + "\old"
		write-host $oldFolder
		
			# Test if the old folder path was already created, if not create it.
			if(-not (test-path $oldFolder)){
				new-item $oldFolder -type directory
			}

		move-item $_.fullname $oldFolder
	
		#Grab Last modified Date
		$fileLastModified = $_.LastWriteTime
	
		#Update Last Modified time on the new file to match the old modified time
		Get-ChildItem $_.fullname | % {$_.LastWriteTime = $fileLastModified}
	} # End of the convertion Loop

# End of script cleanup
$word.Quit()
$word = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()
