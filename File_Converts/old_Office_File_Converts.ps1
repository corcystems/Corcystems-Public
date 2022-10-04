#######################
### Script Settings ###
#######################

# Clear the screen
cls




#######################
###### Functions ######
#######################

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


# A function to grab the wanted file type to convert.
function File_Extension_Form{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    
    # Set the size of your form
    $Form = New-Object System.Windows.Forms.Form
    $Form.width = 500
    $Form.height = 300
    $Form.Text = "Microsoft File Convertion Chooser"
 
    # Set the font of the text to be used within the form
    $Font = New-Object System.Drawing.Font("Times New Roman",12)
    $Form.Font = $Font
 
    # Create a group that will contain your radio buttons
    $MyGroupBox = New-Object System.Windows.Forms.GroupBox
    $MyGroupBox.Location = '40,30'
    $MyGroupBox.size = '400,110'
    $MyGroupBox.text = "What file type do you want to convert?"
    
    # Create the collection of radio buttons
    $RadioButtonDoc = New-Object System.Windows.Forms.RadioButton
    $RadioButtonDoc.Location = '20,40'
    $RadioButtonDoc.size = '350,20'
    $RadioButtonDoc.Checked = $true 
    $RadioButtonDoc.Text = "Microsoft Word 97-2003 .doc to .docx."
 
    $RadioButtonXls = New-Object System.Windows.Forms.RadioButton
    $RadioButtonXls.Location = '20,70'
    $RadioButtonXls.size = '350,20'
    $RadioButtonXls.Checked = $false
    $RadioButtonXls.Text = "Microsoft Excel 97-2003 .xls to .xlsx."
 
    # Add an OK button
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = '130,200'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult=[System.Windows.Forms.DialogResult]::OK
 
    #Add a cancel button
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '255,200'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
 
    # Add all the GroupBox controls on one line
    $MyGroupBox.Controls.AddRange(@($RadioButtonDoc,$RadioButtonXls))
 
    # Add all the Form controls on one line 
    $form.Controls.AddRange(@($MyGroupBox,$OKButton,$CancelButton))
 
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $form.AcceptButton = $OKButton
    $form.CancelButton = $CancelButton
 
    # Activate the form
    $form.Add_Shown({$form.Activate()})    
    
    # Get the results from the button click
    $dialogResult = $form.ShowDialog()
 
    # If the OK button is selected
    if ($dialogResult -eq "OK"){
        
        # Return that doc was chosen.
        if ($RadioButtonDoc.Checked){
            $oldFileType = "doc"
            }
	# Return that xls was chosen.
        elseif ($RadioButtonXls.Checked){
            $oldFileType = "xls"
            }
        }
	return $oldFileType
} # End of the File_Extension_Form function





#######################
###### Variables ######
#######################

# Call the file extention function
$fileTypeSelection = File_Extension_Form

$filetype ="*" + $fileTypeSelection
$newFileExt = "." + $fileTypeSelection +"x"

Write-host "You chose to convert $fileTypeSelection to $newFileExt."

# Make sure a file type was selected.
if ($fileTypeSelection -eq $null){
	Write-host "Filetype was not selected, please run the script again and choose a filetype then hit OK."
	exit
}

# Call the Get-Folder
$folderpath = Get-Folder
Write-host "Converting all $fileTypeSelection files in $folderpath and all sub directories."

# Make sure a folder path was selected.
if ($folderpath -eq $null){
	Write-host "Folderpath was not selected, please run the script again and choose a folderpath then hit OK."
	exit
}




#######################
## User Confirmation ##
#######################

$userConfirm = Read-Host "Can you confirm the info above? (Y/N)"

if ($userConfirm -ne "Y"){
	Write-host "Answer was not Y, exiting script."
	exit
}




#######################
####### Actions #######
#######################

# Create the Word object
Add-Type -AssemblyName Microsoft.Office.Interop.Word
$docFixedFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLDocument
write-host $docFixedFormat
$word = New-Object -ComObject Word.Application
$word.visible = $true

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
 
		$path += $newFileExt
		$document.saveas($path, $docFixedFormat)
		$document.close()

		# Move the old file to another directory
		$oldFolder = $path.substring(0, $path.lastIndexOf("\")) + "\old"
		write-host $oldFolder
		
			# Test if the old folder path was already created, if not create it.
			if(-not (test-path $oldFolder)){
				new-item $oldFolder -type directory
			}

		#Grab Last modified Date
		$fileLastModified = $_.LastWriteTime
		#Update Last Modified time on the new file to match the old modified time
		Get-ChildItem $path | % {$_.LastWriteTime = $fileLastModified}
		
		# Finally move the old file
		move-item $_.fullname $oldFolder
	} # End of the convertion Loop




#######################
####### Cleanup #######
#######################

# End of script cleanup
$word.Quit()
$word = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()
