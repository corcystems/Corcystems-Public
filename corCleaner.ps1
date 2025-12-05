# Paths to Clean
$cleanPath = @()
$cleanPath += "$env:windir\Temp" # Windows Temp folder
$cleanPath += "$env:TEMP" #User Temp folder
$cleanPath += "C:\Windows\SoftwareDistribution\Download" # Windows Update cache
$cleanPath += "$env:LOCALAPPDATA\Microsoft\Windows\INetCache" # Internet Explorer / Edge legacy cache (if exists)


# Create CotTools folder
if(-not (Test-Path "C:\CorTools\")){
	Write-Host "Creating C:\CorTools\ folder."
	mkdir C:\CorTools\
}

# Log file path
$LogFile = "C:\CorTools\DiskCleanup_$(Get-Date -Format 'yyyyMM').log"

# Log Disk Cleanup Started
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"$timestamp - === Disk Cleanup Started ===" | Out-File -FilePath $LogFile -Append
Write-Host "=== Disk Cleanup Started ==="

# Initial Size Pull
$driveSizeStart = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -eq "$($env:SystemDrive)\" }
$driveSizeStartGB = [math]::Round($driveSizeStart.Free / 1GB, 2)

# Log Initial Size Pull
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"$timestamp - $driveSizeStart Drive has $driveSizeStartGB GB free" | Out-File -FilePath $LogFile -Append
Write-Host "$driveSizeStart Drive has $driveSizeStartGB GB free"

# Clean the Paths
foreach ($path in $cleanPath){
	if (Test-Path $path) {
		try {
            Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue |
                Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
				# Log Cleared: $path
				$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				"$timestamp - Cleared: $path" | Out-File -FilePath $LogFile -Append
				Write-Host "Cleared: $path"			
        }
        catch {
			# Log Error clearing $path: $_
			$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
			"$timestamp - Error clearing $path $_" | Out-File -FilePath $LogFile -Append
			Write-Host "Error clearing $path $_"	
        }
    }
    else {
		# Log Path not found: $path
		$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
		"$timestamp - Path not found: $path" | Out-File -FilePath $LogFile -Append
		Write-Host "Path not found: $path"	
    }
}

# Clear the Recycle Bin
try {
    Clear-RecycleBin -Force -ErrorAction SilentlyContinue
	# Log Recycle Bin emptied.
	$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	"$timestamp - Recycle Bin emptied." | Out-File -FilePath $LogFile -Append
	Write-Host "Recycle Bin emptied."
}
catch {
	# Log Error emptying Recycle Bin: $_
	$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	"$timestamp - Error emptying Recycle Bin $_" | Out-File -FilePath $LogFile -Append
	Write-Host "Error emptying Recycle Bin $_"
}

# Ending Size Pull and Saved GB
$driveSizeEnd = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -eq "$($env:SystemDrive)\" }
$driveSizeEndGB = [math]::Round($driveSizeStart.Free / 1GB, 2)
$savedSpaceGB = $driveSizeEndGB - $driveSizeStartGB

# Log Ending Size Pull and saved
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"$timestamp - $driveSizeEnd Drive has $driveSizeEndGB GB free" | Out-File -FilePath $LogFile -Append
Write-Host "$driveSizeEnd Drive has $driveSizeEndGB GB free"
"$timestamp - Saved $savedSpaceGB GB on the $driveSizeEnd Drive." | Out-File -FilePath $LogFile -Append
Write-Host "Saved $savedSpaceGB GB on the $driveSizeEnd Drive."

# Log Disk Cleanup Completed
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"$timestamp - === Disk Cleanup Completed ===" | Out-File -FilePath $LogFile -Append
Write-Host "=== Disk Cleanup Completed ==="	
Write-Host "Cleanup complete. Log saved to: $LogFile"

