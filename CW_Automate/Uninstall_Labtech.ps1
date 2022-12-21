[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$WriteOutput = $True

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#Constants do not change
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LabtechUninstallerURL = "https://labtech.corcystems.com/labtech/service/LabUninstall.exe"
$LabtechUninstallerLocalPath = "C:\LabUninstall.exe"
$LTServices = @("LTSvcMon", "LTService")
$LTProcesses = @("LTSvcMon","LTSVC","LTClient","LTTray")

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#Helper functions do not change
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function LabtechIsInstalled {
   ForEach ($LTService in $LTServices) {
        if (Get-Service -Name $LTService -ErrorAction SilentlyContinue) {
		    if ($WriteOutput) {Write-Host "Found $LTService"}
		    return $True
		}	    
	}
	if ($WriteOutput) {Write-Host "No LT Services Found"}
	return $False
}

if (LabtechIsInstalled) {
    if ($WriteOutput) {Write-Host "LabTech Found"}
} else {
    if ($WriteOutput) {Write-Host "LabTech NOT Found"}
}

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#Generic steps
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#Changes current working directory so that uninstallation files are dropped in root C:\
Set-Location -Path C:\

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#Uninstallation
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

if ($WriteOutput) {Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"}
if ($WriteOutput) {Write-Host "Forced Uninstallation Sequence"}
if ($WriteOutput) {Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"}

if ($True) {
	if ($WriteOutput) {Write-Host "Starting Uninstallation"}

    #Stop and disable LabTech Services and Processes
    if ($WriteOutput) {Write-Host "Stopping Services"}	
    ForEach ($LTService in $LTServices) {
	    Set-Service -Name $LTService -StartupType Disabled -ErrorAction SilentlyContinue
		if ($WriteOutput) {Write-Host "Disabled $LTService"}
		Stop-Service -Name $LTService -Force -ErrorAction SilentlyContinue | Out-Null
		if ($WriteOutput) {Write-Host "Stopped $LTService"}
	}
	if ($WriteOutput) {Write-Host "Stopping Processes"}
	ForEach ($LTProcess in $LTProcesses) {
	    Stop-Process -Name $LTProcess -Force -ErrorAction SilentlyContinue
		if ($WriteOutput) {Write-Host "Stopped $LTProcess"}
	}
	
	#Remove old versions of uninstaller and download a new one
	if (Test-Path -Path $LabtechUninstallerLocalPath) {
	    Remove-Item -Path $LabtechUninstallerLocalPath -Force -ErrorAction SilentlyContinue
		if ($WriteOutput) {Write-Host "Removed $LabtechUninstallerLocalPath"}
	}	
	Invoke-WebRequest -Uri $LabtechUninstallerURL -Outfile $LabtechUninstallerLocalPath
	if ($WriteOutput) {Write-Host "Downloaded new $LabtechUninstallerLocalPath"}
	Start-Sleep -Seconds 5
	
	#Run uninstaller
	C:\LabUninstall.exe /quiet /norestart
	if ($WriteOutput) {Write-Host "Started Uninstaller"}
	if ($WriteOutput) {Write-Host "Waiting for uninstallation to finish"}
	while (Test-Path -Path $LabtechFilesLocalPath) {
	    Start-Sleep -Seconds 1
	}
	if ($WriteOutput) {Write-Host "Uninstallation finished"}
	Start-Sleep -Seconds 10
	if ($WriteOutput) {Write-Host "Waiting an additional 15 seconds"}
}
