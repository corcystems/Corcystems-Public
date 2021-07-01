#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#Per use variables modify as needed
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If ($ClientLocation -eq $null)
{$ClientLocation = '1'}

If ($ClientLocation -eq "")
{$ClientLocation = '1'}

$WriteOutput = $True

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#Constants do not change
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LabtechServerURL = "https://labtech.corcystems.com"
$LabtechUninstallerURL = "https://labtech.corcystems.com/labtech/service/LabUninstall.exe"
$LabtechInstallerURL = "https://labtech.corcystems.com/labtech/service/LabTechRemoteAgent.msi"
$LabtechUninstallerLocalPath = "C:\LabUninstall.exe"
$LabtechInstalerLocalPath = "C:\LabTechRemoteAgent.msi"
$LabtechFilesLocalPath = "C:\Windows\LTSvc"
$LTServices = @("LTSvcMon", "LTService")
$LTProcesses = @("LTSvcMon","LTSVC","LTClient","LTTray")
$LabtechServerPassword = '/STFO7fbHC/H7qighp5SQVQJi3rKlFfM'

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
#Installation
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

if (-not (LabtechIsInstalled)) {
    if ($WriteOutput) {Write-Host "Starting Installation"}
    #Delete any old installers and download a new one
	if (Test-Path -Path $LabtechInstalerLocalPath) {
	    Remove-Item -Path $LabtechInstalerLocalPath -Force -ErrorAction SilentlyContinue
		if ($WriteOutput) {Write-Host "Removed $LabtechInstalerLocalPath"}
	}
    Invoke-WebRequest -Uri $LabtechInstallerURL -Outfile $LabtechInstalerLocalPath
	if ($WriteOutput) {Write-Host "Downloaded new $LabtechInstalerLocalPath"}
	
    #Run the installer with the correct arguements
    msiexec.exe /i C:\LabTechRemoteAgent.msi /quiet /norestart SERVERADDRESS=$LabtechServerURL SERVERPASS=$LabtechServerPassword LOCATION=$ClientLocation
	if ($WriteOutput) {Write-Host "Started installer"}
    
	if ($WriteOutput) {Write-Host "Checking for sucessful install"}	
	#Wait for the services to install
	$ServiceChecks = 0
    while (-not (LabtechIsInstalled)) {
	    Start-Sleep 10
		$ServiceChecks += 1
		if ($ServiceChecks -eq 6) {if ($WriteOutput) {Write-Host "Service not found after 1 minute - exiting script"}; exit}
	}
	
	#Start the services to be safe
	ForEach ($LTService in $LTServices) {
	    Set-Service -Name $LTService -StartupType Automatic -ErrorAction SilentlyContinue
		if ($WriteOutput) {Write-Host "Enabled $LTService"}
		Start-Service -Name $LTService -ErrorAction SilentlyContinue
		if ($WriteOutput) {Write-Host "Started $LTService"}
	}	
} else {
	if ($WriteOutput) {Write-Host "LT Services exist no install performed"}
}
