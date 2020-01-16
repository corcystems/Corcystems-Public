param
(
  [Parameter(Mandatory=$True,Position=1)]
   [string]$clientLocation
)

$FQDN='https://labtech.corcystems.com'
$SERVERPASS='eI11d8I6dGMxW+mqzBwSJA=='
#--------------------------------------------------------

Write-Host "$clientLocation"

$serviceName = 'LTService'
If (Get-Service $serviceName -ErrorAction SilentlyContinue) {
Write-Host "$serviceName service already installed. Exiting Script"
} Else {
Write-Output '----------Download Agent Sequence---------'
$source2 = "$($FQDN)/labtech/service/LabTechRemoteAgent.msi"
$Filename = [System.IO.Path]::GetFileName($source2)
$dest2 = "C:\$Filename"
$wc = New-Object System.Net.WebClient
$file2 = 'C:\LabTechRemoteAgent.msi'
if (!(test-path $file2))
{if ((Test-Path $file2 -OlderThan (Get-Date).AddHours(-24))){
Write-Output '-----------------------------------------'
Write-Output '------File is older than 24 hours old----'
Write-Output '------Deleting Old Installer-----------'
remove-item C:\LabTechRemoteAgent.msi}
Write-Output '-----------------------------------------'
write-Output '--------Downloading Installer Now--------'
$wc.DownloadFile($source2, $dest2)
} Else {
Write-Output '----Installer Already Resides on C:\-----'
Write-Output '---Using c:\LabtechRemoteAgent.msi-------'
Write-Output '---If you have issues with the install---'
Write-Output '---Delete the installer and run again----'}
Write-Output '-----------------------------------------'
Write-Output '----------------Installing---------------'
msiexec.exe /i C:\LabTechRemoteAgent.msi /quiet /norestart SERVERADDRESS=$($FQDN) SERVERPASS=$($SERVERPASS) LOCATION=$($ClientLocation)
Start-Sleep -s 60
Write-Output '-----------------------------------------'
Write-Output '------Verifying Services are Started-----'
sc.exe config "LTService" start= auto
sc.exe config "LTSvcMon" start= auto
sc.exe start ltsvcmon
sc.exe start ltservice
}
