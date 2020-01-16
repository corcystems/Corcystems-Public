if (Test-Path c:\)
{
taskkill /im ltsvcmon.exe /f
taskkill /im lttray.exe /f
taskkill /im ltsvc.exe /f
sc.exe config "LTService" start= disabled
sc.exe config "LTSvcMon" start= disabled
taskkill /im ltsvcmon.exe /f
taskkill /im lttray.exe /f
taskkill /im ltsvc.exe /f
sc.exe config "LTService" start= disabled
sc.exe config "LTSvcMon" start= disabled
$source = "https://labtech.corcystems.com/labtech/service/LabUninstall.exe"
$Filename = [System.IO.Path]::GetFileName($source)
$dest = "C:\$Filename"
$wc = New-Object System.Net.WebClient
if (!(test-path $dest))
{
if ((Test-Path $dest -OlderThan (Get-Date).AddHours(-24)))
{
Write-Output '------File is older than 24 hours old----'
remove-item C:\LabUninstall.exe}
write-Output '------Downloading Uninstaller Now--------'
$wc.DownloadFile($source, $dest)
}
Else
{
Write-Output '---Uninstaller Already Resides on C:\----'
Write-Output '---Using c:\Labuninstall.exe-------------'
Write-Output '---If you have issues with the uninstall-'
Write-Output '---Delete the uninstaller and run again--'
}
Write-Output '---------------Uninstalling--------------'
C:\LabUninstall.exe /quiet /norestart
Write-Output '----Uninstall Started Waiting 90 Secs----'
Start-Sleep 90
}
Else
{
Write-Output '------LTSVC FOLDER DOES NOT EXIST--------'
Write-Output '-------SKIPPING UNINSTALL PROCESS--------'
}