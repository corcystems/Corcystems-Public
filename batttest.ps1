$Current_User = $env:UserName
$OutFile_Dir = "C:\Users\{0}\AppData\Local\Temp\" -f $Current_User
$OutFile_Name = "Battery_Info.xml"
$OutFile_FullPath = $OutFile_Dir + $OutFile_Name
$PercentDifferce = 0.8
​
if (Test-Path -Path $OutFile_FullPath) {
    Remove-Item -Path $OutFile_FullPath
}
​
try {
    POWERCFG /BATTERYREPORT /OUTPUT $OutFile_FullPath /XML | Out-Null
} catch {
    Write-Host "Error running battery report"
    exit
}
​
if (Test-Path -Path $OutFile_FullPath) {
    $XMLData = New-Object System.XML.XMLDocument
    $XMLData.Load($OutFile_FullPath)

    $LastTwoUsageEntries = $XMLData.BatteryReport.RecentUsage.UsageEntry | select -last 2

    $MostRecent = [int]$LastTwoUsageEntries[1].FullChargeCapacity
    $NextMostRecent = [int]$LastTwoUsageEntries[0].FullChargeCapacity
    $CutoffAmount = $NextMostRecent * $PercentDifferce

​    $CutoffAmount
 if ($MostRecent -lt $CutoffAmount) {
    Write-Host "Battery Max Charge Fell by a detrimental amount"
    }
}
