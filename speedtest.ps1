$speedtestURL = "https://install.speedtest.net/app/cli/ookla-speedtest-1.0.0-win64.zip"
$speedtestZip = "C:\Windows\Temp\speedtest.zip"
$speedtestDir = "C:\Windows\Temp\speedtest\"
$speedtestExe = [ScriptBlock]::Create("C:\Windows\Temp\speedtest\speedtest.exe --progress=no --format=human-readable")

Invoke-WebRequest -Uri $speedtestURL -Outfile $speedtestZip
Expand-Archive -LiteralPath $speedtestZip -DestinationPath $speedtestDir -Force
Invoke-Command -ScriptBlock $speedtestExe
