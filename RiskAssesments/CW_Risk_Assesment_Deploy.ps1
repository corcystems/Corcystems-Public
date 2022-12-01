#############   Start Script   ################

###############################################
### CW Risk Assessment Agent Deploy & Run   ###
### This script was written by Aaron Falzon ###
### Use of this script is at your own risk  ###
###############################################

#Get Client Name & Token from the IT Support Portal under: Security > Assessment > (Create Assessment or select client name if already created) > (token is in the link, and starts with 'TKN')
$clientName = "" ## Note: Place Client Name here as setup in the ITS Portal
$token = "" ## Note: Place Token here as Displayed in the ITS Portal (starting with TKN)

#leave true to remove files after the script has ran, this will also remove the log files from the PC, if you want the logs & files to be left set to $false
$deleteFile = $true

$url = https://prod.setup.itsupport247.net/downloadassessment/$clientName/$token
$path = "c:\RiskAssess\"
$file = "$path\ConnectWise-assessment-utility_$token.exe"
if (Test-Path -Path $path) {
    "Path exists!"
} else {
    "Path doesn't exist."
    New-Item -Path "c:\" -Name "RiskAssess" -ItemType "directory"
}

#iwr $url -OutFile $path\temp.tmp
$WebClient = New-Object System.Net.WebClient
$WebClient.DownloadFile($url,"$file")

start-process $file -WorkingDirectory $path -Wait

if ($deleteFile -eq $true){
    Remove-Item -Force -Recurse -Path "$path"
}

#############   End Script   ################
