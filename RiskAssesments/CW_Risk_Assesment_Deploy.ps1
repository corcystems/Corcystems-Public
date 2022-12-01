if($clientURL -eq $null){
	Write-host "No URL variable was passed. Exiting script."
	exit
}

$clientURLArray = $clientURL.Split("/")

$clientName = $clientURLArray[($clientURLArray.count - 2)]
$token = $clientURLArray[($clientURLArray.count - 1)]


#$url = https://prod.setup.itsupport247.net/downloadassessment/$clientName/$token
$path = "c:\RiskAssess\"
$file = "$path\ConnectWise-assessment-utility_$token.exe"
if(Test-Path -Path $path){
    "Path exists!"
}else{
    "Path doesn't exist."
    New-Item -Path "c:\" -Name "RiskAssess" -ItemType "directory"
}

$WebClient = New-Object System.Net.WebClient
$WebClient.DownloadFile($clientURL,"$file")

Start-Process $file -WorkingDirectory $path -Wait

Remove-Item -Force -Recurse -Path "$path"
