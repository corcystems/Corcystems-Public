$errorActionPreference = "SilentlyContinue"
$cwaRegInfo = Get-ItemProperty -Path "HKLM:\SOFTWARE\LabTech\Service"
if($Error) {Write-Host "Not Installed"; exit}
if($null -ne $cwaRegInfo.ID -and $cwaRegInfo.ID -gt 0) {
	$cwaOutput = Get-ItemProperty -Path HKLM:\SOFTWARE\LabTech\Service -Name "Version" | Select -expandproperty Version
	if([DateTime]$cwaRegInfo.LastSuccessStatus -lt (Get-Date).AddDays(-1).Date) {
		$cwaOutput = "Not Checking In"
		if([DateTime]$cwaRegInfo.HeartbeatLastReceived -lt (Get-Date).AddDays(-1).Date) {
			$cwaOutput = "No Heartbeat Received"
			if([DateTime]$cwaRegInfo.HeartbeatLastSent -lt (Get-Date).AddDays(-1).Date) {
				$cwaOutput = "No Heartbeat Sent"
			}
		}
	}
}
$cwaOutput
