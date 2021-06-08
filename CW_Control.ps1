<#
.SYNOPSIS
	Installs CW Control.

.DESCRIPTION
	Downloads the CW Control msi to C:\Windows\Temp\CW_Control.msi then silently installs the application.
	The end user will not know this program has installed.

.LINK
  https://join.corcystems.com

.PARAMETER Company
  This is the Company field in CW Control this agent will install as. Can be left blank (default).
  
.PARAMETER Site
  This is the Site field in CW Control this agent will install as. Can be left blank (default).
  
.PARAMETER Comments
  This is the Comments field in CW Control this agent will install as. Can be left blank (default).
  
.PARAMETER DeviceType
  This is the DeviceType field in CW Control this agent will install as. Can be left blank (default).
  
.EXAMPLE
  $Company = "ACME Company"
  $Site = "Main"
  $Comments = "Test Computer"
  $DeviceType = "Jumpbox"
  $cwControlScript = (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/corcystems/Corcystems-Public/master/CW_Control.ps1')
  Invoke-Expression $cwControlScript

.NOTES
	Version:        1.0
  Author:         Micahel Hauser
	Creation Date:  07/03/2021
  
#>

#Check to make sure PS is version 3.
if (-not ($PSVersionTable))
{Write-Warning 'Powershell 1 Detected. PowerShell Version 3.0 or higher is required.';return}
elseif ($PSVersionTable.PSVersion.Major -eq 2 )
{Write-Warning 'Powershell 2 Detected. PowerShell Version 3.0 or higher is required.';return}
elseif ($PSVersionTable.PSVersion.Major -lt 3 )
{Write-Warning 'Powershell 3+ Not Detected. PowerShell Version 3.0 or higher is required.';return}

#Ignore SSL errors
If ($Null -eq ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type) {
    Add-Type -Debug:$False @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
}
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
#Enable TLS, TLS1.1, TLS1.2, TLS1.3 in this session if they are available
IF([Net.SecurityProtocolType]::Tls) {[Net.ServicePointManager]::SecurityProtocol=[Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls}
IF([Net.SecurityProtocolType]::Tls11) {[Net.ServicePointManager]::SecurityProtocol=[Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls11}
IF([Net.SecurityProtocolType]::Tls12) {[Net.ServicePointManager]::SecurityProtocol=[Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12}
IF([Net.SecurityProtocolType]::Tls13) {[Net.ServicePointManager]::SecurityProtocol=[Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls13}




#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#Per use variables modify as needed
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Param (
	[string] $Company,
	[string] $Site,
	[string] $Comments,
	[string] $DeviceType
)

#Set Nulls to ""
If ($Company -eq $null) {$Company = ""}
If ($Site -eq $null) {$Site = ""}
If ($Comments -eq $null) {$Comments = ""}
If ($DeviceType -eq $null) {$DeviceType = ""}

$WriteOutput = $True



#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#Constants do not change
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#Set URL
$baseURL = "https://join.corcystems.com/Bin/ConnectWiseControl.ClientSetup.msi?h=join.corcystems.com&p=8041&k=BgIAAACkAABSU0ExAAgAAAEAAQAlC9ZHys7DODPwf6K1PP7iY7cwNlfB%2FUzS7ueE3FBLC2llkPpWeHUpL3GXT4QSZo1mRT2CjO8im748tHNnt28d%2F6QpWlcX5rC20AWIEPWZv1brdMrSKMssu91un61I6TkVxrFJoWRJn7JgLY7JDNAmLBz7o%2Fw4brnBY5PbTbrXARArXalsGmfPhllXNauWnUi58toI5s%2FXo%2BeZpix8xv0yW9q6i3JxyfN2TexoLE3dv40Xr2RVDheWe7BNMqGqSUZIxhrfk6fEop3N%2FkqjO17gKLWqi5NhTkopJixK2JE9IMwCU8Non5fW40WQcuHFuQinqtsa9n2XwZoLNPh5PAfS&e=Access&y=Guest&t="

$urlCompany = [uri]::EscapeDataString($Company)
$urlSite = [uri]::EscapeDataString($Site)
$urlComments = [uri]::EscapeDataString($Comments)
$urlDeviceType = [uri]::EscapeDataString($DeviceType)

$cwcURL = $baseURL  + "&c="  + $urlCompany  + "&c="  + $urlSite  + "&c="  + $urlComments  + "&c="  + $urlDeviceType + "&c=&c=&c=&c="


#Set File Location and CW Control Sevice name.
$fileDest = 'C:\Windows\Temp\CW_Control.msi'
$cwcService = 'ScreenConnect Client (8d6cd6b3656cd6f5)'



#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#Check if CW Control is installed and running. Start if not running and install if not installed.
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#Check to see if CW Control is already installed and running.
if (Get-Service $cwcService -ErrorAction SilentlyContinue) {
	#CW Control is already installed.
	#Check to see if the service is running.
    if ((Get-Service $cwcService).Status -eq 'Running')	{
		Write-Host 'CW Control is already installed and running.'
		return
	} else {
		#Service is not running. Start the service.
		Start-Service $cwcService
		Write-Host 'CW Control is already installed but the service was not running. Attempted to start the service.'
		return
	}
} else {
	#CW Control is not installed.
	#Test to make sure the file does not already exsist.
	if (-not(Test-Path -Path $fileDest -PathType Leaf))	{
		try	{
			#Download file then install CW Control.
			Invoke-WebRequest -Uri $cwcURL -OutFile $fileDest; MsiExec.exe /i $fileDest /qn
		} catch	{
			throw $_.Exception.Message
		}
	} else {
		#If the file is already downloaded just run the installer.
		try {
			#Attempt to install CW Control.
			MsiExec.exe /i $fileDest /qn
		} catch {
			throw $_.Exception.Message
		}
	}
}
