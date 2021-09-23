# Grab all users in C:\users and save to C:\users.csv
	Write-Host "Getting the list of users"
	Write-Host "Exporting the list of users to c:\users.csv"
	dir C:\Users | select Name | Export-Csv -Path C:\users.csv -NoTypeInformation
	$list = Test-Path C:\users.csv
	Write-Host "Starting Script..."

# Check for C:\users.csv. Exit if not found.
if ($list) {
	# Clear Mozilla Firefox Cache
		Write-Host "Clearing Mozilla Firefox Caches"
		Import-CSV -Path C:\users.csv -Header Name | foreach {
			Remove-Item -path C:\Users\$($_.Name)\AppData\Local\Mozilla\Firefox\Profiles\*.default\cache\* -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path C:\Users\$($_.Name)\AppData\Local\Mozilla\Firefox\Profiles\*.default\cache\*.* -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path C:\Users\$($_.Name)\AppData\Local\Mozilla\Firefox\Profiles\*.default\cache2\entries\*.* -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path C:\Users\$($_.Name)\AppData\Local\Mozilla\Firefox\Profiles\*.default\thumbnails\* -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path C:\Users\$($_.Name)\AppData\Local\Mozilla\Firefox\Profiles\*.default\cookies.sqlite -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path C:\Users\$($_.Name)\AppData\Local\Mozilla\Firefox\Profiles\*.default\webappsstore.sqlite -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path C:\Users\$($_.Name)\AppData\Local\Mozilla\Firefox\Profiles\*.default\chromeappsstore.sqlite -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			}
 		Write-Host "Firefox Done..."
<#
	# Clear Google Chrome 
		Write-Host "Clearing Google Chrome Caches"
		Import-CSV -Path C:\users.csv -Header Name | foreach {
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Google\Chrome\User Data\Default\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Google\Chrome\User Data\Default\Cache2\entries\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Google\Chrome\User Data\Default\Cookies" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Google\Chrome\User Data\Default\Media Cache" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Google\Chrome\User Data\Default\Cookies-Journal" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			}
		Write-Host "Chrome Done..."

	# Clear Internet Explorer
		Write-Host "Clearing Internet Explorer Caches"
		Import-CSV -Path C:\users.csv | foreach {
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Microsoft\Windows\Temporary Internet Files\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Microsoft\Windows\WER\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\`$recycle.bin\" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			}
		Write-Host "Internet Explorer Done..."

	# All browsers cleared
	Write-Host "All Tasks Done!"
#>
	} else {

	# C:\users.csv not found, exit sctipt.
	Write-Host "C:\users.csv not found, script exit."

	Exit
}
