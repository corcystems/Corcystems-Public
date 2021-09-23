# Grab all users in C:\users and save to C:\users.csv
	Write-Host -ForegroundColor Green "Getting the list of users"
	Write-Host -ForegroundColor Green "Exporting the list of users to c:\users.csv"
	dir C:\Users | select Name | Export-Csv -Path C:\users.csv -NoTypeInformation
	$list = Test-Path C:\users.csv
	Write-Host -ForegroundColor Green "Starting Script..."

# Check for C:\users.csv. Exit if not found.
if ($list) {
	# Clear Mozilla Firefox Cache
		Write-Host -ForegroundColor Green "Clearing Mozilla Firefox Caches"
		Write-Host -ForegroundColor cyan
		Import-CSV -Path C:\users.csv -Header Name | foreach {
			Remove-Item -path C:\Users\$($_.Name)\AppData\Local\Mozilla\Firefox\Profiles\*.default\cache\* -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path C:\Users\$($_.Name)\AppData\Local\Mozilla\Firefox\Profiles\*.default\cache\*.* -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path C:\Users\$($_.Name)\AppData\Local\Mozilla\Firefox\Profiles\*.default\cache2\entries\*.* -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path C:\Users\$($_.Name)\AppData\Local\Mozilla\Firefox\Profiles\*.default\thumbnails\* -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path C:\Users\$($_.Name)\AppData\Local\Mozilla\Firefox\Profiles\*.default\cookies.sqlite -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path C:\Users\$($_.Name)\AppData\Local\Mozilla\Firefox\Profiles\*.default\webappsstore.sqlite -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path C:\Users\$($_.Name)\AppData\Local\Mozilla\Firefox\Profiles\*.default\chromeappsstore.sqlite -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			}
 		Write-Host -ForegroundColor Green "Firefox Done..."

	# Clear Google Chrome 
		Write-Host -ForegroundColor Green "Clearing Google Chrome Caches"
		Write-Host -ForegroundColor cyan
		Import-CSV -Path C:\users.csv -Header Name | foreach {
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Google\Chrome\User Data\Default\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Google\Chrome\User Data\Default\Cache2\entries\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Google\Chrome\User Data\Default\Cookies" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Google\Chrome\User Data\Default\Media Cache" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Google\Chrome\User Data\Default\Cookies-Journal" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			}
		Write-Host -ForegroundColor Green "Chrome Done..."

	# Clear Internet Explorer
		Write-Host -ForegroundColor Green "Clearing Internet Explorer Caches"
		Write-Host -ForegroundColor cyan
		Import-CSV -Path C:\users.csv | foreach {
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Microsoft\Windows\Temporary Internet Files\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Microsoft\Windows\WER\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\Users\$($_.Name)\AppData\Local\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			Remove-Item -path "C:\`$recycle.bin\" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
			}
		Write-Host -ForegroundColor Green "Internet Explorer Done..."

	# All browsers cleared
	Write-Host -ForegroundColor Green "All Tasks Done!"

	Exit

	} else {

	# C:\users.csv not found, exit sctipt.
	Write-Host -ForegroundColor Yellow "C:\users.csv not found, script exit."

	Exit
}