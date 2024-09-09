## Created by Mike Hauser 
## Version 0.1


## Get Hash
$wc = [System.Net.WebClient]::new()
$pkgurl = 'https://raw.githubusercontent.com/corcystems/Corcystems-Public/master/CorTools/Cor_Script.ps1'
$localHash = Get-FileHash C:\Cor\Cor_Script.ps1 -Algorithm SHA256
$FileHash = Get-FileHash -InputStream ($wc.OpenRead($pkgurl))
$FileHash.Hash -eq $localHash
