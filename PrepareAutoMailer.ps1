#.\testExcel.ps1 -excelLocation C:\temp\out.xlsx -sheetName Sheet1 -dateLocation 52 -proxy http://192.168.3.12:8080
##############################################################
##If behind a proxy please use these commands			 	##
##to set up proxy										 	##
##netsh winhttp import proxy source=ie						##
##or														##
##NetSH WinHTTP Set Proxy proxy-server="ProxyServer:8080"	##
##############################################################

param(
	[ValidateScript({ If($_ -match "^http[s]?:\/\/[0-9A-Za-z-_\.]*:[0-9]{2,5}$"){
			$True
		}else{
			Throw "$_ is not a valid proxy. Please chose something like http://192.168.0.12:3128 or http://myproxy:8080 or leave it empty if you do not use a proxy"
		}
	})]
	[string]$proxy = ""
)

$executionPolicy = Get-ExecutionPolicy
Write-Output "Current ExecutionPolicy is $executionPolicy"

#Unrestricted | RemoteSigned | AllSigned | Restricted | Default | Bypass | Undefined

if ($executionPolicy.ToString() -eq "Restricted" -Or $executionPolicy.ToString() -eq "Default" ){
	Write-Output "Setting execution policy ..."
	Set-ExecutionPolicy RemoteSigned
}

if (Get-Module -ListAvailable -Name ImportExcel){
	
	Write-Output "Module ImportExcel already installed."

}else{
	
	Write-Output "Installing module ImportExcel ..."
	
	$version = $PSVersionTable.PSVersion.Major
	
	Write-Output $PSVersionTable.PSVersion
	
	if($version>=5){
		Install-Module ImportExcel
	}else{
		iex (new-object System.Net.WebClient).DownloadString('https://raw.github.com/dfinke/ImportExcel/master/Install.ps1')
	}
}

#store the proxy so that we can reset it after we finish
$proxyBefore=""

if($proxy -ne ""){
	$proxyBefore = ((Netsh WinHTTP Show Proxy) -like "*Proxy Server(s)*")
	
	if($proxyBefore -ne $null){
		$proxyBefore=$proxyBefore.Trim().Split(" ")[4]
	}

	NetSH WinHTTP Set Proxy proxy-server="$proxy"
}

Import-Module ImportExcel

#reset the proxy to the value it had before
if ($proxy){
	if($proxyBefore.length -eq 0){
		Netsh WinHTTP reset proxy
	}else{		
		Netsh WinHTTP Set Proxy proxy-server="$proxyBefore"
	}
}






