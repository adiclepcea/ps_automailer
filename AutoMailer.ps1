param(
	[Parameter(Mandatory=$True)]
	[ValidateScript({If (!(Test-Path $_)) {
			Throw "$_ is not a valid file."			
		}elseif(!($_.endswith(".xlsx","CurrentCultureIgnoreCase"))){
			Throw "$_ is not an excel file"
		}else{
			$True
		}
	})]
	[string]$excelFile,
	[Parameter(Mandatory=$True)]
	[string]$sheetName,
	[Parameter(Mandatory=$True)]
	[string]$dateLocation,
	[Parameter(Mandatory=$False)]
	[string]$mail2Location="",
	[Parameter(Mandatory=$True)]
	[string]$mailLocation,
	[Parameter(Mandatory=$True)]
	[string]$nameLocation,
	[Parameter(Mandatory=$False)]
	[string]$daysBefore=50
)

$historyFile = "./history.csv"
$mailSenderAddress = "http://localhost:8080"
$defaultName = "Your Name"
$defaultMail = "name.surname@mailserver.ro"
$defaultPass = "pass"
function ValidMail 
{
	Param($mailAddress)
	try{
		$rez = new-object system.net.mail.mailaddress($mailAddress)
		return $rez
	}catch{
		return 0
	}
	
}

function ValidDate
{
	Param($date)
	$rez = ""
	try{
		$rez = [System.DateTime]::FromOADate($date)
	}catch{
		return 0
	}
	return 1
}

function GetDate 
{
	Param($date)
	
	$rez = ""
	if(!([DateTime]::TryParse($date,[ref]$rez))){
		throw "$date is an invalid date"
	}
	return $rez
}

function WriteMailSent
{
	Param(
		[String]$mail,		
		[System.DateTime]$date
	)
	
	$d = Get-Date
	
	$newLine = New-Object psobject
	
	Add-Member -InputObject $newLine -MemberType NoteProperty -Name mail -Value $mail
	Add-Member -InputObject $newLine -MemberType NoteProperty -Name date -Value $date.ToString()
	Add-Member -InputObject $newLine -MemberType NoteProperty -Name dateSent -Value $d.ToString()
	
	$newCsv = @()
	$newCsv += $newLine
	
	if (Test-Path -Path $historyFile){
		$csv = Import-Csv $historyFile
		
		$csv | ForEach-Object{
			$newCsv += $_
		}	
	}
	
	$newCsv | Export-Csv  -Path $historyFile -NoTypeInformation
}

function MailAllreadySent
{
	
	Param(
		[String]$mail,		
		[System.DateTime]$date
	)
	
	if (Test-Path -Path $historyFile){
		$csv = Import-Csv $historyFile
		$csv | ForEach-Object {
			if ($_.mail -eq $mail -and $_.date -eq  $date.ToString()){
				return 1
			}
		}
	}
	
	return 0
}

function CheckMustSendMail 
{
	Param($when)
	
	$startDate = Get-Date
	$ts = New-Timespan -Start $startDate -End $when
	if ($ts.Days -gt 0 -and ($ts.Days -lt $daysBefore -or $ts.Days -eq $daysBefore)){		
		return 1
	}
	
	return 0
}

function SendMailIfYouHaveTo
{
	Param(
		$mail,
		$name,
		$date
	)

	#Check if we have a valid mail
	if ($mail -ne $null){
		$vm = ValidMail -mailAddress $mail
		if ($vm -eq 0){
			return
		}
		$d = Get-Date
		$mail = $vm.Address
		$contact = $vm.DisplayName
		#Check if mail is due to be sent
		if (CheckMustSendMail($date)){			
			if ((MailAllreadySent -mail $mail -date $date) -eq 0){
				Write-Output "$($d): Sending mail to $($name), $($date), $($mail)"
				$mail = "adiclepcea@gmail.com"
				$req = @{To = @{Name="$($name)";Address="$($mail)"}; `
						Subject="Contract with Hannes"; `
						Body="In attention of $($contact). Your contract with us will expire on $($date). Please contact us for renewal.";`
						From=@{Name="$($defaultName)";Address="$($defaultMail)"};Password="$($defaultPass)"}
				$json = $req | ConvertTo-Json
				$rez = Invoke-WebRequest -uri "$($mailSenderAddress)/sendmail" -Method Post -Body $json
				if ($rez.StatusCode -ne 200){
					if($rez.Content -ne $null){
						Write-Warning "$($d): Sending mail to $($mail) failed because: $($rez.Content)"
					}else{
						Write-Warning "$($d): Sending mail to $($mail) failed because: $($rez)"
					}
				}else{
					WriteMailSent -mail $mail -date $date
				}
			}
		}
		
	}
}

Import-Module ImportExcel

$excel = Import-Excel $excelFile -WorkSheetname $sheetName

$i = 0

@($excel).length

while($excel[$i] -ne $null `
	-and $excel[$i].$dateLocation -ne "" `
	-and $excel[$i].$nameLocation -ne $null `
	-and $excel[$i].$nameLocation.Trim() -ne ""){
	
	$name = $excel[$i].$nameLocation
	$date = $excel[$i].$dateLocation
	$mail = $excel[$i].$mailLocation
	if ($mail2Location -ne ""){
		$mail2 = $excel[$i].$mail2Location
	}
	
	$date=[System.DateTime]::FromOADate($date)
		
	SendMailIfYouHaveTo -mail $mail -name $name -date $date
	if ($mail2Location -ne ""){
		SendMailIfYouHaveTo -mail $mail2 -name $name -date $date
	}
	
	$i = $i+1
}
