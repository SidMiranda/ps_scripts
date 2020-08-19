# Script:	TotalEmailsSentReceivedPerHour.ps1
# Purpose:	Get the number of e-mails sent and received per hour

$From = Get-Date "28/11/2019"
$To = $From.AddHours(1)

[Int64] $intSent = 0
[Int64] $intRec = 0
[String] $strSent = $strTotalSent = $null
[String] $strRec = $strTotalRec = $null


Do
{
	If ($From.Hour -eq "0") {
		$strTotalSent += "$strSent`n"
		$strTotalRec += "$strRec`n"

		$strSent = "$($From.DayOfWeek),$($From.ToShortDateString()),"
		$strRec = "$($From.DayOfWeek),$($From.ToShortDateString()),"
		
		Write-Host "Searching $From"
	}
	
	
	# It is faster to search the Transport Logs this way then by doing a Get-TransportServer | Get-MessageTrackingLog -ResultSize Unlimited -Start $From -End $To and then checking if it is a sent or received e-mail 
	# Sent E-mails
	$intSent = (Get-TransportService | Get-MessageTrackingLog -ResultSize Unlimited -EventId RECEIVE -Start $From -End $To | Where {$_.Source -eq "STOREDRIVER" -and $_.MessageSubject -ne "Folder Content" -and $_.Sender -notlike "HealthMailbox*" -and $_.Sender -notlike "maildeliveryprobe*" -and $_.Sender -notlike "inboundproxy*" -and $_.Recipients -notmatch "journal@domain.com"}).Count
	$strSent += "$intSent,"
	
	# Received E-mails
	Get-TransportService | Get-MessageTrackingLog -ResultSize Unlimited -EventId DELIVER -Start $From -End $To | Where {$_.MessageSubject -ne "Folder Content" -and $_.Sender -notlike "HealthMailbox*" -and $_.Sender -notlike "maildeliveryprobe*" -and $_.Sender -notlike "inboundproxy*" -and $_.Recipients -notmatch "journal@domain.com"} | ForEach {$intRec += $_.RecipientCount}
	$strRec += "$intRec,"
	$intRec = 0

	$From = $From.AddHours(1)
	$To = $From.AddHours(1)
}
While ($To -lt (Get-Date))

# Print all the results
$strTotalSent += "$strSent`n"
$strTotalRec += "$strRec`n"

Write-Host "`nSent" -ForegroundColor Yellow
Write-Host "DayOfWeek,Date,00:00,01:00,02:00,03:00,04:00,05:00,06:00,07:00,08:00,09:00,10:00,11:00,12:00,13:00,14:00,15:00,16:00,17:00,18:00,19:00,20:00,21:00,22:00,23:00" -ForegroundColor Yellow
$strTotalSent

Write-Host "`nReceived" -ForegroundColor Yellow
Write-Host "DayOfWeek,Date,00:00,01:00,02:00,03:00,04:00,05:00,06:00,07:00,08:00,09:00,10:00,11:00,12:00,13:00,14:00,15:00,16:00,17:00,18:00,19:00,20:00,21:00,22:00,23:00" -ForegroundColor Yellow
$strTotalRec
