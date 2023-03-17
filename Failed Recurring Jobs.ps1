Import-Module JAMS
function Get-SMTicket($EntryID) {
	$dbServerName = "SQL2"
	$sqlStatement = "select top 1 [SEQUENCE] as incidentnum
  FROM [Magic].[_SMDBA_].[_TELMASTE_]
  where [DATE OPEN] > dateadd(hour, -4, sysdatetime())
    and [DESCRIPTION] like '%|replaceme|%'
  order by [DATE OPEN] desc"
	
	$Script:magicTicketNum = ''
	
	# Replace the "|replaceme|" placeholder in the sqlStatement with the 
	# details you want to search for. Put a "%" for spaces
	$searchtext = "JAMS Entry: $EntryID"
	$sqlStatement = $sqlStatement.Replace("|replaceme|", $searchtext)
	
	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$SqlConnection.ConnectionString = "Server = $dbServerName;trusted_connection=true;"
	
	$SqlConnection.Open()
	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
	$SqlCmd.CommandText = $sqlStatement
	$SqlCmd.Connection = $SqlConnection
	
	$Reader = $SqlCmd.ExecuteReader()
	# This returns just the one row which is all we should get back
	while ($Reader.Read()) {
		$Script:magicTicketNum = $Reader["incidentnum"]
	}
	
	$SqlConnection.Close()
	
	# Do something now with the magic ticket
}

$jobobject = @()
New-PSDrive -Name JD -PSProvider JAMS -Root localhost
$Failedjobs = Get-JAMSEntry | Where { $_.FinalSeverity -eq 'Error' -or $_.FinalSeverity -eq 'Warning' -or $_.FinalSeverity -eq 'Fatal' -and ($_.Completiontime -le (get-date).AddHours(-1)) }

Foreach ($Job in $Failedjobs) {
	$Recurring = Get-ChildItem -Path JD: -Recurse | where { $_.Name -like $($Job.Name) -and $_.ResubmitDelay -notlike "" }
	Get-SMTicket -EntryID $Job.Entry
	If ($Recurring.Alerts | where { $_.AlertName -eq 'EntryFailedAlert-Critical' }) {
		$AlertType = 'Critical'
	}
	ElseIf ($Recurring.Alerts | where { $_.AlertName -eq 'EntryFailedAlert-Major' }) {
		$AlertType = 'Major'
	}
	If ($Recurring) {
		$jobobject += [pscustomobject]@{
			JobName = $Job.JobName
			EntryID = $Job.JAMSEntry
			Folder = $Recurring.QualifiedName
			StartTime = $Job.StartTime
			CompletionTime = $Job.CompletionTime
			FinalSeverity = $Job.FinalSeverity
			AlertType = $AlertType
			MagicTicket = $Script:magicTicketNum
		}
	}
}

Remove-PSDrive JD

Write-Output $jobobject

If ($jobobject) {
	$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
	$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
	$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
	$style = $style + "TD{border: 1px solid black; padding: 5px; }"
	$style = $style + "</style>"
	
	$body = $jobobject | ConvertTo-Html -Head $style
	
	$mailparamshtml = @{
		SMTPServer = "smtpmail.acuitybrands.com"
		From = "JAMS@acuitybrands.com"
		To = "EnterpriseEngineering@AcuityBrands.com","JAMSNotification@AcuityBrands.com"
		Subject = "JAMS Recurring Job Failure Report - Recurring jobs that have been failed for more than an hour"
		Body = ($Body | Out-String)
		BodyAsHtml = $true
	}
	
	Send-MailMessage @mailparamshtml
}
