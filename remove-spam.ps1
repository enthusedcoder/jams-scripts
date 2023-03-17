Import-module "C:\Program Files\WindowsPowerShell\Modules\SecretServer\1.0\SecretServer.psm1"
$ssconnect = New-SSConnection -UpdateSecretConfig $True -Uri "http://secretserver/winauthwebservices/sswinauthwebservice.asmx"
$cred2 = (Get-Secret -As Credential -SecretId 2125).Credential
# connect to office 365 exchange
$connect = Connect-MsolService -Credential $cred2
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred2 -Authentication Basic -AllowRedirection
Import-PSSession $Session
#imports csv files.
$files = Get-ChildItem "\\cdcsrvr1\Depts\Engineering Solutions\Spam_script\*.csv"
foreach ($file in $files)
{
	$csv = Import-Csv $file.FullName -Delimiter ','
	#removes spam using the subject and sender address specified in csv object to identify the message from the mailbox which is also identified in the csv object.
	foreach ($cs in $csv)
	{
		$date = [datetime]$($cs.Date).Substring(0,10)
		
		If ($cs.Subject -eq '')
		{
			Try
			{
				Search-Mailbox -Identity $cs.Recipients -SearchQuery "From:$($cs.Sender) Received:$($date.Month)/$($date.Day)/$($date.Year)" -TargetMailbox "ISTEmailReview@AcuityBrands.com" -TargetFolder "Inbox" -SearchDumpster -DeleteContent -Force -Confirm:$false -erroraction stop
			}
			Catch
			{
				
			}
		}
		Else
		{
			Try
			{
				Search-Mailbox -Identity $cs.Recipients -SearchQuery "From:$($cs.Sender) Subject:`"$($cs.Subject)`" Received:$($date.Month)/$($date.Day)/$($date.Year)" -TargetMailbox "ISTEmailReview@AcuityBrands.com" -TargetFolder "Inbox" -SearchDumpster -DeleteContent -Force -Confirm:$false -erroraction stop
			}
			Catch
			{
				
			}
		}
	}
}
Get-PSSession | Remove-PSSession
