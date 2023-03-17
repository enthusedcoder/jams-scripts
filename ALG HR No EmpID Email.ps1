<#

DESCRIPTION

This script would scan all active directory users in a given organizational unit and would email HR with a list of users that did not have the employee id attribute set

#>

$env:ADPS_LoadDefaultDrive = 0
Import-Module ActiveDirectory

$ADOUDN = "OU=ALG Users,DC=AcuityLightingGroup,DC=com"
$ABADABLEmployeePattern = "^([0]{6})|(m[0]{6})|(ca[0]{6})|(sp[0]{6})|(uk[0]{6})$"
$filepath = "C:\temp\HRCleanup.csv"

$SmtpServer = "smtpmail.acuitylightinggroup.com"
$From = "ABL AD Exceptions<ABLexceptions@acuitylightinggroup.com>"
$Title = "ABL AD Employee ID not set"
$ToAddress = "patrick.angus@acuitybrands.com","kali.mayer@acuitybrands.com"

$EMPLIST = Get-ADUSER -SearchBase $ADOUDN -SearchScope Subtree -Filter { Enabled -eq $true } -Properties EmployeeID,whenCreated,mail,physicalDeliveryOfficeName,passwordlastset,lastlogontimestamp | Where-Object { $_.EmployeeID -match $ABADABLEmployeePattern -and $_.physicalDeliveryOfficeName -notin ("CORPORATE")} 

$EMPLIST | Select-Object Name,SurName,GivenName,whenCreated,samAccountName,UserPrincipalName,EmployeeID,mail,physicalDeliveryOfficeName,passwordlastset,@{l='LastLogon';e={[datetime]::FromFileTime($_.lastlogontimestamp)}} |  
Export-Csv -NoTypeInformation -Path $filepath

#region Create Style
            $header = "<style>"
            $header = $header + "H1{border-width: 1;border: solid;text-align:center}"
            #$header = $header + "H2{color:red;}"
            #$header = $header + "TABLE{border-width: 2px;border-style: solid;border-color: black;border-collapse: collapse;}"
            #$header = $header + "TH{border-width: 1px;padding: 1px;border-style: solid;border-color: black;}"
            #$header = $header + "TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;}"
            $header = $header + "</style>"
            #endregion


$body  = ConvertTo-Html -Body "<h1>Report for Missing Employee IDs</h1>Employee AD accounts with 000000 for EmployeeID: $($EMPLIST.Count)" -Title "Report for Missing Employee IDs" -Head $header | Out-String 

send-MailMessage -SmtpServer $SmtpServer -From $From -To $ToAddress -Subject $Title -BodyAsHtml -Body $body -Attachments $filepath
