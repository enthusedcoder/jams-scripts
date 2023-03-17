<#
DESCRIPTION
This jams script would copy office 365 contact objects from one tennant to another

#>

# Sets script home directory
$HomeDir = "C:\TIDALJOBS\GALSync\ABLtoAGTGalSync"

$DeleteCSV = "deleteagtcontacts.csv"
$SendReportByEmail = $true

Set-Location $HomeDir
# Logfile setup information.
$LogFileName = "_ABLtoAGTSyncLog.txt"
$now = Get-Date -format s
$date = $now.substring(0,10)
$FilePath = "$HomeDir\logs\"
$LogFile = $FilePath + $date + $LogFileName

$now = Get-Date -format G
Write-Output "$now - Loading Modules" | Out-File $LogFile -Append

. .\Connect-Function.ps1
Load-ABAutomation | Out-File $LogFile -Append

# Connect to Exchange2013 OnPrem
$OnPremSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://o365exch2.acuitylightinggroup.com/powershell/ -Authentication Kerberos 
Import-PSSession $OnPremSession -Prefix ABL

# Connect to O365 - Corporate
#$UserCredential = Get-Credential
$UserCredential = Get-SecretServerCred 2125
$O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $O365Session -Prefix O365

# Get Office365 Agent Tenant Admin Cred from SecretServer
$O365_AGTAdminCred = Get-SecretServerCred 2061

# Connect to O365 Agent Tenant
$ConnectionUri = "https://outlook.office365.com/powershell-liveid/"
$AGTSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365_AGTAdminCred -Authentication Basic -AllowRedirection
Import-PSSession $AGTSession -Prefix AGT

$ABLArray = @()
$ABLCreateArray = @()
$ABLEditArray = @()
$O365DeleteArray = @()
$ABLCreateError = @()
$ABLEditError = @()
$O365DeleteError = @()
$ErrorActionPreference = "Stop"

# Populate array with groups, users, and shared mailboxes that should be created in AGT O365
$ABLDistGroups = Get-ABLDistributionGroup -ResultSize Unlimited -Filter "
  (Name -notlike '#*') -and
  (Name -notlike '!*') -and
  (Name -notlike 'ABL*') -and
  (Name -notlike 'Agency*') -and
  (Name -notlike 'Agile*') -and
  (Name -notlike '*All*') -and
  (Name -notlike 'C&I*') -and
  (Name -notlike 'Cochran*') -and
  (Name -notlike 'Corporate*') -and
  (Name -notlike 'Database*') -and
  (Name -notlike 'DBA*') -and
  (Name -notlike 'DC*') -and
  (Name -notlike 'Enterprise*') -and
  (Name -notlike 'IST*') -and
  (Name -notlike 'Notices*') -and
  (Name -notlike '*Oracle*') -and
  (Name -notlike 'Sharepoint*') -and
  (Name -notlike 'U.S.*') -and
  (Name -notlike 'WebVPN*') -and
  (Name -notlike 'XOrder*')
  " -WarningAction Silentlycontinue
Foreach ($Group in $ABLDistGroups) {
  If (Get-ABLDistributionGroupMember -ResultSize 1 -Identity $Group.Alias -WarningAction SilentlyContinue -ErrorAction SilentlyContinue) {
    $Alias = "ABLGroup_"+($Group.PrimarySmtpAddress).Split("@")[0]
    $CustomObject = [pscustomobject]@{ 
      'ExternalEmailAddress'=$Group.PrimarySmtpAddress
      'DisplayName'=$Group.DisplayName
      'Alias'=$Alias
    }
    $ABLArray += $CustomObject
  }
}
Get-ABLMailbox -ResultSize Unlimited -Filter { ((RecipientTypeDetails -eq 'UserMailbox') -or (RecipientTypeDetails -eq 'SharedMailbox')) -and (DisplayName -notlike 'x-*') -and (HiddenFromAddressListsEnabled -ne $true) } -WarningAction Silentlycontinue | Where {
  ($_.OrganizationalUnit -notlike '*Cochran*') -and
  ($_.OrganizationalUnit -notlike '*CORPORATE/Shared Mailboxes*') -and
  ($_.OrganizationalUnit -notlike '*Hourly Employees*') -and
  ($_.OrganizationalUnit -notlike '*Monitoring Mailboxes*') -and
  ($_.OrganizationalUnit -notlike '*Neoteris*') -and
  ($_.OrganizationalUnit -notlike '*Temporary and Unknown*') -and
  ($_.OrganizationalUnit -notlike '*Windows Management Groups*') -and
  ($_.OrganizationalUnit -notlike '*z Agents (Do Not Remove)*')
} | Foreach {
  $Alias = "ABLUser_"+($_.PrimarySmtpAddress).Split("@")[0]
  $CustomObject = [pscustomobject]@{ 
    'ExternalEmailAddress'=$_.PrimarySmtpAddress
    'DisplayName'=$_.DisplayName
    'Alias'=$Alias
  }
  $ABLArray += $CustomObject
}

Get-O365Mailbox -ResultSize Unlimited -Filter { ((RecipientTypeDetails -eq 'UserMailbox') -or (RecipientTypeDetails -eq 'SharedMailbox')) -and (DisplayName -notlike 'x-*') -and (HiddenFromAddressListsEnabled -ne $true) } -WarningAction Silentlycontinue | Where {
  ($_.OrganizationalUnit -notlike '*Cochran*') -and
  ($_.OrganizationalUnit -notlike '*CORPORATE/Shared Mailboxes*') -and
  ($_.OrganizationalUnit -notlike '*Hourly Employees*') -and
  ($_.OrganizationalUnit -notlike '*Monitoring Mailboxes*') -and
  ($_.OrganizationalUnit -notlike '*Neoteris*') -and
  ($_.OrganizationalUnit -notlike '*Temporary and Unknown*') -and
  ($_.OrganizationalUnit -notlike '*Windows Management Groups*') -and
  ($_.OrganizationalUnit -notlike '*z Agents (Do Not Remove)*')
} | Foreach {
  $Alias = "ABLUser_"+($_.PrimarySmtpAddress).Split("@")[0]
  $CustomObject = [pscustomobject]@{ 
    'ExternalEmailAddress'=$_.PrimarySmtpAddress
    'DisplayName'=$_.DisplayName
    'Alias'=$Alias
  }
  $ABLArray += $CustomObject
}

# Created in the event that distribution groups with same name already exist.
# Will create contact in outlook online that can be seen by selecting
# "All Contacts" in agent address book
# Caution needs to be exercised when running below command due to new
# microsoft o365 implementation "groups in outlook".  These objects are
# similar to distribution groups, are listed in the same category as
# distribution groups, and are also returned as objects when the 
# get-distributionlist command is run, but will generate an error 
# when the "group in outlook" object is reached because the cmdlet
# doesn't support the object.
Get-o365distributiongroup -Identity "ATS*" | ForEach-Object {
$Alias = $_.DisplayName
  $CustomObject = [pscustomobject]@{ 
    'ExternalEmailAddress'=$_.PrimarySmtpAddress
    'DisplayName'=$_.DisplayName
    'Alias'=$Alias
  }
  $ABLArray += $CustomObject
}

# Get ABL contacts from O365
$O365Array = Get-AGTMailContact -Identity ABL* -ResultSize Unlimited | Select Alias,DisplayName,ExternalEmailAddress,Identity

# Populate Create\Edit Array
$O365Alias = $O365Array | % {($_.Alias).tostring()}
$O365ExternalEmailAddress = $O365Array | % {($_.ExternalEmailAddress).Split(":")[1]}
$O365DisplayName = $O365Array | % {($_.DisplayName).tostring()}

Foreach ($ablobj in $ABLArray) {
  If (!($O365ExternalEmailAddress -contains $ablobj.ExternalEmailAddress)){
    If (!($O365Alias -contains $ablobj.Alias)){
      $ABLCreateArray += $ablobj
    } Else {
      $ABLEditArray += $ablobj
    }
  } ElseIf (!($O365Alias -contains $ablobj.Alias)){
      $ABLEditArray += $ablobj
  } ElseIf (!($O365DisplayName -contains $ablobj.DisplayName)) {
      $ABLEditArray += $ablobj
  }
}

# Populate Delete Array
$ABLExternalEmailAddress = $ABLArray | % {($_.ExternalEmailAddress).tostring()}
Foreach ($o365obj in $O365Array) {
  If (!($ABLExternalEmailAddress -contains (($o365obj.ExternalEmailAddress).Split(":")[1]))) {
    $O365DeleteArray += $o365obj
  }
}

# Create O365 Contacts
Foreach ($createobj in $ABLCreateArray) {
  Try {
    New-AGTMailContact -ExternalEmailAddress ($createobj.ExternalEmailAddress) -Name ($createobj.DisplayName) -Alias ($createobj.Alias) -DisplayName ($createobj.DisplayName) -ErrorAction Stop
    $now = Get-Date -format G
    Write-Output "$now --> $($createobj.ExternalEmailAddress) Created" | Out-File $LogFile -Append
  } Catch [System.Management.Automation.RemoteException]{
    $ABLCreateError += $createobj
    $now = Get-Date -format G
    Write-Output "$now --> Errored on Create --> $($createobj.ExternalEmailAddress) - $($createobj.Alias) - $($createobj.DisplayName)" | Out-File $LogFile -Append
  }
}

# Edit O365 Contacts
Foreach ($editobj in $ABLEditArray) {
  Try {
    Set-AGTMailContact -Identity $editobj.ExternalEmailAddress -ExternalEmailAddress ($editobj.ExternalEmailAddress) -Name ($editobj.DisplayName) -Alias ($editobj.Alias) -DisplayName ($editobj.DisplayName) -ErrorAction Stop
    $now = Get-Date -format G
    Write-Output "$now --> $($editobj.ExternalEmailAddress) Modified" | Out-File $LogFile -Append
  } Catch [System.Management.Automation.RemoteException] {
    $ABLEditError += $editobj
    $now = Get-Date -format G
    Write-Output "$now --> Errored on Edit --> $($editobj.ExternalEmailAddress) - $($editobj.Alias) - $($editobj.DisplayName)" | Out-File $LogFile -Append
  }
}

# Delete O365 Contacts
$DeleteFilePath = "$HomeDir\delete\"
$DeleteFile = $DeleteFilePath + $DeleteCSV
$O365DeleteArray | Export-Csv $DeleteFile -Append

Foreach ($deleteobj in $O365DeleteArray) {
  Try {
    Remove-AGTMailContact -identity (($deleteobj.ExternalEmailAddress).Split(":")[1]) -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
    $now = Get-Date -format G
    Write-Output "$now --> $((($deleteobj.ExternalEmailAddress).Split(":")[1])) Removed" | Out-File $LogFile -Append
  } Catch [System.Management.Automation.RemoteException] {
    $O365DeleteError += $deleteobj
    $now = Get-Date -format G
    Write-Output "$now --> Errored on Delete --> $($deleteobj.ExternalEmailAddress) - $($deleteobj.Alias) - $($deleteobj.DisplayName)" | Out-File $LogFile -Append
  }
}

Remove-PSSession -Session $OnPremSession
Remove-PSSession -Session $O365Session
Remove-PSSession -Session $AGTSession

# Email the log file
$mailparams = @{
  SMTPServer = "smtpmail.acuitybrands.com"
  From = "O365Sync@acuitybrands.com"
  To = "AGTtoABLGalSync@AcuityBrands.com"
  Subject = "Acuity to Agent GALSync Log"
  Body = "GALSync results attached"
  Attachment = $Logfile
}

If ($SendReportByEmail) {
  Send-MailMessage @mailparams
}
