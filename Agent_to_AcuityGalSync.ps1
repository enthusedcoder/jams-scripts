# To create, update, or delete All users, set the following values to $True.  $True will override the Threshold settings.
$CreateAll = $true
$UpdateAll = $true
$DeleteAll = $true

# To create, update, or delete a specific number of users, change the following values
$CreateThreshold = 0
$UpdateThreshold = 0
$DeleteThreshold = 0

$SendReportByEmail = $true
$SendHTMLExceptReportByEmail = $true

# Sets script home directory
$HomeDir = "C:\TIDALJOBS\GALSync\AGTtoABLGalSync"
cd $HomeDir

# Sets logfile name
$LogFileName = "_AGTtoABLSyncLog.txt"

# To create missing OU's in ABL zAgents, set to $True
$OUCreateBool = $true

# OU for agent contacts in ABLAD
$RootOU = "OU=z Agents (Do Not Remove),DC=AcuityLightingGroup,DC=com"
$OUBase = "acuitylightinggroup.com/z Agents (Do Not Remove)/"


# Logfile cleanup and setup 
$now = Get-Date -format s
$date = $now.substring(0, 10)
$FilePath = "$HomeDir\logs\"
Get-ChildItem $Filepath | where { $_.LastWriteTime -le ((Get-Date).AddDays(-15)) } | Remove-Item -Confirm:$false
$LogFile = $FilePath + $date + $LogFileName

$now = Get-Date -format G
Write-Output "$now - Loading Modules" | Out-File $LogFile -Append

Import-Module ABAutomation
Import-Module MSOnline
Import-Module ActiveDirectory
Import-Module EnhancedHTML2

$OnPremSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://o365exch2.acuitylightinggroup.com/powershell/ -Authentication Kerberos
Import-PSSession $OnPremSession -Prefix ABL

# Get Office365 Agent Tenant Admin Cred from SecretServer
$O365_AGTAdminCred = Get-SecretServerCred 2061

# Connect to MSOLService
Try {
	Connect-MsolService -Credential $O365_AGTAdminCred
}
Catch {
	Write-Output $_ | Out-File $LogFile -Append
}

$now = Get-Date -format G
Write-Output "$now - Starting O365 AGT GALSync" | Out-File $LogFile -Append

# Define arrays to hold the objects that need created, updated or deleted.
$CreateArray = @()	# users that need to be created go in this one
$UpdateArray = @()	# users that need to be updated go in this one
$DeleteArray = @()	# users that need to be deleted go in this one
$OUCreateArray = @() # Depts that need to be created in ABLAD
$IncludeOUArray = @() # OU's to include in GAL sync

# Export BPOS users, Department must be set, and no numbers in the Displayname
Try {
	$MsolUserList = Get-MsolUser -All -EnabledFilter EnabledOnly | where { ($_.Department -like "AGT*") -and ($_.Department -notlike "AGTXXX") -and ($_.DisplayName -notmatch "\d" -and $_.DisplayName -like "*, *") }
}
Catch {
	$now = Get-Date -format G
	Write-Output "$now --> " + $_ | Out-File $LogFile -Append
}

$MSOLContacts = @()
Try {
	$ObjectIDs = Get-MsolContact -All | select ObjectID
	foreach ($objid in $ObjectIDs) {
		$AGTCntc = Get-MSOLContact -ObjectId $($objid.ObjectId.ToString()) -ErrorAction SilentlyContinue
		If ($($AGTCntc.Department) -like 'AGT*') {
			$MSOLContacts += $AGTCntc
		}
	}
}
Catch {
	$now = Get-Date -format G
	Write-Output "$now --> " + $_ | Out-File $LogFile -Append
}

$MsolUserList += $MSOLContacts

$MsolUserListSize = $MsolUserList.Length + $MSOLContacts.Length
$now = Get-Date -format G
Write-Output "$now - User list exported.  $MsolUserListSize users exported." | Out-File $LogFile -Append

######################################################################
# Begin OU Compare\Create section
# Get current O365 departments

Write-Output "*****Begin OU Compare\Create Section*****" | Out-File $LogFile -Append
$O365Depts = $MsolUserList.Department | Sort-Object -Unique

# Evaluate if OU exists
Foreach ($OU in $O365Depts) {
	$ABLOU = ""
	$DN = 'OU=' + $OU + ',' + $RootOU
	Try {
		$ABLOU = Get-ADOrganizationalUnit -SearchBase $DN -Filter * -ErrorAction Stop
	}
	Catch {
		$now = Get-Date -format G
		"$now - $DN Produced an error" + $_ | Out-File $LogFile -Append
	}
	If (!$ABLOU) {
		If ($OU -notmatch '(AGT)\d{3}') {
			$now = Get-Date -format G
			Write-Output "$now - $OU is not in the correct format and will not be created" | Out-File $LogFile -Append
		}
		Else {
			$OUCreateArray += $OU
			$name = $OU
			$now = Get-Date -format G
			Write-Output "$now - $name does not exist and needs to be created" | Out-File $LogFile -Append
		}
	}
}

# Create ABLOU if OUCreateBool is set to $True
If ($OUCreateArray) {
	If ($OUCreateBool) {
		Foreach ($OU in $OUCreateArray) {
			New-ADOrganizationalUnit -Name $OU -Path "OU=z Agents (Do Not Remove),DC=AcuityLightingGroup,DC=com"
			$name = $OU
			$now = Get-Date -format G
			Write-Output "$now - $name has been created in ABL AD" | Out-File $LogFile -Append
		}
	}
}

Write-Output "*****End OU Compare\Create Section*****" | Out-File $LogFile -Append

# This is the end of the OU Compare section
######################################################################
# This is the beginning of the Create\Update section

Write-Output "*****Begin to Populate Create\Update Array Section*****" | Out-File $LogFile -Append

$BadOU = $MsolUserList | Where { $_.Department -notmatch '(AGT)\d{3}' }
$HTMLOUObj = @()
Foreach ($User in $BadOU) {
	If ($User -Is [Microsoft.Online.Administration.User]) {
		$props = @{
			'DisplayName' = $User.DisplayName;
			'UserPrincipalName' = $User.UserPrincipalName;
			'Department' = $User.Department
		}
	}
	If ($User -Is [Microsoft.Online.Administration.Contact]) {
		$props = @{
			'DisplayName' = $User.DisplayName;
			'UserPrincipalName' = $User.EmailAddress;
			'Department' = $User.Department
		}
	}
	$HTMLOUObj += New-Object -TypeName PSObject -Property $props
	$now = Get-Date -format G
	Write-Output "$now -  $User.Department needs to be fixed for $User.DisplayName"
}

$GoodOU = $MsolUserList | Where { $_.Department -match '(AGT)\d{3}' }
Foreach ($User in $GoodOU) {
	If ($User -Is [Microsoft.Online.Administration.User]) {
		$UserUPN = $User.UserPrincipalName
	}
	If ($User -Is [Microsoft.Online.Administration.Contact]) {
		$UserUPN = $User.EmailAddress
	}
	$UserDispName = $User.DisplayName
	$UserDept = $User.Department
	$UserOU = $OUBase + $UserDept
	
	# Test if the UserPrincipalName exists in the local exchange as a PrimarySmtpAddress
	Try {
		$MailContact = Get-ABLMailContact -Filter "PrimarySmtpAddress -eq '$UserUPN'"  -OrganizationalUnit $UserOU -ErrorAction stop
	}
	Catch {
		$now = Get-Date -format G
		"$now - $UserUPN doesn't match a ABLAD user PrimarySmtpAddress" + $_ | Out-File $LogFile -Append
	}
	If (!$MailContact) {
		Try {
			$MailContact = Get-ABLMailContact -Filter "DisplayName -eq '$UserDispName'" -OrganizationalUnit $UserOU -ErrorAction stop
		}
		Catch {
			$now = Get-Date -format G
			"$now - $UserDispName doesn't match a ABLAD user DisplayName" + $_ | Out-File $LogFile -Append
		}
	}
	
	# If no, then push to the $CreateArray
	if (!$MailContact.Identity) {
		$CreateArray += $User
		$name = $UserDispName
		$now = Get-Date -format G
		Write-Output "$now - Adding $name to the CreateArray" | Out-File $LogFile -Append
	}
	
	# If yes, then compare the O365 fields to the local exchange and AD fields
	Else {
		# Set a variable to key this section on
		$UpdateBool = $FALSE
		
		# Get the AD object, as the First and Last name is only stored there.
		$ObjGUID = ($MailContact.Guid).Guid
		Try {
			$ADContact = Get-ADObject -Filter "ObjectGUID -eq '$ObjGUID'" -SearchBase $RootOU -Properties DisplayName, givenName, sn -ErrorAction Stop
		}
		Catch {
			$now = Get-Date -format G
			Write-Output "$now - $MailContact.DisplayName with $ObjGUID created an error --> $_" | Out-File $LogFile -Append
		}
		
		# Add GUID to $User object for use in the update or delete array, to make sure we have the correct ADObject
		Add-Member -MemberType NoteProperty -Name ABLGuid -Value $ObjGUID -InputObject $User -Force
		
		# If any do not match, push to the $UpdateArray
		If ($User.Firstname -ne $ADContact.givenName) {
			$UpdateBool = $TRUE
			$name = $UserDispName
			$oldname = $ADContact.givenName
			$newname = $User.Firstname
			$now = Get-Date -format G
			Write-Output "$now - $name - First Name has changed from $oldname to $newname" | Out-File $LogFile -Append
		}
		
		If ($User.LastName -ne $ADContact.sn) {
			$UpdateBool = $TRUE
			$name = $UserDispName
			$oldname = $ADContact.sn
			$newname = $User.LastName
			$now = Get-Date -format G
			Write-Output "$now - $name - Last Name has changed from $oldname to $newname" | Out-File $LogFile -Append
		}
		If ($User.DisplayName -ne $MailContact.DisplayName) {
			$UpdateBool = $TRUE
			$name = $UserDispName
			$oldname = $MailContact.DisplayName
			$newname = $User.DisplayName
			$now = Get-Date -format G
			Write-Output "$now - $name - Display Name has changed from $oldname to $newname" | Out-File $LogFile -Append
		}
		If ($User -Is [Microsoft.Online.Administration.User]) {
			If ($User.UserPrincipalName -ne $MailContact.PrimarySmtpAddress) {
				$UpdateBool = $TRUE
				$name = $UserDispName
				$oldname = $MailContact.PrimarySmtpAddress
				$newname = $User.UserPrincipalName
				$now = Get-Date -format G
				Write-Output "$now - $name - Email address has changed from $oldname to $newname" | Out-File $LogFile -Append
			}
		}
		If ($User -Is [Microsoft.Online.Administration.Contact]) {
			If ($User.EmailAddress -ne $MailContact.PrimarySmtpAddress) {
				$UpdateBool = $TRUE
				$name = $UserDispName
				$oldname = $MailContact.PrimarySmtpAddress
				$newname = $User.EmailAddress
				$now = Get-Date -format G
				Write-Output "$now - $name - Email address has changed from $oldname to $newname" | Out-File $LogFile -Append
			}
		}
		If ($UpdateBool -eq $TRUE) {
			$UpdateArray += $User
			$name = $UserDispName
			$now = Get-Date -format G
			Write-Output "$now - $name needs to be updated in the local Exchange environment" | Out-File $LogFile -Append
		}
	}  #end the compare loop	
} # End user create\update loop

$now = Get-Date -format G
$CreateCount = $CreateArray.Count
$UpdateCount = $UpdateArray.Count

Write-Output "*****$now - $CreateCount users need to be created in ABLAD (zAgents).*****" | Out-File $LogFile -Append
Write-Output "*****$now - $UpdateCount users need to be updated in ABLAD (zAgents).*****" | Out-File $LogFile -Append
Write-Output "*****End Create\Update Array Section*****" | Out-File $LogFile -Append

######################################################################

Write-Output "*****Begin Create\Update Users Section*****" | Out-File $LogFile -Append
# Check if $CreateArray is populated
If ($CreateArray.length -eq 0) {
	$now = Get-Date -format G
	Write-Output "$now - No users to create" | Out-File $LogFile -Append
}
# Create users based on $CreateThreshold setting
Else {
	If ($CreateAll) {
		$CreateThreshold = $CreateArray.Count
		$size = $CreateArray.Count
		$now = Get-Date -format G
		Write-Output "$now -  $CreateThreshold of $size users will be created.  Beginning mail contact creation" | Out-File $LogFile -Append
	}
	Else {
		If ($CreateThreshold -gt 0) {
			$size = $CreateArray.Count
			$now = Get-Date -format G
			Write-Output "$now -  $CreateThreshold of $size users will be created.  Beginning mail contact creation" | Out-File $LogFile -Append
		}
	}
	
	$CreateUsers = $CreateArray | Select -First $CreateThreshold
	Foreach ($User in $CreateUsers) {
		#$SMTPAddress = $User.UserPrincipalName
		If ($User -Is [Microsoft.Online.Administration.User]) {
			$SMTPAddress = $User.UserPrincipalName
		}
		If ($User -Is [Microsoft.Online.Administration.Contact]) {
			$SMTPAddress = $User.EmailAddress
		}
		
		# Create alias in the following format:  firstname.lastname_company
		$SMTPDomain = ($SMTPAddress.split('@'))[1].split(".")[0]
		$Alias = ($SMTPAddress.Replace("@", "_")).split(".")[0]
		
		# If no Department is set, create user in $RootOU
		If (!$User.Department) {
			$name = $User.DisplayName
			$now = Get-Date -format G
			Write-Output "$now - WARNING - Department not set for $name.  User placed in $RootOU" | Out-File $LogFile -Append
			$OU = $RootOU
		}
		Else {
			$OU = 'OU=' + $User.Department + ',' + $RootOU
		}
		
		# Test if OU is present in ABL AD, create new mailcontact
		$TestforOU = Get-ADOrganizationalUnit -SearchBase $OU -Filter *
		If ($TestforOU) {
			Try {
				$NewContact = New-ABLMailContact -Name $User.DisplayName -Alias $Alias -FirstName $User.FirstName -LastName $User.LastName -DisplayName $User.DisplayName -ExternalEmailAddress $SMTPAddress -OrganizationalUnit $OU -PrimarySmtpAddress $SMTPAddress -ErrorAction Stop
			}
			Catch {
				$name = $User.DisplayName
				$now = Get-Date -format G
				Write-Output "$now - $name could not be created --> $_" #| Out-File $LogFile -Append
			}
			If ($NewContact) {
				$name = $User.DisplayName
				$now = Get-Date -format G
				Write-Output "$now - $name has been created in $OU." | Out-File $LogFile -Append
			}
		}
		Else {
			$name = $User.DisplayName
			$now = Get-Date -format G
			Write-Output "$now - WARNING - OU $OU does not exist.  $name not created." | Out-File $LogFile -Append
		}
	}
	
}
If ($CreateThreshold -eq 0) {
	$size = $CreateArray.Count
	$now = Get-Date -format G
	Write-Output "$now - $size users need to be created, however the CreateThreshold is set to 0" | Out-File $LogFile -Append
}
# End CreateArray loop

# Check if $UpdateArray is populated
If ($UpdateArray.length -eq 0) {
	$now = Get-Date -format G
	Write-Output "$now - No users to update" | Out-File $LogFile -Append
}
# Update users based on $UpdateThreshold setting
Else {
	If ($UpdateAll) {
		$UpdateThreshold = $UpdateArray.Count
		$size = $UpdateArray.Count
		$now = Get-Date -format G
		Write-Output "$now -  $UpdateThreshold of $size users will be created.  Beginning mail contact creation" | Out-File $LogFile -Append
	}
	Else {
		If ($UpdateThreshold -gt 0) {
			$size = $UpdateArray.length
			$now = Get-Date -format G
			Write-Output "$now -  $UpdateThreshold of $size users will be updated.  Beginning mail contact updating" | Out-File $LogFile -Append
		}
	}
	$UpdateUsers = $UpdateArray | Select -First $UpdateThreshold
	Foreach ($User in $UpdateUsers) {
		If ($User -Is [Microsoft.Online.Administration.User]) {
			$SMTPAddress = $User.UserPrincipalName
		}
		If ($User -Is [Microsoft.Online.Administration.Contact]) {
			$SMTPAddress = $User.EmailAddress
		}
		
		# Create alias in the following format:  firstname.lastname_company
		$SMTPDomain = ($SMTPAddress.split('@'))[1].split(".")[0]
		$Alias = ($User.FirstName -replace "\s|\.") + '.' + ($User.LastName -replace " ") + '_' + $SMTPDomain
		
		# Run set command fo both MailContact and ADObject.  For simplicity, we pass all parameters, even if they haven't changed.  They will not be updated.
		$now = Get-Date -format G
		Write-Output "$now - Setting MailContact properties for $Alias" | Out-File $LogFile -Append
		Get-ABLMailContact -Identity $User.ABLGuid | Select Alias, DisplayName, ExternalEmailAddress | Out-File $LogFile -Append
		Try {
			Set-ABLMailContact -Identity $User.ABLGuid -Alias $Alias -DisplayName $User.DisplayName -ExternalEmailAddress $SMTPAddress -ForceUpgrade -erroraction stop
		}
		Catch {
			$now = Get-Date -format G
			"$now - Mail contact for $Alias could not be updated" + $_ | Out-File $LogFile -Append
		}
		Get-ABLMailContact -Identity $User.ABLGuid | Select Alias, DisplayName, ExternalEmailAddress | Out-File $LogFile -Append
		
		$now = Get-Date -format G
		Write-Output "$now - Setting ADObject properties for $Alias" | Out-File $LogFile -Append
		Get-ADObject -Identity $User.ABLGuid -Properties givenname, sn | Select Name, givenname, sn | Out-File $LogFile -Append
		Try {
			Set-ADObject -Identity $User.ABLGuid -Replace @{ givenname = $User.Firstname; sn = $User.Lastname } -ErrorAction Stop
		}
		Catch {
			$now = Get-Date -format G
			"$now - AD Object for $Alias could not  be updated" + $_ | Out-File $LogFile -Append
		}
		Get-ADObject -Identity $User.ABLGuid -Properties givenname, sn | Select Name, givenname, sn | Out-File $LogFile -Append
	}
}
If ($UpdateThreshold -eq 0) {
	$size = $UpdateArray.Count
	$now = Get-Date -format G
	Write-Output "$now - $size users need to be updated, however the UpdateThreshold is set to 0" | Out-File $LogFile -Append
}
# end UpdateArray loop

Write-Output "*****End Create\Update Users Section*****" | Out-File $LogFile -Append

# This is the end of the Create\Update section
######################################################################
# This is the beginning of the Delete Array section

Write-Output "*****Begin to Populate Delete Array Section*****" | Out-File $LogFile -Append

$zAgentContactList = Get-ABLMailContact -OrganizationalUnit $RootOU -ResultSize 10000 | where { ($_.OrganizationalUnit).Split("/")[2] -notlike "AEL*" } | Where-Object { $_.PrimarySmtpAddress -notlike "all*" -and $_.PrimarySmtpAddress -notlike "GALSync@Acuitybrandsmail.net" }
$IncludeOUArray += $O365Depts

# Populate the $DeleteArray
Foreach ($zAgentContact in $zAgentContactList) {
	$SMTPAddress = ($zAgentContact.PrimarySmtpAddress).ToString()
	$dn = $zAgentContact.DistinguishedName
	$DeptOU = ($dn.Split(",") | Select-String -Pattern '(AGT)\d{3}').ToString()
	$mcOU = $DeptOU.Split("=")[1]
	
	# Check if OU (Department) exists in O365, to make sure we only delete users from OU's that are hosted in O365
	If ($IncludeOUArray -contains $mcOU) {
		$User = Get-MsolUser -UserPrincipalName $SMTPAddress -ErrorAction SilentlyContinue
		If (!$User) {
			$User = Get-MsolContact -SearchString "$SMTPAddress"
		}
		# If no O365 user, push use to the DeleteArray
		if (!$User) {
			$DeleteArray += $zAgentContact
			$name = $zAgentContact.DisplayName
			$now = Get-Date -format G
			Write-Output "$now - $name Not found in Office365.  Need to delete user in ABLAD (zAgents)." | Out-File $LogFile -Append
		}
	}
}

Write-Output "*****End Delete Array Section*****" | Out-File $LogFile -Append
######################################################################
Write-Output "*****Begin Delete User Section*****" | Out-File $LogFile -Append

If ($DeleteArray.length -eq 0) {
	$now = Get-Date -format G
	Write-Output "$now - No users to delete" | Out-File $LogFile -Append
}
# Delete users based on $DeleteThreshold setting
Else {
	If ($DeleteAll) {
		$DeleteThreshold = $DeleteArray.Count
		$size = $DeleteArray.Count
		$now = Get-Date -format G
		Write-Output "$now -  $DeleteThreshold of $size users will be deleted.  Beginning mail contact deletion" | Out-File $LogFile -Append
	}
	Else {
		If ($DeleteThreshold -gt 0) {
			$size = $DeleteArray.length
			$now = Get-Date -format G
			Write-Output "$now - $size users need to be deleted.  Beginning mail contact deletion" | Out-File $LogFile -Append
		}
	}
	
	$DeleteUsers = $DeleteArray | Select -First $DeleteThreshold
	Foreach ($User in $DeleteUsers) {
		$now = Get-Date -format G
		$name = $User.Name
		$Guid = ($User.Guid).Guid
		Write-Output "$now - Deleting contact for $name" | Out-File $LogFile -Append
		Get-ABLMailContact -Identity $Guid | Out-File $LogFile -Append
		Try {
			Remove-ABLMailContact -Identity $Guid -Confirm:$false -erroraction stop
		}
		Catch {
			$now = Get-Date -format G
			Write-Output "$now - $name could not be deleted" + $_ | Out-File $LogFile -Append
		}
	}
}

If ($DeleteThreshold -eq 0) {
	$size = $DeleteArray.Count
	$now = Get-Date -format G
	Write-Output "$now - $size users need to be deleted, however the DeleteThreshold is set to 0" | Out-File $LogFile -Append
}
# end DeleteArray loop

Write-Output "*****End Delete Section*****" | Out-File $LogFile -Append

# End of Delete Section
######################################################################
####  HTML Code  ####

$style = Get-Content "$HomeDir\CSS.txt"

# Set HTML file path
# $htmlpath = Join-Path -Path $FilePath -ChildPath "GALException.HTML" #Change by ccd01

# HTML OU Section
If (!$HTMLOUObj) {
	$SendHTMLExceptReportByEmail = $false
}
	$params = @{
		'As' = 'Table';
		'PreContent' = '<h2>&diams; Exception in Department Format -- Format should match AGT000 standard (no spaces or extra characters) </h2>';
		'MakeTableDynamic' = $true;
		'TableCssClass' = 'grid'
	}
	
	$html_ou = $HTMLOUObj | ConvertTo-EnhancedHTMLFragment @params
	
	# End HTML OU Section
	
	# HTML Create Section
	# End HTML Create Section
	
	# HTML Delete Section
	# End HTML Create Section
	
	# HTML Assemble Section
	$params = @{
		'CssStyleSheet' = $style;
		'Title' = "GAL Sync Exception Report";
		'PreContent' = "<h1>GAL Sync Exception Report</h1>";
		'HTMLFragments' = @($html_ou)
	}
	
	ConvertTo-EnhancedHTML @params | Out-File -FilePath "$FilePath\GALException.HTML"
	$htmlpath = "$FilePath\GALException.HTML" #Change by ccd01

# End HTML Assemble Section
####  End HTML Code  ####
######################################################################
# SMTP settings.  Used to email the log and exception file

$mailparamslogs = @{
	SMTPServer = "smtpmail.acuitybrands.com"
	From = "O365Sync@acuitybrands.com"
	To = "AGTtoABLGalSync@AcuityBrands.com"
	Subject = "Agent to Acuity GALSync Log"
	Body = "GALSync results attached"
	Attachments = $Logfile
}
$mailparamshtml = @{
	SMTPServer = "smtpmail.acuitybrands.com"
	From = "O365Sync@acuitybrands.com"
	To = "AGTtoABLGalSync@AcuityBrands.com"
	Subject = "$date - GALSync Exception Report"
	Body = "$( Get-Content $htmlpath | out-string )" #Change by wwm01
	BodyAsHtml = $true
	Attachments = $Logfile
}
# Email the HTML Exception Report
If ($SendHTMLExceptReportByEmail) {
	#$HTMLBody = Get-Content $htmlpath
	Send-MailMessage @mailparamshtml
}

# Email the log file
If ($SendReportByEmail) {
	Send-MailMessage @mailparamslogs
}
