function Get-SecretServerCred {
	<#
		.SYNOPSIS
			Gets a PowerShell credential from the SecretServer.

		.DESCRIPTION
			Gets a PowerShell credential from the SecretServer. Pass the function a SecretId for a secret on the Thycotic SecretServer and it will create a credential object.

		.PARAMETER  SecretId
			The SecretId for the secret you want.
			
		.PARAMETER SSUri
			The uri for the SecretServer windows authenticated web service

		.EXAMPLE
			PS C:\> $pscred = Get-SecretServerCred 32
			
		.EXAMPLE
			PS C:\> $pscred = Get-SecretServerCred 32 "http://secretserver/winauthwebservices/sswinauthwebservice.asmx"

		.INPUTS
			System.Int32

		.OUTPUTS
			System.Management.Automation.PSCredential

		.NOTES
			Currently this function uses the the Thycotic SecretServer windows authenticated service. It only works with creds with in a Secret with a "Username" and "Password" fields, and Domain accounts with "Domain" field.

		.LINK
			http://ist/kb/posh/abautomation
	#>
	[CmdletBinding()]
	[OutputType([System.Management.Automation.PSCredential])]
	param(
		[Parameter(Position=0, Mandatory=$true,ValueFromPipeline=$True)]
		[ValidateNotNull()]
		[System.Int32]
		$SecretId,
		[Parameter(Position=1)]
		[alias('Url')]
		[System.String]
		$SSUri = "http://secretserver/winauthwebservices/sswinauthwebservice.asmx"
	)
	
	begin {
		try {
			$ss = New-WebServiceProxy -UseDefaultCredential -Uri $SSUri -ErrorAction Stop -ErrorVariable WebservError
		} catch {
			throw
		}
	}
	
	process {
		$secretdef = $ss.GetSecret($SecretId,$true,$coderesponse).Secret.Items
		$domacct = $false
		$secretdef | foreach { if ($_.FieldName -eq "Domain") {$domacct = $true } }
		
		#region Parse Secret
		$secpasswd = convertto-securestring $($secretdef|where-object {$_.FieldName -eq "Password" }).Value -asPlainText -Force
		$user = $($secretdef|where-object {$_.FieldName -eq "Username" }).Value
		$domain = $($secretdef|where-object {$_.FieldName -eq "Domain" }).Value
		#endregion	
		
        #check to see if the account is a domain account
		if ( $domacct ) { 
			return New-Object System.Management.Automation.PSCredential("$user@$domain",$secpasswd)
			} else {
			return New-Object System.Management.Automation.PSCredential($user,$secpasswd) 
			}
	}
}
Get-PackageProvider -Name NuGet -ForceBootstrap
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
Get-PackageProvider -Name Powershellget -ForceBootstrap
<#
If (!(Test-Path "$env:ProgramFiles\windowspowershell\modules\AzureRM.KeyVault"))
{
	Install-Module AzureRm -AllowClobber -Force -Repository "PSGallery"
}
else
{
	Update-Module Azurerm -Force
}
If (!(Test-Path "$env:ProgramFiles\windowspowershell\modules\PowerShellGet"))
{
	Install-Module PowerShellGet -AllowClobber -Force -Repository "PSGallery"
}
else
{
	Update-Module PowerShellGet -Force
}
If (!(Test-path "F:\JAMSJobResources\box"))
{
	New-Item -Name box -Path "F:\JAMSJobResources" -ItemType directory -Force
}
If (!(Test-Path "F:\JAMSJobResources\box\Box.Dependencies\1.7.0\lib\net45\Box.V2.3.6.0\Box.V2.dll"))
{
	Register-PackageSource -Name "Box.Dependencies" -Location "https://www.myget.org/F/boxdependency/api/v2" -ProviderName "PowershellGet" -Trusted
	save-Package -Name Box.dependencies -path "F:\JAMSJobResources\box" -ProviderName "PowershellGet" -force
}
#>
Get-childitem "F:\JAMSJobResources\box\*.dll" -Recurse | % {[reflection.assembly]::LoadFrom($_.FullName)}
$creds = Get-SecretServerCred 4693
Login-AzureRmAccount -Credential $creds
$clientid = $null
		$clientsecret = $null
		$passphrase = $null
		$privatekey = $null
		$publickeyid = $null
		$enterpriseid = $null
		$clientid2 = $null
		$clientsecret2 = $null
		$passphrase2 = $null
		$privatekey2 = $null
		$publickeyid2 = $null
		$enterpriseid2 = $null
		$secrets = Get-AzureKeyVaultSecret -VaultName boxauth
		Foreach ($item in $secrets)
		{
			$azkey = Get-AzureKeyVaultSecret -name $item.Name -VaultName $item.VaultName
			If ($azkey.Name -like "clientID")
			{
				$clientid = $azkey.SecretValueText
			}
			ElseIf ($azkey.Name -like "clientSecret")
			{
				$clientsecret = $azkey.SecretValueText
			}
			Elseif ($azkey.Name -like "passphrase")
			{
				$passphrase = $azkey.SecretValueText
			}
			ElseIf ($azkey.Name -like "privateKey")
			{
				$privatekey = $azkey.SecretValueText #.replace("\n","`n").ToString()
			}
			Elseif ($azkey.Name -like "publicKeyID")
			{
				$publickeyid = $azkey.SecretValueText
			}
			Elseif ($azkey.Name -like "enterpriseID")
			{
				$enterpriseid = $azkey.SecretValueText
			}
			If ($azkey.Name -like "clientID2")
			{
				$clientid2 = $azkey.SecretValueText
			}
			ElseIf ($azkey.Name -like "clientSecret2")
			{
				$clientsecret2 = $azkey.SecretValueText
			}
			Elseif ($azkey.Name -like "passphrase2")
			{
				$passphrase2 = $azkey.SecretValueText
			}
			ElseIf ($azkey.Name -like "privateKey2")
			{
				$privatekey2 = $azkey.SecretValueText #.replace("\n","`n").ToString()
			}
			Elseif ($azkey.Name -like "publicKeyID2")
			{
				$publickeyid2 = $azkey.SecretValueText
			}
			Elseif ($azkey.Name -like "enterpriseID2")
			{
				$enterpriseid2 = $azkey.SecretValueText
			}
			Else
			{
			}
		}
		
		$boxconfig = New-Object -TypeName Box.V2.Config.Boxconfig($clientid, $clientSecret, $enterpriseID, $privateKey, $passphrase, $publicKeyID)
		$boxJWT = New-Object -TypeName Box.V2.JWTAuth.BoxJWTAuth($boxconfig)
		$boxjwt
		$tokenreal = $boxJWT.AdminToken()
		$adminclient = $boxjwt.AdminClient($tokenreal, "401268528")
		$adminclient
		
		$boxconfig2 = New-Object -TypeName Box.V2.Config.Boxconfig($clientid2, $clientSecret2, $enterpriseID2, $privateKey2, $passphrase2, $publicKeyID2)
		$boxJWT2 = New-Object -TypeName Box.V2.JWTAuth.BoxJWTAuth($boxconfig2)
		$boxjwt2
		$tokenreal2 = $boxJWT2.AdminToken()
		$adminclient2 = $boxjwt2.AdminClient($tokenreal2, "401268528")
		$adminclient2
		$exist = $true
		$curfiles = $null
		If (!(Test-Path "F:\JAMSJobResources\box\currentfiles.txt"))
		{
			$exist = $false
		}
		Else
		{
			$exist = $true
			$curfiles = Get-content "F:\JAMSJobResources\box\currentfiles.txt"
		}

		$folder = $adminclient.FoldersManager.GetItemsAsync("48813635461", 1000, 0)
		$folder.Wait()
		$array = New-object System.Collections.arraylist
		Foreach ($item in $folder.Result.ItemCollection.Entries)
		{
			If ($exist = $true)
			{
				If ($curfiles -notcontains $item.Name)
				{
					$psob = New-object -TypeName PSObject
					$psob | Add-Member -MemberType NoteProperty -Name Name -Value $item.Name
					$psob | Add-Member -MemberType NoteProperty -Name Id -Value $item.Id
					$array.Add($psob)
				}
			}
			Else
			{
				$psob = New-object -TypeName PSObject
					$psob | Add-Member -MemberType NoteProperty -Name Name -Value $item.Name
					$psob | Add-Member -MemberType NoteProperty -Name Id -Value $item.Id
					$array.Add($psob)
			}
		}
If ($array.Count -gt 0)
{
$smtpServer = "smtpmail.acuitybrands.com"
$smtp = New-Object System.Net.Mail.SmtpClient($smtpServer)
$smtp.Port = 25
$Email = new-object System.Net.Mail.MailMessage
$Email.From = "`"Acuity ADP File`" <adp@acuitybrands.com>"
$Email.Subject = "New adp file(s) have been uploaded to box"
$Email.IsBodyHtml = $true
$holstring = $null
$sharejson = [Box.V2.Models.BoxSharedLinkRequest]::new()
$sharejson.Access = "open"
$body = @"
<html>
<body lang=EN-US link=blue vlink=purple style='tab-interval:.5in'>

<div class=WordSection1>

<p class=MsoNormal><o:p>A new adp file has been uploaded to the box adp folder.  See attachment(s) for files.</o:p></p>

</div>

</body>

</html>
"@
$Email.Body = $body
$coll = New-object System.Net.Mail.MailAddressCollection
$collaber = $adminclient.FoldersManager.GetCollaborationsAsync("48813635461")
$collaber.Wait()
foreach ($item in $collaber.Result.Entries)
{
    $mailadd = New-object System.Net.Mail.MailAddress -ArgumentList $item.AccessibleBy.Login, $item.AccessibleBy.Name
    $coll.Add($mailadd)
}
$email.To.Add($coll.ToString())
Foreach ($ob in $array)
{
	$stream = $adminclient.FilesManager.CreateSharedLinkAsync($ob.Id, $sharejson)
	$stream.Wait()
	Invoke-Webrequest $stream.Result.SharedLink.DownloadUrl -OutFile "$env:APPDATA\$($ob.Name)"
	$attach = New-Object System.Net.Mail.Attachment -ArgumentList "$env:APPDATA\$($ob.Name)"
	$Email.Attachments.Add($attach)
	echo $ob.Name | out-file "F:\JAMSJobResources\box\currentfiles.txt" -Append
}
$smtp.Send($Email)
}
