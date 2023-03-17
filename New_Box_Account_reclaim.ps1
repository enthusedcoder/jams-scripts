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
<#
$files = Get-childitem "F:\JAMSJobResources\box\*.dll" -Recurse
Foreach ($file in $files)
{
	Try
	{
		[Reflection.Assembly]::LoadFrom("$($file.FullName)")
	}
	Catch
	{
		Continue
	}
}
#>
$date = Get-Date -Format "yyyy-MM-dd"
$date = ([DateTime]$date).AddMonths(-6)

<#
$boxconfig = New-Object -TypeName Box.V2.Config.Boxconfig($cli.boxAppSettings.clientID, $cli.boxAppSettings.clientSecret, $cli.enterpriseID, $cli.boxAppSettings.appAuth.privateKey, $cli.boxAppSettings.appAuth.passphrase, $cli.boxAppSettings.appAuth.publicKeyID)
$boxJWT = New-Object -TypeName Box.V2.JWTAuth.BoxJWTAuth($boxconfig)
$boxjwt
$tokenreal = $boxJWT.AdminToken()
$adminclient = $boxjwt.AdminClient($tokenreal)
$adminclient
#>
$allevents = New-Object System.Collections.Generic.List[object]
$enteventurl = "https://api.box.com/2.0/events?stream_type=admin_logs&event_type=Login&limit=500&created_after=$date"
$headers = @{ }
$headers.Add("Authorization", "Bearer $tokenreal")
$headers.Add("Accept-Encoding", "gzip, deflate")
$resultevents = Invoke-RestMethod -Uri $enteventurl -Method Get -Headers $headers
$allevents.Add($resultevents.entries)
Do
{
	$streampos = $resultevents.next_stream_position
	$enteventurl = "https://api.box.com/2.0/events?stream_type=admin_logs&event_type=Login&limit=500&created_after=$date&stream_position=$streampos"
	$resultevents = Invoke-RestMethod -Uri $enteventurl -Method Get -Headers $headers
	$allevents.Add($resultevents.entries) > $null
}
Until ($resultevents.chunk_size -lt 10)
$userarray = New-Object System.Collections.ArrayList
foreach ($item in $allevents)
{
	$tempvar = $item.source
	Foreach ($stick in $tempvar)
	{
		$userarray.Add($stick) > $null
	}
}


$hold = 0
$limit = 1000
$userhold = New-Object System.Collections.ArrayList
[string[]]$fields = "role","login","id","name","created_at"
[System.Collections.Generic.IEnumerable[String]]$conv = $fields
$result = $adminclient.UsersManager.GetEnterpriseUsersAsync($null, $hold, $limit, $conv)
$result.Wait()
$totalcount = $result.result.totalcount
$result.result.entries | % {$userhold.Add($_)} > $null
Do
{
    $hold = $hold + $limit
    $result = $adminclient.UsersManager.GetEnterpriseUsersAsync($null, $hold, $limit, $conv)
    $result.Wait()
    $result.result.entries | % {$userhold.Add($_)} > $null
}
While ($hold + $limit -lt $totalcount)
$thdate = Get-date
	$thdate = $thdate.AddMonths(-1)
#loops through all of the Acuity brands box users and will delete the user's box account if they are not in the array of users associated with any of the returned login events and are not admins.
Foreach ($user in $userhold)
{
	Write-Host "Analyzing user $($user.name)."
	$thtime = New-timespan -start $thdate -end $([Datetime]$user.CreatedAt)
    #this portion of if statement executes when the user being analyzed is not found in the array of the users who generated a login event for the past 6 months and if the user is not an admin
	#If (($userarray.id -notcontains $user.id) -and ($user.role -notlike "*admin*"))
If (($userarray.id -notcontains $user.id) -and ($thtime.Days -lt 0))
	{
		Write-Host "$($user.name) has not been active in over 6 months.  Attempting to delete box account."
		$mail = $user.login
        #uses the box sdk to attempt to delete the user's box account.
        $deletetask = $adminclient.UsersManager.DeleteEnterpriseUserAsync($user.id, $false, $false)
				Try
				{
					$deletetask.Wait()
				}
				Catch
				{
				}
        If ($deletetask.IsFaulted -eq $false)
        {
            Write-host "$($user.name)'s account was deleted.  Now removing from Federate-Box AD group."
            Try
			{
                #if the account for the user being analyzed was deleted, the script will then attempt to modify the ad account of that user and remove it from the "Federated - Box" Group
				$aduser = Get-ADUser -Filter "Mail -like `"$mail`"" -ErrorAction Stop
				Remove-ADGroupMember -Identity "Federate - Box" -Members $aduser -Confirm:$false -ErrorAction Stop
                Write-host "$($user.name)'s ad account has been removed from the Federate - box group"
			}
			Catch
			{
				Write-host "$($user.name)'s ad account is not present in the Box - Federate Group.  Moving on...."
                Continue
			}
        }
        Else
        {
            #if the user account could not be deleted using the box sdk, this information is written to the console and the next user in the array will start processing.
            Write-host "Failed to delete $($user.name).  The user probably has managed content in their box account."
            Continue
        }
	}
    #this code runs if the user being analyzed is in the array of users with login events or if the user is an admin.
	Else
	{
		Write-Host "$($user.name)'s account has shown activity in the past 6 months or was created within the last 2 weeks.  Skipping...."
        Continue
	}
}


