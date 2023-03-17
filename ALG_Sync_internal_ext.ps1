$env:ADPS_LoadDefaultDrive = 0
Import-Module ActiveDirectory

$ipaddress = "cdc-uc01"
$today = (Get-Date).ToString('yyyyMMdd')
$fullpath = "\\cdcsrvr1\Depts\IST\Enterprise Engineering\ADPhoneSync\COBRAS_Backup_$($today)_*$($ipaddress)\unitydbdata_backup_$($ipaddress)_*$($today)_*.mdb"
$path = (Get-ChildItem $fullpath).FullName

$ipphone = "ipPhone"
$attr12 = "extensionAttribute12"

$query = "SELECT Subscribers.[Alias], Subscribers.[DTMFAccessID], Subscribers.[DisplayName], Subscribers.[LDAPCCMUserID] FROM Subscribers"

$adOpenStatic = 3
$adLockOptimistic = 3

$cn = new-object -comobject ADODB.Connection
$rs = new-object -comobject ADODB.Recordset
$connectionString = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source = $path"
$cn.Open($connectionString)

$rs.Open($query, $cn, $adOpenStatic, $adLockOptimistic)

$rs.MoveFirst()
$vmlist = @()

do {
	$vmbox = New-Object psobject
	$vmbox | Add-Member -NotePropertyName "AccountName" -NotePropertyValue $rs.Fields.Item("Alias").Value
	$vmbox | Add-Member -NotePropertyName "DTMF" -NotePropertyValue $rs.Fields.Item("DTMFAccessID").Value
	$vmbox | Add-Member -NotePropertyName "LDAPUserID" -NotePropertyValue $rs.Fields.Item("LDAPCCMUserID").Value
	if ($vmbox.DTMF -notlike "1010*" -and $vmbox.DTMF -notlike "9999*" -and $vmbox.LDAPUserID -ne [DBNull]::Value) {
		$vmlist += $vmbox
	}
	$rs.MoveNext()
}
until ($rs.EOF -eq $True)

$rs.Close()
$cn.Close()

#$vmlist | export-csv -NoTypeInformation -Path \\winmgmt\c$\TIDALJOBS\AD\PhoneSync\vmexport_test.csv

$mismatchlist = @()

foreach ($vmbox in $vmlist) {
	#Write-Output "Checking $($vmbox.LDAPUserID)"
	Try {
		$aduser = Get-ADUser -Identity $vmbox.LDAPUserID -Properties ipphone -ErrorAction SilentlyContinue
	}
	Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
		Write-Output "$($vmbox.LDAPUserID) not found in AD"
	}
	If ($aduser) {
		if ($vmbox.DTMF -ne $aduser.ipphone) {
			$mismatch = New-Object psobject
			$mismatch | Add-Member -NotePropertyName "SamAccountName" -NotePropertyValue $($aduser.SamAccountName)
			$mismatch | Add-Member -NotePropertyName "ipphone" -NotePropertyValue $($aduser.ipphone)
			$mismatch | Add-Member -NotePropertyName "DTMF" -NotePropertyValue $($vmbox.DTMF)
			$mismatch | Add-Member -NotePropertyName "Name" -NotePropertyValue $($aduser.Name)
			$mismatchlist += $mismatch
			Set-ADUser -Identity $aduser.SamAccountName -replace @{ $ipphone = $vmbox.DTMF }
		}
	}
}

if ($mismatchlist) {
	Export-Csv -InputObject $mismatchlist -NoTypeInformation -Path "\\cdcsrvr1\Depts\IST\Enterprise Engineering\ADPhoneSync\Log\VM_System_$(get-date -Format 'MMddyyyyhhmm').csv"
}

$list = Get-ADUser -SearchBase "DC=acuitylightinggroup,DC=com" -SearchScope Subtree -Filter * -Properties ipphone, extensionAttribute12 | Where-Object { $_.extensionAttribute12 -ne $_.ipphone -and $_.ipphone -match "^[0-9]{4}$" }

if ($mismatchlist) {
	
	Export-Csv -InputObject $mismatchlist -NoTypeInformation -Path "\\cdcsrvr1\Depts\IST\Enterprise Engineering\ADPhoneSync\Log\Attr12_update_$(get-date -Format 'MMddyyyyhhmm').csv"
	
	foreach ($user in $list) {
		
		Set-ADUser -Identity $user.SamAccountName -replace @{ $attr12 = $user.ipphone }
		$aduser = get-aduser $user.SamAccountName -Properties ipphone, info
		$info = $aduser.info
		
		if ($info) {
			#Populated
			if ($info -match 'Ext: ([0-9]{4})') {
				#Ext:
				if ($info -notmatch "Ext: $($aduser.ipphone)") {
					$info = $info -replace 'Ext: ([0-9]{4})', "Ext: $($aduser.ipphone)"
					Set-ADUser -Identity $aduser.SamAccountName -Replace @{ "info" = $info }
				}
			}
			else {
				#no Ext:
				$info = $info.Insert(0, "Ext: $($aduser.ipphone)`r`n")
				Set-ADUser -Identity $aduser.SamAccountName -Replace @{ "info" = $info }
			}
		}
		else {
			#Not Populated
			Set-ADUser -Identity $aduser.SamAccountName -add @{ "info" = "Ext: $($aduser.ipphone)" }
		}
	}
}
