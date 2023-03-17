Function Get-StringHash([String]$String, $HashName = "MD5")
{
	$seedString = "2uWxs1orY5MfuLcz0YL719V1JP912WCHWO8eQ3do5gPTCaL5HU" + $String
	$StringBuilder = New-Object System.Text.StringBuilder
	[System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash([System.Text.Encoding]::UTF8.GetBytes($seedString)) | ForEach-Object{
		[Void]$StringBuilder.Append($_.ToString("x2"))
	}
	$StringBuilder.ToString()
}

$acuityHashLocations = Get-ChildItem -Path \\cdcsrvr2\images\*.AcuityHash -Recurse
[string[]]$acuityHashes = New-Object string[] $acuityHashLocations.Length
for ($j = 0; $j -lt $acuityHashLocations.Length; $j++)
{
	$acuityHashes[$j] = Get-Content $acuityHashLocations[$j].FullName
}

$imageLocations = Get-ChildItem -Path \\cdcsrvr2\images\*.wim -Recurse
$mailBody = ""
$mailBodyP = ""
$mailBodyW = ""
$numOfUnknownMD5 = 0
$emailSubject = "Windows Image Validation Audit: All Clear"
foreach ($imageLCT in $imageLocations)
{
	$hash = Get-FileHash $imageLCT.PSPath -Algorithm MD5
	$outHash = Get-StringHash $hash.hash "SHA512"
	$foundMD5 = $false
	for ($i = 0; $i -lt $acuityHashes.Length; $i++)
	{
		if ($acuityHashes[$i] -contains $outHash)
		{
			$mailBodyP = $mailBodyP + "`nPass: " + $imageLCT.FullName + " = " + $acuityHashes[$i]
			$i = $acuityHashes.Length
			$foundMD5 = $true
		}
	}
	if ($foundMD5 -eq $false)
	{
		$mailBodyW = $mailBodyW + "`nWARNING! The image found at `"" + $imageLCT.FullName + "`" DOES NOT match any of the known Acuity Hashes, verify this image.`n"
		$numOfUnknownMD5++
	}
}
if ($numOfUnknownMD5 -eq 0)
{
	$mailBody = "All OS images have been validated.`n" + $mailBodyP
}
elseif ($numOfUnknownMD5 -eq 1)
{
	$emailSubject = "WARNING! Windows Image Validation Audit: Unknown Image Found"
	$mailBody = "There is 1 unknown image, verify it below:`n" + $mailBodyW + $mailBodyP
}
else
{
	$emailSubject = "WARNING! Windows Image Validation Audit: Unknown Images Found"
	$mailBody = "There are " + $numOfUnknownMD5 + " unknown images, verrify these below:`n" + $mailBodyW + $mailBodyP
}

$mailparams = @{
	'SMTPServer' = "smtpmail.acuitybrands.com"
	'From'	   = "JAMS@acuitybrands.com"
	'To'		 = "informationsecurity@AcuityBrands.com", "Jeff.Stropoli@AcuityBrands.com"
	'Subject'    = $emailSubject
	'Body'	   = $mailBody
}

$mailBody
Send-MailMessage @mailparams
