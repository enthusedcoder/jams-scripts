$Path = "\\cdcsrvr2\archive\MEXEINVOICE"

$Files = Get-ChildItem -File $Path

Foreach ($File in $Files) {
	$date = $File.CreationTime
	If ($date.Month -le ((Get-Date).AddMonths(-1)).Month) {
		$ChkPath = "$Path\$($date.ToString("yyyy"))\$($date.ToString("MMMM"))"
		If (!(Test-Path $ChkPath)) {
			New-Item -Path $ChkPath -ItemType Directory
		}
		Move-Item $($File.FullName) -Destination $ChkPath
		Write-Output "Moving $File"
	}
}
