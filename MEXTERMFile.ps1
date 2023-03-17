$ErrorActionPreference = "Stop"

Import-module JAMS
$CachCredential = Get-JAMSCredential FTPHRMEX -Server jamssched01
$Yesterday = (Get-Date).AddDays(-1).ToString('ddMMyyyy')

$RCCach = Connect-JFTP ftp.acuitybrands.com -Credential $CachCredential

Try {
	$TermFile = Get-JFSItem -Path "termination_file_$Yesterday.csv" -FileServer $RCCach
}
Catch [System.Management.Automation.ItemNotFoundException] {
	Write-Output "Termination File Is Not Present"
}

If ($TermFile) {
	Receive-JFSItem "termination_file_$Yesterday.csv" \\cdcsrvr1\Depts\CDCHR\ADP_INTERFACE_FILES\ERP_DATA\mexico.csv -FileServer $RCCach -verbose
}

Disconnect-JFS -FileServer $RCCach
