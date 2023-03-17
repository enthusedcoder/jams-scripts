#
#  Make sure that the JAMS Snapin is snapped in
#
Add-PSSnapin MVPSI.JAMS -errorAction SilentlyContinue
#
#  Our JAMS Entry number is passed to us in the $Host.PrivateData object
#  We will use that to retrieve our JAMSEntry object
#
$jamsEntry = Get-JAMSEntry -Entry $Host.PrivateData.JAMSEntry
#
#  Now that we have the JAMSEntry, let's use it!
#
Write-Host "JAMS Entry number is " $jamsEntry.JAMSEntry
Write-Host "Job Name is " $jamsEntry.JobName
Write-Host "Submitted by " $jamsEntry.SubmittedBy
Write-Host "RON is " $jamsEntry.RON



