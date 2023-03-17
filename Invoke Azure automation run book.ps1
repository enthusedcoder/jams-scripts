<#
	Purpose: Triggers processPipeline_TOM runbook in Azure Automation production
	Azure Subscription: IST-Prod-New
	Azure Subscription ID: 190e03c6-8316-4639-b667-940ea6228310
	Automation Account: ABL-Automation-BI-Prod
	Runbook Name: processPipeline_TOM
	Webhook URL: https://s1events.azure-automation.net/webhooks?token=vAOnPHiG0LJM0ys1gMvKxcDJbou%2bEwJJvXZgVrOnZgU%3d
	Description: Uses webhook to trigger runbook. Runbook processes the ADF pipeline
		from illumine to radiant and then processes the tabular object model.
	Created: 22-Jan-2018 Mark Fennell
	Updated: 20-Mar-2018 Mark Fennell
		Changed: added params for model processing and environment
	Updated:
		Changed:
	Updated:
		Changed:
#>

#PROD
$webhookurl = 'https://s5events.azure-automation.net/webhooks?token=aF0PktBl7QJuIuUmWGAwBf%2fNbrLuOujFjEm0ErLC6Us%3d'
#$webhookurl = 'https://s1events.azure-automation.net/webhooks?token=NHrldEu71qAXXY%2fFMWsZCf9kya3l16N%2fgISkevDIBDs%3d'
#dev-test
#$webhookurl = 'https://s1events.azure-automation.net/webhooks?token=TfYh98Zwvj3dHrGBOUb0kU%2ff0Jee%2byUrpODqPG3ehUI%3d'

$startDate = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")
$startDate = Get-Date $startDate -Format o
$endDate = (Get-Date).ToString("yyyy-MM-dd")
$endDate = Get-Date $endDate -Format o

# $envmt is used for PROD or DEV
$envmt = "PROD"

# shall we process the pipeline? YES or NO is good.
$procADF = "YES"

# shall we process the model? YES or NO will suffice.
$procMdl = "YES"

# for regularly scheduled processing, the $proc should be set to INC
# if $proc -eq FULL, the model will clear values and perform a full process
# ignored if procMdl = "NO"
$proc = "INC"

$body = @{"startDate" = "$startDate"; "endDate" = "$endDate"; "proc" = "$proc"; "envmt" = "$envmt"; "procMdl" = "$procMdl"; "procADF" = "$procADF";}

$params = @{
    ContentType = 'application/json'
    Headers = @{'from' = 'JAMS-PowerShell'; 'Date' = "$(Get-Date)"}
    Body = ($body | convertto-json)
    Method = 'Post'
    URI = $webhookurl
}

Invoke-RestMethod @params -Verbose

