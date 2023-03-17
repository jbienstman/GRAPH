#region - variable(s) #########################################################################################################################################
$csvSubFolderName = "getSharePointSiteUsageDetailHistory"
$daysHistory = 27 #SharePoint Site Usage Reports are only available for the past 28 days, including today, so that's 27 day(s) history maximum
#
$doNotReportWeekends = $true
$overwriteExistingReports = $false
$outputEnabled = $false
#
$graphApiRoot = "https://graph.microsoft.com/v1.0"
$graphApiReport = "/reports/"
$graphApiReportName = "getSharePointSiteUsageDetail"
#
#endregion - variable(s) #########################################################################################################################################
#region - include(s) #########################################################################################################################################
$ErrorActionPreference = 'Stop'
$scriptPath = $PSScriptRoot #get path of "this" script
. ($scriptPath.ToString() + "\_includes\Get-GraphApiHeaders.ps1") #returns graph api headers in variable $Headers
. ($scriptPath.ToString() + "\_includes\Get-GraphApiResultsWithAutoThrottle.ps1") #loads function only
. ($scriptPath.ToString() + "\_includes\Rename-ItemAppendCreationDate.ps1") #loads function only
#endregion - include(s) #########################################################################################################################################
#region - getSharePointSiteUsageDetail - History  #########################################################################################################################################
if ($outputEnabled) {Write-Host ("Retrieving ($daysHistory) Data Points from [$graphApiReportName] Report:") -ForegroundColor Cyan}
for ($i = 1; $i -le $daysHistory ; $i++)
    {
    $graphApiDateValue = Get-Date -Date (Get-Date).AddDays(-$i) -Format yyyy-MM-dd #YYYY-MM-DD. As this report is only available for the past 30 days, {date_value} should be a date from that range.
    $dayOfweek = (Get-Date $graphApiDateValue -Format dddd)
    $graphApiDateEndpoint = "(date=$graphApiDateValue)"
    if ($outputEnabled) {Write-Host ("[" + "{0:d2}" -f $i + "/" + $daysHistory + "] - ($graphApiDateValue)") -NoNewline}
    #{0:d3}' -f $_
    if ($doNotReportWeekends -eq $true -and ($dayOfweek -eq "Sunday" -or $dayOfweek -eq "Saturday"))
        {
        if ($outputEnabled) {Write-Host (" - Skipping weekends!") -ForegroundColor Yellow}
        continue
        }
    else
        {
        $reportFileRelativePath = "\$csvSubFolderName\" + $graphApiReportName + "_" + $graphApiDateValue + ".csv"
        $reportFilePath = ($PSScriptRoot.TrimEnd("\") + $reportFileRelativePath)
        if ((Test-Path -Path $reportFilePath) -and $overwriteExistingReports -eq $false)
            {
            if ($outputEnabled) { Write-Host (" - Already Exists!") -ForegroundColor Green}
            continue
            }
        $graphApiReportUri = ($graphApiRoot + $graphApiReport + $graphApiReportName + $graphApiDateEndpoint)
        $report = $null
        $report = Get-GraphApiResultsWithAutoThrottle -uri $graphApiReportUri -headers $Headers -method Get -contentType "application/json" -outputEnabled:$outputEnabled
        #Invoke-RestMethod -Uri $graphApiReportUri -Headers $Headers -Method get -ContentType "application/json"
        $report = $report -Replace "ï»¿", "" | ConvertFrom-Csv
        if ($report)
            {
                $report | Export-Csv -Path $reportFilePath -Encoding UTF8 -NoTypeInformation
                if ($outputEnabled) { Write-Host ("[...$reportFileRelativePath]") -ForegroundColor DarkGray}
            }
        else
            {
                if ($outputEnabled) { Write-Host (" - [No Data]") -ForegroundColor Magenta}
            }
        }
    }
#endregion - getSharePointSiteUsageDetail - History  #########################################################################################################################################
#region - getSharePointSiteUsageDetail - D30  #########################################################################################################################################
#region - variable(s)
$graphApiPeriodValue = "D30"
#endregion - variable(s)
$graphApiPeriodEndPoint = "(period='$graphApiPeriodValue')"
$graphApiReportUri = ($graphApiRoot + $graphApiReport + $graphApiReportName + $graphApiPeriodEndPoint)
$reportFileRelativePath = "\$csvSubFolderName\" + $graphApiReportName + "_" + $graphApiPeriodValue + ".csv"
$reportFilePath = ($PSScriptRoot.TrimEnd("\") + $reportFileRelativePath)
if ($outputEnabled) {Write-Host ("Retrieving [$graphApiReportName] report for ($graphApiPeriodValue) Period:") -ForegroundColor Cyan -NoNewline}
if ((Test-Path -Path $reportFilePath) -and $overwriteExistingReports -eq $false)
    {
    $renamedItemFullPath = Rename-ItemAppendDateProperty -itemFullPath $reportFilePath -propertyName 'Report Refresh Date' -outputEnabled:$outputEnabled
    }
$report = $null
$report = Get-GraphApiResultsWithAutoThrottle -uri $graphApiReportUri -headers $Headers -method Get -contentType "application/json" -outputEnabled:$outputEnabled
$report = $report -Replace "ï»¿", "" | ConvertFrom-Csv
$report | Export-Csv -Path $reportFilePath -Encoding UTF8 -NoTypeInformation
if ($outputEnabled) { Write-Host ("[...$reportFileRelativePath]") -ForegroundColor DarkGray}
#endregion - getSharePointSiteUsageDetail - D30  #########################################################################################################################################