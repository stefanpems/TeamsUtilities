<#
  .VERSION AND AUTHOR
    Script version: v-2020.04.08
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .ATTENTION (UPDATE 2020.10.25)
  The scripts in this workspace that use GraphAPI do not work with the latest version (Oct. 2020 - v. 3.26.2010.0) of the module 
  PnP.PowerShell / SharePointPnPPowerShellOnline.
  We are investigating. Please use any previous version of that module (do not install the latest version by using the command
  "Install-Module SharePointPnPPowerShellOnline" referenced below).
  More info on https://github.com/pnp/PnP-PowerShell/issues/2983

  .SYNOPSIS
  This script downloads the Teams Activity Report for a specific day

  .VARIABLES TO BE SET
    See below

  .PREREQUISITES
   * Use Windows 10 (For earlier versions of Windows, please refer to https://docs.microsoft.com/en-us/microsoft-365/enterprise/connect-to-all-microsoft-365-services-in-a-single-windows-powershell-window?view=o365-worldwide#before-you-begin)
   * If not alreayd done, install the PowerShell module SharePointPnPPowerShellOnline

    To install the PowerShell module, open PowerShell by using the option "Run as administrator" and type:
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
        Install-Module SharePointPnPPowerShellOnline
    
    Details here: https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-all-office-365-services-in-a-single-windows-powershell-window
#>

#########################################################################################################################
# VARIABLES TO BE SET:
#########################################################################################################################

$desiredDate = "2020-04-06" #YYYY-MM-DD. As this report is only available for the past 30 days
$outCsvDirPath = "C:\Temp" #Set a valid path

#########################################################################################################################

function QueryTeamsUserActivity{

    param(
         [string]
         $dateToReport, 
         [string]
         $accessToken
    )

    $returnCsv = $null;
    $GraphURL = "https://graph.microsoft.com/beta" 

    try
    {
        $getTeamsUserActivityFromGraphUrl = "$GraphURL/reports/getTeamsUserActivityUserDetail(date=$dateToReport)" 
        $TeamsUserActivityResponse = Invoke-RestMethod -Uri $getTeamsUserActivityFromGraphUrl -Headers @{Authorization = "Bearer $accesstoken"} -Method Get 
        if($TeamsUserActivityResponse){
            $returnCsv = $TeamsUserActivityResponse.Substring(3,$TeamsUserActivityResponse.Length-3) | ConvertFrom-Csv | sort 'Last Activity Date', 'User Principal Name' -Descending # Team Chat Message Count,Private Chat Message Count,Call Count,Meeting Count,Has Other Action,Report Period
        }        
            
    }
    catch
    {
        Write-Host "Error while executing query - " $_.Exception.Message -ForegroundColor Red
        throw $_
    }

    $returnCsv;

}


Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue 
$arrayOfScopes = @("Reports.Read.All") 
Connect-PnPOnline -Scopes $arrayOfScopes
$accessToken = Get-PnPAccessToken

Write-Host "Execution started - Teams activity report - Day: " $desiredDate

$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
$outCsvFilePath = "$outCsvDirPath\TeamsUsersActivity_Results_$LogStartTime.csv"
If (Test-Path $outCsvFilePath){
	Remove-Item $outCsvFilePath
}

$resultsCsv = QueryTeamsUserActivity -dateToReport $desiredDate -accessToken $accessToken

if($resultsCsv){
    $Count = 0;
    $resultsCsv | ForEach-Object{
        if($Count -eq 0){
            ($_ | ConvertTo-Csv)[1] | Out-File $outCsvFilePath 
        }
        $Count++;
        Write-Host "   (" $Count ")" $_.'User Principal Name' -ForegroundColor Green
        ($_ | ConvertTo-Csv)[2] | Out-File $outCsvFilePath -Append
    }
}
else{
    Write-Host "   No user found" -ForegroundColor Gray
}
Write-Host "Execution ended - Output file: '" $outCsvFilePath "'" -ForegroundColor Yellow

