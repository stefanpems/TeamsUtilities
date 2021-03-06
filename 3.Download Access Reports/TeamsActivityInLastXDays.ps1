﻿<#
  .VERSION AND AUTHOR
    Script version: v-2020.04.08
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .ATTENTION (UPDATE 2020.10.25)
  The script doesn't work with the latest version (Oct. 2020 - v. 3.26.2010.0) of the module 
  PnP.PowerShell / SharePointPnPPowerShellOnline.
  We are investigating. While we find a solution, please do not install the latest version or upgrade your current version, so
  DO NOT install it by simply launching the command referenced below: Install-Module SharePointPnPPowerShellOnline  
  Instead, check if you have the latest version already installed; if so, remove it; then install the previous version.
  => To check if you have the issue, use the command: Connect-PnPOnline -Scopes "Mail.Read","Files.Read","ActivityFeed.Read" 
     If you have the issue, you get an error AADSTS70011. If you don't have the issue, you get a login prompt
  => To check which version is installed use the command: Get-Module -Name sharepointpnppowershell* -ListAvailable
  => To remove any existing version use the command: Uninstall-Module SharePointPnPPowerShellOnline -AllVersions -Force"
  => To install the previous (last working) version, use the command: "Install-Module SharePointPnPPowerShellOnline -RequiredVersion 3.25.2009.1"
   
  More info on https://github.com/pnp/PnP-PowerShell/issues/2983

  .SYNOPSIS
  This script downloads the Teams Activity Report for the last 7 days (or 30 days, if available)

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

$daysBack = 30 #Possibili valori: 7,30,90,180
$outCsvDirPath = "C:\Temp" #Set a valid path

#########################################################################################################################

function QueryTeamsUserActivity{

    param(
         [int]
         $NumberOfDays, 
         [string]
         $accessToken
    )

    $returnCsv = $null;
    $GraphURL = "https://graph.microsoft.com/beta" 

    try
    {
        $getTeamsUserActivityFromGraphUrl = "$GraphURL/reports/getTeamsUserActivityUserDetail(period=%27D$NumberOfDays%27)" 
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

Write-Host "Execution started - Report for the last " $daysBack " days:" 

$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
$outCsvFilePath = "$outCsvDirPath\TeamsUsersActivity_Results_$LogStartTime.csv"
If (Test-Path $outCsvFilePath){
	Remove-Item $outCsvFilePath
}

$resultsCsv = QueryTeamsUserActivity -NumberOfDays $daysBack -accessToken $accessToken

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

