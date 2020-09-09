<#
  .VERSION AND AUTHOR
    Script version: v-2020.05.17
    Author: Stefano Pescosolido, https://www.linkedin.Qcom/in/stefanopescosolido/
    Script published in OneDrive: https://lnkd.in/epZhdAn
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
  This script removes the distribution groups or mails enabled security groups having a specific prefix in their names
  
  .PREREQUISITES
    PowerShell module Azure AD (or AzureADPreview)
    To install the PowerShell module, open PowerShell as administrator and type:
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
        Install-Module -Name AzureAD
            or
        Install-Module -Name AzureADPreview

    Details here: https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0

  .VARIABLES TO BE SET
  Read below.
#>

#########################################################################################################################
# VARIABLES TO BE SET:
#########################################################################################################################

$outCsvDirPath = "C:\Temp" #This is just an example. Set as required
$adminName = "nomeutente@nomescuola.edu.it" #This is just an example. Set as required
$GroupEmailPrexif = "studenti-" #Prefix to identify the groups to be removed
$GroupDisplayNamePrefix = "Studenti " #Prefix to identify the groups to be removed
$testOnly = $false #Set as required

#########################################################################################################################
cls
$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"

$outLogFilePath = "$outCsvDirPath\Remove-OldGroups_Results_$LogStartTime.log"
If (Test-Path $outLogFilePath){
	Remove-Item $outLogFilePath
}

Connect-AzureAD -Account $adminName 

Get-AzureADGroup -All $true | ForEach-Object{
    $g = $_
        
    #Check if it is a Team or a different type of group (security, distribution, office 365)
    $isTeam = $true
    try{
        Get-TeamChannel -GroupId $g.ObjectId | Out-Null
    }
    catch{
        $isTeam = $false
    }
    
    if ( (-not($isTeam)) -and ($g.MailNickName.StartsWith($GroupEmailPrexif)) -and ($g.DisplayName.StartsWith($GroupDisplayNamePrefix)) ){
        if(-not($testOnly)){
            try{
                Remove-AzureADGroup -ObjectId $g.ObjectId
                $line = "Successfully deleted group " + $g.MailNickName; Write-Host $line -ForegroundColor Green ; $line | Out-File $outLogFilePath -Append
            }
            catch{
                $line = "ERROR while deleting group '" + $g.MailNickName + "': " + $_.Exception.Message; Write-Host $line -ForegroundColor Red ; $line | Out-File $outLogFilePath -Append
            }
        }
        else{
            $line = "SIMULATED deletion of group " + $g.MailNickName; Write-Host $line -ForegroundColor Green ; $line | Out-File $outLogFilePath -Append
        }
    }
    else{
        if(-not($isTeam)){
            $line = "Skipped group " + $g.MailNickName; Write-Host $line -ForegroundColor Gray ; $line | Out-File $outLogFilePath -Append
        }
    }
    

}

Write-Host "Execution complete. Log file: " $outLogFilePath