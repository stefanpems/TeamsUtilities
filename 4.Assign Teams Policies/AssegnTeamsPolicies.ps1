<#
  .VERSION AND AUTHOR
    Script version: v-2020.04.13
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
  This script applies Teams policies at a filtered set of Teams users.
  The filter is currently based on the value of the attribute JobTitle.

  .NOTES
  Depending on what you need to do, you may not need all these modules. 
  The recommendation is to install only the modules needed in your script.

  .MORE INFO
  https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-all-office-365-services-in-a-single-windows-powershell-window

  .PREREQUISITES
  * Use Windows 10 (For earlier versions of Windows, please refer to https://docs.microsoft.com/en-us/microsoft-365/enterprise/connect-to-all-microsoft-365-services-in-a-single-windows-powershell-window?view=o365-worldwide#before-you-begin)
  * If not alreayd done, install the PowerShell modules Azure AD (or AzureADPreview) and MicrosoftTeams

    To install the PowerShell modules, open PowerShell by using the option "Run as administrator" and type:
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
        Install-Module -Name AzureAD
            or
        Install-Module -Name AzureADPreview
            and
        Install-Module -Name MicrosoftTeams

    Details here: https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0
    and here: https://docs.microsoft.com/en-us/MicrosoftTeams/teams-powershell-install

#>

#########################################################################################################################
# VARIABLES TO BE SET:
#########################################################################################################################

$outLogDirPath = "C:\Temp" #Set the correct path
$adminName = "nomeadmin@nomescuola.edu.it" #Set the correct value
$filterJobTitle = "Studente" #Set the correct value for filtering your users based on the value of the attribute JobTitle
$teamsPolicyTypes = @(
    "TeamsMeetingPolicy",
    "TeamsMessagingPolicy",
    "TeamsAppSetupPolicy",
    "TeamsCallingPolicy"
    ) #Remove the policy types that you don't want to assign
$policyName = "Education_PrimaryStudent" #Change the policy name as needed

#########################################################################################################################
$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"

$logFile = "$outLogDirPath\AssignPolicies_$LogStartTime.log"
If (Test-Path $logFile){
	Remove-Item $logFile
}
"START: " + $LogStartTime | Out-File $logFile
Connect-AzureAD -AccountId $adminName
Connect-MicrosoftTeams -AccountId $adminName

$users = Get-AzureADUser -All $true -Filter "JobTitle eq '$filterJobTitle'" 

$teamsPolicyTypes| ForEach-Object{

    $time = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
    "START " + $policyType + " - Time: " + $time | Out-File $logFile -Append

    $policyType = $_
    $opName = "Assegnazione "+$policyType+" "+$policyName
    $pao = Get-CsBatchPolicyAssignmentOperation | where "OperationName" -eq $opName

    if($pao){
        $opid = $pao.OperationId
        Write-Host "Existing Policy '"$policyType"': "$opid -ForegroundColor Cyan 
    }
    else{
        $opid = New-CsBatchPolicyAssignmentOperation -PolicyType $policyType -PolicyName $policyName -Identity $users.UserPrincipalName -OperationName $opName
        Write-Host "New Policy '" $policyType "': " $opid " - Sleep 60 sec.: please wait..." -ForegroundColor Green 
        Start-Sleep -s 60
    }
        
    $count=1;
    do { 
        $done=$false;
        try{
            $js = Get-CsBatchPolicyAssignmentOperation -OperationId $opid
            $done = ($js.OverallStatus-eq "Completed")                       
            Write-Host "   (Count: "$count") Policy:" $policyType "- Result:" $done "- Completed:" $js.CompletedCount "- ErrorCount:" $js.ErrorCount "- InProgressCount:" $js.InProgressCount "- NotStartedCount:" $js.NotStartedCount -ForegroundColor Gray
	    }
        catch{}
        if(-not($done)){
            Start-Sleep -s 10
            $count++;
        } 
    }
    While (($done -eq $false) -and ($count -le 360)) 
    $time = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
    "END " + $policyType + " - Result: " + $done + " - Completed: " + $js.CompletedCount + " - ErrorCount: " + $js.ErrorCount + " - InProgressCount: " + $js.InProgressCount + " - NotStartedCount: " + $js.NotStartedCount + " - Time: " + $time | Out-File $logFile -Append

}

"END: " + $LogStartTime | Out-File $logFile -Append