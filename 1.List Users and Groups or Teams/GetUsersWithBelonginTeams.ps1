<#
  .VERSION AND AUTHOR
    Script version: v-2020.04.11
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
  This script generates a CSV files with the list of all the users existing in Azure AD/Office 365.
  If requested, the script can list all the teams to which the users belongs.

  .VARIABLES TO BE SET
    See below

  .PREREQUISITES
  * Use Windows 10 (For earlier versions of Windows, please refer to https://docs.microsoft.com/en-us/microsoft-365/enterprise/connect-to-all-microsoft-365-services-in-a-single-windows-powershell-window?view=o365-worldwide#before-you-begin)
  * If not alreayd done, install the PowerShell modules Azure AD (or AzureADPreview) and Teams

    To install the PowerShell module, open PowerShell by using the option "Run as administrator" and type:
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

$outCsvDirPath = "C:\Temp" #Set the correct path
$adminName = "nomeutente@nomescuola.edu.it" #Set the correct name
$reportAlsoTeams = $false #Set the desider value ($true or $false) 
                          #ATTENTION: when this is set to $true, the execution time is much longer.
                          #           In this case, the script produces 2 CSV files as output

#########################################################################################################################
$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"

$outUsersCsvFilePath = "$outCsvDirPath\DumpUsers_Results_$LogStartTime.csv"
If (Test-Path $outUsersCsvFilePath){
	Remove-Item $outUsersCsvFilePath
}
"UPN;DisplayName;GivenName;Surname;JobTitle;Department;PhysicalDeliveryOfficeName" | Out-File $outUsersCsvFilePath
Connect-AzureAD -AccountId $adminName

if($reportAlsoTeams){
    $outUsersTeamsCsvFilePath = "$outCsvDirPath\DumpUsersAndTeams_Results_$LogStartTime.csv"
    If (Test-Path $outUsersTeamsCsvFilePath){
	    Remove-Item $outUsersTeamsCsvFilePath
    }
    "UPN;DisplayName;GivenName;Surname;JobTitle;Department;PhysicalDeliveryOfficeName;TeamNickName;TeamDisplayName;TeamVisibility;UserRoleInTeam" | Out-File $outUsersTeamsCsvFilePath
    Connect-MicrosoftTeams -AccountId $adminName
}
 
$countU = 0;
Get-AzureADUser -All $true | ForEach-Object{
    $u = $($_)
    $countU++;
    Write-Host "($countU): User '" $u.UserPrincipalName "'" -ForegroundColor Green
    $u.UserPrincipalName+";"+$u.DisplayName+";"+$u.GivenName+";"+$u.Surname+";"+$u.JobTitle+";"+$u.Department+";"+$u.PhysicalDeliveryOfficeName | Out-File $outUsersCsvFilePath -Append

    if($reportAlsoTeams){
        $countT = 0
        Get-Team -User $u.UserPrincipalName | ForEach-Object {
            $t = $($_)
            $countT++;
            $gnn = $t.MailNickName
            $gdn = $t.DisplayName
            Write-Host "     ($countU - $countT): Team '$gnn' - '$gdn'" -ForegroundColor Cyan
            $uit = Get-TeamUser -GroupId $t.GroupId | where User -eq $u.UserPrincipalName
            $u.UserPrincipalName+";"+$u.DisplayName+";"+$u.GivenName+";"+$u.Surname+";"+$u.JobTitle+";"+$u.Department+";"+$u.PhysicalDeliveryOfficeName+";"+$gnn+";"+$gdn+";"+$t.Visibility+";"+$uit.Role | Out-File $outUsersTeamsCsvFilePath -Append       
        }
    }
}
