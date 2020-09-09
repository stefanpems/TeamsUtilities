<#
  .VERSION AND AUTHOR
    Script version: v-2020.05.06
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
  This script assigns a specified user to multiple groups/teams with the specified role

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

$outDirPath = "C:\Temp" #This is just an example
$adminName = "adminName@schoolName.edu" #This is just an example

$user = "userName@schoolName.edu"
$role = "Member" #Valid options: Member, Owner

$groupsOrTeams = @(
"GroupOrTeam1@schoolName.edu";
"GroupOrTeam2@schoolName.edu";
"GroupOrTeam3@schoolName.edu";
"GroupOrTeam4@schoolName.edu";
"GroupOrTeam5@schoolName.edu")

#########################################################################################################################

cls

if( ($role -ne "Member") -and ($role -ne "Owner") ){
    Wrire-Host "Invialid value for the $role variable. Please specify 'Member' or 'Owner" -ForegroundColor Red
    Exit
}

$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
$outFilePrefixName = "AssignUserToMultipleGroupsOrTeams_"

$outLogFilePath = "$outDirPath\$outFilePrefixName$LogStartTime.log"
If (Test-Path $outLogFilePath){
	Remove-Item $outLogFilePath
}

Connect-AzureAD -AccountId $adminName
Connect-MicrosoftTeams -AccountId $adminName

$line = "User: " + $user; Write-Host $line 

$groupsOrTeams | ForEach-Object {

    $tnn = $_
    
    $gm = $tnn.Substring(0,$tnn.IndexOf("@"))
    
    $group = Get-AzureADGroup -Filter "MailNickname eq '$gm'"
        
    if($group){        
        Add-TeamUser -GroupId $group.ObjectId -User $user -Role $role
         
        $line = "Successfully assigned the user '" + $user + "' to the group or team '" + $group.DisplayName + "' with the role of '" +$role + "'"; Write-Host $line -ForegroundColor Green; $line | Out-File $outLogFilePath -Append
    }
    else{
        $line = "ERROR - User '" + $user + "' - Group or team '" + $group.DisplayName + "' Error: '" + $_.Exception.Message + "'"; Write-Host $line -ForegroundColor Red; $line | Out-File $outLogFilePath -Append
    }   
}
