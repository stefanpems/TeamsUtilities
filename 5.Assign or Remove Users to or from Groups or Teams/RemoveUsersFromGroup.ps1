<#
  .VERSION AND AUTHOR
    Script version: v-2020.09.04
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
  This script remove all users from a specified group or teams

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

$groupName = "GroupName" #This is just an example
$testOnly = $false #Set the desired value

#########################################################################################################################

cls
$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
$outFilePrefixName = "RemoveAllUsersFromGroup_"

$outLogFilePath = "$outDirPath\$outFilePrefixName$LogStartTime.log"
If (Test-Path $outLogFilePath){
	Remove-Item $outLogFilePath
}
Connect-AzureAD -AccountId $adminName
Connect-MicrosoftTeams -AccountId $adminName

#Identify the target group
$g = $null
$g = Get-AzureADGroup -Filter "DisplayName eq '$groupName'"
if(-not($g)){
    $line = "No group found with the specified name: $groupName"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
    exit
}

#Check if it is a Team or a different type of group (security, distribution, office 365)
$isTeam = $true
$allTeamUsers = $null
$allGroupMembers = $null
$extraGroupOwners = $null
try{
    $isTeam = Get-TeamChannel -GroupId $g.ObjectId 
    $line = "The taregt group is a Team: $groupName"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green

    $allTeamUsers = Get-TeamUser -GroupId $g.ObjectId
    $auCount = $allTeamUsers.Count
    $line = "It contains $auCount users"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
    $promptMsg = "Found $auCount users within the specified group of type Team. Remove from the team?"
    $extraTeamUsers = $allTeamUsers
}
catch{
    $isTeam = $false
    $line = "The taregt group is not a Team: $groupName"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green

    $allGroupOwners = Get-AzureADGroupOwner -ObjectId $g.ObjectId -All $true
    $allGroupMembers = Get-AzureADGroupMember -ObjectId $g.ObjectId -All $true
    $aoCount = $allGroupOwners.Count;$amCount = $allGroupMembers.Count
    $line = "It contains $aoCount owners and $amCount members"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
    $promptMsg = "Found $aoCount owners and $amCount members within the specified group. Remove from the group?"
    $extraGroupOwners = $allGroupOwners
    $extraGroupMembers = $allGroupMembers
}

#Prompt for prosecution
$uCount = 0; $mCount = 0; $oCount = 0; $xCount = 0; $sCount = 0;
$uCount = $usrs.Count

$promptTitle = "Remove Users From Group"
$promptOptions= echo Yes No
$promptDefault = 1
$promptResponse = $Host.UI.PromptForChoice($promptTitle, $promptMsg, $promptOptions, $promptDefault)
if ($promptResponse -eq 0) {
    if($isTeam){
        $allTeamUsers | ForEach-Object{
            $u = $_
            if(($u.ObjectType -eq "User") -and -not([String]::IsNullOrEmpty($u.UserPrincipalName))){
                if(-not($testOnly)){
                    try{
                        Remove-TeamUser -GroupId $g.ObjectId -User $u
                        $line = "Successfully removed user from team: " + $u.UserPrincipalName; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkGreen -BackgroundColor Green
                    }
                    catch{
                        $line = "ERROR while removing user from team: " + $u.UserPrincipalName + " - " + $_.Exception.Message; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkGreen -BackgroundColor Red
                    }
                }
                else{
                    $line = "[SIMULATED] Removed user from team: " + $u.UserPrincipalName; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkGreen -BackgroundColor Green
                }
            }
            else{
                $line = "[Skipped] This team member is not a user: " + $u.DisplayName; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkGreen -BackgroundColor Gray
            }
        }
    }
    else{
        $allGroupMembers | ForEach-Object{
            $u = $_
            if(($u.ObjectType -eq "User") -and -not([String]::IsNullOrEmpty($u.UserPrincipalName))){
                if(-not($testOnly)){
                    try{
                        Remove-AzureADGroupMember -ObjectId $g.ObjectId -MemberId $u.ObjectId
                        $line = "Successfully removed member from group: " + $u.UserPrincipalName; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkGreen -BackgroundColor Green
                    }
                    catch{
                        $line = "ERROR while removing user from team: " + $u.UserPrincipalName + " - " + $_.Exception.Message; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkGreen -BackgroundColor Red
                    }
                }
                else{
                    $line = "[SIMULATED] Removed member from group: " + $u.UserPrincipalName; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkGreen -BackgroundColor Green
                }
            }
            else{
                $line = "[Skipped] This group member is not a user: " + $u.DisplayName; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkGreen -BackgroundColor Gray
            }
        }
        $allGroupOwners | ForEach-Object{
            $u = $_
            if(($u.ObjectType -eq "User") -and -not([String]::IsNullOrEmpty($u.UserPrincipalName))){
                if(-not($testOnly)){
                    try{
                        Remove-AzureADGroupOwner -ObjectId $g.ObjectId -OwnerId $u.ObjectId
                        $line = "Successfully removed owner from group: " + $u.UserPrincipalName; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkGreen -BackgroundColor Green
                    }
                    catch{
                        $line = "ERROR while removing user from team: " + $u.UserPrincipalName + " - " + $_.Exception.Message; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkGreen -BackgroundColor Red
                    }
                }
                else{
                    $line = "[SIMULATED] Removed owner from group: " + $u.UserPrincipalName; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkGreen -BackgroundColor Green
                }
            }
            else{
                $line = "[Skipped] This group owner is not a user: " + $u.DisplayName; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkGreen -BackgroundColor Gray
            }
        }
    }    
}
else{
    $line = "Execution stopped as requested by the user"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
    exit
}


$line = "Execution ended"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor White    
$line = "---> Generated the log file '$outLogFilePath'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Black -BackgroundColor White
