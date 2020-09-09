<#
  .VERSION AND AUTHOR
    Script version: v-2020.04.29
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
  This script aligns the membership of a group in Azure Active Directory with the results 
  of a query on the Azure Active Directory users based on user attributes.
  The target group can be of any type: security, distribution, Office 365 or a Team.
  The script is useful where Azure AD Dynamic Groups are not available (e.g. for A1 licenses)

  .VARIABLES TO BE SET
    $outDirPath: path of the local folder where the script generates the output log file.
    $adminName: name of the administrative account to be used for the script execution (the password will be prompted).
    $userFilter: oData v3.0 filter statement to identify the users.
    $groupName: name of the group where the identifed users should be added as members or owners.
    $addAllUsersAsGroupOwners: set to $true if the identifed users should be added as owners; set to $false if they should be added as members.
    $specificOwners: (ignored if $addAllUsersAsGroupOwners = $true) array of the login names of the users to be added as owners, separated by ";". 
                     Set to $null if empty. Only users found in the query are considered; others are ignored
    $cleanExtraUsers: remove extra users found in the team or group (members not resulting from the query)
    $testOnly: set to $true if the script should only simulate the execution with no real group membership change.

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
$adminName = "adminName@schoolName.onmicrosoft.com" #This is just an example
$userFilter = "JobTitle eq 'Docente' or startswith(JobTitle,'Dirigente')" #This is just an example
$groupName = "TargetGroupOrTeamName" #This is just an example
$addAllUsersAsGroupOwners = $false #Please choose
$specificOwners = @("user1@schoolname.edu.it";"user2@schoolname.edu.it") #This is just an example. Ignored if $addAllUsersAsGroupOwners = $true
$cleanExtraUsers = $true
$testOnly = $false

#########################################################################################################################

cls
$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
$outFilePrefixName = "AssignUsersToGroup_"

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
try{
    $isTeam = Get-TeamChannel -GroupId $g.ObjectId 
    $line = "The taregt group is a Team: $groupName"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green

    $allTeamUsers = Get-TeamUser -GroupId $g.ObjectId
    $auCount = $allTeamUsers.Count
    $line = "It contains $auCount users"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
    $extraTeamUsers = $allTeamUsers
}
catch{
    $isTeam = $false
    $line = "The taregt group is not a Team: $groupName"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green

    $allGroupOwners = Get-AzureADGroupOwner -ObjectId $g.ObjectId -All $true
    $allGroupMembers = Get-AzureADGroupMember -ObjectId $g.ObjectId -All $true
    $aoCount = $allGroupOwners.Count;$amCount = $allGroupMembers.Count
    $line = "It contains $aoCount owners and $amCount members"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
    $extraGroupOwners = $allGroupOwners
    $extraGroupMembers = $allGroupMembers
}

#Identify the users based on the filter
$usrs = $null
$usrs = Get-AzureADUser -Filter $userFilter -All $true
if(-not($usrs)){
    $line = "No user found with the specified query: $userFilter"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
    exit
}

#Prompt for prosecution
$uCount = 0; $mCount = 0; $oCount = 0; $xCount = 0; $sCount = 0;
$uCount = $usrs.Count
$line = "Found $uCount users with the specified query: $userFilter"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkGreen -BackgroundColor Yellow

$promptTitle = "Assign Users To Group"
$promptMsg = "Found $uCount users with the specified query. Assign to group?"
$promptOptions= echo Yes No
$promptDefault = 1
$promptResponse = $Host.UI.PromptForChoice($promptTitle, $promptMsg, $promptOptions, $promptDefault)
if ($promptResponse -eq 0) {
    $usrs | ForEach-Object{
        $xCount++;
        $uo = $_
        $upn = $uo.userPrincipalName
        $uFoundInTeam = $false;
        $uFoundInGroupAsOwner = $false;
        $uFoundInGroupAsMember = $false;
        
        try{
            if($isTeam){
                $uFoundInTeam = $allTeamUsers | Where {$_.User -EQ $upn}
                $extraTeamUsers = $extraTeamUsers | where {$_.User -ne $upn }

                if($addAllUsersAsGroupOwners){
                    if($uFoundInTeam){                         
                        if($uFoundInTeam.Role -eq "owner"){
                            $sCount++;
                            $line = "$xCount - User '$upn' skipped: already owner of the group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Gray
                        }
                        else{
                            $oCount++;
                            if(-not($testOnly)){
                                $line = "$xCount - Removing the user '$upn' found as member in group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
                                #CHANGE!
                                Remove-TeamUser -GroupId $g.ObjectId -User $upn
                            }
                            else{
                                $line = "[Simulation] $xCount - Removing the user '$upn' found as member in group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
                            }
                            if(-not($testOnly)){
                                $line = "$xCount - Adding the user '$upn' into the group '$groupName' as owner"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Magenta
                                #CHANGE!
                                Add-TeamUser -GroupId $g.ObjectId -User $upn -Role Owner
                            }
                            else{
                                $line = "[Simulation] $xCount - Adding the user '$upn' into the group '$groupName' as owner"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Magenta
                            }
                        }
                    }
                    else{
                        $oCount++;
                        if(-not($testOnly)){
                            $line = "$xCount - Adding the user '$upn' into the group '$groupName' as owner"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
                            #CHANGE!
                            Add-TeamUser -GroupId $g.ObjectId -User $upn -Role Owner
                        }
                        else{
                            $line = "[Simulation] $xCount - Adding the user '$upn' into the group '$groupName' as owner"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
                        }
                    }
                }
                else{
                    if($upn -in $specificOwners){
                        if($uFoundInTeam){
                            if($uFoundInTeam.Role -eq "owner"){
                                $sCount++;
                                $line = "$xCount - User '$upn' skipped: already owner of the group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Gray
                            }
                            else{
                                $oCount++;
                                if(-not($testOnly)){
                                    $line = "$xCount - Removing the user '$upn' found as member in group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
                                    #CHANGE!
                                    Remove-TeamUser -GroupId $g.ObjectId -User $upn
                                }
                                else{
                                    $line = "[Simulation] $xCount - Removing the user '$upn' found as member in group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
                                }
                                if(-not($testOnly)){
                                    $line = "$xCount - Adding the user '$upn' into the group '$groupName' as owner"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
                                    #CHANGE!
                                    Add-TeamUser -GroupId $g.ObjectId -User $upn -Role Owner
                                }
                                else{
                                    $line = "[Simulation] $xCount - Adding the user '$upn' into the group '$groupName' as owner"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
                                }
                            }
                        }
                        else{
                            $oCount++;
                            if(-not($testOnly)){
                                $line = "$xCount - Adding the user '$upn' into the group '$groupName' as owner"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Magenta
                                #CHANGE!
                                Add-TeamUser -GroupId $g.ObjectId -User $upn -Role Owner
                            }
                            else{
                                $line = "[Simulation] $xCount - Adding the user '$upn' into the group '$groupName' as owner"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Magenta
                            }
                        }
                    }
                    else{
                        if($uFoundInTeam){
                            if($uFoundInTeam.Role -eq "member"){
                                $sCount++;
                                $line = "$xCount - User '$upn' skipped: already member of the group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Gray
                            }
                            else{
                                $mCount++;
                                if(-not($testOnly)){
                                    $line = "$xCount - Removing the user '$upn' found as owner in group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
                                    #CHANGE!
                                    Remove-TeamUser -GroupId $g.ObjectId -User $upn
                                }
                                else{
                                    $line = "[Simulation] $xCount - Removing the user '$upn' found as owner in group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
                                }
                                if(-not($testOnly)){
                                    $line = "$xCount - Adding the user '$upn' into the group '$groupName' as member"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
                                    #CHANGE!
                                    Add-TeamUser -GroupId $g.ObjectId -User $upn -Role Member
                                }
                                else{
                                    $line = "[Simulation] $xCount - Adding the user '$upn' into the group '$groupName' as member"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
                                }
                            }
                        }
                        else{
                            $mCount++;
                            if(-not($testOnly)){
                                $line = "$xCount - Adding the user '$upn' into the group '$groupName' as member"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
                                #CHANGE!
                                Add-TeamUser -GroupId $g.ObjectId -User $upn -Role Member
                            }
                            else{
                                $line = "[Simulation] $xCount - Adding the user '$upn' into the group '$groupName' as member"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
                            }
                        }
                    }
                }
            }
            else{
                $uFoundInGroupAsOwner = $allGroupOwners | Where {$_.UserPrincipalName -EQ $upn}
                $uFoundInGroupAsMember = $allGroupMembers | Where {$_.UserPrincipalName -EQ $upn}
                $extraGroupOwners = $extraGroupOwners | where {$_.UserPrincipalName -ne $upn }
                $extraGroupMembers = $extraGroupMembers | where {$_.UserPrincipalName -ne $upn }

                if($addAllUsersAsGroupOwners){
                    if($uFoundInGroupAsOwner){                        
                        $sCount++;
                        $line = "$xCount - User '$upn' skipped: already owner of the group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Gray
                    }
                    else{
                        $oCount++;                        
                        if(-not($testOnly)){
                            $line = "$xCount - Adding the user '$upn' into the group '$groupName' as owner"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Magenta
                            #CHANGE!
                            Add-AzureADGroupOwner -ObjectId $g.ObjectId -RefObjectId $uo.ObjectId
                        }
                        else{
                            $line = "[Simulation] $xCount - Adding the user '$upn' into the group '$groupName' as owner"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Magenta
                        }
                    }
                }
                else{
                    if($upn -in $specificOwners){
                        if($uFoundInGroupAsOwner){                        
                            $sCount++;
                            $line = "$xCount - User '$upn' skipped: already owner of the group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Gray
                        }
                        else{
                            $oCount++;
                            if(-not($testOnly)){
                                $line = "$xCount - Adding the user '$upn' into the group '$groupName' as owner"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Magenta
                                #CHANGE!
                                Add-AzureADGroupOwner -ObjectId $g.ObjectId -RefObjectId $uo.ObjectId
                            }
                            else{
                                $line = "[Simulation] $xCount - Adding the user '$upn' into the group '$groupName' as owner"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Magenta
                            }
                        }
                    }
                    else{                        
                        if($uFoundInGroupAsMember){                        
                            $sCount++;
                            $line = "$xCount - User '$upn' skipped: already member of the group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Gray
                        }
                        else{
                            $mCount++;
                            if(-not($testOnly)){
                                $line = "$xCount - Adding the user '$upn' into the group '$groupName' as member"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
                                #CHANGE!
                                Add-AzureADGroupMember -ObjectId $g.ObjectId -RefObjectId $uo.ObjectId
                            }
                            else{
                                $line = "[Simulation] $xCount - Adding the user '$upn' into the group '$groupName' as member"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
                            }
                        }
                    }
                }
            }
        }
        catch{
            if($_.Exception.Message -contains "One or more added object references already exist for the following modified properties: 'members'"){
                $line = "   Ignored exception"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan    
            }
            else{
                $line = "Error: $_"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red    
            #throw $_
            }
        }
    }   
}
else{
    $line = "Execution stopped as requested by the user"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
    exit
}

$line = "   "; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan    

$etu = 0;$ego = 0; $egm = 0
if($isTeam){
    $etu = $extraTeamUsers.Count
    $newPromptMsg = "Remove the $etu extra team users?"
    $line = "Extra team users:"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor White    
    $extraTeamUsers | ForEach-Object{
        $uex = $_.User
        $line = "   $uex"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor White    
    }
}
else{
    $ego = $extraGroupOwners.Count
    $egm = $extraGroupMembers.Count
    $newPromptMsg = "Remove the $ego extra group owners and $egm group members?"
    $line = "Extra group owners:"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor White     
    $extraGroupOwners | ForEach-Object{
        $ugo = $_.UserPrincipalName
        $line = "   $ugo"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor White    
    }
    $line = "Extra group members:"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor White     
    $extraGroupMembers| ForEach-Object{
        $ugm = $_.UserPrincipalName
        $line = "   $ugm"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor White    
    }
}

if($cleanExtraUsers){
    if(($etu+$ego+$egm)-gt 0){
        #Prompt for prosecution
        $dCount = 0;
        $newPromptTitle = "Remove extra users"
        $promptOptions= echo Yes No
        $promptDefault = 1
        $promptResponse = $Host.UI.PromptForChoice($newPromptTitle, $newPromptMsg, $promptOptions, $promptDefault)
                                                                                                                                                                                    if ($promptResponse -eq 0) {
        if($isTeam){
            $extraTeamUsers | ForEach-Object{ 
                $dCount++;       
                $upn = $_.User
                if(-not($testOnly)){
                    $line = "$dCount - Removing the extra user '$upn' found in the team '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
                    #CHANGE!
                    Remove-TeamUser -GroupId $g.ObjectId -User $upn
                }
                else{
                    $line = "[Simulation] $dCount - Removing the extra user '$upn' found in the team '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
                }
            }
        }
        else{
            $extraGroupOwners | ForEach-Object{
                $dCount++;
                $upn = $_.UserPrincipalName
                $uid = $_.ObjectId
                if(-not($testOnly)){
                    $line = "$dCount - Removing the extra owner '$upn' found in the group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
                    #CHANGE!
                    Remove-AzureADGroupOwner -ObjectId $g.ObjectId -OwnerId $uid
                }
                else{
                    $line = "[Simulation] $dCount - Removing the extra owner '$upn' found in the group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
                }
            }
            $dCount=0;
            $extraGroupMembers | ForEach-Object{
                $dCount++;
                $upn = $_.UserPrincipalName
                $uid = $_.ObjectId
                if(-not($testOnly)){
                    $line = "$dCount - Removing the extra member '$upn' found in the group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
                    #CHANGE!
                    Remove-AzureADGroupMember -ObjectId $g.ObjectId -MemberId $uid
                }
                else{
                    $line = "[Simulation] $dCount - Removing the extra member '$upn' found in the group '$groupName'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
                }
            }

        }
        }
        else{
            $line = "No removal executed"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow    
        }
    }
}

$line = "   "; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan    
$line = "* Number of users skipped because already memebers or owners: $sCount"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkBlue -BackgroundColor Yellow
$line = "* Number of users added to the group as members: $mCount"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkGreen -BackgroundColor Yellow
$line = "* Number of users added to the group as owners: $oCount"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkMagenta -BackgroundColor Yellow

if($isTeam){
    $line = "* Number of users other users in Team: $etu"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkBlue -BackgroundColor Yellow
}
else{
    $line = "* Number of users extra owners in Group: $ego"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkBlue -BackgroundColor Yellow
    $line = "* Number of users extra members in Group: $egm"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor DarkBlue -BackgroundColor Yellow
}





$line = "---> Generated the log file '$outLogFilePath'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Black -BackgroundColor Yellow
