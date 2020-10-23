<#
  .VERSION AND AUTHOR 
    Script version: v-2020.10.23
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
  This script generates a CSV files with the list of all the users existing in Azure AD/Office 365.
  If requested, the script can create also a second file with the list of all the teams to which the users belongs. 

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

$outCsvDirPath = "C:\Temp\OUT" #"C:\Temp" #Set the correct path
$adminName = "adminName@schoolName.edu" #Set the correct name
$reportAlsoTeams = $true #Set the desider value ($true or $false) 
                         #ATTENTION: when this is set to $true, the execution time is much longer.
                         #           In this case, the script produces 2 CSV files as output
$skipMembershipForTheseUsers = @() #Specify the users, if any, for whom you don't want to read the membership. 
                         #Leave @() if the membership must be read for all the users.
                         #Example:
                         #$skipMembershipForTheseUsers = @(
                         #    "name1@schoolName.edu",
                         #    "name1@schoolName.edu"
                         #    ) 

#########################################################################################################################
cls
$StartTimeAsDate = Get-Date
$StartTimeAsString = $StartTimeAsDate.ToString("yyyy-MM-dd_hh.mm.ss")

$outUsersCsvFilePath = "$outCsvDirPath\DumpUsers_Results_$StartTimeAsString.csv"
If (Test-Path $outUsersCsvFilePath){
	Remove-Item $outUsersCsvFilePath
}
"UPN;DisplayName;GivenName;Surname;AccountEnabled;JobTitle;Department;PhysicalDeliveryOfficeName;UserObjectID" | Out-File $outUsersCsvFilePath
Connect-AzureAD -AccountId $adminName

if($reportAlsoTeams){
    $outUsersTeamsCsvFilePath = "$outCsvDirPath\DumpUsersAndTeams_Results_$StartTimeAsString.csv"
    If (Test-Path $outUsersTeamsCsvFilePath){
	    Remove-Item $outUsersTeamsCsvFilePath
    }
    Connect-MicrosoftTeams -AccountId $adminName
    "UPN;DisplayName;GivenName;Surname;AccountEnabled;JobTitle;Department;PhysicalDeliveryOfficeName;UserObjectID;IsGroupOwner;GroupNickName;GroupDisplayName;ObjectType;IsTeam;TeamVisibility;IsArchivedTeam;GroupObjectID;GroupCreatedDateTime;GroupRenewedDateTime" | Out-File $outUsersTeamsCsvFilePath -Append
}

$groupsList = @{} #HashTable used for caching values 
$teamsList = @{} #HashTable used for caching values

$countU = 0;
$allUsers = Get-AzureADUser -All $true
$numUsers = $allUsers.Count

$allUsers | ForEach-Object{

    $percentComplete = [math]::Round(100*$countU/$numUsers,1)
    Write-Progress -Activity "Scanning users" -Status "$percentComplete % users analyzed" -PercentComplete $percentComplete

    $u = $($_)
    $countU++;
    Write-Host "($countU): User '" $u.UserPrincipalName "'" -ForegroundColor Green

    #Write into the Users CSV
    $u.UserPrincipalName+";"+$u.DisplayName+";"+$u.GivenName+";"+$u.Surname+";"+$u.AccountEnabled+";"+$u.JobTitle+";"+$u.Department+";"+$u.PhysicalDeliveryOfficeName+";"+$u.ObjectId | Out-File $outUsersCsvFilePath -Append       
    
    if( ($reportAlsoTeams) -and (-not($skipMembershipForTheseUsers.Contains($u.UserPrincipalName))) )
    {
        $countG = 0
        $countT = 0

        $AllGroups = Get-AzureADUserMembership -ObjectId $u.ObjectId -All $true
        if($AllGroups){
            $AllGroups | ForEach-Object {
                $g = $($_)
                $countG++;
            
                $groupInfo = @{}

                $gnn = $g.MailNickName
                $gdn = $g.DisplayName
                $IsOwner = $false
                $isTeam = $false
                $teamVisibility = ""
                $isArchivedTeam = $false
                $gc = $null
                $gr = $null

                #Check if it is a Team or a different type of group (security, distribution, office 365)

                if(-not([String]::IsNullOrEmpty($g.ObjectId)) -and ($g.ObjectType -eq "Group") ){
                        
                    if($groupsList.ContainsKey($g.ObjectId)){
                        Write-Host "     ($countU - $countG): Reading group '" $gnn "' ('" $gdn "') from cache" -ForegroundColor gray
                        $groupInfo = $groupsList.Item($g.ObjectId)
                
                        $isTeam = $groupInfo.IsTeam
                        $isArchivedTeam = $groupInfo.IsArchivedTeam
                        $teamVisibility = $groupInfo.TeamVisibility
                        $gm = $groupInfo.Users
                        $gc = $groupInfo.Created
                        $gr = $groupInfo.Renewed

                        if(-not([String]::IsNullOrEmpty($u.UserPrincipalName)) -and $gm.ContainsKey($u.UserPrincipalName)) {
                            $IsOwner = $gm.Item($u.UserPrincipalName)
                        }
                    }
                    else{
                        Write-Host "     ($countU - $countG): Reading and caching info for group '" $gnn "' ('" $gdn "') from Azure" -ForegroundColor cyan

                        try{
                            $gi = Get-AzureADMSGroup -Id $g.ObjectId
                            $gc = $gi.CreatedDateTime.ToString("dd/MM/yyyy")
                            $gr = $gi.RenewedDateTime.ToString("dd/MM/yyyy")
                        }
                        catch{
                            Write-Host "     ($countU - $countG): ERROR while executing Get-AzureADMSGroup for group '" $gnn "' ('" $gdn "') from Azure: " $_.Exception.Message -ForegroundColor Red
                        }

                        try{
                            $t = Get-Team -GroupId $g.ObjectId 
                            $isTeam = $true
                            $isArchivedTeam = $t.Archived
                            $teamVisibility = $t.Visibility
                        }
                        catch{
                            $isTeam = $false
                            $isArchivedTeam = $false
                            $teamVisibility = ""
                        }

                        $users = @{}
                        $IsOwner = $false

                        $ow = $null
                        try{
                            $ow = Get-AzureADGroupOwner -ObjectId $g.ObjectId 
                        }
                        catch{
                            Write-Host "     ($countU - $countG): ERROR while executing Get-AzureADGroupOwner for group '" $gnn "' ('" $gdn "') from Azure: " $_.Exception.Message -ForegroundColor Red
                        }

                        if($ow){        
                            if($ow.Count -gt 0){
                                if($ow.Count -eq 1){
                                    $users.Add($ow.UserPrincipalName,$true)
                                    if($ow.UserPrincipalName.toLower() -eq $u.UserPrincipalName.ToLower()){
                                        $IsOwner = $true
                                    }
                                }
                                else{
                                    $ow | ForEach-Object{
                                        $users.Add($_.UserPrincipalName,$true)

                                        if(-not($IsOwner) -and (($ow.UserPrincipalName.toLower() -eq $u.UserPrincipalName.ToLower()))){
                                            $IsOwner = $true
                                        }
                                    }
                                }
                            }
                        }
    
                        try{
                            Get-AzureADGroupMember -ObjectId $g.ObjectId | ForEach-Object{
                                if(-not($users.ContainsKey($_.UserPrincipalName))){
                                    $users.Add($_.UserPrincipalName,$false)
                                }
                            }
                        }
                        catch{
                            Write-Host "     ($countU - $countG): ERROR while executing Get-AzureADGroupMember for group '$gnn' ('$gdn') from Azure: " $_.Exception.Message -ForegroundColor Red
                        }

                        $groupInfo = @{
                            IsTeam = $isTeam
                            IsArchivedTeam = $isArchivedTeam
                            TeamVisibility = $teamVisibility
                            Users = $users
                            Created = $gc
                            Renewed = $gr
                        }

                        $groupsList.Add($g.ObjectId, $groupInfo)

                    }
                }

                #Write into the Users and Groups CSV
                $u.UserPrincipalName+";"+$u.DisplayName+";"+$u.GivenName+";"+$u.Surname+";"+$u.AccountEnabled+";"+$u.JobTitle+";"+$u.Department+";"+$u.PhysicalDeliveryOfficeName+";"+$u.ObjectId+";"+$IsOwner+";"+$gnn+";"+$gdn+";"+$g.ObjectType+";"+$isTeam+";"+$teamVisibility+";"+$isArchivedTeam+";"+$g.ObjectId+";"+$gc+";"+$gr| Out-File $outUsersTeamsCsvFilePath -Append       
            }
        }
        else{
            $u.UserPrincipalName+";"+$u.DisplayName+";"+$u.GivenName+";"+$u.Surname+";"+$u.AccountEnabled+";"+$u.JobTitle+";"+$u.Department+";"+$u.PhysicalDeliveryOfficeName+";"+$u.ObjectId+";"+$false+";;;;"+$false+";;"+$false+";"| Out-File $outUsersTeamsCsvFilePath -Append       
        }
    }
}
$EndTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"

$EndTimeAsDate = Get-Date
$EndTimeAsString = $EndTimeAsDate.ToString("yyyy-MM-dd_hh.mm.ss")

$ExecDuration = New-TimeSpan -Start $StartTimeAsDate -End $EndTimeAsDate

Write-Host "Execution ended - Duration: " $ExecDuration.Hours " hours and " $ExecDuration.Minutes " minutes." -ForegroundColor DarkBlue -BackgroundColor Cyan
