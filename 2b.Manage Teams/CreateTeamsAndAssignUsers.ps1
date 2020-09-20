<#
  .VERSION AND AUTHOR
    Script version: v-2020.09.20
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
    This script massively creates Teams and assign to them members and owners as specified in the 2 input CSV files.
    For each specified team, if it already exsists, the script only update the membership.
    For each specified member/owner, if she/he has already the required role, no action is done.
    The script generate 3 output files:
     * A verbose log
     * A verbose csv with all the actions done and their results
     * A summary csv with all the teams for which at least an error was recorded; this file is useful if you want to run the script again only for the failed teams.

  .PREREQUISITES
   * Use Windows 10 (For earlier versions of Windows, please refer to https://docs.microsoft.com/en-us/microsoft-365/enterprise/connect-to-all-microsoft-365-services-in-a-single-windows-powershell-window?view=o365-worldwide#before-you-begin)
   * If not alreayd done, install the PowerShell modules Azure AD (or AzureADPreview) and SharePointPnPPowerShellOnline

    To install the PowerShell modules, open PowerShell by using the option "Run as administrator" and type:
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
        Install-Module -Name AzureAD
            or
        Install-Module -Name AzureADPreview
            and 
        Install-Module SharePointPnPPowerShellOnline
    Details here: https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0
    and here: https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-all-office-365-services-in-a-single-windows-powershell-window

  .IMPORTANT NOTE 
  Every hour the script produce a credential prompt to be manually confirmed (the execution cannot be completely unattended!)

  .VARIABLES TO BE SET
  Read below.
#>

#########################################################################################################################
# VARIABLES TO BE SET:
#########################################################################################################################

$csvTeams = "C:\Temp\IN\Team.csv"           #EXPECTED COLUMNS: "Team NN","Team DN" -> Team NickName and Team Display Name
                                            #NOTE: In the Team NN do not use special characters. 
                                            #      For schools, we recommend to add the school year as prefix.
                                            #      Example of row: "as2021-1A-italian;1A Italian Language and Culture"
$csvUsers = "C:\Temp\IN\Team-User.csv"      #EXPECTED COLUMNS: "Team NN","User","Role" -> Team NickName, LoginName of the User to be added, Role for the user to be added
                                            #NOTE: the expected values in the column "Role" are "Member" or "Owner".
                                            #      Example of row: "as2021-1A-italian;xxx@schoolName.edu;Member"
$delimiter = ";" #Set the correct value 
$outLogDir = "C:\Temp\OUT" #Set the correct path!

$templateName = "educationClass" 
    #Known values:
    # -> Set "educationClass" for creating teams with the template "Class"
    # -> Set "educationStaff" for creating teams with the template "Staff" ("Personale" in Italian)
    #Details in https://support.microsoft.com/en-us/office/choose-a-team-type-to-collaborate-in-microsoft-teams-0a971053-d640-4555-9fd7-f785c2b99e67

$adminName = "adminName@schoolName.edu" #Set the correct name! NOTE: must be pre-assigned a license including Teams!
    #RECOMMENDATION: do not use a "personal" account as $adminName. This name appears in the first screen of every 
    #created Teams.
$testOnly = $true #Set the desired value

#########################################################################################################################

function CreateEduTeam{

    param(
         [string]
         $TeamName, 
         [string]
         $TeamDescription, 
         [string]
         $TeamTemplate,
         [string]
         $accessToken
    )

    $teamPreferences = @{
        "GuestAllowCreateUpdateChannels" = $false;
        "GuestAllowDeleteChannels" = $false;
        "AllowGiphy" = $true;
        "GiphyContentRating" = "Strict";
        "AllowStickersAndMemes" = $true;
        "AllowCustomMemes" = $true;
        "AllowCreateUpdateChannels" = $false;
        "AllowDeleteChannels" = $false;
        "AllowAddRemoveApps" = $false;
        "AllowCreateUpdateRemoveTabs" = $false;
        "AllowCreateUpdateRemoveConnectors" = $false;
        "AllowUserEditMessages" = $true;
        "AllowUserDeleteMessages" = $true;
        "AllowOwnerDeleteMessages" = $true;
        "AllowTeamMentions" = $true;
        "AllowChannelMentions" = $true;
        }



    #Prepare generic OAuth Bearer token header
    $headers = @{
	    "Content-Type" = "application/json"
	    Authorization = "Bearer $accessToken"
    }

    $GraphURL = "https://graph.microsoft.com/beta" 
    $graphV1Endpoint = "https://graph.microsoft.com/v1.0" 

    try
    {
        # Create the team 
        $getTeamFromGraphUrl = "$GraphURL/groups?`$filter=displayName eq '" + $TeamName + "'" 
        $createTeamRequest = @{
            "template@odata.bind" = "https://graph.microsoft.com/beta/teamsTemplates('"+$TeamTemplate+"')"
            displayName = $TeamName
            description =  $TeamDescription
	        memberSettings = @{
		        allowCreateUpdateChannels = $teamPreferences["GuestAllowCreateUpdateChannels"]
                allowDeleteChannels = $teamPreferences["AllowDeleteChannels"]
                allowAddRemoveApps = $teamPreferences["AllowAddRemoveApps"]
                allowCreateUpdateRemoveTabs = $teamPreferences["AllowCreateUpdateRemoveTabs"]
                allowCreateUpdateRemoveConnectors = $teamPreferences["AllowCreateUpdateRemoveConnectors"]
	        }
	        messagingSettings = @{
		        allowUserEditMessages = $teamPreferences["AllowUserEditMessages"]
		        allowUserDeleteMessages = $teamPreferences["AllowUserDeleteMessages"]
                allowOwnerDeleteMessages = $teamPreferences["AllowOwnerDeleteMessages"]
                allowTeamMentions = $teamPreferences["AllowTeamMentions"]
                allowChannelMentions = $teamPreferences["AllowChannelMentions"]
                    
	        }
	        funSettings = @{
		        allowGiphy = $teamPreferences["AllowGiphy"]
		        giphyContentRating = $teamPreferences["GiphyContentRating"]
                allowStickersAndMemes = $teamPreferences["AllowStickersAndMemes"]
                allowCustomMemes = $teamPreferences["AllowCustomMemes"]
	        }
        }

    
        $createTeamBody = ConvertTo-Json -InputObject $createTeamRequest
    
        $TeamCreationResponse = Invoke-RestMethod -Uri https://graph.microsoft.com/beta/teams -Body $createTeamBody -ContentType "application/json" -Headers @{Authorization = "Bearer $accesstoken"} -Method Post -Verbose -UseBasicParsing

        Start-Sleep -s 2
        $count = 1
        $Stoploop = $false 
        do { 
            $TeamCreationResponse = Invoke-RestMethod -Uri $getTeamFromGraphUrl -Headers @{Authorization = "Bearer $accesstoken"} -Method Get -Verbose   
            $orderedTeams = $TeamCreationResponse.value | Sort-Object -Property createdDateTime -Descending
                
            if($orderedTeams -ne $null){ 
                $Stoploop = $true 
                $orderedTeams = $TeamCreationResponse.value | Sort-Object -Property createdDateTime -Descending
                $TeamID = $orderedTeams.id
            }
            else
            {
                write-host "  Wait... tentative $count/50"
                Start-Sleep -s 2
                $count = $count + 1
            } 
        } 
        While (($Stoploop -eq $false) -or ($count -eq 50))           
            

        #write-host $TeamID
        if ($TeamID -eq $null)
        {
            throw "Could not retrive the new Team"
        }        
            
        Write-Host "Team $TeamName with ID $TeamID has been created successfully..." -ForegroundColor Green
            
    }
    catch
    {
        Write-Host "Error while creating team " $d.TeamName " - " $_.Exception.Message -ForegroundColor Red
        "  Error while creating team '$d.TeamName': '$_'" | Out-File $oLogFile -Append
        throw $_
    }

}


#########################################################################################################################
cls
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue 
Connect-MicrosoftTeams -AccountId $adminName
Connect-AzureAD -AccountId $adminName
$arrayOfScopes = @("Group.Read.All","Group.ReadWrite.All","User.ReadWrite.All", "Directory.Read.All","Reports.Read.All") 

Connect-PnPOnline -Scopes $arrayOfScopes 
$accessToken = Get-PnPAccessToken -Decoded

$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
if($testOnly){
	$oLogFile = "$outLogDir\SIMULATION-CreateTeams_$LogStartTime.log"
    $oCsvFile = "$outLogDir\SIMULATION-CreateTeams_$LogStartTime.csv"
}
else{
	$oLogFile = "$outLogDir\CreateTeams_$LogStartTime.log"
    $oCsvFile = "$outLogDir\CreateTeams_$LogStartTime.csv"
}
If (Test-Path $oLogFile)
{
	Remove-Item $oLogFile
}
If (Test-Path $oCsvFile)
{
	Remove-Item $oCsvFile
}
if($testOnly){
	"SIMULATED EXECUTION STARTED - $LogStartTime" | Out-File $oLogFile     
}
else{
	"EXECUTION STARTED - $LogStartTime" | Out-File $oLogFile 
}
Write-Host "Creato il file di log '$oFile'" -ForegroundColor Yellow
"Target;Action;Result;Details" | Out-File $oCsvFile 

$errCount = 0;
$TeamsToBeRepeated = @{}


#Lettura CSV dei Team
$teamsRows  = Import-Csv -Path $csvTeams -Delimiter $delimiter
$numRows = $teamsRows.Count; if($numRows -eq $null) {$numRows = 1}
$currRow = 0

:TeamsLoop
ForEach ($teamRow in $teamsRows){    

    #Progress bar
    $currRow++;
    $percentComplete = [math]::Round(100*$currRow/$numRows,1)
    Write-Progress -Activity "Scanning rows" -Status "$percentComplete % rows processed" -PercentComplete $percentComplete

    #Searching for the Team 
    $tnn = $teamRow."Team NN".Trim()
    $tdn = $teamRow."Team DN".Trim()
        
    Write-Host "Searching for the team: " $tnn "-" $tdn
    "Searching for the team: $tnn - $tdn" | Out-File $oLogFile -Append

    try{

        $group = Get-Team -MailNickname $tnn

        if($group -eq $null){
            $tnn+";"+"Get-Team;Success;Not found" | Out-File $oCsvFile -Append

            $tokenExpiration = $accessToken.ValidTo.ToLocalTime().AddMinutes(-1);
            $TimeToExpiry = $tokenExpiration - (Get-Date)        
            $sTimeToExpiry = $TimeToExpiry.Minutes.ToString() + " min " + $TimeToExpiry.Seconds.ToString() + " sec"
                
            Write-Host "  Token - TimeToExpiry: " $sTimeToExpiry
            "  Token - TimeToExpiry: '$sTimeToExpiry'" | Out-File $oLogFile -Append            

            $IsExpired = (Get-Date) -gt $tokenExpiration   

            if($IsExpired){
                Write-Host "  Token expired! Acquiring a new token" -ForegroundColor Cyan 
                "  Token expired! Acquiring a new token." | Out-File $oLogFile -Append     
            
                $accessToken = $null;
            
                try{
                    Connect-PnPOnline -Scopes $arrayOfScopes 
                    $accessToken = Get-PnPAccessToken -Decoded      
                }
                catch{
                    Write-Host "  Could not acquire a new token. Forced exit!" -ForegroundColor Red 
                    "  Could not acquire a new token. Forced exit!" | Out-File $oLogFile -Append     
                    Break TeamsLoop
                }

                $tokenExpiration = $accessToken.ValidTo.ToLocalTime();
                $TimeToExpiry = $tokenExpiration - (Get-Date)        
                $sTimeToExpiry = $TimeToExpiry.Minutes.ToString() + " min " + $TimeToExpiry.Seconds.ToString() + " sec"
                
                Write-Host "  New Token - TimeToExpiry: " $sTimeToExpiry
                "  New Token - TimeToExpiry: '$sTimeToExpiry'" | Out-File $oLogFile -Append            

            }
     
            if(-not($testOnly)){
                #CHANGE - Provisioning of the new Team by using GraphAPI
                try{
                    CreateEduTeam -TeamName $tnn -TeamDescription $tdn -TeamTemplate $templateName -accessToken $accessToken.RawData
                    $tnn+";"+"CreateEduTeam;Success;"+$tdn | Out-File $oCsvFile -Append
                }
                catch{
                    $tnn+";"+"CreateEduTeam;Error;" + $_.Exception.Message | Out-File $oCsvFile -Append
                    if(-not($TeamsToBeRepeated.ContainsKey($tnn))){
                        $TeamsToBeRepeated.Add($tnn,$tdn)
                    }
                    $errCount++;
                    Write-Host "ERROR while creating the team '" $tnn "': " $_.Exception.Message "'" -ForegroundColor Red 
                    "  ERROR while creating the team: '$tnn'" | Out-File $oLogFile -Append

                    throw
                }

                #Accessing the new team 
                $group = Get-Team -MailNickname $tnn
                if($group){
                    $tnn+";"+"Get-Team;Success;New team" | Out-File $oCsvFile -Append
                    $gnn = $group.MailNickName
                    Write-Host "  New team created and successfully accessed: " $gnn
                    "  New team created and successfully accessed: '$gnn'" | Out-File $oLogFile -Append
                }
                else{
                    $tnn+";"+"Get-Team;Error;New team" | Out-File $oCsvFile -Append
                    if(-not($TeamsToBeRepeated.ContainsKey($tnn))){
                        $TeamsToBeRepeated.Add($tnn,$tdn)
                    }
                    $errCount++;
                    Write-Host "ERROR while searching the new team: " $tnn -ForegroundColor Red 
                    "  ERROR while searching the new team: '$tnn'" | Out-File $oLogFile -Append
                
                    throw
                }
            }
            else{
                $tnn+";"+"CreateEduTeam;Success;Simulated" | Out-File $oCsvFile -Append
                Write-Host "  Simulated - Created the new team: " $tnn
                "  Simulated - Created the new team: $tnn" | Out-File $oLogFile -Append
            }
        }
        else{
            $tnn+";"+"Get-Team;Success;Already existing" | Out-File $oCsvFile -Append
            Write-Host "  The team already exists: " $group.MailNickName
            "  The team already exists: $tnn" | Out-File $oLogFile -Append
        }
    
    }
    catch{
        $tnn+";"+"Get-Team;Error;" + $_.Exception.Message | Out-File $oCsvFile -Append
        if(-not($TeamsToBeRepeated.ContainsKey($tnn))){
            $TeamsToBeRepeated.Add($tnn,$tdn)
        }
        Write-Host "Skipping the team '" $tnn "' " -ForegroundColor Red 
        "  Skipping the team: '$tnn'" | Out-File $oLogFile -Append

        continue TeamsLoop #STOP executing the following actions for this team!
    }

    #Reading the Team Users
    $usersRows = Import-Csv -Path $csvUsers  -Delimiter $delimiter | 
    Where-Object -Property "Team NN" -eq $tnn 

    :UsersLoop
    Foreach($userRow in $usersRows){           
        $uUPN = $userRow."User".Trim()
        $uR = $userRow."Role".Trim()
                
        try{
            $R = ""
            if($uR.ToLower() -eq "member") { $R = "Member" }
            if($uR.ToLower() -eq "owner") { $R = "Owner" }
            if($R -eq "") {Throw "Invalid role specified for the user "+$uUPN+" in the input file '"+$csvUsers+"'"}

            $uObj = Get-AzureADUser -Filter "UserPrincipalName eq '$uUPN'"

            if($uObj){
                $uUPN+";"+"Get-AzureADUser;Success;UsersLoop" | Out-File $oCsvFile -Append

                Write-Host "  Setting the user as team " $R ": " $uUPN -ForegroundColor Green
                "  Setting the user as team $R : $uUPN" | Out-File $oLogFile -Append

                if(-not($testOnly)){
                    #CHANGE - Setting the user as team member
                    Add-TeamUser -GroupId $group.GroupId -User $uUPN -Role $R
                    $uUPN+";"+"Add-TeamUser;Success;"+$R | Out-File $oCsvFile -Append
                }
                $done=$true;
            }
            else{
                if(-not($TeamsToBeRepeated.ContainsKey($tnn))){
                    $TeamsToBeRepeated.Add($tnn,$tdn)
                }
                $uUPN+";"+"Get-AzureADUser;Error;UsersLoop - User not found" | Out-File $oCsvFile -Append
                Write-Host "  User to be set as " $R " not found in Azure AD: " $uUPN -ForegroundColor Yellow
                "  User to be set as $R not found in Azure AD: $uUPN" | Out-File $oLogFile -Append                        
            }
	    }
        catch{
            if(-not($TeamsToBeRepeated.ContainsKey($tnn))){
                $TeamsToBeRepeated.Add($tnn,$tdn)
            }
            $tnn+";"+"UsersLoop;Error;" + $_.Exception.Message | Out-File $oCsvFile -Append
            $errCount++;
            Write-Host "  ERRORE (" $errCount ") - Cannot set the user as $R of the Team $uUPN" -ForegroundColor Red
            "  ERRORE ($errCount) - Cannot set the user as $R of the Team: $uUPN" | Out-File $oLogFile -Append
        }

    }

    try{       
	    $oldgdn = $group.DisplayName
        Write-Host "  Changing the Display Name of the Team from '" $oldgdn "' to '" $tdn "'"
        "  Changing the Display Name of the Team from '$oldgdn' to '$tdn'"| Out-File $oLogFile -Append 
        
        if(-not($testOnly)){
            #CHANGE - Changing the Display Name of the Team 
            Set-Team -GroupId $group.GroupId -DisplayName $tdn -Description "Lezioni di $tdn" | Out-Null
            $oldgdn+";"+"Set-Team;Success;Change DisplayName" | Out-File $oCsvFile -Append
        }
    }
    catch{
        $uObj.UserPrincipalName+";"+"Set-Team;Error;" + $_.Exception.Message | Out-File $oCsvFile -Append
        if(-not($TeamsToBeRepeated.ContainsKey($tnn))){
            $TeamsToBeRepeated.Add($tnn,$tdn)
        }
        $errCount++;
        Write-Host "ERROR (" $errCount ") - Error while changing the Display Name of the Team from '" $oldgdn "' to '" $tdn "'" " - " $_.Exception.Message -ForegroundColor DarkRed -BackgroundColor Yellow
        "  ERROR ($errCount) - Error while changing the Display Name of the Team $tnn - $_.Exception.Message"| Out-File $oLogFile -Append 
    }
    
}

if($errCount -gt 0){
	Write-Host "ATTENTION: registered " $errCount " errors. Details in the log file '" $oLogFile "'" -ForegroundColor DarkRed -BackgroundColor Yellow
	"ATTENTION: registered $errCount errors." | Out-File $oLogFile -Append
}
else{
	Write-Host "INFORMATION: execution ended successfully. Details in the log file '" $oLogFile "'" -ForegroundColor DarkGreen -BackgroundColor Cyan
	"INFORMATION: execution ended successfully." | Out-File $oLogFile -Append
}

if( ($TeamsToBeRepeated) -and ($TeamsToBeRepeated.Count -gt 0) ){
    $oRepFile = "$outLogDir\Team-Repeat.csv"
    '"Team NN"+$delimiter+"Team DN"' | Out-File $oRepFile

    $TeamsToBeRepeated.Keys | ForEach-Object{
        '"'+ $_ + '"'+$delimiter+'"' + $TeamsToBeRepeated.Item($_) +'"' | Out-File $oRepFile -Append
    }
}

$LogEndTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
if($testOnly){
        "SIMULATED EXECUTION ENDED - $LogEndTime"| Out-File $oLogFile -Append 
}
else{
	"EXECUTION ENDED - $LogEndTime" | Out-File $oLogFile -Append
}

