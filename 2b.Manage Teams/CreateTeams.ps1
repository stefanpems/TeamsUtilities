<#
  .VERSION AND AUTHOR
    Script version: v-2020.09.14
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
    This script massively creates Teams and assign to them members and owners as specified in the 3 CSV files specified in input.

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

$csvTeams = "C:\Temp\Team.csv"           #EXPECTED COLUMNS: "Team NN","Team DN" 
$csvOwners = "C:\Temp\Team-Owner.csv"    #EXPECTED COLUMNS: "Team NN","Owner" 
$csvUsers = "C:\Temp\Team-User.csv"      #EXPECTED COLUMNS: "Team NN","User" 
$delimiter = ";" #Set the correct value of the delimeter in all the 3 CSV files
$outLogDir = "C:\Temp\OUT" #Set the correct path!
$adminName = "admin@schoolname.edu" #Set the correct name! 
    #NOTES for this account: 
    #1) This account must already have assigned a license including Teams
    #2) RECOMMENDATION: do not use a "personal" account as $adminName. This name appears in the home screen in the general channel
    #   of every created created Team.
    #3) ATTENTION: this script does not remove this account from the list of the owners of each created team.
    #   It's important to consider that a single account can be member of at most 1000 teams. 
    #   Please consider this limit. If the account $adminName is already in many other teams or if you are adding more than 1000 teams
    #   you can evaluate the option to add in the code below the removal of the account from each created team or you can 
    #   use a new administrative account or split the execution with different administrative accounts. 
$testOnly = $false #Set the desired value


function CreateEduTeam{

    param(
         [string]
         $TeamName, 
         [string]
         $TeamDescription,
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
            "template@odata.bind" = "https://graph.microsoft.com/beta/teamsTemplates('educationClass')"
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
        "  Error while creating team '$d.TeamName': '$_'" | Out-File $oFile -Append
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
	$oFile = "$outLogDir\SIMULATION-CreateTeams_$LogStartTime.log"
}
else{
	$oFile = "$outLogDir\CreateTeams_$LogStartTime.log"
}
If (Test-Path $oFile)
{
	Remove-Item $oFile
}
if($testOnly){
	"SIMULATED EXECUTION STARTED - $LogStartTime" | Out-File $oFile 
}
else{
	"EXECUTION STARTED - $LogStartTime" | Out-File $oFile 
}
Write-Host "Creato il file di log '$oFile'" -ForegroundColor Yellow

$errCount = 0;


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
    $tnn = $teamRow."Team NN"
    $tdn = $teamRow."Team DN"
    
    Write-Host "Searching for the team: " $tnn "-" $tdn
    "Searching for the team: $tnn - $tdn" | Out-File $oFile -Append

    try{

        $group = Get-Team -MailNickname $tnn

        if($group -eq $null){

            $tokenExpiration = $accessToken.ValidTo.ToLocalTime().AddMinutes(-1);
            $TimeToExpiry = $tokenExpiration - (Get-Date)        
            $sTimeToExpiry = $TimeToExpiry.Minutes.ToString() + " min " + $TimeToExpiry.Seconds.ToString() + " sec"
                
            Write-Host "  Token - TimeToExpiry: " $sTimeToExpiry
            "  Token - TimeToExpiry: '$sTimeToExpiry'" | Out-File $oFile -Append            

            $IsExpired = (Get-Date) -gt $tokenExpiration   

            if($IsExpired){
                Write-Host "  Token expired! Acquiring a new token" -ForegroundColor Cyan 
                "  Token expired! Acquiring a new token." | Out-File $oFile -Append     
            
                $accessToken = $null;
            
                try{
                    Connect-PnPOnline -Scopes $arrayOfScopes 
                    $accessToken = Get-PnPAccessToken -Decoded      
                }
                catch{
                    Write-Host "  Could not acquire a new token. Forced exit!" -ForegroundColor Red 
                    "  Could not acquire a new token. Forced exit!" | Out-File $oFile -Append     
                    Break TeamsLoop
                }

                $tokenExpiration = $accessToken.ValidTo.ToLocalTime();
                $TimeToExpiry = $tokenExpiration - (Get-Date)        
                $sTimeToExpiry = $TimeToExpiry.Minutes.ToString() + " min " + $TimeToExpiry.Seconds.ToString() + " sec"
                
                Write-Host "  New Token - TimeToExpiry: " $sTimeToExpiry
                "  New Token - TimeToExpiry: '$sTimeToExpiry'" | Out-File $oFile -Append            

            }
     
            if(-not($testOnly)){
                #CHANGE - Provisioning of the new Team by using GraphAPI
                try{
                    CreateEduTeam -TeamName $tnn -TeamDescription $tdn -accessToken $accessToken.RawData
                }
                catch{
                    $errCount++;
                    Write-Host "ERROR while creating the team '" $tnn "': " $_.Exception.Message "'" -ForegroundColor Red 
                    "  ERROR while creating the team: '$tnn'" | Out-File $oFile -Append

                    throw
                }

                #Accessing the new team 
                $group = Get-Team -MailNickname $tnn
                if($group){
                    $gnn = $group.MailNickName
                    Write-Host "  New team created and successfully accessed: " $gnn
                    "  New team created and successfully accessed: '$gnn'" | Out-File $oFile -Append
                }
                else{
                    $errCount++;
                    Write-Host "ERROR while searching the new team: " $tnn -ForegroundColor Red 
                    "  ERROR while searching the new team: '$tnn'" | Out-File $oFile -Append
                
                    throw
                }
            }
            else{
                Write-Host "  Simulated - Created the new team: " $tnn
                "  Simulated - Created the new team: $tnn" | Out-File $oFile -Append
            }
        }
        else{
            Write-Host "  The team already exists: " $group.MailNickName
            "  The team already exists: $tnn" | Out-File $oFile -Append
        }
    
    }
    catch{
        Write-Host "Skipping the team '" $tnn "' " -ForegroundColor Red 
        "  Skipping the team: '$tnn'" | Out-File $oFile -Append

        continue TeamsLoop #STOP executing the following actions for this team!
    }

    #Reading the Team owners
    $ownersRows = Import-Csv -Path $csvOwners  -Delimiter $delimiter | 
    Where-Object -Property "Team NN" -eq $tnn 

    :OwnersLoop
    Foreach($ownerRow in $ownersRows){            
        $oUPN = $ownerRow."Owner"
                
        try{
            $oObj = Get-AzureADUser -Filter "UserPrincipalName eq '$oUPN'"

            if($oObj){
                Write-Host "  Setting the user as team owner: " $oUPN -ForegroundColor Green
                "  Setting the user as team owner: $oUPN" | Out-File $oFile -Append
                        
                if(-not($testOnly)){
                    #CHANGE - Setting the user as team owner
                    Add-TeamUser -GroupId $group.GroupId -User $oUPN -Role Owner 
                }
                $done=$true;
            }
            else{
                Write-Host "  User to be set as owner not found in Azure AD: " $oUPN -ForegroundColor Yellow
                "  User to be set as owner not found in Azure AD: $oUPN" | Out-File $oFile -Append                        
            }
	    }
        catch{
            $errCount++;
            Write-Host "  ERROR (" $errCount ") - Cannot set the user as owner of the Team $oUPN" -ForegroundColor Red
            "  ERROR ($errCount) - Cannot set the user as owner of the Team: $oUPN" | Out-File $oFile -Append               
        }

    }

    #Reading the Team Members
    $membersRows = Import-Csv -Path $csvUsers  -Delimiter $delimiter | 
    Where-Object -Property "Team NN" -eq $tnn # $($_."Team NN")

    :MembersLoop
    Foreach($memberRow in $membersRows){           
        $mUPN = $memberRow."User"
                
        try{
            $mObj = Get-AzureADUser -Filter "UserPrincipalName eq '$mUPN'"

            if($mObj){
                Write-Host "  Setting the user as team member: " $mUPN -ForegroundColor Green
                "  Setting the user as team member: $mUPN" | Out-File $oFile -Append

                if(-not($testOnly)){
                    #CHANGE - Setting the user as team member
                    Add-TeamUser -GroupId $group.GroupId -User $mUPN -Role Member
                }
                $done=$true;
            }
            else{
                Write-Host "  User to be set as member not found in Azure AD: " $mUPN -ForegroundColor Yellow
                "  User to be set as member not found in Azure AD: $mUPN" | Out-File $oFile -Append                        
            }
	    }
        catch{
            $errCount++;
            Write-Host "  ERRORE (" $errCount ") - Cannot set the user as member of the Team $mUPN" -ForegroundColor Red
            "  ERRORE ($errCount) - Cannot set the user as member of the Team: $mUPN" | Out-File $oFile -Append
        }

    }

    try{       
	    $oldgdn = $group.DisplayName
        Write-Host "  Changing the Display Name of the Team from '" $oldgdn "' to '" $tdn "'"
        "  Changing the Display Name of the Team from '$oldgdn' to '$tdn'"| Out-File $oFile -Append 
        
        if(-not($testOnly)){
            #CHANGE - Changing the Display Name of the Team 
            Set-Team -GroupId $group.GroupId -DisplayName $tdn -Description "Lezioni di $tdn" | Out-Null
        }
    }
    catch{
        $errCount++;
        Write-Host "ERROR (" $errCount ") - Error while changing the Display Name of the Team from '" $oldgdn "' to '" $tdn "'" " - " $_.Exception.Message -ForegroundColor DarkRed -BackgroundColor Yellow
        "  ERROR ($errCount) - Error while changing the Display Name of the Team $tnn - $_.Exception.Message"| Out-File $oFile -Append 
    }
    
}

if($errCount -gt 0){
	Write-Host "ATTENTION: registered " $errCount " errors. Details in the log file '" $oFile "'" -ForegroundColor DarkRed -BackgroundColor Yellow
	"ATTENTION: registered $errCount errors." | Out-File $oFile -Append
}
else{
	Write-Host "INFORMATION: execution ended successfully. Details in the log file '" $oFile "'" -ForegroundColor DarkGreen -BackgroundColor Cyan
	"INFORMATION: execution ended successfully." | Out-File $oFile -Append
}

$LogEndTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
if($testOnly){
        "SIMULATED EXECUTION ENDED - $LogEndTime"| Out-File $oFile -Append 
}
else{
	"EXECUTION ENDED - $LogEndTime" | Out-File $oFile -Append
}

