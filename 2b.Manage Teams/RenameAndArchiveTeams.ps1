<#
  .VERSION AND AUTHOR
    Script version: v-2020.10.04
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
    This script renames and archive the Teams identified by their name prefix and/or owner

  .PREREQUISITES
    PowerShell modules Azure AD (or AzureADPreview), MicrosoftTeams and ExchangeOnline
    To install the PowerShell module, open PowerShell as administrator and type:
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
        Install-Module -Name <module-name>
    Details here: https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-all-office-365-services-in-a-single-windows-powershell-window

  .VARIABLES TO BE SET
  Read below.
#>

#########################################################################################################################
# VARIABLES TO BE SET:
#########################################################################################################################

$outDirPath = "c:\Tem\OUTs" #Set the correct path!
$adminName = "adminName@schoolname.edu" #Set the correct name!

$csvTeamsToBeConsidered = "" #Leave empty if the teams should be found directly in the online directory.
$columnTeamNN = "Team NN" #Used only of $csvTeamsToBeConsidered is not empty. Specify the name of the column (in the CSV file) containing the NickName of the Teams 
$delimiter = ";" #Used only of $csvTeamsToBeConsidered is not empty. Specify the separator in the CSV file

$filterTeamsNickNamePrefix = "cls" #Leave empty if you don't need to filter-out teams based by the NickName prefix
$filterTeamsDisplayNamePrefix = "5" #Leave empty if you don't need to filter-out teams based by the DisplayName prefix
$filterTeamsMatchOwner = "ownerName@schoolname.edu" #Leave empty if you don't need to filter-out teams based by the presence of a specific owner

$replaceOldPrefixYearInTeamsNickNameAndEmail = "cls" #Leave empty if there is no old prefix to be replaced in the Teams' MailNickName and PrimarySmtpAddress
$newPrefixYearInTeamsNickNameAndEmail = "as1920-" #Leave empty if no prefix must be added to the Teams' MailNickName and PrimarySmtpAddress
$newSuffixYearInTeamsDisplayNameAndDescription = " - A.S. 2019-20"  #Leave empty if no suffix must be appended to the Teams' DisplayName and Description

$testOnly = $true #Set the desired value

#########################################################################################################################
function ProcessTeam{

    param(
         [Object]
         $t
    )

    $skipThisTeam = $false

    #Check if archived
    if($t.Archived){
        $skipThisTeam = $true
        $line = "   SKIPPING Team - Reason: already archived. Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
    }

    #If requested, check if the display name has the specified prefix
    if((-not($skipThisTeam)) -and (-not([String]::IsNullOrEmpty($filterTeamsDisplayNamePrefix)))){
        if(-not($t.DisplayName.ToLower().StartsWith($filterTeamsDisplayNamePrefix.ToLower()))){
            $skipThisTeam = $true
            $line = "   SKIPPING Team - Reason: DisplayName not matching prefix '$filterTeamsDisplayNamePrefix'. Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
        }
    }

    if((-not($skipThisTeam)) -and (-not([String]::IsNullOrEmpty($filterTeamsNickNamePrefix)))){
        if(-not($t.MailNickName.ToLower().StartsWith($filterTeamsNickNamePrefix.ToLower()))){
            $skipThisTeam = $true
            $line = "   SKIPPING Team - Reason: NickName not matching prefix '$filterTeamsNickNamePrefix'. Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
        }
    }

    #If requested, check if the Team has the specified owner
    if((-not($skipThisTeam)) -and (-not([String]::IsNullOrEmpty($filterTeamsMatchOwner)))){
        $ow = Get-AzureADGroupOwner -ObjectId $t.GroupId 
        
        if($ow.Count -gt 0){
            if($ow.Count -eq 1){
                $line = "   Owner: '" + $ow.UserPrincipalName + "'. Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor White
                if($ow.UserPrincipalName.ToLower() -ne $filterTeamsMatchOwner.ToLower()){
                    $skipThisTeam = $true
                }
            }
            else{
                $foundOw = $false

                :innerloop
                ForEach($o1 in $ow){
                    $line = "   Owner: '" + $o1.UserPrincipalName + "'. Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor White
                    if($o1.UserPrincipalName.ToLower() -eq $filterTeamsMatchOwner.ToLower()){
                        $foundOw = $true  
                        break :innerloop                          
                    }
                }                    

                $skipThisTeam = -not($foundOw)
            }

            if($skipThisTeam){
                $line = "   SKIPPING Team - Reason: could not find owner '$filterTeamsMatchOwner'. Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
            }
        }
    }

    if(-not($skipThisTeam)){
        #If required, change the NickName and PrimarySmtpAddress
        if(-not([String]::IsNullOrEmpty($newPrefixYearInTeamsNickNameAndEmail))){

            #Change the NickName
            if(-not($t.MailNickName.ToLower().StartsWith($newPrefixYearInTeamsNickNameAndEmail.ToLower()))){
                if($t.MailNickName.StartsWith($replaceOldPrefixYearInTeamsNickNameAndEmail)){
                    $newMailNickName = $t.MailNickName.Replace($replaceOldPrefixYearInTeamsNickNameAndEmail,$newPrefixYearInTeamsNickNameAndEmail)
                }
                else{
                    $newMailNickName = $newPrefixYearInTeamsNickNameAndEmail + $t.MailNickName
                }
                $newMailNickName = $newMailNickName.ToLower()

                try{
                    if(-not($testOnly)){
                        Set-Team -GroupId $t.GroupId -MailNickName $newMailNickName | Out-Null
                        $line = "   newMailNickName: '$newMailNickName'. Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                    }
                    else{
                        $line = "   [SIMULATED] newMailNickName: '$newMailNickName'. Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                    }
                }
                catch{
                    $line = "   Could not change the MailNickName to '$newMailNickName' for the Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                    $line = "   ---> " + $_.Exception.Message; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                }
            }

            #Change the PrimarySmtpAddress
            $ug=$null
            try{
                $ug = Get-UnifiedGroup -Identity $t.GroupId
            }
            catch{
                $line = "   Could not find the Universal Groups associated to the Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                $line = "   ---> " + $_.Exception.Message; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
            }

            if( ($ug) -and (-not($ug.PrimarySmtpAddress.ToLower().StartsWith($newPrefixYearInTeamsNickNameAndEmail.ToLower())))){
                if($ug.PrimarySmtpAddress.StartsWith($replaceOldPrefixYearInTeamsNickNameAndEmail)){
                    $newMailPrimarySmtpAddress = $ug.PrimarySmtpAddress.Replace($replaceOldPrefixYearInTeamsNickNameAndEmail,$newPrefixYearInTeamsNickNameAndEmail)
                }
                else{
                    $newMailPrimarySmtpAddress = $newPrefixYearInTeamsNickNameAndEmail + $ug.PrimarySmtpAddress
                }
                $newMailPrimarySmtpAddress = $newMailPrimarySmtpAddress.ToLower()

                try{
                    if(-not($testOnly)){
                        Set-UnifiedGroup –Identity $t.GroupId –PrimarySmtpAddress $newMailPrimarySmtpAddress | Out-Null
                        $line = "   newMailPrimarySmtpAddress: '$newMailPrimarySmtpAddress'. Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                    }
                    else{
                        $line = "   [SIMULATED] newMailPrimarySmtpAddress: '$newMailPrimarySmtpAddress'. Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                    }
                }
                catch{
                    $line = "   Could not change the PrimarySmtpAddress to '$newMailNickName' for the Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                    $line = "   ---> " + $_.Exception.Message; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                }
            }
        }

        #If required, change the DisplayName and Description
        if(-not([String]::IsNullOrEmpty($newSuffixYearInTeamsDisplayNameAndDescription))){

            #Change the DisplayName
            if(-not($t.DisplayName.ToLower().EndsWith($newSuffixYearInTeamsDisplayNameAndDescription.ToLower()))){
                $newDisplayName = $t.DisplayName + $newSuffixYearInTeamsDisplayNameAndDescription

                try{
                    if(-not($testOnly)){
                        Set-Team -GroupId $t.GroupId -DisplayName $newDisplayName | Out-Null
                        $line = "   newDisplayName: '$newDisplayName'. Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                    }
                    else{
                        $line = "   [SIMULATED] newDisplayName: '$newDisplayName'. Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                    }
                }
                catch{
                    $line = "   Could not change the DisplayName to '$newDisplayName' for the Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                    $line = "   ---> " + $_.Exception.Message; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                }
            }


            #Change the Description (same suffix used for DisplayName)
            if(-not($t.Description.ToLower().EndsWith($newSuffixYearInTeamsDisplayNameAndDescription.ToLower()))){
                $newDescriptionName = $t.Description + $newSuffixYearInTeamsDisplayNameAndDescription

                try{
                    if(-not($testOnly)){
                        Set-Team -GroupId $t.GroupId -Description $newDescriptionName | Out-Null
                        $line = "   newDescriptionName: '$newDescriptionName'. Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                    }
                    else{
                        $line = "   [SIMULATED] newDescriptionName: '$newDescriptionName'. Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                    }
                }
                catch{
                    $line = "   Could not change the Description to '$newDescriptionName' for the Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                    $line = "   ---> " + $_.Exception.Message; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                }
            }
        }

        #Archive the Team
        try{
            if(-not($testOnly)){                
                Set-TeamArchivedState -GroupId $t.GroupId -Archived $true | Out-Null
                $line = "   Archived! Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
            }
            else{
                $line = "   [SIMULATED] Archived! Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
            }
        }
        catch{
            $line = "   Could not archive the Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
            $line = "   ---> " + $_.Exception.Message; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
        }
    }

}

cls
$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
$outFilePrefixName = "RenameAndArchiveTeams_"

$outLogFilePath = "$outDirPath\$outFilePrefixName$LogStartTime.log"
If (Test-Path $outLogFilePath){
	Remove-Item $outLogFilePath
}

Connect-AzureAD -AccountId $adminName 
Connect-MicrosoftTeams -AccountId $adminName
Connect-ExchangeOnline

if([String]::IsNullOrEmpty($csvTeamsToBeConsidered)){
    Write-Host "Reading from Azure Active Directory" 
    $groups = Get-AzureADGroup -All $true
    $numItems = $groups.Count; if($numItems -eq $null) {$numItems = 1}
    $currItem = 0

    :GroupsLoop
    ForEach ($g in $groups){

        #Progress bar
        $currItem++;
        $percentComplete = [math]::Round(100*$currItem/$numItems,1)
        Write-Progress -Activity "Scanning items" -Status "$percentComplete % rows processed" -PercentComplete $percentComplete

        #Check if it is a Team or a different type of group (security, distribution, office 365)
        $t = $null
        try{
            $t = Get-Team -GroupId $g.ObjectId 
            $line = "The group is a Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
            ProcessTeam($t)      
        }
        catch{
            $line = "The group is not a Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
        }    
    }
}
else{
    Write-Host "Reading from CSV file '" $csvTeamsToBeConsidered "'"
    $rows  = Import-Csv -Path $csvTeamsToBeConsidered -Delimiter $delimiter
    $numItems = $rows.Count; if($numItems -eq $null) {$numItems = 1}
    $currItem = 0

    :RowsLoop
    ForEach ($r in $rows){

        #Progress bar
        $currItem++;
        $percentComplete = [math]::Round(100*$currItem/$numItems,1)
        Write-Progress -Activity "Scanning items" -Status "$percentComplete % rows processed" -PercentComplete $percentComplete

        #Check if it is a Team or a different type of group (security, distribution, office 365)
        $t = $null
        try{
            $tnn = $r."$columnTeamNN"
            $t = Get-Team -MailNickName $tnn
            $line = "Successfully accessed the Team: " + $t.MailNickName + " - (" + $t.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
            ProcessTeam($t)      
        }
        catch{
            $line = "Could not find the Team from the row: '" + $r + "'"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
        }    
    }
}

$line = "Execution completed."; Write-Host $line -ForegroundColor White; $line | Out-File $outLogFilePath -Append
$line = "Output log file: " + $outLogFilePath; Write-Host $line -ForegroundColor White; 