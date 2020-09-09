<#
  .VERSION AND AUTHOR
    Script version: v-2020.08.02
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

$replaceOldPrefixYearInTeamsNickNameAndEmail = "cls" #Leave empty if there is no old prefix to be replaced in the Teams' MailNickName and PrimarySmtpAddress

$newPrefixYearInTeamsNickNameAndEmail = "as1920-" #Leave empty if no prefix must be added to the Teams' MailNickName and PrimarySmtpAddress
$newSuffixYearInTeamsDisplayNameAndDescription = " - A.S. 2019-20"  #Leave empty if no suffix must be appended to the Teams' DisplayName and Description

$filterTeamsDisplayNamePrefix = "5" #Leave empty if you don't need to filter-out teams based by the DisplayName prefix
$filterTeamsMatchOwner = "ownerName@schoolname.edu" #Leave empty if you don't need to filter-out teams based by the presence of a specific owner

$testOnly = $true #Set the desired value

#########################################################################################################################
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

Get-AzureADGroup -All $true | ForEach-Object{
    $t = $null
    $g = $_

    #Check if it is a Team or a different type of group (security, distribution, office 365)
    $isTeam = $true
    try{
        $t = Get-Team -GroupId $g.ObjectId 
        $line = "The group is a Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
    }
    catch{
        $isTeam = $false
        $line = "The group is not a Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
    }
    
    if($isTeam){
        $skipThisTeam = $false

        #Check if archived
        if($t.Archived){
            if( ($g.MailNickName -ne "cls5dl-scienzemotorie") -and ($g.MailNickName -ne "cls5cl-religione")){
                $skipThisTeam = $true
                $line = "   SKIPPING Team - Reason: already archived. Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
            }
        }

        #If requested, check if the display name has the specified prefix
        if((-not($skipThisTeam)) -and (-not([String]::IsNullOrEmpty($filterTeamsDisplayNamePrefix)))){
            if(-not($t.DisplayName.ToLower().StartsWith($filterTeamsDisplayNamePrefix.ToLower()))){
                $skipThisTeam = $true
                $line = "   SKIPPING Team - Reason: DisplayName not matching prefix '$filterTeamsDisplayNamePrefix'. Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
            }
        }

        #If requested, check if the Team has the specified owner
        if((-not($skipThisTeam)) -and (-not([String]::IsNullOrEmpty($filterTeamsMatchOwner)))){
            $ow = Get-AzureADGroupOwner -ObjectId $g.ObjectId 
        
            if($ow.Count -gt 0){
                if($ow.Count -eq 1){
                    $line = "   Owner: '" + $ow.UserPrincipalName + "'. Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor White
                    if($ow.UserPrincipalName.ToLower() -ne $filterTeamsMatchOwner.ToLower()){
                        $skipThisTeam = $true
                    }
                }
                else{
                    $foundOw = $false

                    :innerloop

                    ForEach($o1 in $ow){
                        $line = "   Owner: '" + $o1.UserPrincipalName + "'. Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor White
                        if($o1.UserPrincipalName.ToLower() -eq $filterTeamsMatchOwner.ToLower()){
                            $foundOw = $true  
                            break :innerloop                          
                        }
                    }

                    

                    $skipThisTeam = -not($foundOw)
                }

                if($skipThisTeam){
                    $line = "   SKIPPING Team - Reason: could not find owner '$filterTeamsMatchOwner'. Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
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
                            Set-Team -GroupId $g.ObjectId -MailNickName $newMailNickName | Out-Null
                            $line = "   newMailNickName: '$newMailNickName'. Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                        }
                        else{
                            $line = "   [SIMULATED] newMailNickName: '$newMailNickName'. Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                        }
                    }
                    catch{
                        $line = "   Could not change the MailNickName to '$newMailNickName' for the Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                        $line = "   ---> " + $_.Exception.Message; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                    }
                }

                #Change the PrimarySmtpAddress
                $ug=$null
                try{
                    $ug = Get-UnifiedGroup -Identity $g.ObjectId
                }
                catch{
                    $line = "   Could not find the Universal Groups associated to the Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
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
                            Set-UnifiedGroup –Identity $g.ObjectId –PrimarySmtpAddress $newMailPrimarySmtpAddress | Out-Null
                            $line = "   newMailPrimarySmtpAddress: '$newMailPrimarySmtpAddress'. Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                        }
                        else{
                            $line = "   [SIMULATED] newMailPrimarySmtpAddress: '$newMailPrimarySmtpAddress'. Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                        }
                    }
                    catch{
                        $line = "   Could not change the PrimarySmtpAddress to '$newMailNickName' for the Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
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
                            Set-Team -GroupId $g.ObjectId -DisplayName $newDisplayName | Out-Null
                            $line = "   newDisplayName: '$newDisplayName'. Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                        }
                        else{
                            $line = "   [SIMULATED] newDisplayName: '$newDisplayName'. Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                        }
                    }
                    catch{
                        $line = "   Could not change the DisplayName to '$newDisplayName' for the Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                        $line = "   ---> " + $_.Exception.Message; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                    }
                }


                #Change the Description (same suffix used for DisplayName)
                if(-not($t.Description.ToLower().EndsWith($newSuffixYearInTeamsDisplayNameAndDescription.ToLower()))){
                    $newDescriptionName = $t.Description + $newSuffixYearInTeamsDisplayNameAndDescription

                    try{
                        if(-not($testOnly)){
                            Set-Team -GroupId $g.ObjectId -Description $newDescriptionName | Out-Null
                            $line = "   newDescriptionName: '$newDescriptionName'. Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                        }
                        else{
                            $line = "   [SIMULATED] newDescriptionName: '$newDescriptionName'. Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                        }
                    }
                    catch{
                        $line = "   Could not change the Description to '$newDescriptionName' for the Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                        $line = "   ---> " + $_.Exception.Message; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                    }
                }
            }

            #Archive the Team
            try{
                if(-not($testOnly)){                
                    Set-TeamArchivedState -GroupId $g.ObjectId -Archived $true | Out-Null
                    $line = "   Archived! Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                }
                else{
                    $line = "   [SIMULATED] Archived! Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                }
            }
            catch{
                $line = "   Could not archive the Team: " + $g.MailNickName + " - (" + $g.DisplayName + ")"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
                $line = "   ---> " + $_.Exception.Message; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Red
            }
        }        
    }
    
}
$line = "Execution completed."; Write-Host $line -ForegroundColor White; $line | Out-File $outLogFilePath -Append
$line = "Output log file: " + $outLogFilePath; Write-Host $line -ForegroundColor White; 