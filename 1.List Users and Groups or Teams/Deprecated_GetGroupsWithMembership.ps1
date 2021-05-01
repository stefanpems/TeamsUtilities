<#
  The use of this script is now DEPRECATED!
  Use GetUsersAndBelongingGroups.ps1 instead.
  THIS SCRIPT IS STILL USEFUL ONLY TO GET THE LIST OF THE EMPTY GROUPS

  .VERSION AND AUTHOR
    Script version: v-2021.05.01
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
  This script generates a CSV files with the list of all the Groups (Security, Distribution, Office 365 and Teams) defined 
  in Office 365/Azure Active Directory and, for each group, with the list of users in the group listed as members or owners.

  .VARIABLES TO BE SET
    $outCsvDirPath: path of the local folder where the script generates the output csv file.
    $adminName: name of the administrative account to be used for the script execution (the password will be prompted).

  .PREREQUISITES
  * Use Windows 10 (For earlier versions of Windows, please refer to https://docs.microsoft.com/en-us/microsoft-365/enterprise/connect-to-all-microsoft-365-services-in-a-single-windows-powershell-window?view=o365-worldwide#before-you-begin)
  * If not alreayd done, install the PowerShell module Azure AD (or AzureADPreview) 

    To install the PowerShell module, open PowerShell by using the option "Run as administrator" and type:
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
        Install-Module -Name AzureAD
            or
        Install-Module -Name AzureADPreview

    Details here: https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0

  .NOTES
  1) In the script results, the onwers are not repeated as members (there are type of groups, like Office 365 Groups, where
     the same user listed in the set of owners is also listed in the set of memebers; this is not the case for Teams).
  2) The script needs to be modified to consider also nested groups (for security and distribution groups).
     Nested groups are currently listed as empty UserPrincipalName.
#>

#########################################################################################################################

$outCsvDirPath = "C:\Temp" #This is just a sample
$adminName = "nomeutente@nomescuola.edu.it" #This is just a sample

#########################################################################################################################
$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"

$outUsersCsvFilePath = "$outCsvDirPath\DumpTeams_Results_$LogStartTime.csv"
$outOrphanGroupsCsvFilePath = "$outCsvDirPath\DumpOrphanTeams_Results_$LogStartTime.csv"
If (Test-Path $outUsersCsvFilePath){
	Remove-Item $outUsersCsvFilePath
}
'"GroupNickName","GroupDisplayName","UserPrincipalName","UserType"' | Out-File $outUsersCsvFilePath
'"GroupNickName","GroupDisplayName","ObjectType","DirSyncEnabled","MailEnabled","SecurityEnabled"' | Out-File $outOrphanGroupsCsvFilePath

Connect-AzureAD -AccountId $adminName 

Get-AzureADGroup -All $true | ForEach-Object{
    $g = $_
    
    $orphan = $true

    $ow = Get-AzureADGroupOwner  -ObjectId $g.ObjectId 

    $a = New-Object -TypeName "System.Collections.ArrayList"

    if($ow.Count -gt 0){
        $orphan = $false
        if($ow.Count -eq 1){
            $a.Add($ow.UserPrincipalName)|Out-Null
            Write-Host $g.MailNickName " - " $g.DisplayName " - " $ow.UserPrincipalName " - Owner" -ForegroundColor Magenta
            '"'+$g.MailNickName+'","'+$g.DisplayName+'","'+$ow.UserPrincipalName+'","Owner"' | Out-File $outUsersCsvFilePath -Append
        }
        else{
            $ow | ForEach-Object{
                $a.Add($_.UserPrincipalName)|Out-Null
                Write-Host $g.MailNickName " - " $g.DisplayName " - " $_.UserPrincipalName " - Owner" -ForegroundColor Magenta
                '"'+$g.MailNickName+'","'+$g.DisplayName+'","'+$_.UserPrincipalName+'","Owner"' | Out-File $outUsersCsvFilePath -Append
            }
        }
    }
    

    $og = Get-AzureADGroupMember -ObjectId $g.ObjectId


    if($og.Count -gt 0){
        $orphan = $false
    }

    
    $og | ForEach-Object{
        if(-not($a.Contains($_.UserPrincipalName))){
            Write-Host $g.MailNickName " - " $g.DisplayName " - " $_.UserPrincipalName " - Member" -ForegroundColor Green
            '"'+$g.MailNickName+'","'+$g.DisplayName+'","'+$_.UserPrincipalName+'","Member"' | Out-File $outUsersCsvFilePath -Append
        }
        else{
            Write-Host "Skipped " $g.MailNickName " - " $g.DisplayName " - " $_.UserPrincipalName " - Member (Already Owner)" -ForegroundColor Gray
        }
    }

    if($orphan){
        Write-Host "Empty group: " $g.MailNickName "," $g.DisplayName "," $g.ObjectType "," $g.DirSyncEnabled "," $g.MailEnabled "," $g.SecurityEnabled -ForegroundColor Cyan 
        '"'+$g.MailNickName+'","'+$g.DisplayName+'","'+$g.ObjectType+'","'+$g.DirSyncEnabled+'","'+$g.MailEnabled+'","'+$g.SecurityEnabled+'"'  | Out-File $outOrphanGroupsCsvFilePath -Append
    }


}

Write-Host "Generato i file " $outUsersCsvFilePath " e " $outOrphanGroupsCsvFilePath