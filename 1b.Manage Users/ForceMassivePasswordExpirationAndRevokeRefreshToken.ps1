<#
  .VERSION AND AUTHOR
    Script version: v-2020.08.02
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
    This script force the prompt of the password for all users except guests and administrators.

  .PREREQUISITES
  * Use Windows 10 (For earlier versions of Windows, please refer to https://docs.microsoft.com/en-us/microsoft-365/enterprise/connect-to-all-microsoft-365-services-in-a-single-windows-powershell-window?view=o365-worldwide#before-you-begin)
  * If not alreayd done, install the PowerShell modules Azure AD (or AzureADPreview) and MSOnline

    To install the PowerShell module, open PowerShell by using the option "Run as administrator" and type:
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
        Install-Module -Name AzureAD
            or
        Install-Module -Name AzureADPreview
            and
        Install-Module -Name MSOnline
    
    Details here: https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0
    and here: https://www.powershellgallery.com/packages/MSOnline/

  .ATTENTION.
  To change their password (as forced by this script), users must remember the current password! 
  If you belive that many of your users do not remember their current password and if you haven't yet enabled the functionality
  of Self Service Password Resent, be prepared to manage a huge number of requests of reset lost password!

  .VARIABLES TO BE SET
  Read below.
#>

#########################################################################################################################
# VARIABLES TO BE SET:
#########################################################################################################################

$outDirPath = "c:\Tem\OUTs" #Set the correct path!
$adminName = "adminName@schoolname.edu" #Set the correct name!
$testOnly = $true #Set the desired value

######################################################################
$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
$outLogFilePath = "$outLogDir\ForceMassivePasswordExpirationAndRevokeRefreshToken_$LogStartTime.log"
If (Test-Path $outLogFilePath)
{
	Remove-Item $outLogFilePath
}

cls

Connect-MsolService
Connect-AzureAD -AccountId $adminName


$line = "START - "+(Get-Date); $line| Out-File $outLogFilePath; Write-Host $line 
$count = 0;
Write-Host "Execution Started - Please wait..." -ForegroundColor Yellow

$allMsolUsers = Get-MsolUser -All 
$uCount = $allMsolUsers.Count
$iCount = 0
$allMsolUsers| ForEach-Object {
    $msolu = $($_)
    $iCount++;
    $upn = $msolu.UserPrincipalName
    $line = "("+$icount+"/"+$uCount+") "+$upn; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
    if($upn.IndexOf("#EXT#@") -gt 0){
        $line = "  User SKIPPED - Guest!"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
    }
    else{
        $aadu = Get-AzureADUser -Filter "UserPrincipalName eq '$upn'"
        $aad_urg = Get-AzureADUserMembership -ObjectId $aadu.ObjectId | Group-Object -Property ObjectType -AsHashTable 
        $Roles = $aad_urg.Role 
        if($Roles){
            $line = "  User SKIPPED - Roles:"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
            $Roles | ForEach-Object {
                $r = Get-AzureADDirectoryRole -ObjectId $_.ObjectId 
                $line = "    "+$r.DisplayName; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Cyan
            }
        } 
        else {
            if(-not($testOnly)){
                $line = "  Forcing password expiration"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
                #CHANGE!
                Set-MsolUserPassword -UserPrincipalName $upn -ForceChangePasswordOnly $true -ForceChangePassword $true
            
                $line = "  Forcing refresh token revokation"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Green
                #CHANGE!
                Revoke-AzureADUserAllRefreshToken -ObjectId $aadu.ObjectId
            }
            else{
                $line = "  [Simulation] Forcing password expiration"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
                $line = "  [Simulation] Forcing refresh token revokation"; $line| Out-File $outLogFilePath -Append; Write-Host $line -ForegroundColor Yellow
            }
        }
    }
}

$line = "END - "+(Get-Date); $line| Out-File $outLogFilePath -Append; Write-Host $line 
Write-Host "Log file: '" $outLogFilePath "'" -ForegroundColor Yellow
