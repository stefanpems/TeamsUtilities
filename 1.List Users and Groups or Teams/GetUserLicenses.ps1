<#
  .VERSION AND AUTHOR 
    Script version: v-2021.01.13
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .CREDITS
  Partially reused the script created by Jean-Marie AGBO (https://agbo.blog/2019/08/22/how-is-your-office-365-users-licenses-assigned-direct-or-inherited/)

  .SYNOPSIS
  This script generates a CSV files with the list of all the users existing in Azure AD/Office 365
  and with the evidences of their assigned licenses.
  For each license assignment, the script clarifies if it is "Direct" or "Inherited" (through assignment to a group to whom the user belongs) 

  .VARIABLES TO BE SET
    See below

  .NOTES
  Please modify the Get-LicensePlan() function to add your preferred renaming of the service plans.
  The complete list is here: https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference

  .PREREQUISITES
  * Use Windows 10 (For earlier versions of Windows, please refer to https://docs.microsoft.com/en-us/microsoft-365/enterprise/connect-to-all-microsoft-365-services-in-a-single-windows-powershell-window?view=o365-worldwide#before-you-begin)
  * If not alreayd done, install the PowerShell module MSOnline

    To install the PowerShell module, open PowerShell by using the option "Run as administrator" and type:
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
        Install-Module -Name MSOnline
        
    Details here: https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0
    and here: https://www.powershellgallery.com/packages/MSOnline/

#>
############################################################################################
#                                   VARIABLES TO BE SET
############################################################################################

$ofile = "C:\ZTemp\licensesMLPS.csv" #Path of the CSV file to be generated as output (if existing, it will be overwritten)


############################################################################################
#                                       FUNCTIONS
############################################################################################
function Get-LicensePlan {

    param (

        [Parameter(Mandatory=$true)]
        [String]$SkuId,
        [Parameter(mandatory=$true)]
        [String]$TenantName

    )


    #Please modify the following lines to add your preferred renaming of the service plans.
    #The complete list is here: https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference

    Switch($SkuId){
                      "$($TenantName):OFFICESUBSCRIPTION" {return "O365 ProPlus"}
                      "$($TenantName):AAD_PREMIUM" {return "AAD Premium P1"}
                              "$($TenantName):EMS" {return "EMS E3"}
                       "$($TenantName):EMSPREMIUM" {return "EMS E5"}
                     "$($TenantName):STANDARDPACK" {return "O365 E1"}
                   "$($TenantName):ENTERPRISEPACK" {return "O365 E3"}
                        "$($TenantName):FLOW_FREE" {return "FLOW FREE"}
                      "$($TenantName):INTUNE_A_VL" {return "INTUNE"}
                       "$($TenantName):MCOMEETADV" {return "SFB PSTN Conf"}
        "$($TenantName):MICROSOFT_BUSINESS_CENTER" {return "MBC"}
                     "$($TenantName):POWER_BI_PRO" {return "PBI PRO"}
                "$($TenantName):POWER_BI_STANDARD" {return "PBI STD"}
        "$($TenantName):POWERAPPS_INDIVIDUAL_USER" {return "PAPPS IND User"}
                  "$($TenantName):POWERAPPS_VIRAL" {return "PAPPS and LOGIC FLOW"}
                   "$($TenantName):PROJECTPREMIUM" {return "PJ Online"}
                           "$($TenantName):STREAM" {return "STREAM"}
                "$($TenantName):VISIOONLINE_PLAN1" {return "VISIO P1"}
              "$($TenantName):WACONEDRIVESTANDARD" {return "ODfB P1"}
                      "$($TenantName):WIN_DEF_ATP" {return "MD ATP"}
                      "$($TenantName):ATA" {return "MDI"}
                                           default {return $SkuId.Replace("$($TenantName):","")}
    }

}

function Get-Licenses{

    Begin{
        Write-Host "## Data processing stated at $(Get-date)." -ForegroundColor Yellow
        Write-Host ""
        $TenantName = ((Get-MsolAccountSku).AccountSkuId[0] -split(':'))[0]
    }

    Process{
        Write-Host " "
        Write-Host " "
        Write-Host " "
        Write-Host " "
        Write-Host "Retrieving the list of all the users in the tenant. It may take a few minutes. Please wait..." -ForegroundColor White
        "UPN;LicensePlan;AssignmentType;GroupName" | Out-File $ofile
        $allUsers = Get-MsolUser -All
        $numUsers = $allUsers.Count
        $totCount = 0

        Write-Host " "
        Write-Host "Reading the license assignment for each user..." -ForegroundColor White
        $allUsers | ForEach-Object{
            $totCount++;        
            $UPN = $_.UserPrincipalName

            try{
                $User = Get-MsolUser -UserPrincipalName $UPN

                #Getting assignment paths
                $LicensesTab = $null
                $LicensePlan = $null
                $LicTabCount = 0
                $LicensesTab = $User.Licenses | Select-Object AccountSkuId, GroupsAssigningLicense

                if($LicensesTab){

                    $i = 0 
                    $LicTabCount = $LicensesTab.AccountSkuId.Count

                    Do{

                        $LicensePlan = Get-LicensePlan -SkuId $LicensesTab[$i].AccountSkuId -TenantName $TenantName

                        $numLic = 0
                        if($LicensesTab[$i].GroupsAssigningLicense){
                            $numLic = $LicensesTab[$i].GroupsAssigningLicense.guid.Count
                            foreach ($Guid in $LicensesTab[$i].GroupsAssigningLicense.guid){

                                if($Guid -eq $User.ObjectId.Guid){
                                    
                                    if($numLic -gt 1){
                                        $UPN+";"+$LicensePlan+";AlsoDirect;" | Out-File $ofile -Append
                                    }
                                    else{
                                        $UPN+";"+$LicensePlan+";OnlyDirect;" | Out-File $ofile -Append
                                    }
                                }
                                else{                                    
                                    $UPN+";"+$LicensePlan+";Group;"+(Get-MsolGroup -ObjectId $Guid).DisplayName | Out-File $ofile -Append
                                }

                            }
                        }
                        else{
                            $UPN+";"+$LicensePlan+";OnlyDirect;" | Out-File $ofile -Append
                        }

                        $i++

                    }
                    While ($i -ne $LicTabCount)
                }
                else {
                    $UPN+";;;" | Out-File $ofile -Append
                }
        
                #Manage progress bar
                $percentComplete = [math]::Round(100*$totCount/$numUsers,1)
                Write-Progress -Activity "Scanning users" -Status "$percentComplete % users analyzed" -PercentComplete $percentComplete
            }
            catch{
                $UPN+";Error;Error;" | Out-File $ofile -Append
            }
        }

        Write-Host " "
        Write-Host "Scanned all the users..." -ForegroundColor White
    }

    End{
        Write-Host ""
        Write-Host "## Data Processing ended on $(Get-Date)" -ForegroundColor Yellow
    }

}

############################################################################################
#                                       MAIN
############################################################################################
cls
$fileexists = $false

try{
    Get-Date | Out-File $ofile
    $fileexists = $true
}
catch{
    Write-Host ""
    Write-Host "Please specify a correct path for the output file - Current path: $ofile" -ForegroundColor Red
}

if($fileexists){
    Connect-MsolService
    Get-Licenses
} 
