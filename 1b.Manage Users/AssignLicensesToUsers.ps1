<#
  .VERSION AND AUTHOR
    Script version: v-2020.08.02
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
    This script assign Office 365 Education licenses to Users specified in the input CSV.

  .PREREQUISITES
    PowerShell module Azure AD (or AzureADPreview)
    To install the PowerShell module, open PowerShell as administrator and type:
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
        Install-Module -Name AzureAD
            or
        Install-Module -Name AzureADPreview

    Details here: https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0

  .VARIABLES TO BE SET
  Read below.
#>

#########################################################################################################################
# VARIABLES TO BE SET:
#########################################################################################################################

$outDirPath = "c:\Tem\OUTs" #Set the correct path!
$inputCsvPath = "C:\Temp\INs\Users.csv" #Set the correct path!
$inputCsvDelimiter = ";" #Set the correct value!
$colNameLoginName = "Nome utente" #Name of the column with the "login name" of the user. Do not use special characters in this column name here (only ANSI characters)
$adminName = "adminName@schoolname.edu" #Set the correct name!
$AssignedLicenseSkuId = "94763226-9b3c-4e75-a931-5c89701abe66" #Set the correct value from https://docs.microsoft.com/en-us/MicrosoftTeams/sku-reference-edu.
    #For example:
    #Set "314c4481-f395-4525-be8b-2ec4bb1e9d91" for assigning the 'Office 365 Education for Students' to students (STANDARDWOFFPACK_IW_STUDENT)
    #Set "94763226-9b3c-4e75-a931-5c89701abe66" for assigning the 'Office 365 Education for Faculty' to teachers (STANDARDWOFFPACK_FACULTY)
$testOnly = $true #Set the desired value


#########################################################################################################################
cls
$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
$outLogFilePath = "$outDirPath\AssignLicensesToUsers_Results_$LogStartTime.log" 
If (Test-Path $outLogFilePath){
	Remove-Item $outLogFilePath
}

Connect-AzureAD -AccountId $adminName

$countU = 0;
$numberOfNotExisting = 0;
$numberOfNotAlreadyAssigned = 0;
$numberOfAssigned = 0;
$numberOfErrors = 0;


$allUsersToBeCreated = Import-Csv -Path $inputCsvPath -Delimiter $inputCsvDelimiter
$totNumberOfUsersToBeCreated = $allUsersToBeCreated.count
$allUsersToBeCreated | ForEach-Object{
    $loginName = $($_).$colNameLoginName

    $countU++;
    
    $u = $null
    try{
        $u = Get-AzureADUser -ObjectId $loginName        
    }
    catch{
        $line = "[" + $countU +" / " + $totNumberOfUsersToBeCreated + "]: LoginName '" + $loginName + "' - NOT FOUND in Azure AD"; Write-Host $line -ForegroundColor Yellow; $line  | Out-File $outLogFilePath -Append
    }

    if(-not($u -eq $null)){        

        $foundLic = $false

        if( ($u.AssignedLicenses) -and ($u.AssignedLicenses.Count -gt 0) ){

            :innerloop

            ForEach($lic in $u.AssignedLicenses){
                if($lic.SkuId -eq $AssignedLicenseSkuId){
                    $foundLic = $true  
                    break :innerloop                          
                }
            }
        }
        
        if($foundLic){
            $numberOfNotAlreadyAssigned++;
            $line = "[" + $countU +" / " + $totNumberOfUsersToBeCreated + "]: LoginName '" + $loginName + "' - SKIPPED - License already assigned"; Write-Host $line -ForegroundColor Yellow; $line  | Out-File $outLogFilePath -Append
        }
        else{        
            try{
                $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
                $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
                $license.SkuId = $AssignedLicenseSkuId
                $licenses.AddLicenses = $license

                if(-not($testOnly)){
                    Set-AzureADUserLicense -ObjectId $loginName -AssignedLicenses $licenses
                    $numberOfAssigned++;
                    $line = "[" + $countU +" / " + $totNumberOfUsersToBeCreated + "]: LoginName '" + $loginName + "' - License Assigned"; Write-Host $line -ForegroundColor Green; $line  | Out-File $outLogFilePath -Append
                }
                else{
                    $numberOfAssigned++;
                    $line = "[" + $countU +" / " + $totNumberOfUsersToBeCreated + "]: LoginName '" + $loginName + "' - SIMULATED - License Assigned"; Write-Host $line -ForegroundColor Green; $line  | Out-File $outLogFilePath -Append
                } 
            }
            catch{
                $line = "[" + $countU +" / " + $totNumberOfUsersToBeCreated + "]: LoginName '" + $loginName + "' - ERROR - Cannot assign license - " + $_.Exception.Message; Write-Host $line -ForegroundColor Red; $line  | Out-File $outLogFilePath -Append
                $numberOfErrors++;
            }
        }
    }  
        
    $percentComplete = [math]::Round(100*$countU/$totNumberOfUsersToBeCreated,1)
    Write-Progress -Activity "Scanning users" -Status "$percentComplete % users processed" -PercentComplete $percentComplete

}
$line = "Execution completed."; Write-Host $line -ForegroundColor White; $line | Out-File $outLogFilePath -Append
$line = "Users not existing in Azure AD: " + $numberOfNotExisting; Write-Host $line -ForegroundColor White; $line | Out-File $outLogFilePath -Append
$line = "Users already having the required license: " + $numberOfNotAlreadyAssigned; Write-Host $line -ForegroundColor White; $line | Out-File $outLogFilePath -Append
$line = "Errors: " + $numberOfErrors; Write-Host $line -ForegroundColor White; $line | Out-File $outLogFilePath -Append
$line = "User now licensed: " + $numberOfAssigned; Write-Host $line -ForegroundColor White; $line | Out-File $outLogFilePath -Append
$line = "Output log file: " + $outLogFilePath; Write-Host $line -ForegroundColor White; 