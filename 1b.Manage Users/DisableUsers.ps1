<#
  .VERSION AND AUTHOR
    Script version: v-2020.08.02
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
    This script disable the users specified in the input CSV.

  .PREREQUISITES
  * Use Windows 10 (For earlier versions of Windows, please refer to https://docs.microsoft.com/en-us/microsoft-365/enterprise/connect-to-all-microsoft-365-services-in-a-single-windows-powershell-window?view=o365-worldwide#before-you-begin)
  * If not alreayd done, install the PowerShell module Azure AD (or AzureADPreview) 

    To install the PowerShell module, open PowerShell by using the option "Run as administrator" and type:
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
$colNameLoginName = "LoginName" #Name of the column with the "login name" of the user. Do not use special characters in this column name here (only ANSI characters)
$adminName = "adminName@schoolname.edu" #Set the correct name!
$testOnly = $true #Set the desired value


#########################################################################################################################
cls
$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
$outLogFilePath = "$outDirPath\DisableUsers_Results_$LogStartTime.log"
If (Test-Path $outLogFilePath){
	Remove-Item $outLogFilePath
}

Connect-AzureAD -AccountId $adminName

$countU = 0;
$numberOfNotFoundInAzureAD = 0;
$numberOfAlreadyDisabled = 0;
$numberOfDisabled = 0;


$allUsersToBeDisabled = Import-Csv -Path $inputCsvPath
$totNumberOfUsersToBeDisabled = $allUsersToBeDisabled.count
$allUsersToBeDisabled | ForEach-Object{
    $loginName = $($_).$colNameLoginName
    $countU++;
    
    $u = $null
    try{
        $u = Get-AzureADUser -ObjectId $loginName         
    }
    catch{
        $line = "[" + $countU +" / " + $totNumberOfUsersToBeDisabled + "]: LoginName '" + $loginName + "' - CANNOT FIND USER IN AZURE AD"; Write-Host $line -ForegroundColor Red; $line | Out-File $outLogFilePath -Append
    }

    if($u){        
        if($u.AccountEnabled){
            if(-not($testOnly)){
                Set-AzureADUser -ObjectId $loginName -AccountEnabled $false
                $numberOfDisabled++;
                $line = "[" + $countU +" / " + $totNumberOfUsersToBeDisabled + "]: LoginName '" + $loginName + "' - User disabled successfully!"; Write-Host $line -ForegroundColor Green; $line | Out-File $outLogFilePath -Append
            }
            else{
                $numberOfDisabled++;
                $line = "[" + $countU +" / " + $totNumberOfUsersToBeDisabled + "]: LoginName '" + $loginName + "' - SIMULATION - User disabled successfully!"; Write-Host $line -ForegroundColor Green; $line | Out-File $outLogFilePath -Append
            }
        }
        else{
            $line = "[" + $countU +" / " + $totNumberOfUsersToBeDisabled + "]: LoginName '" + $loginName + "' - User already disabled"; Write-Host $line -ForegroundColor Yellow; $line | Out-File $outLogFilePath -Append
            $numberOfAlreadyDisabled++;
        }
    }   
    else{
        $numberOfNotFoundInAzureAD++;
    } 
    
    
    $percentComplete = [math]::Round(100*$countU/$totNumberOfUsersToBeDisabled,1)
    Write-Progress -Activity "Scanning users" -Status "$percentComplete % users analyzed" -PercentComplete $percentComplete

}
$line = "Execution completed."; Write-Host $line -ForegroundColor White; $line | Out-File $outLogFilePath -Append
$line = "User not found: " + $numberOfNotFoundInAzureAD; Write-Host $line -ForegroundColor White; $line | Out-File $outLogFilePath -Append
$line = "User already disabled: " + $numberOfAlreadyDisabled; Write-Host $line -ForegroundColor White; $line | Out-File $outLogFilePath -Append
$line = "User now disabled: " + $numberOfDisabled; Write-Host $line -ForegroundColor White; $line | Out-File $outLogFilePath -Append
$line = "Output log file: " + $outLogFilePath; Write-Host $line -ForegroundColor White; 