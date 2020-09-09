<#
  .VERSION AND AUTHOR
    Script version: v-2020.08.02
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
    This script creates users in Office 365.
    It accepts in input a CSV with the same format required in the Office 365 Admin Center (https://admin.microsoft.com).
    The advantage of using this script is that you can create more than 250 users at time (250 users per upload is the limit of the Office 365 Admin Center).
    The format of the input CSV is described here: https://docs.microsoft.com/it-IT/microsoft-365/enterprise/add-several-users-at-the-same-time?view=o365-worldwide

    NOTE: in the current verision of the script it is necessary to replace any special characters in the name of the columns of the input CSV.
          For example, a column named "Città" in Italian must be replaced by "Citta" (the "à" must be replaced by "a").
          The special characters on the other data rows are correctly managed and should not be replaced.

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
$inputCsvPath = "C:\Temp\INs\NewUsers.csv" #Set the correct path!
$inputCsvDelimiter = ";" #Set the correct value!
$adminName = "adminName@schoolname.edu" #Set the correct name!
$UsageLocation = "IT" #Set the desired value!
$testOnly = $true #Set the desired value

$colNameLoginName = "Nome utente" #Use local name. Replace special character with ANSI characters (here and in the CSV column name)
$colNameFirstname = "Nome" #Use local value. Use local value. Replace special character with ANSI characters (here and in the CSV column name)
$colNameLastname = "Cognome" #Use local value. Replace special character with ANSI characters (here and in the CSV column name)
$colNameDisplayname = "Nome visualizzato" #Use local value. Replace special character with ANSI characters (here and in the CSV column name)
$colNameJobTitle = "Posizione professionale" #Use local value. Replace special character with ANSI characters (here and in the CSV column name)
$colNameState = "Provincia" #Use local value. Replace special character with ANSI characters (here and in the CSV column name)
$colNameCity = "Citta" #Use local value. Replace special character with ANSI characters (here and in the CSV column name)
$colNamePostalCode = "CAP" #Use local value. Replace special character with ANSI characters (here and in the CSV column name)
$colNameCountry = "Paese" #Use local value. Replace special character with ANSI characters (here and in the CSV column name)



#########################################################################################################################
function GenerateSimpleRandomPassword(){
    $TextInfo = (Get-Culture).TextInfo
    $charsP1 = "bcdfghmnpqrstvwxz".ToCharArray()
    $charsP2 = "aeiou".ToCharArray()
    $charsP3 = "123456789".ToCharArray()
    $newPassword="";$newPasswordP1="";$newPasswordP2=""
    $c1 = $charsP1 | Get-Random
    $c2 = $charsP2 | Get-Random
    $c3 = $charsP1 | Get-Random
    $c4 = $charsP2 | Get-Random
    $newPasswordP1= $c1+$c2+$c3+$c4
    1..4 | ForEach {  $newPasswordP2 += $charsP3 | Get-Random }
    $newPassword = $TextInfo.ToTitleCase($newPasswordP1) + $newPasswordP2
    return $newPassword    
}

cls
$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
$outCsvFilePath = "$outDirPath\CreateNewUsers_Results_$LogStartTime.csv" 
If (Test-Path $outCsvFilePath){
	Remove-Item $outCsvFilePath
}
"LoginName;Password;Result" | Out-File $outCsvFilePath

Connect-AzureAD -AccountId $adminName

$countU = 0;
$numberOfAlreadyExisting = 0;
$numberOfCreated = 0;
$numberOfErrors = 0;


$allUsersToBeCreated = Import-Csv -Path $inputCsvPath -Delimiter $inputCsvDelimiter
$totNumberOfUsersToBeCreated = $allUsersToBeCreated.count
$allUsersToBeCreated | ForEach-Object{
    $loginName = $($_).$colNameLoginName
    $Firstname = $($_).$colNameFirstname
    $Lastname = $($_).$colNameLastname
    $Displayname = $($_).$colNameDisplayname
    $JobTitle = $($_).$colNameJobTitle
    $State = $($_).$colNameState
    $City = $($_).$colNameCity
    $PostalCode = $($_).$colNamePostalCode
    $Country = $($_).$colNameCountry
    $mailNickname = $loginName.Substring(0,$loginName.IndexOf("@"))
    

    $countU++;
    
    $u = $null
    try{
        $u = Get-AzureADUser -ObjectId $loginName 
        $loginName+";"+ $userPwd+";Skipped-AlreadyExisting"| Out-File $outCsvFilePath -Append
        $numberOfAlreadyExisting++;
        $line = "[" + $countU +" / " + $totNumberOfUsersToBeCreated + "]: LoginName '" + $loginName + "' - WARNING - User already existing in Azure AD!"; Write-Host $line -ForegroundColor Yellow; $line 

    }
    catch{}

    if($u -eq $null){        
        $userPwd = GenerateSimpleRandomPassword;
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.Password = $userPwd
        if(-not($testOnly)){
            $success = $false
            try{
                New-AzureADUser -DisplayName $Displayname -MailNickName $mailNickname -GivenName $Firstname -Surname $Lastname -PasswordProfile $PasswordProfile -UserPrincipalName $loginName -AccountEnabled $true -JobTitle $JobTitle -City $City -State $State -PostalCode $PostalCode -Country $Country -UsageLocation $UsageLocation
                $numberOfCreated++;
                $loginName+";"+ $userPwd+";Success"| Out-File $outCsvFilePath -Append                
            }
            catch{
                $numberOfErrors++;
                $loginName+";"+ $userPwd+";Error"| Out-File $outCsvFilePath -Append
                $line = "[" + $countU +" / " + $totNumberOfUsersToBeCreated + "]: LoginName '" + $loginName + "' - ERROR - Could not create the user in Azure AD! " + $_.Exception.Message; Write-Host $line -ForegroundColor Red; $line                 
                pause
            }            
        }
        else{
            $numberOfCreated++;
            "SIMULATED: "+$loginName+";"+ $userPwd+";Success"| Out-File $outCsvFilePath -Append            
        }
    }   
    
    
    $percentComplete = [math]::Round(100*$countU/$totNumberOfUsersToBeCreated,1)
    Write-Progress -Activity "Scanning users" -Status "$percentComplete % users created" -PercentComplete $percentComplete

}
$line = "Execution completed!"; Write-Host $line -ForegroundColor White; 
$line = "User already existing in Azure AD: " + $numberOfAlreadyExisting; Write-Host $line -ForegroundColor White; 
$line = "Errors: " + $numberOfErrors; Write-Host $line -ForegroundColor White; 
$line = "User now created: " + $numberOfCreated; Write-Host $line -ForegroundColor White; 
$line = "Output file (users & passwords): " + $outCsvFilePath; Write-Host $line -ForegroundColor White; 

