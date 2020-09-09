<#
  .VERSION AND AUTHOR
    Script version: v-2020.05.17
    Author: Stefano Pescosolido, https://www.linkedin.Qcom/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .SYNOPSIS
  This script creates distribution groups or mail enabled security groups and associates the members specified in the input CSV.
  This is useful, for example, if you want to have groups of students for each class; the groups of this kind are visible 
  while assigning members to a team

  .VARIABLES TO BE SET
    See below

  .PREREQUISITES
  * Use Windows 10 (For earlier versions of Windows, please refer to https://docs.microsoft.com/en-us/microsoft-365/enterprise/connect-to-all-microsoft-365-services-in-a-single-windows-powershell-window?view=o365-worldwide#before-you-begin)
  * If not alreayd done, install the PowerShell module ExchangeOnlineManagement

    To install the PowerShell module, open PowerShell by using the option "Run as administrator" and type:
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
        Install-Module -Name ExchangeOnlineManagement
        
    INFO and prereqs: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module
#>

#########################################################################################################################
# VARIABLES TO BE SET:
#########################################################################################################################
$workingDirPath = "C:\Temp" #This is just a sample
$inputCsvFileName = "Class-Student.csv" #This is just a sample
$inputCsvFileGroupColName = "ClassName" #This is just a sample. Specify the name of the column, in the CSV, containing the desired Group Name
$inputCsvFileUserColName = "LoginName" #This is just a sample. Specify the name of the column, in the CSV, containing the desired User Name 
$inputCsvFileSeparator = ";" #Separator character in the CSV
$adminName = "nomeutente@nomescuola.edu.it" #This is just a sample
$GroupDisplayNamePrefix = "Studenti " #Prefix to be added to the display names of the groups to be created
$GroupEmailPrexif = "studenti-" #Prefix to be added to the nicknames of the groups to be created
$GroupType = "Security" #Accepted values "Security" or "Distribution"
$testOnly = $true

#########################################################################################################################
$LogStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"

$outUsersLogFilePath = "$workingDirPath\CreateGroups-Results_$LogStartTime.log"
If (Test-Path $outUsersLogFilePath){
	Remove-Item $outUsersLogFilePath
}
"START - "+(Get-Date) | Out-File $outUsersLogFilePath

$inputCsvLogFilePath = "$workingDirPath\$inputCsvFileName"
If (-not(Test-Path $inputCsvLogFilePath)){
	Throw "Input CSV $inputCsvLogFilePath not found"
}

$inputFile = Import-Csv $inputCsvLogFilePath -Delimiter $inputCsvFileSeparator 
Connect-ExchangeOnline -EnableErrorReporting -UserPrincipalName $adminName 
cls

$inputFile | ForEach-Object {

    $class = $_.'$inputCsvFileGroupColName'
    $upn = $_.'$inputCsvFileUserColName'

    $dgAlias = $GroupEmailPrexif+$class
    $dgDN = $GroupDisplayNamePrefix+$class

    $line = $class+" - "+$upn; Write-Host $line -ForegroundColor Green; $line | Out-File $outUsersLogFilePath -Append
    
    $dg = $null
    $dg = Get-DistributionGroup -Filter "(Alias -eq '$dgAlias')"
    
    if($dg -eq $null){
        $line = "  "+$dgAlias+" does not exists"; Write-Host $line -ForegroundColor Cyan; $line | Out-File $outUsersLogFilePath -Append
        if(-not($testOnly)){
            $dg = New-DistributionGroup -Alias $dgAlias -DisplayName $dgDN -Name $dgDN -ManagedBy $adminName -Type $GroupType
            $newDgIdentity = $dg.Identity
            $line = "  Created DG: "+$newDgIdentity; Write-Host $line -ForegroundColor Cyan; $line | Out-File $outUsersLogFilePath -Append
        }
        else{
            $line = "  [Simulated] Created DG: "+$dgAlias; Write-Host $line -ForegroundColor Cyan; $line | Out-File $outUsersLogFilePath -Append
        }            
    }
    else{
        $line = "  "+$dgAlias+" already exists"; Write-Host $line -ForegroundColor Gray; $line | Out-File $outUsersLogFilePath -Append
    }

    if($dg -ne $null){
        $uAlias = $upn.Substring(0,$upn.IndexOf("@"))
        $eu = Get-DistributionGroupMember -Identity $dg.Identity | Where 'RecipientType' -eq 'UserMailbox' | where 'Name' -eq $uAlias

        if($eu){
            $line = "  "+$uAlias+" is already member of Distribution Group "+$dg.Identity; Write-Host $line -ForegroundColor Gray; $line | Out-File $outUsersLogFilePath -Append
        }
        else{
            Add-DistributionGroupMember -Identity $dg.Identity -Member $upn
            $line = "  Added "+$upn+" as member of Distribution Group "+$dg.Identity; Write-Host $line -ForegroundColor Green; $line | Out-File $outUsersLogFilePath -Append
        }
    }
    else{
        if(-$testOnly){
            $line = "  [Simulated] Check if user '"+$upn+"' is member in DG: "+$dgAlias; Write-Host $line -ForegroundColor Cyan; $line | Out-File $outUsersLogFilePath -Append
        }
        else{
            $line = "  [Error] Could not find or create the DG "+$dgAlias; Write-Host $line -ForegroundColor Red; $line | Out-File $outUsersLogFilePath -Append
        }
    }
}

"END - "+(Get-Date) | Out-File $outUsersLogFilePath -Append
Write-Host "Execution ended. Output file: " $outUsersLogFilePath