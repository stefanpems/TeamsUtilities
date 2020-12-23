<#
  .VERSION AND AUTHOR
    Script version: v-2020.12.22
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .ATTENTION (UPDATE 2020.12.22)
  The script doesn't work with the October 2020 version of PnP.PowerShell / SharePointPnPPowerShellOnline. NOTE: I haven't checked with any newer version.
  We are investigating. While we find a solution, please use the September 2020 version of the module.
  So, DO NOT install it by simply launching the command referenced below: Install-Module SharePointPnPPowerShellOnline  
  Instead, check if you have a newer version already installed; if so, remove it; then install the previous version.
  => To check if you have the issue, use the command: Connect-PnPOnline -Scopes "User.Read" 
     If you have the issue, you get an error AADSTS70011. If you get a login prompt, it means that you don't have the issue. 
  => To check which version is installed use the command: Get-Module -Name sharepointpnppowershell* -ListAvailable
  => To remove any existing version use the command: Uninstall-Module SharePointPnPPowerShellOnline -AllVersions -Force"
  => To install the previous (last working) version, use the command: "Install-Module SharePointPnPPowerShellOnline -RequiredVersion 3.25.2009.1"
   
  More info on https://github.com/pnp/PnP-PowerShell/issues/2983

  .SYNOPSIS
    This script massively writes the mobile number to be used by the Azure AD users for MFA. 
    The users (UPN) & mobile numbers are read from an input CSV file.
    NOTES: 
     - The mobile number attribute that is set is only the one used for the strong authentication (MFA), not the one shown in the user profile.
     - If the user has already a mobile number for strong authentication (MFA), that number will not be changed.
     - The script produces in output a CSV file with 3 columns: "UPN;PhoneNumber;Updated". 
       For each user specified in input, the output file specifies the one of these possibile results for the column "Updatesd:
        * True: the mobile phone has been set
        * Simulated: the mobile phone would have been set but the execution was simulated
        * False: the mobile phone is already set (copied in the column "PhoneNumber") and was not changed
        * Not found: the user specified in the row of the input file was not found in Azure AD
    
  .PREREQUISITES
   * Use Windows 10 (For earlier versions of Windows, please refer to https://docs.microsoft.com/en-us/microsoft-365/enterprise/connect-to-all-microsoft-365-services-in-a-single-windows-powershell-window?view=o365-worldwide#before-you-begin)
   * If not alreayd done, install the PowerShell modules MSOnline and SharePointPnPPowerShellOnline

    To install the PowerShell modules, open PowerShell by using the option "Run as administrator" and type:
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned 
        Install-Module -Name MSOnline
        Install-Module SharePointPnPPowerShellOnline ---> Use " -RequiredVersion 3.25.2009.1" (see note above)
    Details here: https://www.powershellgallery.com/packages/MSOnline/
    and here: https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-all-office-365-services-in-a-single-windows-powershell-window

  .VARIABLES TO BE SET
  Read below.
#>

#########################################################################################################################
# VARIABLES TO BE SET:
#########################################################################################################################
$inCsvFile = "C:\temp\in.csv" #Set the correct path. Minimum expected columns: Username, PhoneNumber
$inCsvDelimiter = ";" #Set the correct value
$outLogDir = "C:\temp" #Set the correct path! Here the script creates a CSV file 
$testOnly = $true #Set the desired value


######################################################################
function SetPhoneNumberForUser{

    param(
         [string]
         $upn, 
         [string]
         $phoneNumber, 
         [string]
         $accessToken
    )

    $res = $false;
    $graphURL = "https://graph.microsoft.com/beta" 
    $authMethodUrl = "$graphURL/users/$upn/authentication/phoneMethods" 
    
    $body = @{
        "phoneType" = "mobile"
        "phoneNumber" = $phoneNumber
    }
        
    try
    {
        $response = Invoke-RestMethod -Uri $authMethodUrl -Headers @{Authorization = "Bearer $accesstoken"} -ContentType "application/json" -Method Post -Body ($body | ConvertTo-Json) -Verbose -UseBasicParsing ;
        if($response){
            $res = $true;
        }                    
    }
    catch
    {
        Write-Host "Error while executing query - " $_.Exception.Message -ForegroundColor Red
        throw $_
    }    
}

#Initializations
cls
$ExecStartTime = Get-Date -Format "yyyy-MM-dd_hh.mm.ss"
$setPhoneNumbersFile = "$outLogDir\setPhoneNumbers_$ExecStartTime.csv"

If (Test-Path $setPhoneNumbersFile)
{
	Remove-Item $setPhoneNumbersFile
}
"UPN;PhoneNumber;Updated"| Out-File $setPhoneNumbersFile

#Connections
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue 
$arrayOfScopes = @("UserAuthenticationMethod.Read.All","UserAuthenticationMethod.ReadWrite.All") 

Connect-MsolService
Connect-PnPOnline -Scopes $arrayOfScopes
$accessToken = Get-PnPGraphAccessToken -Decoded 

#Rows loop
$totCount = 0;
$mnCount = 0;
$vCount = 0;
Write-Host "Execution started - Please wait..." -ForegroundColor Yellow
$allUsers = Import-Csv -Path $inCsvFile -Delimiter $inCsvDelimiter

$numUsers = $allUsers.Count
$allUsers | ForEach-Object {
    $r = $($_)
    $totCount++;
    $upn = $r.Username
    $pn = $r.PhoneNumber

    #Search the user in Azure AD by using the legacy MSOL module
    $u = $null
    try{
        $u = Get-MsolUser -UserPrincipalName $upn -ErrorAction SilentlyContinue
    }
    catch{}

    if($u -eq $null){
        $upn + ";" + $pn + ";Not found" | Out-File $setPhoneNumbersFile -Append
    }
    else{

        if( ($u.UserType.toString() -ne "Guest") -and (-not($u.BlockCredential)) ) {
            $vCount++;

            if($u.StrongAuthenticationUserDetails -and $u.StrongAuthenticationUserDetails.PhoneNumber){
                $upn + ";"+ $u.StrongAuthenticationUserDetails.PhoneNumber + ";False" | Out-File $setPhoneNumbersFile -Append
            }
            else{
                $mnCount++;
                if($testOnly){
                    $upn + ";" + $pn + ";Simulated" | Out-File $setPhoneNumbersFile -Append
                }
                else{
                    try{
                        $res = SetPhoneNumberForUser -accessToken $accessToken.RawData -upn $upn -phoneNumber $pn 
                            
                        $upn + ";" + $pn + ";True" | Out-File $setPhoneNumbersFile -Append
                    }
                    catch{
                        $upn + ";" + $pn + ";Error" | Out-File $setPhoneNumbersFile -Append
                    }
                }
            }
        }
    }


    #Manage progress bar
    $percentComplete = [math]::Round(100*$totCount/$numUsers,1)
    Write-Progress -Activity "Scanning users" -Status "$percentComplete % users analyzed" -PercentComplete $percentComplete

    #Manage authentication token expiration
    $tokenExpiration = $accessToken.ValidTo.ToLocalTime().AddMinutes(-1);
    $TimeToExpiry = $tokenExpiration - (Get-Date)        
    $sTimeToExpiry = $TimeToExpiry.Minutes.ToString() + " min " + $TimeToExpiry.Seconds.ToString() + " sec"
                
    Write-Host "  Token - TimeToExpiry: " $sTimeToExpiry
    "  Token - TimeToExpiry: '$sTimeToExpiry'"             

    $IsExpired = (Get-Date) -gt $tokenExpiration   

    if($IsExpired){
        Write-Host "  Token expired! Acquiring a new token" -ForegroundColor Cyan 
        "  Token expired! Acquiring a new token."      
            
        $accessToken = $null;
            
        try{
            Connect-PnPOnline -Scopes $arrayOfScopes 
            $accessToken = Get-PnPAccessToken -Decoded      
        }
        catch{
            Write-Host "  Could not acquire a new token. Forced exit!" -ForegroundColor Red 
            "  Could not acquire a new token. Forced exit!"      
            Break TeamsLoop
        }

        $tokenExpiration = $accessToken.ValidTo.ToLocalTime();
        $TimeToExpiry = $tokenExpiration - (Get-Date)        
        $sTimeToExpiry = $TimeToExpiry.Minutes.ToString() + " min " + $TimeToExpiry.Seconds.ToString() + " sec"
                
        Write-Host "  New Token - TimeToExpiry: " $sTimeToExpiry
        "  New Token - TimeToExpiry: '$sTimeToExpiry'"             

    }
}

Write-Host "Execution ended - Users whose mobile number has been set: $mnCount/$vCount - Output CSV file: '" $setPhoneNumbersFile "'" -ForegroundColor Yellow
