<#
  .VERSION AND AUTHOR
    Script version: v-2020.09.05
    Author: Stefano Pescosolido, https://www.linkedin.com/in/stefanopescosolido/
    Script published in GitHub: https://github.com/stefanpems/TeamsUtilities

  .ATTENTION (UPDATE 2020.10.25)
  The scripts in this workspace that use GraphAPI do not work with the latest version (Oct. 2020 - v. 3.26.2010.0) of the module 
  PnP.PowerShell / SharePointPnPPowerShellOnline.
  We are investigating. While we find a solution, please do not install the latest version or upgrade your current version, so
  DO NOT install it by simply launching the command referenced below: Install-Module SharePointPnPPowerShellOnline  
  Instead, check if you have the latest version already installed; if so, remove it; then install the previous version.
  => To check if you have the issue, use the command: Connect-PnPOnline -Scopes "Mail.Read","Files.Read","ActivityFeed.Read" 
     If you have the issue, you get an error AADSTS70011. If you don't have the issue, you get a login prompt
  => To check which version is installed use the command: Get-Module -Name sharepointpnppowershell* -ListAvailable
  => To remove any existing version use the command: Uninstall-Module SharePointPnPPowerShellOnline -AllVersions -Force"
  => To install the previous (last working) version, use the command: "Install-Module SharePointPnPPowerShellOnline -RequiredVersion 3.25.2009.1"
   
  More info on https://github.com/pnp/PnP-PowerShell/issues/2983

  .SYNOPSIS
  This script install all the modules that can be used to manage the objects in Microsoft Teams
  and related technologies.

  .NOTES
  Depending on what you need to do, you may not need all these modules. 
  The recommendation is to install only the modules needed in your script.

  .MORE INFO
  https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-all-office-365-services-in-a-single-windows-powershell-window

  .PREREQUISITES
  * Use Windows 10 (For earlier versions of Windows, please refer to https://docs.microsoft.com/en-us/microsoft-365/enterprise/connect-to-all-microsoft-365-services-in-a-single-windows-powershell-window?view=o365-worldwide#before-you-begin)
  * Open PowerShell window by using the option "Run as administrator"
#>

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force

Install-Module MSOnline -Force
#INFO: https://www.powershellgallery.com/packages/MSOnline/

Install-Module MicrosoftTeams -Force
#INFO and prereqs: https://docs.microsoft.com/en-us/MicrosoftTeams/teams-powershell-install

Install-Module AzureADPreview -Force #or install AzureAD (not the preview version)
#INFO: https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0

#The commands below are not needed in Windows 10/Windows Server 2016 or newer!
#Install-Module PowershellGet -Force
#Update-Module PowershellGet
#INFO: https://docs.microsoft.com/en-us/powershell/scripting/gallery/installing-psget?view=powershell-7
#      https://docs.microsoft.com/en-us/powershell/module/powershellget/update-module?view=powershell-7

Install-Module SharePointPnPPowerShellOnline -Force
#INFO: https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-pnppowershell

Install-Module -Name ExchangeOnlineManagement -Force
#INFO and prereqs: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module