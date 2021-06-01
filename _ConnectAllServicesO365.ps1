
##############################################################
# Script-Name: ConnectAllServicesO365.ps1
# Version: 1.3
# Last modified: 31.05.2021 (ADVIS/PTH)
##############################################################
# Connect to Office 365 Services
##############################################################
###################### DEFAULT SETTINGS ######################
##############################################################

# List of Modules you want to install/import
$modules = @(
"ExchangeOnlineManagement",
"MSOnline",
"MicrosoftTeams",
"AzureAD",
"AzureADPreview",
"Az"
)

#-------------------------------------------------------------
# Setup Functions
function SetupModule {
    param ([string]$moduleName)

    $installedModules = Get-Module -ListAvailable

    if ($installedModules.Name -contains $moduleName) {   
        ImportModule($moduleName)
    }
    else { 
        Write-Host "Module $moduleName is not installed. Please wait..." -ForegroundColor Yellow
        InstallModule($moduleName)
    }   
}
function InstallModule {
    param ([string]$moduleName)
    try {
        Install-Module $moduleName -Scope AllUsers -Repository PSGallery -Force -AllowClobber
        Write-Host "Install Module success: $moduleName " -ForegroundColor Green
 
        ImportModule($moduleName)
    }
    catch {
        Write-Host "Install Module failed: $moduleName " $_.Exception.Message -ForegroundColor Red 
    }
}
function ImportModule {
    param ([string]$moduleName)
    try {
        Import-Module -Name $moduleName
        Write-Host "Import Module success: $moduleName " -ForegroundColor Green
    }
    catch {
        Write-Host "Import Module failed: $moduleName " $_.Exception.Message -ForegroundColor Red 
    }       
}

# Start importing/installing modules
Write-Host "Setup modules. Please wait..." -ForegroundColor Yellow
foreach ($module in $modules) {
    SetupModule($module)  
}

##############################################################
################ CHOOSE YOUR CONNECTION ######################
##############################################################
Connect-MsolService
Connect-MicrosoftTeams
Connect-ExchangeOnline
Connect-AzureAD
Connect-AzAccount
        
# Teams Telefonie Login
$sfbSessionTeams = New-CsOnlineSession
Import-PSSession $sfbSessionTeams

##############################################################
###################### SETTINGS ##############################
##############################################################

# Get-Mailbox info@webflow.ch | Select-Object MessageCopyForSentAsEnabled

# Get-EXOMailbox -RecipientTypeDetails UserMailbox | Select-Object Identity, UserPrincipalName

# Get-MsolUser -UserPrincipalName christian.mutschler@rhystadt.ch | Format-List

# Get-MsolUser | Where-Object {($_.licenses).AccountSkuId -match "Teams"}

# Get-Mailbox -Identity labor@hauert.com | Select-Object UserPrincipalName, MessageCopyForSentAsEnabled

# Get-MsolUser -ReturnDeletedUsers

# Remove-MsolUser -UserPrincipalName j.fuhrer@luxshearingch.onmicrosoft.com -RemoveFromRecycleBin

# Set-Mailbox -Identity labor@hauert.com -MessageCopyForSentAsEnabled $True

# Set-MsolUserPrincipalName -UserPrincipalName christian.mutschler.rhystadt@rhystadt.ch -NewUserPrincipalName christian.mutschler@novaproperty.ch

# Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Get-MailboxPermission | Select-Object identity,user,accessrights  | Where-Object { ($_.User -like '*@*')   } | Sort-Object user

# Get-Mailbox -RecipientTypeDetails Usermailbox -ResultSize:Unlimited | Get-MailboxPermission | Select-Object identity,user,accessrights  | Where-Object { ($_.User -like '*@*')   } | Sort-Object user

# Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Select-Object DisplayName,PrimarySmtpAddress,@{Name="EmailAddresses";Expression={$_.EmailAddresses }} | Sort-Object DisplayName

# Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize:Unlimited | Select-Object DisplayName,PrimarySmtpAddress,@{Name="EmailAddresses";Expression={$_.EmailAddresses }} | Sort-Object DisplayName

# Get-Mailbox -ResultSize Unlimited -Filter "ForwardingAddress -like '*' -or ForwardingSmtpAddress -like '*'" | Select-Object Name,ForwardingAddress,ForwardingSmtpAddress

# Set-HostedOutboundSpamFilterPolicy -Identity Default -AutoForwardingMode Off

# Get-Mailbox harry.huwiler@huwiler.ch | Select-Object alias | foreach-object {Get-MailboxFolderStatistics -Identity $_.alias | select-object Identity, ItemsInFolder, FolderSize} | Export-csv c:\temp\harryhuwiler_ExchangeOnline.csv -NoTypeInformation -Encoding Unicode

# Get-PublicFolder -Recurse | Get-PublicFolderClientPermission | Select-Object Identity, User, AccessRights
# Get-PublicFolder -Recurse | Get-PublicFolderClientPermission | Select-Object Identity, User, AccessRights | Where-Object User -Match "Schüpbach Philipp"

# Pruefen ob Benutzer Teams erstellen duerfen
# (Get-AzureADDirectorySetting -Id (Get-AzureADDirectorySetting | Where-object -Property Displayname -Value "Group.Unified" -EQ).id).Values

# Set-HostedContentFilterPolicy -Identity Default -BlockedSenders
# Set-HostedContentFilterPolicy -Identity Default -BlockedSenderDomains

#$teams = Get-Team
#foreach ($team in $teams) {
#      $team.DisplayName
#      $team | Get-TeamUser | Select-Object name,role
#}

# $users = Get-CsOnlineUser
# foreach ($user in $users) {
# $user | Select-Object UserPrincipalName, LineURI, OnPremLineURI, OnlineVoiceRoutingPolicy
# }

# Get-CsOnlineUser -Identity 'salome.hostettler@k-p.ch' | Set-CsUser -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -OnPremLineURI "tel:+41313510404;ext=03" 
# Get-CsOnlineUser -Identity 'salome.hostettler@k-p.ch' | Grant-CsOnlineVoiceRoutingPolicy -PolicyName "SwisscomET4T" 

Disconnect-ExchangeOnline -confirm:$false
Disconnect-AzureAD
