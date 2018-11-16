<#
.SYNOPSIS
This script will setup the OWA Disabled Features policy in Exchange for OWA users or groups 
which are not supported by the new Office Add-in model. Such as Tasks, Contacts, Notes and 
Light Mode.

It simplifies the options outlined in the following article:

    https://technet.microsoft.com/en-us/library/dd297989(v=exchg.160).aspx

NOTE: This is only supported in Exchange 2016 or OFfice 365 (Exchange Online)

.DESCRIPTION
Command Line Help:

Usage: Set-OwaDisabledFeatures [-name] [-server] [-policyName] [-disableLight] [-disableContacts] [-disableNotes] [-disableTasks]

[-name]:            The name of the user or group to apply the policy to
[-server]:          On premesis local Exchange server
[-policyName]:      The name of the policy. Defaults to OwaDisabledFeaturesPolicy if left blank
[-identity]:        The user or group to apply the policy to
[-disableLight]:    Disables OWA Light Mode
[-diableContacts]:  Disabled the option to switch to Contacts
[-diableNotes]:     Disabled the option to switch to Notes
[-diableTasks]:     Disabled the option to switch to Tasks

NOTE: If you do not specify anything on the command line it will run in step-by-step 
mode asking you for each value.

.PARAMETER name
This is the email address of the user or the name of the group you want to set the policy for

.PARAMETER disableLight
If present, this will disable the Light Mode in OWA

.PARAMETER disableContacts
If present, this will disable the Contacts view in OWA

.PARAMETER disableNotes
If present, this will disable the Notes view in OWA

.PARAMETER diableTasks
If present, this will disable the Tasks view in OWA

.PARAMETER server
The name of the local Exchange 2016 server if performing this operation on-premesis

.PARAMETER policyName
The name of the policy that these setting will be saved under.
Defaults to OwaDisabledFeaturesPolicy if nothing is specified

.LINK
Policy: https://technet.microsoft.com/en-us/library/dd297989(v=exchg.160).aspx
PowerShell: https://technet.microsoft.com/en-us/library/dd298108(v=exchg.160).aspx

.EXAMPLE
Set-OwaDisabledFeatures -name groupname@contoso.com -disableContacts -disableNotes -disableLight

To disable contacts, notes and light mode for the groupname specified.

.EXAMPLE
Set-OwaDisabledFeatures -name user@contoso.com

To setup Office365 for a single user and being asked for each option

.EXAMPLE
Set-OwaDisabledFeatures

To run in fully interactive/wizard mode, where you will be prompted:
1) User or group
2) User or group name
3) If you want to disable Light Mode
4) If you want to disable Contacts
5) If you want to disable Notes
6) If you want to disable Tasks

.INPUTS
None. You cannot pipe content into this script.

.OUTPUTS
System.String. Displayed output for each completed step with success or failure with error.
#>
Param (
    [string]$name,
    [switch]$user,
    [switch]$group,
    [switch]$disableLight,
    [switch]$disableContacts,
    [switch]$disableNotes,
    [switch]$disableTasks,
    [string]$server,
    [string]$policyName = "OwaDisabledFeaturesPolicy"
)
#Clear-Host
#23456789~123456789~123456789~123456789~123456789~123456789~123456789~1234567890
Write-Host -ForegroundColor:DarkGreen @" 
################################################################################
##                       Set-OwaDisabledFeatures                              ##
##                                                                            ##
## Created by:                                                                ##
## - David E. Craig                                                           ##
## - Version 1.0.1                                                            ##
## - October 24, 2017 11:46AM EST                                             ##
## - http://theofficecontext.com                                              ##
##                                                                            ##
## Usage:                                                                     ##
##                                                                            ## 
##     Set-OwaDisabledFeatures [-name] [-server] [-policyName]                ##
##                             [-disableLight] [-disableContacts]             ## 
##                             [-disableNotes] [-disableTasks]                ##
## Or, for topic help:                                                        ##
##                                                                            ##
##     Get-Help Set-OwaDisabledFeatures -full                                  ##
##                                                                            ##
################################################################################
"@ 
# validate
if((($user -eq $True) -or ($group -eq $True)) -and ((-not $server) -or (-not $policyName))) {
    throw 'Paramaters are invalid. Please see [Get-Help Set-OwaDisabledFeatures.ps1] for more assistance.'
}
if(($user -eq $True) -and ($group -eq $True)) {
    throw 'Parameters are invalid. You can only specify a user or group, not both. Please see [Get-Help Set-OwaDisabledFeatures.ps1] for more assistance.'
}
# if the user did not specify any parameters then we enter the wizard mode
if(($user -eq $False) -and ($group -eq $False)) {
    Write-Host 'No parameters defined. This shell script will run in step-by-step mode.'
    # Ask the user to proceed
    $Answer = Read-Host -Prompt 'Proceed? (Y/n)'
    # If the user answered Y, then proceed
    if (($Answer -eq 'n') -or ($Answer -eq 'N')) {
        Write-Host 'Exited. No changes made.'
        exit
    }
    # Ask the user if it is a group or a user
    $Answer = Read-Host -Prompt 'Install for user or group? (U/g)'
    if (($Answer -eq 'u') -or ($Answer -eq 'U')) {
        $user = $True
        $name = Read-Host -Prompt 'What is the user name?'
    } else {
        $group = $True
        $name = Read-Host -Prompt 'What is the group name?'
    }
    # Ask the user for each setting
    $Answer = Read-Host -Prompt 'Do you want to disable light mode (y/n)'
    if(($Answer -eq 'y') -or ($Answer -eq 'Y')) {
        $disableLight = $True
    }
    $Answer = Read-Host -Prompt 'Do you want to disable Contacts (y/n)'
    if(($Answer -eq 'y') -or ($Answer -eq 'Y')) {
        $disableContacts = $True
    }
    $Answer = Read-Host -Prompt 'Do you want to disable Notes (y/n)'
    if(($Answer -eq 'y') -or ($Answer -eq 'Y')) {
        $disableNotes = $True
    }
    $Answer = Read-Host -Prompt 'Do you want to disable Tasks (y/n)'
    if(($Answer -eq 'y') -or ($Answer -eq 'Y')) {
        $disableTasks = $True
    }
}
# Ask use of this is on-prem of O365
if(-not $server)
{
    # No server specified in the command line so we will prompt the user
    $URLRequest = Read-Host -Prompt 'Is this Exchange Online (Office365) (y/n)?'
    If ($URLRequest -eq 'y')
    {
        # Office 365
        $URL = "https://outlook.office365.com/powershell-liveid/"
    }
    else {
        # On-Prem
        $URL = Read-Host -Prompt 'What is the name (or IP) of your exchange server? ex. exserver01.domain.com, or exserver01 or 192.168.12.11'
        $URL = 'https://' + $URL + '/powershell/'
    } 
} else {
    $URL = 'https://' + $server + '/powershell/'
}
if(-not $policyName) {
    $policyName = "OwaDisabledFeaturesPolicy"
}
Write-Host 'Authentication Required:' -ForegroundColor:Yellow
Write-Host 'You must sign into the Exchange server with an administrator account. Please supply your credentials.' -ForegroundColor:Yellow
$UserCredential = Get-Credential
# connect to the Exchange server
Write-Host 'Connecting to PowerShell session...'
$session_options = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $URL -Credential $UserCredential -Authentication Basic -AllowRedirection -ErrorVariable ProcessError -SessionOption $session_options 
if($ProcessError) {
    Write-Host -ForegroundColor:DarkYellow @"
There was an error trying to get the PowerShell sesssion started. This might be because you have not
enabled PowerShell scripting for the server. Run this command on your server:

    Get-PowerShellVirtualDirectory | Set-PowerShellVirtualDirectory -BasicAuthentication `$true

Please see: 
    https://technet.microsoft.com/en-us/library/dd298108(v=exchg.160).aspx
"@
    exit
}
Import-PSSession $Session -ErrorAction Stop 
$exists = Get-OWAMailboxPolicy $policyName
if(-not $exists) {
    Write-Host 'Creating the ' + $policyName + 'policy...'
    New-OWAMailboxPolicy $policyName -ErrorAction Stop
} else {
    Write-Host 'The' + $policyName + ' policy already exists.'
}
Write-Host 'Configuring the ' + $policyName + ' policy...'
# configure flags for the policy
if($disableContacts -eq $True) {
    $contactsEnabled = $False
} else {
    $contactsEnabled = $True
}
if($disableLight -eq $True) {
    $lightEnabled = $False
} else {
    $lightEnabled = $True
}
if($disableNotes -eq $True) {
    $notesEnabled = $False
} else {
    $notesEnabled = $True
}
if($disableTasks -eq $True) {
    $tasksEnabled = $False
} else {
    $tasksEnabled = $True
}
Get-OWAMailboxPolicy $policyName | Set-OWAMailboxPolicy -TasksEnabled $tasksEnabled -NotesEnabled $notesEnabled -OWALightEnabled $lightEnabled -ContactsEnabled $contactsEnabled -ErrorVariable ProcessError
if($ProcessError) {
    Write-Host -ForegroundColor:DarkYellow @"
There was an error trying to set the Exchange policy. This might have occurred because
you have not updated Active Directory. You will need to run CU6 setup again, but with
the PrepareAD switch:

    Setup.exe /PrepareAD /IAcceptExchangeServerLicenseTerms

Please see: 
    https://technet.microsoft.com/en-us/library/bb125224(v=exchg.160).aspx
"@
    exit
}
Write-Host 'Creating and configuration of ' + $policyName + ' policy complete.'
if($user -eq $True) {
    Write-Host 'Setting policy for user: ' + $name
    Get-User $name -Filter {RecipientTypeDetails -eq $'UserMailBox'}|Set-CASMailbox -OwaMailboxPolicy $policyName -ErrorAction Stop
    Write-Host 'Completed. The ' + $policyName + ' policy has been set for the user ' + $name
} else {
    Write-Host 'Setting policy for group ' + $name
    $targetUsers = Get-Group $name -ErrorAction Stop | Select-Object -ExpandProperty members -ErrorAction Stop
    Write-Host 'Affected users: ' + $targetUsers
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'} -ErrorAction Stop |Set-CASMailbox -OwaMailboxPolicy $policyName -ErrorAction Stop
    Write-Host 'Completed. The ' + $policyName + ' policy has been set for the group ' + $name
}
Write-Host 'Exiting'
exit
