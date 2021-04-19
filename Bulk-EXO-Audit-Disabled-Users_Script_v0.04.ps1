#########################################################################################################
##   
##  Update-All-M365Users-to-have-PSRemoting-Disabled-And-Auditing-Enabled
##
##  Version 0.04
##  
##
##  Created By: Chris Bragg
##  https://github.com/braggaboutme/Bulk-EXO-Audit-Disabled-Users/
##
##  Changelog:
##  v0.01 - 4/15/2021 - First Draft Created
##  v0.02 - 4/16/2021 - Added additional error correction
##  v0.03 - 4/16/2021 - Added logging as a mandatory option which significantly improves performance
##  v0.04 - 4/16/2021 - Boo Boo made, the audit enabled and powershell remoting attributes were set in reverse
##
################################################

#############################

#Get and Set Bulk EXO Users for PowerShell Remoting to Disabled and Auditing enabled

#Within ISE, minimize the function using the minus symbol and you'll see environment specific variables

#############################
function Get-BulkEXOUAuditDisabledUsers
{
 param(
    [Parameter(Mandatory=$true)]
    $CertThumbprint,
    [Parameter(Mandatory=$true)]
    $AppID,
    [Parameter(Mandatory=$true)]
    $TenantName,
    [Parameter(Mandatory=$false)]
    [ValidateSet('O365USGovDoD','O365Default','O365USGovGCCHigh','O365GermanyCloud','O365China')]
    $Environment,
    [Parameter(Mandatory=$false)]
    $InputFile,
    [Parameter(Mandatory=$true)]
    $breakglass
    )
#Sets Azure Environment Name
If ($Environment -eq "O365USGovDoD")
    {
    $Environment = "O365USGovDoD"
    }
ElseIf ($Environment -eq "O365Default")
    {
    $Environment = "O365Default"
    }
ElseIf ($Environment -eq "O365USGovGCCHigh")
    {
    $Environment = "O365USGovGCCHigh"
    }
ElseIf ($Environment -eq "O365GermanyCloud")
    {
    $Environment = "O365GermanyCloud"
    }
ElseIf ($Environment -eq "O365China")
    {
    $Environment = "O365China"
    }
Else
    {
    $Environment = "O365Default"
    }


    #Importing Module
    Try
        {
        Import-Module -Name ExchangeOnlineManagement -MinimumVersion 2.0.4 -ErrorAction Stop
        }
    Catch [System.IO.FileNotFoundException]
        {
        Write-Host "Your Exchange Online Management module version is either not installed, or not at least version 2.0.4" -ForegroundColor Red
        Write-Log "Your Exchange Online Management module version is either not installed, or not at least version 2.0.4"
        }

    #Getting users from CSV
    If ($InputPath -ne $null)
        {
        Write-Host "Getting users from csv" -ForegroundColor Yellow
        Write-Log "Getting users from csv & $_.message"
        $users = get-content -Path $($InputPath)
        }
    
    #Starting timer
    $Sw = [Diagnostics.stopwatch]::StartNew()
    
    #Connecting to Exchange Online
    Try
        {
        Write-Host "Connecting to Exchange Online"
        Write-Log "Connecting to Exchange Online & $_.message"
        $Sw = [Diagnostics.stopwatch]::StartNew()
        Connect-ExchangeOnline -CertificateThumbPrint $CertThumbprint -AppID $AppID -Organization $TenantName -ExchangeEnvironmentName $Environment -ShowBanner:$false -ErrorAction Stop
        }
    Catch [System.ArgumentNullException]
        {
        Write-Host "Missing or incorrect AppID or TenantName" -ForegroundColor Red
        Write-Log "Missing or incorrect AppID or TenantName & $_.message"
        }
    Catch [System.AggregateException]
        {
        Write-Host "Possible Issues
        1) You are possibly using the wrong type of certificate. Please make sure your certificate is a non-CNG cert because CNG certs don't work with this version of the Exchange Online Management PowerShell module
        2) You don't have the correct API's exposed for the Exchange Online Management PowerShell module to connect." -ForegroundColor Red
        Write-Log "Possible Issues
        1) You are possibly using the wrong type of certificate. Please make sure your certificate is a non-CNG cert because CNG certs don't work with this version of the Exchange Online Management PowerShell module
        2) You don't have the correct API's exposed for the Exchange Online Management PowerShell module to connect. & $_.message"
        }
    Catch [System.Management.Automation.Remoting.PSRemotingTransportException]
        {
        Write-Host "Possible Issues
        1) Please make sure you include the proper API permissions in your app registration. When adding the API permissions, search under 'APIs my organization uses' for '00000002-0000-0ff1-ce00-000000000000' to get the Exchange Online APIs. Then grant permissions to the 'Exchange.ManageAsApp' API
        2) You don't have the correct Azure AD Role for your App Registration to perform this task. Please add this App Registration to the Exchange recipient administrator Role if you don't need the auditing. If you require the auditenabled flag set, this app registration will require 'Exchange Administrator'" -ForegroundColor Red
        Write-Log "Possible Issues
        1) Please make sure you include the proper API permissions in your app registration. When adding the API permissions, search under 'APIs my organization uses' for '00000002-0000-0ff1-ce00-000000000000' to get the Exchange Online APIs. Then grant permissions to the 'Exchange.ManageAsApp' API
        2) You don't have the correct Azure AD Role for your App Registration to perform this task. Please add this App Registration to the Exchange recipient administrator Role if you don't need the auditing. If you require the auditenabled flag set, this app registration will require 'Exchange Administrator' & $_.message"
        }


    #Pulling Users from Exchange Online
    If ($InputPath -eq $null)
        {
            Try
                {
                Write-Host "Getting Audit Disabled users from Exchange Online" -ForegroundColor Green
                Write-Log "Getting Audit Disabled users from Exchange Online & $_.message"
                $users = Get-EXOMailbox -ResultSize Unlimited -PropertySets Audit -Filter "AuditEnabled -ne $true -and UserPrincipalName -ne '$($breakglass)'" -ErrorAction Stop
                }
            Catch [Microsoft.Exchange.Management.RestApiClient.RestClientException]
                {
                Write-host "Break Glass account is not defined OR your are not connected to Exchange Online" -ForegroundColor Red
                Write-Log "Break Glass account is not defined OR your are not connected to Exchange Online & $_.message"
                }
            Catch [System.AggregateException]
                {
                Write-host "Invalid Filter" -ForegroundColor Red
                Write-Log "Invalid Filter & $_.Log"
                }
        }
Disconnect-ExchangeOnline -Confirm:$false
$Sw.stop
return $users
}
function Write-Log
{
param($msg)
"$(Get-Date -Format G) : $msg" | Out-File -FilePath $logpath -Append -Force
}
##################
#$CertThumbprint = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
#$AppID = 'XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX'
#$TenantName = 'XXXXXXXXXXXXXXX.ONMICROSOFT.COM'
#$breakglass = 'USERNAME@DOMAIN.onmicrosoft.com'
#$Environment = "O365Default"
#$csv = "C:\users\username\Desktop\allusers.csv"
#$logpath = "C:\users\username\Desktop\TCGScript_LogFile.log"
$sleeptime = '600' #The sleeptime is how long the script will sleep before it starts to run again. Some services will get overloaded if they don't have a pause at least 5 minutes every hour.
$timeout = New-TimeSpan -Minutes 50 #usually 50 minutes before it times out to give the Azure or M365 service some rest time.
##
## There are two options for the get user commands, please read both
##
#OPTION 1: Get from CSV... PREFERRED METHOD... CSV must have no headers with only the UserPrincipalName of all users... CSV pull script not included, but I recommend Graph to pull all users.
#$users = Get-BulkEXOUAuditDisabledUsers -CertThumbprint $CertThumbprint -AppID $AppID -TenantName $TenantName -Environment $Environment -breakglass $breakglass -InputFile $csv | Select UserPrincipalName -ErrorAction Stop
#OPTION 2: Get from tenant... MUCH SLOWER... Higher chance for timeout before the script even runs
Write-Host "Option 2 was selected... Be aware that if you have more than 500K users, this might timeout before it finishes. A CSV might be preferred"
$users = Get-BulkEXOUAuditDisabledUsers -CertThumbprint $CertThumbprint -AppID $AppID -TenantName $TenantName -Environment $Environment -breakglass $breakglass | Select UserPrincipalName -ErrorAction Stop

$Sw = [Diagnostics.stopwatch]::StartNew()
Connect-ExchangeOnline -CertificateThumbPrint $CertThumbprint -AppID $AppID -Organization $TenantName -ExchangeEnvironmentName $Environment -ShowBanner:$false 
    #Processing Users
    foreach ($user in $($users.UserPrincipalName))
        {
        #This while loop checks the timestamp and then re-authenticates every time the stopwatch hits the $timeout variable
        while ($Sw.Elapsed -gt $timeout)
            {
            Write-Log "Disconnecting" -ForegroundColor Yellow
            Write-Log "Current elapsed time is $($sw.elapsed)" -ForegroundColor Yellow
            Disconnect-ExchangeOnline -Confirm:$false
            Write-Log "Last User set was $($user)" -ForegroundColor Green
            Write-Log "Sleeping.... ZZZzzzzzZZZZZzzzz for $($sleeptime) seconds" -ForegroundColor Cyan
            Start-Sleep -Seconds $sleeptime
            write-Log "Connecting" -ForegroundColor Yellow
            Connect-ExchangeOnline -CertificateThumbPrint $CertThumbprint -AppID $AppID -Organization $TenantName -ExchangeEnvironmentName $Environment -ShowBanner:$false -ErrorAction Stop
            $Sw = [Diagnostics.stopwatch]::StartNew()
            Write-Log "Timer Reset" -ForegroundColor Cyan
            }
        #Sets the Audit and PowerShell Remote settings on the user
        If ($user -ne $null)
            {
            Try
                {
                Set-User -Identity $($user) -RemotePowerShellEnabled $false
                Write-Log "$($User) is being set for Remote PowerShell disable"
                }
            Catch [System.Management.Automation.RemoteException]
                {
                Write-Log "There's most likely a problem with the UPN for $($user)"
                }
            Try
                {
                Set-Mailbox -Identity $user -AuditEnabled $true -AuditOwner @{add='MailboxLogin'}
                Write-Log "$($user) enabled for auditing"
                }
            Catch [System.Management.Automation.CommandNotFoundException]
                {
                Write-Log "The app registration running this script doesn't have enough permission in Exchange Online. Please ensure the App registration is granted the 'Exchange Administrator' Role"
                }
            Catch [System.Management.Automation.RemoteException]
                {
                Write-Log "There's most likely a problem with the UPN for $($user)"
                }
            }
    }
$Sw.stop
Disconnect-ExchangeOnline -Confirm:$false
