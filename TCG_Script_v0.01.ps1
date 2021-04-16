#########################################################################################################
##   
##  Update-All-M365Users-to-have-PSRemoting-Disabled-And-Auditing-Enabled
##
##  Version 0.01
##  
##  ##  https://github.com/lithnet/resourcemanagement-webservice/
##
##  Created By: Chris Bragg
##  https://github.com/braggaboutme/MIM-Portal-PSModule
##
##  Changelog:
##  v0.01 - 4/15/2021 - First Draft Created
##
################################################

#############################

#Get and Set Bulk EXO Users for PowerShell Remoting to Disabled and Auditing enabled

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
        }

    #Getting users from CSV
    If ($InputPath -ne $null)
        {
        Write-Host "Getting users from csv" -ForegroundColor Yellow
        $users = get-content -Path $($InputPath)
        }
    
    #Starting timer
    $Sw = [Diagnostics.stopwatch]::StartNew()
    
    #Connecting to Exchange Online
    Try
        {
        Write-Host "Connecting to Exchange Online"
        $Sw = [Diagnostics.stopwatch]::StartNew()
        Connect-ExchangeOnline -CertificateThumbPrint $CertThumbprint -AppID $AppID -Organization $TenantName -ExchangeEnvironmentName $Environment -ErrorAction Stop
        }
    Catch [System.ArgumentNullException]
        {
        Write-Host "Missing or incorrect AppID or TenantName" -ForegroundColor Red
        }

    #Pulling Users from Exchange Online
    If ($InputPath -eq $null)
        {
            Try
                {
                Write-Host "Getting Audit Disabled users from Exchange Online"
                $users = Get-EXOMailbox -ResultSize Unlimited -PropertySets Audit -Filter "AuditEnabled -ne $false -and UserPrincipalName -ne '$($breakglass)'" -ErrorAction Stop
                }
            Catch [Microsoft.Exchange.Management.RestApiClient.RestClientException]
                {
                Write-host "Break Glass account is not defined" -ForegroundColor Red
                }
            Catch [System.AggregateException]
                {
                Write-host "Invalid Filter" -ForegroundColor Red
                }
        }
Disconnect-ExchangeOnline -Confirm:$false
$Sw.stop
return $users
 }

##################
$CertThumbprint = '697A3A003B73B21E91F8503CFA7B4F0078C2CC5B'
$AppID = '0bc7268e-65b5-409d-b2a0-8171027fc3f0'
$TenantName = 'BRAGGABOUTMYCLOUD.ONMICROSOFT.COM'
$breakglass = 'christopher_bragg@braggaboutmycloud.onmicrosoft.com'
$Environment = "O365Default"
$csv = "C:\users\username\Desktop\allusers.csv"
$sleeptime = '600' #The sleeptime is how long the script will sleep before it starts to run again. Some services will get overloaded if they don't have a pause at least 5 minutes every hour.
$timeout = New-TimeSpan -Minutes 50 #usually 50 minutes before it times out to give the Azure or M365 service some rest time.
#################    


##
## There are two options for the get user commands, please read both
##
#OPTION 1: Get from CSV... PREFERRED METHOD... CSV must have no headers with only the UserPrincipalName of all users... CSV pull script not included, but I recommend Graph to pull all users.
#$users = Get-BulkEXOUAuditDisabledUsers -CertThumbprint $CertThumbprint -AppID $AppID -TenantName $TenantName -Environment $Environment -breakglass $breakglass -InputFile $csv | Select UserPrincipalName
#OPTION 2: Get from tenant... MUCH SLOWER... Higher chance for timeout before the script even runs
$users = Get-BulkEXOUAuditDisabledUsers -CertThumbprint $CertThumbprint -AppID $AppID -TenantName $TenantName -Environment $Environment -breakglass $breakglass | Select UserPrincipalName


$Sw = [Diagnostics.stopwatch]::StartNew()
Connect-ExchangeOnline -CertificateThumbPrint $CertThumbprint -AppID $AppID -Organization $TenantName -ExchangeEnvironmentName $Environment
    #Processing Users
    foreach ($user in $users.UserPrincipalName)
        {
        #This while loop checks the timestamp and then re-authenticates every time the stopwatch hits the $timeout variable
        while ($Sw.Elapsed -gt $timeout)
            {
            Write-Host "Disconnecting" -ForegroundColor Yellow
            Write-Host "Current elapsed time is $($sw.elapsed)" -ForegroundColor Yellow
            Disconnect-ExchangeOnline -Confirm:$false
            Write-Host "Sleeping.... ZZZzzzzzZZZZZzzzz for $($sleeptime) seconds" -ForegroundColor Cyan
            Start-Sleep -Seconds $sleeptime
            write-host "Connecting" -ForegroundColor Yellow
            Connect-ExchangeOnline -CertificateThumbPrint $CertThumbprint -AppID $AppID -Organization $TenantName -ExchangeEnvironmentName $Environment
            $Sw = [Diagnostics.stopwatch]::StartNew()
            Write-Host "Timer Reset" -ForegroundColor Cyan
            }

        #Sets the Audit and PowerShell Remote settings on the user, comment out all Write-Host sections here for better performance
        Set-User -Identity $user -RemotePowerShellEnabled $false
        Write-Host "$User disabled for remote PowerShell" -ForegroundColor DarkGreen
        Set-Mailbox -Identity $user -AuditEnabled $true -AuditOwner @{add='MailboxLogin'}
        Write-Host "$($User) enabled for auditing" -ForegroundColor Magenta
        
    }
$Sw.stop
Disconnect-ExchangeOnline -Confirm:$false