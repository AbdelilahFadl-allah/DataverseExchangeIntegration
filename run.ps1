param(
    [Parameter(Mandatory=$false)]
    [string]$WorkingDirectory = $PSScriptRoot
)

# Resolve a sensible default if $PSScriptRoot is empty
if (-not $WorkingDirectory -or $WorkingDirectory -eq '') {
    $WorkingDirectory = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}

Write-Host "Using working directory: $WorkingDirectory"

Set-Location -Path $WorkingDirectory

# Expose working directory to modules loaded in other scopes/jobs
$env:DVEXCHG_WORKING_DIRECTORY = $WorkingDirectory

Import-Module (Join-Path $WorkingDirectory 'test_mod.psm1')

#Logging Setup
$loggingConfig = Get-PSConfig | Select-Object -ExpandProperty Logging
Set-LoggingDefaultLevel -Level $loggingConfig.LogLevel
Add-LoggingTarget -Name "file" -Configuration @{ Path = $loggingConfig.LogFile }



$dvservice = Get-DataverseService
$dvmailboxes = Get-DataverseIncomingMailbox -service $dvservice
$outmailboxes = Get-DataverseOutgoingMailbox -service $dvservice
$mg_token = Get-MgAuthToken


# Function to process a single mailbox
Function Process-Mailbox {
    param(
        [psobject]$mb,
        [psobject]$dvservice,
        [psobject]$token
    )
        
    ######## Process Incoming Emails ###########

    $messages = Get-ExchangeMessages -upn $mb.emailaddress -start $mb.start_processing -token $token
    foreach ($our_message in $messages) {
        $emailid = $null
        #In Case of a user mailbox 
        if ($null -ne $mb.userid) {
            $userpref = Get-UserTrackingPreferences -UserId $mb.userid -service $dvservice
            if ("all" -eq $userpref) {
                $emailid = Add-IncomingEmailInDataverse -service $dvservice -newemail $our_message -upn $mb.emailaddress
            }
            elseif ("correlate" -eq $userpref) {
                $correlatedDvEmail = Get-CorrelatedEmail -service $dvservice -inreplyto ($our_message.inreplyto.Replace("<", "\u003C").Replace(">", "\u003E")) -newEmail $our_message
                if ($null -ne $correlatedDvEmail -and $null -ne $correlatedDvEmail.id) {
                    $emailid = Add-IncomingEmailInDataverse -service $dvservice -newemail $our_message -correlatedemail $correlatedDvEmail -upn ($mb.emailaddress)
                }
            }
        }
        #In Case of a queue, we will just add the email without correlation for the moment, even though the correlation works the same way, just the preferences configs are on the queue
        elseif ($null -ne $mb.queueid -and $null -ne $our_message.messageid) {
            Write-Host "Processing email: $($our_message.subject): $($our_message.messageid)"
            #In Case of a queue, we will just add the email without correlation or preference check
            $emailid = Add-IncomingEmailInDataverse -service $dvservice -newemail $our_message -upn ($mb.emailaddress)
            # add queueitem
            if ($null -ne $emailid) {
                Add-QueueItem -service $dvservice -queueid $mb.queueid -objectid ($emailid | Select-Object -last 1) -objecttype "email"
            }
        }
    }
    
    Write-Host "Completed processing for: $($mb.emailaddress)"
}


Function Send-PendingEmails {
    Param(
        [psobject]$mb,
        [psobject]$dvservice,
        [psobject]$token
    )
    begin {
        Write-Host "Checking for pending outgoing emails for: $($mb.emailaddress)"

        $out_messages = Get-DataverseOutgoingEmail -service $dvservice -mb $mb

        foreach ($email in $out_messages) {
            Write-Host "Sending email: $($email.subject) to $($mb.emailaddress)"
            Send-Email -message $email -service $dvservice -mailbox $mb -token $token
        }
    }
}

foreach($upn in $dvmailboxes) {
    try {
        Process-Mailbox -mb $upn -dvservice $dvservice -token $mg_token
    }
    catch{
        Write-Log -Message "Error processing incoming mailbox $($upn.emailaddress): $_" -Level ERROR
        Write-log -Message "Error Message: $($_.Exception.Message)" -Level ERROR
        Write-log -Message " ..Stack Trace: $($_.Exception.Message)" -Level ERROR
    }
}

foreach($mb in $outmailboxes) {
    try {
        Send-PendingEmails -mb $mb -dvservice $dvservice -token $mg_token
    }
    catch {
        Write-Log -Message "Error processing outgoing mailbox $($mb.emailaddress): $_" -Level ERROR
        Write-log -Message "Error Message: $($_.Exception.Message)" -Level ERROR
        Write-log -Message " ..Stack Trace: $($_.Exception.Message)" -Level ERROR
    }
}



<#TODO: Parallel Processing#>
