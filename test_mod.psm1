Using namespace System.Net
using namespace Microsoft.Xrm.Sdk
using namespace Microsoft.Xrm.Sdk.Query
using namespace Microsoft.Crm.Sdk.Messages
using namespace Microsoft.Xrm.Sdk.Messages
using namespace System.Collections.Generic


## Determine module working directory: prefer env var, then PSScriptRoot
$script:WorkingDirectory = $env:DVEXCHG_WORKING_DIRECTORY
if (-not $script:WorkingDirectory -or $script:WorkingDirectory -eq '') {
    if ($PSScriptRoot -and (Test-Path $PSScriptRoot)) {
        $script:WorkingDirectory = $PSScriptRoot
    }
    else {
        $script:WorkingDirectory = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
    }
}

Import-Module (Join-Path $script:WorkingDirectory 'LoadPackages.psm1')

Load-Packages

Function Get-InMailboxConfig {
    param(
        [string]$ConfigPath = "$script:WorkingDirectory\in_mailbox.json"
    )
    begin {
        Write-Log -Message "Loading incoming mailbox configuration from: $ConfigPath" -Level INFO
        if (-not (Test-Path $ConfigPath)) {
            Write-Log -Message "Configuration file not found: $ConfigPath" -Level ERROR
            throw "Configuration file not found: $ConfigPath"
        }

        $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        return $config
    }
}

Function Get-OutMailboxConfig {
    param(
        [string]$ConfigPath = "$script:WorkingDirectory\out_mailbox.json"
    )
    begin {
        Write-Log -Message "Loading outgoing mailbox configuration from: $ConfigPath" -Level INFO
        if (-not (Test-Path $ConfigPath)) {
            Write-Log -Message "Configuration file not found: $ConfigPath" -Level ERROR
            throw "Configuration file not found: $ConfigPath"
        }

        $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        return $config
    }
}


Function Get-PSConfig {
    param(
        [string]$ConfigPath = "$script:WorkingDirectory\config.json"
    )
    begin {
        if (-not (Test-Path $ConfigPath)) {
            Write-Log -Message "Configuration file not found: $ConfigPath" -Level ERROR
            throw "Configuration file not found: $ConfigPath"
        }

        $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        return $config
    }
}

Function Get-DataverseService {
    param()
    begin {
        $config = Get-PSConfig
        $dataverse = $config.Dataverse
        if ($null -eq $dataverse) {
            throw "Dataverse configuration missing in config file."
        }

        $connStr = @"
            AuthType=$($dataverse.AuthType);
            Url=$($dataverse.Url);
            ClientId=$($dataverse.ClientId);
            ClientSecret=$($dataverse.ClientSecret);
"@
        Write-Log -Message "Establishing connection to Dataverse with URL: $($dataverse.Url)" -Level INFO
        $service = Get-CrmConnection -ConnectionString $connStr

        Write-Output $service
    }
}

Function Get-DataverseIncomingMailbox {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]
        $service
    )
    begin {
        Write-Log -Message "Retrieving incoming mailboxes from Dataverse" -Level INFO
        $integratedIncomingMailboxes = Get-InMailboxConfig
        $inmailbox_query = $integratedIncomingMailboxes | Select-Object -ExpandProperty email |foreach-object { "<value>$($_)</value>" } | Out-String
        $fetch = @"
            <fetch>
                <entity name="mailbox">
                    <attribute name="emailaddress" />
                    <attribute name="mailboxid" />
                    <filter>
                        <condition attribute="incomingemaildeliverymethod" operator="eq" value="2" />
                        <condition attribute="statecode" operator="eq" value="0" />
                        <condition attribute="emailaddress" operator="in">
                            $inmailbox_query
                        </condition>
                    </filter>
                    <link-entity name="queue" from="queueid" to="regardingobjectid" link-type="outer" alias="queue">
                        <attribute name="queueid" />
                    </link-entity>
                    <link-entity name="systemuser" from="systemuserid" to="regardingobjectid" link-type="outer" alias="user">
                        <attribute name="systemuserid" />
                        <filter>
                            <condition attribute="islicensed" operator="eq" value="true" />
                        </filter>
                    </link-entity>
                </entity>
                </fetch>
"@
        $fetchEx = New-Object FetchExpression($fetch)
        $result = $service.RetrieveMultiple($fetchEx)
        $mailboxes = @()
        foreach ($email in $result.Entities) {
            $outo = @{
                emailaddress = [string]$email.Attributes["emailaddress"]
                mailboxid    = [guid]$email.Id
                userid       = if ($email.Attributes.Contains("user.systemuserid")) { [guid]([AliasedValue]$email.Attributes["user.systemuserid"]).Value }else { $null }
                queueid      = if ($email.Attributes.Contains("queue.queueid")) { [guid]([AliasedValue]$email.Attributes["queue.queueid"]).Value }else { $null }
                start_processing = $integratedIncomingMailboxes | Where-Object -Property email -EQ -Value ([string]$email.Attributes["emailaddress"]) | select-object -first 1 -ExpandProperty start_processing
            }
            
            $mb = New-Object psobject -Property $outo
            $mailboxes += $mb
        }
        Write-Log -Message "Retrieved $($mailboxes.Count) mailboxes from Dataverse" -Level INFO
        Write-Output $mailboxes
    }
}

Function Get-EmailTrackingConfiguration {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]
        $service
    )
    begin {
        Write-Log -Message "Retrieving email tracking configuration from Dataverse" -Level INFO
        $fetch = @"
            <fetch>
                <entity name="organization">
                    <attribute name="emailcorrelationenabled" />
                    <attribute name="trackingprefix" />
                    <attribute name="emailconnectionchannel" />
                    <filter>
                        <condition attribute="emailconnectionchannel" operator="eq" value="0" />
                    </filter>
                </entity>
            </fetch>
"@
        $fetchEx = New-Object FetchExpression($fetch)
        $result = $service.RetrieveMultiple($fetchEx)

        
        if (($null -ne $result) -and ($null -ne $result.Entities) -and ($result.Entities.Count -eq 1)) {
            $organization = $result.Entities[0]
            #manage trackingtoken prefixes
            $prefixes = if ($null -ne $organization.Attributes["trackingprefix"]) { $organization.Attributes["trackingprefix"].Split(";") } else { $null }
            $filtered_prefixes = if ($null -ne $prefixes) { $prefixes | Where-Object { $true -ne [string]::IsNullOrEmpty($_) } } else { $null }
            $props = @{}
            $props["emailcorrelationenabled"] = $organization.Attributes["emailcorrelationenabled"]
            $props["trackingprefix"] = $filtered_prefixes
            $props["emailconnectionchannel"] = ([OptionSetValue]$organization.Attributes["emailconnectionchannel"]).Value
            $output_org = New-Object psobject -Property $props
            Write-Output $output_org
        }
        else {
            Write-Log -Message "Failed to retrieve email tracking configuration from Dataverse" -Level ERROR
            Write-Output $null
        }
    }
    process {}
    end {}
}

Function Get-MGAuthToken {
    param ()

    begin {
        <#
        $connectionDetails = @{ 
            'TenantId'     = '<tenant_id>' 
            'ClientId'     = '<app_id>' 
            'ClientSecret' = '<app_secret>' | ConvertTo-SecureString -AsPlainText -Force 
        }
        #>
        $config = Get-PSConfig
        $aad = $config.AzureAD
        if ($null -eq $aad) {
            throw "AzureAD configuration missing in config file."
        }

        $tenantID = $aad.TenantId
        $tokenBody = @{
            Grant_Type    = "client_credentials"
            Scope         = "https://graph.microsoft.com/.default"
            Client_Id     = $aad.ClientId
            Client_Secret = $aad.ClientSecret
         }
         $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $tokenBody
        Write-Log -Message "Retrieved authentication token for tenant: $tenantID" -Level INFO
        Write-Output $tokenResponse
    }
}

#Retrieves incoming messages
Function Get-ExchangeMessages {
    Param(
        [Parameter(Mandatory = $true)]
        [string]
        $upn,
        [Parameter(Mandatory = $true)]
        [string]
        $start,
        [Parameter(Mandatory = $true)]
        [psobject]
        $token
    )
    begin {
        try {
            Write-Log -Message "Retrieving incoming messages for upn: $upn" -Level INFO
            #Get Mail Folder
            $mailFolderUrl = "https://graph.microsoft.com/v1.0/users/$($upn)/mailFolders/Inbox"
            $headers = @{Authorization = ("Bearer " + $token.access_token) }
            $inboxFolder = Invoke-RestMethod -Method Get -Uri $mailFolderUrl -Headers $headers -ErrorAction Continue
            if ($null -eq $inboxFolder) {
                Write-Output @()
            }
            
            Connect-MgGraph -AccessToken ($token.access_token | ConvertTo-SecureString -AsPlainText -Force) -NoWelcome
            $messages = Get-MgUserMailFolderMessage -MailFolderId $inboxFolder.Id `
                -UserId $upn `
                -Filter "not(categories/any(c:c eq 'Tracked In Dataverse')) and receivedDateTime ge $start" `
                -Top 100 `
                -OrderBy "receivedDateTime DESC" `
                -Property "internetMessageHeaders", "body", "torecipients", "from", "ccrecipients", "subject", "internetmessageid", "importance", "hasattachments", "attachments", "categories" `
                -ExpandProperty Attachments
            $out_messages = @()
            foreach ($msg in $messages) {
                if($null -ne $msg -and $null -ne $msg.id) {
                    $props = @{}
                    $props["internalid"] = $msg.Id
                    $props["messageid"] = $msg.internetmessageid
                    $props["inreplyto"] = $msg.InternetMessageHeaders | Where-Object { $_.Name -ieq "In-Reply-To" } | Select-Object -Property Value -ExpandProperty Value
                    $props["subject"] = $msg.Subject
                    $props["body"] = $msg.Body.Content
                    $props["importance"] = $msg.Importance
                    $props["sender"] = $msg.From.EmailAddress.Address
                    $props["torecipients"] = $msg.ToRecipients | Where-Object { $null -ne $_.EmailAddress } | Select-Object -ExpandProperty EmailAddress | Select-Object -ExpandProperty Address
                    $props["ccrecipients"] = $msg.CcRecipients | Where-Object { $null -ne $_.EmailAddress } | Select-Object -ExpandProperty EmailAddress | Select-Object -ExpandProperty Address
                    $props["bccrecipients"] = $msg.BccRecipients | Where-Object { $null -ne $_.EmailAddress } | Select-Object -ExpandProperty EmailAddress | Select-Object -ExpandProperty Address
                    $props["attachments"] = if ($true -eq $msg.HasAttachments) { $msg.Attachments | Select-Object -Property Name, Size, ContentType, @{Name = "Filecontent"; Expression = { $_.AdditionalProperties.contentBytes } } } { @() }
                    $props["categories"] = $msg.categories
                    $message = New-Object psobject -Property $props
                    $out_messages += $message
                } 
            }
            Write-Log -Message "Retrieved $($out_messages.Count) messages for upn: $upn" -Level INFO
            Write-Output $out_messages
        }
        catch {
            Write-Output $null
        }
    }
    process {}
    end {}
}

Function Get-UserTrackingPreferences {
    Param(
        [Parameter(Mandatory = $true)]
        [guid]
        $userid,
        [Parameter(Mandatory = $true)]
        [psobject]
        $service
    )
    begin {
        $fetch = @"
            <fetch>
                <entity name="usersettings">
                    <attribute name="incomingemailfilteringmethod" />
                    <filter>
                    <condition attribute="systemuserid" operator="eq" value="$($userid)" />
                    </filter>
                </entity>
            </fetch>
"@
        $fetchEx = New-Object FetchExpression($fetch)
        $result = $service.RetrieveMultiple($fetchEx)
        
        $out_setting = [string]::Empty
        if ($null -ne $result -and $result.Entities.Count -eq 1) {
            $setting = $result.Entities | Select-Object -First 1
            $setting = ([OptionSetValue]$setting.Attributes["incomingemailfilteringmethod"]).Value
            $out_setting = [string]::Empty
            switch ($setting) {
                0 { $out_setting = "all" }
                1 { $out_setting = "correlate" }
                2 { $out_setting = "correlate" }
                3 { $out_setting = "correlate" }
                Default { $out_setting = "none" }
            }
        }
        Write-Output $out_setting
    }
}

#Find email activity in Dataverse by internetmessageid
Function Get-EmailByMessageId {
    param(
        [Parameter(Mandatory = $true)]
        [guid]
        $userid,
        [Parameter(Mandatory = $true)]
        [psobject]
        $service
    )
    begin {

    }
}

Function Add-IncomingEmailInDataverse {
    param(
        [Parameter(Mandatory = $false)]
        [psobject]
        $correlatedemail,
        [Parameter(Mandatory = $true)]
        [psobject]
        $newemail,
        [Parameter(Mandatory = $true)]
        [psobject]
        $service,
        [Parameter(Mandatory = $true)]
        [string]
        $upn
        
    )
    begin {
        
        write-Log -Message "Adding incoming email to Dataverse for upn: $upn with subject: $($newemail.subject)" -Level INFO
        $email = New-Object Entity("email")
        $created_emailid = $null
        $from_party = New-Object Entity("activityparty")
        $from_party.Attributes.Add("addressused", $newemail.sender)
        $from_party.Attributes.Add("participationtypemask", [OptionSetValue](New-Object OptionSetValue(1)))
        $resolvedFrom = Get-ResolvedAddress -service $service -email $newemail.sender | Sort-Object -Property order | Select-Object -First 1
        
        if ($null -ne $resolvedFrom) {
            $from_party.Attributes.Add("partyid", $resolvedFrom.ref)
        }

        $from_collection = New-Object List[Entity]
        $from_collection.Add($from_party)

        $to_collection = New-Object List[Entity]
        $cc_collection = New-Object List[Entity]
        $bcc_collection = New-Object List[Entity]

        #add to recipients
        foreach ($to in $newemail.torecipients) {
            $resolutions = Get-ResolvedAddress -email $to -service $service | Sort-Object -Property order -Unique
            foreach ($res in $resolutions) {
                $to_party = New-Object Entity("activityparty")
                $to_party.Attributes.Add("participationtypemask", [OptionSetValue](New-Object OptionSetValue(2)))
                $to_party.Attributes.Add("addressused", $to)
                $to_party.Attributes.Add("partyid", $res.ref)
                $to_collection.Add($to_party)
            }
        }
        #add cc recipients
        foreach ($cc in $newemail.ccrecipients) {
            $resolutions = Get-ResolvedAddress -email $cc -service $service
            foreach ($res in $resolutions) {
                $cc_party = New-Object Entity("activityparty")
                $cc_party.Attributes.Add("participationtypemask", [OptionSetValue](New-Object OptionSetValue(3)))
                $cc_party.Attributes.Add("addressused", $cc)
                $cc_party.Attributes.Add("partyid", [EntityReference]$res.ref)
                $cc_collection.Add($cc_party)
            }
        }

        #add bcc recipients
        foreach ($bcc in $newemail.bccrecipients) {
            $resolutions = Get-ResolvedAddress -email $bcc -service $service
            foreach ($res in $resolutions) {
                $bcc_party = New-Object Entity("activityparty")
                $bcc_party.Attributes.Add("participationtypemask", [OptionSetValue](New-Object OptionSetValue(4)))
                $bcc_party.Attributes.Add("addressused", $bcc)
                $bcc_party.Attributes.Add("partyid", $res.ref)
                $bcc_collection.Add($bcc_party)
            }
        }

        $email["from"] = [EntityCollection]::new($from_collection)
        $email["to"] = [EntityCollection]::new($to_collection)
        if ($cc_collection.Count -gt 0) {
            $email["cc"] = [EntityCollection]::new($cc_collection)
        }
        if ($bcc_collection.Count -gt 0) {
            $email["bcc"] = [EntityCollection]::new($bcc_collection)
        }
        $email["subject"] = $newemail.subject
        $email["description"] = $newemail.body
        if ($true -ne [string]::IsNullOrEmpty($newemail.inreplyto)) {
            Write-Log -Message "Trying to correlate email for upn: $upn with subject: $($newemail.subject)" -Level INFO
            #inreplyto is readonly, dont try to set it
            #$email["inreplyto"] = $newemail.inreplyto
            if($null -eq $correlatedemail){
                $correlatedemail = Get-CorrelatedEmail -service $service -inreplyto $newemail.inreplyto.Replace("<", "\u003C").Replace(">", "\u003E") -newEmail $newemail
            }
            if ($null -ne $correlatedemail) {
                Write-Log -Message "Found correlated email for upn: $upn $($correlatedemail.subject)" -Level INFO
                $correlatedactivityid = New-Object EntityReference -Property @{
                    LogicalName = "email"
                    Id          = $correlatedemail.id
                }
                #$email.Attributes.Add("correlatedactivityid", [EntityReference]$correlatedactivityid)
                $parentactivityid = New-Object EntityReference -Property @{
                    LogicalName = "email"
                    Id          = $correlatedemail.id
                }
                $email.Attributes.Add("parentactivityid", [EntityReference]$parentactivityid)
                if ($null -ne $correlatedemail.regardingobjectid) {
                    $email["regardingobjectid"] = [EntityReference]$correlatedemail.regardingobjectid
                }
            }
        }
        #$email["importance"] = $newemail.importance
        $email["messageid"] = $newemail.messageid
        $email["statuscode"] = [OptionSetValue](New-Object OptionSetValue(1))
        $email["statecode"] = [OptionSetValue](New-Object OptionSetValue(0))
        $emailid = $service.Create($email) 
        $created_emailid = $emailid
        #Manage Attachments if any
        if ($null -ne $newemail.attachments -and 0 -ne $newemail.attachments.Count -and $null -ne $emailid) {
            Write-Log -Message "Adding attachments to email for upn: $upn" -Level INFO
            foreach ($att in $newemail.attachments) {
                
                $bytes = [System.Convert]::FromBase64String($att.Filecontent)

                $attachment = New-Object Entity("activitymimeattachment");
                $attachment["objecttypecode"] = "email"
                $attachment["subject"] = $att.Name
                $attachment["filename"] = $att.Name
                $attachment["mimetype"] = $att.ContentType
                $attachment["body"] = [string]$att.Filecontent#[Convert]::ToBase64String($bytes)
                $emailRef = New-Object EntityReference -Property @{
                    LogicalName = "email"
                    Id          = $emailid
                }
                $attachment["objectid"] = [EntityReference]$emailRef
                $attachment["attachmentid"] = [EntityReference]$null

                $service.Create($attachment) | out-null
            }
        }

        $update_email = new-object Entity("email")
        $update_email["activityid"] = $emailid
        $update_email["statuscode"] = [OptionSetValue](New-Object OptionSetValue(4))
        $update_email["statecode"] = [OptionSetValue](New-Object OptionSetValue(1))
        $service.Update($update_email)

        Add-TrackedCategorie -msg $newemail -categorie "trackedingtrak" -upn $upn
        Write-Log -Message "Tracked email for upn: $upn with subject: $($newemail.subject)" -Level INFO
    }
    end {
        Write-Output $created_emailid
    }
}

Function Get-ResolvedAddress {
    Param(
        [Parameter(Mandatory = $true)]
        [string]
        $email,
        [Parameter(Mandatory = $true)]
        [psobject]
        $service
    )
    begin {
        $fetch_emailsearch = @"
        <fetch>
            <entity name="emailsearch">
                <attribute name="emailaddress" />
                <attribute name="parentobjectid" />
                <filter>
                    <condition attribute="emailaddress" operator="eq" value="$($email)" />
                </filter>
                <link-entity name="systemuser" from="systemuserid" to="parentobjectid" link-type="outer" alias="user">
                    <attribute name="systemuserid" />
                    <attribute name="fullname" />
                    <filter>
                        <condition attribute="isdisabled" operator="eq" value="0" />
                    </filter>
                </link-entity>
                <link-entity name="queue" from="queueid" to="parentobjectid" link-type="outer" alias="queue">
                    <attribute name="queueid" />
                    <attribute name="name" />
                    <filter>
                        <condition attribute="statecode" operator="eq" value="0" />
                    </filter>
                </link-entity>
                <link-entity name="contact" from="contactid" to="parentobjectid" link-type="outer" alias="contact">
                    <attribute name="contactid" />
                    <attribute name="fullname" />
                    <filter>
                        <condition attribute="statecode" operator="eq" value="0" />
                    </filter>
                </link-entity>
                <link-entity name="account" from="accountid" to="parentobjectid" link-type="outer" alias="account">
                    <attribute name="accountid" />
                    <attribute name="name" />
                    <filter>
                        <condition attribute="statecode" operator="eq" value="0" />
                    </filter>
                </link-entity>
            </entity>
        </fetch>
"@

        $fetchEx = New-Object FetchExpression($fetch_emailsearch)
        $result = $service.RetrieveMultiple($fetchEx)
        
        if ($result.Entities.Count -gt 0 -and $result.Entities.Count -lt 100) {
            $resolutions = @()
            foreach ($addressresolution in $result.Entities) {
                $ref = $null
                $props = @{}
                #is it user
                if ($addressresolution.Attributes.Contains("user.systemuserid")) {
                    $id = [Guid]([AliasedValue]$addressresolution.Attributes["user.systemuserid"]).Value
                    $name = [string]([AliasedValue]$addressresolution.Attributes["user.fullname"]).Value
                    $ref = New-Object EntityReference -Property @{
                        Id          = $id
                        LogicalName = "systemuser"
                        Name        = $name
                    }
                    
                    $props["order"] = 1
                    $props["ref"] = $ref
                    
                    $r = New-Object psobject -Property $props
                    $resolutions += $r
                }
                #is it queue
                elseif ($addressresolution.Attributes.Contains("queue.queueid")) {
                    $id = [Guid]([AliasedValue]$addressresolution.Attributes["queue.queueid"]).Value
                    $name = [string]([AliasedValue]$addressresolution.Attributes["queue.name"]).Value
                    $ref = New-Object EntityReference -Property @{
                        Id          = $id
                        LogicalName = "queue"
                        Name        = $name
                    }

                    $props["order"] = 2
                    $props["ref"] = $ref
                    
                    $r = New-Object psobject -Property $props
                    $resolutions += $r
                }
                #is it contact
                elseif ($addressresolution.Attributes.Contains("contact.contactid")) {
                    $id = [Guid]([AliasedValue]$addressresolution.Attributes["contact.contactid"]).Value
                    $name = [string]([AliasedValue]$addressresolution.Attributes["contact.fullname"]).Value
                    $ref = New-Object EntityReference -Property @{
                        Id          = $id
                        LogicalName = "contact"
                        Name        = $name
                    }
                    $props["order"] = 3
                    $props["ref"] = $ref
                    
                    $r = New-Object psobject -Property $props
                    $resolutions += $r
                }
                #is it account
                elseif ($addressresolution.Attributes.Contains("account.accountid")) {
                    $id = [Guid]([AliasedValue]$addressresolution.Attributes["account.accountid"]).Value
                    $name = [string]([AliasedValue]$addressresolution.Attributes["account.name"]).Value
                    $ref = New-Object EntityReference -Property @{
                        Id          = $id
                        LogicalName = "account"
                        Name        = $name
                    }

                    $props["order"] = 4
                    $props["ref"] = $ref
                    
                    $r = New-Object psobject -Property $props
                    $resolutions += $r
                }
            }
            
            Write-Output ($resolutions | Sort-Object -Property order) 
        }
        else {
            Write-Output @()
        }
    }
}

#Checks in dataverse if the 
Function Get-CorrelatedEmail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $inreplyto,
        [psobject]
        $newEmail,
        [psobject]
        $service
    )

    begin {

        <#
            $props["emailcorrelationenabled"] = $organization.Attributes["emailcorrelationenabled"]
            $props["trackingprefix"] = $filtered_prefixes
            $props["emailconnectionchannel"] = ([OptionSetValue]$organization.Attributes["emailconnectionchannel"]).Value
            
        #>

        $orgsettings = Get-EmailTrackingConfiguration -service $service
        if ($null -eq $orgsettings -or $false -eq $orgsettings.emailcorrelationenabled -or $null -eq $newEmail.subject -or $false -eq $newEmail.subject.Contains($orgsettings.trackingprefix)) {
            Write-Output $null
        }
        else {
            $trackingNumber = ([string]$newEmail.subject).Split($orgsettings.trackingprefix) | Select-Object -Last 1
            $trackingtoken = ""
            if ($null -ne $trackingNumber) {

                $trackingtoken = "$($orgsettings.trackingprefix)$($trackingNumber)"
                $trackingtokenfilter = "<condition attribute=`"subject`" operator=`"ends-with`" value=`"$($trackingtoken)`" />"
            }
            #1 - check the inreplyto field
            $fetch = @"
            <fetch top="1">
                <entity name="email">
                    <attribute name="inreplyto" />
                    <attribute name="subject" />
                    <attribute name="activityid" />
                    <attribute name="regardingobjectid" />
                    <filter type="or">
                        <condition attribute="inreplyto" operator="eq" value="$($inreplyto)" />
                        $trackingtokenfilter
                    </filter>
                </entity>
            </fetch>
"@
            $fetchEx = New-Object FetchExpression($fetch)
            $emails = $service.RetrieveMultiple($fetchEx)
            
            if ($null -ne $emails -and $emails.Entities.Count -eq 1) {
                $cemail = $emails.Entities | Select-Object -First 1
                $props = @{
                    "id"                = [guid]$cemail.Attributes["activityid"]
                    "regardingobjectid" = [EntityReference]$cemail.Attributes["regardingobjectid"]
                }
                $outemail = New-Object psobject -Property $props
                Write-Output $outemail
            }
            else {
                Write-Output $null
            }
        }
    }
}

Function Get-DataverseOutgoingMailbox {
    param(
        [Parameter(Mandatory = $true)]
        [psobject] $service
    )
    begin {
        Write-Log -Message "Retrieving outgoing mailboxes from Dataverse" -Level INFO
        $outgoingMailboxConfigs = Get-OutMailboxConfig
        $outmailbox_query = $outgoingMailboxConfigs | Select-Object -ExpandProperty email |foreach-object { "<value>$($_)</value>" } | Out-String
        $fetch = @"
            <fetch>
                <entity name="mailbox">
                    <attribute name="emailaddress" />
                    <attribute name="mailboxid" />
                    <filter>
                        <condition attribute="outgoingemaildeliverymethod" operator="eq" value="2" />
                        <condition attribute="statecode" operator="eq" value="0" />
                        <condition attribute="emailaddress" operator="in">
                            $outmailbox_query
                        </condition>
                    </filter>
                    <link-entity name="queue" from="queueid" to="regardingobjectid" link-type="outer" alias="queue">
                        <attribute name="queueid" />
                    </link-entity>
                    <link-entity name="systemuser" from="systemuserid" to="regardingobjectid" link-type="outer" alias="user">
                        <attribute name="systemuserid" />
                    </link-entity>
                </entity>
            </fetch>
"@
        $fetchEx = New-Object FetchExpression($fetch)
        $result = $service.RetrieveMultiple($fetchEx)
        $mailboxes = @()
        Write-Log -Message "Retrieved $($result.Entities.Count) mailboxes from Dataverse" -Level INFO
        foreach ($mb in $result.Entities) {
            if($null -ne $mb.Attributes["queue.queueid"] -or $null -ne $mb.Attributes["user.systemuserid"]) {
                $outo = @{
                    userid       = if ($mb.Attributes.Contains("user.systemuserid")) { [guid]([AliasedValue]$mb.Attributes["user.systemuserid"]).Value }else { $null }
                    queueid      = if ($mb.Attributes.Contains("queue.queueid")) { [guid]([AliasedValue]$mb.Attributes["queue.queueid"]).Value }else { $null }  
                    emailaddress = [string]$mb.Attributes["emailaddress"]
                    mailboxid    = [guid]$mb.Id
                    start_processing = $outgoingMailboxConfigs | Where-Object -Property email -EQ -Value ([string]$mb.Attributes["emailaddress"]) | select-object -first 1 -ExpandProperty start_processing
                }
            
                $mb = New-Object psobject -Property $outo
                $mailboxes += $mb   
            }
        }

        Write-Output $mailboxes
    }
}

Function Get-DataverseOutgoingEmail {
    param(
        [Parameter(Mandatory = $true)]
        [psobject] $service,
        [Parameter(Mandatory = $true)]
        [psobject] $mb
    )
    begin {
        Write-Log -Message "Retrieving outgoing emails from Dataverse for mailbox: $($mb.emailaddress)" -Level INFO
        $fetch = @"
            <fetch top="100">
                <entity name="email">
                    <attribute name="inreplyto" />
                    <attribute name="subject" />
                    <attribute name="activityid" />
                    <attribute name="description" />
                    <attribute name="regardingobjectid" />
                    <attribute name="messageid" />
                    <attribute name="sender" />
                    <attribute name="torecipients" />
                    <attribute name="activityid" />
                    <attribute name="modifiedon" />
                    <order attribute="modifiedon" descending="true" />
                    <filter type="and">
                        <condition attribute="statuscode" operator="eq" value="6" />
                        <condition attribute="directioncode" operator="eq" value="true" />
                        <condition attribute="modifiedon" operator="ge" value="$($mb.start_processing)" />
                        <condition attribute="sendermailboxid" operator="eq" value="$($mb.mailboxid)" />
                    </filter>
                    <link-entity name="activityparty" from="activityid" to="activityid" alias="cc" link-type="outer">
                        <filter>
                            <condition attribute="participationtypemask" operator="eq" value="3" />
                            <condition attribute="addressused" operator="not-null"/>
                        </filter>
                        <attribute name="addressused" />
                    </link-entity>
                    <link-entity name="activityparty" from="activityid" to="activityid" alias="to" link-type="outer">
                        <filter>
                            <condition attribute="participationtypemask" operator="eq" value="2" />
                            <condition attribute="addressused" operator="not-null"/>
                        </filter>
                        <attribute name="addressused" />
                    </link-entity>

                    <link-entity name="activityparty" from="activityid" to="activityid" alias="bcc" link-type="outer">
                        <filter>
                            <condition attribute="participationtypemask" operator="eq" value="4" />
                            <condition attribute="addressused" operator="not-null"/>
                        </filter>
                        <attribute name="addressused" />
                    </link-entity>
                    <link-entity name="activitymimeattachment" from="objectid" to="activityid" link-type="outer" alias="mimeattachments">
                        <attribute name="attachmentid" />
                        <link-entity name="attachment" from="attachmentid" to="attachmentid" link-type="outer" alias="attachments">
                            <attribute name="filename" />
                            <attribute name="body" />
                            <attribute name="mimetype" />
                        </link-entity>
                    </link-entity>
                </entity>
            </fetch>
"@
        $fetchEx = New-Object FetchExpression($fetch)
        $emails = $service.RetrieveMultiple($fetchEx)
      
        if ($null -ne $emails -and $null -ne $emails.Entities -and 0 -lt $emails.Entities.Count) {
            $messageids = $emails.Entities | Sort-Object -Property Id -Unique | Select-Object -ExpandProperty Id
            $outmsgs = @()
            foreach ($id in $messageids) {
                $e = $emails.Entities | Where-Object { $_.Id -eq $id }
                $fe = $e | Select-Object -First 1

                $tos = @()
                foreach ($to in $e) {
                    if ($null -ne $to["to.addressused"]) {
                        $recipient = [string]([AliasedValue]$to["to.addressused"]).Value
                        $tos += @{ EmailAddress = @{ Address = $recipient } }
                    }
                }

                $ccs = @()
                foreach ($cce in $e) {
                    if ($null -ne $cce["cc.addressused"]) {
                        $recipient = [string]([AliasedValue]$cce["cc.addressused"]).Value
                        $ccs += @{ EmailAddress = @{ Address = $recipient } }
                    }
                }
                $bccs = @()
                foreach ($bcce in $e) {
                    if ($null -ne $bcce["bcc.addressused"]) {
                        $recipient = [string]([AliasedValue]$bcce["bcc.addressused"]).Value
                        $bccs += @{ EmailAddress = @{ Address = $recipient } }
                    }
                }
                $atts = @()
                foreach ($att in $e) {
                    if ($null -ne $att["attachments.mimetype"]) {
                        $atts += @{
                            "@odata.type" = "#microsoft.graph.fileAttachment"
                            name          = [string]([AliasedValue]$att["attachments.filename"]).Value
                            contentType   = [string]([AliasedValue]$att["attachments.mimetype"]).Value
                            contentBytes  = [string]([AliasedValue]$att["attachments.body"]).Value
                        }
                    }
                }

                $props = @{
                    "activityid"    = [guid]$fe.Attributes["activityid"]
                    "modifiedon"    = $fe["modifiedon"]
                    "messageid"     = $fe["messageid"]
                    "body"          = $fe["description"]
                    "subject"       = $fe["subject"]
                    "to"            = $tos
                    "from"          = $fe["sender"]
                    "cc"            = $ccs
                    "bcc"           = $bccs
                    "attachments" = $atts
                }
                $outmsgs += (New-Object psobject -Property $props)
            }
            Write-Log -Message "Retrieved $($outmsgs.Count) outgoing emails for mailbox: $($mb.emailaddress)" -Level INFO
            Write-Output $outmsgs
        }
        else {
            Write-Log -Message "No outgoing emails found for mailbox: $($mb.emailaddress)" -Level INFO
            Write-Output $null
        }

    }
}

Function Send-Email {
    param(
        [psobject] $message,
        [psobject] $mailbox,
        [psobject] $service,
        [psobject] $token
    )
    begin {
        <#
        $URLMail = "https://graph.microsoft.com/v1.0/users/$MailFrom/messages"
         $BodyJsonsend = @{
            "Message" = @{
                "Subject" = $message.subject
                "Body" = @{
                    "ContentType" = "Text"
                    "Content" = $message.body
                }
                "ToRecipients" = $message.torecipients | ForEach-Object { @{ EmailAddress = @{ Address = $_ } } }
                "CcRecipients" = $message.ccrecipients | ForEach-Object { @{ EmailAddress = @{ Address = $_ } } }
                "BccRecipients" = $message.bccrecipients | ForEach-Object { @{ EmailAddress = @{ Address = $_ } } }
                "Attachments" = $message.attachments
            }
        } | ConvertTo-Json -Depth 10

        $createdmessage = Invoke-RestMethod -Method POST -Uri $URLMail -Headers $headers -Body $BodyJsonsend
            if($null -ne $createdmessage -and $null -ne $createdmessage.id) {
                foreach($att in $message.attachments) {
                    $attachmentsBody = $att | ConvertTo-Json -Depth 10
                    $URLAttachment = "https://graph.microsoft.com/v1.0/users/$MailFrom/messages/$($createdmessage.id)/attachments"
                    Invoke-RestMethod -Method POST -Uri $URLAttachment -Headers $headers -Body $attachmentsBody
                }
                $urlSend = "https://graph.microsoft.com/v1.0/users/$MailFrom/messages/$($createdmessage.id)/Send"
                Invoke-RestMethod -Method POST -Uri $urlSend -Headers $headers
            }
        #>
        
        
        try {
            Write-Log -Message "Attempting to send email for mailbox: $($mailbox.emailaddress)" -Level INFO
            
            Connect-MgGraph -AccessToken ($token.access_token | ConvertTo-SecureString -AsPlainText -Force) -NoWelcome
            Send-MgUserMail -ErrorAction Stop -UserId ($mailbox.emailaddress) -BodyParameter @{
                Message = @{
                    Subject = $message.subject
                    Body = @{ ContentType = "HTML"; Content = $message.body }
                    ToRecipients = $message.to
                    CcRecipients = $message.cc
                    BccRecipients = $message.bcc
                    Attachments = $message.attachments
                }
            }
            #If all went wel, update the email in dataverse to sent
            $update_email = new-object Entity("email")
            $update_email["activityid"] = $message.activityid
            $update_email["statuscode"] = [OptionSetValue](New-Object OptionSetValue(3))
            $update_email["statecode"] = [OptionSetValue](New-Object OptionSetValue(1))
            $service.Update($update_email)

        } catch{
            Write-Log -Message "Failed to send email for mailbox: $($mailbox.emailaddress)" -Level ERROR
            write-Log -Message $_.Exception.Response.StatusDescription -Level ERROR
            $update_email = new-object Entity("email")
            $update_email["activityid"] = $message.activityid
            $update_email["statuscode"] = [OptionSetValue](New-Object OptionSetValue(5))
            $update_email["statecode"] = [OptionSetValue](New-Object OptionSetValue(1))
            $service.Update($update_email)
        }
    }
}

Function Add-QueueItem {
    param(
        [Parameter(Mandatory = $true)]
        [psobject] $service,
        [Parameter(Mandatory = $true)]
        [guid] $queueid,
        [Parameter(Mandatory = $true)]
        [guid] $objectid,
        [Parameter(Mandatory = $true)]
        [string] $objecttype
    )
    begin {
        Write-Log -Message "Adding queue item for object: $objectid" -Level INFO
        $queueitem = New-Object Entity("queueitem")
        $queueRef = New-Object EntityReference -Property @{
            LogicalName = "queue"
            Id          = $queueid
        }
        $queueitem["queueid"] = [EntityReference]$queueRef
        $objectRef = New-Object EntityReference -Property @{
            LogicalName = $objecttype
            Id          = $objectid
        }
        $queueitem["objectid"] = [EntityReference]$objectRef
        $queueitem["statecode"] = [OptionSetValue](New-Object OptionSetValue(0))
        $queueitem["statuscode"] = [OptionSetValue](New-Object OptionSetValue(1))
        $service.Create($queueitem)
    }
}

Function Add-TrackedCategorie {
    param(
        [Parameter(Mandatory = $true)]
        [string] $upn,
        [Parameter(Mandatory = $true)]
        [string] $categorie,
        [Parameter(Mandatory = $true)]
        [psobject] $msg
    )    
    begin {
        Write-Log -Message "Adding tracked in dataverse category message id: $($msg.messageid)" -Level INFO
        Update-MgUserMessage -UserId $upn -MessageId $msg.internalid -Categories "Tracked In Dataverse"
    }
}

Export-ModuleMember *
