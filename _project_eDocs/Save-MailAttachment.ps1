# Name of the mailbox to pull attachments from
$MailboxName = 'edocs@testcompany.com'

# Path to the XML configuration file
[XML]$configFile = Get-Content ".\config.xml"

# Location to move attachments
$downloadDirectory = '\\testcompany.com\eDocs'
 
# Path to the Web Services dll
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
[VOID][Reflection.Assembly]::LoadFile($dllpath)
 
# Create the new web services object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService 'Exchange2013_SP1',([timezoneinfo]::Utc)
 
# Create the LDAP security string in order to log into the mailbox
#$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
#$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
#$aceuser = [ADSI]$sidbind
 
# Auto discover the URL used to pull the attachments
$service.AutodiscoverUrl($MailboxName)
 
# Get the folder id of the Inbox
$FolderID = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)
$InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
$DeletedItemsFolderID = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems,$MailboxName)

# Find mail in the Inbox with attachments
$Sfha = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
$sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
$sfCollection.add($Sfha)
 
# Grab all the mail that meets the prerequisites
$view = new-object Microsoft.Exchange.WebServices.Data.ItemView(2000)
$frFolderResult = $InboxFolder.FindItems($sfCollection,$view)
 
# Loop through the emails
foreach ($MailMessage in $frFolderResult.Items)
{ 
    # Load the message
    $MailMessage.Load()

    # Proceed only if sender exists in config file
    $MailSender = $MailMessage.Sender.Address
    foreach ($xmlConfigItem in $configFile.configuration.add)
    {
        if ($xmlConfigItem.Email -eq "$MailSender")
        {
            # Get attachment type based on XML config file
            $MailAttachmentType = $xmlConfigItem.AttachmentType
            switch ($MailAttachmentType)
            {
                "RoadSheet"
                {
                    # Loop through the attachments
                    foreach($MailAttachment in $MailMessage.Attachments)
                    {
                        $MailAttachmentName = $MailAttachment.Name
                        
                        # Process PDF files only
                        if ($MailAttachmentName -match "(.*?)\.(pdf|PDF)$")
                        {
                            # Load the attachment
                            $MailAttachment.Load()
                            
                            $MailAttachment.Name.ToString() -match '[A-Z]{2}\d{4}[A-Z]{2}'
                            $CarID = $Matches[0]
                            $Timestamp = Get-Date -Format o | ForEach-Object { $_ -replace “:”,“.” -replace "-","." -replace "T","_" -replace "\+.*",""}
                            $AttachmentFileName = "RoadSheet" + "-" + "$CarID" + "-" + "$Timestamp" + ".pdf"

                            # Save the attachment to the predefined location
                            $MailAttachmentFile = new-object System.IO.FileStream(($downloadDirectory + “\” + $AttachmentFileName), [System.IO.FileMode]::Create)
                            $MailAttachmentFile.Write($MailAttachment.Content, 0, $MailAttachment.Content.Length)
                            $MailAttachmentFile.Close()
                        }
                    }
                    # Mark the email as read
                    $MailMessage.isread = $true
                    $MailMessage.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
 
                    # Move the message to Archive folder
                    $MailMessage.Move($DeletedItemsFolderID) | Out-Null
                }
                "BankStatement"
                {
                    # Loop through the attachments
                    foreach($MailAttachment in $MailMessage.Attachments)
                    {
                        $MailAttachmentName = $MailAttachment.Name
                        
                        # Process PDF files only
                        if ($MailAttachmentName -match "(.*?)\.(pdf|PDF)$")
                        {
                            # Load the attachment
                            $MailAttachment.Load()
                            
                            $MailMessage.Subject.ToString() -match '\b[A-Z]{2}\d{2}[A-Z]{4}\d{14}\b'
                            $IBAN = $Matches[0]
                            $Timestamp = Get-Date -Format o | ForEach-Object { $_ -replace “:”,“.” -replace "-","." -replace "T","_" -replace "\+.*",""}
                            $AttachmentFileName = "BankStatement" + "-" + "$IBAN" + "-" + "$Timestamp" + ".pdf"

                            # Save the attachment to the predefined location
                            $MailAttachmentFile = new-object System.IO.FileStream(($downloadDirectory + “\” + $AttachmentFileName), [System.IO.FileMode]::Create)
                            $MailAttachmentFile.Write($MailAttachment.Content, 0, $MailAttachment.Content.Length)
                            $MailAttachmentFile.Close()
                        }
                    }
                    # Mark the email as read
                    $MailMessage.isread = $true
                    $MailMessage.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
 
                    # Move the message to Archive folder
                    $MailMessage.Move($DeletedItemsFolderID) | Out-Null
                }
                "CCStatement"
                {
                    # Loop through the attachments
                    foreach($MailAttachment in $MailMessage.Attachments)
                    {
                        $MailAttachmentName = $MailAttachment.Name
                        
                        # Process PDF files only
                        if ($MailAttachmentName -match "(.*?)\.(pdf|PDF)$")
                        {
                            # Load the attachment
                            $MailAttachment.Load()
                            
                            $IBAN = $xmlConfigItem.IBAN
                            $Timestamp = Get-Date -Format o | ForEach-Object { $_ -replace “:”,“.” -replace "-","." -replace "T","_" -replace "\+.*",""}
                            $AttachmentFileName = "CCStatement" + "-" + "$IBAN" + "-" + "$Timestamp" + ".pdf"

                            # Save the attachment to the predefined location
                            $MailAttachmentFile = new-object System.IO.FileStream(($downloadDirectory + “\” + $AttachmentFileName), [System.IO.FileMode]::Create)
                            $MailAttachmentFile.Write($MailAttachment.Content, 0, $MailAttachment.Content.Length)
                            $MailAttachmentFile.Close()
                        }
                    }
                    # Mark the email as read
                    $MailMessage.isread = $true
                    $MailMessage.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
 
                    # Move the message to Archive folder
                    $MailMessage.Move($DeletedItemsFolderID) | Out-Null
                }
                default
                {
                    Write-Warning "There is no configured $MailAttachmentType attachment type for $MailSender in the XML configuration file"
                }
            }
        }
    }
}