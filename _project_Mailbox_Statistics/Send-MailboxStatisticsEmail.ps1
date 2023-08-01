# Define functions
function Test-UserIsAMemberOfGroup
{
    param ($Identity, $GroupName)

    Begin
    {
        $UserInGroup = $false
        $GroupMembers = Get-ADGroupMember -Identity $GroupName
        $User = Get-ADUser -Identity $Identity
    }

    Process
    {
        foreach ($GroupMember in $GroupMembers)
        {
            $GroupMemberName = $GroupMember.SamAccountName
            $UserSamAccountName = $User.SamAccountName
            if ($GroupMemberName -eq $UserSamAccountName)
            {
                $UserInGroup = $true
            }
        }
    }

    End
    {
        $UserInGroup
    }
}


# PSRemoting to Microsoft Exchange server
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mail.testcompany.com/PowerShell/ -Authentication Kerberos
Import-PSSession $Session

# This hash table contains all mailbox databases from Microsoft Exhange server with their default mailbox quotas
$MailboxDatabasesHashTable = @{}
$MailboxDatabases = Get-MailboxDatabase
foreach ($MailboxDatabase in $MailboxDatabases) {
    $MailboxDatabaseName = $MailboxDatabase.Name
    $MailboxDatabaseQuota = $MailboxDatabase.ProhibitSendReceiveQuota
    $MailboxDatabasesHashTable.Add("$MailboxDatabaseName","$MailboxDatabaseQuota")
}

# get all mailboxes
$mailboxes = Get-Mailbox -server EX1901


foreach ($mailbox in $mailboxes) {

    $UserPrincipalName = $mailbox.UserPrincipalName

        if ($UserPrincipalName -ne $null) {
    
            $user = Get-Mailbox -Identity $UserPrincipalName | select SamAccountName, CustomAttribute4, DisplayName, PrimarySmtpAddress, MaxSendSize, MaxReceiveSize, ProhibitSendReceiveQuota, ArchiveQuota, ArchiveDatabase, UseDatabaseQuotaDefaults, Database
            $MailboxStatistics = Get-MailboxStatistics -Identity $UserPrincipalName | select TotalItemSize
            $userDatabaseQuotaDefaults = $user.UseDatabaseQuotaDefaults
            $globalSendReceiveSize = Get-TransportConfig
            $globalMaxReceiveSize = $globalSendReceiveSize.MaxReceiveSize
            $globalMaxSendSize = $globalSendReceiveSize.MaxSendSize
            $userSamAccountName = $user.SamAccountName
            $usernameEN = $user.DisplayName
            $usernameBG = $user.CustomAttribute4
                if ($username -eq $null) {
                    $username = $user.DisplayName
                }
            $usermail = $user.PrimarySmtpAddress
            $maxsendsize = $user.MaxSendSize
            if ($maxsendsize -eq "unlimited") {
                $maxsendsize = $globalMaxSendSize
            }
            $maxreceivesize = $user.MaxReceiveSize
            if ($maxreceivesize -eq "unlimited") {
                $maxreceivesize = $globalMaxReceiveSize
            }
            if ($userDatabaseQuotaDefaults -eq $True)
            {
                $userDatabase = $user.Database
                $ProhibitSendReceiveQuota = $MailboxDatabasesHashTable.Item("$userDatabase")
            }
            else 
            {
                $ProhibitSendReceiveQuota = $user.ProhibitSendReceiveQuota
            }
            $MailboxTotalItemSize = $MailboxStatistics.TotalItemSize
            
            $UserArchiveDatabase = $user.ArchiveDatabase
            if ($UserArchiveDatabase -eq $null)
            {
                $ArchiveQuota = "no archive"
                $ArchiveMailboxTotalItemSize = "no archive"
            }
            else 
            {
            $ArchiveMailboxStatistics = Get-MailboxStatistics -Identity $UserPrincipalName -Archive | select TotalItemSize
            $ArchiveQuota = $user.ArchiveQuota
            $ArchiveMailboxTotalItemSize = $ArchiveMailboxStatistics.TotalItemSize
            }

        }

$SubjectBG = "Лимити на съобщенията през ел.поща"
$SubjectEN = "Limits on e-mail messages"

$MessageBG = "
<!DOCTYPE html>
<html>
<body>
<font size=2 face=verdana color=#003366>
Здравейте г-н/г-жа $usernameBG,<br><br>  
Изпращаме Ви информация за лимитите на съобщенията, които получавате и изпращате през корпоративната електронна поща $usermail`:<br><br>      
<pre>
Максималният размер на входящото съобщение:&#9;$maxreceivesize
Максималният размер на изходящото съобщение:&#9;$maxsendsize<br>
Общ лимит на пощенската кутия:&#9;&#9;&#9;$ProhibitSendReceiveQuota
Текущ размер на пощенската кутия:&#9;&#9;$MailboxTotalItemSize<br>
Общ лимит на архивната кутия:&#9;&#9;&#9;$ArchiveQuota
Текущ размер на архивната кутия:&#9;&#9;$ArchiveMailboxTotalItemSize<br><br>
</pre>
При необходимост общият лимит на пощенската кутия може да бъде увеличен.<br><br>
С Уважение,<br>
Екипът на ИТР Сървисиз ЕООД<br><br><br>
<i><small>Забележка: Това е автоматично генерирано съобщение. Моля не отговаряйте.</i></small>
</font>
</body>
</html>"


$MessageEN = "
<!DOCTYPE html>
<html>
<body>
<font size=2 face=verdana color=#003366>
Hello, Mr/Mrs $usernameEN,<br><br>
We are sending you information about the limits on messages you receive and send through the corporate e-mail $usermail`:<br><br>
<pre>
Maximum size of an incoming message:&#9;$maxreceivesize
Maximum size of an outgoing message:&#9;$maxsendsize<br>
Total mailbox size limit:&#9;&#9;$ProhibitSendReceiveQuota
Current mailbox size:&#9;&#9;&#9;$MailboxTotalItemSize<br>
Total archive mailbox size limit:&#9;$ArchiveQuota
Current archive mailbox size:&#9;&#9;$ArchiveMailboxTotalItemSize<br><br>
</pre>
Your mailbox size can be increased if needed.<br><br>
Yours sincerely,<br>
The team of ITR Services EOOD<br><br><br>
<i><small>Note: This is an automated message. Please do not reply.</i></small>
</font>
</body>
</html>"


if (Test-UserIsAMemberOfGroup -Identity $userSamAccountName -GroupName "Mail Notifications - English")
{
    # Send email message in English
    Send-MailMessage -to "$usermail" -From "Do Not Reply <noreply@testcompany.com>" -SmtpServer mail.testcompany.com -Subject "$SubjectEN" -Body "$MessageEN" -BodyAsHtml -Encoding UTF8
}
else
{        
    # Send email message in Bulgarian
    Send-MailMessage -to "$usermail" -From "Do Not Reply <noreply@testcompany.com>" -SmtpServer mail.testcompany.com -Subject "$SubjectBG" -Body "$MessageBG" -BodyAsHtml -Encoding UTF8     
}
}

# Close PSRemoting session to MS Exchange server
Remove-PSSession $Session