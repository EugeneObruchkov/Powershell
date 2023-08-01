$query = "SELECT * FROM [dbo].[tbl_UsersForRemoval] WHERE RemovalDate = CAST(GETDATE() AS DATE)"

$UsersForRemove = Invoke-sqlcmd -ServerInstance 'MSSQL\MSSQL_M_DB' -Database 'ITR1' -Query $query


if ($UsersForRemove -ne $null)
{
    foreach ($userForRemove in $UsersForRemove)
    {
        $SamAccountName = $userForRemove.SamAccountName
        $user = Get-ADUser -Filter {SamAccountName -eq $SamAccountName} -Properties extensionAttribute4, Company -ErrorAction SilentlyContinue
        
        if ($user -ne $null)
        {
            $userForRemoveTicketLink = $userForRemove.TicketLink
            $extensionAttribute4 = $user.extensionAttribute4
            $Company = $user.Company
            $subject = "Изтриване на потребителски акаунт: $SamAccountName"
            $Message = "
            <!DOCTYPE html>
            <html>
            <body>
            <font size=2 face=verdana color=#003366>
            Нотификация за изтриване на потребителски акаунт на напуснал служител.<br><br>
            SamAccountName: $SamAccountName<br>
            Име: $extensionAttribute4<br>
            Фирма: $Company<br>
            Заявка за освобождаване: $userForRemoveTicketLink<br>
            </font>
            </body>
            </html>"
            
            Send-MailMessage -To "support@testcompany.com" -From "Do Not Reply <noreply@testcompany.com>" -SmtpServer EX1901.testcompany.com -Subject "$subject" -Body "$Message" -BodyAsHtml -Encoding UTF8
        }     
    }
}

