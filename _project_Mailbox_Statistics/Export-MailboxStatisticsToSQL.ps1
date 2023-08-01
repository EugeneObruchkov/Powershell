# Import sql server module
Import-Module SQLserver

# PSRemoting to Microsoft Exchange server
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://EX1901.testcompany.com/PowerShell/ -Authentication Kerberos
Import-PSSession $Session

# define date
$date = Get-Date -UFormat "%Y-%m-%d %H:%M:%S"

# define SQL server parameters
$params = @{'server'='MSSQL\MSSQL_M_DB';'Database'='ITR1'}
$sqlServer='MSSQL'

# define functions
function Write-MailboxInfoToSQL
{
param($server,$CN_UserName,$UserNameBG,$CompanyENG,$CompanyBG,$Email,$MBoxSize,$ArchMBoxSize,$TotalSize,$AccountType,$Enabled,$QueryDate)
    $InsertResults = @"
        INSERT INTO [ITR1].[dbo].[tbl_EmailStatistics](CN_UserName,UserNameBG,CompanyENG,CompanyBG,Email,MBoxSize,ArchMBoxSize,MailTotalSize,AccountType,Enabled,QueryDate)
        VALUES ('$CN_UserName','$UserNameBG','$CompanyENG','$CompanyBG','$Email','$MBoxSize','$ArchMBoxSize','$TotalSize','$AccountType','$Enabled','$QueryDate')
"@      
    Invoke-sqlcmd @params -Query $InsertResults
}

# get all mailboxes
$mailboxes = Get-Mailbox -server EX1901

foreach ($mailbox in $mailboxes) 
{
    $UserSamAccountName = $mailbox.SamAccountName
    if ($UserSamAccountName -ne $null) 
    {
        # get all mailbox properties
        $user = Get-Mailbox -Identity $UserSamAccountName | Select-Object `
        UserPrincipalName, `
        CustomAttribute4, `
        DisplayName, `
        CustomAttribute3, `
        PrimarySmtpAddress, `
        RecipientType, `
        IsMailboxEnabled, `
        Database, `
        ArchiveDatabase

        $userCompanyENG = Get-ADUser -Identity $UserSamAccountName -Properties Company
        $UserDatabase = $user.Database
        $UserArchiveDatabase = $user.ArchiveDatabase
        
        # get mailbox size in bytes
        if ($UserDatabase -ne $null)
        {
            $userMailboxTotalItemSize = (Get-MailboxStatistics -Identity "TESTCOMPANY\$UserSamAccountName").TotalItemSize -replace ".*?\(", "" -replace ",", "" -replace " bytes\)",""
            $userMailboxTotalItemSize = [int64]$userMailboxTotalItemSize
        }
        else 
        {
            $userMailboxTotalItemSize = $null
        }

        # get archive mailbox size in bytes
        if ($UserArchiveDatabase -ne $null)
        {
            $userArchiveMailboxTotalItemSize = (Get-MailboxStatistics -Identity "TESTCOMPANY\$UserSamAccountName" -Archive).TotalItemSize -replace ".*?\(", "" -replace ",", "" -replace " bytes\)",""
            $userArchiveMailboxTotalItemSize = [int64]$userArchiveMailboxTotalItemSize
        }
        else
        {
            $userArchiveMailboxTotalItemSize = $null
        }

        # calculate total (mailbox + archive) size
        $userTotalItemSize = $userMailboxTotalItemSize + $userArchiveMailboxTotalItemSize

        

        # write info to SQL server
        Write-MailboxInfoToSQL $sqlServer $UserSamAccountName $user.CustomAttribute4 $userCompanyENG.Company $user.CustomAttribute3 $user.PrimarySmtpAddress $userMailboxTotalItemSize $userArchiveMailboxTotalItemSize $userTotalItemSize $user.RecipientType $user.IsMailboxEnabled $date

        




    }
}

Remove-PSSession $Session