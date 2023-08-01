<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '1100,800'
$Form.text                       = "Напускане на служител"
$Form.TopMost                    = $false

###############################################################################################

$Users_Groupbox                  = New-Object system.Windows.Forms.Groupbox
$Users_Groupbox.height           = 140
$Users_Groupbox.width            = 520
$Users_Groupbox.text             = "Информация за потребителите"
$Users_Groupbox.location         = New-Object System.Drawing.Point(20,21)

$fired_username_label            = New-Object system.Windows.Forms.Label
$fired_username_label.text       = "Напускащ служител (UserName):"
$fired_username_label.AutoSize   = $true
$fired_username_label.width      = 25
$fired_username_label.height     = 10
$fired_username_label.location   = New-Object System.Drawing.Point(30,45)
$fired_username_label.Font       = 'Microsoft Sans Serif,10'

$fired_username_textbox          = New-Object system.Windows.Forms.TextBox
$fired_username_textbox.multiline  = $false
$fired_username_textbox.width    = 200
$fired_username_textbox.height   = 20
$fired_username_textbox.location  = New-Object System.Drawing.Point(30,65)
$fired_username_textbox.Font     = 'Microsoft Sans Serif,10'

$fired_username_status_label     = New-Object system.Windows.Forms.Label
$fired_username_status_label.text  = ""
$fired_username_status_label.visible  = $false
$fired_username_status_label.AutoSize  = $true
$fired_username_status_label.width  = 25
$fired_username_status_label.height  = 10
$fired_username_status_label.location  = New-Object System.Drawing.Point(244,70)
$fired_username_status_label.Font  = 'Microsoft Sans Serif,10'

$substitute_username_label       = New-Object system.Windows.Forms.Label
$substitute_username_label.text  = "Заместващ служител (UserName):"
$substitute_username_label.AutoSize  = $true
$substitute_username_label.width  = 25
$substitute_username_label.height  = 10
$substitute_username_label.location  = New-Object System.Drawing.Point(30,95)
$substitute_username_label.Font  = 'Microsoft Sans Serif,10'

$substitute_username_textbox     = New-Object system.Windows.Forms.TextBox
$substitute_username_textbox.multiline  = $false
$substitute_username_textbox.width  = 200
$substitute_username_textbox.height  = 20
$substitute_username_textbox.location  = New-Object System.Drawing.Point(30,115)
$substitute_username_textbox.Font  = 'Microsoft Sans Serif,10'

$substitute_username_status_label   = New-Object system.Windows.Forms.Label
$substitute_username_status_label.text  = ""
$substitute_username_status_label.visible  = $false
$substitute_username_status_label.AutoSize  = $true
$substitute_username_status_label.width  = 25
$substitute_username_status_label.height  = 10
$substitute_username_status_label.location  = New-Object System.Drawing.Point(245,120)
$substitute_username_status_label.Font  = 'Microsoft Sans Serif,10'

###############################################################################################

$Processing_Groupbox                  = New-Object system.Windows.Forms.Groupbox
$Processing_Groupbox.height           = 360
$Processing_Groupbox.width            = 520
$Processing_Groupbox.text             = "Конфигурация"
$Processing_Groupbox.location         = New-Object System.Drawing.Point(20,170)

$disable_fired_username_account_checkbox           = New-Object system.Windows.Forms.CheckBox
$disable_fired_username_account_checkbox.text      = "Спиране на акаунта на напускащия служител"
$disable_fired_username_account_checkbox.AutoSize  = $false
$disable_fired_username_account_checkbox.Checked   = $true
$disable_fired_username_account_checkbox.width     = 340
$disable_fired_username_account_checkbox.height    = 20
$disable_fired_username_account_checkbox.location  = New-Object System.Drawing.Point(30,195)
$disable_fired_username_account_checkbox.Font      = 'Microsoft Sans Serif,10'

$email_forward_checkbox          = New-Object system.Windows.Forms.CheckBox
$email_forward_checkbox.text     = "Пренасочване на входящата поща към заместващ служител"
$email_forward_checkbox.AutoSize  = $false
$email_forward_checkbox.Checked = $true
$email_forward_checkbox.width    = 420
$email_forward_checkbox.height   = 20
$email_forward_checkbox.location  = New-Object System.Drawing.Point(30,220)
$email_forward_checkbox.Font     = 'Microsoft Sans Serif,10'

$email_forward_to_both_checkbox          = New-Object system.Windows.Forms.CheckBox
$email_forward_to_both_checkbox.text     = "Да се доставят писмата и в кутията на напускащият служител"
$email_forward_to_both_checkbox.AutoSize  = $false
$email_forward_to_both_checkbox.Enabled  = $true
$email_forward_to_both_checkbox.Checked = $false
$email_forward_to_both_checkbox.width    = 450
$email_forward_to_both_checkbox.height   = 20
$email_forward_to_both_checkbox.location  = New-Object System.Drawing.Point(45,245)
$email_forward_to_both_checkbox.Font     = 'Microsoft Sans Serif,10'

$email_export_checkbox           = New-Object system.Windows.Forms.CheckBox
$email_export_checkbox.text      = "Запазване на наличната кореспонденция"
$email_export_checkbox.AutoSize  = $false
$email_export_checkbox.Checked   = $true
$email_export_checkbox.width     = 300
$email_export_checkbox.height    = 20
$email_export_checkbox.location  = New-Object System.Drawing.Point(30,270)
$email_export_checkbox.Font      = 'Microsoft Sans Serif,10'

$email_forward_time_label_1      = New-Object system.Windows.Forms.Label
$email_forward_time_label_1.text  = "след:"
$email_forward_time_label_1.AutoSize  = $true
$email_forward_time_label_1.width  = 25
$email_forward_time_label_1.height  = 10
$email_forward_time_label_1.location  = New-Object System.Drawing.Point(85,400)
$email_forward_time_label_1.Font  = 'Microsoft Sans Serif,10'

$email_forward_time_combobox     = New-Object system.Windows.Forms.ComboBox
$email_forward_time_combobox.text  = "1"
$email_forward_time_combobox.items.Add("1")
$email_forward_time_combobox.items.Add("2")
$email_forward_time_combobox.items.Add("3")
$email_forward_time_combobox.items.Add("4")
$email_forward_time_combobox.items.Add("5")
$email_forward_time_combobox.items.Add("6")
$email_forward_time_combobox.width  = 40
$email_forward_time_combobox.height  = 20
$email_forward_time_combobox.location  = New-Object System.Drawing.Point(130,398)
$email_forward_time_combobox.Font  = 'Microsoft Sans Serif,10'

$email_forward_time_label_2      = New-Object system.Windows.Forms.Label
$email_forward_time_label_2.text  = "месец(а)"
$email_forward_time_label_2.AutoSize  = $true
$email_forward_time_label_2.width  = 25
$email_forward_time_label_2.height  = 10
$email_forward_time_label_2.location  = New-Object System.Drawing.Point(175,400)
$email_forward_time_label_2.Font  = 'Microsoft Sans Serif,10'

$email_export_label              = New-Object system.Windows.Forms.Label
$email_export_label.text         = "с достъп на:"
$email_export_label.AutoSize     = $true
$email_export_label.width        = 25
$email_export_label.height       = 10
$email_export_label.location     = New-Object System.Drawing.Point(43,295)
$email_export_label.Font         = 'Microsoft Sans Serif,10'

$email_export_combobox           = New-Object system.Windows.Forms.ComboBox
$email_export_combobox.text      = "заместващ служител"
$email_export_combobox.Items.Add("заместващ служител")
$email_export_combobox.Items.Add("друг служител")
$email_export_combobox.width     = 175
$email_export_combobox.height    = 20
$email_export_combobox.location  = New-Object System.Drawing.Point(131,293)
$email_export_combobox.Font      = 'Microsoft Sans Serif,10'

$archive_access_username_textbox   = New-Object system.Windows.Forms.TextBox
$archive_access_username_textbox.multiline  = $false
$archive_access_username_textbox.visible  = $false
$archive_access_username_textbox.width  = 200
$archive_access_username_textbox.height  = 20
$archive_access_username_textbox.location  = New-Object System.Drawing.Point(30,320)
$archive_access_username_textbox.Font  = 'Microsoft Sans Serif,10'

$archive_access_username_status_label   = New-Object system.Windows.Forms.Label
$archive_access_username_status_label.text  = ""
$archive_access_username_status_label.AutoSize  = $true
$archive_access_username_status_label.Visible = $false
$archive_access_username_status_label.width  = 25
$archive_access_username_status_label.height  = 10
$archive_access_username_status_label.location  = New-Object System.Drawing.Point(245,325)
$archive_access_username_status_label.Font  = 'Microsoft Sans Serif,10'

$disable_activesync_checkbox          = New-Object system.Windows.Forms.CheckBox
$disable_activesync_checkbox.text     = "Спиране на ActiveSync"
$disable_activesync_checkbox.AutoSize  = $false
$disable_activesync_checkbox.Checked = $true
$disable_activesync_checkbox.width    = 420
$disable_activesync_checkbox.height   = 20
$disable_activesync_checkbox.location  = New-Object System.Drawing.Point(30,350)
$disable_activesync_checkbox.Font     = 'Microsoft Sans Serif,10'

$remove_user_notification_checkbox           = New-Object system.Windows.Forms.CheckBox
$remove_user_notification_checkbox.text      = "Нотификация за изтриване на акаунт"
$remove_user_notification_checkbox.AutoSize  = $false
$remove_user_notification_checkbox.Checked   = $true
$remove_user_notification_checkbox.width     = 300
$remove_user_notification_checkbox.height    = 20
$remove_user_notification_checkbox.location  = New-Object System.Drawing.Point(30,375)
$remove_user_notification_checkbox.Font      = 'Microsoft Sans Serif,10'

$ticket_number_label             = New-Object system.Windows.Forms.Label
$ticket_number_label.text        = "Номер на заявката:"
$ticket_number_label.AutoSize    = $true
$ticket_number_label.width       = 20
$ticket_number_label.height      = 10
$ticket_number_label.location    = New-Object System.Drawing.Point(32,425)
$ticket_number_label.Font        = 'Microsoft Sans Serif,10'

$ticket_number_textbox           = New-Object system.Windows.Forms.TextBox
$ticket_number_textbox.multiline  = $false
$ticket_number_textbox.width     = 274
$ticket_number_textbox.height    = 20
$ticket_number_textbox.location  = New-Object System.Drawing.Point(32,450)
$ticket_number_textbox.Font      = 'Microsoft Sans Serif,10'

###############################################################################################

$Status_Groupbox                  = New-Object system.Windows.Forms.Groupbox
$Status_Groupbox.height           = 250
$Status_Groupbox.width            = 520
$Status_Groupbox.text             = "Статус"
$Status_Groupbox.location         = New-Object System.Drawing.Point(555,170)

$status_textbox                  = New-Object system.Windows.Forms.TextBox
$status_textbox.multiline        = $true
$status_textbox.BackColor        = "#000000"
$status_textbox.ForeColor        = "#34AD3F"
$status_textbox.width            = 500
$status_textbox.height           = 220
$status_textbox.location         = New-Object System.Drawing.Point(565,190)
$status_textbox.Font             = 'Microsoft Sans Serif,10'

###############################################################################################

$Buttons_Groupbox                  = New-Object system.Windows.Forms.Groupbox
$Buttons_Groupbox.height           = 100
$Buttons_Groupbox.width            = 520
$Buttons_Groupbox.text             = "Управление"
$Buttons_Groupbox.location         = New-Object System.Drawing.Point(555,430)

$begin_button                    = New-Object system.Windows.Forms.Button
$begin_button.text               = "Begin"
$begin_button.Enabled            = $false 
$begin_button.width              = 80
$begin_button.height             = 30
$begin_button.location           = New-Object System.Drawing.Point(565,455)
$begin_button.Font               = 'Microsoft Sans Serif,10,style=Bold'

$copy_to_clipboard_button        = New-Object system.Windows.Forms.Button
$copy_to_clipboard_button.text   = "Копирай отговор за тикет"
$copy_to_clipboard_button.width  = 180
$copy_to_clipboard_button.height  = 30
$copy_to_clipboard_button.Enabled = $true
$copy_to_clipboard_button.location  = New-Object System.Drawing.Point(565,490)
$copy_to_clipboard_button.Font   = 'Microsoft Sans Serif,10'

###############################################################################################

#ADD ITR logo
$ITRlogo_picturebox                     = New-Object system.Windows.Forms.PictureBox
$ITRlogo_picturebox.width               = 1000
$ITRlogo_picturebox.height              = 80
$ITRlogo_picturebox.location            = New-Object System.Drawing.Point(330,50)
$ITRlogo_picturebox.imageLocation       = ".\logo.png"
$ITRlogo_picturebox.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$Form.controls.AddRange(@($ITRlogo_picturebox))

$Form.controls.AddRange(@($fired_username_label,
$fired_username_textbox,
$fired_username_status_label,
$substitute_username_label,
$substitute_username_textbox,
$substitute_username_status_label,
$disable_fired_username_account_checkbox,
$email_forward_checkbox,
$email_forward_to_both_checkbox,
$email_forward_time_label_1,
$email_forward_time_combobox,
$email_forward_time_label_2,
$email_export_label,
$email_export_checkbox,
$email_export_combobox,
$archive_access_username_textbox,
$archive_access_username_status_label,
$disable_activesync_checkbox,
$remove_user_notification_checkbox,
$ticket_number_label,
$ticket_number_textbox,
$status_textbox,
$begin_button,
$copy_to_clipboard_button,
$Users_Groupbox,
$Processing_Groupbox,
$Status_Groupbox,
$Buttons_Groupbox,
$ITRlogo_picturebox))

#region gui events {
$fired_username_textbox.Add_Validating({
    if ($fired_username_textbox.Text -ne "")
    {
        $fired_username = Get-ADUser -Filter {samAccountName -eq $fired_username_textbox.Text} -Properties *

        if ($fired_username)
        {
            $fired_username_status_label.visible = $true
            $fired_username_status_label.Text = $fired_username.Name + "    |    " + $fired_username.Company
            $fired_username_status_label.ForeColor = "#000000"
        }
        else
        {
            $fired_username_status_label.visible = $true
            $fired_username_status_label.Text = "Няма такъв потребител"
            $fired_username_status_label.ForeColor = "#fa0101"
        }
    }
    else
    {
        $fired_username_status_label.visible = $true
        $fired_username_status_label.Text = "Моля въведете потребителско име"
        $fired_username_status_label.ForeColor = "#fa0101"
    }
    Check-PrerequisitesForBeginButton
})

$substitute_username_textbox.Add_Validating({
    if ($substitute_username_textbox.Text -ne "")
    {
        if ($substitute_username_textbox.Text -eq $fired_username_textbox.Text)
        {
            $substitute_username_status_label.visible = $true
            $substitute_username_status_label.Text = "Моля въведете потребителско име, `r`nразлично от това на напускащия служител"
            $substitute_username_status_label.ForeColor = "#fa0101"
        }
        else
        {        
            $substitute_username = Get-ADUser -Filter {samAccountName -eq $substitute_username_textbox.Text} -Properties *

            if ($substitute_username)
            {
                $substitute_username_status_label.visible = $true
                $substitute_username_status_label.Text = $substitute_username.Name + "    |    " + $substitute_username.Company
                $substitute_username_status_label.ForeColor = "#000000"
            }
            else
            {
                $substitute_username_status_label.visible = $true
                $substitute_username_status_label.Text = "Няма такъв потребител"
                $substitute_username_status_label.ForeColor = "#fa0101"
            }
        }
    }
    else
    {
        $substitute_username_status_label.visible = $true
        $substitute_username_status_label.Text = "Моля въведете потребителско име"
        $substitute_username_status_label.ForeColor = "#fa0101"
    }
    Check-PrerequisitesForBeginButton
})

$email_export_combobox.Add_TextChanged({
    if ($email_export_combobox.Text -eq "друг служител")
    {
        $archive_access_username_textbox.Visible = $true
        $archive_access_username_status_label.Visible = $true
    }
    else
    {
        $archive_access_username_textbox.Visible = $false
        $archive_access_username_status_label.Visible = $false
    }
    Check-PrerequisitesForBeginButton
})

$archive_access_username_textbox.Add_Validating({
    if ($archive_access_username_textbox.Text -ne "")
    {
       if ($archive_access_username_textbox.Text -eq $fired_username_textbox.Text)
       {
            $archive_access_username_status_label.visible = $true
            $archive_access_username_status_label.Text = "Моля въведете потребителско име,`r`nразлично от това на напускащия служител"
            $archive_access_username_status_label.ForeColor = "#fa0101"
       }
       else
       {
            $archive_access_username = Get-ADUser -Filter {samAccountName -eq $archive_access_username_textbox.Text} -Properties *

            if ($archive_access_username)
            {
                $archive_access_username_status_label.visible = $true
                $archive_access_username_status_label.Text = $archive_access_username.Name + "    |    " + $archive_access_username.Company
                $archive_access_username_status_label.ForeColor = "#000000"
            }
            else
            {
                $archive_access_username_status_label.visible = $true
                $archive_access_username_status_label.Text = "Няма такъв потребител"
                $archive_access_username_status_label.ForeColor = "#fa0101"
            }
        }
    }
    else
    {
        $archive_access_username_status_label.visible = $true
        $archive_access_username_status_label.Text = "Моля въведете потребителско име"
        $archive_access_username_status_label.ForeColor = "#fa0101"
    }
    Check-PrerequisitesForBeginButton
})

$email_forward_checkbox.Add_CheckedChanged({
    if ($email_forward_checkbox.Checked -eq $true)
    {
        $email_forward_to_both_checkbox.Enabled  = $true
    }
    else 
    {
        $email_forward_to_both_checkbox.Enabled  = $false
    }
})

$remove_user_notification_checkbox.Add_CheckedChanged({
    if ($remove_user_notification_checkbox.Checked -eq $true)
    {
        $ticket_number_label.Enabled = $true
        $ticket_number_textbox.Enabled = $true
        $email_forward_time_combobox.Enabled = $true
        $email_forward_time_label_1.Enabled = $true
        $email_forward_time_label_2.Enabled = $true
    }
    else 
    {
        $ticket_number_label.Enabled = $false
        $ticket_number_textbox.Enabled = $false
        $email_forward_time_combobox.Enabled = $false
        $email_forward_time_label_1.Enabled = $false
        $email_forward_time_label_2.Enabled = $false
    }
})

# CHECK IF SQLserver module is installed
Import-Module SQLserver

if (Get-Module SQLserver)
{
    # do nothing
}

else
{
    Update-StatusTextbox "WARNING! Please run the following command from Powershell:" -NewLine
    Update-StatusTextbox "Install-Module SQLserver" -NewLine
    $remove_user_notification_checkbox.Enabled = $false
    $ticket_number_label.Enabled = $false
    $ticket_number_textbox.Enabled = $false
}


# BEGIN button
$begin_button.Add_Click({

# Disable user account in Active Directory
if ($disable_fired_username_account_checkbox.Checked -eq $true)
{   
    Update-StatusTextbox "Disable user account" -Tabulation
    Disable-ADAccount -Identity $fired_username_textbox.Text
    Start-Sleep -Seconds 5
    if ((Get-ADUser -Identity $fired_username_textbox.Text -Properties Enabled).Enabled -eq $true)
    {
        Update-StatusTextbox "ERROR" -NewLine 
    }
    else
    {
        Update-StatusTextbox "OK" -NewLine 
    }
}

# PSRemoting to Microsoft Exchange server
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mail.testcompany.com/PowerShell/ -Authentication Kerberos
Import-PSSession $ExchangeSession

# Enable email forwarding
if ($email_forward_checkbox.Checked -eq $true)
{
    Update-StatusTextbox -Text "Forward inbox messages" -Tabulation
    Enable-MailForwarding
}

# Export email to PST archive
if ($email_export_checkbox.Checked -eq $true)
{
    Update-StatusTextbox "Export email to archive" -Tabulation
    Export-EmailToArchive
}

# Disable ActiveSync
if ($disable_activesync_checkbox.Checked -eq $true)
    {
    Update-StatusTextbox "Disable ActiveSync" -Tabulation
    Disable-ActiveSync
    $ActiveSyncStatus = (Get-CASMailbox -Identity $fired_username_textbox.Text).ActiveSyncEnabled
    if ($ActiveSyncStatus -eq $false) 
    {
        Update-StatusTextbox "OK" -NewLine
    }
    else 
    {
        Update-StatusTextbox "ERROR" -NewLine
    }
}

# Remove PSSession to Exchange
Remove-PSSession $ExchangeSession

# Create automatic notification for account removal
if ($remove_user_notification_checkbox.Checked -eq $true)
{
    $fired_username = Get-ADUser -Filter {samAccountName -eq $fired_username_textbox.Text}
    $fired_username = $fired_username.SamAccountName
    $RemovalDate = (Get-Date).AddMonths($email_forward_time_combobox.text)
    $ticket_number = $ticket_number_textbox.Text

Export-ToSQLServer -SQLserverInstance 'MSSQL\MSSQL_M_DB' `
-SQLserverDatabase 'ITR1' `
-SQLserverSchema 'dbo' `
-SQLserverTable 'tbl_UsersForRemoval' `
-ColumnNames "SamAccountName","RemovalDate","TicketLink" `
-InsertData "$fired_username","$RemovalDate","$ticket_number"

    Update-StatusTextbox -Text "Create email notification for account removal" -Tabulation
    Update-StatusTextbox "OK" -NewLine 
}

})


$copy_to_clipboard_button.Add_Click({

    $fired_username = Get-ADUser -Filter {samAccountName -eq $fired_username_textbox.Text} -Properties *
    $substitute_username = Get-ADUser -Filter {samAccountName -eq $substitute_username_textbox.Text} -Properties *

    $CompanyAbbreviation = Get-CompanyAbbreviation -Company $fired_username.Company
    $mailArchiveDirectory = "\\testcompany.com\_Archives\" + $CompanyAbbreviation + "_" + $fired_username.SamAccountName
    
    $Message = Generate-ReplyMessage -FiredUser $fired_username.extensionAttribute4 -SubstituteUser $substitute_username.extensionAttribute4 -Months $email_forward_time_combobox.text -Archive $mailArchiveDirectory
    $Message | Set-Clipboard

    Update-StatusTextbox -Text "Reply message is succesfully copied to clipboard" -NewLine
})
#endregion events }

#endregion GUI }


#Write your logic code here
function Check-PrerequisitesForBeginButton
{
    if (($fired_username_status_label.Text -eq "Няма такъв потребител") -or ($fired_username_status_label.Text -eq ""))
    {
        $begin_button.Enabled = $false
    }
    elseif (($substitute_username_status_label.Text -eq "Няма такъв потребител") -or ($substitute_username_status_label.Text -eq ""))
    {
        $begin_button.Enabled = $false
    }
    elseif ($substitute_username_status_label.Text -eq "Моля въведете потребителско име,`r`nразлично от това на напускащия служител")
    {
        $begin_button.Enabled = $false
    }
    elseif ($email_export_combobox.Text -eq "друг служител")
    {
        if (($archive_access_username_status_label.Text -eq "Моля въведете потребителско име,`r`nразлично от това на напускащия служител") -or ($archive_access_username_status_label.Text -eq "") -or ($archive_access_username_status_label.Text -eq ""))
        {
            $begin_button.Enabled = $false
        }
        else
        {
            $begin_button.Enabled = $true
        }
    }
    else
    {
        $begin_button.Enabled = $true
    }
}

function Update-StatusTextbox
{
    param($Text,[switch]$NewLine,[switch]$Tabulation)

    Begin
    {
        # Calculate tabulation
        $MaxStringLength = 50
        $tabulationCount = ($MaxStringLength - $Text.Length)
        $tabulationCount = [System.Math]::Floor($tabulationCount)
        $tabulationText = " . " * $tabulationCount
    }

    Process
    {
        if ($NewLine)
        {
            $status_textbox.Text = $status_textbox.Text + $Text + "`r`n"
        }
        else
        {
            if ($Tabulation)
            {
                $status_textbox.Text = $status_textbox.Text + $Text + $tabulationText
            }
            else
            {
                $status_textbox.Text = $status_textbox.Text + $Text
            }
        }
    }

    End
    {
    }
}

function Get-CompanyAbbreviation
{
    param($Company)

    Begin
    {
        # Load XML with all company information
        [xml]$Company_BG_Xml = Get-Content -Path \\testcompany.com\_XML\companies.xml -Encoding UTF8
    }
    Process
    {
        $CompanyAbbreviation = $Company_BG_Xml.CompaniesTable.CompaniesRows | Where-Object {$_.Companies_ActiveDirectory -eq "$Company"} | select Abbreviation
        $CompanyAbbreviation = $CompanyAbbreviation.Abbreviation
        if ($CompanyAbbreviation -eq $null)
        {
            $CompanyAbbreviation = "Can't find company: $Company"
        }
    }
    End
    {
        $CompanyAbbreviation
    }
}

function Enable-MailForwarding
{
    $fired_username = Get-ADUser -Filter {samAccountName -eq $fired_username_textbox.Text} -Properties *
    $substitute_username = Get-ADUser -Filter {samAccountName -eq $substitute_username_textbox.Text} -Properties *
    
    Set-Mailbox -identity $fired_username.SamAccountName -ForwardingAddress $substitute_username.SamAccountName
    
    Start-Sleep -Seconds 2

    $forwardingAddress = Get-Mailbox -identity $fired_username.SamAccountName | select ForwardingAddress
    $forwardingAddress = $forwardingAddress.ForwardingAddress
    $substitute_username_canonicalname = $substitute_username.CanonicalName

    if ($forwardingAddress -eq $substitute_username_canonicalname)
    {
        Update-StatusTextbox "OK" -NewLine
    }
    else
    {
        Update-StatusTextbox "ERROR" -NewLine
    }
}

function Export-EmailToArchive
{
    #param($fired_userame,$access_to_archive_username,$exchange_server_name,$file_path)
    
    Begin
    {
        $fired_username = Get-ADUser -Filter {samAccountName -eq $fired_username_textbox.Text} -Properties *
        $fired_username_UserPrincipalName = $fired_username.UserPrincipalName
        $fired_username_SamAccountName = $fired_username.SamAccountName
        $CompanyAbbreviation = Get-CompanyAbbreviation -Company $fired_username.Company
        $mailArchiveDirectory = "\\FS-06\_Archives\" + $CompanyAbbreviation + "_" + $fired_username.SamAccountName
        $mailbox_file_path = $mailArchiveDirectory + "\" + $fired_username.SamAccountName + "UMB.pst"
        $archive_mailbox_file_path = $mailArchiveDirectory + "\" + $fired_username.SamAccountName + "AMB.pst"
    }
    Process
    {
        if($email_export_combobox.text -eq "заместващ служител")
        {
            $archive_access_username = Get-ADUser -Filter {samAccountName -eq $substitute_username_textbox.Text} -Properties *
            $archive_access_username = $archive_access_username.SamAccountName
        }
        else
        {
            $archive_access_username = Get-ADUser -Filter {samAccountName -eq $archive_access_username_textbox.Text} -Properties *
            $archive_access_username = $archive_access_username.SamAccountName
        }

        # Create folder for export and set appropriate permissions to it
        New-Item -Path "$mailArchiveDirectory" -ItemType Directory
        $ACL = Get-Acl -Path "$mailArchiveDirectory"
        $ACLarguments = New-Object system.security.accesscontrol.filesystemaccessrule("$archive_access_username",'ReadAndExecute','ContainerInherit,ObjectInherit','None','Allow')
        $ACL.AddAccessRule($ACLarguments)
        Set-Acl "$mailArchiveDirectory" $ACL

        # Export mailbox
        New-MailboxExportRequest -Mailbox $fired_username_UserPrincipalName -FilePath $mailbox_file_path -Name "$fired_username_SamAccountName`UMB"
        Update-StatusTextbox "PENDING" -NewLine
        
        # Check if online archive is enabled
        $archive_mailbox_status = Get-Mailbox -identity $fired_username_UserPrincipalName | select ArchiveDatabase
        $archive_mailbox_status = $archive_mailbox_status.ArchiveDatabase

        # Export archive mailbox
        if ($archive_mailbox_status -ne $null)
        {
            New-MailboxExportRequest -Mailbox $fired_username_UserPrincipalName -isArchive -FilePath $archive_mailbox_file_path -Name "$fired_username_SamAccountName`AMB"
        }
    }
    End
    {
    }
}

function Generate-ReplyMessage
{
    param($FiredUser,$SubstituteUser,$Months,$Archive)
    
    Begin
    {
        $Message = [System.Text.Encoding]::Default
    }

    Process
    {
        $Message = "Здравейте,

Входящата поща на $FiredUser е пренасочена към $SubstituteUser в рамките на $Months месец(а).
Наличната кореспонденция от пощата на $FiredUser е експортирана във вид на два .PST файла и се намира тук:
{code}
$Archive
{code}
Акаунта на $FiredUser ще бъде изтрит след $Months месец(а)."
    }

    End
    {
        $Message
    }

}

function Disable-ActiveSync
{
    Begin
    {
        $fired_username = Get-ADUser -Filter {samAccountName -eq $fired_username_textbox.Text} -Properties *
        $fired_username_UserPrincipalName = $fired_username.UserPrincipalName
    }

    Process
    {
        Set-CASMailbox -Identity $fired_username_UserPrincipalName -ActiveSyncEnabled $false
    }

    End
    {
        
    }
}

# Import sql server module
#Install-Module -Name SqlServer 
#Import-Module SQLserver

<#
.DESCRIPTION
   This function can be used to add data into previously created SQL Table.
   IMPORTANT! You need "SQLserver" module to use this function. Please use "Import-Module SQLserver" before using this function.
.EXAMPLE
   #Export-ToSQLServer -SQLserverInstance 
#>
function Export-ToSQLServer
{
    param($SQLserverInstance,$SQLserverDatabase,$SQLserverSchema,$SQLserverTable,[string[]]$ColumnNames,[string[]]$InsertData)

    $InsertResults = "INSERT INTO [$SQLserverSchema].[$SQLserverTable]("
    foreach ($ColumnName in $ColumnNames)
    {
        $InsertResults = $InsertResults + "$ColumnName" + ","
    }

    $InsertResults = $InsertResults.TrimEnd(',')
    $InsertResults = $InsertResults + ") "
    $InsertResults = $InsertResults + "VALUES ("
    foreach ($InsertValue in $InsertData)
    {
        $InsertResults = $InsertResults + "`'$InsertValue`'" + ","
    }
    $InsertResults = $InsertResults.TrimEnd(',')
    $InsertResults = $InsertResults + ")"
    Invoke-sqlcmd -ServerInstance $SQLserverInstance -Database $SQLserverDatabase -Query $InsertResults
}

[void]$Form.ShowDialog()

#