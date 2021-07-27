
<#PSScriptInfo

.VERSION 1.1

.GUID db05befa-dd3c-4666-88d2-cc31b84fb4f5

.AUTHOR dlinds

.COMPANYNAME 

.COPYRIGHT 

.TAGS 

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES


#>

<# 

.DESCRIPTION 
 Edit Calendar Permissions on Users/Rooms/Equipment Mailboxes, or on a user's Sub Calendar Folders 

#> 
Param()


#region XAML Main
$inputxml='<Window x:Class="ChangeRoomPermissions.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:ChangeRoomPermissions"
        mc:Ignorable="d"
        Title="Assign Room Permissions" Height="370" Width="600">
    <Grid Visibility="Visible">
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
            <ColumnDefinition Width="0*"/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <Button Name="Button_Collect365Creds" Content="Collect 365 Credentials" HorizontalAlignment="Left" Margin="233,150,0,0" VerticalAlignment="Top" Width="138"  Background="#FFF3F3F3" Height="20">
            <Button.Effect>
                <DropShadowEffect Color="#FFB8B8B8" BlurRadius="2" ShadowDepth="2" Opacity="0.6"/>
            </Button.Effect>
        </Button>
        <Label x:Name="Label_MailboxName" Content="Mailbox Name" HorizontalAlignment="Left" Margin="8,19,0,0" VerticalAlignment="Top" Height="26" Width="87"/>
        <TextBox HorizontalAlignment="Left" Height="23" Margin="12,40,0,0" TextWrapping="Wrap" Name="Textbox_RoomName" VerticalAlignment="Top" Width="189">
            <TextBox.Effect>
                <DropShadowEffect Color="#FFB8B8B8" BlurRadius="2" ShadowDepth="2" Opacity="0.6"/>
            </TextBox.Effect>
        </TextBox>
        <ComboBox x:Name="ComboBox_DifferentFolders" HorizontalAlignment="Left" Margin="12,77,0,0" VerticalAlignment="Top" Width="189" Height="22">
            <ComboBox.Effect>
                <DropShadowEffect Color="#FFB8B8B8" BlurRadius="2" ShadowDepth="2" Opacity="0.6"/>
            </ComboBox.Effect>
            <ComboBoxItem Content="Calendar" HorizontalAlignment="Left" Width="189"/>
            <ComboBoxItem Content="Contacts" HorizontalAlignment="Left" Width="189"/>
        </ComboBox>
        <Button x:Name="Button_FindPerms" Content="Find" HorizontalAlignment="Left" Margin="227,40,0,0" VerticalAlignment="Top" Width="100" Background="#FFF3F3F3" Height="21">
            <Button.Effect>
                <DropShadowEffect Color="#FFB8B8B8" BlurRadius="2" ShadowDepth="2" Opacity="0.6"/>
            </Button.Effect>
        </Button>
        <DataGrid HorizontalAlignment="Left" Margin="12,121,0,0" VerticalAlignment="Top"   
   Height="171" Width="560" x:Name="Datagrid_PermsList" AutoGenerateColumns="False"  
   AlternatingRowBackground="#FFF7F7F7" GridLinesVisibility="Horizontal"  RowHeaderWidth="0"
   SelectionUnit="FullRow" SelectionMode="Single" BorderBrush="#FFD1D1D1" Background="White" Visibility="Visible" CanUserResizeRows="False" CanUserResizeColumns="False">
            <DataGrid.Effect>
                <DropShadowEffect Color="#FFB8B8B8" BlurRadius="2" ShadowDepth="2" Opacity="0.6"/>
            </DataGrid.Effect>
            <DataGrid.Columns>
                <DataGridTextColumn Binding="{Binding Folder}" Header="Folder Name" IsReadOnly="True" Width="186" />
                <DataGridTextColumn Binding="{Binding Username}" Header="User" IsReadOnly="True" Width="186" />
                <DataGridTextColumn Binding="{Binding PermissionLevel}" Header="Permissions Level" IsReadOnly="True" Width="186" />
            </DataGrid.Columns>
        </DataGrid>
        <Button Name="Button_AddUsersToFolder" Content="Add" HorizontalAlignment="Left" Margin="242,300,0,0" VerticalAlignment="Top" Width="100" Background="#FFF3F3F3" Height="21">
            <Button.Effect>
                <DropShadowEffect Color="#FFB8B8B8" BlurRadius="2" ShadowDepth="2" Opacity="0.6"/>
            </Button.Effect>
        </Button>
        <Button Name="Button_EditUserPerms" Content="Edit" HorizontalAlignment="Left" Margin="357,300,0,0" VerticalAlignment="Top" Width="100" Background="#FFF3F3F3" Height="21">
            <Button.Effect>
                <DropShadowEffect Color="#FFB8B8B8" BlurRadius="2" ShadowDepth="2" Opacity="0.6"/>
            </Button.Effect>
        </Button>
        <Button Name="Button_RemoveUser" Content="Remove" HorizontalAlignment="Left" Margin="472,300,0,0" VerticalAlignment="Top" Width="100" Background="#FFF3F3F3" Height="21">
            <Button.Effect>
                <DropShadowEffect Color="#FFB8B8B8" BlurRadius="2" ShadowDepth="2" Opacity="0.6"/>
            </Button.Effect>
        </Button>

    </Grid>
</Window>
'
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
 
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{
    $Form=[Windows.Markup.XamlReader]::Load( $reader )
}
catch{
    Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
    throw
}
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
   
$xaml.SelectNodes("//*[@Name]") | %{"trying item $($_.Name)";
    try {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop}
    catch{throw}
    }

#endregion

#region XAML - FORM FOR EDITING EXISTING
$inputxml='<Window x:Class="ChangeRoomPermissions.EditExisting"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:ChangeRoomPermissions"
        mc:Ignorable="d"
        Title="Edit Existing Permissions" Height="244.667" Width="327.556">
    <Grid>
        <GroupBox Name="Groupbox_EditingDisplayName" Header="Folder: Calendar" HorizontalAlignment="Left" Height="162" Margin="25,25,0,0" VerticalAlignment="Top" Width="270">
            <Grid Name="Grid_PermOptions" HorizontalAlignment="Left" Height="135" VerticalAlignment="Top" Width="255">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="0*"/>
                </Grid.ColumnDefinitions>
                <Button Content="OK" Name="Button_OK" HorizontalAlignment="Left" Margin="4,107,0,0" VerticalAlignment="Top" Width="75"  Background="#FFF3F3F3" Height="20">
                    <Button.Effect>
                        <DropShadowEffect Color="#FFB8B8B8" BlurRadius="2" ShadowDepth="2" Opacity="0.6"/>
                    </Button.Effect>
                </Button>
                <Button Content="Cancel" Name="Button_Cancel" HorizontalAlignment="Left" Margin="100,107,0,0" VerticalAlignment="Top" Width="75"  Background="#FFF3F3F3" Height="20">
                    <Button.Effect>
                        <DropShadowEffect Color="#FFB8B8B8" BlurRadius="2" ShadowDepth="2" Opacity="0.6"/>
                    </Button.Effect>
                </Button>
                <ComboBox Name="Combo_PermissionLevel" HorizontalAlignment="Left" Margin="4,59,0,0" VerticalAlignment="Top" Width="241" Height="22">
                    <ComboBox.Effect>
                        <DropShadowEffect Color="#FFB8B8B8" BlurRadius="2" ShadowDepth="2" Opacity="0.6"/>
                    </ComboBox.Effect>
                    <ComboBoxItem Content="Owner" Name="ComboItem_Owner" HorizontalAlignment="Left" Width="240"/>
                    <ComboBoxItem Content="Publishing Editor" Name="ComboItem_PublishingEditor" HorizontalAlignment="Left" Width="240"/>
                    <ComboBoxItem Content="Editor" Name="ComboItem_Editor" HorizontalAlignment="Left" Width="240"/>
                    <ComboBoxItem Content="Publishing Author" Name="ComboItem_PublishingAuthor" HorizontalAlignment="Left" Width="240"/>
                    <ComboBoxItem Content="Author" Name="ComboItem_Author" HorizontalAlignment="Left" Width="240"/>
                    <ComboBoxItem Content="Nonediting Author" Name="ComboItem_NonEditingAuthor" HorizontalAlignment="Left" Width="240"/>
                    <ComboBoxItem Content="Reviewer" Name="ComboItem_Reviewer" HorizontalAlignment="Left" Width="240"/>
                    <ComboBoxItem Content="Contributor" Name="ComboItem_Contributor" HorizontalAlignment="Left" Width="240"/>
                    <ComboBoxItem Content="Free/Busy Time, Subject, Location" Name="ComboItem_FreeBusyWithMore" HorizontalAlignment="Left" Width="240"/>
                    <ComboBoxItem Content="Free/Busy Time" Name="ComboItem_FreeBusy" HorizontalAlignment="Left" Width="240"/>
                    <ComboBoxItem Content="None" Name="ComboItem_None" HorizontalAlignment="Left" Width="240"/>
                </ComboBox>
                <Label Content="Set Permission Level to:" HorizontalAlignment="Left" Margin="-1,36,0,0" VerticalAlignment="Top" Height="26" Width="134"/>
                <Label Name="Label_DisplayName" Content="User: Display Name" HorizontalAlignment="Left" Margin="-2,0,0,0" VerticalAlignment="Top" Height="25" Width="223"/>
                <TextBox HorizontalAlignment="Left" Height="23" Name="Textbox_NewName" Margin="32,4,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="189" Visibility="Hidden"/>
            </Grid>
        </GroupBox>
    </Grid>
</Window>

'
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
 
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{
    $EditExistingForm=[Windows.Markup.XamlReader]::Load( $reader )
}
catch{
    Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
    throw
}
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
  
$xaml.SelectNodes("//*[@Name]") | %{"trying item $($_.Name)";
    try {Set-Variable -Name "EE_WPF$($_.Name)" -Value $EditExistingForm.FindName($_.Name) -ErrorAction Stop}
    catch{throw}
    }

#endregion

#region FUNCTIONS FOR ALL FORMS

function Connect-365 {

    <#
        .SYNOPSIS
        Collects the 365 credentials, and then connects PS to it

        .DESCRIPTION
        When the button Connect to 365 button is pushed, this function is called. It will credential prompt and then connect the session to 365.

        To test is connection is successful, Get-MsolDomain is called. If it works, HideorShowAll function is called. If it doesn't error pops up function returns false

    #>


    $credential365 = Get-Credential -Message "Enter in your Office 365 credentials"
    Connect-MsolService -Credential $credential365
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri “https://outlook.office365.com/powershell-liveid/” -Credential $credential365 -Authentication “Basic” –AllowRedirection
    Import-PSSession $session -AllowClobber
    try {
        Get-MsolDomain -ErrorAction Stop | Out-Null
        $WPFButton_Collect365Creds.Visibility = 'hidden'
        $WPFButton_Collect365Creds.IsEnabled = $false
        HideorShowAll -status 'show'
        }
    catch {
        [System.Windows.Forms.MessageBox]::Show("The supplied Office 365 credentials didn't work. Please try again")
        Write-Error "The supplied Office 365 credentials didn't work. Please try again"
        return $false
    }
    

}

function HideorShowAll {

    <#
        .SYNOPSIS
        Hides or shows elements

        .DESCRIPTION
        Parameter status is either hide or show, hides elements.

    #>

    param (
        $status
    )
    if ($status -eq 'hide') {
        $WPFLabel_MailboxName.Visibility = 'hidden'
        $WPFTextbox_RoomName.Visibility = 'hidden'
        $WPFButton_FindPerms.Visibility = 'hidden'
        $WPFDatagrid_PermsList.Visibility = 'hidden'
        $WPFComboBox_DifferentFolders.Visibility = 'hidden'
        $WPFButton_AddUsersToFolder.Visibility = 'hidden'
        $WPFButton_EditUserPerms.Visibility = 'hidden'
        $WPFButton_RemoveUser.Visibility = 'hidden'
    }
    if ($status -eq 'show') {
        $WPFLabel_MailboxName.Visibility = 'visible'
        $WPFTextbox_RoomName.Visibility = 'visible'
        $WPFComboBox_DifferentFolders.Visibility = 'visible'
        $WPFButton_FindPerms.Visibility = 'visible'
        $WPFDatagrid_PermsList.Visibility = 'visible'
        $WPFButton_AddUsersToFolder.Visibility = 'visible'
        $WPFButton_EditUserPerms.Visibility = 'visible'
        $WPFButton_RemoveUser.Visibility = 'visible'
    }
}

function ClearAndDisableValues {

    <#
        .SYNOPSIS
        Clears the values and essentially resets the form

        .DESCRIPTION
        Clears the values and essentially resets the form to the moment after 365 PS log in is successful. 

    #>

    $WPFDatagrid_PermsList.Items.Clear()
    $WPFComboBox_DifferentFolders.Items.Clear()
    if ($WPFTextbox_RoomName.Text.Length -eq 0) {
        $WPFButton_FindPerms.IsEnabled = $false
    } elseif ($WPFTextbox_RoomName.Text.Length -gt 0) {
        $WPFButton_FindPerms.IsEnabled = $true
    }
    $WPFButton_EditUserPerms.IsEnabled = $false
    $WPFComboBox_DifferentFolders.IsEnabled = $false
    $WPFButton_AddUsersToFolder.IsEnabled = $false
    $WPFButton_RemoveUser.IsEnabled = $false
}

function Find-MailboxFolders {

    <#
        .SYNOPSIS
        Checks the mailbox to see what Calendar folders are available

        .DESCRIPTION
        When a mailbox name is typed and searched in the text field, this will be called. 
        
        First checks to see if the mailbox exists, if it doesn't will return false and kill the function

        If it does exists, will get Calendar and Sub Calendars to combo box and enable clicking of the combobox


    #>


    try {
        Get-Mailbox -identity $WPFTextbox_RoomName.Text -erroraction Stop | Out-Null
        }
    catch {
        [System.Windows.Forms.MessageBox]::Show("That mailbox couldn't be found")
        Write-Error "That mailbox couldn't be found"
        return $false
    }
    $calendarFolders = Get-MailboxFolderStatistics -Identity $WPFTextbox_RoomName.Text -FolderScope Calendar
    $contactFolders = Get-MailboxFolderStatistics -Identity $WPFTextbox_RoomName.Text -FolderScope Contacts
    foreach ($item in $calendarFolders) {
        if ($item.Name -eq 'Calendar') {
            $WPFComboBox_DifferentFolders.Items.Add("$($item.Name)")
        } else {
            $WPFComboBox_DifferentFolders.Items.Add("Calendar\$($item.Name)")
        }
    }
    
    foreach ($item in $contactFolders) {
        if ($item.Name -eq 'Contacts') {
            $WPFComboBox_DifferentFolders.Items.Add("$($item.Name)")
        }
    }
    
    $WPFComboBox_DifferentFolders.IsEnabled = $true
}

function Get-FolderPerms {

    <#
        .SYNOPSIS
        Gets the permission is the selected Folder

        .DESCRIPTION
        When a folder is chosen in the drop down, this function is called and displays the permissions each user has on the folder.


    #>
    if ($WPFComboBox_DifferentFolders.Items.Count -gt 0) {
        $WPFDatagrid_PermsList.Items.Clear() #clear out the previous entries
        $WPFButton_AddUsersToFolder.IsEnabled = $true #enable button to add more users
        $perms = Get-MailboxFolderPermission -identity "$($WPFTextbox_RoomName.Text):\$($WPFComboBox_DifferentFolders.SelectedItem)"
        foreach ($item in $perms) { #display each permission
            $WPFDatagrid_PermsList.AddChild([pscustomobject]@{Folder=$($WPFComboBox_DifferentFolders.SelectedItem);Username=$item.User.DisplayName;PermissionLevel=$item.AccessRights;})
        }
    }
}

function Load-EditExistingFolderPermsForm {
        <#
        .SYNOPSIS
        Sets the Edit form

        .DESCRIPTION
        When Edit is selected, this function is called and sets the Add/Edit permissions form with Edit settings. Each time this function is called it will reset the
        form, as it's not generating it but hiding/unhiding so previous options will need to be reset. "Set" button calls Set-UserFolderPermissions

        #>
    
    $EditExistingForm.Icon = $bitmap #set the icon on the pop up
    



    #can't set Free/Busy on sub calendars, so hide these options unless it's main "Calendar"
    switch ($WPFComboBox_DifferentFolders.SelectedItem) {
    'Calendar' {
                $EE_WPFComboItem_FreeBusyWithMore.Visibility = "Visible"
                $EE_WPFComboItem_FreeBusy.Visibility = "Visible"
                }
    default    {
                $EE_WPFComboItem_FreeBusyWithMore.Visibility = "Collapsed"
                $EE_WPFComboItem_FreeBusy.Visibility = "Collapsed"
                }
    }
    
    #this hides the text box that is used for adding a new user to perms via Add-Functions
    $EE_WPFTextbox_NewName.Visibility = 'hidden'
    $EE_WPFGroupbox_EditingDisplayName.Header = "Folder: $($WPFComboBox_DifferentFolders.SelectedItem)"

    #before unhiding the form, null the selection of permission drop down
    $EE_WPFCombo_PermissionLevel.SelectedIndex = $null
    $EE_WPFLabel_DisplayName.Content = "User: $($WPFDatagrid_PermsList.SelectedItem.Username)"

    #can't close a form. ShowDialog should only run once 
    if ($EditExistingForm.Visibility -eq 'hidden') {
        $EditExistingForm.Visibility = 'visible'
    } else {
        $EditExistingForm.ShowDialog() | Out-null
        }
}

function Set-UserFolderPermissions {

        <#
        .SYNOPSIS
        Changes or adds the permissions via the $EditExistingForm form

        .DESCRIPTION
        When Add or Edit is selected, $EditExistingForm form is loaded. When any changes are made on that form, this function is called and sets the changes

        #>

    #"Free/Busy Time" and "Free/Busy Time, Subject, Location" don't match up with the AccessRights name like everything else does (at least once spaces are removed)
    #switch sets the accessrights variable to the correct name
    switch ($EE_WPFCombo_PermissionLevel.SelectedItem.Content.ToString()) {
        "Free/Busy Time" {
                           $accessRights = "AvailabilityOnly"
                        }
        "Free/Busy Time, Subject, Location" {
                           $accessRights = "LimitedDetails"
                        }
        default {
                           $accessRights = $EE_WPFCombo_PermissionLevel.SelectedItem.Content.ToString() -replace '\s',''
                }
    }
    
    try {
        # change the permissions. Anonymous/Default users in Exchange don't have the ability to delete, so they can only be set. Switch item 'default' (which is different than Default user)
        # does, this is all users that have ever or will ever be added that aren't named Anonymous/Defaul
        switch ($WPFDatagrid_PermsList.SelectedItem.Username) {
            "Anonymous" { Set-MailboxFolderPermission -Identity "$($WPFTextbox_RoomName.Text):\$($WPFComboBox_DifferentFolders.SelectedItem)" -User $($WPFDatagrid_PermsList.SelectedItem.Username) -AccessRights $accessRIghts -ErrorAction Stop | Out-Null
                        }

            "Default"   { Set-MailboxFolderPermission -Identity "$($WPFTextbox_RoomName.Text):\$($WPFComboBox_DifferentFolders.SelectedItem)" -User $($WPFDatagrid_PermsList.SelectedItem.Username) -AccessRights $accessRIghts -ErrorAction Stop | Out-Null
                        }

            default     {
                           if ($accessRights -eq 'None') {
                                Remove-MailboxFolderPermission -identity "$($WPFTextbox_RoomName.Text):\$($WPFComboBox_DifferentFolders.SelectedItem)" -User $($WPFDatagrid_PermsList.SelectedItem.Username) -Confirm:$False -ErrorAction Stop | Out-Null 
                           } else {
                                Set-MailboxFolderPermission -Identity "$($WPFTextbox_RoomName.Text):\$($WPFComboBox_DifferentFolders.SelectedItem)" -User $($WPFDatagrid_PermsList.SelectedItem.Username) -AccessRights $accessRIghts -ErrorAction Stop | Out-Null
                                }
                        }
        }
        [System.Windows.Forms.MessageBox]::Show("Permissions successfully changed!")
        return $true
        }
    catch {
        Write-Host $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Something went wrong and the permissions couldn't be changed. `n`nHere is the error: $($_.Exception.Message)")
        return $false
         }
}
function Load-AddNewUserPermsForm {
        <#
        .SYNOPSIS
        Opens the Edit/Add form

        .DESCRIPTION
        When Add is selected, this function is called and opens the Add/Edit permissions form as Add. Each time this function is called it will reset the
        form, as it's not generating it but hiding/unhiding so previous options will need to be reset. "Set" button calls Set-UserFolderPermissions

        #>

    $EditExistingForm.Icon = $bitmap
    $EE_WPFTextbox_NewName.Visibility = 'visible' #unhide it
    $EE_WPFGroupbox_EditingDisplayName.Header = "Folder: $($WPFComboBox_DifferentFolders.SelectedItem)"
    #reset the form
    $EE_WPFCombo_PermissionLevel.SelectedIndex = $null
    $EE_WPFLabel_DisplayName.Content = "User:"
    #can't close a form. ShowDialog should only run once 
    if ($EditExistingForm.Visibility -eq 'hidden') {
        $EditExistingForm.Visibility = 'visible'
    } else {
        $EditExistingForm.ShowDialog() | Out-null
        }
}

function Add-UserFolderPermissions {
        <#
        .SYNOPSIS
        Adds the permissions via the $EditExistingForm form

        .DESCRIPTION
        When Add is selected, $EditExistingForm form is loaded. When any additions are added on that form, this function is called and adds the addition

        #>

    #"Free/Busy Time" and "Free/Busy Time, Subject, Location" don't match up with the AccessRights name like everything else does (at least once spaces are removed)
    #switch sets the accessrights variable to the correct name
    switch ($EE_WPFCombo_PermissionLevel.SelectedItem.Content.ToString()) {
        "Free/Busy Time" {
                           $accessRights = "AvailabilityOnly"
                        }
        "Free/Busy Time, Subject, Location" {
                           $accessRights = "LimitedDetails"
                        }
        default {
                           $accessRights = $EE_WPFCombo_PermissionLevel.SelectedItem.Content.ToString() -replace '\s',''
                }
    }

    try {
            #don't add anything is the accessright is None. Otherwise it will add
            if ($accessRIghts -ne 'None') {
                Add-MailboxFolderPermission -Identity "$($WPFTextbox_RoomName.Text):\$($WPFComboBox_DifferentFolders.SelectedItem)" -User $EE_WPFTextbox_NewName.Text -AccessRights $accessRIghts -ErrorAction Stop | Out-Null
            }
            [System.Windows.Forms.MessageBox]::Show("Permissions added!")
            return $true
        }
    catch {
            Write-Host $_.Exception.Message
            [System.Windows.Forms.MessageBox]::Show("Something went wrong and the permissions couldn't be added.`n`nHere is the error: $($_.Exception.Message)")
            return $false
         }

}

function Remove-UserPerms {
        <#
        .SYNOPSIS
        Removes a user from the Calendar permissions

        .DESCRIPTION
        When a user is selected that has permission to a folder, the remove button will become available. When it's pushed, this function will run and then will remove the permissions (after confirming)

        The Remove button will only be allowed for actual users. It is still disabled for Anonymous/Default as they cannot be removed.

        #>

    
    $permfullName = $WPFDatagrid_PermsList.SelectedItem.Username
    if ([System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove $permfullName from $($WPFComboBox_DifferentFolders.SelectedItem)?" , "Please Confirm Removal" , 4)  -eq "Yes") {
        try {
            Remove-MailboxFolderPermission -identity "$($WPFTextbox_RoomName.Text):\$($WPFComboBox_DifferentFolders.SelectedItem)" -User "$permfullName" -Confirm:$False -ErrorAction Stop | Out-Null
            Disable-AddRemoveButtons
            Get-FolderPerms
            }
        catch {
            Write-Host $_.Exception.Message
            [System.Windows.Forms.MessageBox]::Show("Unable to remove. Sorry about that!`n`nHere is the error: $($_.Exception.Message)")
            return $false
            }
    }


}
function Disable-AddRemoveButtons {
        <#
        .SYNOPSIS
        After an add/edit/remove, nulls and disabled buttons

        .DESCRIPTION
        After an add/edit/remove, this function is called and will unselect the datagrid and then disable the Edit/Remove (as nothing is selected)

        #>

    
    $WPFDatagrid_PermsList.UnselectAll()
    $WPFButton_EditUserPerms.IsEnabled = $false
    $WPFButton_RemoveUser.IsEnabled = $false
}

#endregion

#region REQUIRED START CODE
try {
    Import-Module MSOnline -ErrorAction Stop | Out-Null
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName PresentationFramework
    }
catch {
    [System.Windows.Forms.MessageBox]::Show("MSOnline PS Module not installed. Please install and try again")
    throw "MSOnline PS Module not installed. Please install and try again"
    exit
}

#at first, just the 365 creds collection box is shown.
$WPFButton_Collect365Creds.Visibility = 'visible'
HideorShowAll -status 'hide' #hides all the variables except 365 creds
ClearAndDisableValues
#endregion

#region USER INTERACTIONS MAIN FORM
$WPFButton_Collect365Creds.Add_Click({
    #Office 365 Credential Collection
    Connect-365
})

$WPFTextbox_RoomName.Add_TextChanged({
    #runs when the email textbox value is changed
    ClearAndDisableValues        
})

$WPFTextbox_RoomName.Add_GotFocus({
    #if the email textbox is selected, that means a user isn't selected. Runs function to disable Add/Remove/Edit
    Disable-AddRemoveButtons
})

$WPFComboBox_DifferentFolders.Add_GotFocus({
    #if the combobox is selected, that means a user isn't selected. Runs function to disable Add/Remove/Edit
    Disable-AddRemoveButtons
})

$WPFTextbox_RoomName.add_KeyDown({
    #runs only if the key that's pressed in the email textbox is Enter/Return. Then it searches for the mailbox and its folders
    if ($_.Key -eq 'Return')
    {
        if ($WPFTextbox_RoomName.Text.Length -gt 0) {
            Find-MailboxFolders
        }
    }
})

$WPFButton_FindPerms.Add_Click({
    #runs when the Find button is pushed. It searches for the mailbox and its folders after clearing the values
    if ($WPFTextbox_RoomName.Text.Length -gt 0) {
        ClearAndDisableValues
        Find-MailboxFolders
    }
})

$WPFComboBox_DifferentFolders.add_SelectionChanged({
    #when the combobox selection changes, this will run
    Get-FolderPerms
})

$WPFDatagrid_PermsList.Add_SelectionChanged({

    #if something is selected in main datagrid, this runs. 
    #certain options are enabled/disabled based on mailbox
    #Default User/Anonymous can only be edited
    switch ($WPFDatagrid_PermsList.SelectedItem.Username) {
        $null {
                $WPFButton_EditUserPerms.IsEnabled = $false 
                $WPFButton_RemoveUser.IsEnabled = $false
                break
              }
        "Anonymous" {
                $WPFButton_EditUserPerms.IsEnabled = $true
                $WPFButton_RemoveUser.IsEnabled = $false
                break
                }
        "Default" {
                $WPFButton_EditUserPerms.IsEnabled = $true
                $WPFButton_RemoveUser.IsEnabled = $false
                break
                }
        default {
                $WPFButton_EditUserPerms.IsEnabled = $true 
                $WPFButton_RemoveUser.IsEnabled = $true
                break
                }
        }
})

$WPFButton_EditUserPerms.Add_Click({
    #loads Edit/Add form with Edit settings when Edit button is pushed
    Load-EditExistingFolderPermsForm
})

$WPFButton_AddUsersToFolder.Add_Click({
    #loads Edit/Add form with Add settings when Add button is pushed
    Load-AddNewUserPermsForm
})

$WPFButton_RemoveUser.Add_Click({
    #when Remove button is pushed, will remove the user selected
    Remove-UserPerms
})
#endregion

#region USER INTERACTION EDIT EXISTING PERMS

#this is only the pop up form for Edit/Adds

$EditExistingForm.Add_Closing({
    #When the form is closed, this runs
    $_.Cancel = $true #only hides it, doesn't allow closing of form
    $EditExistingForm.Visibility = 'hidden'
    $EE_WPFTextbox_NewName.Text = ''
})
$EE_WPFButton_Cancel.Add_Click({
    #When Cancel button is selected, form is hidden
    $EditExistingForm.Visibility = 'hidden'
    $EE_WPFTextbox_NewName.Text = ''
})
$EE_WPFButton_OK.Add_Click({
    #when OK button is pushed, will make the changes desired
    
    #Will first check if the Add button or Edit button was called
    #if it was add, will run first If. For Edit, will run elseif
    if ($EE_WPFTextbox_NewName.IsVisible -eq $true) {    
        if (Add-UserFolderPermissions) {
            $EditExistingForm.Visibility = 'hidden'
            $EE_WPFTextbox_NewName.Text = ''
            Disable-AddRemoveButtons
            Get-FolderPerms
        }
    } elseif ($EE_WPFTextbox_NewName.IsVisible -eq $false) {    
        if (Set-UserFolderPermissions) {
            $EditExistingForm.Visibility = 'hidden'
            $EE_WPFTextbox_NewName.Text = ''
            Disable-AddRemoveButtons
            Get-FolderPerms
        }
    }
})

$Form.Add_Closing({
    #when the form is closed, will remove PSS Sessions
    Remove-PSSession * -ErrorAction SilentlyContinue
    Exit-PSSession
    Exit
})

#endregion

#change base64 string to your icon. Google "Convert image to base64" to find a converter. String below is generic 365 icon
$base64 = 'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAKZSURBVEhLY/hPOvjz9ePn5uL3EcZfpnb8fnIPKooDkGbB19mT3phyvBRheGPM+MaM/Y0hwytZhheiDB/zIr6tng1VhAqIsuD7ntXv4l1ecDG80Wd4Y8H51ob3rTUy4nljwvxam+EFJ8NbZ4kvPdW/n9yF6sRvwZ8Pbz6VJr8UY3ityvDGlAXDXCRkBSZteN+YcwBtemsrATUClwXfZk95Y8YJCQqgAxEGAY2AhIwyw2t9oBRMHAm9MWN7H6wLNQjNgh/7N7xP8gD6FBEUYPTGlPWNAcMrRYaXXAxvfVU+t1b+vnP5Y5Y3UBzNdCDCbsGXyS0vgEGhggiKNyYsb/QYXsowvBRieBdp/qWv/vf18xDFEPAxwf6NCdEWvPfXgDoHHJpA975P9/o6tevP80cQBZjgQ7wdKRaEGr4xZYMqsuV9wY09bpABZRbwDgILPg55HwyyIPowCH3wfjSSCYFBZgEZqYjmkUyJBTzPGQlbQGocGLwxZwdWCa9UGV5KMnwsTYGI4wEkWhCgBawD3kXb/Ni1CiJCEJBmwZe5rT/2roGwiQSkWQAE/35//1hR8hLYkJJi+BBv86W1/vvKCb+vnIFKYwCSLYCDv7++f6pOBrYBXvKDK31hhhf8DK+tud/HWn9qrPo6s+Xn6UNAZZ9yA4ANAzTTgYiwBXDw89zBjzlRL/kYXuswgBoZZmxvjBje6DK8UmEANple6zGgGQ1qjlhxv9ZneOcqCzUCvwVw8G3FnHf+msBUAGxtvLHkfmvDB2nawM0FJnFgU+qlKKhB9rEg/vuBjVCdRFoAB18mtLzSBKXjN8ZMwFYi0EOvlRlAyS/C7OuUzr/fPkPVIQHSLICAPy+ffK7Pfxes9aWn8eel41BRrOD/fwDTB5MQyM2byQAAAABJRU5ErkJggg=='
$bitmap = New-Object System.Windows.Media.Imaging.BitMapImage
$bitmap.BeginInit()
$bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($base64)
$bitmap.EndInit()
$bitmap.Freeze()

$form.Icon = $bitmap
$Form.ShowDialog() | Out-Null
