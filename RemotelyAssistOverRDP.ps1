$Global:sessionList = New-Object System.Collections.ArrayList

function Get-PC {
<#


.SYNOPSIS  
    Checks to see if the PC is online or not

.DESCRIPTION
    Checks to see if PC is online

#>
    try {
        Test-Connection $WPFTextBox_PCName.Text -Count 1 -ErrorAction Stop | Out-null
        return $true
        }
    catch {
        return $false
        throw
        }
}

function Get-Sessions {

<#


.SYNOPSIS  
    Gets the sessions and IDs from Get-UserSession and then displays them in window

.DESCRIPTION
    First will confirm PC is in AD via Get-PC. If so, then Get-UserSession is called. Based on what Get-UserSession returns, the listbox in form will be updated and the button to connect enables

#>

    if (!(Get-PC)) {
        $WPFLabel_PCNotFound.Visibility = 'visible'
    } else {
        #if the PC is online, this will run
        $Global:sessionList = Get-usersession -computername $WPFTextbox_PCName.Text
        #just because PC is online doesn't mean there are sessions actively logged in
        #first checks to see if there are sessions or not in array
        if ($Global:sessionList -eq $false) {
            $itm = new-object System.Windows.Controls.ListboxItem
            $itm.FontWeight = 'thin'
            $itm.Content = 'No active sessions found'
            $WPFListbox_UserList.IsEnabled = $false
            $WPFButton_Connect.IsEnabled = $false
            $WPFListbox_UserList.Items.Add($itm)
        } else {
            #if there are sessions in array, then it will add them to list box
            foreach ($session in $Global:sessionList) { 
                $dn = get-aduser -Identity $session.Username -Properties Name | Select Name        
                #format each listbox item to be pretty
                $itm = new-object System.Windows.Controls.ListboxItem
                $itm.FontWeight = 'thin'
                $itm.Content = $dn.name
                $WPFListbox_UserList.Items.Add($itm)
                $session | Add-Member NoteProperty "displayName" $dn.name
            }
            #enable the connect buttons and hide PC Not Found label, if it'd been set earlier    
            $WPFButton_Connect.IsEnabled = $true
            $WPFLabel_PCNotFound.Visibility = 'hidden'
            $WPFListbox_UserList.IsEnabled = $true
          }
    }
}
function Get-UserSession {
<# 

**************
This function is taken from https://www.powershellgallery.com/packages/WFTools/0.1.37/Content/Get-UserSession.ps1
**************

.SYNOPSIS
    Retrieves all user sessions from local or remote computers(s)

.DESCRIPTION
    Retrieves all user sessions from local or remote computer(s).
    
    Note: Requires query.exe in order to run
    Note: This works against Windows Vista and later systems provided the following registry value is in place
            HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\AllowRemoteRPC = 1
    Note: If query.exe takes longer than 15 seconds to return, an error is thrown and the next computername is processed. Suppress this with -erroraction silentlycontinue
    Note: If $sessions is empty, we return a warning saying no users. Suppress this with -warningaction silentlycontinue

.PARAMETER computername
    Name of computer(s) to run session query against
              
.parameter parseIdleTime
    Parse idle time into a timespan object

.parameter timeout
    Seconds to wait before ending query.exe process. Helpful in situations where query.exe hangs due to the state of the remote system.
                    
.FUNCTIONALITY
    Computers

.EXAMPLE
    Get-usersession -computername "server1"

    Query all current user sessions on 'server1'

.EXAMPLE
    Get-UserSession -computername $servers -parseIdleTime | ?{$_.idletime -gt [timespan]"1:00"} | ft -AutoSize

    Query all servers in the array $servers, parse idle time, check for idle time greater than 1 hour.

.NOTES
    Thanks to Boe Prox for the ideas - http://learn-powershell.net/2010/11/01/quick-hit-find-currently-logged-on-users/

.LINK
    http://gallery.technet.microsoft.com/Get-UserSessions-Parse-b4c97837
#> 
    [cmdletbinding()]
    Param(
        [Parameter(
            Position = 0,
            ValueFromPipeline = $True)]
        [string[]]$ComputerName = "localhost",

        [switch]$ParseIdleTime,

        [validaterange(0,120)]
        [int]$Timeout = 15
    )             
    Process
    {
        ForEach($computer in $ComputerName)
        {
        
            #start query.exe using .net and cmd /c.  We do this to avoid cases where query.exe hangs

                #build temp file to store results.  Loop until we see the file
                    Try
                    {
                        $Started = Get-Date
                        $tempFile = [System.IO.Path]::GetTempFileName()
                        
                        Do{
                            start-sleep -Milliseconds 300
                            
                            if( ((Get-Date) - $Started).totalseconds -gt 10)
                            {
                                Throw "Timed out waiting for temp file '$TempFile'"
                            }
                        }
                        Until(Test-Path -Path $tempfile)
                    }
                    Catch
                    {
                        Write-Error "Error for '$Computer': $_"
                        Continue
                    }

                #Record date.  Start process to run query in cmd.  I use starttime independently of process starttime due to a few issues we ran into
                    $Started = Get-Date
                    $p = Start-Process -FilePath C:\windows\system32\cmd.exe -ArgumentList "/c query user /server:$computer > $tempfile" -WindowStyle hidden -passthru

                #we can't read in info or else it will freeze.  We cant run waitforexit until we read the standard output, or we run into issues...
                #handle timeouts on our own by watching hasexited
                    $stopprocessing = $false
                    do
                    {
                    
                        #check if process has exited
                            $hasExited = $p.HasExited
                
                        #check if there is still a record of the process
                            Try
                            {
                                $proc = Get-Process -id $p.id -ErrorAction stop
                            }
                            Catch
                            {
                                $proc = $null
                            }

                        #sleep a bit
                            start-sleep -seconds .5

                        #If we timed out and the process has not exited, kill the process
                            if( ( (Get-Date) - $Started ).totalseconds -gt $timeout -and -not $hasExited -and $proc)
                            {
                                $p.kill()
                                $stopprocessing = $true
                                Remove-Item $tempfile -force
                                Write-Error "$computer`: Query.exe took longer than $timeout seconds to execute"
                            }
                    }
                    until($hasexited -or $stopProcessing -or -not $proc)
                    
                    if($stopprocessing)
                    {
                        Continue
                    }

                    #if we are still processing, read the output!
                        try
                        {
                            $sessions = Get-Content $tempfile -ErrorAction stop
                            Remove-Item $tempfile -force
                        }
                        catch
                        {
                            Write-Error "Could not process results for '$computer' in '$tempfile': $_"
                            continue
                        }
        
            #handle no results
            if($sessions){

                1..($sessions.count - 1) | Foreach-Object {
            
                    #Start to build the custom object
                    $temp = "" | Select ComputerName, Username, SessionName, Id, State, IdleTime, LogonTime
                    $temp.ComputerName = $computer

                    #The output of query.exe is dynamic. 
                    #strings should be 82 chars by default, but could reach higher depending on idle time.
                    #we use arrays to handle the latter.

                    if($sessions[$_].length -gt 5){
                        
                        #if the length is normal, parse substrings
                        if($sessions[$_].length -le 82){
                           
                            $temp.Username = $sessions[$_].Substring(1,22).trim()
                            $temp.SessionName = $sessions[$_].Substring(23,19).trim()
                            $temp.Id = $sessions[$_].Substring(42,4).trim()
                            $temp.State = $sessions[$_].Substring(46,8).trim()
                            $temp.IdleTime = $sessions[$_].Substring(54,11).trim()
                            $logonTimeLength = $sessions[$_].length - 65
                            try{
                                $temp.LogonTime = Get-Date $sessions[$_].Substring(65,$logonTimeLength).trim() -ErrorAction stop
                            }
                            catch{
                                #Cleaning up code, investigate reason behind this.  Long way of saying $null....
                                $temp.LogonTime = $sessions[$_].Substring(65,$logonTimeLength).trim() | Out-Null
                            }

                        }
                        
                        #Otherwise, create array and parse
                        else{                                       
                            $array = $sessions[$_] -replace "\s+", " " -split " "
                            $temp.Username = $array[1]
                
                            #in some cases the array will be missing the session name.  array indices change
                            if($array.count -lt 9){
                                $temp.SessionName = ""
                                $temp.Id = $array[2]
                                $temp.State = $array[3]
                                $temp.IdleTime = $array[4]
                                try
                                {
                                    $temp.LogonTime = Get-Date $($array[5] + " " + $array[6] + " " + $array[7]) -ErrorAction stop
                                }
                                catch
                                {
                                    $temp.LogonTime = ($array[5] + " " + $array[6] + " " + $array[7]).trim()
                                }
                            }
                            else{
                                $temp.SessionName = $array[2]
                                $temp.Id = $array[3]
                                $temp.State = $array[4]
                                $temp.IdleTime = $array[5]
                                try
                                {
                                    $temp.LogonTime = Get-Date $($array[6] + " " + $array[7] + " " + $array[8]) -ErrorAction stop
                                }
                                catch
                                {
                                    $temp.LogonTime = ($array[6] + " " + $array[7] + " " + $array[8]).trim()
                                }
                            }
                        }

                        #if specified, parse idle time to timespan
                        if($parseIdleTime){
                            $string = $temp.idletime
                
                            #quick function to handle minutes or hours:minutes
                            function Convert-ShortIdle {
                                param($string)
                                if($string -match "\:"){
                                    [timespan]$string
                                }
                                else{
                                    New-TimeSpan -Minutes $string
                                }
                            }
                
                            #to the left of + is days
                            if($string -match "\+"){
                                $days = New-TimeSpan -days ($string -split "\+")[0]
                                $hourMin = Convert-ShortIdle ($string -split "\+")[1]
                                $temp.idletime = $days + $hourMin
                            }
                            #. means less than a minute
                            elseif($string -like "." -or $string -like "none"){
                                $temp.idletime = [timespan]"0:00"
                            }
                            #hours and minutes
                            else{
                                $temp.idletime = Convert-ShortIdle $string
                            }
                        }
                
                        #Output the result
                        return $temp
                    }
                }
            }            
            else
            {
                Write-Warning "'$computer': No sessions found"
                return $false
            }
        }
    }
}

#define the WPF XML

$inputXML = '<Window x:Class="RDPAssist.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:RDPAssist"
        mc:Ignorable="d"
        Title="RDP Assist" Height="302.528" Width="243.944">
    <Grid Background="White">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="476*"/>
            <ColumnDefinition Width="5*"/>
        </Grid.ColumnDefinitions>
        <Label Name="Label_EnterInThePCName" Content="Enter in the PC name" HorizontalAlignment="Left" Margin="10,14,0,0" VerticalAlignment="Top"/>
        <TextBox Name="Textbox_PCName" HorizontalAlignment="Left" Height="23" Margin="15,34,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="196">
            <TextBox.Effect>
                <DropShadowEffect Color="#FFEAEAEA" ShadowDepth="2"/>
            </TextBox.Effect>
        </TextBox>
        <Button Name="Button_Locate" IsEnabled="False" Content="Locate" HorizontalAlignment="Left" Margin="15,62,0,0" VerticalAlignment="Top" Width="107" Background="#FFF7F7F7" BorderBrush="#FFBFBEBE" Height="25">
            <Button.Effect>
                <DropShadowEffect Color="#FFEAEAEA" ShadowDepth="2"/>
            </Button.Effect>
        </Button>
        <Label Name="Label_ActiveUserList" Content="Active User List" HorizontalAlignment="Left" Margin="11,104,0,0" VerticalAlignment="Top"/>
        <ListBox Name="Listbox_UserList" HorizontalAlignment="Left" Height="86" Margin="15,125,0,0" VerticalAlignment="Top" Width="196">
            <ListBox.Effect>
                <DropShadowEffect Color="#FFEAEAEA" ShadowDepth="2"/>
            </ListBox.Effect>
        </ListBox>
        <Button Name="Button_Connect" IsEnabled="False" Content="Connect" HorizontalAlignment="Left" Margin="15,216,0,0" VerticalAlignment="Top" Width="107" Background="#FFF7F7F7" BorderBrush="#FFBFBEBE" Height="25">
            <Button.Effect>
                <DropShadowEffect Color="#FFEAEAEA" ShadowDepth="2"/>
            </Button.Effect>
        </Button>
        <Label Name="Label_PCNotFound" Content="PC Not Found!" HorizontalAlignment="Left" Margin="127,62,0,0" VerticalAlignment="Top" Height="25" FontWeight="SemiBold" Foreground="#FFB44949" Visibility="Hidden"/>
    </Grid>
</Window>'
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
  
$xaml.SelectNodes("//*[@Name]") | %{
    try {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop}
    catch{throw}
    }

#when the button is clicked to search for active sessions, it will clear out any previous users listed
$WPFButton_Locate.Add_Click({
    $WPFListbox_UserList.Items.Clear()
    Get-Sessions
})

#if a PC name is entered, and Return is hit, then it will get the sessions
$WPFTextbox_PCName.add_KeyDown({
    if ($_.Key -eq 'Return')
    {
        Get-Sessions
    }
})

#connect to the selected session
$WPFButton_Connect.Add_Click({
    #this button is only enabled if the search comes back with sessions and a session is actually clicked on
    $id = $Global:sessionList[$WPFListBox_userlist.selectedIndex].id
    $computer = $WPFTextbox_PCName.Text
    mstsc /shadow:$id /v:$computer /control
})

#enable the button when a session is selected
$WPFListBox_userlist.Add_SelectionChanged({ 
    #once a session is clicked, the Connect button is enabled
    $WPFButton_Connect.IsEnabled = $true
})


#essentially resets the form if text is changed.
#If PC1 is entered and sessions are found, and then you change the textbox to PC, it wouldn't be able to mstsc to the PC/Session2 since PC wasn't what was searched

$WPFTextbox_PCName.Add_TextChanged({
    $WPFLabel_PCNotFound.Visibility = 'hidden'
    $WPFListbox_UserList.Items.Clear()
    $WPFButton_Connect.IsEnabled = $false
    try {
        $Global:sessionList.Clear()
    } catch { echo 'Session List appears to not currently be an array' }
    if ($WPFTextbox_PCName.Text.Length -gt 0) {
        $WPFButton_Locate.IsEnabled = $true
    } else { $WPFButton_Locate.IsEnabled = $false }
})

$Form.ShowDialog() | Out-Null
