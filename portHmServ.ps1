Import-Module ActiveDirectory            

# This script helps in bulk import of hMailServer user accounts from Active Directory using Async GUI
# Provides complete control to sync accounts from specific Organization Unit or Group
# CSV/TXT File can be use to import existing AD accounts to hMailServer
# (C) 2015 mfahim provided under MS-LPL license - https://programmingpakistan.com

#The Active Directory domain to use
$domain = Get-ADDomain 

#Current Directory from which script is being executed
$cScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
 
#Synchronize multithreaded access to objects
$syncHashForm = [hashtable]::Synchronized(@{})
$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"         
$newRunspace.Open()
#Sharing Variables and Live Objects Between PowerShell Runspaces
$newRunspace.SessionStateProxy.SetVariable("syncHashForm",$syncHashForm)     
$newRunspace.SessionStateProxy.SetVariable("cScriptPathSVar",$cScriptPath)   
$newRunspace.SessionStateProxy.SetVariable("domainSVar",$domain)        
$psCmd = [PowerShell]::Create().AddScript({ 
 
#The base DN for accounts (DC=appanoxstudios,DC=com)
$dn = $domainSVar.DistinguishedName 

#The domain used for emails (appanoxstudios.com) 
$dnsroot = $domainSVar.DNSRoot; 

#Log file headers (columns)
$errors_file_headers = @("User","FullName","Message","HelpLink","At")
#Log file data array
$error_logs = @()

#Set Current Directory from cScriptPathSVar
$cScriptPath = $cScriptPathSVar

#WPF XAML GUI between the @" "@ 
$inputXML = @"
<Window x:Class="portHmServ.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:portHmServ"
        mc:Ignorable="d"
        ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen"
        Title="portHmServ - Import AD/CSV file users to hmailServer" Height="480" Width="910">

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height=".6*" />
            <RowDefinition Height=".6*" />
            <RowDefinition Height="1*" />
            <RowDefinition Height=".12*" />
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="510" />
            <ColumnDefinition Width="20" />
            <ColumnDefinition Width="380" />
        </Grid.ColumnDefinitions>
        <GroupBox x:Name="hmGroupBox" Grid.Row="0" Grid.Column="0" HorizontalAlignment="Left" Margin="20,10,0,0" VerticalAlignment="Top" Height="105" Width="485">
            <GroupBox.Header>
                <TextBlock FontWeight="Bold">HmailServer</TextBlock>
            </GroupBox.Header>
            <Grid HorizontalAlignment="Left" Height="77" Margin="10,0,-2,-12" VerticalAlignment="Top" Width="477">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="0.4*" />
                    <ColumnDefinition Width=".6*" />
                    <ColumnDefinition Width="0.3*" />
                    <ColumnDefinition Width=".6*" />
                    <ColumnDefinition Width=".2*" />
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="1*" />
                    <RowDefinition Height="1*" />
                </Grid.RowDefinitions>
                <Label x:Name="usernameLabel" Content="Username" HorizontalAlignment="left" Margin="0,10,0,0" Grid.Row="0" Grid.Column="0" VerticalAlignment="Top" Width="81" VerticalContentAlignment="Top" HorizontalContentAlignment="Right"/>
                <TextBox x:Name="usernameTextBox" HorizontalAlignment="Left" Height="26" Margin="0,10,0,0" Grid.Row="0" Grid.Column="1" TextWrapping="NoWrap" Text="Administrator" VerticalAlignment="Top" VerticalContentAlignment="Center" Width="120"/>
                <Label x:Name="passLabel" Content="Password" HorizontalAlignment="Right" Margin="0,10,5,0" Grid.Row="0" Grid.Column="2" VerticalAlignment="Top" VerticalContentAlignment="Top" HorizontalContentAlignment="Right" Width="63" />
                <PasswordBox x:Name="passTextBox" Password="" HorizontalAlignment="Left" Height="26" Margin="0,10,0,0" Grid.Row="0" Grid.Column="3" VerticalAlignment="Top" VerticalContentAlignment="Center" Width="120"/>
                <Label x:Name="hmailmaxSizeLabel" Content="MailBox (MB)" HorizontalAlignment="Right" Margin="0,10,10,0" Grid.Row="1" Grid.Column="0" VerticalAlignment="Top" Width="81" VerticalContentAlignment="Top"/>
                <TextBox x:Name="hmailmaxSizeTextBox" HorizontalAlignment="Left" Height="26" Margin="0,10,0,0" Grid.Row="1" Grid.Column="1" TextWrapping="NoWrap" Text="50" VerticalAlignment="Top" VerticalContentAlignment="Center" Width="120"/>
            </Grid>
        </GroupBox>
        <GroupBox x:Name="hmImportGroupBox" Grid.Row="1" Grid.Column="0" HorizontalAlignment="Left" Margin="20,10,0,0" VerticalAlignment="Top" Height="107" Width="485">
            <GroupBox.Header>
                <TextBlock FontWeight="Bold">HmailServer Import (AD)</TextBlock>
            </GroupBox.Header>
            <Grid HorizontalAlignment="Left" Height="75" Margin="10,0,-14,-12" VerticalAlignment="Top" Width="477">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="0.4*" />
                    <ColumnDefinition Width=".6*" />
                    <ColumnDefinition Width="0.3*" />
                    <ColumnDefinition Width=".7*" />
                    <ColumnDefinition Width=".3*" />
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="1*" />
                    <RowDefinition Height="1*" />
                </Grid.RowDefinitions>
                <Label x:Name="importFromLabel" Content="From" HorizontalAlignment="Right" Margin="0,10,10,0" Grid.Row="0" Grid.Column="0" VerticalAlignment="Top" HorizontalContentAlignment="Right" Width="67" VerticalContentAlignment="Top"/>
                <ComboBox x:Name="importFromComboBox" SelectedIndex="0" Height="26" Margin="0,10,0,0" Grid.Row="0" Grid.Column="1" HorizontalAlignment="Left" VerticalAlignment="top" VerticalContentAlignment="Center" Width="120">
                    <ComboBoxItem Name="adall">AD ALL</ComboBoxItem>
                    <ComboBoxItem Name="adou">AD OU</ComboBoxItem>
                    <ComboBoxItem Name="adgroup">AD Group</ComboBoxItem>
                    <ComboBoxItem Name="datafile">Data File</ComboBoxItem>
                </ComboBox>
                <Label x:Name="cusParamLabel" Content="Param" HorizontalAlignment="Right" Margin="0,10,5,0" Grid.Row="0" Grid.Column="2" VerticalAlignment="Top" Width="50" VerticalContentAlignment="Top" HorizontalContentAlignment="Right" Visibility="Hidden"/>
                <TextBox x:Name="cusParamTextBox" HorizontalAlignment="Left" Height="26" Margin="0,10,0,0" Grid.Row="0" Grid.Column="3" TextWrapping="NoWrap" Text="" VerticalAlignment="Top" VerticalContentAlignment="Center" Width="142" Visibility="Hidden"/>
                <ComboBox x:Name="cusParamComboBox" IsEditable="True" IsTextSearchEnabled="True" Height="26" Margin="0,10,0,0" Grid.Row="0" Grid.Column="3" HorizontalAlignment="Left" VerticalAlignment="top" VerticalContentAlignment="Center" Width="142" Visibility="Hidden"/>
                <Button x:Name="cusParamButton" HorizontalAlignment="Left" Height="26" Margin="0,10,0,0" Grid.Row="0" Grid.Column="4" Content="+" FontWeight="Bold" VerticalAlignment="Top" VerticalContentAlignment="Center" Width="30" Visibility="Hidden"/>
                <Label x:Name="mappingLabel" Content="Mapping" HorizontalAlignment="Right" Margin="0,10,10,0" Grid.Row="1" Grid.Column="0" VerticalAlignment="Top" HorizontalContentAlignment="Right" Width="67" VerticalContentAlignment="Top"/>                <ComboBox x:Name="mappingComboBox" SelectedIndex="0" Height="26" Margin="0,10,0,0" Grid.Row="1" Grid.Column="1" HorizontalAlignment="Left" VerticalAlignment="top" VerticalContentAlignment="Center" Width="120">
                    <ComboBoxItem Name="adsamaccn">Username</ComboBoxItem>
                    <ComboBoxItem Name="ademail">Email</ComboBoxItem>
                </ComboBox>
            </Grid>
        </GroupBox>
        <GroupBox x:Name="csvGroupBox" IsEnabled="False" Grid.Row="2" Grid.Column="0" HorizontalAlignment="Left" Margin="20,10,0,0" VerticalAlignment="Top" Height="148" Width="485">
            <GroupBox.Header>
                <TextBlock FontWeight="Bold">CSV to AD (Mapping)</TextBlock>
            </GroupBox.Header>
            <Grid HorizontalAlignment="Left" Height="120" Margin="10,0,-14,-75" VerticalAlignment="Top" Width="477">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="0.5*" />
                    <ColumnDefinition Width=".6*" />
                    <ColumnDefinition Width="0.5*" />
                    <ColumnDefinition Width=".7*" />
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="1*" />
                    <RowDefinition Height="1*" />
                    <RowDefinition Height="1*" />
                </Grid.RowDefinitions>
                <Label x:Name="samANLabel" Content="SamAccountName" HorizontalAlignment="Right" Margin="0,10,10,0" Grid.Row="0" Grid.Column="0" VerticalAlignment="Top" HorizontalContentAlignment="Right" Width="84" VerticalContentAlignment="Top"/>
                <TextBox x:Name="samANTextBox" Height="26" Margin="10,10,0,0" Grid.Row="0" Grid.Column="1" TextWrapping="NoWrap" HorizontalAlignment="Left" VerticalAlignment="top" VerticalContentAlignment="Center" Width="104"/>
                <Label x:Name="givNLabel" Content="GivenName" HorizontalAlignment="Right" Margin="0,10,5,0" Grid.Row="0" Grid.Column="2" VerticalAlignment="Top" Width="89" VerticalContentAlignment="Top" HorizontalContentAlignment="Right" Visibility="Visible"/>
                <TextBox x:Name="givNTextBox" HorizontalAlignment="Left" Height="26" Margin="10,10,0,0" Grid.Row="0" Grid.Column="3" TextWrapping="NoWrap" Text="" VerticalAlignment="Top" VerticalContentAlignment="Center" Width="104" Visibility="Visible"/>

                <Label x:Name="surNLabel" Content="Surname" HorizontalAlignment="Right" Margin="0,10,10,0" Grid.Row="1" Grid.Column="0" VerticalAlignment="Top" HorizontalContentAlignment="Right" Width="84" VerticalContentAlignment="Top"/>
                <TextBox x:Name="surNTextBox" Height="26" Margin="10,10,0,0" Grid.Row="1" Grid.Column="1" TextWrapping="NoWrap" HorizontalAlignment="Left" VerticalAlignment="top" VerticalContentAlignment="Center" Width="104"/>
                <Label x:Name="statusNLabel" Content="Enabled" HorizontalAlignment="Right" Margin="0,10,5,0" Grid.Row="1" Grid.Column="2" VerticalAlignment="Top" Width="89" VerticalContentAlignment="Top" HorizontalContentAlignment="Right" Visibility="Visible"/>
                <TextBox x:Name="enabledStatusNTextBox" HorizontalAlignment="Left" Height="26" Margin="10,10,0,0" Grid.Row="1" Grid.Column="3" VerticalAlignment="Top" VerticalContentAlignment="Center" Width="104" Visibility="Visible"/>

                <Label x:Name="emailALabel" Content="EmailAddress" HorizontalAlignment="Right" Margin="0,10,10,0" Grid.Row="2" Grid.Column="0" VerticalAlignment="Top" HorizontalContentAlignment="Right" Width="84" VerticalContentAlignment="Top"/>
                <TextBox x:Name="emailANTextBox" Height="26" Margin="10,10,0,0" Grid.Row="2" Grid.Column="1" TextWrapping="NoWrap" HorizontalAlignment="Left" VerticalAlignment="top" VerticalContentAlignment="Center" Width="104"/>

            </Grid>
        </GroupBox>

        <StackPanel Orientation="Horizontal" Grid.Column="1" Grid.Row="0" Grid.RowSpan="3" Margin="0,17,0,20">
            <Separator Style="{StaticResource {x:Static ToolBar.SeparatorStyleKey}}" Opacity="0" />
        </StackPanel>


        <Grid Grid.Row="0" Grid.Column="3" Grid.RowSpan="3">
            <Grid.RowDefinitions>
                <RowDefinition Height="0.92*" />
                <RowDefinition Height="0.92*" />
                <RowDefinition Height="1*" />
                <RowDefinition Height="1.1*" />
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="50" />
                <ColumnDefinition Width="100" />
                <ColumnDefinition Width="100" />
                <ColumnDefinition Width="100" />
            </Grid.ColumnDefinitions>
            <Label x:Name="importForProLabel" Content="AD/Files" HorizontalAlignment="Right" Margin="0,10,10,0" Grid.Row="0" Grid.Column="0" VerticalAlignment="Top" HorizontalContentAlignment="Center" Width="84" VerticalContentAlignment="Top">
                <Label.LayoutTransform>
                    <RotateTransform Angle="-90" />
                </Label.LayoutTransform>
            </Label>
            <Grid Grid.Row="0" Grid.Column="1" Canvas.Left="90" Canvas.Top="50" Width="80" Height="auto" Margin="10">
                <Ellipse Stroke="Black" StrokeThickness="1" />
                <StackPanel
                    Orientation="Vertical" VerticalAlignment="Center">
                    <TextBlock x:Name="usersADFBlock" Text="0" HorizontalAlignment="Center" FontSize="30"></TextBlock>
                    <TextBlock Text="users" HorizontalAlignment="Center"></TextBlock>
                </StackPanel>
            </Grid>
            <Grid Grid.Row="0" Grid.Column="2" Canvas.Left="90" Canvas.Top="50" Width="80" Height="auto" Margin="10">
                <Ellipse Stroke="Black" StrokeThickness="1" />
                <StackPanel
                    Orientation="Vertical" VerticalAlignment="Center">
                    <TextBlock x:Name="ousADFBlock" Text="0" HorizontalAlignment="Center" FontSize="30"></TextBlock>
                    <TextBlock Text="ou's" HorizontalAlignment="Center"></TextBlock>
                </StackPanel>
            </Grid>
            <Grid Grid.Row="0" Grid.Column="3" Canvas.Left="90" Canvas.Top="50" Width="80" Height="auto" Margin="10">
                <Ellipse Stroke="Black" StrokeThickness="1" />
                <StackPanel
                    Orientation="Vertical" VerticalAlignment="Center">
                    <TextBlock x:Name="groupsADFBlock" Text="0" HorizontalAlignment="Center" FontSize="30"></TextBlock>
                    <TextBlock Text="groups" HorizontalAlignment="Center"></TextBlock>
                </StackPanel>
            </Grid>


            <Label x:Name="hmailServerProLabel" Content="hmailServer" HorizontalAlignment="Right" Margin="0,10,10,0" Grid.Row="1" Grid.Column="0" VerticalAlignment="Top" HorizontalContentAlignment="Center" Width="84" VerticalContentAlignment="Top">
                <Label.LayoutTransform>
                    <RotateTransform Angle="-90" />
                </Label.LayoutTransform>
            </Label>
            <Grid Grid.Row="1" Grid.Column="1" Canvas.Left="90" Canvas.Top="50" Width="80" Height="auto" Margin="10">
                <Ellipse Stroke="Black" StrokeThickness="1" />
                <StackPanel
                    Orientation="Vertical" VerticalAlignment="Center">
                    <TextBlock x:Name="accountsHmBlock" Text="0" HorizontalAlignment="Center" FontSize="30"></TextBlock>
                    <TextBlock Text="accounts" HorizontalAlignment="Center"></TextBlock>
                </StackPanel>
            </Grid>
            <Grid Grid.Row="1" Grid.Column="2" Canvas.Left="90" Canvas.Top="50" Width="80" Height="auto" Margin="10">
                <Ellipse Stroke="Black" StrokeThickness="1" />
                <StackPanel
                    Orientation="Vertical" VerticalAlignment="Center">
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                        <TextBlock x:Name="mailboxSizeHmBlock" Text="0" HorizontalAlignment="Center" FontSize="30"></TextBlock>
                        <TextBlock x:Name="mailboxSizeHmBlockUnit" Text="/MB" HorizontalAlignment="Left" VerticalAlignment="Center" FontSize="8"></TextBlock>
                    </StackPanel>
                    <TextBlock Text="mailbox" HorizontalAlignment="Center"></TextBlock>
                </StackPanel>
            </Grid>

            <StackPanel Orientation="Vertical" Grid.Row="2" Grid.ColumnSpan="4"  Margin="30,0,10,5" VerticalAlignment="Center">
                <ProgressBar x:Name="importProBar" Value="0" Height="10" Visibility="Hidden"></ProgressBar>
                <TextBlock x:Name="importValueProBar" Text="" HorizontalAlignment="Right" FontSize="10" Height="12" ></TextBlock>
            </StackPanel>

            <TextBlock x:Name="developerInfo" Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="4" FontSize="8" TextAlignment="Right" Text="© 2015 mfahim provided under MS-LPL license - https://programmingpakistan.com
" Margin="5,0,10,5"/>
            <CheckBox  x:Name="errorLogsCheckBox" Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2" Content="Error Logs as CSV" Margin="25,40,10,5"
VerticalAlignment="Center"></CheckBox>
            <Button x:Name="startButton" Grid.Row="3" Grid.Column="2" Content="Start" Margin="5,40,10,5" 
VerticalAlignment="Center"></Button>
            <Button x:Name="cancelButton" Grid.Row="3" Grid.Column="3" Content="Cancel" Margin="5,40,10,5" 
VerticalAlignment="Center"></Button>
        </Grid>

        <DockPanel Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="3">
            <StatusBar DockPanel.Dock="Bottom" BorderThickness="1" BorderBrush="#FFB6BDC5">
                <StatusBar.ItemsPanel>
                    <ItemsPanelTemplate>
                        <Grid Margin="10,0,30,0">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="100" />
                                <ColumnDefinition Width="Auto" />
                                <ColumnDefinition Width="*" />
                                <ColumnDefinition Width="Auto" />
                                <ColumnDefinition Width="*" />
                                <ColumnDefinition Width="Auto" />
                                <ColumnDefinition Width="Auto" />
                            </Grid.ColumnDefinitions>
                        </Grid>
                    </ItemsPanelTemplate>
                </StatusBar.ItemsPanel>
                <StatusBarItem>
                    <TextBlock x:Name="processItemSBarItem" Foreground="#a94442" Text="" />
                </StatusBarItem>
                <Separator Grid.Column="1" Style="{StaticResource {x:Static ToolBar.SeparatorStyleKey}}" Opacity="1" />
                <StatusBarItem Grid.Column="2">
                    <TextBlock x:Name="itemDetailsSBarItem" Foreground="#a94442" Text="" />
                </StatusBarItem>
                <Separator Grid.Column="3" Style="{StaticResource {x:Static ToolBar.SeparatorStyleKey}}" Opacity="1" />
                <StatusBarItem Grid.Column="4">
                    <TextBlock x:Name="distNameDetailsSBarItem" Text="" />
                </StatusBarItem>
                <Separator Grid.Column="5" Style="{StaticResource {x:Static ToolBar.SeparatorStyleKey}}" Opacity="1" />
                <StatusBarItem Grid.Column="6">
                    <TextBlock x:Name="domainDetailsSBarItem" Text="" />
                </StatusBarItem>
            </StatusBar>
        </DockPanel>
    </Grid>
</Window>
"@       
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML

#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$syncHashForm=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}

#Create hooks to each named object in the WPF XAML   
$XAML.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $syncHashForm.FindName($_.Name)}

#Digital Storage (Unit conversion)
Function Format-MailBoxSize {
param ([decimal]$Type,
       [string]$out
)
     # MB to Bytes Conversion
     $Type = $Type * (1024 * 1024.0)
     if($out -eq "v"){
         If ($Type -ge 1GB) { return [string]::Format("{0:N1}", $Type / 1GB)}
         ElseIf ($Type -ge 1MB) { return [string]::Format("{0:0}", $Type / 1MB)}
         #ElseIf ($Type -ge 1KB) {[string]::Format("{0:0} KB ", $Type / 1KB)}
         #ElseIf ($Type -gt 0) {[string]::Format("{0:0} Bytes", $Type)}
         Else { return "0"}
     }else{
        If ($Type -ge 1GB) { return "/GB"}
        ElseIf ($Type -ge 1MB) { return "/MB"}
         #ElseIf ($Type -ge 1KB) {[string]::Format("{0:0} KB ", $Type / 1KB)}
         #ElseIf ($Type -gt 0) {[string]::Format("{0:0} Bytes", $Type)}
         Else { return "/MB"}
     }
}

#Add AD User Account to hMailServer
Function Add-UserToHmailServer {
param([parameter(Mandatory=$true)] $AdUser,
      [parameter(Mandatory=$true)] $hmdom)
    $importFrom = $WPFimportFromComboBox.SelectedItem.Content.ToString()
    $mappingAD = $WPFmappingComboBox.SelectedItem.Content.ToString()
    $hmact = $hmdom.Accounts.Add() 
    
    #Create Account using CSV File
    if($importFrom -eq "Data File"){
        # Create hmailserver accounts base on AD Username or Email
        if($mappingAD -eq "Username"){ 
            $AdUserAddress = Invoke-Expression "`$AdUser.$($WPFsamANTextBox.Text)"
            $hmact.Address = "$($AdUserAddress)@$($dnsroot)"
        }else{
            $hmact.Address = Invoke-Expression "`$AdUser.$($WPFemailANTextBox.Text)"
        }
        $AdUsersStatus = Invoke-Expression "`$AdUser.$($WPFenabledStatusNTextBox.Text)" 
        $hmact.Active = [System.Convert]::ToBoolean($AdUsersStatus)
        $hmact.IsAD =  [System.Convert]::ToBoolean($AdUsersStatus)
        $hmact.MaxSize = $WPFhmailmaxSizeTextBox.Text
        $hmact.ADDomain = $dnsroot
        $hmact.ADUsername = Invoke-Expression "`$AdUser.$($WPFsamANTextBox.Text)"
        $hmact.PersonFirstName = Invoke-Expression "`$AdUser.$($WPFgivNTextBox.Text)"
        $hmact.PersonLastName = Invoke-Expression "`$AdUser.$($WPFsurNTextBox.Text)"  
    
    #Filling attributes for the email account based on current information, including AD integration for the password 
    }else{
        #Create hmailserver accounts base on AD Username or Email
        if($mappingAD -eq "Username"){ 
            $hmact.Address = "$($AdUser.SamAccountName)@$($dnsroot)"
        }else{
            $hmact.Address = $($AdUser.Mail)
        }
        $hmact.IsAD = $AdUser.Enabled  
        $hmact.Active = $AdUser.Enabled  
        $hmact.MaxSize = $WPFhmailmaxSizeTextBox.Text
        $hmact.ADDomain = $dnsroot
        $hmact.ADUsername = $AdUser.SamAccountName
        $hmact.PersonFirstName = $AdUser.GivenName
        $hmact.PersonLastName = $AdUser.Surname
    }
    
    #Error creating hMailServer Account
    try{
        $hmact.save() 
        return $true
    } Catch {
        if($importFrom -eq "Data File"){
            $ErrorUser = Invoke-Expression "`$AdUser.$($WPFsamANTextBox.Text)"
        }else{
            $ErrorUser = $AdUser.SamAccountName
        }
        $FailedFullName = $_.Exception.GetType().FullName
        $ErrorMessage = $_.Exception.Message
        $ErrorHelpLink = $_.Exception.HelpLink
        $ErrorAt = (Get-Date).ToString('MM/dd/yyyy hh:mm:ss tt')
        $properties  = @{User=$ErrorUser;FullName=$FailedFullName;Message=$ErrorMessage;HelpLink=$ErrorHelpLink;At=$ErrorAt}
        $obj = New-Object –TypeName PSObject –Prop $properties
        $global:error_logs += $obj

        #Set Error Adding User to StatusBar
        $WPFprocessItemSBarItem.Dispatcher.Invoke("Background",[action]{$WPFprocessItemSBarItem.Text = $ErrorUser})
        $WPFitemDetailsSBarItem.Dispatcher.Invoke("Background",[action]{$WPFitemDetailsSBarItem.Text = $ErrorMessage})

        #On Error return false
        return $false
    }
}

#Add AD User Accounts to hMailServer and control gui progress  
Function Get-UserQryGroupBy {
param(
    [parameter(Mandatory=$true)] $UsersSet
)
    $importUsersBy = $WPFimportFromComboBox.SelectedItem.Content.ToString()
    #List to hold all Unique Groups
    $Groups = New-Object System.Collections.ArrayList
    $UsersListT = $UsersSet

    $MailBoxSize = [decimal]$WPFhmailmaxSizeTextBox.Text

    #Creates a COM object against the hMailServer API 
    $hm = New-Object -ComObject hMailServer.Application 
    $hm.Authenticate($WPFusernameTextBox.Text, $WPFpassTextBox.Password) | Out-Null 
    $hmdom = $hm.Domains.ItemByName($dnsroot) 
    $accountsCreatedCount = 0

    #Check for Progress Bar Visibility
    $CStateWPFimportProBar = $WPFimportProBar.Visibility 
    if($CStateWPFimportProBar -eq "Hidden"){
        $WPFimportProBar.Visibility = "Visible"
    }

    foreach ($user in $UsersSet){

        $usersAddStatus = Add-UserToHmailServer $user $hmdom

        #Getting user AD account details 
        if($importUsersBy -eq "Data File"){
            $UserName = Invoke-Expression "`$user.$($WPFsamANTextBox.Text)"
        } else {
            $UserName = $user.SamAccountName
        }

        #Getting each User AD Groups (memberOf)
        $GroupsSet = Get-ADPrincipalGroupMembership $UserName | Select-Object name
        foreach ( $Group in $GroupsSet){
            if($Groups -notcontains  $Group.Name){
                $Groups.Add($Group.Name)
                $WPFgroupsADFBlock.Dispatcher.Invoke("Background",[action]{$WPFgroupsADFBlock.Text = $Groups.Count})
            }
        }

        #Update UI AD Users Count
        $WPFusersADFBlock.Dispatcher.Invoke("Background",[action]{$WPFusersADFBlock.Text = ($UsersSet | Measure-Object).Count})

        #Progress Indicator
        $i++
        $WPFimportProBar.Dispatcher.Invoke("Background",[action]{
                $perc = (($i / $UsersSet.Count)  * 100)
                $WPFimportProBar.Value = $perc
                $WPFimportValueProBar.Text = "{0:N0}%" -f $perc
        })

        #Update hMailserver accounts creation progress on success
        if($usersAddStatus -eq $true){
            $accountsCreatedCount++
            $WPFmailboxSizeHmBlock.Dispatcher.Invoke("Background",[action]{$WPFmailboxSizeHmBlock.Text = Format-MailBoxSize $MailBoxSize 'v'})
            $WPFmailboxSizeHmBlockUnit.Dispatcher.Invoke("Background",[action]{$WPFmailboxSizeHmBlockUnit.Text = Format-MailBoxSize $MailBoxSize 'u' })
            $WPFaccountsHmBlock.Dispatcher.Invoke("Background",[action]{$WPFaccountsHmBlock.Text = $accountsCreatedCount })
            
            $MailBoxSize = $MailBoxSize + [decimal]$WPFhmailmaxSizeTextBox.Text
        }

    }
    #Get count of all Users Unique OU's 
    $OUs = $UsersListT | select @{l='OU';e={$_.DistinguishedName.split(',')[1].split('=')[1]}} -Unique | measure
    #Update UI AD OU's Count
    $WPFousADFBlock.Dispatcher.Invoke("Background",[action]{$WPFousADFBlock.Text = $OUs.Count})

    #Export Error Logs to Csv on Completion
    if($WPFerrorLogsCheckBox.IsChecked -eq $true){
        $error_logs | Export-Csv -path "$cScriptPath\error_logs.csv" -NoTypeInformation -Append
    }
    #Clear Logs on Completion
    $global:error_logs = @()
}


$WPFstartButton.Add_Click({
    #Reset UI Elements
    $WPFgroupsADFBlock.Dispatcher.Invoke("Background",[action]{$WPFgroupsADFBlock.Text = "0"})
    $WPFusersADFBlock.Dispatcher.Invoke("Background",[action]{$WPFusersADFBlock.Text = "0"})
    $WPFousADFBlock.Dispatcher.Invoke("Background",[action]{$WPFousADFBlock.Text = "0"})
    $WPFmailboxSizeHmBlock.Dispatcher.Invoke("Background",[action]{$WPFmailboxSizeHmBlock.Text = "0"})
    $WPFmailboxSizeHmBlockUnit.Dispatcher.Invoke("Background",[action]{$WPFmailboxSizeHmBlockUnit.Text = "/MB" })
    $WPFaccountsHmBlock.Dispatcher.Invoke("Background",[action]{$WPFaccountsHmBlock.Text = "0" })
    
    $command = ''

    $importUsersBy = $WPFimportFromComboBox.SelectedItem.Content.ToString()
    $CValWPFcusParamComboBoxSBtn = $WPFcusParamComboBox.Text.ToString()
    $mappingAD = $WPFmappingComboBox.SelectedItem.Content.ToString()

    #Options for importing AD Users to hMailServer
    if($importUsersBy -eq "AD ALL") { $command = "Get-ADUser -Filter *" } 
    Elseif($importUsersBy -eq "AD OU") { $command = "Get-ADUser -Filter * -SearchBase `"$CValWPFcusParamComboBoxSBtn`"" } `
    ElseIf($importUsersBy -eq "AD Group")  { $command = "Get-ADGroupMember -identity `"$CValWPFcusParamComboBoxSBtn`" -Recursive | Get-ADUser" }
    ElseIf($importUsersBy -eq "Data File") { $command = "Import-CSV '" + $WPFcusParamTextBox.Text + "'" }

    #Getting user AD account details 
    if($importUsersBy -eq "Data File"){
        $resultSet = Invoke-Expression "$command" 
    } else {
        #Exclude all Users without an Email Address - Mapping Email
        if($mappingAD -eq "Email"){
                $resultSet = Invoke-Expression "$command -Properties SamAccountName, GivenName, Surname, Enabled, Mail, DistinguishedName | Select SamAccountName, GivenName, Surname, Enabled, Mail, DistinguishedName | where {`$_.Mail -ne `$Null}"
        }else{
             $resultSet = Invoke-Expression "$command -Properties SamAccountName, GivenName, Surname, Enabled, Mail, DistinguishedName | Select SamAccountName, GivenName, Surname, Enabled, Mail, DistinguishedName"
        }
    }

    #Pass Users Query Resultset to Get-UserQryGroupBy
    Get-UserQryGroupBy -UsersSet $resultSet
})

$WPFcancelButton.Add_Click({
    #Export Error Logs to Csv on Cancel
    if($WPFerrorLogsCheckBox.IsChecked -eq $true){
        $error_logs | Export-Csv -path "$cScriptPath\error_logs.csv" -NoTypeInformation -Append
    }
    $syncHashForm.Close()
})

Function Get-FileName { 
     Param(
     [Parameter(Mandatory=$true)] [string]$initialDirectory,
     [Parameter(Mandatory=$true)] [string]$title
     )

     #Add-Type -AssemblyName System.Windows.Forms 
     $OpenFileDialog = New-Object Microsoft.Win32.OpenFileDialog
     $OpenFileDialog.Title = $title
     $OpenFileDialog.initialDirectory = [Environment]::GetFolderPath($initialDirectory)
     $OpenFileDialog.filter = "CSV and TXT Files (*.csv,*.txt)|*.csv;*.txt"
     #Dialog box opens on front 
     $OpenFileDialog.ShowHelp = $true
     $OpenFileDialog.ShowDialog() | Out-Null
     return $OpenFileDialog.filename
} #end function Get-FileName


$WPFcusParamButton.Add_Click({
    $WPFcusParamTextBox.Dispatcher.Invoke("Background",[action]{
        $WPFcusParamTextBox.Text = Get-FileName -initialDirectory "Desktop" -title "Select the Active Directory Users CSV File"
    })

})

Function WPFcsvUIControlsState {
    Param(
    [Parameter(Mandatory=$true)] [string]$cusParamLabelVal,
    [Parameter(Mandatory=$true)] [string]$cusParamComboBoxVal,
    [Parameter(Mandatory=$true)] [string]$cusParamTextBoxVal,
    [Parameter(Mandatory=$true)] [string]$cusParamButtonVal
    )

    $WPFcusParamLabel.Visibility = $cusParamLabelVal
    $WPFcusParamComboBox.Visibility = $cusParamComboBoxVal
    $WPFcusParamTextBox.Visibility = $cusParamTextBoxVal
    $WPFcusParamButton.Visibility = $cusParamButtonVal
}

#Source for importing hMailServer User Accounts 
$WPFimportFromComboBox.add_SelectionChanged({

    $CValWPFcusParamTextBox = $WPFimportFromComboBox.SelectedItem.Content.ToString()
    $WPFcusParamComboBox.Items.Clear()

    $CStateWPFcsvGroupBox = $WPFcsvGroupBox.IsEnabled  
    if($CValWPFcusParamTextBox -eq "Data File"){
        if(!$CStateWPFcsvGroupBox){
            $WPFcsvGroupBox.IsEnabled = $True
         }
         WPFcsvUIControlsState "Visible" "Hidden" "Visible" "Visible"
         $WPFcusParamTextBox.IsReadOnly = $True
         $WPFcusParamLabel.Content = "File"
    }elseif($CValWPFcusParamTextBox -eq "AD ALL"){
        if($CStateWPFcsvGroupBox){
            $WPFcsvGroupBox.IsEnabled = $False
         }
         WPFcsvUIControlsState "Hidden" "Hidden" "Hidden" "Hidden"
    }elseif($CValWPFcusParamTextBox -eq "AD OU"){
        if($CStateWPFcsvGroupBox){
            $WPFcsvGroupBox.IsEnabled = $False
         }
         WPFcsvUIControlsState "Visible" "Visible" "Hidden" "Hidden"
         $WPFcusParamLabel.Content = "OU"

         $cmdlets = @(Get-ADOrganizationalUnit -Filter * | Select -ExpandProperty DistinguishedName)
         $WPFcusParamComboBox.Dispatcher.Invoke("Background",[action]{$WPFcusParamComboBox.itemsSource = $cmdlets})

    }elseif($CValWPFcusParamTextBox -eq "AD Group"){
        if($CStateWPFcsvGroupBox){
            $WPFcsvGroupBox.IsEnabled = $False
        }
        WPFcsvUIControlsState "Visible" "Visible" "Hidden" "Hidden"
        $WPFcusParamLabel.Content = "Group"


        $cmdlets = @(get-adgroup -Filter * | Select -ExpandProperty Name)
        $WPFcusParamComboBox.Dispatcher.Invoke("Background",[action]{$WPFcusParamComboBox.itemsSource = $cmdlets})
}

#Update ComboBox List on TextInput
<#$WPFcusParamComboBox.Add_PreviewTextInput({
    $CValWPFcusParamTextBox = $WPFimportFromComboBox.SelectedItem.Content.ToString()
    $CValWPFcusParamComboBox = $WPFcusParamComboBox.Text
    $WPFcusParamComboBox.Items.Clear()

    if($CValWPFcusParamTextBox -eq "AD OU" -and $CValWPFcusParamComboBox -ne "" ){
        $cmdlets = @(Get-ADOrganizationalUnit -Filter "name -like `"`*$($CValWPFcusParamComboBox)`*`"" | Select -ExpandProperty DistinguishedName)
        $WPFcusParamComboBox.Dispatcher.Invoke("Background",[action]{$WPFcusParamComboBox.itemsSource = $cmdlets})
    }elseif($CValWPFcusParamTextBox -eq "AD Group" -and $CValWPFcusParamComboBox -ne ""){
        $cmdlets = @(get-adgroup -Filter "name -like `"`*$($CValWPFcusParamComboBox)`*`"" | Select -ExpandProperty Name)
        $WPFcusParamComboBox.Dispatcher.Invoke("Background",[action]{$WPFcusParamComboBox.itemsSource = $cmdlets})
    }elseif($CValWPFcusParamTextBox -eq "AD OU" -and $CValWPFcusParamComboBox -eq ""){
        $cmdlets = @(Get-ADOrganizationalUnit -Filter * | Select -ExpandProperty DistinguishedName)
        $WPFcusParamComboBox.Dispatcher.Invoke("Background",[action]{$WPFcusParamComboBox.itemsSource = $cmdlets})
    }elseif($CValWPFcusParamTextBox -eq "AD Group" -and $CValWPFcusParamComboBox -eq ""){
        $cmdlets = @(get-adgroup -Filter * | Select -ExpandProperty Name)
        $WPFcusParamComboBox.Dispatcher.Invoke("Background",[action]{$WPFcusParamComboBox.itemsSource = $cmdlets})
    }
})#>

#Update ComboBox List on KeyDown (include missing PreviewTextInput events)
$WPFcusParamComboBox.Add_PreviewKeyDown({
    $CValWPFcusParamTextBox = $WPFimportFromComboBox.SelectedItem.Content.ToString()
    $CValWPFcusParamComboBox = $WPFcusParamComboBox.Text
    $WPFcusParamComboBox.Items.Clear()
     
    if($CValWPFcusParamTextBox -eq "AD OU" -and $CValWPFcusParamComboBox -ne "" ){
        $cmdlets = @(Get-ADOrganizationalUnit -Filter "name -like `"`*$($CValWPFcusParamComboBox)`*`"" | Select -ExpandProperty DistinguishedName)
        $WPFcusParamComboBox.Dispatcher.Invoke("Background",[action]{$WPFcusParamComboBox.itemsSource = $cmdlets})
    }elseif($CValWPFcusParamTextBox -eq "AD Group" -and $CValWPFcusParamComboBox -ne ""){
        $cmdlets = @(get-adgroup -Filter "name -like `"`*$($CValWPFcusParamComboBox)`*`"" | Select -ExpandProperty Name)
        $WPFcusParamComboBox.Dispatcher.Invoke("Background",[action]{$WPFcusParamComboBox.itemsSource = $cmdlets})
    }elseif($CValWPFcusParamTextBox -eq "AD OU" -and $CValWPFcusParamComboBox -eq ""){
        $cmdlets = @(Get-ADOrganizationalUnit -Filter * | Select -ExpandProperty DistinguishedName)
        $WPFcusParamComboBox.Dispatcher.Invoke("Background",[action]{$WPFcusParamComboBox.itemsSource = $cmdlets})
    }elseif($CValWPFcusParamTextBox -eq "AD Group" -and $CValWPFcusParamComboBox -eq ""){
        $cmdlets = @(get-adgroup -Filter * | Select -ExpandProperty Name)
        $WPFcusParamComboBox.Dispatcher.Invoke("Background",[action]{$WPFcusParamComboBox.itemsSource = $cmdlets})
    }
})

}) 

Function Show-Form{
    $WPFdeveloperInfo.Text = "© $((Get-Date).year) mfahim provided under MS-LPL license - https://programmingpakistan.com"
    #Setting Domain Details to StatusBar
    $WPFdistNameDetailsSBarItem.Text = $dn
    $WPFdomainDetailsSBarItem.Text = $dnsroot

    $syncHashForm.ShowDialog() | out-null
    $syncHashForm.Erorr = $Error
}
Show-Form

})
#Thread Synchronization
$psCmd.Runspace = $newRunspace
$data = $psCmd.BeginInvoke()