param([switch]$SelfTest, [string]$PreviewPath = '', [int]$PreviewTab = 0, [switch]$FeatureTest)
$ErrorActionPreference = 'Stop'
if ($FeatureTest -and -not $SelfTest) {throw 'FeatureTest requires SelfTest and fictional test data.'}
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$script:appRoot = $PSScriptRoot
$script:projectRoot = Split-Path -Parent $script:appRoot
$script:scriptsRoot = Join-Path $script:projectRoot 'scripts'
try {
    $runtime = & (Join-Path $script:scriptsRoot 'Set up runtime.ps1')
    $script:python = $runtime.PythonPath
    $script:taskName = & (Join-Path $script:scriptsRoot 'Get task name.ps1')
} catch {
    if ($SelfTest) { throw }
    [System.Windows.MessageBox]::Show($_.Exception.Message,'Setup needed') | Out-Null
    exit 1
}
$script:job = $null
$script:busy = $false
$script:state = $null
$script:autoSyncBlocked = $false
$script:updatingYears = $false
. (Join-Path $PSScriptRoot 'Import dialogs.ps1')
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="TNG Expense Agent" Width="1060" Height="760" MinWidth="900" MinHeight="650" WindowStartupLocation="CenterScreen" Background="#F4F6FA" FontFamily="Segoe UI" FontSize="14">
 <Window.Resources>
  <Style TargetType="Button"><Setter Property="Padding" Value="15,9"/><Setter Property="Margin" Value="0,0,10,0"/><Setter Property="Background" Value="#FFFFFF"/><Setter Property="BorderBrush" Value="#CDD6E2"/><Setter Property="Foreground" Value="#18324F"/><Setter Property="Cursor" Value="Hand"/></Style>
  <Style TargetType="TextBox"><Setter Property="Padding" Value="9,7"/><Setter Property="BorderBrush" Value="#CDD6E2"/></Style>
  <Style TargetType="PasswordBox"><Setter Property="Padding" Value="9,7"/><Setter Property="BorderBrush" Value="#CDD6E2"/></Style>
  <Style TargetType="DataGrid"><Setter Property="AutoGenerateColumns" Value="False"/><Setter Property="IsReadOnly" Value="True"/><Setter Property="CanUserAddRows" Value="False"/><Setter Property="GridLinesVisibility" Value="Horizontal"/><Setter Property="HorizontalGridLinesBrush" Value="#E8EDF3"/><Setter Property="RowHeight" Value="34"/><Setter Property="HeadersVisibility" Value="Column"/><Setter Property="BorderThickness" Value="0"/><Setter Property="AlternatingRowBackground" Value="#F4F7FB"/><Setter Property="SelectionMode" Value="Single"/></Style>
 </Window.Resources>
 <Grid Margin="28,22" Background="#F4F6FA">
  <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
  <StackPanel Grid.Row="0" Margin="0,0,0,20">
   <TextBlock Text="TNG Expense Agent" Foreground="#18324F" FontSize="28" FontWeight="SemiBold"/>
   <TextBlock Text="Add an expense below or upload a statement. Excel saves entries by date into the correct month." Foreground="#637286" Margin="0,6,0,0"/>
  </StackPanel>
  <Border Grid.Row="1" Background="#18324F" CornerRadius="8" Padding="22,16" Margin="0,0,0,20">
   <Grid><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
    <StackPanel><TextBlock Text="EXPENSES IN SELECTED YEAR" Foreground="#C5D5E7" FontSize="11"/><TextBlock x:Name="CountLabel" Text="0" Foreground="White" FontSize="25" FontWeight="SemiBold" Margin="0,4,0,0"/></StackPanel>
    <StackPanel Grid.Column="1"><TextBlock Text="SELECTED YEAR TOTAL" Foreground="#C5D5E7" FontSize="11"/><TextBlock x:Name="TotalLabel" Text="RM 0.00" Foreground="White" FontSize="25" FontWeight="SemiBold" Margin="0,4,0,0"/></StackPanel>
    <StackPanel Grid.Column="2"><TextBlock Text="PROCESSING" Foreground="#C5D5E7" FontSize="11"/><TextBlock Text="On this PC" Foreground="White" FontSize="17" Margin="0,8,0,0"/></StackPanel>
   </Grid>
  </Border>
  <TabControl Grid.Row="2" x:Name="Tabs" Background="White" BorderBrush="#DCE3ED" Padding="18">
   <TabItem Header="  Expenses  ">
    <Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
     <WrapPanel Margin="0,0,0,14">
      <Button x:Name="UploadButton" Content="Upload PDF" Background="#2458A6" Foreground="White" BorderBrush="#2458A6"/>
      <Button x:Name="OpenButton" Content="Open Excel"/>
      <Button x:Name="RefreshButton" Content="Refresh Excel"/>
      <Button x:Name="DeleteExpenseButton" Content="Delete selected"/>
      <ComboBox x:Name="MonthFilter" MinWidth="160" Margin="5,0,0,0" VerticalContentAlignment="Center"/>
     </WrapPanel>
     <WrapPanel Grid.Row="1" Margin="0,0,0,12">
      <Button x:Name="PreviousYearButton" Content="&lt; Previous year"/>
      <ComboBox x:Name="YearFilter" Width="100" Margin="0,0,10,0" VerticalContentAlignment="Center"/>
      <Button x:Name="NextYearButton" Content="Next year &gt;"/>
      <TextBlock x:Name="YearFileLabel" Foreground="#637286" VerticalAlignment="Center"/>
     </WrapPanel>
     <TextBlock Grid.Row="2" Text="Purchases and outgoing transfers only. Reloads, incoming money and GO+ movements are excluded." Foreground="#637286" TextWrapping="Wrap" Margin="0,0,0,12"/>
     <Grid Grid.Row="3" Margin="0,0,0,16">
      <Grid.ColumnDefinitions><ColumnDefinition Width="155"/><ColumnDefinition Width="*"/><ColumnDefinition Width="120"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
      <StackPanel Margin="0,0,12,0"><TextBlock Text="Date" Margin="0,0,0,6"/><DatePicker x:Name="ExpenseDate" SelectedDateFormat="Short" Height="36"/></StackPanel>
      <StackPanel Grid.Column="1" Margin="0,0,12,0"><TextBlock Text="Description" Margin="0,0,0,6"/><TextBox x:Name="ExpenseDescription" MaxLength="300"/></StackPanel>
      <StackPanel Grid.Column="2" Margin="0,0,12,0"><TextBlock Text="Price (RM)" Margin="0,0,0,6"/><TextBox x:Name="ExpensePrice" MaxLength="12" ToolTip="For example, 12.50"/></StackPanel>
      <Button Grid.Column="3" x:Name="AddExpenseButton" Content="Save to Excel" VerticalAlignment="Bottom" Background="#2458A6" Foreground="White" BorderBrush="#2458A6" Margin="0"/>
     </Grid>
     <DataGrid Grid.Row="4" x:Name="TransactionsGrid">
      <DataGrid.Columns>
       <DataGridTextColumn Header="Date" Binding="{Binding date}" Width="115"/>
       <DataGridTextColumn Header="Description" Binding="{Binding display_description}" Width="*"/>
       <DataGridTextColumn Header="Price (RM)" Binding="{Binding amount}" Width="110"/>
       <DataGridTextColumn Header="Original description" Binding="{Binding description}" Width="230"/>
      </DataGrid.Columns>
     </DataGrid>
    </Grid>
   </TabItem>
   <TabItem Header="  Description rules  ">
    <Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
     <TextBlock Text="Save a merchant phrase and the name you want in Excel column B. Rules update saved and future transactions." TextWrapping="Wrap" Margin="0,0,0,15" Foreground="#637286"/>
     <Grid Grid.Row="1" Margin="0,0,0,12"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="130"/></Grid.ColumnDefinitions>
      <StackPanel Margin="0,0,12,0"><TextBlock Text="Merchant phrase" Margin="0,0,0,6"/><TextBox x:Name="PatternInput"/></StackPanel>
      <StackPanel Grid.Column="1" Margin="0,0,12,0"><TextBlock Text="Use this description" Margin="0,0,0,6"/><TextBox x:Name="ReplacementInput"/></StackPanel>
      <StackPanel Grid.Column="2"><TextBlock Text="Match" Margin="0,0,0,6"/><ComboBox x:Name="RuleMode" SelectedIndex="0" Padding="8"><ComboBoxItem Content="Contains"/><ComboBoxItem Content="Exact"/></ComboBox></StackPanel>
     </Grid>
     <WrapPanel Grid.Row="2" Margin="0,0,0,14"><Button x:Name="SaveRuleButton" Content="Save rule"/><Button x:Name="DeleteRuleButton" Content="Remove selected rule"/><TextBlock Text="Exact matches take priority. Otherwise, the first matching rule wins." FontSize="12" Foreground="#637286" VerticalAlignment="Center"/></WrapPanel>
     <DataGrid Grid.Row="3" x:Name="RulesGrid"><DataGrid.Columns><DataGridTextColumn Header="Merchant phrase" Binding="{Binding pattern}" Width="*"/><DataGridTextColumn Header="Description in Excel" Binding="{Binding replacement}" Width="*"/><DataGridTextColumn Header="Match" Binding="{Binding mode}" Width="110"/></DataGrid.Columns></DataGrid>
    </Grid>
   </TabItem>
   <TabItem Header="  Settings  ">
    <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel MaxWidth="760" HorizontalAlignment="Left">
     <TextBlock Text="PDF password" FontWeight="SemiBold" Margin="0,0,0,6"/>
     <PasswordBox x:Name="PasswordInput" Width="550" HorizontalAlignment="Left"/>
     <TextBlock x:Name="PasswordHint" Text="Blank keeps the saved password." Foreground="#637286" FontSize="12" Margin="0,5,0,14"/>
     <Button x:Name="SaveSettingsButton" Content="Save settings" HorizontalAlignment="Left" Margin="0,14,0,10"/>
     <Button x:Name="ForgetPasswordButton" Content="Forget saved password" HorizontalAlignment="Left" Margin="0,0,0,10"/>
     <TextBlock Text="Statements are read and processed on this PC. Your PDF password and transaction data stay on this PC. Saved passwords use Windows encryption for your account." TextWrapping="Wrap" Foreground="#637286" FontSize="12" Margin="0,0,0,18"/>
     <Separator/>
     <TextBlock Text="Monthly sheets" FontSize="17" FontWeight="SemiBold" Margin="0,14,0,7"/>
     <TextBlock x:Name="ScheduleHint" Text="The app creates the current month on launch. Enable the daily Windows check to create new months while the app is closed." TextWrapping="Wrap" Foreground="#637286" FontSize="12" Margin="0,0,0,10"/>
     <Button x:Name="ScheduleButton" Content="Enable automatic monthly sheets" HorizontalAlignment="Left"/>
     <TextBlock Text="Each year has its own file, such as Expenses-2026.xlsx, with month sheets named August, September, etc. Use the year selector or Previous/Next year to choose which workbook to open. Missed months are created after the PC starts again." TextWrapping="Wrap" Foreground="#637286" FontSize="12" Margin="0,9,0,18"/>
     <Separator/>
     <TextBlock Text="Workbook recovery" FontSize="17" FontWeight="SemiBold" Margin="0,14,0,7"/>
     <TextBlock Text="Make description changes through the app. Rebuild Excel replaces the workbook from saved transactions and rules, keeping a backup of the previous file. Use this after editing Excel directly." TextWrapping="Wrap" Foreground="#637286" FontSize="12" Margin="0,0,0,10"/>
     <Button x:Name="RebuildButton" Content="Back up and rebuild selected year" HorizontalAlignment="Left"/>
    </StackPanel></ScrollViewer>
   </TabItem>
   <TabItem Header="  Import history  "><DataGrid x:Name="HistoryGrid"><DataGrid.Columns><DataGridTextColumn Header="PDF" Binding="{Binding filename}" Width="*"/><DataGridTextColumn Header="Imported" Binding="{Binding imported_at}" Width="210"/><DataGridTextColumn Header="Method" Binding="{Binding engine}" Width="155"/><DataGridTextColumn Header="Expenses" Binding="{Binding row_count}" Width="85"/><DataGridTextColumn Header="Excluded" Binding="{Binding excluded_count}" Width="85"/></DataGrid.Columns></DataGrid></TabItem>
  </TabControl>
  <Border Grid.Row="3" Margin="0,16,0,0" Padding="12" Background="#E8EEF6" CornerRadius="5">
   <StackPanel><ProgressBar x:Name="Progress" Height="3" IsIndeterminate="True" Visibility="Collapsed" Margin="0,0,0,8"/><TextBlock x:Name="StatusText" Text="Loading your saved transactions..." TextWrapping="Wrap" Foreground="#18324F"/></StackPanel>
  </Border>
 </Grid>
</Window>
'@
$script:window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
$script:ui = @{}
# Find named controls from the trusted application markup.
foreach ($match in [regex]::Matches($xaml.OuterXml,'x:Name="([^"]+)"')) {
    $name=$match.Groups[1].Value
    $script:ui[$name]=$window.FindName($name)
}

$ui.ExpenseDate.SelectedDate=[DateTime]::Today
$script:manualEntryId=[guid]::NewGuid().ToString()

function Set-ExpenseRows {
    if (-not $script:state -or $script:updatingYears) { return }
    $year=[string]$ui.YearFilter.SelectedItem
    $selected = [string]$ui.MonthFilter.SelectedItem
    $rows = @($script:state.records | Where-Object {$_.date.StartsWith($year)})
    if ($selected -and $selected -ne 'All months') {
        $month = @($script:state.months | Where-Object { $_.sheet -eq $selected -and $_.month.StartsWith($year) })[0]
        $rows = @($rows | Where-Object { $_.date.StartsWith($month.month) })
    }
    $ui.TransactionsGrid.ItemsSource = $rows
    $ui.DeleteExpenseButton.IsEnabled=($null -ne $ui.TransactionsGrid.SelectedItem -and -not $script:busy)
}
function Get-SelectedYearInfo {
    $year=[string]$ui.YearFilter.SelectedItem
    return @($script:state.years | Where-Object {$_.year -eq $year}) | Select-Object -First 1
}
function Set-YearView {
    if (-not $script:state -or $script:updatingYears) {return}
    $info=Get-SelectedYearInfo
    if (-not $info) {return}
    $script:updatingYears=$true
    try {
        $selection=[string]$ui.MonthFilter.SelectedItem
        $ui.MonthFilter.ItemsSource=@('All months')+@($script:state.months | Where-Object {$_.month.StartsWith($info.year)} | Sort-Object month | ForEach-Object {$_.sheet})
        if ($selection -and $ui.MonthFilter.Items.Contains($selection)) {$ui.MonthFilter.SelectedItem=$selection} else {$ui.MonthFilter.SelectedIndex=0}
        $ui.CountLabel.Text=[string]$info.count
        $ui.TotalLabel.Text='RM '+([decimal]$info.total).ToString('N2')
        $ui.OpenButton.Content='Open '+$info.year+' Excel'
        $ui.YearFileLabel.Text=[IO.Path]::GetFileName($info.workbook)
        $ui.RebuildButton.Content='Back up and rebuild '+$info.year
        $ui.PreviousYearButton.IsEnabled=($ui.YearFilter.SelectedIndex -gt 0)
        $ui.NextYearButton.IsEnabled=($ui.YearFilter.SelectedIndex -lt $ui.YearFilter.Items.Count-1)
    } finally {$script:updatingYears=$false}
    Set-ExpenseRows
}
function Update-State($state) {
    $script:state=$state
    $ui.PasswordHint.Text=if ($state.has_password) {'Password saved for this Windows account. Blank keeps it.'} else {'Enter the password used to open your statement.'}
    $ui.RulesGrid.ItemsSource=@($state.rules)
    $ui.HistoryGrid.ItemsSource=@($state.imports)
    $year=[string]$ui.YearFilter.SelectedItem
    $script:updatingYears=$true
    try {
        $ui.YearFilter.ItemsSource=@($state.years | ForEach-Object {$_.year})
        if ($year -and $ui.YearFilter.Items.Contains($year)) {$ui.YearFilter.SelectedItem=$year}
        elseif ($ui.YearFilter.Items.Contains([string][DateTime]::Today.Year)) {$ui.YearFilter.SelectedItem=[string][DateTime]::Today.Year}
        elseif ($ui.YearFilter.Items.Count) {$ui.YearFilter.SelectedIndex=$ui.YearFilter.Items.Count-1}
    } finally {$script:updatingYears=$false}
    Set-YearView
    $task=Get-ScheduledTask -TaskName $script:taskName -ErrorAction SilentlyContinue
    if ($task -and $task.State -ne 'Disabled') {$ui.ScheduleHint.Text='Automatic checks are enabled in Windows. A new sheet is created at the first daily or logon check of the month.'}
    elseif ($task) {$ui.ScheduleHint.Text='The Windows task is disabled. Use the button below to enable automatic checks again.'}
}
function Start-Action($request, $message) {
    if ($script:busy) {return}
    $script:busy=$true
    $ui.StatusText.Text=$message
    $ui.Progress.Visibility='Visible'
    foreach ($name in @('UploadButton','DeleteExpenseButton','AddExpenseButton','ExpenseDate','ExpenseDescription','ExpensePrice','SaveRuleButton','DeleteRuleButton','SaveSettingsButton','RefreshButton','RebuildButton','ScheduleButton')) {$ui[$name].IsEnabled=$false}
    $info=New-Object System.Diagnostics.ProcessStartInfo
    $info.FileName=$script:python
    $info.Arguments='-B -E "'+(Join-Path $script:appRoot 'agent.py')+'"'
    $info.WorkingDirectory=$script:appRoot
    $info.UseShellExecute=$false
    $info.CreateNoWindow=$true
    $info.RedirectStandardInput=$true
    $info.RedirectStandardOutput=$true
    $info.RedirectStandardError=$true
    $info.EnvironmentVariables['PYTHONIOENCODING']='utf-8'
    $process=New-Object System.Diagnostics.Process
    $process.StartInfo=$info
    try {
        $null=$process.Start()
        $outTask=$process.StandardOutput.ReadToEndAsync()
        $errTask=$process.StandardError.ReadToEndAsync()
        $bytes=[System.Text.Encoding]::UTF8.GetBytes(($request | ConvertTo-Json -Depth 8 -Compress))
        $process.StandardInput.BaseStream.Write($bytes,0,$bytes.Length)
        $process.StandardInput.Close()
        $script:job=@{Process=$process; Output=$outTask; Error=$errTask; Action=$request.action; Request=$request}
    } catch {
        $ui.StatusText.Text='Could not start the agent. Close the window and launch it again.'
        $script:busy=$false
        $ui.Progress.Visibility='Collapsed'
        foreach ($name in @('UploadButton','DeleteExpenseButton','AddExpenseButton','ExpenseDate','ExpenseDescription','ExpensePrice','SaveRuleButton','DeleteRuleButton','SaveSettingsButton','RefreshButton','RebuildButton','ScheduleButton')) {$ui[$name].IsEnabled=$true}
    }
}
$timer=New-Object Windows.Threading.DispatcherTimer
$timer.Interval=[TimeSpan]::FromMilliseconds(250)
$timer.Add_Tick({
    if (-not $script:job -or -not $script:job.Process.HasExited) {return}
    $job=$script:job
    $script:job=$null
    $script:busy=$false
    $ui.Progress.Visibility='Collapsed'
    foreach ($name in @('UploadButton','DeleteExpenseButton','AddExpenseButton','ExpenseDate','ExpenseDescription','ExpensePrice','SaveRuleButton','DeleteRuleButton','SaveSettingsButton','RefreshButton','RebuildButton','ScheduleButton')) {$ui[$name].IsEnabled=$true}
    try {
        $response=$job.Output.Result | ConvertFrom-Json
        if ($response.ok) {
            Update-State $response.status
            if ($job.Action -eq 'import' -and $response.result.needs_review) {
                $script:busy=$true
                try {$decisions=Get-MatchDecisions $response.result.candidates}
                finally {$script:busy=$false}
                if ($null -ne $decisions) {
                    $job.Request.resolutions=$decisions
                    $job.Request.expected_hash=$response.result.pdf_hash
                    Start-Action $job.Request 'Saving the reviewed PDF expenses...'
                } else {$ui.StatusText.Text='Upload cancelled. No transactions from this PDF were added.'}
                return
            }
            $script:autoSyncBlocked=($response.result.reason -eq 'workbook_modified')
            if ($job.Action -eq 'add_expense' -and $response.result.saved) {
                $ui.ExpenseDescription.Clear()
                $ui.ExpensePrice.Clear()
                $script:manualEntryId=[guid]::NewGuid().ToString()
                $ui.YearFilter.SelectedItem=$response.result.month.Substring(0,4)
                $savedMonth=@($response.status.months | Where-Object {$_.month -eq $response.result.month})
                if ($savedMonth.Count) {$ui.MonthFilter.SelectedItem=$savedMonth[0].sheet}
            }
            if ($job.Action -eq 'import') {
                $ui.StatusText.Text=('{0} added. {1} duplicates skipped. {2} non-expense rows excluded. {3}' -f $response.result.added,$response.result.duplicates,$response.result.excluded,$response.result.message)
            } elseif ($response.result.message) {$ui.StatusText.Text=$response.result.message}
            elseif ($job.Action -eq 'save_settings') {$ui.PasswordInput.Clear(); $ui.StatusText.Text='Password saved using Windows encryption. You can now upload a PDF.'}
            else {$ui.StatusText.Text='Ready. Enter a date, description and price, then click Save to Excel.'}
        } else {$ui.StatusText.Text=$response.error}
    } catch {$ui.StatusText.Text='The agent could not finish. Your saved data remains available. Please try again.'}
    finally {$job.Process.Dispose(); $job.Request=$null}
})
$timer.Start()
$ui.MonthFilter.Add_SelectionChanged({Set-ExpenseRows})
$ui.YearFilter.Add_SelectionChanged({Set-YearView})
$ui.PreviousYearButton.Add_Click({if($ui.YearFilter.SelectedIndex -gt 0){$ui.YearFilter.SelectedIndex--}})
$ui.NextYearButton.Add_Click({if($ui.YearFilter.SelectedIndex -lt $ui.YearFilter.Items.Count-1){$ui.YearFilter.SelectedIndex++}})
$ui.TransactionsGrid.Add_SelectionChanged({$ui.DeleteExpenseButton.IsEnabled=($null -ne $ui.TransactionsGrid.SelectedItem -and -not $script:busy)})
$ui.DeleteExpenseButton.Add_Click({
    if ($script:busy -or -not $ui.TransactionsGrid.SelectedItem) {return}
    $entry=$ui.TransactionsGrid.SelectedItem
    $script:busy=$true
    try {$confirmed=Confirm-ExpenseDeletion $entry}
    finally {$script:busy=$false}
    if ($confirmed) {Start-Action @{action='delete_expense';id=$entry.id} 'Deleting the selected expense and updating its yearly workbook...'}
})
$ui.UploadButton.Add_Click({
    $dialog=New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Filter='PDF statements (*.pdf)|*.pdf'
    $dialog.Multiselect=$false
    if ($dialog.ShowDialog($window)) {
        $request=Get-UploadRequest $dialog.FileName
        if ($null -ne $request) {Start-Action $request 'Reading the PDF and saving new expenses on this PC...'}
    }
})
$ui.OpenButton.Add_Click({
    $info=Get-SelectedYearInfo
    if ($info -and (Test-Path -LiteralPath $info.workbook)) {
        try {Start-Process -FilePath $info.workbook}
        catch {$ui.StatusText.Text='The workbook is saved at '+$info.workbook+'. Open it with Excel or another app that supports .xlsx files.'}
    } else {$ui.StatusText.Text='The workbook is not ready yet. Click Refresh Excel.'}
})
$ui.RefreshButton.Add_Click({Start-Action @{action='refresh'} 'Updating your Excel workbook...'})
$ui.AddExpenseButton.Add_Click({
    if (-not $ui.ExpenseDate.SelectedDate) {$ui.StatusText.Text='Choose a valid expense date.'; return}
    $expenseDay=([datetime]$ui.ExpenseDate.SelectedDate).ToString('yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)
    Start-Action @{action='add_expense';date=$expenseDay;description=$ui.ExpenseDescription.Text;price=$ui.ExpensePrice.Text;entry_id=$script:manualEntryId} 'Saving your expense and updating Excel...'
})
$ui.SaveSettingsButton.Add_Click({
    $request=@{action='save_settings';pdf_password=$ui.PasswordInput.Password}
    Start-Action $request 'Saving your settings...'
})
$ui.ForgetPasswordButton.Add_Click({Start-Action @{action='forget_password'} 'Removing the saved password...'})
$ui.SaveRuleButton.Add_Click({
    $mode=if ($ui.RuleMode.SelectedIndex -eq 1) {'exact'} else {'contains'}
    Start-Action @{action='add_rule';pattern=$ui.PatternInput.Text;replacement=$ui.ReplacementInput.Text;mode=$mode} 'Saving your rule and updating Excel...'
})
$ui.DeleteRuleButton.Add_Click({if ($ui.RulesGrid.SelectedItem) {Start-Action @{action='delete_rule';id=$ui.RulesGrid.SelectedItem.id} 'Removing the rule and restoring original descriptions...'}})
$ui.RulesGrid.Add_SelectionChanged({if ($ui.RulesGrid.SelectedItem) {$ui.PatternInput.Text=$ui.RulesGrid.SelectedItem.pattern; $ui.ReplacementInput.Text=$ui.RulesGrid.SelectedItem.replacement; $ui.RuleMode.SelectedIndex=if($ui.RulesGrid.SelectedItem.mode -eq 'exact'){1}else{0}}})
$ui.RebuildButton.Add_Click({if($ui.YearFilter.SelectedItem){Start-Action @{action='rebuild';year=[string]$ui.YearFilter.SelectedItem} 'Backing up and rebuilding the selected year...'}})
$ui.ScheduleButton.Add_Click({
    try {
        & (Join-Path $script:scriptsRoot 'Enable monthly sheets.ps1')
        $ui.ScheduleHint.Text='Automatic checks are enabled in Windows. New months are created at the first daily or logon check.'
        $ui.StatusText.Text='Automatic monthly sheets enabled.'
    } catch {$ui.StatusText.Text='Windows could not enable the schedule. New months will still be created when you open the app.'}
})
$monthlyTimer=New-Object Windows.Threading.DispatcherTimer
$monthlyTimer.Interval=[TimeSpan]::FromHours(1)
$monthlyTimer.Add_Tick({if (-not $script:busy) {Start-Action @{action='ensure_month'} 'Checking for a new monthly sheet...'}})
$monthlyTimer.Start()
function Try-AutoSync {
    if ($script:busy -or -not $script:state -or -not $script:state.dirty) {return}
    # One open or externally edited year must not stop another pending year.
    $ready=$false
    foreach($info in @($script:state.years | Where-Object {$_.pending -and $_.reason -ne 'workbook_modified'})) {
        if (-not (Test-Path -LiteralPath $info.workbook)) {$ready=$true;break}
        try {$probe=[IO.File]::Open($info.workbook,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None);$probe.Dispose();$ready=$true;break}
        catch {continue}
    }
    if (-not $ready) {return}
    Start-Action @{action='ensure_month'} 'Saving pending entries to their monthly Excel sheets...'
}
$syncTimer=New-Object Windows.Threading.DispatcherTimer
$syncTimer.Interval=[TimeSpan]::FromSeconds(15)
$syncTimer.Add_Tick({Try-AutoSync})
$syncTimer.Start()
$window.Add_Closing({param($sender,$eventArgs) if ($script:busy) {$eventArgs.Cancel=$true; $ui.StatusText.Text='Please wait for the current operation to finish before closing.'}})
if ($SelfTest) {
    $stateRaw = '{"action":"status"}' | & $script:python (Join-Path $script:appRoot 'agent.py')
    $parsed=$stateRaw | ConvertFrom-Json
    if (-not $parsed.ok) {throw 'Backend status failed'}
    Update-State $parsed.status
    Start-Action @{action='status'} 'Checking saved expenses...'
    $checkUntil=[DateTime]::Now.AddSeconds(20)
    while ($script:busy -and [DateTime]::Now -lt $checkUntil) {
        $window.Dispatcher.Invoke([action]{},[Windows.Threading.DispatcherPriority]::ApplicationIdle)
        Start-Sleep -Milliseconds 100
    }
    if ($script:busy -or -not $script:state) {throw 'UI background worker failed'}
    if ($FeatureTest) {
        # Only the isolated package test creates this explicitly named fixture.
        $fixture=@($script:state.records | Where-Object {$_.description -eq 'Fictional delete-through-UI'})
        if (-not $env:TNG_AGENT_DATA -or $fixture.Count -ne 1 -or $ui.YearFilter.Items.Count -lt 2) {throw 'Fictional feature-test fixtures are missing.'}
        $ui.YearFilter.SelectedIndex=0
        if ($ui.PreviousYearButton.IsEnabled) {throw 'Previous year should be disabled at the first year.'}
        $firstYear=[string]$ui.YearFilter.SelectedItem
        $ui.NextYearButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
        if ([string]$ui.YearFilter.SelectedItem -eq $firstYear) {throw 'Next year button did not change the year.'}
        $ui.PreviousYearButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
        if ([string]$ui.YearFilter.SelectedItem -ne $firstYear) {throw 'Previous year button did not return to the first year.'}
        $ui.YearFilter.SelectedIndex=$ui.YearFilter.Items.Count-1
        if ($ui.NextYearButton.IsEnabled) {throw 'Next year should be disabled at the last year.'}
        $ui.YearFilter.SelectedItem=$fixture[0].date.Substring(0,4)
        $ui.MonthFilter.SelectedIndex=0
        if (@($ui.TransactionsGrid.ItemsSource | Where-Object {-not $_.date.StartsWith([string]$ui.YearFilter.SelectedItem)}).Count) {throw 'Year filter includes another year.'}
        function Start-Process {param([string]$FilePath) $script:openedWorkbook=$FilePath}
        $ui.OpenButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
        if ($script:openedWorkbook -ne (Get-SelectedYearInfo).workbook) {throw 'Open Excel used the wrong year.'}
        $ui.TransactionsGrid.SelectedItem=@($ui.TransactionsGrid.ItemsSource | Where-Object {$_.id -eq $fixture[0].id})[0]
        function Confirm-ExpenseDeletion {param($entry) return $false}
        $before=$script:state.count
        $ui.DeleteExpenseButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
        if ($script:busy -or $script:state.count -ne $before) {throw 'Cancelling deletion changed data.'}
        function Confirm-ExpenseDeletion {param($entry) return $true}
        $ui.DeleteExpenseButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
        $limit=[DateTime]::Now.AddSeconds(20)
        while ($script:busy -and [DateTime]::Now -lt $limit) {
            $window.Dispatcher.Invoke([action]{},[Windows.Threading.DispatcherPriority]::ApplicationIdle)
            Start-Sleep -Milliseconds 100
        }
        if ($script:busy -or $script:state.count -ne $before-1 -or @($script:state.records | Where-Object {$_.id -eq $fixture[0].id}).Count) {throw 'Confirmed deletion did not remove the selected fixture.'}
        Write-Output 'Year navigation, boundaries, year filtering, Open Excel destination, cancel deletion and confirmed deletion passed.'
    }
    $ui.Tabs.SelectedIndex=$PreviewTab
    $ui.StatusText.Text='Ready. Upload a PDF or enter a date, description and price to save expenses.'
    $window.Width=1060; $window.Height=760
    $window.WindowStartupLocation='Manual'
    $window.Left=-10000; $window.Top=-10000
    $window.ShowActivated=$false; $window.ShowInTaskbar=$false
    $window.Show()
    $window.Dispatcher.Invoke([action]{},[Windows.Threading.DispatcherPriority]::ApplicationIdle)
    $window.Measure((New-Object Windows.Size(1060,760)))
    $window.Arrange((New-Object Windows.Rect(0,0,1060,760)))
    $window.UpdateLayout()
    if ($PreviewPath) {
        $bitmap=New-Object Windows.Media.Imaging.RenderTargetBitmap(1060,760,96,96,[Windows.Media.PixelFormats]::Pbgra32)
        $bitmap.Render($window.Content)
        $encoder=New-Object Windows.Media.Imaging.PngBitmapEncoder
        $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
        $stream=[IO.File]::Create($PreviewPath)
        try {$encoder.Save($stream)} finally {$stream.Dispose()}
    }
    Write-Output ('UI loaded with '+$parsed.status.count+' transactions and '+$ui.MonthFilter.Items.Count+' month choices.')
    $timer.Stop(); $monthlyTimer.Stop(); $syncTimer.Stop()
    $window.Close()
} else {
    $window.Add_Loaded({
        # Install under the real user's account when they first launch the app.
        try {
            if (-not (Get-ScheduledTask -TaskName $script:taskName -ErrorAction SilentlyContinue)) {
                & (Join-Path $script:scriptsRoot 'Enable monthly sheets.ps1') | Out-Null
            }
        } catch { }
        Start-Action @{action='ensure_month'} 'Checking your workbook and monthly sheets...'
    })
    $null=$window.ShowDialog()
}
