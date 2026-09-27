function Get-UploadRequest($path) {
    [xml]$markup = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Open PDF statement" Width="490" SizeToContent="Height" ResizeMode="NoResize" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="14">
 <StackPanel Margin="24">
  <TextBlock Text="Password for this upload" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,12"/>
  <TextBlock Text="Enter the PDF password here, or leave it blank to use your saved password. This field is never saved." TextWrapping="Wrap" Margin="0,0,0,14"/>
  <PasswordBox x:Name="OncePassword" Padding="9"/>
  <TextBlock Text="To remember a password, save it in Settings using Windows encryption." TextWrapping="Wrap" Foreground="#637286" Margin="0,12,0,18"/>
  <StackPanel Orientation="Horizontal" HorizontalAlignment="Right"><Button x:Name="Cancel" Content="Cancel" Padding="16,9" Margin="0,0,10,0" IsCancel="True"/><Button x:Name="Continue" Content="Read PDF" Padding="16,9" IsDefault="True"/></StackPanel>
 </StackPanel>
</Window>
'@
    $prompt = [Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $markup))
    $prompt.Owner = $script:window
    # Capture the owning dialog explicitly for each event handler.
    $button = $prompt.FindName('Continue')
    $button.Add_Click({$prompt.DialogResult=$true}.GetNewClosure())
    $prompt.FindName('Cancel').Add_Click({$prompt.DialogResult=$false}.GetNewClosure())
    if (-not $prompt.ShowDialog()) {return $null}
    $request = @{action='import';path=$path}
    $password = $prompt.FindName('OncePassword').Password
    if ($password.Length) {$request.pdf_password=$password}
    $prompt.FindName('OncePassword').Clear()
    return $request
}

function Confirm-ExpenseDeletion($entry) {
    [xml]$markup = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Delete expense" Width="530" SizeToContent="Height" ResizeMode="NoResize" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="14">
 <StackPanel Margin="24">
  <TextBlock Text="Delete this saved expense?" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,14"/>
  <TextBlock x:Name="ExpenseDetails" TextWrapping="Wrap" Margin="0,0,0,14"/>
  <TextBlock Text="It will be removed from the app and its yearly Excel file. A database backup is kept first. Its known PDF reference will stay excluded from future uploads." TextWrapping="Wrap" Foreground="#637286" Margin="0,0,0,20"/>
  <StackPanel Orientation="Horizontal" HorizontalAlignment="Right"><Button x:Name="Cancel" Content="Keep expense" Padding="16,9" Margin="0,0,10,0" IsCancel="True" IsDefault="True"/><Button x:Name="Delete" Content="Delete expense" Padding="16,9"/></StackPanel>
 </StackPanel>
</Window>
'@
    $prompt=[Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $markup))
    $prompt.Owner=$script:window
    $prompt.FindName('ExpenseDetails').Text=('{0} | {1} | RM {2}' -f $entry.date,$entry.display_description,$entry.amount)
    $prompt.FindName('Delete').Add_Click({$prompt.DialogResult=$true}.GetNewClosure())
    $prompt.FindName('Cancel').Add_Click({$prompt.DialogResult=$false}.GetNewClosure())
    return $prompt.ShowDialog() -eq $true
}

function Get-MatchDecisions($candidates) {
    $decisions=@{}
    $used=@{}
    foreach ($candidate in $candidates) {
        [xml]$markup = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Review a possible duplicate" Width="660" SizeToContent="Height" ResizeMode="NoResize" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="14">
 <StackPanel Margin="24">
  <TextBlock Text="Is this expense already saved?" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,14"/>
  <TextBlock x:Name="PdfRow" TextWrapping="Wrap" Margin="0,0,0,14"/>
  <TextBlock Text="These manual entries have the same date and amount. Choose a matching entry, or keep both if they are different purchases." TextWrapping="Wrap" Margin="0,0,0,12"/>
  <ListBox x:Name="Matches" DisplayMemberPath="label" Height="150"/>
  <TextBlock Text="Matching keeps your manual description and counts the expense only once. Cancel leaves the whole PDF unimported." TextWrapping="Wrap" Foreground="#637286" Margin="0,14,0,18"/>
  <WrapPanel HorizontalAlignment="Right"><Button x:Name="Cancel" Content="Cancel upload" Padding="12,9" Margin="0,0,8,0"/><Button x:Name="Keep" Content="Keep both" Padding="12,9" Margin="0,0,8,0"/><Button x:Name="Match" Content="Use selected entry" Padding="12,9"/></WrapPanel>
 </StackPanel>
</Window>
'@
        $prompt = [Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $markup))
        $prompt.Owner=$script:window
        $prompt.FindName('PdfRow').Text=('PDF: {0} | {1} | RM {2}' -f $candidate.date,$candidate.description,$candidate.amount)
        $choices=@($candidate.matches | Where-Object {-not $used.ContainsKey($_.id)} | ForEach-Object {
            [pscustomobject]@{id=$_.id;label=($_.description+' | RM '+$_.amount+' | '+$_.id.Substring(7,8))}
        })
        $list=$prompt.FindName('Matches'); $list.ItemsSource=$choices
        if ($choices.Count) {$list.SelectedIndex=0} else {$prompt.FindName('Match').IsEnabled=$false}
        $prompt.Tag=$null
        $prompt.FindName('Keep').Add_Click({$prompt.Tag='keep';$prompt.DialogResult=$true}.GetNewClosure())
        $prompt.FindName('Match').Add_Click({if($list.SelectedItem){$prompt.Tag=$list.SelectedItem.id;$prompt.DialogResult=$true}}.GetNewClosure())
        $prompt.FindName('Cancel').Add_Click({$prompt.DialogResult=$false}.GetNewClosure())
        if (-not $prompt.ShowDialog()) {return $null}
        $decisions[$candidate.id]=[string]$prompt.Tag
        if ($prompt.Tag -ne 'keep') {$used[[string]$prompt.Tag]=$true}
    }
    return $decisions
}
