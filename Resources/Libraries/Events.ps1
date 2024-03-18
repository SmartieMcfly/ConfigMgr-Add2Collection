$UI.Window.Add_Loaded({
    $This.Activate()
    # Create registry keys in the CU hive if they don't exist
    Add-RegSetting

    # Read the registry keys
    Read-RegSetting


    # If we have a SQL server and database...
    If ($ui.SessionData[0] -and $UI.SessionData[1] -and $UI.SessionData[2])
    {
        # Check if new version is available in a background job
        $Code = {
            Param($UI)
            Test-UpdateAvailable -UI $UI
        }
        $Job = [BackgroundJob]::new($Code,@($UI),@("Function:\Test-UpdateAvailable","Function:\Show-BalloonTip"))
        $UI.Jobs += $Job
        $Job.Start()
    }

})

$UI.SelectCollection.Add_Click({
    . "$(Join-Path $ResourcePath 'Show-CollectionsTreeView.ps1')"
    $UI.MembersGrid.ItemsSource = $null
})

$UI.Populate.Add_Click({
    Try
    {
        Get-CollectionMembers -ErrorAction Stop
    }
    Catch
    {
        $Customerror = $_.Exception.Message.Replace('"',"'")
        Show-WpfMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
        Return
    }

})

$Text = "Enter or paste resource names here, each resource on a new line.`nFor devices, add the computer name. For users or user groups, add the user or group name (no domain prefix required)."

$UI.Resources.Add_TextChanged({
    If ($This.Text -ne "" -and $This.text -ne $Text)
    {
        $UI.AddResources.IsEnabled = "True"
    }
    Else
    {
        $UI.AddResources.IsEnabled = $False
        $This.Foreground = "Gray"
    }
})

$UI.Resources.Add_GotFocus({
    If ($This.Text -eq "")
    {
        $This.Text = $Text
        $This.Foreground = "Gray"
    }
    ElseIf ($This.text -eq $Text)
    {
        $This.Text = ""
        $This.Foreground = "Black"
    }
})

$UI.Resources.Add_LostFocus({
    If ($This.Text -eq "")
    {
        $This.Text = $Text
        $This.Foreground = "Gray"
    }
})

$UI.AddResources.Add_Click({

    $UI.StatusBarData[1] = "Adding resources"
    $UI.StatusBarData[2] = 0
    $UI.StatusBarData[3] = 0
    $UI.StatusBarData[4] = 0
    $UI.SessionData[3] = "True"

    Try
    {
        Add-ResourcesToCollection -ErrorAction Stop
    }
    Catch
    {
        $Customerror = $_.Exception.Message.Replace('"',"'")
        Show-WpfMessageBox -Title "SQL Error" -Content $Customerror -BorderThickness 1 -BorderBrush Red -Sound 'Windows Error' -TitleBackground Red -TitleTextForeground GhostWhite -TitleFontWeight Bold -TitleFontSize 20
        Return
    }

})

$UI.Btn_Settings.Add_Click({
    Show-SettingsView
})

$UI.Btn_About.Add_Click({
    Show-AboutView
})

$UI.Btn_Help.Add_Click({
    Show-HelpView
})

$UI.Btn_Exit.Add_Click({
    $UI.Window.Close()
})
