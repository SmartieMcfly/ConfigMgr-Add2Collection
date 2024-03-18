#---[ Initialization ]---#
# Paths
$MyRoot = 'C:\Users\smcollier\Source\GitHub\ConfigMgr-Add2Collection'
$MyRoot = $PSScriptRoot
$ResourcePath = Join-Path $MyRoot 'Resources'
$AssetPath = Join-Path $ResourcePath 'Assets'
$LibraryPath = Join-Path $ResourcePath 'Libraries'
$ViewPath = Join-Path $ResourcePath 'Views'
# Controls
$TestHashes = $False

#---[ Assemblies ]---#
Add-Type -AssemblyName PresentationFramework
Add-Type -Path (Join-Path $LibraryPath 'MaterialDesignColors.dll')
Add-Type -Path (Join-Path $LibraryPath 'MaterialDesignThemes.Wpf.dll')

#---[ Functions ]---#
. (Join-Path $LibraryPath 'Functions.ps1')

#---[ Classes ]---#
. (Join-Path $LibraryPath 'Classes.ps1')

#---[ Traps ]---#
# PowerShell Version Check
if ($PSVersionTable.PSVersion.Major -lt 5) {
    $Content = 'ConfigMgr Add2Collection cannot start because it requires PowerShell 5 or greater. Please upgrade your PowerShell version.'
    Show-WpfMessageBox -Content $Content -Title 'Oops!' -TitleBackground Orange -TitleTextForeground Yellow -TitleFontSize 20 -TitleFontWeight Bold -BorderThickness 1 -BorderBrush Orange -Sound 'Windows Exclamation'
    Break
}

#---[ File Hash Check ]---#
if ($TestHashes) {
    # Define File Hashes
    $FileHashes = @{
        'Show-CollectionsTreeView.ps1' = 'ED5A75926DE110E2E9531327DDAE48DCF9473352D2B4DC49525949F370E693C3'
        'add2-1.png'                   = '8F1F31F42EBBB6D437B22AAC43442297BE0C4974429EB4C43F629AE1EF9E3C5B'
        'add2-2.png'                   = 'F55BE06DF33D314F833D6FD7792A326E5BD9B89E9E485C1897D9789CDA95D4A5'
        'col.ico'                      = 'DFD1D78262DA9A7CC21C89228CBB0336B30EC3C39FAC987C006A3296AD8DC9D7'
        'col.png'                      = 'A68916A4AA47EDBF195747FC04B3E3187B019CEF5825ABB84BD7FFD3C99B75E7'
        'help-circle.png'              = 'DAF81B62D27979BA726C22606D166BD85699DDB6829289E7DED44CECE913ED3D'
        'Icon0.bmp'                    = '4126031997384D38C7F472D6AA786AA5901559CFA8FAE41CDEFDDDBB1A726F63'
        'Icon194.bmp'                  = '87EE64DFFA186683CD1DDDA393E6FC8CD27EBE097895ABEE14F15F8D54A8C86D'
        'Icon21.bmp'                   = '806E1D0BBF8C7BF925CA2D51789E7242696DDF0059DC488D20D32E3A37321F7B'
        'Icon96.bmp'                   = 'B0A6FE6B16B7CACD9081BD9E2B91ACD8401A5EFB9153C057882E703E6F46475A'
        'information-outline.png'      = '5034BC311E6E76789B5E0E834D15826857E20A48DBA7C7ACAB75F63C68F532C3'
        'settings.png'                 = '1659BB8CDF22D86F1B3BBF3F962D080FA717BD1408C215CA1DD6E34865B71EE1'
        'Classes.ps1'                  = '09D2308E75524576796759E95237E11FB03433AFF6E4FFE10F78F3453870023D'
        'Events.ps1'                   = 'FA1BB0F3A799A5A6C5366FCB0BE34C5061BB6BD92F73CC03CEB53FBF810D2429'
        'Functions.ps1'                = 'A77B39ED42138C4FC2C30A6B2E62E80D24405EAEC55484411B29F9DD9F2BC8DB'
        'MaterialDesignColors.dll'     = '778C199154A97E3501A98283B517CD4FCD20481C0A7CE898EBA5D44257E4157D'
        'MaterialDesignThemes.Wpf.dll' = 'C09EDF6C96C352B5D108DD6A685FF6D855B582CF6ED3C4A1475F1A26718775BF'
        'About.xaml'                   = 'A9994056D48CE205A14987372304E6FDB460AE5BBE362C235E39D4A9C73DA038'
        'App.xaml'                     = 'B4BB3F9005EA28E60543B00C0A41BA6CF2C16873983FA0A0F3D2710E379692BE'
        'Help.Xaml'                    = 'E52DC0F561D41A74522EBDC7660A77A89606E324BCE74769F27492EAD7797812'
        'HelpFlowDocument.xaml'        = 'C11CD14C0B554C986E310AA0B2E2DB2571D6FFD59AB79D334B6E01C03EB219B1'
        'Settings.xaml'                = '4E9071DAC7371F235E23FA15908D12B985938952EEE49CABA26EBEF0B33CA08B'
    }

    # Get Resource Files
    $ResourceFiles = Get-ChildItem -Path $ResourcePath -Recurse -File

    # Check Resource Hashes
    $ResourceFiles | ForEach-Object {
        if ((Get-FileHash -Path $_.FullName).Hash -ne $FileHashes.$($_.Name)) {
            $Content = "Hash check failed on: $($_.FullName)"
            Show-WpfMessageBox -Content $Content -Title 'Security Error' -TitleBackground Red -TitleTextForeground Black -TitleFontSize 20 -TitleFontWeight Bold -BorderThickness 1 -BorderBrush Black -Sound 'Windows Exclamation'
            Break
        }
    }
}

#---[ Load App ]---#
# Read App.xaml
[XML]$AppXaml = [System.IO.File]::ReadAllLines("$(Join-Path $ViewPath 'App.xaml')")

# Create Synchronized Hash Table
$UI = [System.Collections.Hashtable]::Synchronized(@{})
$UI.Host = $Host
$UI.Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $AppXaml))
$AppXaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object -Process {
    $UI.$($_.Name) = $UI.Window.FindName($_.Name)
}

# Set Window Icon
$UI.Window.Icon = "$(Join-Path $AssetPath 'col.png')"

# Load Events Library
. $(Join-Path $LibraryPath 'Events.ps1')

# Set Session Data
$UI.SessionData = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$UI.SessionData.Add($null) # SQL Server
$UI.SessionData.Add($null) # Database
$UI.SessionData.Add($null) # Site server
$UI.SessionData.Add('False')
$UI.SessionData.Add('HKCU:\SOFTWARE\CBE\DeviceManager')  # Reg branch
$UI.SessionData.Add($null)
$UI.SessionData.Add([double]1.0) # current version
$UI.SessionData.Add($MyRoot)
$UI.SessionData.Add($null) # SQLServer
$UI.SessionData.Add($null) # Database
$UI.SessionData.Add($null) # AdminUIServer
$UI.SessionData.Add($null) # Changes xml
$UI.SessionData.Add($null) # Change table
$UI.Window.DataContext = $UI.SessionData

# Collection Info
$UI.CollectionInfo = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$UI.CollectionInfo.Add($null)
$UI.CollectionInfo.Add($null)
$UI.CollectionInfo.Add($null)
$UI.CollectionInfo.Add($null)
$UI.CollectionInfo.Add($null)
$UI.CollectionInfo.Add($null)
$UI.CollectionInfo.Add($null)
$UI.CollectionInfo.Add($null)
$UI.WrapPanel.DataContext = $UI.CollectionInfo

# Add Resource Results
$UI.Results = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$UI.Results.Add($null)
$UI.ResultsGrid.DataContext = $UI.Results

# Add Status Bar
$UI.StatusBarData = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$UI.StatusBarData.Add('Idle')
$UI.StatusBarData.Add(0)
$UI.StatusBarData.Add(0)
$UI.StatusBarData.Add(0)
$UI.StatusBarData.Add(0)
$UI.StatusBar.DataContext = $UI.StatusBarData

#---[ Launch App ]---#
# Make PowerShell Disappear
$WindowCode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$AsyncWindow = Add-Type -MemberDefinition $WindowCode -Name Win32ShowWindowAsync -Namespace Win32Functions -PassThru
$null = $AsyncWindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

# Load App
$App = New-Object -TypeName Windows.Application
$App.Run($UI.Window) | Out-Null
