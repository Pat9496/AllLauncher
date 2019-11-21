param
(
  [Parameter(Position = 0)]
  [string]
  $System,
  [Parameter(Position = 1)]
  [string]
  $Game,
  [Parameter(Position = 2)]
  [Alias('Reset','Back','Return','Launcher')]
  [switch]
  $Menu
)

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue

'Ending other instances'
$Explorer = Get-Process -Name explorer
$OtherInstances = Get-Process -Name powershell, powershell32, AllLauncher, cmd, powershell_ise -ErrorAction SilentlyContinue
ForEach ($OtherInstance in $OtherInstances) 
{
  If ($OtherInstance.Id -ne $pid) 
  {
    Stop-Process -Id $OtherInstance.Id -Force -ErrorAction SilentlyContinue
  }
}

filter timestamp 
{
  ('{0}: {1}' -f (Get-Date -Format 'yyyy-MM-dd, HH:mm:ss'), $_)
}
$CurrentDir = (Get-Location | Select-Object -ExpandProperty Path)
#Manually Set a Directory:
$CurrentDir = "$env:HOMEDRIVE\AllLauncher"
$LogFileDir = ('{0}\Logs' -f $CurrentDir)
$CFGDir = ('{0}\cfg' -f $CurrentDir)
$INIFile = ('{0}\AllLauncher.ini' -f $CFGDir)
$LogFile = ('{0}\LastGame.txt' -f $LogFileDir)
$Transcript = ('{0}\AllLauncher.log' -f $LogFileDir)
$ProcessBlacklistFile = ('{0}\ProcessBlacklist.txt' -f $CFGDir)
$ServiceBlacklistFile = ('{0}\ServiceBlacklist.txt' -f $CFGDir)
$ServiceWhitelistFile = ('{0}\ServiceWhitelist.txt' -f $CFGDir)
$DS4BlacklistFile = ('{0}\DS4SystemBlacklist.txt' -f $CFGDir)
$DS4WindowsBlacklistFile = ('{0}\DS4WindowsBlacklist.txt' -f $CFGDir)
$BorderlessGamesList = ('{0}\BorderlessGamesList.txt' -f $CFGDir)
$MIDISynthGamesFile = ('{0}\MIDISynthGames.txt' -f $CFGDir)
$MIDISynthSystemsFile = ('{0}\MIDISynthSystems.txt' -f $CFGDir)


#Remove-Item -Path $Transcript -Force -ErrorAction SilentlyContinue

'Starting Log!'
'Starting Log!' | timestamp > $LogFile

Start-Transcript -Path $Transcript -Force -ErrorAction SilentlyContinue
'Transcript started.'
('Writing transcript to {0}.' -f $Transcript) | timestamp > $LogFile

'Loading general functions...'
#These are general functions
New-Item "$($profile | Split-Path)\Modules\AudioDeviceCmdlets" -Type directory -Force -ErrorAction SilentlyContinue
Copy-Item -Path "$CurrentDir\DLLs\AudioDeviceCmdlets.dll" -Destination "$($profile | Split-Path)\Modules\AudioDeviceCmdlets\AudioDeviceCmdlets.dll" -Force -ErrorAction SilentlyContinue
Set-Location -Path "$($profile | Split-Path)\Modules\AudioDeviceCmdlets" -ErrorAction SilentlyContinue
Get-ChildItem | Unblock-File
Import-Module -Name AudioDeviceCmdlets -ErrorAction SilentlyContinue
Set-Location -Path $CurrentDir
$wshell = New-Object -ComObject wscript.shell

function Hide-DesktopIcons 
{
  $signature = @" 
[DllImport("user32.dll")]  
public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);  
[DllImport("user32.dll")]  
public static extern bool ShowWindow(IntPtr hWnd,int nCmdShow); 
"@ 

  $icons = Add-Type -MemberDefinition $signature -Name Win32Window -Namespace ScriptFanatic.WinAPI -PassThru 
  $hWnd = $icons::FindWindow('Progman','Program Manager') 
  $null = $icons::ShowWindow($hWnd,0) 
  'Hiding Desktop Icons'
} 
function Show-DesktopIcons 
{
  $signature = @" 
[DllImport("user32.dll")]  
public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);  
[DllImport("user32.dll")]  
public static extern bool ShowWindow(IntPtr hWnd,int nCmdShow); 
"@ 

  $icons = Add-Type -MemberDefinition $signature -Name Win32Window -Namespace ScriptFanatic.WinAPI -PassThru 
  $hWnd = $icons::FindWindow('Progman','Program Manager') 
  $null = $icons::ShowWindow($hWnd,5) 
  'Showing Desktop Icons'
} 
function Activate-App 
{
  param
  (
    [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
    $Application
  )
  Add-Type -AssemblyName microsoft.VisualBasic
  Add-Type -AssemblyName System.Windows.Forms
  [Microsoft.VisualBasic.Interaction]::AppActivate($Application)
  ('Activating {0}' -f $Application)
}
function Send-Keys 
{
  param
  (
    [Parameter(Mandatory, Position = 0)]
    $KeysToSend
  )
  Add-Type -AssemblyName microsoft.VisualBasic
  Add-Type -AssemblyName System.Windows.Forms
  [Windows.Forms.SendKeys]::SendWait($KeysToSend)
  ('Sending {0}' -f $KeysToSend)
}
function Get-INIValue 
{
  param
  (
    [Parameter(Mandatory, Position = 0)]
    $Path,
    [Parameter(Mandatory, Position = 1)]
    [string]
    $Section,
    [Parameter(Mandatory, Position = 2)]
    [string]
    $Key
  )
  
  $signature = @'
[DllImport("kernel32.dll")]
public static extern uint GetPrivateProfileString(
string lpAppName,
string lpKeyName,
string lpDefault,
StringBuilder lpReturnedString,
uint nSize,
string lpFileName);
'@
  $type = Add-Type -MemberDefinition $signature -Name IniRead -Namespace INIAPI -UsingNamespace System.Text -PassThru
  $builder = New-Object -TypeName System.Text.StringBuilder -ArgumentList 1024
  $len = [INIAPI.IniRead]::GetPrivateProfileString($Section, $Key, '', $builder, $builder.Capacity, $Path)
  $builder.ToString() 
}
function Set-INIValue 
{  
  param
  (
    [Parameter(Mandatory, Position = 0)]
    $Path,
    [Parameter(Mandatory, Position = 1)]
    [string]
    $Section,
    [Parameter(Mandatory, Position = 2)]
    [string]
    $Key,
    [Parameter(Mandatory, Position = 3)]
    [string]
    $Value,
    [Switch]
    $Open
  )
  
  $signature = @' 
[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] 
[return: MarshalAs(UnmanagedType.Bool)] 
public static extern bool WritePrivateProfileString( 
string lpAppName, 
string lpKeyName, 
string lpString, 
string lpFileName); 
'@ 
  
  
  Add-Type -MemberDefinition $signature -Name IniWrite -Namespace INIAPI -UsingNamespace System.Text 
  [INIAPI.IniWrite]::WritePrivateProfileString($Section, $Key, $Value, $Path)
   
  if ($Open) 
  {
    & "$env:windir\system32\notepad.exe" $Path
  }
}
function Get-Shortcut 
{
  [CmdletBinding()]
  param(
    $Path = $null
  )

  $obj = New-Object -ComObject WScript.Shell

  if ($Path -eq $null) 
  {
    $pathUser = [System.Environment]::GetFolderPath('StartMenu')
    $pathCommon = $obj.SpecialFolders.Item('AllUsersStartMenu')
    $Path = Get-ChildItem $pathUser, $pathCommon -Filter *.lnk -Recurse 
  }
  if ($Path -is [string]) 
  {
    $Path = Get-ChildItem $Path -Filter *.lnk
  }
  $Path | ForEach-Object -Process { 
    if ($_ -is [string]) 
    {
      $_ = Get-ChildItem $_ -Filter *.lnk
    }
    if ($_) 
    {
      $link = $obj.CreateShortcut($_.FullName)

      $info = @{}
      $info.Hotkey = $link.Hotkey
      $info.TargetPath = $link.TargetPath
      $info.LinkPath = $link.FullName
      $info.Arguments = $link.Arguments
      $info.Target = try 
      {
        Split-Path -Path $info.TargetPath -Leaf
      }
      catch 
      {
        'n/a'
      }
      $info.Link = try 
      {
        Split-Path -Path $info.LinkPath -Leaf
      }
      catch 
      {
        'n/a'
      }
      $info.WindowStyle = $link.WindowStyle
      $info.IconLocation = $link.IconLocation

      New-Object -TypeName PSObject -Property $info
    }
  }
}
function Wait-ProcessToCalm
{
  param
  (
    [Parameter(Position = 0)]
    [string]
    $ProcessToCalm,
    [Parameter(Position = 1)]
    [int]
    $CalmThreshold
  )

  If ($ProcessToCalm -eq '') 
  {
    'No process to wait for until it is calm.'
  }
  else
  { 
    If ($CalmThreshold -eq '') 
    {
      $CalmThreshold = 5
    } 
    $WaitTime = 100
    "Waiting for process $ProcessToCalm to calm down before continuing..."

    $WaitCounter = 0
    While ((!(Get-Process -Name $ProcessToCalm -ErrorAction SilentlyContinue)) -AND ($WaitCounter -lt $WaitTime)) 
    { 
      Start-Sleep -Milliseconds 250
      $WaitCounter = $WaitCounter + 1
    }

    Start-Sleep -Milliseconds 500
  
    $WaitCounter = 0
    $CurrentProcessPercentage = (((Get-Counter -Counter "\Process($ProcessToCalm)\% Processor Time" -ErrorAction SilentlyContinue).CounterSamples).CookedValue)
    While (($CurrentProcessPercentage -gt $CalmThreshold) -AND ($WaitCounter -lt $WaitTime)) 
    {
      Start-Sleep -Milliseconds 500
      $WaitCounter = $WaitCounter + 1
      $CurrentProcessPercentage = (((Get-Counter -Counter "\Process($ProcessToCalm)\% Processor Time" -ErrorAction SilentlyContinue).CounterSamples).CookedValue)
    }
  }
}



'General functions loaded.'

###################################### End of functions #########################################

Stop-Process -Name DS4Windows -ErrorAction SilentlyContinue
Stop-Process -Name Explorer -ErrorAction SilentlyContinue


('Getting Settings from {0}...' -f $INIFile)
('Getting Settings from {0}...' -f $INIFile) | timestamp >> $LogFile

IF ($System -eq '' -or $System -eq $null)
{
  'No System selected. Getting INI-entry...' | timestamp >> $LogFile
  $System = (Get-INIValue -Path $INIFile -Section 'LastGame' -Key 'System')
  ('Gotten {0} from INI-file.' -f $System) | timestamp >> $LogFile
}
else 
{
  ('Gotten {0} from parameter.' -f $System) | timestamp >> $LogFile
  Set-INIValue -Path $INIFile -Section 'LastGame' -Key 'System' -Value $System
}
IF ($Game -eq '' -or $Game -eq $null)
{
  'No game selected. Getting INI-entry...' | timestamp >> $LogFile
  $Game = (Get-INIValue -Path $INIFile -Section $System -Key 'Game')
  ('Gotten {0} from INI-file.' -f $Game) | timestamp >> $LogFile
}
else 
{
  ('Gotten {0} from parameter.' -f $Game) | timestamp >> $LogFile
  Set-INIValue -Path $INIFile -Section 'LastGame' -Key 'Game' -Value $Game
}
$LauncherName = (Get-INIValue -Path $INIFile -Section 'Launcher' -Key 'Name')
$LauncherDir = (Get-INIValue -Path $INIFile -Section 'Launcher' -Key 'Folder')
$LauncherExe = (Get-INIValue -Path $INIFile -Section 'Launcher' -Key 'Executable')
$LauncherProcess = $LauncherExe -Replace '.exe'
$LauncherVariables = (Get-INIValue -Path $INIFile -Section 'Launcher' -Key 'Variables')
$LauncherWindow = (Get-Process -Name $LauncherProcess -ErrorAction SilentlyContinue | Select-Object -ExpandProperty MainWindowTitle)
$LauncherQuitKey = (Get-INIValue -Path $INIFile -Section 'Launcher' -Key 'QuitKey')
[string]$LauncherQuit = (Get-INIValue -Path $INIFile -Section 'Launcher' -Key 'Quit')
IF ($LauncherQuit -eq '1')
{
  [bool]$LauncherQuit = $true
}
else 
{
  [bool]$LauncherQuit = $false
}
$Launcher = ('{0}\{1}' -f $LauncherDir, $LauncherExe)
$EmulatorINIEntry = (Get-INIValue -Path $INIFile -Section 'Emulators' -Key $System)
$GameDocsDir = (Get-INIValue -Path $INIFile -Section 'Directories' -Key 'GameDocsFolder')
$TrainersDir = (Get-INIValue -Path $INIFile -Section 'Directories' -Key 'TrainersFolder')
$DS4Folder = (Get-INIValue -Path $INIFile -Section 'Controllers' -Key 'DS4Folder')
$UsePS4Pad = (Get-INIValue -Path $INIFile -Section 'Controllers' -Key 'UsePS4Pad')
IF ($UsePS4Pad -eq '1')
{
  [bool]$UsePS4Pad = $true
}
else 
{
  [bool]$UsePS4Pad = $false
}[string]$UseCheats = (Get-INIValue -Path $INIFile -Section 'Options' -Key 'UseCheats')
IF ($UseCheats -eq '1')
{
  [bool]$UseCheats = $true
}
else 
{
  [bool]$UseCheats = $false
}
[bool]$WasLauncherRunning = ([bool](Get-Process -Name $LauncherProcess -ErrorAction SilentlyContinue))
$StartCurrentSystem = 'Start-' + $System
$GameExt = ($Game -split '\.')[-1]
$GameName = ($Game -Replace ('.{0}' -f $GameExt) -split '\\')[-1]
$ArtMoneyProcess = (Get-INIValue -Path $INIFile -Section 'ArtMoney' -Key 'Executable') -replace '.exe'
$CheatEngineProcess = (Get-INIValue -Path $INIFile -Section 'CheatEngine' -Key 'Executable') -replace '.exe'
$CoSMOSProcess = (Get-INIValue -Path $INIFile -Section 'CoSMOS' -Key 'Executable') -replace '.exe'
[bool]$WasArtMoneyRunning = ([bool](Get-Process -Name $ArtMoneyProcess -ErrorAction SilentlyContinue))
[string]$UseDisplayFusion = (Get-INIValue -Path $INIFile -Section 'DisplayFusion' -Key 'UseDisplayFusion')
IF ($UseDisplayFusion -eq '1')
{
  [bool]$UseDisplayFusion = $true
}
else 
{
  [bool]$UseDisplayFusion = $false
}
$GameTrainer = $TrainersDir + '\' + $GameName + '.exe'
$Additionaltrainer = $TrainersDir + '\' + $GameName + '_add.exe'
$ArtMoneyName = (Get-INIValue -Path $INIFile -Section 'ArtMoney' -Key 'Name')
$ArtMoneyExt = (Get-INIValue -Path $INIFile -Section 'ArtMoney' -Key 'Extension')
$ArtMoneyDir = (Get-INIValue -Path $INIFile -Section 'ArtMoney' -Key 'Folder')
$ArtMoneyTablesDir = (Get-INIValue -Path $INIFile -Section 'ArtMoney' -Key 'TablesFolder')
$ArtMoneyExe = (Get-INIValue -Path $INIFile -Section 'ArtMoney' -Key 'Executable')
[string]$ArtMoneyLoad = (Get-INIValue -Path $INIFile -Section 'ArtMoney' -Key 'AlwaysLoad')
IF ($ArtMoneyLoad -eq '1')
{
  [bool]$ArtMoneyLoad = $true
}
else 
{
  [bool]$ArtMoneyLoad = $false
}
[string]$UseSystemsSubfolder = (Get-INIValue -Path $INIFile -Section 'ArtMoney' -Key 'UseSystemsSubfolder')
IF ($UseSystemsSubfolder -eq '1')
{
  [bool]$UseSystemsSubfolder = $true
}
else 
{
  [bool]$UseSystemsSubfolder = $false
}
If ($UseSystemsSubfolder -eq $true) 
{
  $ArtMoneySystemDir = $ArtMoneyTablesDir + '\' + $System
}
else 
{
  $ArtMoneySystemDir = $ArtMoneyDir
} 
$ArtMoneyProcess = $ArtMoneyExe -Replace '.exe'
$TrainerProcess = $GameName
$ArtMoney = $ArtMoneyDir + '\' + $ArtMoneyExe
$CheatEngineName = (Get-INIValue -Path $INIFile -Section 'CheatEngine' -Key 'Name')
$CheatEngineExt = (Get-INIValue -Path $INIFile -Section 'CheatEngine' -Key 'Extension')
$CheatEngineDir = (Get-INIValue -Path $INIFile -Section 'CheatEngine' -Key 'Folder')
$CheatEngineTablesDir = (Get-INIValue -Path $INIFile -Section 'CheatEngine' -Key 'TablesFolder')
$CheatEngineExe = (Get-INIValue -Path $INIFile -Section 'CheatEngine' -Key 'Executable')
[string]$CheatEngineLoad = (Get-INIValue -Path $INIFile -Section 'CheatEngine' -Key 'AlwaysLoad')
IF ($CheatEngineLoad -eq '1')
{
  [bool]$CheatEngineLoad = $true
}
else 
{
  [bool]$CheatEngineLoad = $false
}
$CheatEngineSystemDir = $CheatEngineDir
$CheatEngineProcess = $CheatEngineExe -Replace '.exe'
$CheatEngine = $CheatEngineDir + '\' + $CheatEngineExe
$CoSMOSName = (Get-INIValue -Path $INIFile -Section 'CoSMOS' -Key 'Name')
$CoSMOSExt = (Get-INIValue -Path $INIFile -Section 'CoSMOS' -Key 'Extension')
$CoSMOSDir = (Get-INIValue -Path $INIFile -Section 'CoSMOS' -Key 'Folder')
$CoSMOSTablesDir = (Get-INIValue -Path $INIFile -Section 'CoSMOS' -Key 'TablesFolder')
$CoSMOSExe = (Get-INIValue -Path $INIFile -Section 'CoSMOS' -Key 'Executable')
[string]$CoSMOSLoad = (Get-INIValue -Path $INIFile -Section 'CoSMOS' -Key 'AlwaysLoad')
IF ($CoSMOSLoad -eq '1')
{
  [bool]$CoSMOSLoad = $true
}
else 
{
  [bool]$CoSMOSLoad = $false
}
$CoSMOSSystemDir = $CoSMOSDir
$CoSMOSProcess = $CoSMOSExe -Replace '.exe'
$CoSMOS = $CoSMOSDir + '\' + $CoSMOSExe
$BorderlessEXE = (Get-INIValue -Path $INIFile -Section 'Borderless' -Key 'BorderlessEXE')
$BorderlessProcess = ($BorderlessEXE.Split('\\')[-1]) -replace('.exe')
$MIDISynthEXE = (Get-INIValue -Path $INIFile -Section 'Audio' -Key 'MIDISynthEXE')


$DefaultAudioDevice = Get-INIValue -Path $INIFile -Section 'Audio' -Key 'DefaultAudio'
$DefaultAudioID = Get-AudioDevice -List |
Where-Object -Property Type -EQ -Value 'Playback' |
Where-Object -Property Name -Like -Value "*$DefaultAudioDevice*" |
Select-Object -ExpandProperty ID
$VRAudioDevice = Get-INIValue -Path $INIFile -Section 'Audio' -Key 'VRAudio'
$VRAudioID = Get-AudioDevice -List |
Where-Object -Property Type -EQ -Value 'Playback' |
Where-Object -Property Name -Like -Value "*$VRAudioDevice*" |
Select-Object -ExpandProperty ID
$HeadphonesDevice = Get-INIValue -Path $INIFile -Section 'Audio' -Key 'Headphones'
$HeadphonesID = Get-AudioDevice -List |
Where-Object -Property Type -EQ -Value 'Playback' |
Where-Object -Property Name -Like -Value "*$HeadphonesDevice*" |
Select-Object -ExpandProperty ID
$DefaultVolume = (Get-INIValue -Path $INIFile -Section 'Audio' -Key 'DefaultVolume')
$VRVolume = (Get-INIValue -Path $INIFile -Section 'Audio' -Key 'VRVolume')
$HeadphonesVolume = (Get-INIValue -Path $INIFile -Section 'Audio' -Key 'HeadphonesVolume')

$VRAudioRecID = Get-AudioDevice -List |
Where-Object -Property Type -EQ -Value 'Recording' |
Where-Object -Property Name -Like -Value "*$VRAudioDevice*" |
Select-Object -ExpandProperty ID
$HeadphonesRecID = Get-AudioDevice -List |
Where-Object -Property Type -EQ -Value 'Recording' |
Where-Object -Property Name -Like -Value "*$HeadphonesDevice*" |
Select-Object -ExpandProperty ID

Set-AudioDevice -ID $VRAudioRecID
Set-AudioDevice -ID $DefaultAudioID
Set-AudioDevice -PlaybackVolume $DefaultVolume

If ($HeadphonesID -ne $null) 
{
  Set-AudioDevice -ID $HeadphonesRecID
  Set-AudioDevice -ID $HeadphonesID
  Set-AudioDevice -PlaybackVolume $HeadphonesVolume  
}

$CurrentGameDocsDir = $GameDocsDir + '\' + $GameName
$LogFileSystemDir = $LogFileDir + '\' + $System

$AllLauncherDirs = $LogFileDir, $ArtMoneyTablesDir, $ArtMoneySystemDir, $TrainersDir, $GameDocsDir, $LogFileSystemDir

ForEach ($AllLauncherDir in $AllLauncherDirs) 
{
  If (!(Test-Path -Path $AllLauncherDir)) 
  {
    mkdir -Path $AllLauncherDir
  }
}

'Settings loaded.' 
'Settings loaded.' | timestamp >> $LogFile

###################### These are AllLauncher specific functions ################################

function Start-Launcher 
{
  ('Starting {0} and finishing up here...' -f $LauncherName) | timestamp >> $LogFile
  $shell = New-Object -ComObject 'Shell.Application'
  $shell.minimizeall()
  Start-Process -FilePath $Launcher -WorkingDirectory $LauncherDir -Verb RunAs
  Start-Sleep -Milliseconds 500
  Set-AudioDevice -ID $VRAudioRecID
  Set-AudioDevice -ID $DefaultAudioID
  Set-AudioDevice -PlaybackVolume $DefaultVolume
  If ($HeadphonesID -ne $null) 
  {
    Set-AudioDevice -ID $HeadphonesRecID
    Set-AudioDevice -ID $HeadphonesID
    Set-AudioDevice -PlaybackVolume $HeadphonesVolume  
  }
  Close-UnneededStuff
  IF ($UsePS4Pad -eq $true) 
  {
    IF (!(Get-Process -Name DS4Windows -ErrorAction SilentlyContinue)) 
    { 
      'Making sure DS4Windows is running'
      Get-PnpDevice -FriendlyName '*game*' -ErrorAction SilentlyContinue | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
      Start-Sleep -Milliseconds 250
      Get-PnpDevice -FriendlyName '*game*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
      Start-Process -FilePath "$DS4Folder\DS4Windows.exe" -WorkingDirectory $DS4Folder -Verb RunAs
    }
  }
  Start-WhitelistedServices
  If ($UseDisplayFusion -eq $true)
  {
    'Preparing display for your launcher...' | timestamp >> $LogFile
    If ($System -eq 'VR') 
    {
      Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "Return screen from VR game"') -Wait
      Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "Ready Launcher VR"') -Wait
    }
    else 
    {
      Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "Return screen from game"') -Wait
      Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "Ready Launcher"') -Wait
    }
  }
  Stop-Process -Name VirtualMIDISynch -Force -ErrorAction SilentlyContinue
  Stop-Process -Name Explorer -Force -ErrorAction SilentlyContinue
  
  If (!(Get-Process -Name explorer)) 
  {
    Start-Process -FilePath $Explorer.Path -ErrorAction SilentlyContinue
  }
  'Finished!' | timestamp >> $LogFile
  
  While ((Get-Process -Name $LauncherProcess -ErrorAction SilentlyContinue | Select-Object -ExpandProperty MainWindowTitle) -ne $LauncherWindow) 
  {
    Start-Sleep -Milliseconds 500
  }
  
  $OtherInstances = Get-Process -Name powershell, AllLauncher, cmd -ErrorAction SilentlyContinue
  ForEach ($OtherInstance in $OtherInstances) 
  {
    If ($OtherInstance.Id -ne $pid) 
    {
      Stop-Process -Id $OtherInstance.Id -Force -ErrorAction SilentlyContinue
    }
  }
  
  'Ending.'
  'Ending.' | timestamp >> $LogFile
  Stop-Transcript -ErrorAction SilentlyContinue
  Get-Process -Name 'powershell', 'alllauncher' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  If ((Get-Process -Id $pid -ErrorAction SilentlyContinue).Name -ne 'powershell_ise') 
  { 
    Stop-Process -Id $pid -Force -ErrorAction SilentlyContinue
    exit
  }
  Stop-Process -Name AllLauncher -Force -ErrorAction SilentlyContinue
  Stop-Process -Name Powershell -Force -ErrorAction SilentlyContinue
  Exit
}
function Quit-Launcher 
{
  'Quitting Launcher...'
  IF ([bool](Get-Process -Name $LauncherProcess -ErrorAction SilentlyContinue)-and($LauncherQuitKey -ne '' -or $null)) 
  {
    ('Using {0} to quit {1}...' -f $LauncherQuitKey, $LauncherName) | timestamp >> $LogFile
    $LauncherQuitKey = $LauncherQuitKey -replace '\+' -replace ',' -replace ' ' -replace 'ctrl', '^' -replace 'shift', '+' -replace 'alt', '%'
    Activate-App -Application $LauncherProcess -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
    Send-Keys -KeysToSend $LauncherQuitKey -ErrorAction SilentlyContinue
    Wait-Process -Name $LauncherProcess -Timeout 3 -ErrorAction SilentlyContinue
  }
  IF ([bool](Get-Process -Name $LauncherProcess -ErrorAction SilentlyContinue))
  {
    ('Closing {1}...' -f $LauncherQuitKey, $LauncherName) | timestamp >> $LogFile
    Get-Process -Name $LauncherProcess -ErrorAction SilentlyContinue | ForEach-Object -Process {
      $_.CloseMainWindow()
    } 
  }  
  Wait-Process -Name $LauncherProcess -Timeout 3
  IF ([bool](Get-Process -Name $LauncherProcess -ErrorAction SilentlyContinue))
  {
    ('Something did not work correctly. Killing {1}...' -f $LauncherQuitKey, $LauncherName)
    Stop-Process -Name $LauncherProcess -Force -ErrorAction SilentlyContinue
  }
}
function Open-GameDocs 
{
  ('Gamedocs: {0}' -f $CurrentGameDocsDir) | timestamp >> $LogFile
  ('Gamedocs: {0}' -f $CurrentGameDocsDir)
  $UnsortedGamePDF = $GameDocsDir + '\' + $GameName + '.pdf'
  'Looking for game docs...' | timestamp >> $LogFile
  'Looking for game docs...'
  If (Test-Path -Path $CurrentGameDocsDir) 
  {
    ('Opening directory {0}.' -f $CurrentGameDocsDir) | timestamp >> $LogFile
    Invoke-Item -Path $CurrentGameDocsDir
    Start-Sleep -Milliseconds 500    
    $GamePDFs = Get-ChildItem -Path $CurrentGameDocsDir -Filter '*.pdf' | Select-Object -ExpandProperty FullName
    ForEach($GamePDF in $GamePDFs) 
    {
      ('Opening {0}.' -f $GamePDF) | timestamp >> $LogFile
      Invoke-Item -Path $GamePDF
      Start-Sleep -Milliseconds 500
    }
    $GameURLs = Get-ChildItem -Path $CurrentGameDocsDir -Filter '*.url' | Select-Object -ExpandProperty FullName
    ForEach($GameURL in $GameURLs) 
    {
      ('Opening {0}.' -f $GameURL) | timestamp >> $LogFile
      Invoke-Item -Path $GameURL
      Start-Sleep -Milliseconds 500
    }
    $GameUHSs = Get-ChildItem -Path $CurrentGameDocsDir -Filter '*.uhs' | Select-Object -ExpandProperty FullName
    ForEach($GameUHS in $GameUHSs) 
    {
      ('Opening {0}.' -f $GameUHS) | timestamp >> $LogFile
      Invoke-Item -Path $GameUHS
      Start-Sleep -Milliseconds 500
    }
  }
  else 
  {
    New-Item -Path $CurrentGameDocsDir -ItemType Directory
  }
  
  If (Test-Path -Path $UnsortedGamePDF) 
  {
    ('Opening {0}.' -f $UnsortedGamePDF) | timestamp >> $LogFile
    Invoke-Item -Path $UnsortedGamePDF
    Start-Sleep -Milliseconds 500
  }
  If (Get-Process -Name '*pdf*', '*reader*', '*foxit*', '*acro*' -ErrorAction SilentlyContinue)
  {
    Start-Sleep -Seconds 1
    $CalmingTMPProcesses = (Get-Process -Name '*pdf*', '*reader*', '*foxit*', '*acro*' -ErrorAction SilentlyContinue).Name
    ForEach ($CalmingTMPProcess in $CalmingTMPProcesses) 
    {
      Wait-ProcessToCalm -ProcessToCalm $CalmingTMPProcess -CalmThreshold 3 -ErrorAction SilentlyContinue
    }
  }
  Start-Sleep -Seconds 1
}
function Start-Cheats 
{
  'Looking for cheats...' | timestamp >> $LogFile
  If (Test-Path -Path $GameTrainer) 
  {
    ('Opening {0}.' -f $GameTrainer) | timestamp >> $LogFile
    Start-Process -FilePath $GameTrainer -WorkingDirectory $TrainersDir -Verb runas
    Wait-ProcessToCalm -ProcessToCalm $GameName -CalmThreshold 5
    Start-Sleep -Seconds 3
    If (Get-Process -Name '*.tmp'-ErrorAction SilentlyContinue) 
    {
      $CalmingTMPProcess = (Get-Process -Name '*.tmp'-ErrorAction SilentlyContinue).Name
      Wait-ProcessToCalm -ProcessToCalm $CalmingTMPProcess -CalmThreshold 5
      Start-Sleep -Seconds 3
    }
  }
  If (Test-Path -Path $Additionaltrainer) 
  {
    ('Opening {0}.' -f $Additionaltrainer) | timestamp >> $LogFile
    Start-Process -FilePath $Additionaltrainer -WorkingDirectory $TrainersDir -Verb runas
    Start-Sleep -Seconds 5
    If (Get-Process -Name '*.tmp'-ErrorAction SilentlyContinue) 
    {
      $CalmingTMPProcess = (Get-Process -Name '*.tmp'-ErrorAction SilentlyContinue).Name
      Wait-ProcessToCalm -ProcessToCalm $CalmingTMPProcess -CalmThreshold 5
      Start-Sleep -Seconds 3
    }
  }
  $ArtMoneyTable = $ArtMoneySystemDir + '\' + $GameName + '.' + $ArtMoneyExt
  $GameFileCheatTable = $ArtMoneyTablesDir + '\Files\' + $GameName + '.' + $ArtMoneyExt
  If (Test-Path -Path $GameFileCheatTable) 
  {
    ('Loading {0} with {1}.' -f $GameFileCheatTable, $ArtMoneyName) | timestamp >> $LogFile
    $ArtMoneyArgument = '"' + $GameFileCheatTable + '"'
    Invoke-Item -Path $GameFileCheatTable
    Start-Sleep -Seconds 4
  }  
  If (Test-Path -Path $ArtMoneyTable) 
  {
    ('Loading {0} with {1}.' -f $ArtMoneyTable, $ArtMoneyName) | timestamp >> $LogFile
    $ArtMoney = $ArtMoneyDir + '\' + $ArtMoneyExe
    $ArtMoneyTableDir = $ArtMoneyTablesDir + '\' + $System
    $ArtMoneyArgument = '"' + $ArtMoneyTable + '"'
    Invoke-Item -Path $ArtMoneyTable
    #Start-Process -FilePath $ArtMoney -ArgumentList $ArtMoneyArgument -WorkingDirectory $ArtMoneyTableDir -Verb RunAs
    Wait-ProcessToCalm -ProcessToCalm $ArtMoneyProcess
    Start-Sleep -Seconds 1
    Get-Process -Name $ArtMoneyProcess -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty Id |
    Activate-App -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 750
    Send-Keys -KeysToSend 'n'
    Start-Sleep -Milliseconds 250
  }
  $CheatEngineTable = $CheatEngineTablesDir + '\' + $GameName + '.' + $CheatEngineExt
  If (Test-Path -Path $CheatEngineTable) 
  {
    ('Loading {0} with {1}.' -f $CheatEngineTable, $CheatEngineName) | timestamp >> $LogFile
    Invoke-Item -Path $CheatEngineTable
    Wait-ProcessToCalm -ProcessToCalm $CheatEngineProcess
    Start-Sleep -Seconds 1
  }
  $CoSMOSTable = $CoSMOSTablesDir + '\' + $GameName + '.' + $CoSMOSExt
  If (Test-Path -Path $CoSMOSTable) 
  {
    ('Loading {0} with {1}.' -f $CoSMOSTable, $CoSMOSName) | timestamp >> $LogFile
    Invoke-Item -Path $CoSMOSTable
    Wait-ProcessToCalm -ProcessToCalm $CoSMOSProcess
    Start-Sleep -Seconds 1
  }
  'Cheat-searching completed.' | timestamp >> $LogFile
  Start-Sleep -Seconds 2
  $CheatsReady = 1
}
function Set-DS4 
{
  $DS4Blacklist = Get-Content -Path $DS4BlacklistFile -ErrorAction SilentlyContinue
  Stop-Process -Name 'DS4Windows' -Force -ErrorAction SilentlyContinue
  Start-Sleep -Milliseconds 250
  Get-PnpDevice -FriendlyName '*game*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
  Start-Sleep -Milliseconds 500
  Get-PnpDevice -FriendlyName '*game*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
  
  [bool]$UseDS4 = $true 
  
  If ($DS4Blacklist -contains $System) 
  { 
    $UseDS4 = $false
    "Stopping DS4 because $System does not need it..." 
    "Stopping DS4 because $System does not need it..." | timestamp >> $LogFile
    Stop-Process -Name 'DS4Windows' -Force -ErrorAction SilentlyContinue
  }
  
  If ($System -eq 'Windows') 
  {
    $DS4WindowsBlacklist = Get-Content -Path $DS4WindowsBlacklistFile -ErrorAction SilentlyContinue
    If ($DS4WindowsBlacklist -contains $GameName) 
    { 
      $UseDS4 = $false
      "Stopping DS4 because $GameName does not need it..." 
      "Stopping DS4 because $GameName does not need it..." | timestamp >> $LogFile
      Stop-Process -Name 'DS4Windows' -Force -ErrorAction SilentlyContinue
    }
  }
  
  IF ($UseDS4 -eq $true) 
  {
    'Making sure DS4Windows is running'
    Start-Sleep -Milliseconds 500
    Start-Process -FilePath "$DS4Folder\DS4Windows.exe" -WorkingDirectory $DS4Folder -Verb RunAs
  }
}
function Close-UnneededStuff 
{
  'Close-UnneededStuff:'
  'Minimizing all existing windows.' | timestamp >> $LogFile
  $shell = New-Object -ComObject 'Shell.Application'
  $shell.minimizeall()
  
  'Closing all explorer windows.'
  $a = (New-Object -ComObject Shell.Application).Windows() |
  Where-Object -FilterScript {
    $_.FullName -ne $null
  } |
  Where-Object -FilterScript {
    $_.FullName.toLower().Endswith('\explorer.exe')
  } 
  $a | ForEach-Object -Process {
    $_.Quit()
  }

  'Ending game leftovers...'
  'Ending game leftovers...' | timestamp >> $LogFile
  Stop-Process -Name $GameName -Force -ErrorAction SilentlyContinue
  Stop-Process -Name $TrainerProcess -Force -ErrorAction SilentlyContinue
  Stop-Process -Name $Additionaltrainer -Force -ErrorAction SilentlyContinue
  Get-Process -Name '*.tmp' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  Get-Process -Name '*trainer*' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  Get-Process -Name '*midi*' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  Get-Process -Name '*MIDISynth*' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  
  If ($System -ne 'VR') 
  {
    'Non-VR. Removing VR stuff...'
    'Non-VR. Removing VR stuff...' | timestamp >> $LogFile
    $VRProcesses = 'steamvr_tutorial', 'Steamtours', 'steamvr', 'vrmonitor', 'vrdashboard', 'vrcompositor', 'vrserver', 'OculusVR', 'Home-Win64-Shipping', 'Home2-Win64-Shipping', 'OculusClient'
    
    ForEach ($VRPRocess in $VRProcesses) 
    {
      "Closing $VRPRocess..."
      Get-Process -Name $VRPRocess -ErrorAction SilentlyContinue | ForEach-Object -Process {
        $_.CloseMainWindow()
        Start-Sleep -Milliseconds 250
      }
    }
    While (Get-Process -Name OculusClient -ErrorAction SilentlyContinue) 
    {
      Start-Sleep -Seconds 1
      Get-Process -Name OculusClient -ErrorAction SilentlyContinue | ForEach-Object -Process {
        $_.CloseMainWindow()
      }
      Start-Sleep -Milliseconds 250
    }
    'Making sure VR services are stopped...'
    'Making sure VR services are stopped...' | timestamp >> $LogFile   
    Set-Service -Name OVRLibraryService -StartupType Disabled
    Set-Service -Name OVRService -StartupType Disabled
    Stop-Service -Name OVRLibraryService -Force -ErrorAction SilentlyContinue
    Stop-Service -Name OVRService -Force -ErrorAction SilentlyContinue
    
    $VRProcesses = 'steamvr_tutorial', 'Steamtours', 'steamvr', 'vrmonitor', 'vrdashboard', 'vrcompositor', 'vrserver', 'OculusVR', 'Home-Win64-Shipping', 'Home2-Win64-Shipping', 'OculusClient', 'OVRRedir', 'OVRServiceLauncher', 'OVRServer_x64'
    ForEach ($VRPRocess in $VRProcesses) 
    {
      "Killing $VRPRocess..."
      Stop-Service -Name 'OVRService' -ErrorAction SilentlyContinue
      Get-Process -Name $VRPRocess -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
  }
              
  'Ending unneeded processes...'
  'Ending unneeded processes...' | timestamp >> $LogFile
  Start-Sleep -Milliseconds 250
  ('Getting {0} and {1}...' -f $ProcessBlacklistFile, $ServiceBlacklistFile)
  $ProcessBlacklist = Get-Content -Path $ProcessBlacklistFile -ErrorAction SilentlyContinue
  $ServiceBlacklist = Get-Content -Path $ServiceBlacklistFile -ErrorAction SilentlyContinue
  
  'Closing all blacklisted processes...' | timestamp >> $LogFile
  'Closing all blacklisted processes...'
  ForEach ($BlacklistProcess in $ProcessBlacklist) 
  {
    (' Checking for {0}' -f $BlacklistProcess)
    If ([bool](Get-Process -Name $BlacklistProcess -ErrorAction SilentlyContinue) -eq $true) 
    { 
      ('  Stopping {0}...' -f $BlacklistProcess)
      Stop-Process -Name $BlacklistProcess -ErrorAction SilentlyContinue -Force
    }
  }
  $EverythingClosed = 1
  'Stopping unneeded services...' 
  'Stopping unneeded services...' | timestamp >> $LogFile
  Start-Sleep -Milliseconds 250
  ForEach ($BlacklistService in $ServiceBlacklist) 
  {
    (' Stopping {0}...' -f $BlacklistService)
    Stop-Service -Name $BlacklistService -ErrorAction SilentlyContinue -Force
  }

  'Everything unneeded has been closed successfully.' | timestamp >> $LogFile
}
function Start-WhitelistedServices 
{
  $ServiceWhitelist = (Get-Content -Path $ServiceWhitelistFile -ErrorAction SilentlyContinue)
  'Starting whitelisted services again...' | timestamp >> $LogFile
  ForEach ($WhitelistService in $ServiceWhitelist) 
  {
    Start-Service -Name $WhitelistService -ErrorAction SilentlyContinue
  }
}
function Set-Borderless
{
  $BorderlessGames = Get-Content -Path $BorderlessGamesList -ErrorAction SilentlyContinue
  $BorderlessProcess = ($BorderlessEXE.Split('\\')[-1]) -replace('.exe')
  
  If ($BorderlessGames -contains $GameName) 
  {
    'Activating Borderless Gaming...' | timestamp >> $LogFile
    'Activating Borderless Gaming...'
    Start-Process -FilePath $BorderlessEXE -Verb runas -WindowStyle Minimized
    Wait-ProcessToCalm -ProcessToCalm $BorderlessProcess -CalmThreshold 1
  }
}


################# Here are the different system functions #############################
function Start-Windows 
{
  'Starting a Windows-game...' | timestamp >> $LogFile
  $SpecialSystem = 1
  IF ($GameExt -eq 'htm' -or $GameExt -eq 'html' -or $GameExt -eq 'url') 
  {
    'This seems to be a Steam or Uplay game. Launching appropriate service...' | timestamp >> $LogFile
    'This seems to be a Steam or Uplay game. Launching appropriate service...'
    Invoke-Item -Path $Game
    'Waiting for Steam/UPlay'
    While (!(Get-Process '*steam*', 'upc', 'uplay' -ErrorAction SilentlyContinue)) 
    {
      Start-Sleep -Milliseconds 500
    }
    Start-Sleep -Seconds 10
    Get-Process -Name '*update*' -ErrorAction SilentlyContinue | Wait-Process -ErrorAction SilentlyContinue
    While (!(Get-Process '*steam*', 'upc', 'uplay' -ErrorAction SilentlyContinue)) 
    {
      Start-Sleep -Milliseconds 500
    }
    Start-Sleep -Seconds 10
  }
  IF ($GameExt -eq 'lnk') 
  {
    'It is a shortcut. Let us see, where this goes...'
    $GameEXE = (Get-Shortcut -Path $Game -ErrorAction SilentlyContinue)
    $GamePath = $GameEXE.TargetPath -replace $GameEXE.Target
    $GameProcess = ($GameEXE.Target).Replace('.exe','')
    ('It is {0} at {1},' -f $GameEXE.Target, $GamePath) | timestamp >> $LogFile
    ('It is {0} at {1},' -f $GameEXE.Target, $GamePath) 
    If ($GameEXE.Arguments -eq '' -or $null) 
    {
      'but there do not seem to be any arguments.' | timestamp >> $LogFile
      $GameEXE.Arguments = ' '
    }
    else 
    {
      ('and there is an argument string: {0}' -f $GameEXE.Arguments) | timestamp >> $LogFile
    }
    $ENBInjector = $GamePath + '\ENBInjector.exe'
    If (Test-Path -Path $ENBInjector -ErrorAction SilentlyContinue) 
    {
      'There is ENBInjector present. I will start it.' 
      'There is ENBInjector present. I will start it.' | timestamp >> $LogFile
      Start-Process -FilePath $ENBInjector -WorkingDirectory $GamePath
      Start-Sleep -Milliseconds 500
    }
    
    If ($GameEXE.Target -like '*launcher*')
    {
      'This looks like a launcher.' | timestamp >> $LogFile
      $StartLauncherProcesses = (Get-Process -ErrorAction SilentlyContinue |
        Where-Object -FilterScript {
          $_.MainWindowTitle -ne '' -or $null
        } |
        Where-Object -FilterScript {
          $_.ProcessName -ne 'explorer'
        } |
        Where-Object -FilterScript {
          $_.ProcessName -notlike 'vrmonitor'
        } |
        Where-Object -FilterScript {
          $_.ProcessName -notlike 'oculusclient'
        }|
        Where-Object -FilterScript {
          $_.ProcessName -notlike '*powershell*'
        }|
        Where-Object -FilterScript {
          $_.ProcessName -notlike 'shellexperiencehost'
      })
      Start-Process -FilePath $GameEXE.TargetPath -WorkingDirectory $GamePath -ArgumentList $GameEXE.Arguments -Verb RunAs
      Start-Sleep -Seconds 25
      $StopLauncherProcesses = (Get-Process -ErrorAction SilentlyContinue |
        Where-Object -FilterScript {
          $_.MainWindowTitle -ne '' -or $null
        } |
        Where-Object -FilterScript {
          $_.ProcessName -ne 'explorer'
        } |
        Where-Object -FilterScript {
          $_.ProcessName -notlike 'vrmonitor'
        } |
        Where-Object -FilterScript {
          $_.ProcessName -notlike 'oculusclient'
        }|
        Where-Object -FilterScript {
          $_.ProcessName -notlike '*powershell*'
        }|
        Where-Object -FilterScript {
          $_.ProcessName -notlike 'shellexperiencehost'
      })
      $NewProcesses = (Compare-Object -ReferenceObject $StartLauncherProcesses -DifferenceObject $StopLauncherProcesses -PassThru | Select-Object -ExpandProperty ProcessName)
      If ($NewProcesses -eq $null) 
      {
        Get-Process $GameProcess | Wait-Process -ErrorAction SilentlyContinue
      }
      else
      {
        'Waiting for:'
        $NewProcesses
        Wait-Process -Name $NewProcesses -ErrorAction SilentlyContinue
      }
    }
    else 
    {
      'Launching game...' | timestamp >> $LogFile
      Start-Process -FilePath $GameEXE.TargetPath -WorkingDirectory $GamePath -ArgumentList $GameEXE.Arguments -Verb RunAs
      Start-Sleep -Seconds 15
      Wait-Process -Name $GameProcess -ErrorAction SilentlyContinue
    }
  }
  IF ($GameExt -eq 'exe') 
  {
    'It is an ordinary program file.' | timestamp >> $LogFile
    Start-Process -FilePath $Game -Wait
  }
  
  Get-Process -Name '*origin*' -ErrorAction SilentlyContinue | Wait-Process -ErrorAction SilentlyContinue
  Get-Process -Name '*GalaxyClient*' -ErrorAction SilentlyContinue | Wait-Process -ErrorAction SilentlyContinue
  Get-Process -Name '*upc*' -ErrorAction SilentlyContinue | Wait-Process -ErrorAction SilentlyContinue
  Get-Process -Name '*uplay*' -ErrorAction SilentlyContinue | Wait-Process -ErrorAction SilentlyContinue
  Get-Process -Name '*Steam*' -ErrorAction SilentlyContinue | Wait-Process -ErrorAction SilentlyContinue
  'Quitting game...' | timestamp >> $LogFile
}

function Start-VR
{
  If ($env:OculusBase -eq $null) 
  {
    $env:OculusBase = "$env:ProgramW6432\Oculus\"
  }
  $SpecialSystem = 1
    
  If ((Get-Service -Name OVRService -ErrorAction SilentlyContinue).Status -ne 'Running') 
  { 
    Set-Service -Name OVRLibraryService -StartupType Manual
    Set-Service -Name OVRService -StartupType Manual
    Start-Service -Name OVRService -ErrorAction SilentlyContinue
    While ((Get-Service -Name OVRService -ErrorAction SilentlyContinue).Status -ne 'Running') 
    {
      Start-Sleep -Seconds 2
    }
    Wait-ProcessToCalm -ProcessToCalm OVRServiceLauncher
    Wait-ProcessToCalm -ProcessToCalm OVRServer_x64
    Start-Sleep -Seconds 3
  }
  If (!(Get-Process -Name OculusClient -ErrorAction SilentlyContinue)) 
  { 
    Start-Process -FilePath "$env:OculusBase\Support\oculus-client\OculusClient.exe" -WorkingDirectory "$env:OculusBase\Support\oculus-client" -Verb RunAs
    Start-Sleep -Seconds 3
    Wait-ProcessToCalm -ProcessToCalm OculusClient
    Start-Sleep -Seconds 5
  }
      
  Set-AudioDevice -ID $VRAudioRecID
  Set-AudioDevice -ID $VRAudioID
  Set-AudioDevice -PlaybackVolume $VRVolume
  
  Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "' + (Get-INIValue -Path $INIFile -Section 'DisplayFusion' -Key 'FunctionVR') + '"') -Wait
    
  IF ($GameExt -eq 'lnk') 
  {
    'Checking game shortcut...'
    $GameEXE = (Get-Shortcut -Path $Game -ErrorAction SilentlyContinue)
    $GamePath = $GameEXE.TargetPath -replace $GameEXE.Target
    $steamAPI = $GamePath + 'steam_api64.dll'
    $steamVRPath = 'C:\STEAM\steamapps\common\SteamVR\bin\win64'
    $steamVRServer = $steamVRPath + '\vrserver.exe'
    $steamVR = $steamVRPath + '\vrmonitor.exe'
          
    If (Test-Path -Path $steamAPI) 
    { 
      'SteamVR is needed. Starting it...'
      Start-Sleep -Seconds 3
      Start-Process -FilePath $steamVRServer -WorkingDirectory $steamVRPath -Verb RunAs -ErrorAction SilentlyContinue
      Start-Sleep -Seconds 2
      Wait-ProcessToCalm -ProcessToCalm vrserver
      Start-Process -FilePath $steamVR -WorkingDirectory $steamVRPath -Verb RunAs -ErrorAction SilentlyContinue
      Start-Sleep -Seconds 3
      Wait-ProcessToCalm -ProcessToCalm vrmonitor
    }
  }
  Start-Windows
  
  Get-Process -Name 'vr*' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "' + (Get-INIValue -Path $INIFile -Section 'DisplayFusion' -Key 'Function') + '"')
}

function Start-Amiga 
{
  'Starting an Amiga-game...' | timestamp >> $LogFile
  $SpecialSystem = 1
  $Emulator = Get-Item -Path (((Get-INIValue -Path $INIFile -Section 'Emulators' -Key 'Amiga') -split '.exe')[0] + '.exe')
  $EmulatorDir = $Emulator.Directory.FullName
  $EmulatorProcess = $Emulator.Name.Replace('.exe','')
  $EmulatorArguments = ((Get-INIValue -Path $INIFile -Section 'Emulators' -Key 'Amiga') -split '.exe')[1]
  $EmulatorSaveDir = $EmulatorDir + '\Savestates'
  $GameSave = $EmulatorSaveDir + '\' + $GameName + '.uss'
  $SaveState = $EmulatorSaveDir + '\default.uss'
  Copy-Item -Path $GameSave -Destination $SaveState -Force
  Start-Process -FilePath $Emulator -ArgumentList ('"' + $Game + '"') -Wait
  'Quitting game...' | timestamp >> $LogFile
  Copy-Item -Path $SaveState -Destination $GameSave -Force
}

function Start-Mednafen 
{
  'Starting a Mednafen-game...' | timestamp >> $LogFile
  $SpecialSystem = 1
  Stop-Process -Name 'DS4Windows' -Force -ErrorAction SilentlyContinue
  $Emulator = Get-Item -Path (((Get-INIValue -Path $INIFile -Section 'Emulators' -Key 'Mednafen') -split '.exe')[0] + '.exe')
  $EmulatorDir = $Emulator.Directory.FullName
  $EmulatorProcess = $Emulator.Name.Replace('.exe','')
  $EmulatorArguments = ((Get-INIValue -Path $INIFile -Section 'Emulators' -Key 'Mednafen') -split '.exe')[1]
  Start-Process -FilePath $Emulator -ArgumentList ('"' + $Game + '"') -Wait -Verb RunAs -WorkingDirectory $Emulator.Directory
  'Quitting Mednafen...' | timestamp >> $LogFile
}

function Start-Emulator 
{
  If ($EmulatorINIEntry -eq 'mednafen') 
  {
    'Starting Mednafen...' | timestamp >> $LogFile
    Stop-Process -Name 'DS4Windows' -Force -ErrorAction SilentlyContinue
    $EmulatorINIEntry = Get-INIValue -Path $INIFile -Section 'Emulators' -Key 'Mednafen'
  }
  else
  {
    'Starting a custom emulator...' | timestamp >> $LogFile
  }
  If ($EmulatorINIEntry -notlike '*.exe') 
  {
    'There is no Emulator for this System!' | timestamp >> $LogFile
  }
  else
  { 
    $Emulator = Get-Item -Path (($EmulatorINIEntry -split '.exe')[0] + '.exe')
    $EmulatorDir = $Emulator.Directory.FullName
    $EmulatorProcess = $Emulator.Name.Replace('.exe','')
    $EmulatorArguments = ($EmulatorINIEntry -split '.exe')[1]
    ('Executing: {0} "{1}" {2}' -f $Emulator, $Game, $EmulatorArguments) | timestamp >> $LogFile
    Start-Process -FilePath $Emulator -ArgumentList ('"' + $Game + '"' + $EmulatorArguments) -Wait -Verb RunAs -WorkingDirectory $Emulator.Directory
    'Quitting emulator...' | timestamp >> $LogFile
  }
}

function Start-ScummVM 
{
  'Starting a ScummVM-game...' | timestamp >> $LogFile
  $SpecialSystem = 1
  $EmulatorDir = "$env:HOMEDRIVE\ScummVM"
  $EmulatorProcess = 'scummvm'
  $Emulator = $EmulatorDir + '\' + $EmulatorProcess + '.exe'
  
  $ScummVMSaveDir = 'G:\ScummVM-Files\SAVES'
  $ScummVMPath = ($Game -split '\.')[0]
  $ScummVMGame = Get-Content -Path $Game
  $ScummVMArguments = ('--fullscreen --aspect-ratio --subtitles --debuglevel=0 --no-console --savepath="{0}" --path="{1}" {2}' -f $ScummVMSaveDir, $ScummVMPath, $ScummVMGame)

  Start-Sleep -Seconds 1
  Start-Process -FilePath $Emulator -ArgumentList $ScummVMArguments -WorkingDirectory $EmulatorDir -Wait -Verb RunAs
  'Quitting game...' | timestamp >> $LogFile
}

function Start-C64
{
  'Starting a C64-game...' | timestamp >> $LogFile
  $SpecialSystem = 1
  $C64Volume = (Get-AudioDevice -PlaybackVolume).replace('%','')
  If ($C64Volume -gt 30) 
  {
    $C64Volume = $C64Volume - 10
  }
  else 
  {
    $C64Volume = $C64Volume - 5
  }
  Set-AudioDevice -PlaybackVolume $C64Volume
  $Emulator = Get-Item -Path (((Get-INIValue -Path $INIFile -Section 'Emulators' -Key 'C64') -split '.exe')[0] + '.exe')
  $EmulatorDir = $Emulator.Directory.FullName
  $EmulatorProcess = $Emulator.Name.Replace('.exe','')
  $EmulatorArguments = ((Get-INIValue -Path $INIFile -Section 'Emulators' -Key 'C64') -split '.exe')[1]
  $ROMPath = 'G:\C64-Games'
  $EmulatorSaveDir = $EmulatorDir + '\C64'
  $GameSaveDir = $EmulatorSaveDir + '\SAVE\' + $GameName
  $Arguments = '-fullscreen -autostart'
  Remove-Item -Path ('{0}\*.vsf' -f $EmulatorSaveDir) -Force
  Copy-Item -Path ('{0}\*.vsf' -f $GameSaveDir) -Destination $EmulatorSaveDir -Force -ErrorAction SilentlyContinue
  $C64ROM = Get-ChildItem -Path $ROMPath -Filter ('{0}.*' -f $GameName) -Exclude '*.vfl' -Recurse | Select-Object -ExpandProperty FullName
  If (($C64ROM -split '\.')[-1] -eq 'crt')
  {
    $Arguments = '-fullscreen -cartcrt "' + $Game + '"'
  }
  If (($C64ROM -split '\.')[-1] -eq 'bin')
  {
    $Arguments = '-fullscreen -cartgmod2 "' + $Game + '"'
  }
  If (($C64ROM -split '\.')[-1] -eq 't64')
  {
    $Arguments = '-fullscreen -autostart "' + $C64ROM + '"'
  }
  If (($C64ROM -split '\.')[-1] -eq 'tap')
  {
    $Arguments = '-fullscreen -autostart "' + $C64ROM + '"'
  }
  If (($C64ROM -split '\.')[-1] -eq 'd64')
  {
    $Arguments = '-fullscreen -autostart "' + $C64ROM + '" -flipname "' + $Game + '"'
  }
  Start-Process -FilePath $Emulator -ArgumentList $Arguments -WorkingDirectory $EmulatorDir -Wait -Verb RunAs
  'Quitting game...' | timestamp >> $LogFile
  Copy-Item -Path ('{0}\*.vsf' -f $EmulatorSaveDir) -Destination $GameSaveDir -Force
  Remove-Item -Path ('{0}\*.vsf' -f $EmulatorSaveDir) -Force
}

function Start-DOS 
{
  'Starting a DOS-game...' | timestamp >> $LogFile
  $SpecialSystem = 1
  $DOSGame = Get-Shortcut -Path $Game
  $SaveDir = 'D:\EMULATOREN\DOSBox\captures\save'
  $GameSaveDir = ('{0}\{1}' -f $SaveDir, $GameName)
  $Emulator = ($DOSGame.Target -replace '.exe')
  mkdir -Path $GameSaveDir
  Copy-Item -Path ('{0}\*.sav' -f $GameSaveDir) -Destination $SaveDir -Force -ErrorAction SilentlyContinue
  ('Starting DOS-Box with {0}' -f $DOSGame.Arguments) | timestamp >> $LogFile
  $DOSVolume = (Get-AudioDevice -PlaybackVolume).replace('%','')
  If ($DOSVolume -gt 30) 
  {
    $DOSVolume = $DOSVolume - 15
  }
  else
  {
    $DOSVolume = $DOSVolume - 5
  }
  Set-AudioDevice -PlaybackVolume $DOSVolume
  Start-Sleep -Seconds 1
  Start-Process -FilePath $DOSGame.TargetPath -ArgumentList $DOSGame.Arguments -WorkingDirectory ($DOSGame.TargetPath -replace $DOSGame.Target) -Verb RunAs
  Start-Sleep -Seconds 10
  Wait-Process -Name $Emulator
  Wait-Process -Name 'DOSBox' -ErrorAction SilentlyContinue
  Wait-Process -Name 'dosbox_x64' -ErrorAction SilentlyContinue
  Wait-Process -Name '*dosbox*' -ErrorAction SilentlyContinue
  Start-Sleep -Milliseconds 500
  'Quitting DOS game...' | timestamp >> $LogFile
  Copy-Item -Path ('{0}\*.sav' -f $SaveDir) -Destination $GameSaveDir -Force -ErrorAction SilentlyContinue
}

function Start-DOSBox 
{
  Start-DOS
}

function Start-InteractiveFiction
{
  'Starting an Interactive Fiction-game...' | timestamp >> $LogFile
  $SpecialSystem = 1
  $IFGame = Get-Shortcut -Path $Game
  Start-Sleep -Seconds 2
  If ($IFGame.Arguments -eq '' -or $null) 
  {
    'There do not seem to be any arguments.'
    $IFGame.Arguments = ' '
  }
  Start-Process -FilePath $IFGame.TargetPath -ArgumentList $IFGame.Arguments -WorkingDirectory ($IFGame.TargetPath -replace $IFGame.Target)
  Start-Sleep -Seconds 5
  Wait-Process -Name ($IFGame.Target -replace '.exe') -ErrorAction SilentlyContinue
  'Quitting game...' | timestamp >> $LogFile
}

function Start-NDS 
{
  'Starting a Nintendo-DS-game...' | timestamp >> $LogFile
  $EmulatorDir = 'D:\EMULATOREN\DS'
  $EmulatorProcess = 'DeSmuME-x64'
  $Emulator = $EmulatorDir + '\' + $EmulatorProcess + '.exe'
  $Arguments = '"' + $Game + '"'
  Start-Process -FilePath $Emulator -ArgumentList $Arguments -WorkingDirectory $EmulatorDir -Verb RunAs
  While (!(Get-Process -Name $EmulatorProcess  -ErrorAction SilentlyContinue)) 
  {
    Start-Sleep -Milliseconds 500
  }
  Start-Sleep -Seconds 1
  Send-Keys -KeysToSend '%{ENTER}'
  Start-Sleep -Milliseconds 250
  Wait-Process -Name $EmulatorProcess
  'Quitting game...' | timestamp >> $LogFile
}

function Start-PS2 
{
  'Starting a PS2-game...' | timestamp >> $LogFile
  $SpecialSystem = 1
  $EmulatorDir = 'D:\EMULATOREN\PS2'
  $EmulatorProcess = 'pcsx2'
  $Emulator = $EmulatorDir + '\' + $EmulatorProcess + '.exe'
  $EmulatorSaveDir = $EmulatorDir + '\_MemoryCard'
  $GameSave = $EmulatorSaveDir + '\' + $GameName + '.ps2'
  $EmptySave = $EmulatorSaveDir + '\Empty.ps2'
  $MemoryCard = $EmulatorDir + '\memcards\Mcd001.ps2'
  $GameConfig = $EmulatorDir + '\inis\' + $GameName + '.ini'
  $OriginalConfig = $EmulatorDir + '\inis\LilyPad.ini'
  If (Test-Path -Path $GameConfig) 
  {
    $RenameConfigBack = $true
    Rename-Item -Path $OriginalConfig -NewName ($OriginalConfig + '.bak') -Force
    Rename-Item -Path $GameConfig -NewName $OriginalConfig -Force
  } 
  If (Test-Path -Path $GameSave) 
  {
    Copy-Item -Path $GameSave -Destination $MemoryCard -Force
  }
  else 
  {
    Copy-Item -Path $EmptySave -Destination $MemoryCard -Force
  }
  Start-Process -FilePath $Emulator -ArgumentList ('"' + $Game + '" --fullscreen --fullboot') -Wait -Verb RunAs
  'Quitting game...' | timestamp >> $LogFile
  If ($RenameConfigBack -eq $true) 
  {
    Rename-Item -Path $OriginalConfig -NewName $GameConfig -Force
    Rename-Item -Path ($OriginalConfig + '.bak') -NewName $OriginalConfig -Force  
  }
  Copy-Item -Path $MemoryCard -Destination $GameSave -Force
}

function Start-Dolphin
{
  'Starting a Dolphin-game...' | timestamp >> $LogFile
  $SpecialSystem = 1
  $EmulatorDir = 'D:\EMULATOREN\Gamecube'
  $EmulatorProcess = 'Dolphin'
  $Emulator = $EmulatorDir + '\' + $EmulatorProcess + '.exe'
  $EmulatorSaveDir = $env:USERPROFILE + '\Documents\Dolphin Emulator\GC'
  $GameSave = $EmulatorSaveDir + '\' + $GameName + '.raw'
  $EmptySave = $EmulatorSaveDir + '\Empty.raw'
  $MemoryCard = $EmulatorSaveDir + '\MemoryCardA.USA.raw'
  $GameConfig = $env:USERPROFILE + '\Documents\Dolphin Emulator\Config\' + $GameName + '.ini'
  $OriginalConfigGC = $env:USERPROFILE + '\Documents\Dolphin Emulator\Config\GCPadNew.ini'
  $OriginalConfigWii = $env:USERPROFILE + '\Documents\Dolphin Emulator\Config\WiimoteNew.ini'
  If (Test-Path -Path $GameConfig) 
  {
    $RenameConfigBack = $true
    Rename-Item -Path $OriginalConfigGC -NewName ($OriginalConfigGC + '.bak') -Force
    Rename-Item -Path $OriginalConfigWii -NewName ($OriginalConfigWii + '.bak') -Force
    Rename-Item -Path $GameConfig -NewName $OriginalConfigGC -Force
    Rename-Item -Path $GameConfig -NewName $OriginalConfigWii -Force
  } 
  If (Test-Path -Path $GameSave) 
  {
    Copy-Item -Path $GameSave -Destination $MemoryCard -Force
  }
  else 
  {
    Copy-Item -Path $EmptySave -Destination $MemoryCard -Force
  }
  Start-Process -FilePath $Emulator -ArgumentList ('/e "' + $Game + '"') -Wait
  'Quitting game...' | timestamp >> $LogFile
  If ($RenameConfigBack -eq $true) 
  {
    Rename-Item -Path $OriginalConfigWii -NewName $GameConfig -Force
    Rename-Item -Path ($OriginalConfigGC + '.bak') -NewName $OriginalConfigGC -Force  
    Rename-Item -Path ($OriginalConfigWii + '.bak') -NewName $OriginalConfigWii -Force  
  }
  Copy-Item -Path $MemoryCard -Destination $GameSave -Force
}

function Start-RetroArch
{
  'Starting a RetroArch-game...' | timestamp >> $LogFile
  Stop-Process -Name 'DS4Windows' -Force -ErrorAction SilentlyContinue
  $Emulator = Get-Item -Path (((Get-INIValue -Path $INIFile -Section 'Emulators' -Key 'RetroArch') -split '.exe')[0] + '.exe')
  $EmulatorDir = $Emulator.Directory.FullName
  $EmulatorProcess = $Emulator.Name.Replace('.exe','')
  $EmulatorArguments = ((Get-INIValue -Path $INIFile -Section 'Emulators' -Key 'RetroArch') -split '.exe')[1]
  $Arguments = '-L "' + $EmulatorDir + '\cores\' + $EmulatorINIEntry + '" "' + $Game + '"'
  ('Using Core: {0}' -f $EmulatorINIEntry) | timestamp >> $LogFile
  'Starting RetroArch'
  ('Executing: {0} {1} {2}' -f $Emulator, $Game, $Arguments) | timestamp >> $LogFile
  Start-Process -FilePath $Emulator -ArgumentList $Arguments -WorkingDirectory $Emulator.Directory -Verb RunAs -Wait
  'Finished'
  'Quitting game...' | timestamp >> $LogFile
}
######################################################################################


If ($Menu -eq $true) 
{
  'The resetting parameter to return back to the menu has been used.' | timestamp >> $LogFile
  'Loading your Launcher and exiting AllLauncher.' | timestamp >> $LogFile
  '-menu has been used: Starting Launcher'
  $System = 'Windows'
  $Game = 'AllLauncher Menu'
  Start-Launcher
  Start-Sleep -Seconds 360
}

Stop-Process -Name 'DS4Windows' -Force -ErrorAction SilentlyContinue
Hide-DesktopIcons

'Writing INI setting for this game'
Set-INIValue -Path $INIFile -Section 'LastGame' -Key 'System' -Value $System
Set-INIValue -Path $INIFile -Section 'LastGame' -Key 'Game' -Value $Game
Set-INIValue -Path $INIFile -Section $System -Key 'Game' -Value $Game

'Checking for PS4 Gamepad'
Start-Sleep -Milliseconds 100
If ($UsePS4Pad -eq $true) 
{
  Set-DS4
}

Close-UnneededStuff
'Everything is closed!'
Hide-DesktopIcons

Stop-Process -Name Explorer -ErrorAction SilentlyContinue
    
Hide-DesktopIcons
  
If ($LauncherQuit -eq $true) 
{
  If ([bool](Get-Process -Name $LauncherProcess -ErrorAction SilentlyContinue) -eq $true) 
  {
    Quit-Launcher
  }
  Wait-Process -Name $LauncherProcess -Timeout 15 -ErrorAction SilentlyContinue
}

If ($System -eq 'VR') 
{
  'This is a VR game. Making sure hardware is ready...'
  'This is a VR game. Making sure hardware is ready...' | timestamp >> $LogFile
  If ($env:OculusBase -eq $null) 
  {
    $env:OculusBase = "$env:ProgramW6432\Oculus\"
  }
  
  $VRProcesses = 'steamvr_tutorial', 'Steamtours', 'steamvr', 'vrmonitor', 'vrdashboard', 'vrcompositor', 'vrserver', 'Home-Win64-Shipping', 'Home2-Win64-Shipping'
  ForEach ($VRPRocess in $VRProcesses) 
  {
    Get-Process -Name $VRPRocess -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  }
  
  Remove-Item -Path "$env:LOCALAPPDATA\openvr" -Force -Recurse -ErrorAction SilentlyContinue
  
  If ((Get-Service -Name OVRService -ErrorAction SilentlyContinue).Status -ne 'Running') 
  { 
    'Oculus Service not running. Starting it...'
    'Oculus Service not running. Starting it...' | timestamp >> $LogFile
    Get-PnpDevice -FriendlyName '*rift*', '*oculus*' -ErrorAction SilentlyContinue | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
    Get-PnpDevice -FriendlyName '*rift*', '*oculus*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
    Set-Service -Name OVRLibraryService -StartupType Manual -ErrorAction SilentlyContinue
    Set-Service -Name OVRService -StartupType Manual -ErrorAction SilentlyContinue
    Get-PnpDevice -FriendlyName '*rift*', '*oculus*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
    Start-Service -Name OVRService -ErrorAction SilentlyContinue
    While ((Get-Service -Name OVRService -ErrorAction SilentlyContinue).Status -ne 'Running') 
    {
      Start-Sleep -Seconds 3
    }
    Wait-ProcessToCalm -ProcessToCalm OVRServiceLauncher
    Wait-ProcessToCalm -ProcessToCalm OVRServer_x64
    Start-Sleep -Seconds 5
    #Start-Service -Name OVRLibraryService -ErrorAction SilentlyContinue
  }
  If (!(Get-Process -Name OculusClient -ErrorAction SilentlyContinue)) 
  { 
    "Oculus Client not running. Starting $env:OculusBase\Support\oculus-client\OculusClient.exe"
    "Oculus Client not running. Starting $env:OculusBase\Support\oculus-client\OculusClient.exe" | timestamp >> $LogFile
    Start-Service -Name OVRService -ErrorAction SilentlyContinue
    Start-Process -FilePath "$env:OculusBase\Support\oculus-client\OculusClient.exe" -WorkingDirectory "$env:OculusBase\Support\oculus-client" -Verb RunAs
    Start-Sleep -Seconds 5
    Wait-ProcessToCalm -ProcessToCalm OculusClient
    Start-Sleep -Seconds 7
  }
}
Hide-DesktopIcons

$ProcessesBefore = Get-Process

If (Test-Path -Path $MIDISynthEXE) 
{
  $MIDISynthSystems = Get-Content -Path $MIDISynthSystemsFile -ErrorAction SilentlyContinue
  If ($MIDISynthSystems -contains $System) 
  { 
    "$System uses a MIDISynth: Starting it..." 
    "$System uses a MIDISynth: Starting it..." | timestamp >> $LogFile
    Start-Process -FilePath $MIDISynthEXE -ErrorAction SilentlyContinue
  }
  
  $MIDISynthGames = Get-Content -Path $MIDISynthGamesFile -ErrorAction SilentlyContinue
  If ($MIDISynthGames -contains $GameName) 
  {
    "$GameName uses a MIDISynth: Starting it..." 
    "$GameName uses a MIDISynth: Starting it..." | timestamp >> $LogFile
    Start-Process -FilePath $MIDISynthEXE -ErrorAction SilentlyContinue
  }
}

If ($UseCheats -eq $true) 
{
  Start-Cheats
  'Cheats ready'
}

Open-GameDocs
'Docs ready'

If ($UseDisplayFusion -eq $true)
{
  'Using DisplayFusion...' | timestamp >> $LogFile
  Start-Sleep -Milliseconds 500
  If ($System -eq 'VR') 
  {
    Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "' + (Get-INIValue -Path $INIFile -Section 'DisplayFusion' -Key 'Function') + '"') -Wait
  }
  else
  {
    Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "' + (Get-INIValue -Path $INIFile -Section 'DisplayFusion' -Key 'Function') + '"') -Wait
  }
  Start-Sleep -Seconds 3
  Get-Process -Name DisplayFusion* -ErrorAction SilentlyContinue |
  Where-Object -FilterScript {
    $_.MainWindowTitle -ne ''
  } |
  Select-Object -Property MainWindowTitle |
  ForEach-Object -Process {
    $_.CloseMainWindow()
    $wshell.AppActivate('DisplayFusion')
    Start-Sleep -Milliseconds 100
    $wshell.Sendkeys('{enter}')
  }
}

#Get-Process -Name 'explorer' |
#Select-Object -ExpandProperty Id |
#Activate-App
Hide-DesktopIcons


#Start-Sleep -Seconds 1
$StartDate = Get-Date

'Getting ready for:' | timestamp >> $LogFile


########## Here starts the game ################
''
''
'Getting ready for the game!'
'' | Out-File $LogFile -Append
('{0}' -f $GameName) | timestamp >> $LogFile
('Going to {0}' -f $StartCurrentSystem)
Try 
{
  "Trying $StartCurrentSystem..."
  . ($StartCurrentSystem)
}
Catch 
{
  If ($SpecialSystem -ne 1) 
  { 
    'Using emulator settings from ini...'
    If ($EmulatorINIEntry -like '*libretro.dll') 
    {
      'Trying RetroArch...'
      . Start-RetroArch
    }
    else
    {
      'Trying emulator from ini...'
      . Start-Emulator
    }
  }
}
'' | Out-File $LogFile -Append
('{0} has ended.' -f $GameName) | timestamp >> $LogFile
''
''
################################################


$EndDate = Get-Date

Start-Sleep -Milliseconds 500
$ProcessesAfter = Get-Process

$ProcessesChanged = Compare-Object -ReferenceObject $ProcessesAfter -DifferenceObject $ProcessesBefore -PassThru | Select-Object -Property Name, Id

ForEach ($ChangedProcess in $ProcessesChanged) 
{ 
  If ($ChangedProcess.Name -ne 'explorer') 
  { 
    Get-Process -Id $ChangedProcess.Id -ErrorAction SilentlyContinue | ForEach-Object -Process {
      $_.CloseMainWindow()
    } 
    Start-Sleep -Milliseconds 200
  }
}
Start-Sleep -Milliseconds 500
Stop-Process -Name $GameName -Force -ErrorAction SilentlyContinue
Stop-Process -Name $TrainerProcess -Force -ErrorAction SilentlyContinue
Stop-Process -Name $Additionaltrainer -Force -ErrorAction SilentlyContinue
Stop-Process -Name $BorderlessProcess -Force -ErrorAction SilentlyContinue
Get-Process -Name '*.tmp' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Get-Process -Name '*trainer*' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Get-Process -Name "*$GameName*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Get-Process -Name '*Borderless*' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

ForEach ($ChangedProcess in $ProcessesChanged) 
{ 
  If ($ChangedProcess.Name -ne 'explorer') 
  { 
    Get-Process -Id $ChangedProcess.Id -ErrorAction SilentlyContinue | ForEach-Object -Process {
      $_.CloseMainWindow()
      Start-Sleep -Milliseconds 100
    }
  }
}

$Playtime = $EndDate - $StartDate
'Calculating time played...' | timestamp >> $LogFile
If ($Playtime.TotalMinutes -gt 3) 
{
  ('Played for {0} minutes.' -f $Playtime.TotalMinutes) | timestamp >> $LogFile
  $LogFileName = $LogFileDir + '\Complete_Game_Log_' + $StartDate.Year + '.txt'
  $GameLogName = $LogFileSystemDir + '\' + $GameName + '.txt'
  If (Test-Path -Path $GameLogName)
  {
    $GameLog = Import-Csv -Path $GameLogName -Delimiter "`t"
    If ($GameLog[-1].Total -ne $null) 
    {
      $LastTime = ([TimeSpan]($GameLog[-1].Total))
    }
    else 
    {
      $LastTime = ([TimeSpan]'00:00')
    }
  }
  else 
  {
    $LastTime = ([TimeSpan]'00:00')
  }
  $ThisDate = (Get-Date -Date $StartDate.Date -Format 'ddd, dd. MMM yyyy')
  $ThisStart = ('{0:HH}:{0:mm}' -f $StartDate)
  $ThisEnd = ('{0:HH}:{0:mm}' -f $EndDate)
  $ThisPlayed = ('{0}:{1}' -f $Playtime.Hours.ToString('00'), ($Playtime.Minutes).ToString('00'))
  $TotalTime = ([TimeSpan]$LastTime + [TimeSpan]$ThisPlayed)
  $ThisTotal = ('{0}:{1}' -f $TotalTime.Hours.ToString('00'), $TotalTime.Minutes.ToString('00'))
  $ThisGameName = $GameName
  $ThisSystem = $System
  ('Complete playtime: {0}.' -f $ThisTotal) | timestamp >> $LogFile
  $GameLogHash = [ordered]@{
    Date   = $ThisDate
    Start  = $ThisStart
    End    = $ThisEnd
    Played = $ThisPlayed
    Total  = $ThisTotal
  }
  New-Object  -TypeName PSObject -Property $GameLogHash | Export-Csv -Path $GameLogName -Append -Delimiter "`t" -Force -NoTypeInformation
  $CompleteLogHash = [ordered]@{
    Date     = $ThisDate
    Start    = $ThisStart
    End      = $ThisEnd
    Played   = $ThisPlayed
    GameName = $ThisGameName
    System   = $ThisSystem
  }
  New-Object  -TypeName PSObject -Property $CompleteLogHash | Export-Csv -Path $LogFileName -Delimiter "`t" -Append -Force -NoTypeInformation
}

Set-AudioDevice -ID $VRAudioRecID
Set-AudioDevice -ID $DefaultAudioID
Set-AudioDevice -PlaybackVolume $DefaultVolume
If ($HeadphonesID -ne $null) 
{
  Set-AudioDevice -ID $HeadphonesRecID
  Set-AudioDevice -ID $HeadphonesID
  Set-AudioDevice -PlaybackVolume $HeadphonesVolume  
}
If ((Get-ChildItem -Path $CurrentGameDocsDir -Recurse | Measure-Object).Count -eq 0) 
{
  Remove-Item -Path $CurrentGameDocsDir -Force -ErrorAction SilentlyContinue
}


If ($WasLauncherRunning -eq $true) 
{
  ('Returning to {0}.' -f $LauncherName) | timestamp >> $LogFile
  Start-Launcher
}
else 
{ 
  'Not started from Launcher, so not going back' | timestamp >> $LogFile
  Close-UnneededStuff
  Start-WhitelistedServices  
  If ($UseDisplayFusion -eq $true)
  {
    'Preparing display for your launcher...' | timestamp >> $LogFile
    If ($System -eq 'VR') 
    {
      Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "Return screen from VR game"') -Wait
    }
    else 
    {
      Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "Return screen from game"') -Wait
    }
  }
  Stop-Process -Name Explorer -ErrorAction SilentlyContinue
  
  If (!(Get-Process -Name explorer)) 
  {
    Start-Process -FilePath $Explorer.Path -ErrorAction SilentlyContinue
  }
  Stop-Transcript -ErrorAction SilentlyContinue
  ('Everything went well, removing transcript {0}.' -f $Transcript) | timestamp > $LogFile
  Remove-Item -Path $Transcript -Force
  Get-Process -Name 'powershell', 'alllauncher' -ErrorAction SilentlyContinue | Stop-Process -Force
  Exit
}
