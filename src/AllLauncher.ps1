param
(
  [Parameter(Position = 0)]
  [string]
  $System,
  [Parameter(Position = 1)]
  [string]
  $Game,
  [Parameter]
  [Alias('Reset','Back','Return','Launcher')]
  [switch]
  $Menu,
  [Parameter]
  [Alias('Last','Replay','Again','Previous','PreviousGame')]
  [switch]
  $LastGame
)

If ((($System -eq $null) -or ($System -eq '')) -and ($LastGame -ne $true))
{
  [bool]$Menu = $true
}
If ($LastGame -eq $true) 
{
  [bool]$Menu = $false
}

Clear-Host
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue

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
$IgnoreProcessesFile = ('{0}\IgnoreProcesses.txt' -f $CFGDir)
$IgnoreDirectoriesFile = ('{0}\IgnoreDirectories.txt' -f $CFGDir)
$DS4BlacklistFile = ('{0}\DS4SystemBlacklist.txt' -f $CFGDir)
$DS4WindowsBlacklistFile = ('{0}\DS4WindowsBlacklist.txt' -f $CFGDir)
$HOTASGamesFile = ('{0}\HOTASGames.txt' -f $CFGDir)
$BorderlessGamesList = ('{0}\BorderlessGamesList.txt' -f $CFGDir)
$MIDISynthGamesFile = ('{0}\MIDISynthGames.txt' -f $CFGDir)
$MIDISynthSystemsFile = ('{0}\MIDISynthSystems.txt' -f $CFGDir)


#Remove-Item -Path $Transcript -Force -ErrorAction SilentlyContinue
Stop-Process -Name DS4Windows -Force -ErrorAction SilentlyContinue
$shell = New-Object -ComObject 'Shell.Application'


'Starting Log!'
'Starting Log!' | timestamp > $LogFile

Remove-Item -Path $Transcript -Force -ErrorAction SilentlyContinue
Start-Transcript -Path $Transcript -Force -ErrorAction SilentlyContinue
'Transcript started.'
('Writing transcript to {0}.' -f $Transcript) | timestamp > $LogFile

'Loading general functions...'
#These are general functions
$AudioLocation = ($profile | Split-Path) + '\Modules\AudioDeviceCmdlets'
If (!(Test-Path -Path $AudioLocation)) 
{
  New-Item $AudioLocation -Type directory -Force -ErrorAction SilentlyContinue
}
If (!(Test-Path -Path "$AudioLocation\AudioDeviceCmdlets.dll")) 
{
  Copy-Item -Path "$CurrentDir\DLLs\AudioDeviceCmdlets.dll" -Destination "$AudioLocation\AudioDeviceCmdlets.dll" -Force -ErrorAction SilentlyContinue
}
Set-Location -Path $AudioLocation -ErrorAction SilentlyContinue
Get-ChildItem | Unblock-File
Import-Module -Name AudioDeviceCmdlets -ErrorAction SilentlyContinue
Set-Location -Path $CurrentDir
$wshell = New-Object -ComObject wscript.shell
$CSharpSig = @'
[DllImport("user32.dll", EntryPoint = "SystemParametersInfo")]
public static extern bool SystemParametersInfo(
                 uint uiAction,
                 uint uiParam,
                 uint pvParam,
                 uint fWinIni);
'@
$CursorRefresh = Add-Type -MemberDefinition $CSharpSig -Name WinAPICall -Namespace SystemParamInfo -PassThru -ErrorAction SilentlyContinue
$CursorRefresh::SystemParametersInfo(0x0057,0,$null,0)
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
$Displays = [System.Windows.Forms.Screen]::AllScreens
$PrimaryMonitor = $Displays[0].Bounds
$SecondaryMonitor = $Displays[1].Bounds
$PrimaryScreen = $Displays[0].WorkingArea
$SecondaryScreen = $Displays[1].WorkingArea

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
Function Get-WindowSizeAndPos 
{
  [OutputType('System.Automation.WindowInfo')]
  [cmdletbinding()]
  Param (
    [parameter(ValueFromPipelineByPropertyName = $true)]
    $ProcessName,
    [parameter(ValueFromPipelineByPropertyName = $true)]
    $MainWindowHandle
  )
  Begin {
    Try
    {
      [void][Window]
    }
    Catch 
    {
      Add-Type -TypeDefinition @"
              using System;
              using System.Runtime.InteropServices;
              public class Window {
                [DllImport("user32.dll")]
                [return: MarshalAs(UnmanagedType.Bool)]
                public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
              }
              public struct RECT
              {
                public int Left;        // x position of upper-left corner
                public int Top;         // y position of upper-left corner
                public int Right;       // x position of lower-right corner
                public int Bottom;      // y position of lower-right corner
              }
"@
    }
  }
  Process {        
    If (($MainWindowHandle -eq $null) -or ($MainWindowHandle -eq '') -or ($MainWindowHandle -eq 0)) 
    {
      $Handles = (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue ).MainWindowHandle
    }
    else 
    {
      $Handles = $MainWindowHandle
    }          
    ForEach ($Handle in $Handles) 
    {
      #$Handle = $_.MainWindowHandle
      $Rectangle = New-Object -TypeName RECT
      $Return = [Window]::GetWindowRect($Handle,[ref]$Rectangle)
      If ($Return) 
      {
        $Height = $Rectangle.Bottom - $Rectangle.Top
        $Width = $Rectangle.Right - $Rectangle.Left
        $Size = New-Object -TypeName System.Management.Automation.Host.Size -ArgumentList $Width, $Height
        $TopLeft = New-Object -TypeName System.Management.Automation.Host.Coordinates -ArgumentList $Rectangle.Left, $Rectangle.Top
        $BottomRight = New-Object -TypeName System.Management.Automation.Host.Coordinates -ArgumentList $Rectangle.Right, $Rectangle.Bottom
        $Object = [pscustomobject]@{
          ProcessName = $ProcessName
          Handle      = $Handle
          Size        = $Size
          TopLeft     = $TopLeft
          BottomRight = $BottomRight
        }
        $Object.PSTypeNames.insert(0,'System.Automation.WindowInfo')
        $Object
      }
    }
  }
}
Function Move-Window 
{
  [OutputType('System.Automation.WindowInfo')]
  [cmdletbinding()]
  Param (
    [parameter(ValueFromPipelineByPropertyName = $true)]
    $ProcessName,
    [int]$X,
    [int]$Y,
    [int]$Width,
    [int]$Height,
    $MainWindowHandle,
    [int]$Screen,
    [switch]$center,
    [switch]$Passthru
  )
  Try
  {
    [void][Window]
  }
  Catch 
  {
    Add-Type -TypeDefinition @"
              using System;
              using System.Runtime.InteropServices;
              public class Window {
                [DllImport("user32.dll")]
                [return: MarshalAs(UnmanagedType.Bool)]
                public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
                [DllImport("User32.dll")]
                public extern static bool MoveWindow(IntPtr handle, int x, int y, int width, int height, bool redraw);
              }
              public struct RECT
              {
                public int Left;
                public int Top;
                public int Right;
                public int Bottom;
              }
"@
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
  }
  
  $Displays = [System.Windows.Forms.Screen]::AllScreens
  $PrimaryMonitor = $Displays[0].Bounds
  $SecondaryMonitor = $Displays[1].Bounds
  $PrimaryScreen = $Displays[0].WorkingArea
  $SecondaryScreen = $Displays[1].WorkingArea
  
  $Rectangle = New-Object -TypeName RECT
  $Handles = (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue).MainWindowHandle
  ForEach ($Handle in $Handles) 
  { 
    #$WindowSizeandPos = Get-WindowSizeAndPos -MainWindowHandle $Handle
    $Rectangle = New-Object -TypeName RECT
    $null = [Window]::GetWindowRect($Handle,[ref]$Rectangle)
    $Return = [Window]::GetWindowRect($Handle,[ref]$Rectangle)
    If (($Rectangle.Left -ge $PrimaryMonitor.Left) -and ($Rectangle.Right -le $PrimaryMonitor.Right))
    {
      $UsePrimaryDisplay = $true
      $IsOnPrimaryDisplay = $true
    } 
    else
    {
      $UsePrimaryDisplay = $false
      $IsOnPrimaryDisplay = $false
    } 
    If ($Screen -eq 1) 
    {
      $UsePrimaryDisplay = $true
    } 
    If ($Screen -eq 2) 
    {
      $UsePrimaryDisplay = $false
    } 

    If ($IsOnPrimaryDisplay -eq $true)
    {
      [int]$XPosPercent = (100/$PrimaryScreen.Size.Width)*$Rectangle.Left
      [int]$YPosPercent = (100/$PrimaryScreen.Size.Height)*$Rectangle.Top
    }
    else
    {
      [int]$XPosPercent = (100/$SecondaryScreen.Size.Width)*$Rectangle.Left
      [int]$YPosPercent = (100/$SecondaryScreen.Size.Height)*$Rectangle.Top
    }
    If ([int]$XPosPercent -lt 0) 
    {
      [int]$XPosPercent = 100 + [int]$XPosPercent
    }
    If ([int]$YPosPercent -lt 0) 
    {
      [int]$YPosPercent = 100 + [int]$YPosPercent
    }
    If ($center -eq $true)
    {
      If ($UsePrimaryDisplay -eq $true)
      { 
        [int]$NewPosX = ($PrimaryScreen.Left + (($PrimaryScreen.Size.Width - ($Rectangle.Right - $Rectangle.Left))/2))
        [int]$NewPosY = ($PrimaryScreen.Top + (($PrimaryScreen.Size.Height - ($Rectangle.Bottom - $Rectangle.Top))/2))
      }
      else
      { 
        [int]$NewPosX = ($SecondaryScreen.Left + (($SecondaryScreen.Size.Width - ($Rectangle.Right - $Rectangle.Left))/2))
        [int]$NewPosY = ($SecondaryScreen.Top + (($SecondaryScreen.Size.Height - ($Rectangle.Bottom - $Rectangle.Top))/2))
      }
    }
    else
    {
      If ($UsePrimaryDisplay -eq $true)
      { 
        [int]$NewPosX = ($PrimaryScreen.Left + (($PrimaryScreen.Size.Width/100)*$XPosPercent))
        [int]$NewPosY = ($PrimaryScreen.Top + (($PrimaryScreen.Size.Height/100)*$YPosPercent))
      }
   
      else
      { 
        [int]$NewPosX = ($SecondaryScreen.Left + (($SecondaryScreen.Size.Width/100)*$XPosPercent))
        [int]$NewPosY = ($SecondaryScreen.Top + (($SecondaryScreen.Size.Height/100)*$XPosPercent))
      }
    }
    
      
    If (-NOT $PSBoundParameters.ContainsKey('Width')) 
    {
      [int]$Width = $Rectangle.Right - $Rectangle.Left
    }
    If (-NOT $PSBoundParameters.ContainsKey('Height')) 
    {
      [int]$Height = $Rectangle.Bottom - $Rectangle.Top
    }
    If (-NOT $PSBoundParameters.ContainsKey('X')) 
    {
      [int]$X = $NewPosX
    }
    If (-NOT $PSBoundParameters.ContainsKey('Y')) 
    {
      [int]$Y = $NewPosY
    }
    If ($Return) 
    {
      $Return = [Window]::MoveWindow($Handle, $X, $Y, $Width, $Height,$true)
    }
    If ($PSBoundParameters.ContainsKey('Passthru')) 
    {
      $Rectangle = New-Object -TypeName RECT
      $Return = [Window]::GetWindowRect($Handle,[ref]$Rectangle)
      If ($Return) 
      {
        $Height = $Rectangle.Bottom - $Rectangle.Top
        $Width = $Rectangle.Right - $Rectangle.Left
        $Size = New-Object -TypeName System.Management.Automation.Host.Size -ArgumentList $Width, $Height
        $TopLeft = New-Object -TypeName System.Management.Automation.Host.Coordinates -ArgumentList $Rectangle.Left, $Rectangle.Top
        $BottomRight = New-Object -TypeName System.Management.Automation.Host.Coordinates -ArgumentList $Rectangle.Right, $Rectangle.Bottom
        $Object = [pscustomobject]@{
          ProcessName = $ProcessName
          Handle      = $Handle
          Size        = $Size
          TopLeft     = $TopLeft
          BottomRight = $BottomRight
        }
        $Object.PSTypeNames.insert(0,'System.Automation.WindowInfo')
        $Object            
      }
    }
  }
}
function Set-WindowStyle 
{
  param(
    [Parameter()]
    [ValidateSet('FORCEMINIMIZE', 'HIDE', 'MAXIMIZE', 'MINIMIZE', 'RESTORE', 
        'SHOW', 'SHOWDEFAULT', 'SHOWMAXIMIZED', 'SHOWMINIMIZED', 
    'SHOWMINNOACTIVE', 'SHOWNA', 'SHOWNOACTIVATE', 'SHOWNORMAL')]
    $Style = 'SHOW',
    [Parameter(ValueFromPipelineByPropertyName = $true)]
    $ProcessName
    #$MainWindowHandle = (Get-Process -Id $pid).MainWindowHandle
  )
  $WindowStates = @{
    FORCEMINIMIZE   = 11
    HIDE            = 0
    MAXIMIZE        = 3
    MINIMIZE        = 6
    RESTORE         = 9
    SHOW            = 5
    SHOWDEFAULT     = 10
    SHOWMAXIMIZED   = 3
    SHOWMINIMIZED   = 2
    SHOWMINNOACTIVE = 7
    SHOWNA          = 8
    SHOWNOACTIVATE  = 4
    SHOWNORMAL      = 1
  }
  Write-Verbose -Message ('Set Window Style {1} on handle {0}' -f $MainWindowHandle, $($WindowStates[$Style]))

  $MainWindowHandles = (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue).MainWindowHandle
  ForEach ($MainWindowHandle in $MainWindowHandles)
  {
    $Win32ShowWindowAsync = Add-Type -MemberDefinition @" 
    [DllImport("user32.dll")] 
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name 'Win32ShowWindowAsync' -Namespace Win32Functions -PassThru

    $null = $Win32ShowWindowAsync::ShowWindowAsync($MainWindowHandle, $WindowStates[$Style])
  }
}
function Wait-WindowAppear
{
  Param (
    [parameter(ValueFromPipelineByPropertyName = $true)]
  $ProcessName)
  
  If (($ProcessName -eq $null) -or ($ProcessName -eq '')) 
  {
    $ProcessName = 'zzzzzzzzzzzz'
  }
  [int]$WaitTime = 0
  [bool]$WaitingFinished = $false
  If ((Get-Process -Name $ProcessName -ErrorAction SilentlyContinue).Count -gt 1) 
  {
    "More than one process with the name $ProcessName."
  }
  else
  {
    While ((Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) -and ($WaitingFinished -eq $false))
    { 
      While (([int](Get-Process -Name $ProcessName -ErrorAction SilentlyContinue).MainWindowHandle -eq 0) -and ($WaitTime -lt 300))
      {
        Start-Sleep -Milliseconds 50
        $WaitTime = $WaitTime + 1
      }
      $WaitingFinished = $true
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

  If (($ProcessToCalm -eq '') -or ($ProcessToCalm -eq $null))
  {
    'No process to wait for until it is calm.'
  }
  else
  { 
    If (($CalmThreshold -eq '') -or ($CalmThreshold -eq $null))
    {
      $CalmThreshold = 10
    } 
    $WaitTime = 100
    
    If ($ProcessToCalm.GetType().Name -eq 'Process') 
    {
      $ProcessToCalm = $ProcessToCalm.Name
    }
    
    [int]$WaitCounter = 0
    
    If (!(Get-Process -Name $ProcessToCalm -ErrorAction SilentlyContinue)) 
    { 
      "Waiting for process $ProcessToCalm to appear..."
      While ((!(Get-Process -Name $ProcessToCalm -ErrorAction SilentlyContinue)) -and ($WaitCounter -lt $WaitTime)) 
      { 
        Start-Sleep -Milliseconds 100
        $WaitCounter = $WaitCounter + 1
      }
    }

    If (Get-Process -Name $ProcessToCalm -ErrorAction SilentlyContinue) 
    { 
      ('Waiting for process {0} to calm down before continuing...' -f $ProcessToCalm)
    
      [int]$WaitCounter = 0
      Write-Host -Object ('Waiting for the process to calm down to {0}...' -f $CalmThreshold) -NoNewline
    
      While ((((((Get-Counter -Counter ('\Process({0})\% Processor Time' -f $ProcessToCalm) -ErrorAction SilentlyContinue).CounterSamples).CookedValue) -gt $CalmThreshold) -or ((((Get-Counter -Counter '\physicaldisk(_total)\% disk time' -ErrorAction SilentlyContinue).CounterSamples).CookedValue) -gt $CalmThreshold)) -AND ($WaitCounter -lt $WaitTime))
      {
        Start-Sleep -Milliseconds 250
        $WaitCounter = $WaitCounter + 1
      }
      'OK'
    }
    else
    {
      "Process $ProcessToCalm does not exist."
    }
  }
}
function Request-Close
{
  param([string]$Name)
	
  $Close_src = @'
  using System;
  using System.Runtime.InteropServices;
  public static class Win32 {
	public static uint WM_CLOSE = 0x10;

	[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
	public static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);
  }
'@

  Try 
  {
    Add-Type -TypeDefinition $Close_src -ErrorAction SilentlyContinue -InformationAction SilentlyContinue
  }
  Catch 
  {

  }
  $Zero = [IntPtr]::Zero
  $ProcessToClose = @(Get-Process $Name -ErrorAction SilentlyContinue)[0]
  Try 
  {
    [Win32]::SendMessage($ProcessToClose.MainWindowHandle, [Win32]::WM_CLOSE, $Zero, $Zero) > $null
  }
  Catch 
  {

  }
}
function Focus-Process 
{
  param(
    [Parameter(Mandatory,ValueFromPipelineByPropertyName = $true)]
    [string] $ProcessName
  )

  # As a courtesy, strip '.exe' from the name, if present.
  $ProcessName = $ProcessName -replace '\.exe$'

  # Get the PID of the first instance of a process with the given name
  # that has a non-empty window title.
  # NOTE: If multiple instances have visible windows, it is undefined
  #       which one is returned.
  $hWnd = (Get-Process -ErrorAction Ignore $ProcessName).Where({
      $_.MainWindowTitle
  }, 'First').MainWindowHandle

  if (-not $hWnd) 
  {
    "Unable to focus on $ProcessName. No process with a non-empty window title found."
  }

  $type = Add-Type -PassThru -Namespace Util -Name SetFgWin -MemberDefinition @'
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);    
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool IsIconic(IntPtr hWnd);    // Is the window minimized?
'@ 

  # Note: 
  #  * This can still fail, because the window could have bee closed since
  #    the title was obtained.
  #  * If the target window is currently minimized, it gets the *focus*, but its
  #    *not restored*.
  $null = $type::SetForegroundWindow($hWnd)
  # If the window is minimized, restore it.
  # Note: We don't call ShowWindow() *unconditionally*, because doing so would
  #       restore a currently *maximized* window instead of activating it in its current state.
  if ($type::IsIconic($hWnd)) 
  {
    $type::ShowWindow($hWnd, 9) # SW_RESTORE
  }
}
function Say-Something
{
  param
  (
    [Parameter(Position = 0)]
    $About,
    [Parameter(Position = 1)]
    $Modifier
  )
  If ($TTSFeature -eq $true) 
  {
    If (($speak -eq $null) -or ($speak -eq '')) 
    { 
      Add-Type -AssemblyName System.Speech
      $speak = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
      If (($speak.GetInstalledVoices().VoiceInfo.Name -notcontains 'Microsoft Eva') -and ($speak.GetInstalledVoices().VoiceInfo.Name -notcontains 'Microsoft Eva Mobile'))
      {
        $speak.Speak('Installing Eva voice.')
        
        $RegKeys = 'Microsoft\Speech', 'Microsoft\Speech_OneCore', 'WOW6432Node\Microsoft\SPEECH'

        ForEach ($RegKey in $RegKeys) 
        { 
          New-Item -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM" -ErrorAction SilentlyContinue
          New-Item -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM\Attributes" -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM" -Name '(Default)' -Value 'Microsoft Eva - English (United States)' -PropertyType 'String' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM" -Name '409' -Value 'Microsoft Eva - English (United States)' -PropertyType 'String' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM" -Name 'CLSID' -Value '{179F3D56-1B0B-42B2-A962-59B7EF59FE1B}' -PropertyType 'String' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM" -Name 'LangDataPath' -Value '%windir%\Speech_OneCore\Engines\TTS\en-US\MSTTSLocenUS.dat' -PropertyType 'ExpandString' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM" -Name 'LangUpdateDataDirectory' -Value '%SystemDrive%\Data\SharedData\Speech_OneCore\Engines\TTS\en-US' -PropertyType 'ExpandString' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM" -Name 'VoicePath' -Value '%windir%\Speech_OneCore\Engines\TTS\en-US\M1033Eva' -PropertyType 'ExpandString' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM" -Name 'VoiceUpdateDataDirectory' -Value '%SystemDrive%\Data\SharedData\Speech_OneCore\Engines\TTS\en-US' -PropertyType 'ExpandString' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM\Attributes" -Name 'Age' -Value 'Adult' -PropertyType 'String' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM\Attributes" -Name 'Gender' -Value 'Female' -PropertyType 'String' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM\Attributes" -Name 'Version' -Value '11.0' -PropertyType 'String' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM\Attributes" -Name 'Language' -Value '409' -PropertyType 'String' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM\Attributes" -Name 'Name' -Value 'Microsoft Eva' -PropertyType 'String' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM\Attributes" -Name 'SharedPronunciation' -Value '' -PropertyType 'String' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM\Attributes" -Name 'Vendor' -Value 'Microsoft' -PropertyType 'String' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM\Attributes" -Name 'DataVersion' -Value '11.0.2013.1022' -PropertyType 'String' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM\Attributes" -Name 'SayAsSupport' -Value 'spell=NativeSupported; cardinal=GlobalSupported; ordinal=NativeSupported; date=GlobalSupported; time=GlobalSupported; telephone=NativeSupported; currency=NativeSupported; net=NativeSupported; url=NativeSupported; address=NativeSupported; alphanumeric=NativeSupported; Name=NativeSupported; media=NativeSupported; message=NativeSupported; companyName=NativeSupported; computer=NativeSupported; math=NativeSupported; duration=NativeSupported' -PropertyType 'String' -Force -ErrorAction SilentlyContinue
          New-ItemProperty -Path "HKLM:\SOFTWARE\$RegKey\Voices\Tokens\MSTTS_V110_enUS_EvaM\Attributes" -Name 'PersonalAssistant' -Value '1' -PropertyType 'String' -Force -ErrorAction SilentlyContinue
        }
  
        #$speak.SelectVoice('Microsoft Eva')
        $speak.Speak('Eva voice installation complete. New voice will be used next time you reboot your system.')
      }
      If ($speak.GetInstalledVoices().VoiceInfo.Name -contains 'Microsoft Eva Mobile')
      {
        $speak.SelectVoice('Microsoft Eva Mobile')
      }
      If ($speak.GetInstalledVoices().VoiceInfo.Name -contains 'Microsoft Eva')
      {
        $speak.SelectVoice('Microsoft Eva')
      }
      $speak.Volume = 100
      $speak.Rate = 0
    }
  
    $SayThis = ''
  
    $SystemPronounciation = $System
    If ($System -eq 'MAME') 
    {
      $SystemPronounciation = 'Arcade'
    }
    If ($System -eq 'DOS') 
    {
      $SystemPronounciation = 'DOSS'
    }
    If ($System -eq 'C64') 
    {
      $SystemPronounciation = 'Commodore 64'
    }
    If ($System -eq 'SNES') 
    {
      $SystemPronounciation = 'Super Nintendo'
    }
    If ($System -eq 'NES') 
    {
      $SystemPronounciation = 'Nintendo Entertainment System'
    }
    If ($System -eq 'PCE') 
    {
      $SystemPronounciation = 'Turbografix'
    }
    If ($System -eq 'PSX') 
    {
      $SystemPronounciation = 'Playstation'
    }
    If ($System -eq 'PS2') 
    {
      $SystemPronounciation = 'Playstation 2'
    }

    If (($SystemPronounciation -eq '') -or ($SystemPronounciation -eq $null)) 
    {
      $SystemPronounciation = 'Computer'
    }
    If ($SystemPronounciation[0] -match '[aeiou]') 
    {
      $SystemPronoun = 'an'
    }
    else 
    {
      $SystemPronoun = 'a'
    }
  
    $SayPlease = (Get-Random -Maximum ('', 'Please'))
    $SaySystem = (Get-Random -Maximum ('', "$SystemPronounciation"))

    If ($About -eq 'Nothing') 
    {
      $SayThis = ''
    }
  
    If ($About -eq 'Greeting') 
    { 
      If (($Game -eq 'AllLauncher Menu') -or ($Game -eq '') -or ($Game -eq $null))
      {
        $NormalGreeting = Get-Random -Maximum ('Hello', 'Greetings', 'Welcome', 'Hello', 'Hey there', 'Hi there', 'Hi')
        $Greeting = $NormalGreeting
        If ([int](Get-Date -Format HH) -lt 9) 
        {
          $Greeting = Get-Random -Maximum ($NormalGreeting, 'Good Morning')
        }
        If ([int](Get-Date -Format HH) -lt 7) 
        {
          $Greeting = 'Good Morning'
        }
        If ([int](Get-Date -Format HH) -gt 20) 
        {
          $Greeting = Get-Random -Maximum ($NormalGreeting, 'Good Evening')
        }
        If ([int](Get-Date -Format HH) -gt 22) 
        {
          $Greeting = 'Good Evening'
        }
      
        $Callsign = Get-Random -Maximum ('', "$env:Username", "$MyName", "$CallMe")
        $Today = Get-Random -Maximum ('Today is', "It's", 'It is', "Today's")
        $TodaysDay = Get-Date -Format 'dddd'
        $TodaysDate = (Get-Date -Format ' d').Replace(' ','')
        $MonthAndYear = Get-Date -Format 'MMMM yyyy'
        $DatesTH = 'th'
        If (($TodaysDate[-1] -eq 3) -or ($TodaysDate[-1] -eq '3')) 
        {
          $DatesTH = 'rd'
        }
        If (($TodaysDate[-1] -eq 2) -or ($TodaysDate[-1] -eq '2')) 
        {
          $DatesTH = 'nd'
        }
        If (($TodaysDate[-1] -eq 1) -or ($TodaysDate[-1] -eq '1') -or ($TodaysDate -eq '11') -or ($TodaysDate -eq '12') -or ($TodaysDate -eq '13')) 
        {
          $DatesTH = 'st'
        }
        If (($TodaysDate -eq 11) -or ($TodaysDate -eq 12) -or ($TodaysDate -eq 13)) 
        {
          $DatesTH = 'th'
        }
        
        If ((Get-Date -Format ddMM) -eq 2412) 
        {
          $SpecialGreeting = "Merry Christmas $Callsign."
          $Today = "$Today christmas eve, "
        }
        If ((Get-Date -Format ddMM) -eq 2512) 
        {
          $SpecialGreeting = "Merry Christmas $Callsign."
          $Today = "$Today christmas day, "
        }
        If ((Get-Date -Format ddMM) -eq 3112) 
        {
          $Today = "$Today new year's eve, "
        }
        If ((Get-Date -Format ddMM) -eq 0101) 
        {
          $SpecialGreeting = "Happy new year $Callsign."
          $Today = "$Today new year's day, "
        }
            
        $SayThis = ("$Greeting $Callsign. - $Today $TodaysDay, the $TodaysDate$DatesTH of $MonthAndYear. $SpecialGreeting")
      }
      else
      {
        $speak.Rate = -1
        $speak.Speak(((Get-Random -Maximum ('Starting', 'Initiating', ' ', 'Readying', 'Loading', 'Preparing', 'Launching', 'Commencing', 'Opening'))))
        $speak.Rate = -2
        $speak.Speak("$GameName")
        $speak.Rate = 0
        $SayThis = ''
      }
    }

    If ($About -eq 'ReturnToLauncher') 
    {
      $ReturningTo = (Get-Random -Maximum ('Returning to', 'Going back to', 'Reloading', 'Back to', 'Reverting to', 'And now back to', 'reopening', 'restarting'))
      $IsRestarting = (Get-Random -Maximum ('is restarting', 'is reloading', 'is returning', 'is coming back'))
      $SayThis = (Get-Random -Maximum ("$ReturningTo $LauncherName.", "$LauncherName $IsRestarting.", "Displaying $LauncherName again.", "Showing $LauncherName again."))
    }

    If ($About -eq 'HOTAS') 
    {
      $HOTAS = (Get-Random -Maximum ('HOTAS control', 'flightstick control', 'use of HOTAS controller', 'use of flightstick controllers'))
      $SayThis = "This game is configured for $HOTAS."
    }
    
    If ($About -eq 'StartLauncher') 
    {
      $ReturningTo = (Get-Random -Maximum ('Starting', 'Opening', 'Loading', "Let's go to", 'Displaying', 'Initializing', 'Showing', 'Starting', 'Loading'))
      $IsRestarting = (Get-Random -Maximum ('is starting', 'is loading', 'is opening', 'is initializing'))
      $SayThis = (Get-Random -Maximum ("$ReturningTo $LauncherName.", "$LauncherName $IsRestarting."))
    }
  
    If ($About -eq 'ReadyVR') 
    {
      $MakingSure = (Get-Random -Maximum ('Making sure', 'Making certain', 'Checking that'))
      $VRis = (Get-Random -Maximum ('is ready', 'is working correctly', 'is prepared', 'is initialized', 'drivers are loaded.'))
      $SayThis = (Get-Random -Maximum ("$MakingSure VR hardware $VRis.", 'Checking VR hardware readyness...', 'Checking status of VR hardware....'))
    }
  
    If ($About -eq 'StartOculusClient') 
    {
      $SayThis = (Get-Random -Maximum ('Starting Oculus Client...', 'Starting Client for Oculus Rift...', 'Starting Oculus Rift Client...', 'Starting Oculus Client...'))
    }
  
    If ($About -eq 'StartOculusServices') 
    {
      $SayThis = (Get-Random -Maximum ('Starting Oculus services...', 'Starting Oculus Rift services...', 'Starting services for Oculus Rift...', 'Oculus Rift services are being started...', 'Oculus services are being started...'))
    }
  
    If ($About -eq 'RestartOculus') 
    {
      $SayThis = (Get-Random -Maximum ('Restarting both Oculus services and Oculus Client...', 'Trying to restart Oculus services and Oculus Client...', 'Restarting Oculus Rift services and Client...'))
    }

    If ($About -eq 'StartSteamVR') 
    {  
      $SVRneeded = (Get-Random -Maximum ('needed', 'required', 'necessary', 'demanded', 'essential'))
      $SVRneeds = (Get-Random -Maximum ('needs', 'requires', 'demands', 'uses'))
      $SVRstart = (Get-Random -Maximum ('starting', 'loading'))
      $SayThis = (Get-Random -Maximum ("$SVRstart Steam VR...", "Game $SVRneeds Steam VR. $SVRstart it...", "Steam VR $SVRneeded. $SVRstart it...", "$GameName $SVRneeds Steam VR. $SVRstart it..."))
    }
  
    If ($About -eq 'SteamVRLoaded') 
    {  
      $SVRloaded = (Get-Random -Maximum ('loaded', 'ready', 'prepared'))
      $SVRfinally = (Get-Random -Maximum ('', 'now', 'finally'))
      $SVRstartG = (Get-Random -Maximum ('starting', 'loading', 'entering'))
      $SayThis = "Steam VR $SVRloaded. $SVRfinally $SVRstartG game."
    }

    If ($About -eq 'JoyMapper') 
    {
      $SayThis = (Get-Random -Maximum ("A controller mapping is used for this $Modifier.", "Starting JoyMapper associated with this $Modifier.", "For this $Modifier, keyboard commands are mapped to your controller.", "Controller Mapping profile for this $Modifier found."))
    }
      
    If ($About -eq 'OpenGameDocs') 
    {
      $GDOpening = (Get-Random -Maximum ('Opening', 'Displaying', 'Readying', 'Loading', 'Preparing'))
      $SayThis = (Get-Random -Maximum ("$GDOpening game documents.", "$GDOpening multiple documents.", "There are several game documents. $GDOpening them.", "Several game documents found. $GDOpening them.", "Multiple game documents found. $GDOpening them."))
    }
    
    If ($About -eq 'OpenOneGameDoc') 
    {
      $GDOpening = (Get-Random -Maximum ('Opening', 'Displaying', 'Readying', 'Loading', 'Preparing'))
      $SayThis = (Get-Random -Maximum ("$GDOpening game document.", "$GDOpening a document.", "There is a game document. $GDOpening it.", "One game document found. $GDOpening it.", "A single game document has been found. $GDOpening it."))
    }
  
    If ($About -eq 'OpenGameWebpage') 
    {
      $SayThis = (Get-Random -Maximum ("Opening game web page$Modifier...", "Displaying game web page$Modifier...", "Readying game web page$Modifier...", "Opening game URL$Modifier..."))
    }
    
    If ($About -eq 'OpenGameTool') 
    {
      $SayThis = (Get-Random -Maximum ("Opening game related link$Modifier...", "Opening game related shortcut$Modifier...", "Readying additional game tool$Modifier..."))
    }
  
    If ($About -eq 'OpenUHS') 
    {
      $SayThis = (Get-Random -Maximum ('Opening universal hint system file...', 'U H S file found. Opening it...', 'Displaying Universal Hint System file...', 'Opening U H S file...'))
    }
  
    If ($About -eq 'LoadTrainer') 
    {
      $TrainerLoading = (Get-Random -Maximum ('Loading', 'Activating', 'Starting', 'Initializing'))
      $SayThis = (Get-Random -Maximum ("$TrainerLoading trainer.", "Trainer is $TrainerLoading.", "Trainer found. $TrainerLoading it."))
    }
  
    If ($About -eq 'LoadCheatTable') 
    {
      $TableLoading = (Get-Random -Maximum ('Loading', 'Opening'))
      $SayThis = (Get-Random -Maximum ("$TableLoading $Modifier table.", "$TableLoading table for $Modifier", "$Modifier table found. $TableLoading it."))
    }

    If ($About -eq 'MountDisc') 
    {
      $SayThis = (Get-Random -Maximum ("Mounting disc-Image$Modifier...", "Mounting ISO-Image$Modifier...", "Disc-Image$Modifier found. Mounting...", "ISO-Image$Modifier found. Mounting...", "Mounting CD or DVD image$Modifier...", "CD or DVD Image$Modifier found. Mounting..."))
    }
  
    If ($About -eq 'TimePlayed') 
    {
      $TPtime = [timespan]$ThisPlayed
      $TPhave = (Get-Random -Maximum ('', 'have'))
      $TPgame = (Get-Random -Maximum ('', 'the game'))
      $SayThis = (Get-Random -Maximum ("You $TPhave played $TPgame for $TPtime.", "You were playing $TPgame for $TPtime", "The game ran for $TPtime."))
    }

    If ($About -eq 'TotalTimePlayed') 
    {
      $TTPintotal = (Get-Random -Maximum ('', 'In total,'))
      $TTPplayed = (Get-Random -Maximum ('played', 'have played', 'have been playing'))
      $TTPgame = (Get-Random -Maximum ('the game', $GameName, 'this game'))
      $SayThis = (Get-Random -Maximum ("$TTPintotal you $TTPplayed $TTPgame for more than $TotalHours hours.", "You now spent more than $TotalHours hours with $TTPgame.", "$TTPintotal you $TTPplayed $TTPgame for more than $TotalHours hours."))
    }


    If ($About -eq 'StartGame') 
    {
      $SGStarting = (Get-Random -Maximum ('Starting', 'Loading', 'Initiating', 'Activating', 'Ready for', 'Playing'))
      $SGReady = (Get-Random -Maximum ('ready', 'initialized', 'prepared'))
      $StartingSpeech = (Get-Random -Maximum ("$SGStarting $SaySystem game...", "$SGStarting $GameName...", "System $SGReady."))
      $SayThis = $StartingSpeech
      If (($System -ne 'Windows') -and ($System -ne 'VR') -and ($System -ne 'Computer')) 
      {
        $StartingEmu = "$SGStarting $SaySystem emulator..."
        $SayThis = (Get-Random -Maximum ($StartingSpeech, $StartingEmu, $StartingSpeech))
      }
      If ($EmulatorINIEntry -like '*.dll') 
      {
        $StartingRA = "$SGStarting RetroArch..."
        $SayThis = (Get-Random -Maximum ($StartingSpeech, $StartingEmu, $StartingSpeech, $StartingRA))
      }
      If ($EmulatorINIEntry -eq 'mednafen') 
      {
        $StartingME = "$SGStarting Mednafen..."
        $SayThis = (Get-Random -Maximum ($StartingSpeech, $StartingEmu, $StartingSpeech, $StartingME))      
      }
    }
    
    If ($About -eq 'Problem') 
    { 
      $AProblem = (Get-Random -Maximum ('a problem', 'an issue', 'an error', 'a complication'))
      $SayThis = (Get-Random -Maximum ("There has been $AProblem.", "There was $AProblem.", "There is $AProblem.", "$AProblem has occured."))
    }
  
    If ($About -eq 'TakesLonger') 
    {
      $SayThis = (Get-Random -Maximum ('This takes longer than expected.', 'This takes longer than anticipated.'))
    }
  
    If ($About -eq 'Loading') 
    {
      $SayThis = (Get-Random -Maximum ('Loading... -', 'Working... -', 'Getting things ready... -', 'Getting ready... -', 'Preparing... -', 'Initializing... -', 'Working on it... -'))
    }

    If ($About -eq 'PleaseWait') 
    {
      $SayThis = (Get-Random -Maximum ("$SayPlease wait...", "$SayPlease Hold on...", "Just a second $SayPlease...", "$SayPlease Hang in there...", "$SayPlease Stand by...", "$SayPlease Hold tight...", "Patience $SayPlease...", "One moment $SayPlease...", "Just a minute $SayPlease...", "Bear with me $SayPlease..."))
    }
  
    If ($About -eq 'OpeningSpeech') 
    {
      $OSSaySystem = (Get-Random -Maximum ('', 'system'))
      $OSPreparingFor = (Get-Random -Maximum ('This is', "Preparing $OSSaySystem for", "Getting $OSSaySystem ready for", "Readying $OSSaySystem for"))
      $NormalOpening = "$OSPreparingFor $SystemPronoun $SystemPronounciation game."
      $SayThis = $NormalOpening
      $MednafenOpening = (Get-Random -Maximum ("Using Mednafen to emulate $SystemPronoun $SystemPronounciation.", "Using Mednafen to emulate $SystemPronoun $SystemPronounciation system."))
      If (($System -ne 'Windows') -and ($System -ne 'VR') -and ($System -ne 'Computer')) 
      {
        $EmulatorOpening = (Get-Random -Maximum ("Emulating $SystemPronoun $SystemPronounciation $OSSaySystem.", "Preparing to emulate $SystemPronoun $SystemPronounciation $OSSaySystem.", "Preparing Emulator for $SystemPronoun $SystemPronounciation $OSSaySystem."))
        $SayThis = (Get-Random -Maximum ($NormalOpening, $NormalOpening, $EmulatorOpening))
      }
      If ($EmulatorINIEntry -like '*.dll') 
      {
        $RetroArchOpening = (Get-Random -Maximum ("Using RetroArch to emulate $SystemPronoun $SystemPronounciation $OSSaySystem.", "Preparing RetroArch to emulate $SystemPronoun $SystemPronounciation $OSSaySystem.", "Emulating $SystemPronoun $SystemPronounciation $OSSaySystem with RetroArch.", "Preparing to Emulate $SystemPronoun $SystemPronounciation $OSSaySystem with RetroArch."))
        $SayThis = (Get-Random -Maximum ($NormalOpening, $NormalOpening, $EmulatorOpening, $RetroArchOpening))
      }
      If ($EmulatorINIEntry -eq 'mednafen') 
      {
        $MednafenOpening = (Get-Random -Maximum ("Using Mednafen to emulate $SystemPronoun $SystemPronounciation $OSSaySystem.", "Preparing Mednafen to emulate $SystemPronoun $SystemPronounciation $OSSaySystem.", "Emulating $SystemPronoun $SystemPronounciation $OSSaySystem with Mednafen.", "Preparing to Emulate $SystemPronoun $SystemPronounciation $OSSaySystem with Mednafen."))
        $SayThis = (Get-Random -Maximum ($NormalOpening, $NormalOpening, $EmulatorOpening, $MednafenOpening))
      }
    }
  
    If ($About -eq 'GameEnded') 
    {
      $GEsaySystem = (Get-Random -Maximum ('', "$SystemPronounciation"))
      $GEgame = 'game'
      $GEhas = (Get-Random -Maximum ('', 'has', 'has been'))
      $GEended = (Get-Random -Maximum ('ended', 'finished', 'closed', 'terminated', 'unloaded'))
      If (($System -ne 'Windows') -and ($System -ne 'VR') -and ($System -ne 'Computer')) 
      {
        $GEgame = (Get-Random -Maximum ('game', 'emulator'))
      }
      If ($EmulatorINIEntry -like '*.dll')  
      {
        $GEgame = (Get-Random -Maximum ('game', 'emulator', 'RetroArch'))
      }
      If ($EmulatorINIEntry -eq 'mednafen') 
      {
        $GEgame = (Get-Random -Maximum ('game', 'emulator', 'Mednafen'))
      }
      $SayThis = (Get-Random -Maximum ("$GEsaySystem $GEgame $GEhas $GEended.", "$GameName $GEhas $GEended.", 'Game, Over.', "$GEsaySystem $GEgame $GEhas $GEended.", "$GameName $GEhas $GEended.", 'Game Over', "$GEsaySystem $GEgame process $GEhas $GEended." ))
    }
  
    If ($About -eq 'Finished') 
    {
      $SayThis = (Get-Random -Maximum ('Finished.', 'Everything is finished.', 'Ending.', 'Done.'))
    }
    If ($About -eq 'Ready') 
    {
      $SayThis = (Get-Random -Maximum ('Ready.', 'Ready!', 'Okay, ready.', 'We are ready.', "We're ready.", 'Everything is ready.', "Everything's ready.", "$LauncherName is ready.", "$LauncherName ready.", 'Ready for you.', 'Ready for you, Sir.'))
    }
      
    $speak.Speak($SayThis)
  }
}


###################### These are AllLauncher specific functions ################################

function Start-Launcher 
{
  ('Starting {0} and finishing up here...' -f $LauncherName) | timestamp >> $LogFile
  'Beginning with Start-Launcher...'
  [bool]$Menu = $true
  Stop-Process -Name VirtualMIDISynch -Force -ErrorAction SilentlyContinue
  
  . Close-UnneededStuff
  While ($EverythingClosed -ne $true) 
  {
    Start-Sleep -Seconds 1
  }
  'Everything closed.'
    
  'Resetting Explorer...'
  Stop-Process -Name Explorer -ErrorAction SilentlyContinue -Force
  
  If ($UsePS4Pad -eq $true) 
  {
    . Set-DS4
    While ($DS4Finished -ne $true) 
    {
      Start-Sleep -Seconds 1
    }
    'DS4 finished.'
  }


  If ($WasLauncherRunning -eq $true) 
  {
    Say-Something -About ReturnToLauncher
  }
  else
  {
    Say-Something -About StartLauncher
  }

  
  If (!(Get-Process -Name explorer)) 
  {
    Start-Process -FilePath $Explorer.Path -ErrorAction SilentlyContinue
  }
    
  $CursorRefresh::SystemParametersInfo(0x0057,0,$null,0)
  Start-WhitelistedServices
  
  'Starting Launcher'
  'Starting Launcher...' | timestamp >> $LogFile
  Start-Process -FilePath $Launcher -WorkingDirectory $LauncherDir -Verb RunAs -WindowStyle Maximized
  Wait-WindowAppear -ProcessName $LauncherProcess
  Focus-Process -ProcessName $LauncherProcess -ErrorAction SilentlyContinue
  Move-Window -ProcessName $LauncherProcess -X $PrimaryScreen.Top -Y $PrimaryScreen.Left -Width $PrimaryScreen.Right -Height $PrimaryScreen.Bottom -Screen 1
  Set-WindowStyle -ProcessName $LauncherProcess -Style MAXIMIZE -ErrorAction SilentlyContinue
    
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
  
  Wait-ProcessToCalm -ProcessToCalm $LauncherProcess -CalmThreshold 60
  Move-Window -ProcessName $LauncherProcess -X $PrimaryScreen.Top -Y $PrimaryScreen.Left -Width $PrimaryScreen.Right -Height $PrimaryScreen.Bottom -Screen 1
  Set-WindowStyle -ProcessName $LauncherProcess -Style MAXIMIZE -ErrorAction SilentlyContinue
  #Focus-Process -ProcessName $LauncherProcess -ErrorAction SilentlyContinue
    
  Say-Something -About Ready
  
  'Finished!' | timestamp >> $LogFile
  'Finished!'
  
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
  Stop-Process -Name AllLauncher -Force -ErrorAction SilentlyContinue
  Stop-Process -Name Powershell -Force -ErrorAction SilentlyContinue
  If ((Get-Process -Id $pid -ErrorAction SilentlyContinue).Name -ne 'powershell_ise') 
  { 
    Stop-Process -Id $pid -Force -ErrorAction SilentlyContinue
    exit
  }
  Exit
}
function Quit-Launcher 
{
  'Quitting Launcher...'
  IF (Get-Process -Name $LauncherProcess -ErrorAction SilentlyContinue)
  {
    ('Closing {0}...' -f $LauncherName) | timestamp >> $LogFile
    Request-Close -Name $LauncherProcess
    (Get-Process -Name $LauncherProcess -ErrorAction SilentlyContinue).CloseMainWindow()
    Wait-Process -Name $LauncherProcess -ErrorAction SilentlyContinue -Timeout 2
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
    Invoke-Item -Path $CurrentGameDocsDir -ErrorAction SilentlyContinue
    $GamePDFs = Get-ChildItem -Path $CurrentGameDocsDir -Filter '*.pdf' | Select-Object -ExpandProperty FullName

    If ($GamePDFs.Count -gt 1) 
    {
      Say-Something -About OpenGameDocs
    }
    If ($GamePDFs.Count -eq 1) 
    {
      Say-Something -About OpenOneGameDoc
    }

    ForEach($GamePDF in $GamePDFs) 
    {
      ('Opening {0}.' -f $GamePDF) | timestamp >> $LogFile
      Invoke-Item -Path $GamePDF -ErrorAction SilentlyContinue
    }
    $GameURLs = Get-ChildItem -Path $CurrentGameDocsDir -Filter '*.url' | Select-Object -ExpandProperty FullName

    If ($GameURLs.Count -gt 1) 
    {
      Say-Something -About OpenGameWebpage -Modifier s
    }
    If ($GameURLs.Count -eq 1) 
    {
      Say-Something -About OpenGameWebpage
    }

    ForEach($GameURL in $GameURLs) 
    {
      ('Opening {0}.' -f $GameURL) | timestamp >> $LogFile
      Invoke-Item -Path $GameURL -ErrorAction SilentlyContinue
      #Start-Sleep -Milliseconds 750
    }
    $GameToolShortcuts = Get-ChildItem -Path $CurrentGameDocsDir -Filter '*.lnk' | Select-Object -ExpandProperty FullName

    If ($GameToolShortcuts.Count -gt 1) 
    {
      Say-Something -About OpenGameTool -Modifier s
    }
    If ($GameToolShortcuts.Count -eq 1) 
    {
      Say-Something -About OpenGameTool
    }

    ForEach($GameToolShortcut in $GameToolShortcuts) 
    {
      ('Opening {0}.' -f $GameToolShortcut) | timestamp >> $LogFile
      Invoke-Item -Path $GameToolShortcut -ErrorAction SilentlyContinue
      #Start-Sleep -Seconds 3
    }
    $GameUHSs = Get-ChildItem -Path $CurrentGameDocsDir -Filter '*.uhs' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
    ForEach($GameUHS in $GameUHSs) 
    {
      ('Opening {0}.' -f $GameUHS) | timestamp >> $LogFile

      Say-Something -About OpenUHS

      Invoke-Item -Path $GameUHS -ErrorAction SilentlyContinue
      #Start-Sleep -Milliseconds 500
    }
  }
  else 
  {
    New-Item -Path $CurrentGameDocsDir -ItemType Directory
  }
  
  If (Test-Path -Path $UnsortedGamePDF) 
  {
    ('Opening {0}.' -f $UnsortedGamePDF) | timestamp >> $LogFile

    Say-Something -About OpenOneGameDoc

    Invoke-Item -Path $UnsortedGamePDF -ErrorAction SilentlyContinue
    #Start-Sleep -Milliseconds 500
  }
  If (Get-Process -Name '*pdf*', '*reader*', '*foxit*', '*acro*', '*tmp', '*uhs*', '*spot*', '*fire*' -ErrorAction SilentlyContinue)
  {
    #Start-Sleep -Milliseconds 250
    $CalmingTMPProcesses = (Get-Process -Name '*pdf*', '*reader*', '*foxit*', '*acro*', '*tmp', '*spot*', '*fire*' -ErrorAction SilentlyContinue).Name | Get-Unique
    ForEach ($CalmingTMPProcess in $CalmingTMPProcesses) 
    {
      Wait-ProcessToCalm -ProcessToCalm $CalmingTMPProcess.Name -CalmThreshold 12
    }
  }
}
function Start-Cheats 
{
  'Looking for cheats...'
  'Looking for cheats...' | timestamp >> $LogFile
  $GameTrainer = $TrainersDir + '\' + $GameName + '.exe'
  $Additionaltrainer = $TrainersDir + '\' + $GameName + '_add.exe'
  If (Test-Path -Path $GameTrainer) 
  {
    If (!(Test-Path -Path ($TrainersDir + '\' + $GameName + ' Trainer.exe'))) 
    {
      Rename-Item -Path $GameTrainer -NewName "$GameName Trainer.exe"
    }
    else 
    {
      Remove-Item -Path $GameTrainer -Force
    }
  }
  If (Test-Path -Path $Additionaltrainer) 
  {
    If (!(Test-Path -Path ($TrainersDir + '\' + $GameName + ' Trainer 2.exe'))) 
    {
      Rename-Item -Path $Additionaltrainer -NewName "$GameName Trainer 2.exe"
    }
    else 
    {
      Remove-Item -Path $Additionaltrainer -Force
    }
  }
    
  'Looking for trainers...'
  $GameTrainers = Get-ChildItem -Path $TrainersDir -Filter '*.exe' | Where-Object -FilterScript {
    $_.Name -like "$GameName Trainer*"
  }
  
  If (($GameTrainers -ne '') -and ($GameTrainers -ne $null)) 
  {
    ForEach ($GameTrainer in $GameTrainers) 
    {
      Say-Something -About LoadTrainer

      ('Opening {0}.' -f $GameTrainer) | timestamp >> $LogFile
      ('Opening {0}.' -f $GameTrainer)
      Start-Process -FilePath $GameTrainer -WorkingDirectory $TrainersDir -Verb runas
      $GameTrainerProcess = Get-Process -Name (((($GameTrainer.Name).split('\'))[-1]).replace('.exe','')) -ErrorAction SilentlyContinue
      Wait-ProcessToCalm -ProcessToCalm $GameTrainerProcess.Name -CalmThreshold 15
      If (Get-Process -Name '*.tmp'-ErrorAction SilentlyContinue) 
      {
        $CalmingTMPProcesses = (Get-Process -Name '*.tmp'-ErrorAction SilentlyContinue).Name
        ForEach ($CalmingTMPProcess in $CalmingTMPProcesses)
        {
          Wait-ProcessToCalm -ProcessToCalm $CalmingTMPProcess -CalmThreshold 20
        }
      }
    }
  }

  $ArtMoneyTable = $ArtMoneySystemDir + '\' + $GameName + '.' + $ArtMoneyExt
  $GameFileCheatTable = $ArtMoneyTablesDir + '\Files\' + $GameName + '.' + $ArtMoneyExt
  If (Test-Path -Path $GameFileCheatTable) 
  {
    Say-Something -About LoadCheatTable -Modifier 'ArtMoney FileEditor'
    
    ('Loading {0} with {1}.' -f $GameFileCheatTable, $ArtMoneyName) | timestamp >> $LogFile
    $ArtMoneyArgument = '"' + $GameFileCheatTable + '"'
    Invoke-Item -Path $GameFileCheatTable -ErrorAction SilentlyContinue
    Wait-ProcessToCalm -ProcessToCalm $ArtmoneyProcess -CalmThreshold 15
  }  
  If (Test-Path -Path $ArtMoneyTable) 
  {
    Say-Something -About LoadCheatTable -Modifier ArtMoney
    
    ('Loading {0} with {1}.' -f $ArtMoneyTable, $ArtMoneyName) | timestamp >> $LogFile
    $ArtMoney = $ArtMoneyDir + '\' + $ArtMoneyExe
    $ArtMoneyTableDir = $ArtMoneyTablesDir + '\' + $System
    $ArtMoneyArgument = '"' + $ArtMoneyTable + '"'
    Invoke-Item -Path $ArtMoneyTable -ErrorAction SilentlyContinue
    Wait-ProcessToCalm -ProcessToCalm $ArtmoneyProcess -CalmThreshold 15
    Get-Process -Name $ArtmoneyProcess -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty Id |
    Activate-App -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 250
    Send-Keys -KeysToSend 'n'
  }
  $CheatEngineTable = $CheatEngineTablesDir + '\' + $GameName + '.' + $CheatEngineExt
  If (Test-Path -Path $CheatEngineTable) 
  {
    Say-Something -About LoadCheatTable -Modifier CheatEngine
    
    ('Loading {0} with {1}.' -f $CheatEngineTable, $CheatEngineName) | timestamp >> $LogFile
    Invoke-Item -Path $CheatEngineTable -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    Wait-ProcessToCalm -ProcessToCalm $CheatEngineProcess -CalmThreshold 15
  }
  $CoSMOSTable = $CoSMOSTablesDir + '\' + $GameName + '.' + $CoSMOSExt
  If (Test-Path -Path $CoSMOSTable) 
  {
    ('Loading {0} with {1}.' -f $CoSMOSTable, $CoSMOSName) | timestamp >> $LogFile
    Invoke-Item -Path $CoSMOSTable -ErrorAction SilentlyContinue
    Wait-ProcessToCalm -ProcessToCalm $CoSMOSProcess
  }
  If (Test-Path -Path $UHSFile -ErrorAction SilentlyContinue) 
  {
    ('Opening {0}.' -f $UHSFile) | timestamp >> $LogFile

    Say-Something -About OpenUHS
    
    Invoke-Item -Path $UHSFile -ErrorAction SilentlyContinue
    #Start-Sleep -Milliseconds 500
  }
  'Cheat-searching completed.' | timestamp >> $LogFile
  $CheatsReady = $true
}
function Set-HOTAS
{
  $HOTASGames = Get-Content -Path $HOTASGamesFile -ErrorAction SilentlyContinue
  
  If (($HOTASGames -contains $GameName) -and (Get-PnpDevice -FriendlyName "*$HOTAS*"))
  { 
    Say-Something -About HOTAS
    $UseDS4 = $false
    
    Get-PnpDevice -FriendlyName '*xbox*' -ErrorAction SilentlyContinue | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
    Stop-Process -Name 'DS4Windows' -Force -ErrorAction SilentlyContinue
    Get-PnpDevice -FriendlyName '*game*controller*' -ErrorAction SilentlyContinue | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
    Get-PnpDevice -FriendlyName '*game*controller*' -ErrorAction SilentlyContinue | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
    Get-PnpDevice -Class HIDClass -ErrorAction SilentlyContinue |
    Where-Object -FilterScript {
      $_.HardwareID.Contains('HID_DEVICE_SYSTEM_GAME')
    } |
    Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
    Get-PnpDevice -FriendlyName "*$HOTAS*" -ErrorAction SilentlyContinue | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
      
    Get-PnpDevice -FriendlyName "*$HOTAS*" -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
    Get-PnpDevice -FriendlyName "*$HOTAS*" -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
    Get-PnpDevice -FriendlyName '*hid*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
  }
}
function Set-DS4 
{
  $DS4Blacklist = Get-Content -Path $DS4BlacklistFile -ErrorAction SilentlyContinue
  Get-PnpDevice -FriendlyName '*xbox*' -ErrorAction SilentlyContinue | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
  Stop-Process -Name 'DS4Windows' -Force -ErrorAction SilentlyContinue
  Get-PnpDevice -FriendlyName '*game*controller*' -ErrorAction SilentlyContinue | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
  Get-PnpDevice -FriendlyName '*game*controller*' -ErrorAction SilentlyContinue | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
  Get-PnpDevice -Class HIDClass -ErrorAction SilentlyContinue |
  Where-Object -FilterScript {
    $_.HardwareID.Contains('HID_DEVICE_SYSTEM_GAME')
  } |
  Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
  
  Get-PnpDevice -FriendlyName '*game*controller*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
  Get-PnpDevice -FriendlyName '*game*controller*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
  Get-PnpDevice -FriendlyName '*hid*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
  
    
  $UseDS4 = $true 
  
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
  
  IF ($Menu -eq $true) 
  {
    $UseDS4 = $true
  }
  
  IF ($UseDS4 -eq $true) 
  {
    'Making sure DS4Windows is running'
    Start-Process -FilePath "$DS4Folder\DS4Windows.exe" -WorkingDirectory $DS4Folder -Verb RunAs
    Wait-ProcessToCalm -ProcessToCalm DS4Windows -CalmThreshold 11
    Get-PnpDevice -FriendlyName '*xbox*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
  }
  $DS4Finished = $true
}
function Set-JoyMapper
{
  'Checking for profile file...'
  $JoyMapperSystemProfile = $JoyMapperDir + '\' + $System + '.' + $JoyMapperExt
  $JoyMapperGameProfile = $JoyMapperDir + '\' + $GameName + '.' + $JoyMapperExt
  Stop-Process -Name "*$JoyMapperProcess*" -Force -ErrorAction SilentlyContinue
  If (Test-Path $JoyMapperSystemProfile -ErrorAction SilentlyContinue) 
  {
    "$JoyMapperSystemProfile found!"
    Say-Something -About JoyMapper -Modifier System
    Invoke-Item $JoyMapperSystemProfile -ErrorAction SilentlyContinue
  }
  If (Test-Path $JoyMapperGameProfile -ErrorAction SilentlyContinue) 
  {
    "$JoyMapperGameProfile found!"
    Stop-Process -Name "*$JoyMapperProcess*" -Force -ErrorAction SilentlyContinue
    Say-Something -About JoyMapper -Modifier Game
    Invoke-Item $JoyMapperGameProfile -ErrorAction SilentlyContinue
  }
}
function Manage-DRMSystem
{
  $SteamPath = Get-Item -Path (((Get-Item -Path HKCU:\Software\Valve\Steam).GetValue('SteamEXE')).ToString()).Replace('steam.exe','')
  $SteamEXE = Get-Item -Path ((Get-Item -Path HKCU:\Software\Valve\Steam).GetValue('SteamEXE')).ToString()
  $SteamEXEs = (Get-ChildItem -Path $SteamPath -Filter '*.exe').FullName
  $OriginPath = Get-Item -Path (((Get-Item -Path HKLM:\SOFTWARE\WOW6432Node\Origin).GetValue('ClientPath')).ToString()).Replace('Origin.exe','')
  $OriginEXE = Get-Item -Path (((Get-Item -Path HKLM:\SOFTWARE\WOW6432Node\Origin).GetValue('ClientPath')).ToString())
  $OriginEXEs = (Get-ChildItem -Path $OriginPath -Filter '*.exe').FullName
  
  If ($System -eq 'Steam') 
  { 
    $DRMSystem = 'Steam'
    $System = 'Windows'
  }

  If ($System -eq 'Uplay') 
  { 
    $DRMSystem = 'Uplay'
    $System = 'Windows'
  }

  If ($System -eq 'Origin') 
  { 
    $DRMSystem = 'Origin'
    $System = 'Windows'
  }

  If (($System -eq 'GOG') -or ($System -like '*galaxy*')) 
  { 
    $DRMSystem = 'GalaxyClient'
    $System = 'Windows'
  }

  If ($System -eq 'Windows')
  {
    IF (($GameExt -eq 'htm') -or ($GameExt -eq 'html') -or ($GameExt -eq 'url')) 
    {
      If ((Get-Content -Path $Game -ErrorAction SilentlyContinue) -like '*uplay:*') 
      {
        'This is a UPlay (upc) game.'
        'This is a UPlay (upc) game.' | timestamp >> $LogFile
        $DRMSystem = 'Uplay'
      }
      If ((Get-Content -Path $Game -ErrorAction SilentlyContinue) -like '*steam:*') 
      {
        'This is a Steam game.'
        'This is a Steam game.' | timestamp >> $LogFile
        $DRMSystem = 'Steam'
      }
    }
    IF ($GameExt -eq 'lnk') 
    {
      'It is a shortcut. Let us see, where this goes...'
      $GameEXE = (Get-Shortcut -Path $Game -ErrorAction SilentlyContinue)
      $GamePath = $GameEXE.TargetPath -replace $GameEXE.Target
      $GameProcess = ($GameEXE.Target).Replace('.exe','')
      ('It is {0} at {1},' -f $GameEXE.Target, $GamePath) | timestamp >> $LogFile
      ('It is {0} at {1},' -f $GameEXE.Target, $GamePath) 
      If (($GameEXE.Arguments -ne '') -and ($GameEXE.Arguments -ne $null)) 
      {
        ('and there is an argument string: {0}' -f $GameEXE.Arguments)
      }
      else 
      {
        'there are no arguments.'
        $GameEXE.Arguments = ' '
      }
    
      If ($GameEXE.TargetPath -like '*GalaxyClient.exe*') 
      {
        'This is a Galaxy (GOG) game.'
        'This is a Galaxy (GOG) game.' | timestamp >> $LogFile
        $DRMSystem = 'Galaxy'
      }
      If ($GameEXE.TargetPath -like '*\Origin\*') 
      {
        'This is a Origin game.'
        'This is a Origin game.' | timestamp >> $LogFile
        $DRMSystem = 'Origin'
      }
    } 
  }

  IF ($DRMSystem -ne 'Steam')
  {
    If (Get-Process -Name 'Steam' -ErrorAction SilentlyContinue) 
    {
      Start-Process -FilePath $SteamEXE -ArgumentList '-shutdown' -Wait
      'Closing Steam'
      Stop-Process -Name Steam -Force -ErrorAction SilentlyContinue
    }
  }  
  
  IF ($DRMSystem -ne 'Galaxy')
  {
    If (Get-Process -Name 'GalaxyClient' -ErrorAction SilentlyContinue) 
    {
      'Closing GOG Galaxy Client'
      Stop-Process -Name GalaxyClient -Force -ErrorAction SilentlyContinue
    }
  }  
  
  IF ($DRMSystem -ne 'UPlay')
  {
    If (Get-Process -Name 'upc' -ErrorAction SilentlyContinue) 
    {
      'Closing UPlay (UPC)'
      Stop-Process -Name upc -Force -ErrorAction SilentlyContinue
    }
  }  
  
  IF ($DRMSystem -ne 'Origin')
  {
    If (Get-Process -Name 'origin' -ErrorAction SilentlyContinue) 
    {
      Start-Process -FilePath $OriginEXE -ArgumentList 'origin://Quit' -Wait
      'Closing Origin'
      Stop-Process -Name origin -Force -ErrorAction SilentlyContinue
    }
  }
}
function Close-UnneededStuff 
{
  'Closing unneeded stuff...' | timestamp >> $LogFile
  'Close-UnneededStuff:'
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

  Get-Process -Name '*.tmp' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  Get-Process -Name '*trainer*' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  Get-Process -Name '*midi*' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  Get-Process -Name '*MIDISynth*' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  Get-Process -Name "*$JoyMapperProcess*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  $CloseProcesses = Get-Process -Name '*launch*', '*start*', '*conf*', '*setup*', '*inst*', '*store*', '*upd*' -ErrorAction SilentlyContinue |
  Where-Object -Property ProcessName -NE -Value $LauncherProcess |
  Where-Object -Property Id -NE -Value $pid
  ForEach ($CloseProcess in $CloseProcesses) 
  {
    $null = $CloseProcess.CloseMainWindow()
  }
    
  ForEach ($DeviceNumber in 0..9) 
  {
    Get-DiskImage -DevicePath "\\.\CDROM$DeviceNumber" -ErrorAction SilentlyContinue | Dismount-DiskImage
  }
    
  'Stopping unneeded services...' 
  'Stopping unneeded services...' | timestamp >> $LogFile
  $ProcessBlacklist = Get-Content -Path $ProcessBlacklistFile -ErrorAction SilentlyContinue
  $ServiceBlacklist = Get-Content -Path $ServiceBlacklistFile -ErrorAction SilentlyContinue
  ForEach ($BlacklistService in $ServiceBlacklist) 
  {
    If ((Get-Service $BlacklistService -ErrorAction SilentlyContinue).Status -eq 'Running')
    { 
      (' Stopping {0}...' -f $BlacklistService)
      Stop-Service -Name $BlacklistService -ErrorAction SilentlyContinue -Force -NoWait
    }
  }
  
  'Ending unneeded processes...'
  'Ending unneeded processes...' | timestamp >> $LogFile
  ('Getting {0} and {1}...' -f $ProcessBlacklistFile, $ServiceBlacklistFile)
  $ProcessBlacklist = Get-Content -Path $ProcessBlacklistFile -ErrorAction SilentlyContinue
  $ServiceBlacklist = Get-Content -Path $ServiceBlacklistFile -ErrorAction SilentlyContinue
  
  'Closing all blacklisted processes...' | timestamp >> $LogFile
  'Closing all blacklisted processes...'
  ForEach ($BlacklistProcess in $ProcessBlacklist) 
  {
    Stop-Process -Name $BlacklistProcess -ErrorAction SilentlyContinue -Force
  }
  
  #'Resetting Explorer...'
  #Stop-Process -Name Explorer -ErrorAction SilentlyContinue -Force
  
  'Everything unneeded has been closed successfully.' | timestamp >> $LogFile
  'Everything unneeded has been closed successfully.'
  $shell.minimizeall()
  $EverythingClosed = $true
  'Finished with Close-Unneededstuff'
}
function Start-WhitelistedServices 
{
  $ServiceWhitelist = (Get-Content -Path $ServiceWhitelistFile -ErrorAction SilentlyContinue)
  'Starting whitelisted services again...' | timestamp >> $LogFile
  'Starting whitelisted services again...'
  ForEach ($WhitelistService in $ServiceWhitelist) 
  {
    Start-Service -Name $WhitelistService -ErrorAction SilentlyContinue
  }
  Get-Service -DisplayName '*defender*' -ErrorAction SilentlyContinue | Start-Service -ErrorAction SilentlyContinue
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
    Wait-ProcessToCalm -ProcessToCalm $BorderlessProcess
  }
}
function Set-Audio
{
  'Setting Audiodevice...'
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
  [int]$DefaultVolume = [int](Get-INIValue -Path $INIFile -Section 'Audio' -Key 'DefaultVolume')
  [int]$VRVolume = [int](Get-INIValue -Path $INIFile -Section 'Audio' -Key 'VRVolume')
  [int]$HeadphonesVolume = [int](Get-INIValue -Path $INIFile -Section 'Audio' -Key 'HeadphonesVolume')
  $Nightmode = Get-INIValue -Path $INIFile -Section 'Audio' -Key 'Nightmode'
  $NightmodeBegin = Get-Date '23:59:59'
  $NightmodeBegin = Get-Date (Get-INIValue -Path $INIFile -Section 'Audio' -Key 'NightmodeBegin')
  $NightmodeEnd = Get-Date '00:00:00'
  $NightmodeEnd = Get-Date (Get-INIValue -Path $INIFile -Section 'Audio' -Key 'NightmodeEnd')
  $NightmodeVolume = Get-INIValue -Path $INIFile -Section 'Audio' -Key 'NightmodeVolume'
    
  If (!((Get-INIValue -Path $INIFile -Section 'Audio' -Key ($System + 'Volume')) -eq '')) 
  {
    If (([string](Get-INIValue -Path $INIFile -Section 'Audio' -Key ($System + 'Volume')) -like '+*') -or ([string](Get-INIValue -Path $INIFile -Section 'Audio' -Key ($System + 'Volume')) -like '-*'))
    {
      [String]$SystemVolume = [String](Get-INIValue -Path $INIFile -Section 'Audio' -Key ($System + 'Volume'))
    }
    else
    {
      [int]$SystemVolume = [int](Get-INIValue -Path $INIFile -Section 'Audio' -Key ($System + 'Volume'))
    }
  }
  $VRAudioRecID = Get-AudioDevice -List |
  Where-Object -Property Type -EQ -Value 'Recording' |
  Where-Object -Property Name -Like -Value "*$VRAudioDevice*" |
  Select-Object -ExpandProperty ID
  $HeadphonesRecID = Get-AudioDevice -List |
  Where-Object -Property Type -EQ -Value 'Recording' |
  Where-Object -Property Name -Like -Value "*$HeadphonesDevice*" |
  Select-Object -ExpandProperty ID
  

  If (($DefaultVolume -eq $null) -or ($DefaultVolume -eq 0))
  {
    [int]$CurrentVolume = [int](Get-AudioDevice -PlaybackVolume).Replace('%','')
    [int]$DefaultVolume = [int]$CurrentVolume
  }

  If (($DefaultAudioID -ne $null) -and ($DefaultAudioID -ne ''))
  {
    Set-AudioDevice -ID $DefaultAudioID -ErrorAction SilentlyContinue
  }
  Set-AudioDevice -PlaybackMute 0 -ErrorAction SilentlyContinue
  Set-AudioDevice -PlaybackVolume $DefaultVolume -ErrorAction SilentlyContinue
  "Setting default volume to $DefaultVolume."

  If (($VRAudioRecID -ne $null) -and ($VRAudioRecID -ne ''))
  {
    Set-AudioDevice -ID $VRAudioRecID -ErrorAction SilentlyContinue
  }
  
  If (($Nightmode -eq 1) -or ($Nightmode -eq '1') -or ($Nightmode -eq $true))
  {
    If (((Get-Date) -lt $NightmodeEnd) -or ((Get-Date) -gt $NightmodeBegin))
    {
      Set-AudioDevice -PlaybackVolume ((Get-AudioDevice -PlaybackVolume).replace('%','') - $NightmodeVolume.Replace('-','')) -ErrorAction SilentlyContinue
      $CurrentVolume = (Get-AudioDevice -PlaybackVolume).replace('%','')
      "Adjusting default volume to $CurrentVolume because of nightmode."
    }
  }
         
  If ($HeadphonesID -ne $null) 
  {
    If ($HeadphonesRecID -ne $null) 
    {
      Set-AudioDevice -ID $HeadphonesRecID -ErrorAction SilentlyContinue
    }
    If ($HeadphonesID -ne $null) 
    {
      'Headphones detected.'
      Set-AudioDevice -ID $HeadphonesID -ErrorAction SilentlyContinue
      Set-AudioDevice -PlaybackMute 0 -ErrorAction SilentlyContinue
      Set-AudioDevice -PlaybackVolume $HeadphonesVolume -ErrorAction SilentlyContinue
      "Setting headphones volume to $HeadphonesVolume."
    }
  }

  [int]$CurrentVolume = (Get-AudioDevice -PlaybackVolume).replace('%','')
}
function Ready-VR
{
  If ($env:OculusBase -eq $null) 
  {
    $env:OculusBase = "$env:ProgramW6432\Oculus\"
  }
  
  $OculusClient = $env:OculusBase + 'Support\oculus-client\OculusClient.exe'
  $OculusClientDir = $env:OculusBase + 'Support\oculus-client\'
  
  'Starting Oculus drivers...'
  Get-PnpDevice -FriendlyName '*oculus*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
  Get-PnpDevice -FriendlyName '*rift*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
    
  $VRProcesses = 'steamvr_tutorial', 'Steamtours', 'vrmonitor', 'steamvr', 'vrdashboard', 'vrcompositor', 'vrserver', 'Home-Win64-Shipping', 'Home2-Win64-Shipping'
  ForEach ($VRPRocess in $VRProcesses) 
  {
    If (Get-Process -Name $VRPRocess -ErrorAction SilentlyContinue)
    {
      $CloseProcess = Get-Process -Name $VRPRocess -ErrorAction SilentlyContinue
      $CloseProcess.CloseMainWindow()
    }
    Get-Process -Name $VRPRocess -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  }
  
  Remove-Item -Path "$env:LOCALAPPDATA\openvr" -Force -Recurse -ErrorAction SilentlyContinue
  
  If ((Get-Service -Name OVRService -ErrorAction SilentlyContinue).Status -ne 'Running') 
  { 
    'Oculus Service not running. Starting it...'
    'Oculus Service not running. Starting it...' | timestamp >> $LogFile

    Say-Something -About StartOculusServices
    Say-Something -About PleaseWait
    
    Set-Service -Name OVRLibraryService -StartupType Manual -ErrorAction SilentlyContinue
    Set-Service -Name OVRService -StartupType Manual -ErrorAction SilentlyContinue
    Start-Service -Name OVRService -ErrorAction SilentlyContinue
    Wait-ProcessToCalm -ProcessToCalm OVRServer_x64 -CalmThreshold 15
  }
  If (!(Get-Process -Name OculusClient -ErrorAction SilentlyContinue)) 
  { 
    "Oculus Client not running. Starting $OculusClient"
    "Oculus Client not running. Starting $OculusClient" | timestamp >> $LogFile

    Say-Something -About StartOculusClient
    Say-Something -About PleaseWait
    
    Start-Service -Name OVRService -ErrorAction SilentlyContinue
    Start-Process -FilePath $OculusClient -WorkingDirectory $OculusClientDir -Verb RunAs
    #Wait-WindowAppear -ProcessName OculusClient -ErrorAction SilentlyContinue
    Wait-ProcessToCalm -ProcessToCalm OculusClient -CalmThreshold 15
    If (!(Get-Process -Name OculusClient -ErrorAction SilentlyContinue)) 
    {
      Start-Process -FilePath $OculusClient -WorkingDirectory $OculusClientDir -Verb RunAs
      #Wait-WindowAppear -ProcessName OculusClient -ErrorAction SilentlyContinue
      Wait-ProcessToCalm -ProcessToCalm OculusClient -CalmThreshold 15
    }
    Wait-ProcessToCalm -ProcessToCalm OculusDash -CalmThreshold 20
    $WaitforOculus = 0
    While (!(Get-Process -Name OculusClient -ErrorAction SilentlyContinue)) 
    {
      Start-Sleep -Milliseconds 250
      $WaitforOculus = $WaitforOculus + 1
      If ($WaitforOculus -eq 40) 
      {
        Restart-Service -Name OVRService -ErrorAction SilentlyContinue
        Start-Process -FilePath $OculusClient -WorkingDirectory $OculusClientDir -Verb RunAs    
      }
    }
    If (!(Get-Process -Name OculusClient -ErrorAction SilentlyContinue)) 
    { 
      'Oculus Client still not running. Starting it for the last time...'
      'Oculus Client still not running. Starting it for the last time...' | timestamp >> $LogFile

      Say-Something -About Problem
      Say-Something -About RestartOculus
      Say-Something -About PleaseWait
      
      Restart-Service -Name OVRService -ErrorAction SilentlyContinue
      Start-Process -FilePath $OculusClient -WorkingDirectory $OculusClientDir -Verb RunAs
    }
    Wait-ProcessToCalm -ProcessToCalm OculusClient
  }
  $WaitforOculus = 0
  While (!(Get-Process -Name OculusClient -ErrorAction SilentlyContinue)) 
  {
    Start-Sleep -Milliseconds 250
    $WaitforOculus = $WaitforOculus + 1
    If ($WaitforOculus -eq 60) 
    {
      Say-Something -About TakesLonger
      Say-Something -About PleaseWait
      
      Start-Service -Name OVRService -ErrorAction SilentlyContinue
      Start-Process -FilePath $OculusClient -WorkingDirectory $OculusClientDir -Verb RunAs    
    }
  }
  
  While (!(Get-Process -Name OculusDash -ErrorAction SilentlyContinue)) 
  {
    Start-Sleep -Milliseconds 250
  }
  
  $OculusEXEs = (Get-ChildItem -Path $env:OculusBase -Filter '*.exe' -Recurse).FullName
  
  Wait-ProcessToCalm -ProcessToCalm OculusClient
  Wait-ProcessToCalm -ProcessToCalm OculusDash
  . Set-Audio
}
function Stop-VR
{
  $VRProcesses = 'steamvr_tutorial', 'Steamtours', 'steamvr', 'vrmonitor', 'vrdashboard', 'vrcompositor', 'vrserver', 'OculusVR', 'Home-Win64-Shipping', 'Home2-Win64-Shipping', 'OculusClient'
  
  ForEach ($VRPRocess in $VRProcesses) 
  {
    If (Get-Process -Name $VRPRocess -ErrorAction SilentlyContinue)
    { 
      "Closing $VRPRocess..."
      $CloseProcesses = Get-Process -Name $VRPRocess -ErrorAction SilentlyContinue
      ForEach ($CloseProcess in $CloseProcesses) 
      {
        $CloseProcess.CloseMainWindow()
        Wait-Process -Name $CloseProcess -Timeout 3 -ErrorAction SilentlyContinue
        Stop-Process -Name $CloseProcess -ErrorAction SilentlyContinue -Force
      }
    }
  }
  $CloseProcesses = Get-Process -Name OculusClient -ErrorAction SilentlyContinue
  ForEach ($CloseProcess in $CloseProcesses) 
  {
    $CloseProcess.CloseMainWindow()
    Wait-Process -Name $CloseProcess -Timeout 3 -ErrorAction SilentlyContinue
    Stop-Process -Name $CloseProcess -ErrorAction SilentlyContinue -Force
  }
  'Making sure VR services are stopped...'
  'Making sure VR services are stopped...' | timestamp >> $LogFile   
  Stop-Service -Name OVRLibraryService -Force -ErrorAction SilentlyContinue -NoWait
  Stop-Service -Name OVRService -Force -ErrorAction SilentlyContinue -NoWait
  Set-Service -Name OVRLibraryService -StartupType Disabled
  Set-Service -Name OVRService -StartupType Disabled
  Stop-Service -Name OVRLibraryService -Force -ErrorAction SilentlyContinue -NoWait
  Stop-Service -Name OVRService -Force -ErrorAction SilentlyContinue -NoWait
  Stop-Process -Name OVRLibraryService -ErrorAction SilentlyContinue -Force
  Stop-Process -Name OVRService -ErrorAction SilentlyContinue -Force
      
  $VRProcesses = 'steamvr_tutorial', 'Steamtours', 'steamvr', 'vrmonitor', 'vrdashboard', 'vrcompositor', 'vrserver', 'OculusVR', 'Home-Win64-Shipping', 'Home2-Win64-Shipping', 'OculusClient', 'OVRRedir', 'OVRServiceLauncher', 'OVRServer_x64'
  ForEach ($VRPRocess in $VRProcesses) 
  {
    If (Get-Process -Name $VRPRocess -ErrorAction SilentlyContinue)
    {
      "Killing $VRPRocess..."
      $CloseProcesses = Get-Process -Name $VRPRocess -ErrorAction SilentlyContinue
      ForEach ($CloseProcess in $CloseProcesses) 
      {
        $CloseProcess.CloseMainWindow()
        Wait-Process -Name $CloseProcess -Timeout 3 -ErrorAction SilentlyContinue
        Stop-Process -Name $CloseProcess -ErrorAction SilentlyContinue -Force
      }
    }
  }
  Get-PnpDevice -FriendlyName '*oculus*' -ErrorAction SilentlyContinue | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
  Get-PnpDevice -FriendlyName '*rift*' -ErrorAction SilentlyContinue | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
}
function Wait-GameProcess
{
  $IgnoreProcesses = (Get-Content -Path $IgnoreProcessesFile -ErrorAction SilentlyContinue)
  $IgnoreDirectories = (Get-Content -Path $IgnoreDirectoriesFile -ErrorAction SilentlyContinue)
  $IgnoreDirectoriesEXEs = (Get-ChildItem -Path $IgnoreDirectories -Filter '*.exe').FullName
  'Starting detection of new processes...'
  If ($BeforeGameProcesses.Count -lt 1) 
  {
    $BeforeGameProcesses = (Get-Process -ErrorAction SilentlyContinue |
      Where-Object -Property Path -NotIn -Value $IgnoreDirectoriesEXEs |
      Where-Object -Property ProcessName -NotIn -Value $IgnoreProcesses |
    Where-Object -Property ProcessName -NotLike -Value '*.tmp').ProcessName | Get-Unique
  }
  Start-Sleep -Seconds 1
  $AfterGameProcesses = (Get-Process -ErrorAction SilentlyContinue |
    Where-Object -Property ProcessName -NotIn -Value $IgnoreProcesses |
    Where-Object -Property Path -NotIn -Value $IgnoreDirectoriesEXEs |
  Where-Object -Property ProcessName -NotLike -Value '*.tmp').ProcessName | Get-Unique
  $NewProcesses = (Compare-Object -ReferenceObject $BeforeGameProcesses -DifferenceObject $AfterGameProcesses -ErrorAction SilentlyContinue | Where-Object -Property SideIndicator -EQ -Value '=>').InputObject
  
  While ($NewProcesses.Count -lt 1) 
  { 
    Start-Sleep -Seconds 1
    $AfterGameProcesses = (Get-Process -ErrorAction SilentlyContinue |
      Where-Object -Property ProcessName -NotIn -Value $IgnoreProcesses |
      Where-Object -Property Path -NotIn -Value $IgnoreDirectoriesEXEs |
    Where-Object -Property ProcessName -NotLike -Value '*.tmp').ProcessName | Get-Unique
    $NewProcesses = (Compare-Object -ReferenceObject $BeforeGameProcesses -DifferenceObject $AfterGameProcesses -ErrorAction SilentlyContinue | Where-Object -Property SideIndicator -EQ -Value '=>').InputObject
  }
  
  "Process found: $NewProcesses"
  
  If ($NewProcesses -like '*launch*') 
  {
    'Found something that looks like a launcher.'
    While ($NewProcesses.Count -lt 2) 
    { 
      Start-Sleep -Seconds 2
      $AfterGameProcesses = (Get-Process -ErrorAction SilentlyContinue |
        Where-Object -Property ProcessName -NotIn -Value $IgnoreProcesses |
        Where-Object -Property Path -NotIn -Value $IgnoreDirectoriesEXEs |
      Where-Object -Property ProcessName -NotLike -Value '*.tmp').ProcessName | Get-Unique
      $NewProcesses = (Compare-Object -ReferenceObject $BeforeGameProcesses -DifferenceObject $AfterGameProcesses -ErrorAction SilentlyContinue | Where-Object -Property SideIndicator -EQ -Value '=>').InputObject
    }
    "Additional Process found: $NewProcesses"
  }
}
function Update-ReShade
{
  If (($Reshadesource -eq $null) -or ($Reshadesource -eq '')) 
  {
    $Reshadesource = 'None'
  }
  If (Test-Path $Reshadesource -ErrorAction SilentlyContinue) 
  { 
    IF ($Game -like '*.lnk') 
    {
      $GamePath = ''
      $GameShortcut = Get-Shortcut -Path $Game
      If ($GameShortcut.TargetPath -like '*:\*') 
      {
        $GamePath = ($GameShortcut.TargetPath).Replace($GameShortcut.Target,'')
      }
      $ReshadeFiles = Get-ItemProperty -Path ($GamePath + 'd3d9.dll'), ($GamePath + 'dxgi.dll'), ($GamePath + 'opengl32.dll') -ErrorAction SilentlyContinue
    }

    'Checking for ReShade Update...'
    ForEach ($ReshadeFile in $ReshadeFiles)
    {
      If ($ReshadeFile.VersionInfo.ProductName -eq 'ReShade') 
      { 
        $OldReshadeVersion = $ReshadeFile.VersionInfo.ProductVersion
        $ReShadeSourceDLL = (Get-ChildItem -Path $Reshadesource -Include $ReshadeFile.Name -Recurse | Sort-Object -Descending -Property LastWriteTime)[0]
        $NewReshadeVersion = $ReShadeSourceDLL.VersionInfo.ProductVersion
        If ($NewReshadeVersion -ne $OldReshadeVersion)
        {
          ('Updating Version {0} of {1} to {2}...' -f $OldReshadeVersion, $ReshadeFile.Name, $NewReshadeVersion)
          Copy-Item -Path $ReShadeSourceDLL -Destination $ReshadeFile -Force
          $ReshadeINI = ($ReshadeFile.FullName).Replace('.dll','.ini')
          Set-INIValue -Path $ReshadeINI -Section General -Key PerformanceMode -Value 1
          Set-INIValue -Path $ReshadeINI -Section General -Key TutorialProgress -Value 4
          Set-INIValue -Path $ReshadeINI -Section General -Key ShowClock -Value 0
          Set-INIValue -Path $ReshadeINI -Section General -Key ShowFPS -Value 0
          Set-INIValue -Path $ReshadeINI -Section General -Key NoReloadOnInit -Value 0
        }
      }
    }
  }
}
function Mount-CDROMImages
{
  'Checking for CD-ROM images...'
  ForEach ($DeviceNumber in 0..9) 
  {
    Get-DiskImage -DevicePath "\\.\CDROM$DeviceNumber" -ErrorAction SilentlyContinue | Dismount-DiskImage
  }
  IF ($Game -like '*.lnk') 
  {
    $GameShortcut = Get-Shortcut -Path $Game
    $GamePath = ($GameShortcut.TargetPath).Replace($GameShortcut.Target,'')
    $ISOImages = Get-ChildItem -Path $GamePath -Filter *.iso -ErrorAction SilentlyContinue | Sort-Object -Property Name
  }
  If ($ISOImages.Count -ge 1) 
  {
    ('Found {0} ISO image(s)...' -f $ISOImages.Count)
    ForEach ($ISOImage in $ISOImages) 
    {
      ('Mounting {0}.' -f $ISOImage.Name)
      Mount-DiskImage -ImagePath $ISOImage.FullName
    }

    If ($ISOImages.Count -eq 1) 
    {
      Say-Something -About MountDisc
      Start-Sleep -Seconds 1
    }
    If ($ISOImages.Count -gt 1) 
    {
      Say-Something -About MountDisc -Modifier s
      Start-Sleep -Seconds 2
    }
  }
}
function Detect-GameProcess
{
  param
  (
    [Parameter(Position = 0)]
    [string]
    $DetectionPhase
  )
  
  $IgnoreProcesses = (Get-Content -Path $IgnoreProcessesFile -ErrorAction SilentlyContinue)
  $IgnoreDirectories = (Get-Content -Path $IgnoreDirectoriesFile -ErrorAction SilentlyContinue)
  $IgnoreDirectoriesEXEs = (Get-ChildItem -Path $IgnoreDirectories -Filter '*.exe').FullName | Get-Unique
  $SteamPath = Get-Item -Path (((Get-Item -Path HKCU:\Software\Valve\Steam).GetValue('SteamEXE')).ToString()).Replace('steam.exe','')
  $SteamEXE = Get-Item -Path ((Get-Item -Path HKCU:\Software\Valve\Steam).GetValue('SteamEXE')).ToString()
  $SteamEXEs = (Get-ChildItem -Path $SteamPath, "$SteamPath\bin" -Filter '*.exe').FullName
  $OriginPath = Get-Item -Path (((Get-Item -Path HKLM:\SOFTWARE\WOW6432Node\Origin).GetValue('ClientPath')).ToString()).Replace('Origin.exe','')
  $OriginEXE = Get-Item -Path (((Get-Item -Path HKLM:\SOFTWARE\WOW6432Node\Origin).GetValue('ClientPath')).ToString())
  $OriginEXEs = (Get-ChildItem -Path $OriginPath -Filter '*.exe').FullName
    
  If ($DetectionPhase -eq 'prepare')
  {
    $TimeOfPreparation = Get-Date
    Remove-Variable -Name NewProcesses -Force -ErrorAction SilentlyContinue
    'Getting processes running...'
    $BeforeGameProcesses = (Get-Process -ErrorAction SilentlyContinue |
      Where-Object -Property ProcessName -NotIn -Value $IgnoreProcesses |
      Where-Object -Property Path -NotIn -Value $IgnoreDirectoriesEXEs |
      Where-Object -Property Path -NotIn -Value $SteamEXEs |
      Where-Object -Property Path -NotIn -Value $OriginEXEs |
      Where-Object -Property Path -NotLike -Value "$env:windir*" |
      Where-Object -Property Path -NotLike -Value "$env:OculusBase*" |
      Where-Object -Property Path -NotLike -Value '*\SteamVR\*' |
    Where-Object -Property ProcessName -NotLike -Value '*.tmp').ProcessName | Get-Unique
    ('Processes found: {0}' -f $BeforeGameProcesses.Count)
    $BeforeGameProcesses
    ''
  }
  else
  {
    Start-Sleep -Seconds 5
    Remove-Variable -Name NewProcesses -Force -ErrorAction SilentlyContinue
    'Looking for the game process...'
    While ($NewProcesses.Count -lt 1) 
    {
      Start-Sleep -Seconds 5
      $AfterGameProcesses = (Get-Process -ErrorAction SilentlyContinue |
        Select-Object -Property ProcessName, Path |
        Where-Object -Property ProcessName -NotIn -Value $IgnoreProcesses |
        Where-Object -Property Path -NotIn -Value $IgnoreDirectoriesEXEs |
        Where-Object -Property Path -NotIn -Value $SteamEXEs |
        Where-Object -Property Path -NotIn -Value $OriginEXEs |
        Where-Object -Property Path -NotLike -Value "$env:windir*" |
        Where-Object -Property Path -NotLike -Value "$env:OculusBase*" |
        Where-Object -Property Path -NotLike -Value '*\SteamVR\*' |
      Where-Object -Property ProcessName -NotLike -Value '*.tmp').ProcessName | Get-Unique
      $NewProcesses = (Compare-Object -ReferenceObject $BeforeGameProcesses -DifferenceObject $AfterGameProcesses -ErrorAction SilentlyContinue | Where-Object -Property SideIndicator -EQ -Value '=>').InputObject
    }
    $NewProcess = $NewProcesses
    "Processes found: $NewProcess"
    ''
  
    If (($NewProcess -like '*launch*') -or ($NewProcess -like '*start*') -or ($NewProcess -like '*update*') -or ($NewProcess -like '*install*') -or ($NewProcess -like '*config*') -or ($NewProcess -like '*setup*'))
    {
      Remove-Variable -Name NewProcesses -Force -ErrorAction SilentlyContinue
      $RunningGameLauncher = $NewProcess
      "Found something that looks like a launcher or installer: $RunningGameLauncher"
      'Looking for the proper game process...'
      $WaitingForProperGameProcess = 0
      While (($NewProcesses.Count -lt 1) -and ($WaitingForProperGameProcess -lt 12))
      {
        Start-Sleep -Seconds 5
        $AfterGameProcesses = (Get-Process -ErrorAction SilentlyContinue |
          Where-Object -Property ProcessName -NotIn -Value $IgnoreProcesses |
          Where-Object -Property Path -NotIn -Value $IgnoreDirectoriesEXEs |
          Where-Object -Property ProcessName -NotLike -Value '*.tmp' |
        Where-Object -Property ProcessName -NE -Value $RunningGameLauncher).ProcessName | Get-Unique
        $WaitingForProperGameProcess = $WaitingForProperGameProcess + 1
        $NewProcesses = (Compare-Object -ReferenceObject $BeforeGameProcesses -DifferenceObject $AfterGameProcesses -ErrorAction SilentlyContinue | Where-Object -Property SideIndicator -EQ -Value '=>').InputObject
      }
      If ($WaitingForProperGameProcess -ge 12)
      {
        'Nothing new has been started after 1 minute.'
      }
      else
      { 
        $NewProcess = $NewProcesses
        ('Additional Process found: {0}' -f $NewProcess)
      }
    }
    If ($NewProcess.Count -gt 1)
    {
      $NewProcess = $NewProcess[0]
    }
  }
}

'General functions loaded.'

###################################### End of functions #########################################


('Getting Settings from {0}...' -f $INIFile)
('Getting Settings from {0}...' -f $INIFile) | timestamp >> $LogFile

IF (($System -eq '') -or ($System -eq $null) -or ($LastGame -eq $true))
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
IF (($Game -eq '') -or ($Game -eq $null) -or ($LastGame -eq $true))
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
$GameExt = ($Game -split '\.')[-1]
$GameName = ($Game -Replace ('.{0}' -f $GameExt) -split '\\')[-1]

. Set-Audio

$TTSFeature = (Get-INIValue -Path $INIFile -Section 'Audio' -Key 'TTSFeature')


  
IF ($Menu -eq $true) 
{
  $System = 'Windows'
  $Game = 'AllLauncher Menu'
  $Greeting = Get-Random -Maximum ('Hello', 'Good day', 'Greetings', 'Welcome', 'Hello')
  If ([int](Get-Date -Format HH) -lt 9) 
  {
    $Greeting = Get-Random -Maximum ($Greeting, 'Good Morning')
  }
  If ([int](Get-Date -Format HH) -lt 7) 
  {
    $Greeting = 'Good Morning'
  }
  If ([int](Get-Date -Format HH) -gt 20) 
  {
    $Greeting = Get-Random -Maximum ($Greeting, 'Good Evening')
  }
  $Callsign = Get-Random -Maximum ('', $env:Username, 'Sir')
  $GameName = ''
  $DRMSystem = ''
}

IF ($TTSFeature -eq '1')
{
  [bool]$TTSFeature = $true
  $MyName = (Get-INIValue -Path $INIFile -Section 'Options' -Key 'MyName')
  $CallMe = (Get-INIValue -Path $INIFile -Section 'Options' -Key 'CallMe')
  . Say-Something
}
else 
{
  [bool]$TTSFeature = $false
}


Say-Something -About Greeting

  
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
$UHSDir = (Get-INIValue -Path $INIFile -Section 'Directories' -Key 'UHSFolder')
$UHSFile = "$UHSDir\$GameName.uhs"
[string]$MoveWindowsTo2ndScreen = (Get-INIValue -Path $INIFile -Section 'Options' -Key 'MoveWindowsTo2ndScreen')
IF ($MoveWindowsTo2ndScreen -eq '1')
{
  [bool]$MoveWindowsTo2ndScreen = $true
}
else 
{
  [bool]$MoveWindowsTo2ndScreen = $false
}
$DS4Folder = (Get-INIValue -Path $INIFile -Section 'Controllers' -Key 'DS4Folder')
[string]$UsePS4Pad = (Get-INIValue -Path $INIFile -Section 'Controllers' -Key 'UsePS4Pad')
IF ($UsePS4Pad -eq '1')
{
  [bool]$UsePS4Pad = $true
}
else 
{
  [bool]$UsePS4Pad = $false
}
[string]$UseJoyMapper = (Get-INIValue -Path $INIFile -Section 'JoyMapper' -Key 'UseJoyMapper')
IF ($UseJoyMapper -eq '1')
{
  [bool]$UseJoyMapper = $true
}
else 
{
  [bool]$UseJoyMapper = $false
}
$JoyMapperName = (Get-INIValue -Path $INIFile -Section 'JoyMapper' -Key 'Name')
$JoyMapperExt = (Get-INIValue -Path $INIFile -Section 'JoyMapper' -Key 'Extension')
$JoyMapperDir = (Get-INIValue -Path $INIFile -Section 'JoyMapper' -Key 'Folder')
$JoyMapperExe = (Get-INIValue -Path $INIFile -Section 'JoyMapper' -Key 'Executable')
$JoyMapperProcess = $JoyMapperExe.Replace('.exe','')
[string]$UseCheats = (Get-INIValue -Path $INIFile -Section 'Options' -Key 'UseCheats')
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
$ArtmoneyProcess = (Get-INIValue -Path $INIFile -Section 'ArtMoney' -Key 'Executable') -replace '.exe'
$CheatEngineProcess = (Get-INIValue -Path $INIFile -Section 'CheatEngine' -Key 'Executable') -replace '.exe'
$CoSMOSProcess = (Get-INIValue -Path $INIFile -Section 'CoSMOS' -Key 'Executable') -replace '.exe'
[bool]$WasArtMoneyRunning = ([bool](Get-Process -Name $ArtmoneyProcess -ErrorAction SilentlyContinue))
[string]$UseDisplayFusion = (Get-INIValue -Path $INIFile -Section 'DisplayFusion' -Key 'UseDisplayFusion')
IF ($UseDisplayFusion -eq '1')
{
  [bool]$UseDisplayFusion = $true
}
else 
{
  [bool]$UseDisplayFusion = $false
}
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
$ArtmoneyProcess = $ArtMoneyExe -Replace '.exe'
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
$Reshadesource = (Get-INIValue -Path $INIFile -Section 'Options' -Key 'ReshadeSource')
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
$ThisPlayed = '0:00'
$HOTAS = (Get-INIValue -Path $INIFile -Section 'Controllers' -Key 'HOTAS')
$BorderlessEXE = (Get-INIValue -Path $INIFile -Section 'Borderless' -Key 'BorderlessEXE')
$BorderlessProcess = ($BorderlessEXE.Split('\\')[-1]) -replace('.exe')
$MIDISynthEXE = (Get-INIValue -Path $INIFile -Section 'Audio' -Key 'MIDISynthEXE')
$GameTrainers = Get-ChildItem -Path $TrainersDir -Filter '*.exe' | Where-Object -FilterScript {
  $_.Name -like "$GameName Trainer*"
}
$IgnoreProcesses = (Get-Content -Path $IgnoreProcessesFile -ErrorAction SilentlyContinue)
$IgnoreDirectories = (Get-Content -Path $IgnoreDirectoriesFile -ErrorAction SilentlyContinue)
$IgnoreDirectoriesEXEs = (Get-ChildItem -Path $IgnoreDirectories -Filter '*.exe').FullName
$SteamPath = Get-Item -Path (((Get-Item -Path HKCU:\Software\Valve\Steam).GetValue('SteamEXE')).ToString()).Replace('steam.exe','')
$SteamEXE = Get-Item -Path ((Get-Item -Path HKCU:\Software\Valve\Steam).GetValue('SteamEXE')).ToString()
$OriginPath = Get-Item -Path (((Get-Item -Path HKLM:\SOFTWARE\WOW6432Node\Origin).GetValue('ClientPath')).ToString()).Replace('Origin.exe','')
$OriginEXE = Get-Item -Path (((Get-Item -Path HKLM:\SOFTWARE\WOW6432Node\Origin).GetValue('ClientPath')).ToString())
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

Get-PnpDevice -Class HIDClass -ErrorAction SilentlyContinue |
Where-Object -FilterScript {
  $_.HardwareID.Contains('HID_DEVICE_SYSTEM_GAME')
} |
Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
Get-PnpDevice -FriendlyName '*xbox*control*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
Get-PnpDevice -FriendlyName '*game*controller*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
Get-PnpDevice -FriendlyName '*hid*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
If (!(Get-Process -Name explorer)) 
{
  Start-Process -FilePath $Explorer.Path -ErrorAction SilentlyContinue
}
$CursorRefresh::SystemParametersInfo(0x0057,0,$null,0)


''
'Settings loaded.' 
'Settings loaded.' | timestamp >> $LogFile



##########################################################################################
#################### Here are the different system functions #############################
##########################################################################################

function Start-Windows 
{
  'Starting a Windows-game...' | timestamp >> $LogFile
  $SpecialSystem = 1
  ''
  
  IF ($Game -like '*.lnk') 
  {
    'Checking for reshade...'
    $GameShortcut = Get-Shortcut -Path $Game
    If (!(($GameShortcut.TargetPath -notlike '*\*') -or ($GameShortcut.Target -eq 'n/a')))
    { 
      $GamePath = ($GameShortcut.TargetPath).Replace($GameShortcut.Target,'')
      $ReshadeFiles = Get-ItemProperty -Path ($GamePath + 'd3d9.dll'), ($GamePath + 'dxgi.dll'), ($GamePath + 'opengl32.dll') -ErrorAction SilentlyContinue
      ForEach ($ReshadeFile in $ReshadeFiles)
      {
        If ($ReshadeFile.VersionInfo.ProductName -eq 'ReShade') 
        {
          'Reshade found.'
          $Reshade = $true
        }
      }
      If ($Reshade -eq $true) 
      {
        . Update-ReShade
      }
    }
  }
  
  . Detect-GameProcess -DetectionPhase prepare
  
  "Invoking:  $Game"
  $GameLink = Get-Item -Path $Game
  Try
  {
    Invoke-Item -Path $GameLink -ErrorAction SilentlyContinue
  }
  Catch
  {
    Invoke-Item -Path "$Game" -ErrorAction SilentlyContinue
  }
  ''

  If (Get-Process -Name Origin -ErrorAction SilentlyContinue) 
  {
    Wait-ProcessToCalm -ProcessToCalm origin -CalmThreshold 10
  }
  If (Get-Process -Name Steam -ErrorAction SilentlyContinue) 
  {
    Wait-ProcessToCalm -ProcessToCalm steam -CalmThreshold 10
  }
  If ($System -eq 'VR') 
  {
    Wait-ProcessToCalm -ProcessToCalm oculusclient -CalmThreshold 10
    Wait-ProcessToCalm -ProcessToCalm steamvr -CalmThreshold 10
  }
  Get-Process '*update*', '*install*' -ErrorAction SilentlyContinue | ForEach-Object -Process {
    Wait-ProcessToCalm -ProcessToCalm oculusclient -CalmThreshold 5
  }

  . Detect-GameProcess
  
  Start-Sleep -Seconds 5
  [int]$WaitingForGameProcess = 0
  While ($WaitingForGameProcess -lt 5)
  { 
    If (Get-Process -Name $NewProcess -ErrorAction SilentlyContinue) 
    { 
      ('Setting {0} to high.' -f $NewProcess)
      Try 
      {
        Wait-WindowAppear -ProcessName $NewProcess -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 5
        If ([int](Get-Process -Name $NewProcess -ErrorAction SilentlyContinue).MainWindowHandle -gt 0)
        {
          Focus-Process -ProcessName $NewProcess -ErrorAction SilentlyContinue
        }
        (Get-Process -Name $NewProcess -ErrorAction SilentlyContinue).PriorityClass = 'HIGH'
      }
      Catch 
      {
        'Cannot set Priority.'
      }
    }  
  
    ('Waiting for:  {0}' -f $NewProcess)
    If (Get-Process -Name $NewProcess -ErrorAction SilentlyContinue) 
    {
      Try 
      {
        Wait-WindowAppear -ProcessName $NewProcess -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 5
        If ([int](Get-Process -Name $NewProcess -ErrorAction SilentlyContinue).MainWindowHandle -gt 0)
        {
          Focus-Process -ProcessName $NewProcess -ErrorAction SilentlyContinue
        }
        (Get-Process -Name $NewProcess -ErrorAction SilentlyContinue).PriorityClass = 'HIGH'
      }
      Catch 
      {
        'Cannot set Priority.'
      }
    }
    Get-Process -Name $NewProcess -ErrorAction SilentlyContinue | Wait-Process
    $WaitingForGameProcess = $WaitingForGameProcess + 1
  }

  'Game has ended.'  
  'Game has ended.' | timestamp >> $LogFile
}

function Start-VR
{
  If ($env:OculusBase -eq $null) 
  {
    $env:OculusBase = "$env:ProgramW6432\Oculus\"
  }
  $OculusClient = $env:OculusBase + 'Support\oculus-client\OculusClient.exe'
  $OculusClientDir = $env:OculusBase + 'Support\oculus-client\'
  $VRAudioDevice = Get-INIValue -Path $INIFile -Section 'Audio' -Key 'VRAudio'
  $VRAudioID = Get-AudioDevice -List |
  Where-Object -Property Type -EQ -Value 'Playback' |
  Where-Object -Property Name -Like -Value "*$VRAudioDevice*" |
  Select-Object -ExpandProperty ID
  $SpecialSystem = 1
    
  If ((Get-Service -Name OVRService -ErrorAction SilentlyContinue).Status -ne 'Running') 
  { 
    Set-Service -Name OVRLibraryService -StartupType Manual
    Set-Service -Name OVRService -StartupType Manual
    Start-Service -Name OVRService -ErrorAction SilentlyContinue
    Wait-ProcessToCalm -ProcessToCalm OVRServiceLauncher -CalmThreshold 5
    Wait-ProcessToCalm -ProcessToCalm OVRServer_x64 -CalmThreshold 5
  }
  If (!(Get-Process -Name OculusClient -ErrorAction SilentlyContinue)) 
  {
    Say-Something -About StartOculusClient
    
    Start-Process -FilePath $OculusClient -WorkingDirectory $OculusClientDir -Verb RunAs
    #Wait-WindowAppear -ProcessName OculusClient -ErrorAction SilentlyContinue
    Wait-ProcessToCalm -ProcessToCalm OculusClient -CalmThreshold 5
  }
  $WaitforOculus = 0
  While (!(Get-Process -Name OculusClient -ErrorAction SilentlyContinue)) 
  {
    Start-Sleep -Milliseconds 250
    $WaitforOculus = $WaitforOculus + 1
    If ($WaitforOculus -eq 40) 
    {
      Start-Service -Name OVRService -ErrorAction SilentlyContinue

      Say-Something -About PleaseWait
      
      Start-Process -FilePath $OculusClient -WorkingDirectory $OculusClientDir -Verb RunAs    
    }
  }
  
  #Wait-ProcessToCalm -ProcessToCalm OVRServiceLauncher -CalmThreshold 5
  #Wait-ProcessToCalm -ProcessToCalm OVRServer_x64 -CalmThreshold 5
  Wait-ProcessToCalm -ProcessToCalm OculusClient -CalmThreshold 5
  Wait-ProcessToCalm -ProcessToCalm OculusDash -CalmThreshold 5
    
  Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "' + (Get-INIValue -Path $INIFile -Section 'DisplayFusion' -Key 'FunctionVR') + '"') -Wait

      
  IF ($GameExt -eq 'lnk') 
  {
    'Checking game shortcut...'
    $GameEXE = (Get-Shortcut -Path $Game -ErrorAction SilentlyContinue)
    $GamePath = $GameEXE.TargetPath -replace $GameEXE.Target
    $steamAPI = $GamePath + 'steam_api64.dll'
    $steamVRPath = 'C:\STEAM\steamapps\common\SteamVR\bin\win64'
    $steamVRDLL = $steamVRPath + '\steam*.dll'
    $steamVRServer = $steamVRPath + '\vrserver.exe'
    $steamVR = $steamVRPath + '\vrmonitor.exe'
          
    If (Test-Path -Path $steamAPI) 
    { 
      'SteamVR is needed. Starting it...'

      Say-Something -About StartSteamVR
      
      #Copy-Item -Path $steamVRDLL -Destination $steamAPI -Force
      #Start-Sleep -Seconds 3
      Start-Process -FilePath 'steam://rungameid/250820'
      #Start-Process -FilePath $steamVRServer -WorkingDirectory $steamVRPath -Verb RunAs -ErrorAction SilentlyContinue
      While (!(Get-Process -Name vrmonitor -ErrorAction SilentlyContinue)) 
      {
        Start-Sleep -Seconds 1
      }
      Wait-ProcessToCalm -ProcessToCalm steam -CalmThreshold 3
      Wait-ProcessToCalm -ProcessToCalm vrserver -CalmThreshold 3
      #Start-Process -FilePath $steamVR -WorkingDirectory $steamVRPath -Verb RunAs -ErrorAction SilentlyContinue
      Wait-ProcessToCalm -ProcessToCalm vrmonitor -CalmThreshold 3
      Wait-ProcessToCalm -ProcessToCalm vrdashboard -CalmThreshold 3

      Say-Something -About SteamVRLoaded
    }
  }
  
  If (($VRAudioRecID -ne $null) -and ($VRAudioRecID -ne ''))
  {
    Set-AudioDevice -ID $VRAudioRecID -ErrorAction SilentlyContinue
  }
  If (($VRAudioID -ne $null) -and ($VRAudioID -ne ''))
  {
    'VR Headset detected.'
    Set-AudioDevice -ID $VRAudioID -ErrorAction SilentlyContinue
    Set-AudioDevice -PlaybackMute 0 -ErrorAction SilentlyContinue
    If (($VRVolume -ne $null) -and ($VRVolume -ne '')) 
    {
      Set-AudioDevice -PlaybackVolume $VRVolume -ErrorAction SilentlyContinue
      "Setting VR-headset volume to $VRVolume."
    }
  }
  
  . Start-Windows
  
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
  Start-Process -FilePath $Emulator -ArgumentList ('"' + $Game + '"') -Verb RunAs
  Start-Sleep -Seconds 3
  $CurrentGameProcess = Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
  Try 
  {
    Wait-WindowAppear -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
    (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue).PriorityClass = 'HIGH'
    Focus-Process -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
  }
  Catch 
  {
    'Cannot set Priority.'
  }
  Wait-Process -Name $CurrentGameProcess -ErrorAction SilentlyContinue
  Wait-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
  Wait-Process -Name winuae -ErrorAction SilentlyContinue
  Wait-Process -Name winuae64 -ErrorAction SilentlyContinue
  'Closing Amiga...' | timestamp >> $LogFile
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
  $DummyROM = ('"' + $Emulator.Directory + '\dummy.sfc"')
  $STDOUT = Get-Item -LiteralPath ('{0}\stdout.txt' -f $Emulator.Directory)
  $MednafenCFGs = Get-ChildItem -Path $Emulator.Directory -Name -Filter '*.cfg' -Recurse
  $IDRegex = '0x[0-9A-Fa-f]{32}'
  'Getting controller ID...'
  Start-Process -FilePath $Emulator -ArgumentList $DummyROM -Verb RunAs -WorkingDirectory $Emulator.Directory -Wait
  $NewJoystickIDs = [regex]::Matches((Get-Content -Path $STDOUT), $IDRegex)
  If ($NewJoystickIDs.Count -gt 1) 
  {
    $NewJoystickID = $NewJoystickIDs[0]
  }
  else 
  {
    $NewJoystickID = $NewJoystickIDs
  }
  ForEach ($MednafenCFG in $MednafenCFGs)
  {
    $MednafenCFGfile = ('{0}\{1}' -f $Emulator.Directory, $MednafenCFG)
    If ((Get-Content -Path $MednafenCFGfile).Length -gt 0) 
    { 
      $OldJoystickIDs = [regex]::Matches((Get-Content -Path $MednafenCFGfile), $IDRegex) | Get-Unique
      ForEach ($OldJoystickID in $OldJoystickIDs)
      {
        If ($OldJoystickID -notlike $NewJoystickID)
        {
          ((Get-Content -Path $MednafenCFGfile -Raw) -replace $OldJoystickID, $NewJoystickID) | Set-Content -Path $MednafenCFGfile
        }
      }
    }
  }
  
  'Starting Game with Mednafen...'
  Stop-Process -Name mednafen -Force -ErrorAction SilentlyContinue
  Start-Process -FilePath $Emulator -ArgumentList ('"' + $Game + '"') -Verb RunAs -WorkingDirectory $Emulator.Directory
  Start-Sleep -Seconds 2
  $CurrentGameProcess = Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
  Try 
  {
    If (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue) 
    {
      Wait-WindowAppear -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
      Focus-Process -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
    }
    $CurrentGameProcess.PriorityClass = 'HIGH'
    (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue).PriorityClass = 'HIGH'
    Focus-Process -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
  }
  Catch 
  {
    'Cannot set Priority.'
  }
  If (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue) 
  {
    Wait-WindowAppear -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
    Focus-Process -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
  }
  Wait-Process -Name $CurrentGameProcess -ErrorAction SilentlyContinue
  Wait-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
  Wait-Process -Name mednafen -ErrorAction SilentlyContinue
  'Quitting Mednafen...' | timestamp >> $LogFile
}

function Start-Emulator 
{
  If ($EmulatorINIEntry -eq 'mednafen') 
  {
    'Starting Mednafen...' | timestamp >> $LogFile
    Stop-Process -Name 'DS4Windows' -Force -ErrorAction SilentlyContinue
    $EmulatorINIEntry = Get-INIValue -Path $INIFile -Section 'Emulators' -Key 'Mednafen'
    . Start-Mednafen
  }
  else
  {
    'Starting a custom emulator...' | timestamp >> $LogFile
  
    If ($EmulatorINIEntry -eq $null)
    {
      "There is no emulator entry for $System in INI file!" | timestamp >> $LogFile
    }
    else
    { 
      $Emulator = Get-Item -Path (($EmulatorINIEntry -split '.exe')[0] + '.exe')
      $EmulatorDir = $Emulator.Directory.FullName
      $EmulatorProcess = $Emulator.Name.Replace('.exe','')
      $EmulatorArguments = ($EmulatorINIEntry -split '.exe')[1]
      ('Executing: {0} "{1}" {2}' -f $Emulator, $Game, $EmulatorArguments) | timestamp >> $LogFile
      Start-Process -FilePath $Emulator -ArgumentList ('"' + $Game + '"' + $EmulatorArguments) -Verb RunAs -WorkingDirectory $Emulator.Directory
      Start-Sleep -Seconds 3
      $GameProcesses = Get-Process $EmulatorProcess, 'medafen' -ErrorAction SilentlyContinue 
      ForEach ($CurrentGameProcess in $GameProcesses) 
      {
        Try 
        {
          If (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue) 
          {
            Wait-WindowAppear -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
          }
          $CurrentGameProcess.PriorityClass = 'HIGH'
          (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue).PriorityClass = 'HIGH'
          Focus-Process -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
        }
        Catch 
        {
          'Cannot set Priority.'
        }
      }
      Wait-Process -Name $GameProcesses.ProcessName -ErrorAction SilentlyContinue
      'Quitting emulator...' | timestamp >> $LogFile
    }
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

  Start-Process -FilePath $Emulator -ArgumentList $ScummVMArguments -WorkingDirectory $EmulatorDir -Verb RunAs
  Start-Sleep -Seconds 1
  $CurrentGameProcess = Get-Process -Name scummvm -ErrorAction SilentlyContinue
  Try 
  {
    If (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue) 
    {
      Wait-WindowAppear -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
    }
    $CurrentGameProcess.PriorityClass = 'HIGH'
    (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue).PriorityClass = 'HIGH'
    Focus-Process -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
  }
  Catch 
  {
    'Cannot set Priority.'
  }
  Wait-Process -Name $CurrentGameProcess, $EmulatorProcess, scummvm -ErrorAction SilentlyContinue
  Start-Sleep -Seconds 1
  Wait-Process -Name $CurrentGameProcess, $EmulatorProcess, scummvm -ErrorAction SilentlyContinue
  'Quitting game...' | timestamp >> $LogFile
}

function Start-C64
{
  'Starting a C64-game...' | timestamp >> $LogFile
  $SpecialSystem = 1
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
  Start-Process -FilePath $Emulator -ArgumentList $Arguments -WorkingDirectory $EmulatorDir -Verb RunAs
  Start-Sleep -Seconds 3
  $CurrentGameProcess = Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
  Try 
  {
    If (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue) 
    {
      Wait-WindowAppear -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
    }
    $CurrentGameProcess.PriorityClass = 'HIGH'
    (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue).PriorityClass = 'HIGH'
    Focus-Process -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
  }
  Catch 
  {
    'Cannot set Priority.'
  }
  Wait-Process -Name $CurrentGameProcess -ErrorAction SilentlyContinue
  Wait-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
  Wait-Process -Name x64 -ErrorAction SilentlyContinue
  Wait-Process -Name c64 -ErrorAction SilentlyContinue
  'Closing C64...' | timestamp >> $LogFile
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
      
  Start-Process -FilePath $DOSGame.TargetPath -ArgumentList $DOSGame.Arguments -WorkingDirectory ($DOSGame.TargetPath -replace $DOSGame.Target) -Verb RunAs -Wait
  Start-Sleep -Seconds 3
  $CurrentGameProcess = Get-Process -Name '*dosbox*' -ErrorAction SilentlyContinue
  Try 
  {
    If (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue) 
    {
      Wait-WindowAppear -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
    }
    $CurrentGameProcess.PriorityClass = 'HIGH'
    (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue).PriorityClass = 'HIGH'
    Focus-Process -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
  }
  Catch 
  {
    'Cannot set Priority.'
  }
  Wait-Process -Name $CurrentGameProcess -ErrorAction SilentlyContinue
  Wait-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
  Start-Sleep -Milliseconds 500
  'Quitting DOS game...' | timestamp >> $LogFile
  Copy-Item -Path ('{0}\*.sav' -f $SaveDir) -Destination $GameSaveDir -Force -ErrorAction SilentlyContinue
}

function Start-DOSBox 
{
  . Start-DOS
}

function Start-InteractiveFiction
{
  'Starting an Interactive Fiction-game...' | timestamp >> $LogFile
  $SpecialSystem = 1
  $IFGame = Get-Shortcut -Path $Game
  If (($IFGame.Arguments -eq '') -or ($IFGame.Arguments -eq $null)) 
  {
    'There do not seem to be any arguments.'
    $IFGame.Arguments = ' '
  }
  Start-Process -FilePath $IFGame.TargetPath -ArgumentList $IFGame.Arguments -WorkingDirectory ($IFGame.TargetPath -replace $IFGame.Target) -Wait
  #Start-Sleep -Seconds 5
  #Wait-Process -Name ($IFGame.Target -replace '.exe') -ErrorAction SilentlyContinue
  'Quitting game...' | timestamp >> $LogFile
}

function Start-NintendoDS 
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
  Start-Sleep -Seconds 3
  Send-Keys -KeysToSend '%{ENTER}'
  Start-Sleep -Seconds 1
  $CurrentGameProcess = Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
  Try 
  {
    Wait-WindowAppear -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
    $CurrentGameProcess.PriorityClass = 'HIGH'
    (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue).PriorityClass = 'HIGH'
    Focus-Process -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
  }
  Catch 
  {
    'Cannot set Priority.'
  }
  Wait-Process -Name $CurrentGameProcess -ErrorAction SilentlyContinue
  Wait-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
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
  Start-Process -FilePath $Emulator -ArgumentList ('"' + $Game + '" --fullscreen --fullboot') -Verb RunAs
  Start-Sleep -Seconds 3
  $CurrentGameProcess = Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
  Try 
  {
    Wait-WindowAppear -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
    $CurrentGameProcess.PriorityClass = 'HIGH'
    (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue).PriorityClass = 'HIGH'
    Focus-Process -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
  }
  Catch 
  {
    'Cannot set Priority.'
  }
  Wait-Process -Name $CurrentGameProcess -ErrorAction SilentlyContinue
  Wait-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
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
  "Starting $Emulator"
  'with Arguments:'
  ('-e "' + $Game + '"')
  Start-Process -FilePath $Emulator -ArgumentList ('-e "' + $Game + '"')
  Start-Sleep -Seconds 5
  $CurrentGameProcess = Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
  Try 
  {
    Wait-WindowAppear -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
    $CurrentGameProcess.PriorityClass = 'HIGH'
    (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue).PriorityClass = 'HIGH'
    Focus-Process -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
  }
  Catch 
  {
    'Cannot set Priority.'
  }
  Wait-Process -Name $CurrentGameProcess -ErrorAction SilentlyContinue
  Wait-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
  Wait-Process -Name Dolphin -ErrorAction SilentlyContinue
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
  $EmulatorINIEntry = (Get-INIValue -Path $INIFile -Section 'Emulators' -Key $System)
  $EmulatorArguments = ((Get-INIValue -Path $INIFile -Section 'Emulators' -Key 'RetroArch') -split '.exe')[1]
  $Arguments = '-L "' + $EmulatorDir + '\cores\' + $EmulatorINIEntry + '" "' + $Game + '"'
  ('Using Core: {0}' -f $EmulatorINIEntry) | timestamp >> $LogFile
  'Starting RetroArch'
  ('Executing: {0} {1}' -f $Emulator, $Arguments) | timestamp >> $LogFile
  Stop-Process -Name $EmulatorProcess, RetroArch, libretro -Force -ErrorAction SilentlyContinue
  Start-Process -FilePath $Emulator -ArgumentList $Arguments -WorkingDirectory $EmulatorDir -Verb RunAs
  Start-Sleep -Seconds 3
  $CurrentGameProcess = Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
  Try 
  {
    Wait-WindowAppear -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
    (Get-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue).PriorityClass = 'HIGH'
    Focus-Process -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
    $CurrentGameProcess.PriorityClass = 'HIGH'
  }
  Catch 
  {
    'Cannot set Priority.'
  }
  Wait-Process -Name $CurrentGameProcess -ErrorAction SilentlyContinue
  Wait-Process -Name $EmulatorProcess -ErrorAction SilentlyContinue
  Wait-Process -Name RetroArch -ErrorAction SilentlyContinue
  'Finished'
  'Quitting game...' | timestamp >> $LogFile
}

function Start-Android 
{
  'Starting an Android-game...' | timestamp >> $LogFile
  $SpecialSystem = 1
  $IgnoreProcesses = (Get-Content -Path $IgnoreProcessesFile -ErrorAction SilentlyContinue)
  $IgnoreDirectories = (Get-Content -Path $IgnoreDirectoriesFile -ErrorAction SilentlyContinue)
  $IgnoreDirectoriesEXEs = (Get-ChildItem -Path $IgnoreDirectories -Filter '*.exe').FullName
  
  Get-Service -Name MEmuSVC -ErrorAction SilentlyContinue | Start-Service -ErrorAction SilentlyContinue
  
  'Getting processes running...'
  $BeforeGameProcesses = (Get-Process -ErrorAction SilentlyContinue |
    Where-Object -Property ProcessName -NotIn -Value $IgnoreProcesses |
    Where-Object -Property Path -NotIn -Value $IgnoreDirectoriesEXEs |
  Where-Object -Property ProcessName -NotLike -Value '*.tmp').ProcessName | Get-Unique
  ('Processes found: {0}' -f $BeforeGameProcesses.Count)

  "Invoking $Game"
  Invoke-Item -Path $Game -ErrorAction SilentlyContinue

  Start-Sleep -Seconds 1
  $AfterGameProcesses = (Get-Process -ErrorAction SilentlyContinue |
    Where-Object -Property ProcessName -NotIn -Value $IgnoreProcesses |
    Where-Object -Property Path -NotIn -Value $IgnoreDirectoriesEXEs |
  Where-Object -Property ProcessName -NotLike -Value '*.tmp').ProcessName | Get-Unique
  $NewProcesses = (Compare-Object -ReferenceObject $BeforeGameProcesses -DifferenceObject $AfterGameProcesses -ErrorAction SilentlyContinue | Where-Object -Property SideIndicator -EQ -Value '=>').InputObject
  
  While ($NewProcesses.Count -lt 1) 
  { 
    Start-Sleep -Seconds 1
    $AfterGameProcesses = (Get-Process -ErrorAction SilentlyContinue |
      Where-Object -Property ProcessName -NotIn -Value $IgnoreProcesses |
      Where-Object -Property Path -NotIn -Value $IgnoreDirectoriesEXEs |
    Where-Object -Property ProcessName -NotLike -Value '*.tmp').ProcessName | Get-Unique
    $NewProcesses = (Compare-Object -ReferenceObject $BeforeGameProcesses -DifferenceObject $AfterGameProcesses -ErrorAction SilentlyContinue | Where-Object -Property SideIndicator -EQ -Value '=>').InputObject
  }
  $RunningGameProcess = $NewProcesses
  "Process found: $RunningGameProcess"
  
  Start-Sleep -Seconds 3
  
  If (Get-Process -Name $RunningGameProcess -ErrorAction SilentlyContinue) 
  { 
    ('Setting {0} to high.' -f $RunningGameProcess)
    Try 
    {
      Wait-WindowAppear -ProcessName $EmulatorProcess -ErrorAction SilentlyContinue
      (Get-Process -Name $RunningGameProcess -ErrorAction SilentlyContinue).PriorityClass = 'HIGH'
      Focus-Process -ProcessName $RunningGameProcess -ErrorAction SilentlyContinue
    }
    Catch 
    {
      'Cannot set Priority.'
    }
  }  
  
  
  Start-Sleep -Seconds 5
  ('Waiting for:  {0}' -f $RunningGameProcess)
  Get-Process -Name $RunningGameProcess -ErrorAction SilentlyContinue | Wait-Process
  Get-Process -Name 'MEMU', 'Bluestacks' -ErrorAction SilentlyContinue | Wait-Process
  'Game has ended.'  
  'Game has ended.' | timestamp >> $LogFile
  Stop-Process -Name 'MEMUConsole' -Force -ErrorAction SilentlyContinue
  Get-Service -Name MEmuSVC -ErrorAction SilentlyContinue | Stop-Service -Force -ErrorAction SilentlyContinue -NoWait
  Stop-Process -Name 'MEMUConsole', 'MemuService' -Force -ErrorAction SilentlyContinue
}
######################################################################################

. Manage-DRMSystem

If ($Menu -eq $true) 
{
  'The resetting parameter to return back to the menu has been used.' | timestamp >> $LogFile
  'Loading your Launcher and exiting AllLauncher.' | timestamp >> $LogFile
  '-menu has been used: Starting Launcher'
  $System = 'Windows'
  $Game = 'AllLauncher Menu'
  $DRMSystem = ''
  . Stop-VR
  . Start-Launcher
  Start-Sleep -Seconds 360
  EXIT
}

$shell.minimizeall()
Hide-DesktopIcons

If ($LauncherQuit -eq $true) 
{
  If (Get-Process -Name $LauncherProcess -ErrorAction SilentlyContinue)
  {
    'Going to Quit-Launcher'
    . Quit-Launcher
  }
  Wait-Process -Name $LauncherProcess -Timeout 15 -ErrorAction SilentlyContinue
}


Say-Something -About OpeningSpeech


. Close-UnneededStuff
'Everything is closed!'

Hide-DesktopIcons

$HOTASGames = Get-Content -Path $HOTASGamesFile -ErrorAction SilentlyContinue
If (($HOTASGames -contains $GameName) -and (Get-PnpDevice -FriendlyName "*$HOTAS*"))
{
  $UsePS4Pad = $false
  . Set-HOTAS 
}
  
'Checking for PS4 Gamepad'
If ($UsePS4Pad -eq $true) 
{
  . Set-DS4
  'DS4 ready.'
}
else
{
  'Using PS4-Pad is disabled.'
}

'Checking for Joy Mapping'
If ($UseJoyMapper -eq $true) 
{
  . Set-JoyMapper
  'JoyMapper finished.'
}
else
{
  'Using a controller mapping is disabled.'
}

'Writing INI setting for this game'
Set-INIValue -Path $INIFile -Section 'LastGame' -Key 'System' -Value $System
Set-INIValue -Path $INIFile -Section 'LastGame' -Key 'Game' -Value $Game
Set-INIValue -Path $INIFile -Section $System -Key 'Game' -Value $Game

If ($System -eq 'VR') 
{
  'This is a VR game. Making sure hardware is ready...'

  Say-Something -About ReadyVR
  
  'This is a VR game. Making sure hardware is ready...' | timestamp >> $LogFile
  . Ready-VR
}
else
{
  'Non-VR. Removing VR stuff...'
  'Non-VR. Removing VR stuff...' | timestamp >> $LogFile
  . Stop-VR
}

$shell.minimizeall()
Hide-DesktopIcons

IF ($Game -like '*.lnk') 
{
  . Mount-CDROMImages
}

$IgnoreDirectories = (Get-Content -Path $IgnoreDirectoriesFile -ErrorAction SilentlyContinue)
$IgnoreDirectoriesEXEs = (Get-ChildItem -Path $IgnoreDirectories -Filter '*.exe').FullName
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
  . Start-Cheats
  While ([bool]$CheatsReady -ne $true) 
  {
    Start-Sleep -Seconds 1
  }
  'Cheats ready'
}

. Open-GameDocs
'Docs ready'


'Waiting for the computer to calm down...'
$SWaitCounter = 0

If ((((((Get-Counter -Counter '\Processor(_total)\% Processor Time' -ErrorAction SilentlyContinue).CounterSamples).CookedValue) -ge 20) -or ((((Get-Counter -Counter '\physicaldisk(_total)\% disk time' -ErrorAction SilentlyContinue).CounterSamples).CookedValue) -ge 20)))
{
  'High CPU or disc usage!'

  Say-Something -About Loading
  Say-Something -About PleaseWait
}

While ((((((Get-Counter -Counter '\Processor(_total)\% Processor Time' -ErrorAction SilentlyContinue).CounterSamples).CookedValue) -gt 10) -or ((((Get-Counter -Counter '\physicaldisk(_total)\% disk time' -ErrorAction SilentlyContinue).CounterSamples).CookedValue) -gt 10)) -and ($SWaitCounter -lt 35))
{
  Start-Sleep -Seconds 1
  $SWaitCounter = $SWaitCounter + 1
}

If ($UseDisplayFusion -eq $true)
{
  'Using DisplayFusion...' | timestamp >> $LogFile
  If ($System -eq 'VR') 
  {
    Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "' + (Get-INIValue -Path $INIFile -Section 'DisplayFusion' -Key 'FunctionVR') + '"') -Wait
  }
  else
  {
    Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "' + (Get-INIValue -Path $INIFile -Section 'DisplayFusion' -Key 'Function') + '"') -Wait
  }
}
If ($MoveWindowsTo2ndScreen -eq $true)
{
  'Moving Windows to 2nd screen...'
  $OpenWindows = Get-Process -ErrorAction SilentlyContinue |
  Where-Object -Property ProcessName -NE -Value '' |
  Where-Object -Property MainWindowHandle -NE -Value 0 |
  Where-Object -Property ProcessName -NE -Value $null |
  Where-Object -Property ProcessName -NE -Value 'explorer'
  
  ForEach ($OpenWindow in $OpenWindows)
  {
    If ((Get-Process -Name $OpenWindow.ProcessName -ErrorAction SilentlyContinue) -and (!(((Get-WindowSizeAndPos -ProcessName $OpenWindow.ProcessName -ErrorAction SilentlyContinue).TopLeft.X -lt 0) -and ((Get-WindowSizeAndPos -ProcessName $OpenWindow.ProcessName -ErrorAction SilentlyContinue).TopLeft.Y -lt 0))))
    { 
      ('Moving {0} to display 2...' -f $OpenWindow.ProcessName)
      #Set-WindowStyle -ProcessName (Get-Process -Name $OpenWindow.ProcessName -ErrorAction SilentlyContinue).ProcessName -Style RESTORE -ErrorAction SilentlyContinue
      Move-Window -ProcessName (Get-Process -Name $OpenWindow.ProcessName -ErrorAction SilentlyContinue).ProcessName -center -Screen 2 -ErrorAction SilentlyContinue
    }
  }
  ''
  #Get-Process opera, firefox, chrome, steam, origin -ErrorAction SilentlyContinue | Set-WindowStyle -Style MAXIMIZE -ErrorAction SilentlyContinue
  #Get-Process -Name *read* -ErrorAction SilentlyContinue | Focus-Process -ErrorAction SilentlyContinue
  If (Get-Process -Name *read* -ErrorAction SilentlyContinue) 
  {
    Get-Process -Name *read* -ErrorAction SilentlyContinue | Set-WindowStyle -Style RESTORE -ErrorAction SilentlyContinue
    Get-Process -Name *read* -ErrorAction SilentlyContinue | Move-Window -center -Screen 2 -ErrorAction SilentlyContinue
    Get-Process -Name *read* -ErrorAction SilentlyContinue | Set-WindowStyle -Style SHOWMAXIMIZED -ErrorAction SilentlyContinue
  }
}

$SWaitCounter = 0

While ((((((Get-Counter -Counter '\Processor(_total)\% Processor Time' -ErrorAction SilentlyContinue).CounterSamples).CookedValue) -gt 7) -or ((((Get-Counter -Counter '\physicaldisk(_total)\% disk time' -ErrorAction SilentlyContinue).CounterSamples).CookedValue) -gt 7)) -and ($SWaitCounter -lt 25))
{
  Start-Sleep -Seconds 1
  $SWaitCounter = $SWaitCounter + 1
}

$PreviousVolume = (Get-AudioDevice -PlaybackVolume).replace('%','')

If (($SystemVolume -ne $null) -and ($SystemVolume -ne '') -and ($System -ne 'VR'))
{
  If ($SystemVolume -like '+*') 
  {
    "$System is set to a volume adjustment of $SystemVolume."
    [int]$CurrentSystemVolume = [int]$PreviousVolume + ($SystemVolume.Replace('+',''))
  }
  If ($SystemVolume -like '-*') 
  {
    "$System is set to a volume adjustment of $SystemVolume."
    [int]$CurrentSystemVolume = [int]$PreviousVolume - ($SystemVolume.Replace('-',''))
  }
  If (($SystemVolume -notlike '+*') -and ($SystemVolume -notlike '-*'))
  {
    "$System is set to a volume of $SystemVolume."
    [int]$CurrentSystemVolume = $SystemVolume
  }
}


Say-Something -About StartGame


If (($CurrentSystemVolume -ge 0) -and ($CurrentSystemVolume -le 100) -and ($System -ne 'VR')) 
{
  "Setting volume from $PreviousVolume to $CurrentSystemVolume..."
  Set-AudioDevice -PlaybackVolume $CurrentSystemVolume -ErrorAction SilentlyContinue
}

'Getting ready for:'
'Getting ready for:' | timestamp >> $LogFile

$StartDate = Get-Date



######################## Here starts the game #################################
''
''
'' | Out-File $LogFile -Append
('{0}' -f $GameName)
('{0}' -f $GameName) | timestamp >> $LogFile
('Going to {0}' -f $StartCurrentSystem)

Try 
{
  "Trying $StartCurrentSystem..."
  . $StartCurrentSystem
}
Catch 
{
  'No such entry!'
  'Using emulator settings from ini...'
  If ($EmulatorINIEntry -like '*:\*.exe*')
  {
    'Trying emulator from ini...'
    . Start-Emulator
  }
  If ($EmulatorINIEntry -like '*libretro.dll') 
  {
    'Trying RetroArch...'
    . Start-RetroArch
  }
  If ($EmulatorINIEntry -like 'mednafen') 
  {
    'Trying Mednafen...'
    Stop-Process -Name 'DS4Windows' -Force -ErrorAction SilentlyContinue
    $EmulatorINIEntry = Get-INIValue -Path $INIFile -Section 'Emulators' -Key 'Mednafen'
    . Start-Mednafen
  }
}
'' | Out-File $LogFile -Append
('{0} has ended.' -f $GameName)
('{0} has ended.' -f $GameName) | timestamp >> $LogFile
''
''
###############################################################################

$EndDate = Get-Date

. Set-Audio
Say-Something -About GameEnded


$IgnoreProcesses = (Get-Content -Path $IgnoreProcessesFile -ErrorAction SilentlyContinue)
$IgnoreDirectories = (Get-Content -Path $IgnoreDirectoriesFile -ErrorAction SilentlyContinue)
$IgnoreDirectoriesEXEs = (Get-ChildItem -Path $IgnoreDirectories -Filter '*.exe').FullName


ForEach ($GameTrainer in $GameTrainers) 
{
  Stop-Process -Name $GameTrainer -Force -ErrorAction SilentlyContinue
}

$CloseProcesses = Get-Process -Name '*launch*', '*start*', '*conf*', '*setup*' -ErrorAction SilentlyContinue |
Where-Object -Property ProcessName -NE -Value $LauncherProcess |
Where-Object -Property Id -NE -Value $pid
ForEach ($CloseProcess in $CloseProcesses) 
{
  $null = $CloseProcess.CloseMainWindow()
}

[timespan]$Playtime = $EndDate - $StartDate
'Calculating time played...' | timestamp >> $LogFile
'Calculating time played...'

$Playtime

If ($Playtime.TotalMinutes -lt 5)
{
  'Played for less than 5 Minutes, this does not count.' | timestamp >> $LogFile 
  'Played for less than 5 Minutes, this does not count.'
}
else 
{
  ('Played for {0} minutes.' -f $Playtime.TotalMinutes) | timestamp >> $LogFile
  $LogFileName = $LogFileDir + '\Complete_Game_Log_' + $StartDate.Year + '.txt'
  $GameLogName = $LogFileSystemDir + '\' + $GameName + '.txt'
  If (Test-Path -Path $GameLogName)
  {
    $GameLog = Import-Csv -Path $GameLogName -Delimiter "`t" -ErrorAction SilentlyContinue
    If ($GameLog[-1].Total -ne $null) 
    {
      $LastEntry = ($GameLog[-1].Total)
      [int]$TotalHoursLT = ($LastEntry.Split(':'))[0]
      [int]$DaysCalc = [Math]::floor($TotalHoursLT / 24)
      [String]$HoursLT = $TotalHoursLT - ($DaysCalc * 24)
      [String]$MinutesLT = ($LastEntry.Split(':'))[1]
      [string]$DaysLT = [Math]::floor($TotalHoursLT / 24)
      $LastTime = [TimeSpan]([String]($DaysLT + ':' + $HoursLT + ':' + $MinutesLT + ':0'))
    }
    else 
    {
      $LastTime = ([TimeSpan]'0:00')
    }
  }
  else 
  {
    $LastTime = ([TimeSpan]'0:00')
  }
  $ThisDate = (Get-Date -Date $StartDate.Date -Format 'ddd, dd. MMM yyyy')
  $ThisStart = ('{0:HH}:{0:mm}' -f $StartDate)
  $ThisEnd = ('{0:HH}:{0:mm}' -f $EndDate)
  $ThisPlayed = ('{0}:{1}' -f $Playtime.Hours.ToString('0'), ($Playtime.Minutes).ToString('00'))
  $TotalTime = ([TimeSpan]$LastTime + [TimeSpan]$ThisPlayed)
  $ThisTotal = ('{0}:{1}' -f [Math]::floor($TotalTime.TotalHours).ToString('0'), $TotalTime.Minutes.ToString('00'))
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
  

  Say-Something -About TimePlayed
}


'Getting processes still open...'
$ProcessesAfter = Get-Process 
$ProcessesChanged = Compare-Object -ReferenceObject $ProcessesBefore -DifferenceObject $ProcessesAfter -PassThru |
Where-Object -Property SideIndicator -EQ -Value '=>' |
Get-Unique
'Open processes: '
$ProcessesChanged.ProcessName
If ($ProcessesChanged.Count -lt 1) 
{
  $ProcessesChanged = Get-Process -Name explorer -ErrorAction SilentlyContinue
}

'Closing open processes...'
ForEach ($ChangedProcess in $ProcessesChanged) 
{ 
  If ($ChangedProcess.ProcessName -ne 'explorer')
  {
    $CloseProcesses = Get-Process -Name $ChangedProcess.ProcessName -ErrorAction SilentlyContinue
    ForEach ($CloseProcess in $CloseProcesses) 
    {
      If ($CloseProcess -ne $null) 
      {
        If (Get-Process -Name $CloseProcess.ProcessName -ErrorAction SilentlyContinue) 
        {
          $null = (Get-Process -Name $CloseProcess.ProcessName -ErrorAction SilentlyContinue).CloseMainWindow()
        }
      }
    }
  }
}

$shell.MinimizeAll()

If ((Get-ChildItem -Path $CurrentGameDocsDir -Recurse | Measure-Object).Count -eq 0) 
{
  Remove-Item -Path $CurrentGameDocsDir -Force -ErrorAction SilentlyContinue
}

If ($JoyMapperProcess -ne $null)
{ 
  If (Get-Process $JoyMapperProcess -ErrorAction SilentlyContinue) 
  {
    (Get-Process $JoyMapperProcess -ErrorAction SilentlyContinue).CloseMainWindow()
  }
}

If ($ThisTotal -like '*:*') 
{ 
  $TotalHours = (($ThisTotal).Split(':')[0])
  If ([int]$TotalHours -gt '1') 
  {
    Say-Something -About TotalTimePlayed
  }
}


If ($UseDisplayFusion -eq $true)
{
  'Using DisplayFusion...' | timestamp >> $LogFile
  If ($System -eq 'VR') 
  {
    Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "' + (Get-INIValue -Path $INIFile -Section 'DisplayFusion' -Key 'ReturnVR') + '"') -Wait
  }
  else
  {
    Start-Process -FilePath "${env:ProgramFiles(x86)}\DisplayFusion\displayfusioncommand.exe" -ArgumentList ('-functionrun "' + (Get-INIValue -Path $INIFile -Section 'DisplayFusion' -Key 'Return') + '"') -Wait
  }
}

Get-PnpDevice -Class HIDClass -ErrorAction SilentlyContinue |
Where-Object -FilterScript {
  $_.HardwareID.Contains('HID_DEVICE_SYSTEM_GAME')
} |
Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
Get-PnpDevice -FriendlyName '*xbox*control*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
Get-PnpDevice -FriendlyName '*game*controller*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
Get-PnpDevice -FriendlyName '*hid*' -ErrorAction SilentlyContinue | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue

If ($WasLauncherRunning -eq $true) 
{
  ('Returning to {0}.' -f $LauncherName) | timestamp >> $LogFile
  Start-Launcher
}
else 
{ 
  'Not started from Launcher, so not going back' | timestamp >> $LogFile
  . Close-UnneededStuff
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
  'Resetting Explorer...'
  Stop-Process -Name Explorer -ErrorAction SilentlyContinue -Force
  
  If (!(Get-Process -Name explorer)) 
  {
    Start-Process -FilePath $Explorer.Path -ErrorAction SilentlyContinue
  }
  $CursorRefresh::SystemParametersInfo(0x0057,0,$null,0)

  Say-Something -About Finished
  
  Stop-Transcript -ErrorAction SilentlyContinue
  #('Everything went well, removing transcript {0}.' -f $Transcript) | timestamp > $LogFile
  #Remove-Item -Path $Transcript -Force
  Get-Process -Name 'powershell', 'alllauncher' -ErrorAction SilentlyContinue | Stop-Process -Force
  Exit
}
