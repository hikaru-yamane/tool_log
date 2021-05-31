########################### CSharp ###########################
$src = @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public static class WindowController {
    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    public static bool ActivateWindow(IntPtr handle) {
        bool hasRestored = ShowWindowAsync(handle, 9);
        bool hasMoved = SetForegroundWindow(handle);
        return hasRestored && hasMoved;
    }

    public static void SendCommand(string command) {
        SendKeys.SendWait(command);
    }
}
"@
Add-Type -TypeDefinition $src -Language CSharp -ReferencedAssemblies 'System.Windows.Forms'
function activateWindowCSharp($handle) {
  [WindowController]::ActivateWindow($handle)
}
function sendCommandCSharp($command) {
  [WindowController]::SendCommand($command)
}


########################### class ###########################
class User {
    [Object[]]$connectInfoList
    [Object[]]$fileList
    [Object[]]$teratermList
    
    User() {
        $this.connectInfoList = @(
            [ConnectInfo]::new('localhost', '2222', 'test')
            #[ConnectInfo]::new('192.168.56.150', '22', 'test')
        )
        $this.fileList = @(
            [File]::new('1')
            [File]::new('2')
        )
        $this.teratermList = @()
        0..($this.connectInfoList.Count - 1) |
        % {
            $this.teratermList += [Teraterm]::new($this.connectInfoList[$_], $this.fileList[$_])
        }
    }

    [void] connect() {
        $this.teratermList | % { $_.connect() }
    }

    [void] disconnect() {
        $this.teratermList | % { $_.disconnect() }
    }

    [void] startLogging() {
        $this.teratermList | % { $_.startLogging() }
    }

    [void] stopLogging() {
        $this.teratermList | % { $_.stopLogging() }
    }
}
class Teraterm {
    [ConnectInfo]$connectInfo
    [File]$file
    [Window]$window
    [TeratermCommand]$command
    [Time]$time
    
    Teraterm([ConnectInfo]$connectInfo, [File]$file) {
        $this.connectInfo = $connectInfo
        $this.file = $file
        $this.window = [Window]::new($connectInfo.get_hostName(), $connectInfo.get_userName())
        $this.command = [TeratermCommand]::new()
        $this.time = [Time]::new()
    }
    
    [void] connect() {
        Start-Process `
            -FilePath $this.file.get_exe() `
            -ArgumentList $this.file.get_ttl(),
                          $this.connectInfo.get_hostName(),
                          $this.connectInfo.get_portNum(),
                          $this.connectInfo.get_userName(),
                          $this.file.get_key()
        $this.waitToConnect($this.window.get_popName(), -1)
        $this.init()
    }

    [void] disconnect() {
        if ($this.activateWindow()) {
            $winNum = $this.getNumberOfWindow($this.window.get_mainName())
            $this.sendCommand($this.command.get_commandToCloseMainWindow())
            $this.waitToExecute($this.window.get_mainName(), $winNum - 1)
        }
    }

    [void] startLogging() {
        if ($this.activateWindow()) {
            $winNum = $this.getNumberOfWindow($this.window.get_logName())
            $this.sendCommand($this.command.get_commandToOpenFileMenu())
            $this.sendCommand($this.command.get_commandToSelectViewLog())
            $this.waitForFormToOpen()
            $this.sendCommand((Invoke-Expression $this.file.get_log()))
            $this.sendCommand($this.command.get_commandToEnter())
            $this.sendCommand($this.command.get_commandToOpenFileMenu())
            $this.sendCommand($this.command.get_commandToSelectShowLogDialog())
            $this.waitToExecute($this.window.get_logName(), $winNum + 1)
        }
    }

    [void] stopLogging() {
        if ($this.activateWindow()) {
            $this.sendCommand($this.command.get_commandToOpenFileMenu())
            $this.sendCommand($this.command.get_commandToSelectShowLogDialog())
            $winNum = $this.getNumberOfWindow($this.window.get_logName())
            $this.sendCommand($this.command.get_commandToCloseSubWindow())
            $this.waitToExecute($this.window.get_logName(), $winNum - 1)
        }
    }
    
    [void] init() {
        $winHandle = (
            Get-Process |
            ? { $_.MainWindowTitle -eq $this.window.get_mainName() } |
            % { $_.MainWindowHandle }
        )[-1]
        $this.window.set_mainHandle($winHandle)
    }
    
    [Int32] getNumberOfWindow($windowName) {
        return (Get-Process | ? { $_.MainWindowTitle -eq $windowName }).count
    }
    
    [Boolean] activateWindow() {
        return activateWindowCSharp $this.window.get_mainHandle()
    }

    [void] sendCommand($command) {
        sendCommandCSharp $command
    }
    
    [void] waitToConnect($targetWindowName, $targetChangeNum) {
        $startTime = Get-Date
        $oldWinNum = -1
        while ($true) {
            $winNum = $this.getNumberOfWindow($targetWindowName)
            $diff = $winNum - $oldWinNum
            $hasConnected = $diff -eq $targetChangeNum
            $isTimeout = (New-TimeSpan $startTime (Get-Date)).TotalMilliseconds -gt $this.time.get_waitToConnectTimeout()
            if ($hasConnected) { break }
            if ($isTimeout) { exit }
            $oldWinNum = $winNum
            Start-Sleep -Milliseconds $this.time.get_loopInterval()
        }
    }

    [void] waitToExecute($targetWindowName, $targetReachNum) {
        $startTime = Get-Date
        while ($true) {
            $winNum = $this.getNumberOfWindow($targetWindowName)
            $hasExecuted = $winNum -eq $targetReachNum
            $isTimeout = (New-TimeSpan $startTime (Get-Date)).TotalMilliseconds -gt $this.time.get_waitToExecuteTimeout()
            if ($hasExecuted) { break }
            if ($isTimeout) { exit }
            Start-Sleep -Milliseconds $this.time.get_loopInterval()
        }
    }

    [void] waitForFormToOpen() {
        $startTime = Get-Date
        Set-Clipboard ' '
        while ($true) {
            $sending = 'dummy'
            $this.sendCommand($sending)
            $this.sendCommand('^a')
            $this.sendCommand('^c')
            $sent = Get-Clipboard -Format Text
            $hasOpened = $sent -eq $sending
            $isTimeout = (New-TimeSpan $startTime (Get-Date)).TotalMilliseconds -gt $this.time.get_waitForFormToOpenTimeout()
            if ($hasOpened) { break }
            if ($isTimeout) { exit }
            Start-Sleep -Milliseconds $this.time.get_loopInterval()
        }
    }
}
class ConnectInfo {
    [String]$hostName
    [String]$portNum
    [String]$userName
    
    ConnectInfo([String]$hostName, [String]$portNum, [String]$userName) {
        $this.hostName = $hostName
        $this.portNum  = $portNum
        $this.userName = $userName
    }
}
class File {
    [String]$exe
    [String]$ttl
    [String]$key
    [String]$log

    File([String]$logName) {
        $this.exe = 'C:\Program Files (x86)\teraterm\ttpmacro.exe'
        $this.ttl = Join-Path $PSScriptRoot 'log.ttl'
        $this.key = Join-Path $PSScriptRoot 'id_rsa'
        $this.log = "`"C:\Users\yamane\Documents\`$(Get-Date -Format 'yyyyMMdd_HHmmss')_$($logName).log`""
    }
}
class Window {
    [String]$mainName
    [String]$popName
    [String]$logName
    [IntPtr]$mainHandle

    Window([String]$hostName, [String]$userName) {
        $this.mainName = "$($hostName) - $($userName)@localhost:~ VT"
        $this.popName  = 'MACRO - log.ttl'
        $this.logName  = 'Tera Term: ログ'
    }
}
class TeratermCommand {
    [String]$commandToOpenFileMenu
    [String]$commandToSelectViewLog
    [String]$commandToSelectShowLogDialog
    [String]$commandToEnter
    [String]$commandToCloseMainWindow
    [String]$commandToCloseSubWindow

    TeratermCommand() {
        $this.commandToOpenFileMenu = '%f'
        $this.commandToSelectViewLog = 'l'
        $this.commandToSelectShowLogDialog = 'w'
        $this.commandToEnter = '{ENTER}'
        $this.commandToCloseMainWindow = '^uexit{ENTER}'
        $this.commandToCloseSubWindow = '%{F4}'
    }
}
class Time {
    [Int32]$waitToConnectTimeout
    [Int32]$waitToExecuteTimeout
    [Int32]$waitForFormToOpenTimeout
    [Int32]$loopInterval

    Time() {
        $this.waitToConnectTimeout = 10000
        $this.waitToExecuteTimeout = 2000
        $this.waitForFormToOpenTimeout = 2000
        $this.loopInterval = 200
    }
}


########################### main関数 ###########################
function main() {
    ## 初期化
    $user = [User]::new()
    $message = ''
    $isConnecting = $true
    $operation = '0'
    $START_LOG_OPERATION  = '1'
    $STOP_LOG_OPERATION   = '2'
    $DISCONNECT_OPERATION = '3'
    
    ## 接続
    Write-Host 'サーバに接続します。接続中は操作しないでください。'
    Write-Host '接続中...'
    $user.connect() | Out-Null
    Write-Host "接続が完了しました。`r`n"

    ## 操作
    while ($isConnecting) {
        $message = ''
        switch ($operation) {
            $START_LOG_OPERATION {
                try {
                    $message += "2. ログ取得終了`r`n"
                    $message += '操作を選択してください[2]'
                    [ValidateSet('2')]$operation = Read-Host $message
                    $user.stopLogging() | Out-Null
                } catch {
                    if ($_.Exception -is [System.Management.Automation.ValidationMetadataException]) {
                        # 何もしない
                    }
                }
            
            }
            default {
                try {
                    $message += "1. ログ取得開始`r`n"
                    $message += "3. 切断`r`n"
                    $message += '操作を選択してください[1/3]'
                    [ValidateSet('1', '3')]$operation = Read-Host $message
                    switch ($operation) {
                        $START_LOG_OPERATION {
                            $user.startLogging() | Out-Null
                        }
                        $DISCONNECT_OPERATION {
                            $user.disconnect() | Out-Null
                            $isConnecting = $false
                        }
                    }
                } catch {
                    if ($_.Exception -is [System.Management.Automation.ValidationMetadataException]) {
                        # 何もしない
                    }
                }
            }
        }
    }
}

main
