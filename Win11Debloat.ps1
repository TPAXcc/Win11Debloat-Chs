#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess)]
param (
    [switch]$Silent,
    [switch]$Sysprep,
    [string]$User,
    [switch]$RunAppsListGenerator, [switch]$RunAppConfigurator,
    [switch]$RunDefaults, [switch]$RunWin11Defaults,
    [switch]$RunSavedSettings,
    [switch]$RemoveApps, 
    [switch]$RemoveAppsCustom,
    [switch]$RemoveGamingApps,
    [switch]$RemoveCommApps,
    [switch]$RemoveDevApps,
    [switch]$RemoveW11Outlook,
    [switch]$ForceRemoveEdge,
    [switch]$DisableDVR,
    [switch]$DisableTelemetry,
    [switch]$DisableBingSearches, [switch]$DisableBing,
    [switch]$DisableDesktopSpotlight,
    [switch]$DisableLockscrTips, [switch]$DisableLockscreenTips,
    [switch]$DisableWindowsSuggestions, [switch]$DisableSuggestions,
    [switch]$ShowHiddenFolders,
    [switch]$ShowKnownFileExt,
    [switch]$HideDupliDrive,
    [switch]$TaskbarAlignLeft,
    [switch]$HideSearchTb, [switch]$ShowSearchIconTb, [switch]$ShowSearchLabelTb, [switch]$ShowSearchBoxTb,
    [switch]$HideTaskview,
    [switch]$DisableStartRecommended,
    [switch]$DisableCopilot,
    [switch]$DisableRecall,
    [switch]$DisableWidgets, [switch]$HideWidgets,
    [switch]$DisableChat, [switch]$HideChat,
    [switch]$EnableEndTask,
    [switch]$ClearStart,
    [switch]$ClearStartAllUsers,
    [switch]$RevertContextMenu,
    [switch]$DisableMouseAcceleration,
    [switch]$DisableStickyKeys,
    [switch]$HideHome,
    [switch]$HideGallery,
    [switch]$ExplorerToHome,
    [switch]$ExplorerToThisPC,
    [switch]$ExplorerToDownloads,
    [switch]$ExplorerToOneDrive,
    [switch]$DisableOnedrive, [switch]$HideOnedrive,
    [switch]$Disable3dObjects, [switch]$Hide3dObjects,
    [switch]$DisableMusic, [switch]$HideMusic,
    [switch]$DisableIncludeInLibrary, [switch]$HideIncludeInLibrary,
    [switch]$DisableGiveAccessTo, [switch]$HideGiveAccessTo,
    [switch]$DisableShare, [switch]$HideShare
)


# 当 PowerShell 执行环境受安全策略限制时显示错误
if ($ExecutionContext.SessionState.LanguageMode -ne "FullLanguage") {
    Write-Host "错误：Win11Debloat 无法在您的系统上运行，PowerShell 执行受到安全策略限制" -ForegroundColor Red
    Write-Output ""
    Write-Output "按回车键退出..."
    Read-Host | Out-Null
    Exit
}


# 显示应用程序选择表单，允许用户选择要删除或保留的应用程序
function ShowAppSelectionForm {
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

    # 初始化表单对象
    $form = New-Object System.Windows.Forms.Form
    $label = New-Object System.Windows.Forms.Label
    $button1 = New-Object System.Windows.Forms.Button
    $button2 = New-Object System.Windows.Forms.Button
    $selectionBox = New-Object System.Windows.Forms.CheckedListBox 
    $loadingLabel = New-Object System.Windows.Forms.Label
    $onlyInstalledCheckBox = New-Object System.Windows.Forms.CheckBox
    $checkUncheckCheckBox = New-Object System.Windows.Forms.CheckBox
    $initialFormWindowState = New-Object System.Windows.Forms.FormWindowState

    $global:selectionBoxIndex = -1

    # 保存按钮事件处理
    $handler_saveButton_Click= 
    {
        if ($selectionBox.CheckedItems -contains "Microsoft.WindowsStore" -and -not $Silent) {
            $warningSelection = [System.Windows.Forms.Messagebox]::Show('您确定要卸载Microsoft Store吗？此应用无法轻易重新安装。', '确认操作', 'YesNo', 'Warning')
        
            if ($warningSelection -eq 'No') {
                return
            }
        }

        $global:SelectedApps = $selectionBox.CheckedItems

        # 创建存储已选应用的文件（如果不存在）
        if (!(Test-Path "$PSScriptRoot/CustomAppsList")) {
            $null = New-Item "$PSScriptRoot/CustomAppsList"
        } 

        Set-Content -Path "$PSScriptRoot/CustomAppsList" -Value $global:SelectedApps

        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }

    # 取消按钮事件处理
    $handler_cancelButton_Click= 
    {
        $form.Close()
    }

    $selectionBox_SelectedIndexChanged= 
    {
        $global:selectionBoxIndex = $selectionBox.SelectedIndex
    }

    $selectionBox_MouseDown=
    {
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            if ([System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Shift) {
                if ($global:selectionBoxIndex -ne -1) {
                    $topIndex = $global:selectionBoxIndex

                    if ($selectionBox.SelectedIndex -gt $topIndex) {
                        for (($i = ($topIndex)); $i -le $selectionBox.SelectedIndex; $i++){
                            $selectionBox.SetItemChecked($i, $selectionBox.GetItemChecked($topIndex))
                        }
                    }
                    elseif ($topIndex -gt $selectionBox.SelectedIndex) {
                        for (($i = ($selectionBox.SelectedIndex)); $i -le $topIndex; $i++){
                            $selectionBox.SetItemChecked($i, $selectionBox.GetItemChecked($topIndex))
                        }
                    }
                }
            }
            elseif ($global:selectionBoxIndex -ne $selectionBox.SelectedIndex) {
                $selectionBox.SetItemChecked($selectionBox.SelectedIndex, -not $selectionBox.GetItemChecked($selectionBox.SelectedIndex))
            }
        }
    }

    $check_All=
    {
        for (($i = 0); $i -lt $selectionBox.Items.Count; $i++){
            $selectionBox.SetItemChecked($i, $checkUncheckCheckBox.Checked)
        }
    }

    $load_Apps=
    {
        # 调整表单初始状态以防止最大化问题
        $form.WindowState = $initialFormWindowState

        # 重置选择框索引
        $global:selectionBoxIndex = -1
        $checkUncheckCheckBox.Checked = $False

        # 显示加载指示器
        $loadingLabel.Visible = $true
        $form.Refresh()

        # 清空选择框
        $selectionBox.Items.Clear()

        # 设置应用列表文件路径
        $appsFile = "$PSScriptRoot/Appslist.txt"
        $listOfApps = ""

        if ($onlyInstalledCheckBox.Checked -and ($global:wingetInstalled -eq $true)) {
            # 尝试通过 winget 获取已安装应用列表（10秒超时）
            $job = Start-Job { return winget list --accept-source-agreements --disable-interactivity }
            $jobDone = $job | Wait-Job -TimeOut 10

            if (-not $jobDone) {
                [System.Windows.MessageBox]::Show('无法通过 winget 加载已安装应用列表，部分应用可能未显示。', '错误', 'Ok', 'Error')
            }
            else {
                $listOfApps = Receive-Job -Job $job
            }
        }

        # 加载应用列表
        Foreach ($app in (Get-Content -Path $appsFile | Where-Object { $_ -notmatch '^\s*$' -and $_ -notmatch '^#  .*' -and $_ -notmatch '^# -* #' } )) { 
            $appChecked = $true

            if ($app.StartsWith('#')) {
                $app = $app.TrimStart("#")
                $appChecked = $false
            }

            if (-not ($app.IndexOf('#') -eq -1)) {
                $app = $app.Substring(0, $app.IndexOf('#'))
            }
            
            $app = $app.Trim()
            $appString = $app.Trim('*')

            if ($appString.length -gt 0) {
                if ($onlyInstalledCheckBox.Checked) {
                    if (-not ($listOfApps -like ("*$appString*")) -and -not (Get-AppxPackage -Name $app)) {
                        continue
                    }
                    if (($appString -eq "Microsoft.Edge") -and -not ($listOfApps -like "* Microsoft.Edge *")) {
                        continue
                    }
                }

                $selectionBox.Items.Add($appString, $appChecked) | Out-Null
            }
        }
        
        $loadingLabel.Visible = $False
        $selectionBox.Sorted = $True
    }

    $form.Text = "Win11Debloat 应用选择"
    $form.Name = "appSelectionForm"
    $form.DataBindings.DefaultDataSourceUpdateMode = 0
    $form.ClientSize = New-Object System.Drawing.Size(400,502)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $False

    $button1.TabIndex = 4
    $button1.Name = "saveButton"
    $button1.UseVisualStyleBackColor = $True
    $button1.Text = "确认"
    $button1.Location = New-Object System.Drawing.Point(27,472)
    $button1.Size = New-Object System.Drawing.Size(75,23)
    $button1.DataBindings.DefaultDataSourceUpdateMode = 0
    $button1.add_Click($handler_saveButton_Click)

    $form.Controls.Add($button1)

    $button2.TabIndex = 5
    $button2.Name = "cancelButton"
    $button2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $button2.UseVisualStyleBackColor = $True
    $button2.Text = "取消"
    $button2.Location = New-Object System.Drawing.Point(129,472)
    $button2.Size = New-Object System.Drawing.Size(75,23)
    $button2.DataBindings.DefaultDataSourceUpdateMode = 0
    $button2.add_Click($handler_cancelButton_Click)

    $form.Controls.Add($button2)

    $label.Location = New-Object System.Drawing.Point(13,5)
    $label.Size = New-Object System.Drawing.Size(400,14)
    $Label.Font = 'Microsoft Sans Serif,8'
    $label.Text = '勾选您希望卸载的应用，取消勾选您希望保留的应用'

    $form.Controls.Add($label)

    $loadingLabel.Location = New-Object System.Drawing.Point(16,46)
    $loadingLabel.Size = New-Object System.Drawing.Size(300,418)
    $loadingLabel.Text = '正在加载应用...'
    $loadingLabel.BackColor = "White"
    $loadingLabel.Visible = $false

    $form.Controls.Add($loadingLabel)

    $onlyInstalledCheckBox.TabIndex = 6
    $onlyInstalledCheckBox.Location = New-Object System.Drawing.Point(230,474)
    $onlyInstalledCheckBox.Size = New-Object System.Drawing.Size(150,20)
    $onlyInstalledCheckBox.Text = '仅显示已安装应用'
    $onlyInstalledCheckBox.add_CheckedChanged($load_Apps)

    $form.Controls.Add($onlyInstalledCheckBox)

    $checkUncheckCheckBox.TabIndex = 7
    $checkUncheckCheckBox.Location = New-Object System.Drawing.Point(16,22)
    $checkUncheckCheckBox.Size = New-Object System.Drawing.Size(150,20)
    $checkUncheckCheckBox.Text = '全选/取消全选'
    $checkUncheckCheckBox.add_CheckedChanged($check_All)

    $form.Controls.Add($checkUncheckCheckBox)

    $selectionBox.FormattingEnabled = $True
    $selectionBox.DataBindings.DefaultDataSourceUpdateMode = 0
    $selectionBox.Name = "selectionBox"
    $selectionBox.Location = New-Object System.Drawing.Point(13,43)
    $selectionBox.Size = New-Object System.Drawing.Size(374,424)
    $selectionBox.TabIndex = 3
    $selectionBox.add_SelectedIndexChanged($selectionBox_SelectedIndexChanged)
    $selectionBox.add_Click($selectionBox_MouseDown)

    $form.Controls.Add($selectionBox)

    $initialFormWindowState = $form.WindowState
    $form.add_Load($load_Apps)
    $form.Add_Shown({$form.Activate(); $selectionBox.Focus()})

    return $form.ShowDialog()
}


# 从指定文件读取应用列表
function ReadAppslistFromFile {
    param (
        $appsFilePath
    )

    $appsList = @()

    Foreach ($app in (Get-Content -Path $appsFilePath | Where-Object { $_ -notmatch '^#.*' -and $_ -notmatch '^\s*$' } )) { 
        if (-not ($app.IndexOf('#') -eq -1)) {
            $app = $app.Substring(0, $app.IndexOf('#'))
        }

        $app = $app.Trim()
        $appString = $app.Trim('*')
        $appsList += $appString
    }

    return $appsList
}


# 卸载指定应用
function RemoveApps {
    param (
        $appslist
    )

    Foreach ($app in $appsList) { 
        Write-Output "正在尝试卸载 $app..."

        if (($app -eq "Microsoft.OneDrive") -or ($app -eq "Microsoft.Edge")) {
            if ($global:wingetInstalled -eq $false) {
                Write-Host "错误：WinGet 未安装或版本过旧，无法移除 $app" -ForegroundColor Red
            }
            else {
                Strip-Progress -ScriptBlock { winget uninstall --accept-source-agreements --disable-interactivity --id $app } | Tee-Object -Variable wingetOutput 

                If (($app -eq "Microsoft.Edge") -and (Select-String -InputObject $wingetOutput -Pattern "Uninstall failed with exit code")) {
                    Write-Host "无法通过 Winget 卸载 Microsoft Edge" -ForegroundColor Red
                    Write-Output ""

                    if ($( Read-Host -Prompt "是否要强制卸载Edge？(不推荐) (y/n)" ) -eq 'y') {
                        Write-Output ""
                        ForceRemoveEdge
                    }
                }
            }
        }
        else {
            $app = '*' + $app + '*'

            if ($WinVersion -ge 22000){
                try {
                    Get-AppxPackage -Name $app -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction Continue
                }
                catch {
                    if($DebugPreference -ne "SilentlyContinue") {
                        Write-Host "无法为所有用户移除 $app" -ForegroundColor Yellow
                    }
                }
            }
            else {
                try {
                    Get-AppxPackage -Name $app | Remove-AppxPackage -ErrorAction SilentlyContinue
                }
                catch {}
                
                try {
                    Get-AppxPackage -Name $app -PackageTypeFilter Main, Bundle, Resource -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
                }
                catch {}
            }

            try {
                Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -like $app } | ForEach-Object { Remove-ProvisionedAppxPackage -Online -AllUsers -PackageName $_.PackageName }
            }
            catch {
                Write-Host "无法从系统镜像移除 $app" -ForegroundColor Yellow
            }
        }
    }
            
    Write-Output ""
}


# 强制移除 Microsoft Edge
function ForceRemoveEdge {
    Write-Output "> 正在强制卸载 Microsoft Edge..."

    $regView = [Microsoft.Win32.RegistryView]::Registry32
    $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $regView)
    $hklm.CreateSubKey('SOFTWARE\Microsoft\EdgeUpdateDev').SetValue('AllowUninstall', '')

    $edgeStub = "$env:SystemRoot\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe"
    New-Item $edgeStub -ItemType Directory | Out-Null
    New-Item "$edgeStub\MicrosoftEdge.exe" | Out-Null

    $uninstallRegKey = $hklm.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge')
    if ($null -ne $uninstallRegKey) {
        Write-Output "正在运行卸载程序..."
        $uninstallString = $uninstallRegKey.GetValue('UninstallString') + ' --force-uninstall'
        Start-Process cmd.exe "/c $uninstallString" -WindowStyle Hidden -Wait

        Write-Output "正在移除残留文件..."
        $edgePaths = @(
            "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Tombstones\Microsoft Edge.lnk",
            "$env:PUBLIC\Desktop\Microsoft Edge.lnk",
            "$env:USERPROFILE\Desktop\Microsoft Edge.lnk",
            "$edgeStub"
        )

        foreach ($path in $edgePaths){
            if (Test-Path -Path $path) {
                Remove-Item -Path $path -Force -Recurse -ErrorAction SilentlyContinue
                Write-Host "  Removed $path" -ForegroundColor DarkGray
            }
        }

        Write-Output "正在清理注册表..."
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "MicrosoftEdgeAutoLaunch_A9F6DCE4ABADF4F51CF45CD7129E3C6C" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "Microsoft Edge Update" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" /v "MicrosoftEdgeAutoLaunch_A9F6DCE4ABADF4F51CF45CD7129E3C6C" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" /v "Microsoft Edge Update" /f *>$null

        Write-Output "Microsoft Edge 已卸载"
    }
    else {
        Write-Output ""
        Write-Host "错误：无法找到 Microsoft Edge 卸载程序" -ForegroundColor Red
    }
    
    Write-Output ""
}


# 执行命令并清理进度显示
function Strip-Progress {
    param(
        [ScriptBlock]$ScriptBlock
    )

    $progressPattern = 'Γû[Æê]|^\s+[-\\|/]\s+$'
    $sizePattern = '(\d+(\.\d{1,2})?)\s+(B|KB|MB|GB|TB|PB) /\s+(\d+(\.\d{1,2})?)\s+(B|KB|MB|GB|TB|PB)'

    & $ScriptBlock 2>&1 | ForEach-Object {
        if ($_ -is [System.Management.Automation.ErrorRecord]) {
            "错误：$($_.Exception.Message)"
        } else {
            $line = $_ -replace $progressPattern, '' -replace $sizePattern, ''
            if (-not ([string]::IsNullOrWhiteSpace($line)) -and -not ($line.StartsWith('  '))) {
                $line
            }
        }
    }
}


# Import & execute regfile
function RegImport {
    param (
        $message,
        $path
    )

    Write-Output $message

    if ($global:Params.ContainsKey("Sysprep")) {
        $defaultUserPath = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), '\Default\NTUSER.DAT'
        
        reg load "HKU\Default" $defaultUserPath | Out-Null
        reg import "$PSScriptRoot\Regfiles\Sysprep\$path"
        reg unload "HKU\Default" | Out-Null
    }
    elseif ($global:Params.ContainsKey("User")) {
        $userPath = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), "\$($global:Params.Item("User"))\NTUSER.DAT"
        
        reg load "HKU\Default" $userPath | Out-Null
        reg import "$PSScriptRoot\Regfiles\Sysprep\$path"
        reg unload "HKU\Default" | Out-Null
        
    }
    else {
        reg import "$PSScriptRoot\Regfiles\$path"  
    }

    Write-Output ""
}


# Restart the Windows Explorer process
function RestartExplorer {
    if ($global:Params.ContainsKey("Sysprep") -or $global:Params.ContainsKey("User")) {
        return
    }

    Write-Output "> 正在重启 Windows 资源管理器进程以应用更改... (可能会出现闪烁)"

    if ($global:Params.ContainsKey("DisableMouseAcceleration")) {
        Write-Host "警告：指针精确度增强设置的更改需要重启后才能生效" -ForegroundColor Yellow
    }

    if ($global:Params.ContainsKey("DisableStickyKeys")) {
        Write-Host "警告：粘滞键设置的更改需要重启后才能生效" -ForegroundColor Yellow
    }

    # Only restart if the powershell process matches the OS architecture.
    # Restarting explorer from a 32bit Powershell window will fail on a 64bit OS
    if ([Environment]::Is64BitProcess -eq [Environment]::Is64BitOperatingSystem) {
        Stop-Process -processName: Explorer -Force
    }
    else {
        Write-Warning "无法自动重启 Windows 资源管理器进程，请手动重启计算机以应用所有更改。"
    }
}


# Replace the startmenu for all users, when using the default startmenuTemplate this clears all pinned apps
# Credit: https://lazyadmin.nl/win-11/customize-windows-11-start-menu-layout/
function ReplaceStartMenuForAllUsers {
    param (
        $startMenuTemplate = "$PSScriptRoot/Start/start2.bin"
    )

    Write-Output "> 正在为所有用户移除开始菜单中的所有固定应用..."

    # Check if template bin file exists, return early if it doesn't
    if (-not (Test-Path $startMenuTemplate)) {
        Write-Host "错误：无法清除开始菜单，脚本目录中缺少 start2.bin 文件" -ForegroundColor Red
        Write-Output ""
        return
    }

    # Get path to start menu file for all users
    $userPathString = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), "\*\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState"
    $usersStartMenuPaths = get-childitem -path $userPathString

    # Go through all users and replace the start menu file
    ForEach ($startMenuPath in $usersStartMenuPaths) {
        ReplaceStartMenu "$($startMenuPath.Fullname)\start2.bin" $startMenuTemplate
    }

    # Also replace the start menu file for the default user profile
    $defaultStartMenuPath = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), '\Default\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState'

    # Create folder if it doesn't exist
    if (-not(Test-Path $defaultStartMenuPath)) {
        new-item $defaultStartMenuPath -ItemType Directory -Force | Out-Null
        Write-Output "已为默认用户配置文件创建 LocalState 文件夹"
    }

    # Copy template to default profile
    Copy-Item -Path $startMenuTemplate -Destination $defaultStartMenuPath -Force
    Write-Output "已替换默认用户配置文件的开始菜单"
    Write-Output ""
}


# Replace the startmenu for all users, when using the default startmenuTemplate this clears all pinned apps
# Credit: https://lazyadmin.nl/win-11/customize-windows-11-start-menu-layout/
function ReplaceStartMenu {
    param (
        $startMenuBinFile = "$env:LOCALAPPDATA\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start2.bin",
        $startMenuTemplate = "$PSScriptRoot/Start/start2.bin"
    )

    # Change path to correct user if a user was specified
    if ($global:Params.ContainsKey("User")) {
        $startMenuBinFile = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), "\$(GetUserName)\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start2.bin"
    }

    # Check if template bin file exists, return early if it doesn't
    if (-not (Test-Path $startMenuTemplate)) {
        Write-Host "错误：无法清除开始菜单，脚本目录中缺少 start2.bin 文件" -ForegroundColor Red
        return
    }

    # Check if bin file exists, return early if it doesn't
    if (-not (Test-Path $startMenuBinFile)) {
        Write-Host "错误：无法为用户 $(GetUserName) 清除开始菜单，找不到 start2.bin 文件" -ForegroundColor Red
        return
    }

    $backupBinFile = $startMenuBinFile + ".bak"

    # Backup current start menu file
    Move-Item -Path $startMenuBinFile -Destination $backupBinFile -Force

    # Copy template file
    Copy-Item -Path $startMenuTemplate -Destination $startMenuBinFile -Force

    Write-Output "已替换用户 $(GetUserName) 的开始菜单"
}


# Add parameter to script and write to file
function AddParameter {
    param (
        $parameterName,
        $message
    )

    # Add key if it doesn't already exist
    if (-not $global:Params.ContainsKey($parameterName)) {
        $global:Params.Add($parameterName, $true)
    }

    # Create or clear file that stores last used settings
    if (!(Test-Path "$PSScriptRoot/SavedSettings")) {
        $null = New-Item "$PSScriptRoot/SavedSettings"
    } 
    elseif ($global:FirstSelection) {
        $null = Clear-Content "$PSScriptRoot/SavedSettings"
    }
    
    $global:FirstSelection = $false

    # Create entry and add it to the file
    $entry = "$parameterName#- $message"
    Add-Content -Path "$PSScriptRoot/SavedSettings" -Value $entry
}


function PrintHeader {
    param (
        $title
    )

    $fullTitle = " Win11Debloat 脚本 - $title"

    if ($global:Params.ContainsKey("Sysprep")) {
        $fullTitle = "$fullTitle (Sysprep 模式)"
    }
    else {
        $fullTitle = "$fullTitle (用户：$(GetUserName))"
    }

    Clear-Host
    Write-Output "-------------------------------------------------------------------------------------------"
    Write-Output $fullTitle
    Write-Output "-------------------------------------------------------------------------------------------"
}


function PrintFromFile {
    param (
        $path
    )

    Clear-Host

    # Get & print script menu from file
    Foreach ($line in (Get-Content -Path $path )) {   
        Write-Output $line
    }
}


function AwaitKeyToExit {
    # Suppress prompt if Silent parameter was passed
    if (-not $Silent) {
        Write-Output ""
        Write-Output "按任意键退出..."
        $null = [System.Console]::ReadKey()
    }
}


function GetUserName {
    if ($global:Params.ContainsKey("User")) { 
        return $global:Params.Item("User") 
    } 
    else { 
        return $env:USERNAME 
    }
}


function DisplayCustomModeOptions {
    # Get current Windows build version to compare against features
    $WinVersion = Get-ItemPropertyValue 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' CurrentBuild
            
    PrintHeader '自定义模式'

    # Show options for removing apps, only continue on valid input
    Do {
        Write-Host "选项：" -ForegroundColor Yellow
        Write-Host " (n) 不移除任何应用" -ForegroundColor Yellow
        Write-Host " (1) 仅移除 Appslist.txt 中的默认预装应用" -ForegroundColor Yellow
        Write-Host " (2) 移除默认预装应用，以及邮件日历应用、开发者工具和游戏相关应用"  -ForegroundColor Yellow
        Write-Host " (3) 自定义选择要移除的应用" -ForegroundColor Yellow
        $RemoveAppsInput = Read-Host "是否移除预装应用？(n/1/2/3)" 

        # Show app selection form if user entered option 3
        if ($RemoveAppsInput -eq '3') {
            $result = ShowAppSelectionForm

            if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
                # User cancelled or closed app selection, show error and change RemoveAppsInput so the menu will be shown again
                Write-Output ""
                Write-Host "已取消应用选择，请重试" -ForegroundColor Red

                $RemoveAppsInput = 'c'
            }
            
            Write-Output ""
        }
    }
    while ($RemoveAppsInput -ne 'n' -and $RemoveAppsInput -ne '0' -and $RemoveAppsInput -ne '1' -and $RemoveAppsInput -ne '2' -and $RemoveAppsInput -ne '3') 

    # Select correct option based on user input
    switch ($RemoveAppsInput) {
        '1' {
            AddParameter 'RemoveApps' '移除默认预装应用'
        }
        '2' {
            AddParameter 'RemoveApps' '移除默认预装应用'
            AddParameter 'RemoveCommApps' '移除邮件、日历和联系人应用'
            AddParameter 'RemoveW11Outlook' '移除新版 Outlook for Windows 应用'
            AddParameter 'RemoveDevApps' '移除开发者相关应用'
            AddParameter 'RemoveGamingApps' '移除 Xbox 应用和 Xbox 游戏栏'
            AddParameter 'DisableDVR' '禁用 Xbox 游戏/屏幕录制'
        }
        '3' {
            Write-Output "您已选择移除 $($global:SelectedApps.Count) 个应用"

            AddParameter 'RemoveAppsCustom' "移除 $($global:SelectedApps.Count) 个应用："

            Write-Output ""

            if ($( Read-Host -Prompt "是否禁用 Xbox 游戏/屏幕录制？同时禁用游戏弹窗 (y/n)" ) -eq 'y') {
                AddParameter 'DisableDVR' '禁用 Xbox 游戏/屏幕录制'
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "是否禁用遥测、诊断数据、活动历史、应用启动追踪和定向广告？(y/n)" ) -eq 'y') {
        AddParameter 'DisableTelemetry' '禁用遥测、诊断数据、活动历史、应用启动追踪和定向广告'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "是否禁用开始菜单、设置、通知、文件资源管理器中的提示、技巧、建议和广告？(y/n)" ) -eq 'y') {
        AddParameter 'DisableSuggestions' '禁用开始菜单、设置、通知和文件资源管理器中的提示、技巧、建议和广告'
        AddParameter 'DisableDesktopSpotlight' '禁用 Windows Spotlight 桌面背景选项'
        AddParameter 'DisableLockscreenTips' '禁用锁屏界面的提示和技巧'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "是否禁用并移除 Windows 搜索中的 Bing 网页搜索、Bing AI 和 Cortana？(y/n)" ) -eq 'y') {
        AddParameter 'DisableBing' '禁用并移除 Windows 搜索中的 Bing 网页搜索、Bing AI 和 Cortana'
    }

    # Only show this option for Windows 11 users running build 22621 or later
    if ($WinVersion -ge 22621){
        Write-Output ""

        if ($( Read-Host -Prompt "是否为所有用户禁用并移除 Microsoft Copilot 和 Windows Recall 快照？(y/n)" ) -eq 'y') {
            AddParameter 'DisableCopilot' '禁用并移除 Microsoft Copilot'
            AddParameter 'DisableRecall' '禁用并移除 Windows Recall 快照'
        }
    }

    # Only show this option for Windows 11 users running build 22000 or later
    if ($WinVersion -ge 22000){
        Write-Output ""

        if ($( Read-Host -Prompt "是否恢复 Windows 10 样式的右键菜单？(y/n)" ) -eq 'y') {
            AddParameter 'RevertContextMenu' '恢复 Windows 10 样式的右键菜单'
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "是否关闭指针精确度增强（鼠标加速）？(y/n)" ) -eq 'y') {
        AddParameter 'DisableMouseAcceleration' '关闭指针精确度增强（鼠标加速）'
    }

    # Only show this option for Windows 11 users running build 26100 or later
    if ($WinVersion -ge 26100){
        Write-Output ""

        if ($( Read-Host -Prompt "是否禁用粘滞键快捷键？(y/n)" ) -eq 'y') {
            AddParameter 'DisableStickyKeys' '禁用粘滞键快捷键'
        }
    }

    # Only show option for disabling context menu items for Windows 10 users or if the user opted to restore the Windows 10 context menu
    if ((get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'") -or $global:Params.ContainsKey('RevertContextMenu')){
        Write-Output ""

        if ($( Read-Host -Prompt "是否需要禁用某些右键菜单选项？(y/n)" ) -eq 'y') {
            Write-Output ""

            if ($( Read-Host -Prompt "   隐藏右键菜单中的 '包含到库中' 选项？(y/n)" ) -eq 'y') {
                AddParameter 'HideIncludeInLibrary' "隐藏右键菜单中的 '包含到库中' 选项"
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   隐藏右键菜单中的 '授予访问权限' 选项？(y/n)" ) -eq 'y') {
                AddParameter 'HideGiveAccessTo' "隐藏右键菜单中的 '授予访问权限' 选项"
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   隐藏右键菜单中的 '共享' 选项？(y/n)" ) -eq 'y') {
                AddParameter 'HideShare' "隐藏右键菜单中的 '共享' 选项"
            }
        }
    }

    # Only show this option for Windows 11 users running build 22621 or later
    if ($WinVersion -ge 22621){
        Write-Output ""

        if ($( Read-Host -Prompt "是否需要修改开始菜单？(y/n)" ) -eq 'y') {
            Write-Output ""

            if ($global:Params.ContainsKey("Sysprep")) {
                if ($( Read-Host -Prompt "是否为所有现有用户和新用户移除开始菜单中的所有固定应用？(y/n)" ) -eq 'y') {
                    AddParameter 'ClearStartAllUsers' '为所有用户移除开始菜单中的固定应用'
                }
            }
            else {
                Do {
                    Write-Host "   选项：" -ForegroundColor Yellow
                    Write-Host "    (n) 不移除任何固定应用" -ForegroundColor Yellow
                    Write-Host "    (1) 仅移除当前用户 ($(GetUserName)) 的固定应用" -ForegroundColor Yellow
                    Write-Host "    (2) 为所有现有用户和新用户移除固定应用"  -ForegroundColor Yellow
                    $ClearStartInput = Read-Host "   移除开始菜单中的固定应用？(n/1/2)" 
                }
                while ($ClearStartInput -ne 'n' -and $ClearStartInput -ne '0' -and $ClearStartInput -ne '1' -and $ClearStartInput -ne '2') 

                # Select correct option based on user input
                switch ($ClearStartInput) {
                    '1' {
                        AddParameter 'ClearStart' "移除当前用户 ($(GetUserName)) 的固定应用"
                    }
                    '2' {
                        AddParameter 'ClearStartAllUsers' "为所有用户移除开始菜单中的固定应用"
                    }
                }
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   是否为所有用户禁用并隐藏开始菜单的推荐区域？(y/n)" ) -eq 'y') {
                AddParameter 'DisableStartRecommended' '禁用并隐藏开始菜单的推荐区域'
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "是否需要修改任务栏及相关服务？(y/n)" ) -eq 'y') {
        # Only show these specific options for Windows 11 users running build 22000 or later
        if ($WinVersion -ge 22000){
            Write-Output ""

            if ($( Read-Host -Prompt "   是否将任务栏图标左对齐？(y/n)" ) -eq 'y') {
                AddParameter 'TaskbarAlignLeft' '任务栏图标左对齐'
            }

            # Show options for search icon on taskbar, only continue on valid input
            Do {
                Write-Output ""
                Write-Host "   选项：" -ForegroundColor Yellow
                Write-Host "    (n) 保持当前设置" -ForegroundColor Yellow
                Write-Host "    (1) 隐藏任务栏搜索图标" -ForegroundColor Yellow
                Write-Host "    (2) 显示任务栏搜索图标" -ForegroundColor Yellow
                Write-Host "    (3) 显示带标签的搜索图标" -ForegroundColor Yellow
                Write-Host "    (4) 显示搜索框" -ForegroundColor Yellow
                $TbSearchInput = Read-Host "   是否修改任务栏搜索图标？(n/1/2/3/4)" 
            }
            while ($TbSearchInput -ne 'n' -and $TbSearchInput -ne '0' -and $TbSearchInput -ne '1' -and $TbSearchInput -ne '2' -and $TbSearchInput -ne '3' -and $TbSearchInput -ne '4') 

            # Select correct taskbar search option based on user input
            switch ($TbSearchInput) {
                '1' {
                    AddParameter 'HideSearchTb' '隐藏任务栏搜索图标'
                }
                '2' {
                    AddParameter 'ShowSearchIconTb' '显示任务栏搜索图标'
                }
                '3' {
                    AddParameter 'ShowSearchLabelTb' '显示带标签的搜索图标'
                }
                '4' {
                    AddParameter 'ShowSearchBoxTb' '显示搜索框'
                }
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   是否隐藏任务视图按钮？(y/n)" ) -eq 'y') {
                AddParameter 'HideTaskview' '隐藏任务视图按钮'
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   是否禁用小组件服务并隐藏任务栏图标？(y/n)" ) -eq 'y') {
            AddParameter 'DisableWidgets' '禁用小组件服务并隐藏任务栏图标'
        }

        # Only show this options for Windows users running build 22621 or earlier
        if ($WinVersion -le 22621){
            Write-Output ""

            if ($( Read-Host -Prompt "   是否隐藏任务栏聊天（立即开会）图标？(y/n)" ) -eq 'y') {
                AddParameter 'HideChat' '隐藏任务栏聊天图标'
            }
        }
        
        # Only show this options for Windows users running build 22631 or later
        if ($WinVersion -ge 22631){
            Write-Output ""

            if ($( Read-Host -Prompt "   是否在任务栏右键菜单启用 '结束任务' 选项？(y/n)" ) -eq 'y') {
                AddParameter 'EnableEndTask' "启用任务栏右键菜单的 '结束任务' 选项"
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "是否要修改文件资源管理器设置？(y/n)" ) -eq 'y') {
        # 修改文件资源管理器默认位置的选项
        Do {
            Write-Output ""
            Write-Host "   选项：" -ForegroundColor Yellow
            Write-Host "    (n) 不修改" -ForegroundColor Yellow
            Write-Host "    (1) 文件资源管理器默认打开'主页'" -ForegroundColor Yellow
            Write-Host "    (2) 文件资源管理器默认打开'此电脑'" -ForegroundColor Yellow
            Write-Host "    (3) 文件资源管理器默认打开'下载'" -ForegroundColor Yellow
            Write-Host "    (4) 文件资源管理器默认打开'OneDrive'" -ForegroundColor Yellow
            $ExplSearchInput = Read-Host "   请选择文件资源管理器的默认打开位置 (n/1/2/3/4)" 
        }
        while ($ExplSearchInput -ne 'n' -and $ExplSearchInput -ne '0' -and $ExplSearchInput -ne '1' -and $ExplSearchInput -ne '2' -and $ExplSearchInput -ne '3' -and $ExplSearchInput -ne '4') 

        # 根据用户输入选择设置
        switch ($ExplSearchInput) {
            '1' {
                AddParameter 'ExplorerToHome' "将文件资源管理器默认打开位置设置为'主页'"
            }
            '2' {
                AddParameter 'ExplorerToThisPC' "将文件资源管理器默认打开位置设置为'此电脑'"
            }
            '3' {
                AddParameter 'ExplorerToDownloads' "将文件资源管理器默认打开位置设置为'下载'"
            }
            '4' {
                AddParameter 'ExplorerToOneDrive' "将文件资源管理器默认打开位置设置为'OneDrive'"
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   是否显示隐藏的文件、文件夹和驱动器？(y/n)" ) -eq 'y') {
            AddParameter 'ShowHiddenFolders' '显示隐藏的文件、文件夹和驱动器'
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   是否显示已知文件类型的扩展名？(y/n)" ) -eq 'y') {
            AddParameter 'ShowKnownFileExt' '显示已知文件类型的扩展名'
        }

        # 仅适用于Windows 11 22000或更高版本
        if ($WinVersion -ge 22000){
            Write-Output ""

            if ($( Read-Host -Prompt "   是否从文件资源管理器侧边栏隐藏'主页'部分？(y/n)" ) -eq 'y') {
                AddParameter 'HideHome' '从文件资源管理器侧边栏隐藏主页部分'
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   是否从文件资源管理器侧边栏隐藏'图库'部分？(y/n)" ) -eq 'y') {
                AddParameter 'HideGallery' '从文件资源管理器侧边栏隐藏图库部分'
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   是否从文件资源管理器侧边栏隐藏重复的可移动驱动器条目？（仅在'此电脑'下显示）(y/n)" ) -eq 'y') {
            AddParameter 'HideDupliDrive' '从文件资源管理器侧边栏隐藏重复的可移动驱动器条目'
        }

        # 仅Windows 10显示特定文件夹隐藏选项
        if (get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'"){
            Write-Output ""

            if ($( Read-Host -Prompt "是否要从文件资源管理器侧边栏隐藏某些文件夹？(y/n)" ) -eq 'y') {
                Write-Output ""

                if ($( Read-Host -Prompt "   是否隐藏OneDrive文件夹？(y/n)" ) -eq 'y') {
                    AddParameter 'HideOnedrive' '在文件资源管理器侧边栏隐藏OneDrive文件夹'
                }

                Write-Output ""
                
                if ($( Read-Host -Prompt "   是否隐藏3D对象文件夹？(y/n)" ) -eq 'y') {
                    AddParameter 'Hide3dObjects' "在文件资源管理器的'此电脑'中隐藏3D对象文件夹" 
                }
                
                Write-Output ""

                if ($( Read-Host -Prompt "   是否隐藏音乐文件夹？(y/n)" ) -eq 'y') {
                    AddParameter 'HideMusic' "在文件资源管理器的'此电脑'中隐藏音乐文件夹"
                }
            }
        }
    }

    # 静默模式下不显示确认提示
    if (-not $Silent) {
        Write-Output ""
        Write-Output ""
        Write-Output ""
        Write-Output "按回车确认选择并执行脚本，或按CTRL+C退出..."
        Read-Host | Out-Null
    }

    PrintHeader '自定义模式'
}

##################################################################################################################
#                                                                                                                #
#                                                  脚本开始执行                                                  #
#                                                                                                                #
##################################################################################################################

# 检查是否安装winget（v1.4+）
if ((Get-AppxPackage -Name "*Microsoft.DesktopAppInstaller*") -and ([int](((winget -v) -replace 'v','').split('.')[0..1] -join '') -gt 14)) {
    $global:wingetInstalled = $true
}
else {
    $global:wingetInstalled = $false

    # 非静默模式下显示警告
    if (-not $Silent) {
        Write-Warning "未检测到winget或版本过低（需要v1.4+），可能会影响应用卸载功能"
        Write-Output ""
        Write-Output "按任意键继续..."
        $null = [System.Console]::ReadKey()
    }
}

# 获取当前Windows版本
$WinVersion = Get-ItemPropertyValue 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' CurrentBuild

$global:Params = $PSBoundParameters
$global:FirstSelection = $true
$SPParams = 'WhatIf', 'Confirm', 'Verbose', 'Silent', 'Sysprep', 'Debug', 'User'
$SPParamCount = 0

# 统计SPParams参数数量
foreach ($Param in $SPParams) {
    if ($global:Params.ContainsKey($Param)) {
        $SPParamCount++
    }
}

# 隐藏进度条（静默模式）
if (-not ($global:Params.ContainsKey("Verbose"))) {
    $ProgressPreference = 'SilentlyContinue'
}
else {
    Read-Host "已启用详细模式，按回车继续" 
    $ProgressPreference = 'Continue'
}

# Sysprep模式检查
if ($global:Params.ContainsKey("Sysprep")) {
    $defaultUserPath = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), '\Default\NTUSER.DAT'

    if (-not (Test-Path "$defaultUserPath")) {
        Write-Host "错误：无法找到默认用户配置文件：'$defaultUserPath'" -ForegroundColor Red
        AwaitKeyToExit
        Exit
    }
    if ($WinVersion -lt 22000) {
        Write-Host "错误：Sysprep模式不支持Windows 10" -ForegroundColor Red
        AwaitKeyToExit
        Exit
    }
}

# 用户模式检查
if ($global:Params.ContainsKey("User")) {
    $userPath = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), "\$($global:Params.Item("User"))\NTUSER.DAT"

    if (-not (Test-Path "$userPath")) {
        Write-Host "错误：无法找到用户 $($global:Params.Item("User")) 的配置文件：'$userPath'" -ForegroundColor Red
        AwaitKeyToExit
        Exit
    }
}

# 清理空配置文件
if ((Test-Path "$PSScriptRoot/SavedSettings") -and ([String]::IsNullOrWhiteSpace((Get-content "$PSScriptRoot/SavedSettings")))) {
    Remove-Item -Path "$PSScriptRoot/SavedSettings" -recurse
}

# 应用列表生成模式
if ($RunAppConfigurator -or $RunAppsListGenerator) {
    PrintHeader "自定义应用列表生成器"

    $result = ShowAppSelectionForm

    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "应用选择未保存" -ForegroundColor Red
    }
    else {
        Write-Output "您的应用选择已保存至以下路径的'CustomAppsList'文件："
        Write-Host "$PSScriptRoot" -ForegroundColor Yellow
    }

    AwaitKeyToExit
    Exit
}

# 脚本执行模式选择
if ((-not $global:Params.Count) -or $RunDefaults -or $RunWin11Defaults -or $RunSavedSettings -or ($SPParamCount -eq $global:Params.Count)) {
    if ($RunDefaults -or $RunWin11Defaults) {
        $Mode = '1'
    }
    elseif ($RunSavedSettings) {
        if(-not (Test-Path "$PSScriptRoot/SavedSettings")) {
            PrintHeader '自定义模式'
            Write-Host "错误：未找到已保存的设置" -ForegroundColor Red
            AwaitKeyToExit
            Exit
        }

        $Mode = '4'
    }
    else {
        Do { 
            $ModeSelectionMessage = "请选择一个选项 (1/2/3/0)" 

            PrintHeader '主菜单'

            Write-Output "(1) 默认模式：应用默认设置"
            Write-Output "(2) 自定义模式：按需修改脚本设置"
            Write-Output "(3) 应用移除模式：仅选择并移除应用"

            if (Test-Path "$PSScriptRoot/SavedSettings") {
                Write-Output "(4) 应用上次保存的自定义设置"
                $ModeSelectionMessage = "请选择一个选项 (1/2/3/4/0)" 
            }

            Write-Output ""
            Write-Output "(0) 显示更多信息"
            Write-Output ""
            Write-Output ""

            $Mode = Read-Host $ModeSelectionMessage

            if ($Mode -eq '0') {
                PrintFromFile "$PSScriptRoot/Assets/Menus/Info"

                Write-Output ""
                Write-Output "按任意键返回..."
                $null = [System.Console]::ReadKey()
            }
            elseif (($Mode -eq '4')-and -not (Test-Path "$PSScriptRoot/SavedSettings")) {
                $Mode = $null
            }
        }
        while ($Mode -ne '1' -and $Mode -ne '2' -and $Mode -ne '3' -and $Mode -ne '4') 
    }

    # 根据模式添加参数
    switch ($Mode) {
        '1' { 
            if (-not $Silent) {
                PrintFromFile "$PSScriptRoot/Assets/Menus/DefaultSettings"

                Write-Output ""
                Write-Output "按回车开始执行，或按CTRL+C退出..."
                Read-Host | Out-Null
            }

            $DefaultParameterNames = 'RemoveApps','DisableTelemetry','DisableBing','DisableLockscreenTips','DisableSuggestions','ShowKnownFileExt','DisableWidgets','HideChat','DisableCopilot'

            PrintHeader '默认模式'

            foreach ($ParameterName in $DefaultParameterNames) {
                if (-not $global:Params.ContainsKey($ParameterName)){
                    $global:Params.Add($ParameterName, $true)
                }
            }

            if ((get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'") -and (-not $global:Params.ContainsKey('Hide3dObjects'))) {
                $global:Params.Add('Hide3dObjects', $Hide3dObjects)
            }
        }

        '2' { 
            DisplayCustomModeOptions
        }

        '3' {
            PrintHeader "应用移除"

            $result = ShowAppSelectionForm

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Output "已选择 $($global:SelectedApps.Count) 个待移除应用"
                AddParameter 'RemoveAppsCustom' "移除 $($global:SelectedApps.Count) 个应用："

                if (-not $Silent) {
                    Write-Output ""
                    Write-Output "按回车开始移除应用，或按CTRL+C退出..."
                    Read-Host | Out-Null
                    PrintHeader "应用移除"
                }
            }
            else {
                Write-Host "已取消操作，未移除任何应用" -ForegroundColor Red
                Write-Output ""
            }
        }

        '4' {
            PrintHeader '自定义模式'
            Write-Output "即将应用以下设置："

            Foreach ($line in (Get-Content -Path "$PSScriptRoot/SavedSettings" )) { 
                $line = $line.Trim()
            
                if (-not ($line.IndexOf('#') -eq -1)) {
                    $parameterName = $line.Substring(0, $line.IndexOf('#'))

                    if ($parameterName -eq "RemoveAppsCustom") {
                        if (-not (Test-Path "$PSScriptRoot/CustomAppsList")) {
                            continue
                        }
                        
                        $appsList = ReadAppslistFromFile "$PSScriptRoot/CustomAppsList"
                        Write-Output "- 移除 $($appsList.Count) 个应用："
                        Write-Host $appsList -ForegroundColor DarkGray
                    }
                    else {
                        Write-Output $line.Substring(($line.IndexOf('#') + 1), ($line.Length - $line.IndexOf('#') - 1))
                    }

                    if (-not $global:Params.ContainsKey($parameterName)){
                        $global:Params.Add($parameterName, $true)
                    }
                }
            }

            if (-not $Silent) {
                Write-Output ""
                Write-Output ""
                Write-Output "按回车开始执行，或按CTRL+C退出..."
                Read-Host | Out-Null
            }

            PrintHeader '自定义模式'
        }
    }
}
else {
    PrintHeader '自定义模式'
}

# 以下内容是用户上传的文档解析结果：
# 如果 SPParams 中的键数量与 Params 中的键数量相同，则表示用户未选择或添加任何修改
# 脚本可以直接退出而不做任何更改
if ($SPParamCount -eq $global:Params.Keys.Count) {
    Write-Output "脚本执行完毕，未进行任何修改。"

    AwaitKeyToExit
}
else {
    # 执行所有选择/提供的参数
    switch ($global:Params.Keys) {
        'RemoveApps' {
            $appsList = ReadAppslistFromFile "$PSScriptRoot/Appslist.txt" 
            Write-Output "> 正在移除默认选择的 $($appsList.Count) 个应用..."
            RemoveApps $appsList
            continue
        }
        'RemoveAppsCustom' {
            if (-not (Test-Path "$PSScriptRoot/CustomAppsList")) {
                Write-Host "> 错误：无法从文件加载自定义应用列表，未移除任何应用" -ForegroundColor Red
                Write-Output ""
                continue
            }
            
            $appsList = ReadAppslistFromFile "$PSScriptRoot/CustomAppsList"
            Write-Output "> 正在移除 $($appsList.Count) 个应用..."
            RemoveApps $appsList
            continue
        }
        'RemoveCommApps' {
            Write-Output "> 正在移除邮件、日历和人脉应用..."
            
            $appsList = 'Microsoft.windowscommunicationsapps', 'Microsoft.People'
            RemoveApps $appsList
            continue
        }
        'RemoveW11Outlook' {
            $appsList = 'Microsoft.OutlookForWindows'
            Write-Output "> 正在移除新版 Windows Outlook 应用..."
            RemoveApps $appsList
            continue
        }
        'RemoveDevApps' {
            $appsList = 'Microsoft.PowerAutomateDesktop', 'Microsoft.RemoteDesktop', 'Windows.DevHome'
            Write-Output "> 正在移除开发相关应用..."
            RemoveApps $appsList
            continue
        }
        'RemoveGamingApps' {
            $appsList = 'Microsoft.GamingApp', 'Microsoft.XboxGameOverlay', 'Microsoft.XboxGamingOverlay'
            Write-Output "> 正在移除游戏相关应用..."
            RemoveApps $appsList
            continue
        }
        "ForceRemoveEdge" {
            ForceRemoveEdge
            continue
        }
        'DisableDVR' {
            RegImport "> 正在禁用 Xbox 游戏/屏幕录制..." "Disable_DVR.reg"
            continue
        }
        'DisableTelemetry' {
            RegImport "> 正在禁用遥测、诊断数据、活动历史记录、应用启动跟踪和定向广告..." "Disable_Telemetry.reg"
            continue
        }
        {$_ -in "DisableSuggestions", "DisableWindowsSuggestions"} {
            RegImport "> 正在禁用 Windows 中的技巧提示、建议和广告..." "Disable_Windows_Suggestions.reg"
            continue
        }
        'DisableDesktopSpotlight' {
            RegImport "> 正在禁用 'Windows 聚焦' 桌面背景选项..." "Disable_Desktop_Spotlight.reg"
            continue
        }
        {$_ -in "DisableLockscrTips", "DisableLockscreenTips"} {
            RegImport "> 正在禁用锁屏界面上的技巧提示..." "Disable_Lockscreen_Tips.reg"
            continue
        }
        {$_ -in "DisableBingSearches", "DisableBing"} {
            RegImport "> 正在禁用 Windows 搜索中的必应网页搜索、必应 AI 和小娜..." "Disable_Bing_Cortana_In_Search.reg"
            
            # 同时移除必应搜索应用包
            $appsList = 'Microsoft.BingSearch'
            RemoveApps $appsList
            continue
        }
        'DisableCopilot' {
            RegImport "> 正在禁用并移除 Microsoft Copilot..." "Disable_Copilot.reg"

            # 同时移除 Copilot 应用包
            $appsList = 'Microsoft.Copilot'
            RemoveApps $appsList
            continue
        }
        'DisableRecall' {
            RegImport "> 正在禁用 Windows Recall 快照功能..." "Disable_AI_Recall.reg"
            continue
        }
        'RevertContextMenu' {
            RegImport "> 正在恢复旧版 Windows 10 样式右键菜单..." "Disable_Show_More_Options_Context_Menu.reg"
            continue
        }
        'DisableMouseAcceleration' {
            RegImport "> 正在关闭鼠标指针精度增强..." "Disable_Enhance_Pointer_Precision.reg"
            continue
        }
        'DisableStickyKeys' {
            RegImport "> 正在禁用粘滞键快捷键..." "Disable_Sticky_Keys_Shortcut.reg"
            continue
        }
        'ClearStart' {
            Write-Output "> 正在为用户 $(GetUserName) 移除开始菜单所有固定应用..."
            ReplaceStartMenu
            Write-Output ""
            continue
        }
        'ClearStartAllUsers' {
            ReplaceStartMenuForAllUsers
            continue
        }
        'DisableStartRecommended' {
            RegImport "> 正在禁用并隐藏开始菜单推荐版块..." "Disable_Start_Recommended.reg"
            continue
        }
        'TaskbarAlignLeft' {
            RegImport "> 正在将任务栏按钮左对齐..." "Align_Taskbar_Left.reg"
            continue
        }
        'HideSearchTb' {
            RegImport "> 正在隐藏任务栏搜索图标..." "Hide_Search_Taskbar.reg"
            continue
        }
        'ShowSearchIconTb' {
            RegImport "> 正在将任务栏搜索设置为仅显示图标..." "Show_Search_Icon.reg"
            continue
        }
        'ShowSearchLabelTb' {
            RegImport "> 正在将任务栏搜索设置为图标加文字..." "Show_Search_Icon_And_Label.reg"
            continue
        }
        'ShowSearchBoxTb' {
            RegImport "> 正在将任务栏搜索设置为完整搜索框..." "Show_Search_Box.reg"
            continue
        }
        'HideTaskview' {
            RegImport "> 正在隐藏任务栏任务视图按钮..." "Hide_Taskview_Taskbar.reg"
            continue
        }
        {$_ -in "HideWidgets", "DisableWidgets"} {
            RegImport "> 正在禁用小组件服务并隐藏任务栏小组件图标..." "Disable_Widgets_Taskbar.reg"
            continue
        }
        {$_ -in "HideChat", "DisableChat"} {
            RegImport "> 正在隐藏任务栏聊天图标..." "Disable_Chat_Taskbar.reg"
            continue
        }
        'EnableEndTask' {
            RegImport "> 正在启用任务栏右键菜单中的 '结束任务' 选项..." "Enable_End_Task.reg"
            continue
        }
        'ExplorerToHome' {
            RegImport "> 正在将文件资源管理器默认打开位置设置为 `主页`..." "Launch_File_Explorer_To_Home.reg"
            continue
        }
        'ExplorerToThisPC' {
            RegImport "> 正在将文件资源管理器默认打开位置设置为 `此电脑`..." "Launch_File_Explorer_To_This_PC.reg"
            continue
        }
        'ExplorerToDownloads' {
            RegImport "> 正在将文件资源管理器默认打开位置设置为 `下载`..." "Launch_File_Explorer_To_Downloads.reg"
            continue
        }
        'ExplorerToOneDrive' {
            RegImport "> 正在将文件资源管理器默认打开位置设置为 `OneDrive`..." "Launch_File_Explorer_To_OneDrive.reg"
            continue
        }
        'ShowHiddenFolders' {
            RegImport "> 正在显示隐藏文件、文件夹和驱动器..." "Show_Hidden_Folders.reg"
            continue
        }
        'ShowKnownFileExt' {
            RegImport "> 正在启用已知文件类型的扩展名显示..." "Show_Extensions_For_Known_File_Types.reg"
            continue
        }
        'HideHome' {
            RegImport "> 正在隐藏文件资源管理器导航窗格中的主页版块..." "Hide_Home_from_Explorer.reg"
            continue
        }
        'HideGallery' {
            RegImport "> 正在隐藏文件资源管理器导航窗格中的图库版块..." "Hide_Gallery_from_Explorer.reg"
            continue
        }
        'HideDupliDrive' {
            RegImport "> 正在隐藏文件资源管理器导航窗格中的重复可移动驱动器..." "Hide_duplicate_removable_drives_from_navigation_pane_of_File_Explorer.reg"
            continue
        }
        {$_ -in "HideOnedrive", "DisableOnedrive"} {
            RegImport "> 正在隐藏文件资源管理器导航窗格中的 OneDrive 文件夹..." "Hide_Onedrive_Folder.reg"
            continue
        }
        {$_ -in "Hide3dObjects", "Disable3dObjects"} {
            RegImport "> 正在隐藏文件资源管理器导航窗格中的 3D 对象文件夹..." "Hide_3D_Objects_Folder.reg"
            continue
        }
        {$_ -in "HideMusic", "DisableMusic"} {
            RegImport "> 正在隐藏文件资源管理器导航窗格中的音乐文件夹..." "Hide_Music_folder.reg"
            continue
        }
        {$_ -in "HideIncludeInLibrary", "DisableIncludeInLibrary"} {
            RegImport "> 正在隐藏右键菜单中的 '包含到库中' 选项..." "Disable_Include_in_library_from_context_menu.reg"
            continue
        }
        {$_ -in "HideGiveAccessTo", "DisableGiveAccessTo"} {
            RegImport "> 正在隐藏右键菜单中的 '授予访问权限' 选项..." "Disable_Give_access_to_context_menu.reg"
            continue
        }
        {$_ -in "HideShare", "DisableShare"} {
            RegImport "> 正在隐藏右键菜单中的 '共享' 选项..." "Disable_Share_from_context_menu.reg"
            continue
        }
    }

    RestartExplorer

    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output "脚本执行完毕！请检查上方是否有错误信息。"

    AwaitKeyToExit
    }    
