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


# �� PowerShell ִ�л����ܰ�ȫ��������ʱ��ʾ����
if ($ExecutionContext.SessionState.LanguageMode -ne "FullLanguage") {
    Write-Host "����Win11Debloat �޷�������ϵͳ�����У�PowerShell ִ���ܵ���ȫ��������" -ForegroundColor Red
    Write-Output ""
    Write-Output "���س����˳�..."
    Read-Host | Out-Null
    Exit
}


# ��ʾӦ�ó���ѡ����������û�ѡ��Ҫɾ��������Ӧ�ó���
function ShowAppSelectionForm {
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

    # ��ʼ��������
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

    # ���水ť�¼�����
    $handler_saveButton_Click= 
    {
        if ($selectionBox.CheckedItems -contains "Microsoft.WindowsStore" -and -not $Silent) {
            $warningSelection = [System.Windows.Forms.Messagebox]::Show('��ȷ��Ҫж��Microsoft Store�𣿴�Ӧ���޷��������°�װ��', 'ȷ�ϲ���', 'YesNo', 'Warning')
        
            if ($warningSelection -eq 'No') {
                return
            }
        }

        $global:SelectedApps = $selectionBox.CheckedItems

        # �����洢��ѡӦ�õ��ļ�����������ڣ�
        if (!(Test-Path "$PSScriptRoot/CustomAppsList")) {
            $null = New-Item "$PSScriptRoot/CustomAppsList"
        } 

        Set-Content -Path "$PSScriptRoot/CustomAppsList" -Value $global:SelectedApps

        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }

    # ȡ����ť�¼�����
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
        # ��������ʼ״̬�Է�ֹ�������
        $form.WindowState = $initialFormWindowState

        # ����ѡ�������
        $global:selectionBoxIndex = -1
        $checkUncheckCheckBox.Checked = $False

        # ��ʾ����ָʾ��
        $loadingLabel.Visible = $true
        $form.Refresh()

        # ���ѡ���
        $selectionBox.Items.Clear()

        # ����Ӧ���б��ļ�·��
        $appsFile = "$PSScriptRoot/Appslist.txt"
        $listOfApps = ""

        if ($onlyInstalledCheckBox.Checked -and ($global:wingetInstalled -eq $true)) {
            # ����ͨ�� winget ��ȡ�Ѱ�װӦ���б�10�볬ʱ��
            $job = Start-Job { return winget list --accept-source-agreements --disable-interactivity }
            $jobDone = $job | Wait-Job -TimeOut 10

            if (-not $jobDone) {
                [System.Windows.MessageBox]::Show('�޷�ͨ�� winget �����Ѱ�װӦ���б�����Ӧ�ÿ���δ��ʾ��', '����', 'Ok', 'Error')
            }
            else {
                $listOfApps = Receive-Job -Job $job
            }
        }

        # ����Ӧ���б�
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

    $form.Text = "Win11Debloat Ӧ��ѡ��"
    $form.Name = "appSelectionForm"
    $form.DataBindings.DefaultDataSourceUpdateMode = 0
    $form.ClientSize = New-Object System.Drawing.Size(400,502)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $False

    $button1.TabIndex = 4
    $button1.Name = "saveButton"
    $button1.UseVisualStyleBackColor = $True
    $button1.Text = "ȷ��"
    $button1.Location = New-Object System.Drawing.Point(27,472)
    $button1.Size = New-Object System.Drawing.Size(75,23)
    $button1.DataBindings.DefaultDataSourceUpdateMode = 0
    $button1.add_Click($handler_saveButton_Click)

    $form.Controls.Add($button1)

    $button2.TabIndex = 5
    $button2.Name = "cancelButton"
    $button2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $button2.UseVisualStyleBackColor = $True
    $button2.Text = "ȡ��"
    $button2.Location = New-Object System.Drawing.Point(129,472)
    $button2.Size = New-Object System.Drawing.Size(75,23)
    $button2.DataBindings.DefaultDataSourceUpdateMode = 0
    $button2.add_Click($handler_cancelButton_Click)

    $form.Controls.Add($button2)

    $label.Location = New-Object System.Drawing.Point(13,5)
    $label.Size = New-Object System.Drawing.Size(400,14)
    $Label.Font = 'Microsoft Sans Serif,8'
    $label.Text = '��ѡ��ϣ��ж�ص�Ӧ�ã�ȡ����ѡ��ϣ��������Ӧ��'

    $form.Controls.Add($label)

    $loadingLabel.Location = New-Object System.Drawing.Point(16,46)
    $loadingLabel.Size = New-Object System.Drawing.Size(300,418)
    $loadingLabel.Text = '���ڼ���Ӧ��...'
    $loadingLabel.BackColor = "White"
    $loadingLabel.Visible = $false

    $form.Controls.Add($loadingLabel)

    $onlyInstalledCheckBox.TabIndex = 6
    $onlyInstalledCheckBox.Location = New-Object System.Drawing.Point(230,474)
    $onlyInstalledCheckBox.Size = New-Object System.Drawing.Size(150,20)
    $onlyInstalledCheckBox.Text = '����ʾ�Ѱ�װӦ��'
    $onlyInstalledCheckBox.add_CheckedChanged($load_Apps)

    $form.Controls.Add($onlyInstalledCheckBox)

    $checkUncheckCheckBox.TabIndex = 7
    $checkUncheckCheckBox.Location = New-Object System.Drawing.Point(16,22)
    $checkUncheckCheckBox.Size = New-Object System.Drawing.Size(150,20)
    $checkUncheckCheckBox.Text = 'ȫѡ/ȡ��ȫѡ'
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


# ��ָ���ļ���ȡӦ���б�
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


# ж��ָ��Ӧ��
function RemoveApps {
    param (
        $appslist
    )

    Foreach ($app in $appsList) { 
        Write-Output "���ڳ���ж�� $app..."

        if (($app -eq "Microsoft.OneDrive") -or ($app -eq "Microsoft.Edge")) {
            if ($global:wingetInstalled -eq $false) {
                Write-Host "����WinGet δ��װ��汾���ɣ��޷��Ƴ� $app" -ForegroundColor Red
            }
            else {
                Strip-Progress -ScriptBlock { winget uninstall --accept-source-agreements --disable-interactivity --id $app } | Tee-Object -Variable wingetOutput 

                If (($app -eq "Microsoft.Edge") -and (Select-String -InputObject $wingetOutput -Pattern "Uninstall failed with exit code")) {
                    Write-Host "�޷�ͨ�� Winget ж�� Microsoft Edge" -ForegroundColor Red
                    Write-Output ""

                    if ($( Read-Host -Prompt "�Ƿ�Ҫǿ��ж��Edge��(���Ƽ�) (y/n)" ) -eq 'y') {
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
                        Write-Host "�޷�Ϊ�����û��Ƴ� $app" -ForegroundColor Yellow
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
                Write-Host "�޷���ϵͳ�����Ƴ� $app" -ForegroundColor Yellow
            }
        }
    }
            
    Write-Output ""
}


# ǿ���Ƴ� Microsoft Edge
function ForceRemoveEdge {
    Write-Output "> ����ǿ��ж�� Microsoft Edge..."

    $regView = [Microsoft.Win32.RegistryView]::Registry32
    $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $regView)
    $hklm.CreateSubKey('SOFTWARE\Microsoft\EdgeUpdateDev').SetValue('AllowUninstall', '')

    $edgeStub = "$env:SystemRoot\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe"
    New-Item $edgeStub -ItemType Directory | Out-Null
    New-Item "$edgeStub\MicrosoftEdge.exe" | Out-Null

    $uninstallRegKey = $hklm.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge')
    if ($null -ne $uninstallRegKey) {
        Write-Output "��������ж�س���..."
        $uninstallString = $uninstallRegKey.GetValue('UninstallString') + ' --force-uninstall'
        Start-Process cmd.exe "/c $uninstallString" -WindowStyle Hidden -Wait

        Write-Output "�����Ƴ������ļ�..."
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

        Write-Output "��������ע���..."
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "MicrosoftEdgeAutoLaunch_A9F6DCE4ABADF4F51CF45CD7129E3C6C" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "Microsoft Edge Update" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" /v "MicrosoftEdgeAutoLaunch_A9F6DCE4ABADF4F51CF45CD7129E3C6C" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" /v "Microsoft Edge Update" /f *>$null

        Write-Output "Microsoft Edge ��ж��"
    }
    else {
        Write-Output ""
        Write-Host "�����޷��ҵ� Microsoft Edge ж�س���" -ForegroundColor Red
    }
    
    Write-Output ""
}


# ִ��������������ʾ
function Strip-Progress {
    param(
        [ScriptBlock]$ScriptBlock
    )

    $progressPattern = '���0�4[�0�4��]|^\s+[-\\|/]\s+$'
    $sizePattern = '(\d+(\.\d{1,2})?)\s+(B|KB|MB|GB|TB|PB) /\s+(\d+(\.\d{1,2})?)\s+(B|KB|MB|GB|TB|PB)'

    & $ScriptBlock 2>&1 | ForEach-Object {
        if ($_ -is [System.Management.Automation.ErrorRecord]) {
            "����$($_.Exception.Message)"
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

    Write-Output "> �������� Windows ��Դ������������Ӧ�ø���... (���ܻ������˸)"

    if ($global:Params.ContainsKey("DisableMouseAcceleration")) {
        Write-Host "���棺ָ�뾫ȷ����ǿ���õĸ�����Ҫ�����������Ч" -ForegroundColor Yellow
    }

    if ($global:Params.ContainsKey("DisableStickyKeys")) {
        Write-Host "���棺ճ�ͼ����õĸ�����Ҫ�����������Ч" -ForegroundColor Yellow
    }

    # Only restart if the powershell process matches the OS architecture.
    # Restarting explorer from a 32bit Powershell window will fail on a 64bit OS
    if ([Environment]::Is64BitProcess -eq [Environment]::Is64BitOperatingSystem) {
        Stop-Process -processName: Explorer -Force
    }
    else {
        Write-Warning "�޷��Զ����� Windows ��Դ���������̣����ֶ������������Ӧ�����и��ġ�"
    }
}


# Replace the startmenu for all users, when using the default startmenuTemplate this clears all pinned apps
# Credit: https://lazyadmin.nl/win-11/customize-windows-11-start-menu-layout/
function ReplaceStartMenuForAllUsers {
    param (
        $startMenuTemplate = "$PSScriptRoot/Start/start2.bin"
    )

    Write-Output "> ����Ϊ�����û��Ƴ���ʼ�˵��е����й̶�Ӧ��..."

    # Check if template bin file exists, return early if it doesn't
    if (-not (Test-Path $startMenuTemplate)) {
        Write-Host "�����޷������ʼ�˵����ű�Ŀ¼��ȱ�� start2.bin �ļ�" -ForegroundColor Red
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
        Write-Output "��ΪĬ���û������ļ����� LocalState �ļ���"
    }

    # Copy template to default profile
    Copy-Item -Path $startMenuTemplate -Destination $defaultStartMenuPath -Force
    Write-Output "���滻Ĭ���û������ļ��Ŀ�ʼ�˵�"
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
        Write-Host "�����޷������ʼ�˵����ű�Ŀ¼��ȱ�� start2.bin �ļ�" -ForegroundColor Red
        return
    }

    # Check if bin file exists, return early if it doesn't
    if (-not (Test-Path $startMenuBinFile)) {
        Write-Host "�����޷�Ϊ�û� $(GetUserName) �����ʼ�˵����Ҳ��� start2.bin �ļ�" -ForegroundColor Red
        return
    }

    $backupBinFile = $startMenuBinFile + ".bak"

    # Backup current start menu file
    Move-Item -Path $startMenuBinFile -Destination $backupBinFile -Force

    # Copy template file
    Copy-Item -Path $startMenuTemplate -Destination $startMenuBinFile -Force

    Write-Output "���滻�û� $(GetUserName) �Ŀ�ʼ�˵�"
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

    $fullTitle = " Win11Debloat �ű� - $title"

    if ($global:Params.ContainsKey("Sysprep")) {
        $fullTitle = "$fullTitle (Sysprep ģʽ)"
    }
    else {
        $fullTitle = "$fullTitle (�û���$(GetUserName))"
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
        Write-Output "��������˳�..."
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
            
    PrintHeader '�Զ���ģʽ'

    # Show options for removing apps, only continue on valid input
    Do {
        Write-Host "ѡ�" -ForegroundColor Yellow
        Write-Host " (n) ���Ƴ��κ�Ӧ��" -ForegroundColor Yellow
        Write-Host " (1) ���Ƴ� Appslist.txt �е�Ĭ��ԤװӦ��" -ForegroundColor Yellow
        Write-Host " (2) �Ƴ�Ĭ��ԤװӦ�ã��Լ��ʼ�����Ӧ�á������߹��ߺ���Ϸ���Ӧ��"  -ForegroundColor Yellow
        Write-Host " (3) �Զ���ѡ��Ҫ�Ƴ���Ӧ��" -ForegroundColor Yellow
        $RemoveAppsInput = Read-Host "�Ƿ��Ƴ�ԤװӦ�ã�(n/1/2/3)" 

        # Show app selection form if user entered option 3
        if ($RemoveAppsInput -eq '3') {
            $result = ShowAppSelectionForm

            if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
                # User cancelled or closed app selection, show error and change RemoveAppsInput so the menu will be shown again
                Write-Output ""
                Write-Host "��ȡ��Ӧ��ѡ��������" -ForegroundColor Red

                $RemoveAppsInput = 'c'
            }
            
            Write-Output ""
        }
    }
    while ($RemoveAppsInput -ne 'n' -and $RemoveAppsInput -ne '0' -and $RemoveAppsInput -ne '1' -and $RemoveAppsInput -ne '2' -and $RemoveAppsInput -ne '3') 

    # Select correct option based on user input
    switch ($RemoveAppsInput) {
        '1' {
            AddParameter 'RemoveApps' '�Ƴ�Ĭ��ԤװӦ��'
        }
        '2' {
            AddParameter 'RemoveApps' '�Ƴ�Ĭ��ԤװӦ��'
            AddParameter 'RemoveCommApps' '�Ƴ��ʼ�����������ϵ��Ӧ��'
            AddParameter 'RemoveW11Outlook' '�Ƴ��°� Outlook for Windows Ӧ��'
            AddParameter 'RemoveDevApps' '�Ƴ����������Ӧ��'
            AddParameter 'RemoveGamingApps' '�Ƴ� Xbox Ӧ�ú� Xbox ��Ϸ��'
            AddParameter 'DisableDVR' '���� Xbox ��Ϸ/��Ļ¼��'
        }
        '3' {
            Write-Output "����ѡ���Ƴ� $($global:SelectedApps.Count) ��Ӧ��"

            AddParameter 'RemoveAppsCustom' "�Ƴ� $($global:SelectedApps.Count) ��Ӧ�ã�"

            Write-Output ""

            if ($( Read-Host -Prompt "�Ƿ���� Xbox ��Ϸ/��Ļ¼�ƣ�ͬʱ������Ϸ���� (y/n)" ) -eq 'y') {
                AddParameter 'DisableDVR' '���� Xbox ��Ϸ/��Ļ¼��'
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "�Ƿ����ң�⡢������ݡ����ʷ��Ӧ������׷�ٺͶ����棿(y/n)" ) -eq 'y') {
        AddParameter 'DisableTelemetry' '����ң�⡢������ݡ����ʷ��Ӧ������׷�ٺͶ�����'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "�Ƿ���ÿ�ʼ�˵������á�֪ͨ���ļ���Դ�������е���ʾ�����ɡ�����͹�棿(y/n)" ) -eq 'y') {
        AddParameter 'DisableSuggestions' '���ÿ�ʼ�˵������á�֪ͨ���ļ���Դ�������е���ʾ�����ɡ�����͹��'
        AddParameter 'DisableDesktopSpotlight' '���� Windows Spotlight ���汳��ѡ��'
        AddParameter 'DisableLockscreenTips' '���������������ʾ�ͼ���'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "�Ƿ���ò��Ƴ� Windows �����е� Bing ��ҳ������Bing AI �� Cortana��(y/n)" ) -eq 'y') {
        AddParameter 'DisableBing' '���ò��Ƴ� Windows �����е� Bing ��ҳ������Bing AI �� Cortana'
    }

    # Only show this option for Windows 11 users running build 22621 or later
    if ($WinVersion -ge 22621){
        Write-Output ""

        if ($( Read-Host -Prompt "�Ƿ�Ϊ�����û����ò��Ƴ� Microsoft Copilot �� Windows Recall ���գ�(y/n)" ) -eq 'y') {
            AddParameter 'DisableCopilot' '���ò��Ƴ� Microsoft Copilot'
            AddParameter 'DisableRecall' '���ò��Ƴ� Windows Recall ����'
        }
    }

    # Only show this option for Windows 11 users running build 22000 or later
    if ($WinVersion -ge 22000){
        Write-Output ""

        if ($( Read-Host -Prompt "�Ƿ�ָ� Windows 10 ��ʽ���Ҽ��˵���(y/n)" ) -eq 'y') {
            AddParameter 'RevertContextMenu' '�ָ� Windows 10 ��ʽ���Ҽ��˵�'
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "�Ƿ�ر�ָ�뾫ȷ����ǿ�������٣���(y/n)" ) -eq 'y') {
        AddParameter 'DisableMouseAcceleration' '�ر�ָ�뾫ȷ����ǿ�������٣�'
    }

    # Only show this option for Windows 11 users running build 26100 or later
    if ($WinVersion -ge 26100){
        Write-Output ""

        if ($( Read-Host -Prompt "�Ƿ����ճ�ͼ���ݼ���(y/n)" ) -eq 'y') {
            AddParameter 'DisableStickyKeys' '����ճ�ͼ���ݼ�'
        }
    }

    # Only show option for disabling context menu items for Windows 10 users or if the user opted to restore the Windows 10 context menu
    if ((get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'") -or $global:Params.ContainsKey('RevertContextMenu')){
        Write-Output ""

        if ($( Read-Host -Prompt "�Ƿ���Ҫ����ĳЩ�Ҽ��˵�ѡ�(y/n)" ) -eq 'y') {
            Write-Output ""

            if ($( Read-Host -Prompt "   �����Ҽ��˵��е� '����������' ѡ�(y/n)" ) -eq 'y') {
                AddParameter 'HideIncludeInLibrary' "�����Ҽ��˵��е� '����������' ѡ��"
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   �����Ҽ��˵��е� '�������Ȩ��' ѡ�(y/n)" ) -eq 'y') {
                AddParameter 'HideGiveAccessTo' "�����Ҽ��˵��е� '�������Ȩ��' ѡ��"
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   �����Ҽ��˵��е� '����' ѡ�(y/n)" ) -eq 'y') {
                AddParameter 'HideShare' "�����Ҽ��˵��е� '����' ѡ��"
            }
        }
    }

    # Only show this option for Windows 11 users running build 22621 or later
    if ($WinVersion -ge 22621){
        Write-Output ""

        if ($( Read-Host -Prompt "�Ƿ���Ҫ�޸Ŀ�ʼ�˵���(y/n)" ) -eq 'y') {
            Write-Output ""

            if ($global:Params.ContainsKey("Sysprep")) {
                if ($( Read-Host -Prompt "�Ƿ�Ϊ���������û������û��Ƴ���ʼ�˵��е����й̶�Ӧ�ã�(y/n)" ) -eq 'y') {
                    AddParameter 'ClearStartAllUsers' 'Ϊ�����û��Ƴ���ʼ�˵��еĹ̶�Ӧ��'
                }
            }
            else {
                Do {
                    Write-Host "   ѡ�" -ForegroundColor Yellow
                    Write-Host "    (n) ���Ƴ��κι̶�Ӧ��" -ForegroundColor Yellow
                    Write-Host "    (1) ���Ƴ���ǰ�û� ($(GetUserName)) �Ĺ̶�Ӧ��" -ForegroundColor Yellow
                    Write-Host "    (2) Ϊ���������û������û��Ƴ��̶�Ӧ��"  -ForegroundColor Yellow
                    $ClearStartInput = Read-Host "   �Ƴ���ʼ�˵��еĹ̶�Ӧ�ã�(n/1/2)" 
                }
                while ($ClearStartInput -ne 'n' -and $ClearStartInput -ne '0' -and $ClearStartInput -ne '1' -and $ClearStartInput -ne '2') 

                # Select correct option based on user input
                switch ($ClearStartInput) {
                    '1' {
                        AddParameter 'ClearStart' "�Ƴ���ǰ�û� ($(GetUserName)) �Ĺ̶�Ӧ��"
                    }
                    '2' {
                        AddParameter 'ClearStartAllUsers' "Ϊ�����û��Ƴ���ʼ�˵��еĹ̶�Ӧ��"
                    }
                }
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   �Ƿ�Ϊ�����û����ò����ؿ�ʼ�˵����Ƽ�����(y/n)" ) -eq 'y') {
                AddParameter 'DisableStartRecommended' '���ò����ؿ�ʼ�˵����Ƽ�����'
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "�Ƿ���Ҫ�޸�����������ط���(y/n)" ) -eq 'y') {
        # Only show these specific options for Windows 11 users running build 22000 or later
        if ($WinVersion -ge 22000){
            Write-Output ""

            if ($( Read-Host -Prompt "   �Ƿ�������ͼ������룿(y/n)" ) -eq 'y') {
                AddParameter 'TaskbarAlignLeft' '������ͼ�������'
            }

            # Show options for search icon on taskbar, only continue on valid input
            Do {
                Write-Output ""
                Write-Host "   ѡ�" -ForegroundColor Yellow
                Write-Host "    (n) ���ֵ�ǰ����" -ForegroundColor Yellow
                Write-Host "    (1) ��������������ͼ��" -ForegroundColor Yellow
                Write-Host "    (2) ��ʾ����������ͼ��" -ForegroundColor Yellow
                Write-Host "    (3) ��ʾ����ǩ������ͼ��" -ForegroundColor Yellow
                Write-Host "    (4) ��ʾ������" -ForegroundColor Yellow
                $TbSearchInput = Read-Host "   �Ƿ��޸�����������ͼ�ꣿ(n/1/2/3/4)" 
            }
            while ($TbSearchInput -ne 'n' -and $TbSearchInput -ne '0' -and $TbSearchInput -ne '1' -and $TbSearchInput -ne '2' -and $TbSearchInput -ne '3' -and $TbSearchInput -ne '4') 

            # Select correct taskbar search option based on user input
            switch ($TbSearchInput) {
                '1' {
                    AddParameter 'HideSearchTb' '��������������ͼ��'
                }
                '2' {
                    AddParameter 'ShowSearchIconTb' '��ʾ����������ͼ��'
                }
                '3' {
                    AddParameter 'ShowSearchLabelTb' '��ʾ����ǩ������ͼ��'
                }
                '4' {
                    AddParameter 'ShowSearchBoxTb' '��ʾ������'
                }
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   �Ƿ�����������ͼ��ť��(y/n)" ) -eq 'y') {
                AddParameter 'HideTaskview' '����������ͼ��ť'
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   �Ƿ����С�����������������ͼ�ꣿ(y/n)" ) -eq 'y') {
            AddParameter 'DisableWidgets' '����С�����������������ͼ��'
        }

        # Only show this options for Windows users running build 22621 or earlier
        if ($WinVersion -le 22621){
            Write-Output ""

            if ($( Read-Host -Prompt "   �Ƿ��������������죨�������ᣩͼ�ꣿ(y/n)" ) -eq 'y') {
                AddParameter 'HideChat' '��������������ͼ��'
            }
        }
        
        # Only show this options for Windows users running build 22631 or later
        if ($WinVersion -ge 22631){
            Write-Output ""

            if ($( Read-Host -Prompt "   �Ƿ����������Ҽ��˵����� '��������' ѡ�(y/n)" ) -eq 'y') {
                AddParameter 'EnableEndTask' "�����������Ҽ��˵��� '��������' ѡ��"
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "�Ƿ�Ҫ�޸��ļ���Դ���������ã�(y/n)" ) -eq 'y') {
        # �޸��ļ���Դ������Ĭ��λ�õ�ѡ��
        Do {
            Write-Output ""
            Write-Host "   ѡ�" -ForegroundColor Yellow
            Write-Host "    (n) ���޸�" -ForegroundColor Yellow
            Write-Host "    (1) �ļ���Դ������Ĭ�ϴ�'��ҳ'" -ForegroundColor Yellow
            Write-Host "    (2) �ļ���Դ������Ĭ�ϴ�'�˵���'" -ForegroundColor Yellow
            Write-Host "    (3) �ļ���Դ������Ĭ�ϴ�'����'" -ForegroundColor Yellow
            Write-Host "    (4) �ļ���Դ������Ĭ�ϴ�'OneDrive'" -ForegroundColor Yellow
            $ExplSearchInput = Read-Host "   ��ѡ���ļ���Դ��������Ĭ�ϴ�λ�� (n/1/2/3/4)" 
        }
        while ($ExplSearchInput -ne 'n' -and $ExplSearchInput -ne '0' -and $ExplSearchInput -ne '1' -and $ExplSearchInput -ne '2' -and $ExplSearchInput -ne '3' -and $ExplSearchInput -ne '4') 

        # �����û�����ѡ������
        switch ($ExplSearchInput) {
            '1' {
                AddParameter 'ExplorerToHome' "���ļ���Դ������Ĭ�ϴ�λ������Ϊ'��ҳ'"
            }
            '2' {
                AddParameter 'ExplorerToThisPC' "���ļ���Դ������Ĭ�ϴ�λ������Ϊ'�˵���'"
            }
            '3' {
                AddParameter 'ExplorerToDownloads' "���ļ���Դ������Ĭ�ϴ�λ������Ϊ'����'"
            }
            '4' {
                AddParameter 'ExplorerToOneDrive' "���ļ���Դ������Ĭ�ϴ�λ������Ϊ'OneDrive'"
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   �Ƿ���ʾ���ص��ļ����ļ��к���������(y/n)" ) -eq 'y') {
            AddParameter 'ShowHiddenFolders' '��ʾ���ص��ļ����ļ��к�������'
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   �Ƿ���ʾ��֪�ļ����͵���չ����(y/n)" ) -eq 'y') {
            AddParameter 'ShowKnownFileExt' '��ʾ��֪�ļ����͵���չ��'
        }

        # ��������Windows 11 22000����߰汾
        if ($WinVersion -ge 22000){
            Write-Output ""

            if ($( Read-Host -Prompt "   �Ƿ���ļ���Դ���������������'��ҳ'���֣�(y/n)" ) -eq 'y') {
                AddParameter 'HideHome' '���ļ���Դ�����������������ҳ����'
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   �Ƿ���ļ���Դ���������������'ͼ��'���֣�(y/n)" ) -eq 'y') {
                AddParameter 'HideGallery' '���ļ���Դ���������������ͼ�ⲿ��'
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   �Ƿ���ļ���Դ����������������ظ��Ŀ��ƶ���������Ŀ��������'�˵���'����ʾ��(y/n)" ) -eq 'y') {
            AddParameter 'HideDupliDrive' '���ļ���Դ����������������ظ��Ŀ��ƶ���������Ŀ'
        }

        # ��Windows 10��ʾ�ض��ļ�������ѡ��
        if (get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'"){
            Write-Output ""

            if ($( Read-Host -Prompt "�Ƿ�Ҫ���ļ���Դ���������������ĳЩ�ļ��У�(y/n)" ) -eq 'y') {
                Write-Output ""

                if ($( Read-Host -Prompt "   �Ƿ�����OneDrive�ļ��У�(y/n)" ) -eq 'y') {
                    AddParameter 'HideOnedrive' '���ļ���Դ���������������OneDrive�ļ���'
                }

                Write-Output ""
                
                if ($( Read-Host -Prompt "   �Ƿ�����3D�����ļ��У�(y/n)" ) -eq 'y') {
                    AddParameter 'Hide3dObjects' "���ļ���Դ��������'�˵���'������3D�����ļ���" 
                }
                
                Write-Output ""

                if ($( Read-Host -Prompt "   �Ƿ����������ļ��У�(y/n)" ) -eq 'y') {
                    AddParameter 'HideMusic' "���ļ���Դ��������'�˵���'�����������ļ���"
                }
            }
        }
    }

    # ��Ĭģʽ�²���ʾȷ����ʾ
    if (-not $Silent) {
        Write-Output ""
        Write-Output ""
        Write-Output ""
        Write-Output "���س�ȷ��ѡ��ִ�нű�����CTRL+C�˳�..."
        Read-Host | Out-Null
    }

    PrintHeader '�Զ���ģʽ'
}

##################################################################################################################
#                                                                                                                #
#                                                  �ű���ʼִ��                                                  #
#                                                                                                                #
##################################################################################################################

# ����Ƿ�װwinget��v1.4+��
if ((Get-AppxPackage -Name "*Microsoft.DesktopAppInstaller*") -and ([int](((winget -v) -replace 'v','').split('.')[0..1] -join '') -gt 14)) {
    $global:wingetInstalled = $true
}
else {
    $global:wingetInstalled = $false

    # �Ǿ�Ĭģʽ����ʾ����
    if (-not $Silent) {
        Write-Warning "δ��⵽winget��汾���ͣ���Ҫv1.4+�������ܻ�Ӱ��Ӧ��ж�ع���"
        Write-Output ""
        Write-Output "�����������..."
        $null = [System.Console]::ReadKey()
    }
}

# ��ȡ��ǰWindows�汾
$WinVersion = Get-ItemPropertyValue 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' CurrentBuild

$global:Params = $PSBoundParameters
$global:FirstSelection = $true
$SPParams = 'WhatIf', 'Confirm', 'Verbose', 'Silent', 'Sysprep', 'Debug', 'User'
$SPParamCount = 0

# ͳ��SPParams��������
foreach ($Param in $SPParams) {
    if ($global:Params.ContainsKey($Param)) {
        $SPParamCount++
    }
}

# ���ؽ���������Ĭģʽ��
if (-not ($global:Params.ContainsKey("Verbose"))) {
    $ProgressPreference = 'SilentlyContinue'
}
else {
    Read-Host "��������ϸģʽ�����س�����" 
    $ProgressPreference = 'Continue'
}

# Sysprepģʽ���
if ($global:Params.ContainsKey("Sysprep")) {
    $defaultUserPath = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), '\Default\NTUSER.DAT'

    if (-not (Test-Path "$defaultUserPath")) {
        Write-Host "�����޷��ҵ�Ĭ���û������ļ���'$defaultUserPath'" -ForegroundColor Red
        AwaitKeyToExit
        Exit
    }
    if ($WinVersion -lt 22000) {
        Write-Host "����Sysprepģʽ��֧��Windows 10" -ForegroundColor Red
        AwaitKeyToExit
        Exit
    }
}

# �û�ģʽ���
if ($global:Params.ContainsKey("User")) {
    $userPath = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), "\$($global:Params.Item("User"))\NTUSER.DAT"

    if (-not (Test-Path "$userPath")) {
        Write-Host "�����޷��ҵ��û� $($global:Params.Item("User")) �������ļ���'$userPath'" -ForegroundColor Red
        AwaitKeyToExit
        Exit
    }
}

# ����������ļ�
if ((Test-Path "$PSScriptRoot/SavedSettings") -and ([String]::IsNullOrWhiteSpace((Get-content "$PSScriptRoot/SavedSettings")))) {
    Remove-Item -Path "$PSScriptRoot/SavedSettings" -recurse
}

# Ӧ���б�����ģʽ
if ($RunAppConfigurator -or $RunAppsListGenerator) {
    PrintHeader "�Զ���Ӧ���б�������"

    $result = ShowAppSelectionForm

    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Ӧ��ѡ��δ����" -ForegroundColor Red
    }
    else {
        Write-Output "����Ӧ��ѡ���ѱ���������·����'CustomAppsList'�ļ���"
        Write-Host "$PSScriptRoot" -ForegroundColor Yellow
    }

    AwaitKeyToExit
    Exit
}

# �ű�ִ��ģʽѡ��
if ((-not $global:Params.Count) -or $RunDefaults -or $RunWin11Defaults -or $RunSavedSettings -or ($SPParamCount -eq $global:Params.Count)) {
    if ($RunDefaults -or $RunWin11Defaults) {
        $Mode = '1'
    }
    elseif ($RunSavedSettings) {
        if(-not (Test-Path "$PSScriptRoot/SavedSettings")) {
            PrintHeader '�Զ���ģʽ'
            Write-Host "����δ�ҵ��ѱ��������" -ForegroundColor Red
            AwaitKeyToExit
            Exit
        }

        $Mode = '4'
    }
    else {
        Do { 
            $ModeSelectionMessage = "��ѡ��һ��ѡ�� (1/2/3/0)" 

            PrintHeader '���˵�'

            Write-Output "(1) Ĭ��ģʽ��Ӧ��Ĭ������"
            Write-Output "(2) �Զ���ģʽ�������޸Ľű�����"
            Write-Output "(3) Ӧ���Ƴ�ģʽ����ѡ���Ƴ�Ӧ��"

            if (Test-Path "$PSScriptRoot/SavedSettings") {
                Write-Output "(4) Ӧ���ϴα�����Զ�������"
                $ModeSelectionMessage = "��ѡ��һ��ѡ�� (1/2/3/4/0)" 
            }

            Write-Output ""
            Write-Output "(0) ��ʾ������Ϣ"
            Write-Output ""
            Write-Output ""

            $Mode = Read-Host $ModeSelectionMessage

            if ($Mode -eq '0') {
                PrintFromFile "$PSScriptRoot/Assets/Menus/Info"

                Write-Output ""
                Write-Output "�����������..."
                $null = [System.Console]::ReadKey()
            }
            elseif (($Mode -eq '4')-and -not (Test-Path "$PSScriptRoot/SavedSettings")) {
                $Mode = $null
            }
        }
        while ($Mode -ne '1' -and $Mode -ne '2' -and $Mode -ne '3' -and $Mode -ne '4') 
    }

    # ����ģʽ��Ӳ���
    switch ($Mode) {
        '1' { 
            if (-not $Silent) {
                PrintFromFile "$PSScriptRoot/Assets/Menus/DefaultSettings"

                Write-Output ""
                Write-Output "���س���ʼִ�У���CTRL+C�˳�..."
                Read-Host | Out-Null
            }

            $DefaultParameterNames = 'RemoveApps','DisableTelemetry','DisableBing','DisableLockscreenTips','DisableSuggestions','ShowKnownFileExt','DisableWidgets','HideChat','DisableCopilot'

            PrintHeader 'Ĭ��ģʽ'

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
            PrintHeader "Ӧ���Ƴ�"

            $result = ShowAppSelectionForm

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Output "��ѡ�� $($global:SelectedApps.Count) �����Ƴ�Ӧ��"
                AddParameter 'RemoveAppsCustom' "�Ƴ� $($global:SelectedApps.Count) ��Ӧ�ã�"

                if (-not $Silent) {
                    Write-Output ""
                    Write-Output "���س���ʼ�Ƴ�Ӧ�ã���CTRL+C�˳�..."
                    Read-Host | Out-Null
                    PrintHeader "Ӧ���Ƴ�"
                }
            }
            else {
                Write-Host "��ȡ��������δ�Ƴ��κ�Ӧ��" -ForegroundColor Red
                Write-Output ""
            }
        }

        '4' {
            PrintHeader '�Զ���ģʽ'
            Write-Output "����Ӧ���������ã�"

            Foreach ($line in (Get-Content -Path "$PSScriptRoot/SavedSettings" )) { 
                $line = $line.Trim()
            
                if (-not ($line.IndexOf('#') -eq -1)) {
                    $parameterName = $line.Substring(0, $line.IndexOf('#'))

                    if ($parameterName -eq "RemoveAppsCustom") {
                        if (-not (Test-Path "$PSScriptRoot/CustomAppsList")) {
                            continue
                        }
                        
                        $appsList = ReadAppslistFromFile "$PSScriptRoot/CustomAppsList"
                        Write-Output "- �Ƴ� $($appsList.Count) ��Ӧ�ã�"
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
                Write-Output "���س���ʼִ�У���CTRL+C�˳�..."
                Read-Host | Out-Null
            }

            PrintHeader '�Զ���ģʽ'
        }
    }
}
else {
    PrintHeader '�Զ���ģʽ'
}

# �����������û��ϴ����ĵ����������
# ��� SPParams �еļ������� Params �еļ�������ͬ�����ʾ�û�δѡ�������κ��޸�
# �ű�����ֱ���˳��������κθ���
if ($SPParamCount -eq $global:Params.Keys.Count) {
    Write-Output "�ű�ִ����ϣ�δ�����κ��޸ġ�"

    AwaitKeyToExit
}
else {
    # ִ������ѡ��/�ṩ�Ĳ���
    switch ($global:Params.Keys) {
        'RemoveApps' {
            $appsList = ReadAppslistFromFile "$PSScriptRoot/Appslist.txt" 
            Write-Output "> �����Ƴ�Ĭ��ѡ��� $($appsList.Count) ��Ӧ��..."
            RemoveApps $appsList
            continue
        }
        'RemoveAppsCustom' {
            if (-not (Test-Path "$PSScriptRoot/CustomAppsList")) {
                Write-Host "> �����޷����ļ������Զ���Ӧ���б�δ�Ƴ��κ�Ӧ��" -ForegroundColor Red
                Write-Output ""
                continue
            }
            
            $appsList = ReadAppslistFromFile "$PSScriptRoot/CustomAppsList"
            Write-Output "> �����Ƴ� $($appsList.Count) ��Ӧ��..."
            RemoveApps $appsList
            continue
        }
        'RemoveCommApps' {
            Write-Output "> �����Ƴ��ʼ�������������Ӧ��..."
            
            $appsList = 'Microsoft.windowscommunicationsapps', 'Microsoft.People'
            RemoveApps $appsList
            continue
        }
        'RemoveW11Outlook' {
            $appsList = 'Microsoft.OutlookForWindows'
            Write-Output "> �����Ƴ��°� Windows Outlook Ӧ��..."
            RemoveApps $appsList
            continue
        }
        'RemoveDevApps' {
            $appsList = 'Microsoft.PowerAutomateDesktop', 'Microsoft.RemoteDesktop', 'Windows.DevHome'
            Write-Output "> �����Ƴ��������Ӧ��..."
            RemoveApps $appsList
            continue
        }
        'RemoveGamingApps' {
            $appsList = 'Microsoft.GamingApp', 'Microsoft.XboxGameOverlay', 'Microsoft.XboxGamingOverlay'
            Write-Output "> �����Ƴ���Ϸ���Ӧ��..."
            RemoveApps $appsList
            continue
        }
        "ForceRemoveEdge" {
            ForceRemoveEdge
            continue
        }
        'DisableDVR' {
            RegImport "> ���ڽ��� Xbox ��Ϸ/��Ļ¼��..." "Disable_DVR.reg"
            continue
        }
        'DisableTelemetry' {
            RegImport "> ���ڽ���ң�⡢������ݡ����ʷ��¼��Ӧ���������ٺͶ�����..." "Disable_Telemetry.reg"
            continue
        }
        {$_ -in "DisableSuggestions", "DisableWindowsSuggestions"} {
            RegImport "> ���ڽ��� Windows �еļ�����ʾ������͹��..." "Disable_Windows_Suggestions.reg"
            continue
        }
        'DisableDesktopSpotlight' {
            RegImport "> ���ڽ��� 'Windows �۽�' ���汳��ѡ��..." "Disable_Desktop_Spotlight.reg"
            continue
        }
        {$_ -in "DisableLockscrTips", "DisableLockscreenTips"} {
            RegImport "> ���ڽ������������ϵļ�����ʾ..." "Disable_Lockscreen_Tips.reg"
            continue
        }
        {$_ -in "DisableBingSearches", "DisableBing"} {
            RegImport "> ���ڽ��� Windows �����еı�Ӧ��ҳ��������Ӧ AI ��С��..." "Disable_Bing_Cortana_In_Search.reg"
            
            # ͬʱ�Ƴ���Ӧ����Ӧ�ð�
            $appsList = 'Microsoft.BingSearch'
            RemoveApps $appsList
            continue
        }
        'DisableCopilot' {
            RegImport "> ���ڽ��ò��Ƴ� Microsoft Copilot..." "Disable_Copilot.reg"

            # ͬʱ�Ƴ� Copilot Ӧ�ð�
            $appsList = 'Microsoft.Copilot'
            RemoveApps $appsList
            continue
        }
        'DisableRecall' {
            RegImport "> ���ڽ��� Windows Recall ���չ���..." "Disable_AI_Recall.reg"
            continue
        }
        'RevertContextMenu' {
            RegImport "> ���ڻָ��ɰ� Windows 10 ��ʽ�Ҽ��˵�..." "Disable_Show_More_Options_Context_Menu.reg"
            continue
        }
        'DisableMouseAcceleration' {
            RegImport "> ���ڹر����ָ�뾫����ǿ..." "Disable_Enhance_Pointer_Precision.reg"
            continue
        }
        'DisableStickyKeys' {
            RegImport "> ���ڽ���ճ�ͼ���ݼ�..." "Disable_Sticky_Keys_Shortcut.reg"
            continue
        }
        'ClearStart' {
            Write-Output "> ����Ϊ�û� $(GetUserName) �Ƴ���ʼ�˵����й̶�Ӧ��..."
            ReplaceStartMenu
            Write-Output ""
            continue
        }
        'ClearStartAllUsers' {
            ReplaceStartMenuForAllUsers
            continue
        }
        'DisableStartRecommended' {
            RegImport "> ���ڽ��ò����ؿ�ʼ�˵��Ƽ����..." "Disable_Start_Recommended.reg"
            continue
        }
        'TaskbarAlignLeft' {
            RegImport "> ���ڽ���������ť�����..." "Align_Taskbar_Left.reg"
            continue
        }
        'HideSearchTb' {
            RegImport "> ������������������ͼ��..." "Hide_Search_Taskbar.reg"
            continue
        }
        'ShowSearchIconTb' {
            RegImport "> ���ڽ���������������Ϊ����ʾͼ��..." "Show_Search_Icon.reg"
            continue
        }
        'ShowSearchLabelTb' {
            RegImport "> ���ڽ���������������Ϊͼ�������..." "Show_Search_Icon_And_Label.reg"
            continue
        }
        'ShowSearchBoxTb' {
            RegImport "> ���ڽ���������������Ϊ����������..." "Show_Search_Box.reg"
            continue
        }
        'HideTaskview' {
            RegImport "> ��������������������ͼ��ť..." "Hide_Taskview_Taskbar.reg"
            continue
        }
        {$_ -in "HideWidgets", "DisableWidgets"} {
            RegImport "> ���ڽ���С�����������������С���ͼ��..." "Disable_Widgets_Taskbar.reg"
            continue
        }
        {$_ -in "HideChat", "DisableChat"} {
            RegImport "> ������������������ͼ��..." "Disable_Chat_Taskbar.reg"
            continue
        }
        'EnableEndTask' {
            RegImport "> ���������������Ҽ��˵��е� '��������' ѡ��..." "Enable_End_Task.reg"
            continue
        }
        'ExplorerToHome' {
            RegImport "> ���ڽ��ļ���Դ������Ĭ�ϴ�λ������Ϊ `��ҳ`..." "Launch_File_Explorer_To_Home.reg"
            continue
        }
        'ExplorerToThisPC' {
            RegImport "> ���ڽ��ļ���Դ������Ĭ�ϴ�λ������Ϊ `�˵���`..." "Launch_File_Explorer_To_This_PC.reg"
            continue
        }
        'ExplorerToDownloads' {
            RegImport "> ���ڽ��ļ���Դ������Ĭ�ϴ�λ������Ϊ `����`..." "Launch_File_Explorer_To_Downloads.reg"
            continue
        }
        'ExplorerToOneDrive' {
            RegImport "> ���ڽ��ļ���Դ������Ĭ�ϴ�λ������Ϊ `OneDrive`..." "Launch_File_Explorer_To_OneDrive.reg"
            continue
        }
        'ShowHiddenFolders' {
            RegImport "> ������ʾ�����ļ����ļ��к�������..." "Show_Hidden_Folders.reg"
            continue
        }
        'ShowKnownFileExt' {
            RegImport "> ����������֪�ļ����͵���չ����ʾ..." "Show_Extensions_For_Known_File_Types.reg"
            continue
        }
        'HideHome' {
            RegImport "> ���������ļ���Դ���������������е���ҳ���..." "Hide_Home_from_Explorer.reg"
            continue
        }
        'HideGallery' {
            RegImport "> ���������ļ���Դ���������������е�ͼ����..." "Hide_Gallery_from_Explorer.reg"
            continue
        }
        'HideDupliDrive' {
            RegImport "> ���������ļ���Դ���������������е��ظ����ƶ�������..." "Hide_duplicate_removable_drives_from_navigation_pane_of_File_Explorer.reg"
            continue
        }
        {$_ -in "HideOnedrive", "DisableOnedrive"} {
            RegImport "> ���������ļ���Դ���������������е� OneDrive �ļ���..." "Hide_Onedrive_Folder.reg"
            continue
        }
        {$_ -in "Hide3dObjects", "Disable3dObjects"} {
            RegImport "> ���������ļ���Դ���������������е� 3D �����ļ���..." "Hide_3D_Objects_Folder.reg"
            continue
        }
        {$_ -in "HideMusic", "DisableMusic"} {
            RegImport "> ���������ļ���Դ���������������е������ļ���..." "Hide_Music_folder.reg"
            continue
        }
        {$_ -in "HideIncludeInLibrary", "DisableIncludeInLibrary"} {
            RegImport "> ���������Ҽ��˵��е� '����������' ѡ��..." "Disable_Include_in_library_from_context_menu.reg"
            continue
        }
        {$_ -in "HideGiveAccessTo", "DisableGiveAccessTo"} {
            RegImport "> ���������Ҽ��˵��е� '�������Ȩ��' ѡ��..." "Disable_Give_access_to_context_menu.reg"
            continue
        }
        {$_ -in "HideShare", "DisableShare"} {
            RegImport "> ���������Ҽ��˵��е� '����' ѡ��..." "Disable_Share_from_context_menu.reg"
            continue
        }
    }

    RestartExplorer

    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output "�ű�ִ����ϣ������Ϸ��Ƿ��д�����Ϣ��"

    AwaitKeyToExit
    }    
