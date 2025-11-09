# Windows Script "AO Niii"
# Run as Administrator for full functionality

function Show-Header {
    Write-Host "`n" + ("=" * 70) -ForegroundColor Cyan
    Write-Host "                 WINDOWS SCRIPT ""AO Niii""" -ForegroundColor White
    Write-Host ("=" * 70) -ForegroundColor Cyan
}

function Test-AdminRights {
    return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Show-Menu {
    param($SelectedApps, $SelectedSettings, $StartMenuSelected, $CurrentSelection)
    
    Clear-Host
    Show-Header
    
    # Admin status
    if (Test-AdminRights) {
        Write-Host "Status: Running with administrator rights" -ForegroundColor Green
    } else {
        Write-Host "Status: Run as administrator for full functionality" -ForegroundColor White
    }
    
    Write-Host "`n1. APPLICATIONS TO REMOVE:" -ForegroundColor White
    for ($i = 0; $i -lt $AppList.Count; $i++) {
        $status = if ($SelectedApps.Contains($i)) { "[X]" } else { "[ ]" }
        $color = if ($SelectedApps.Contains($i)) { "Green" } else { "Gray" }
        $selectionMarker = if ($CurrentSelection -eq $i) { ">> " } else { "   " }
        Write-Host "   $selectionMarker$status $($AppList[$i].Name)" -ForegroundColor $color
    }
    
    Write-Host "`n2. WINDOWS SETTINGS:" -ForegroundColor White
    $settingsStartIndex = $AppList.Count
    for ($i = 0; $i -lt $RegistrySettings.Count; $i++) {
        $globalIndex = $settingsStartIndex + $i
        $status = if ($SelectedSettings.Contains($i)) { "[X]" } else { "[ ]" }
        $color = if ($SelectedSettings.Contains($i)) { "Green" } else { "Gray" }
        $selectionMarker = if ($CurrentSelection -eq $globalIndex) { ">> " } else { "   " }
        Write-Host "   $selectionMarker$status $($RegistrySettings[$i].Name)" -ForegroundColor $color
    }
    
    Write-Host "`n3. START MENU CLEANUP:" -ForegroundColor White
    $startMenuIndex = $AppList.Count + $RegistrySettings.Count
    $status = if ($StartMenuSelected) { "[X]" } else { "[ ]" }
    $color = if ($StartMenuSelected) { "Green" } else { "Gray" }
    $selectionMarker = if ($CurrentSelection -eq $startMenuIndex) { ">> " } else { "   " }
    Write-Host "   $selectionMarker$status ULTIMATE Start Menu Cleanup" -ForegroundColor $color
    Write-Host "      Removes ALL tiles: Edge, Office, Entertainment, Tools" -ForegroundColor Gray
    
    Write-Host "`n" + ("-" * 70) -ForegroundColor Cyan
    
    # Control hints
    Write-Host "SPACE - mark/unmark" -ForegroundColor White
    Write-Host "ENTER - start execution" -ForegroundColor White
}

# Application list for removal
$AppList = @(
    @{Name="Microsoft Store"; Patterns=@("*WindowsStore*", "*StorePurchaseApp*")},
    @{Name="Cortana"; Patterns=@("*Cortana*", "*Microsoft.549981C3F5F10*")},
    @{Name="Feedback Hub"; Patterns=@("*FeedbackHub*")},
    @{Name="Tips"; Patterns=@("*Microsoft.GetHelp*", "*GetStarted*")},
    @{Name="PC Health Check"; Patterns=@("*PCHealthCheck*")},
    @{Name="Support"; Patterns=@("*Microsoft.WindowsSupport*")},
    @{Name="Mail and Calendar"; Patterns=@("*windowscommunicationsapps*")},
    @{Name="Your Phone"; Patterns=@("*YourPhone*")},
    @{Name="Skype"; Patterns=@("*Skype*")},
    @{Name="People"; Patterns=@("*People*")},
    @{Name="Weather"; Patterns=@("*BingWeather*")},
    @{Name="News"; Patterns=@("*BingNews*")},
    @{Name="Maps"; Patterns=@("*WindowsMaps*")},
    @{Name="Camera"; Patterns=@("*WindowsCamera*")},
    @{Name="Voice Recorder"; Patterns=@("*SoundRecorder*")},
    @{Name="Movies & TV"; Patterns=@("*ZuneVideo*", "*Microsoft.MoviesTV*")},
    @{Name="Groove Music"; Patterns=@("*ZuneMusic*", "*Microsoft.MediaPlayer*")},
    @{Name="Photos"; Patterns=@("*Photos*", "*Microsoft.Windows.Photos*")},
    @{Name="Clipchamp"; Patterns=@("*Clipchamp*")},
    @{Name="Sticky Notes"; Patterns=@("*StickyNotes*")},
    @{Name="Microsoft To Do"; Patterns=@("*Todo*")},
    @{Name="Calculator"; Patterns=@("*Calculator*")},
    @{Name="Alarms & Clock"; Patterns=@("*Alarms*", "*WindowsAlarms*")},
    @{Name="Snip & Sketch"; Patterns=@("*ScreenSketch*", "*Snip*")},
    @{Name="Paint 3D"; Patterns=@("*Paint*", "*Microsoft.MSPaint*")},
    @{Name="Photoshop Express"; Patterns=@("*Photoshop*")},
    @{Name="3D Viewer"; Patterns=@("*3DViewer*")},
    @{Name="Microsoft Solitaire Collection"; Patterns=@("*Solitaire*")},
    @{Name="Xbox and Gaming Services"; Patterns=@("*Xbox*", "*XboxGamingOverlay*")},
    @{Name="Mixed Reality Portal"; Patterns=@("*MixedReality*", "*HoloLens*")},
    @{Name="Office (UWP)"; Patterns=@("*Office*", "*OneNote*")}
)

# Registry settings
$RegistrySettings = @(
    @{Name="Change search icon"; Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"; Key="SearchboxTaskbarMode"; Value=1; Type="DWORD"},
    @{Name="Hide Cortana button"; Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Key="ShowCortanaButton"; Value=0; Type="DWORD"},
    @{Name="Hide Task View button"; Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Key="ShowTaskViewButton"; Value=0; Type="DWORD"},
    @{Name="Remove favorites from taskbar"; Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"; Key="Favorites"; Action="Delete"}
)

# Functions
function Remove-OneDrive {
    Write-Host "Removing OneDrive..." -ForegroundColor White
    
    # Stop OneDrive processes
    Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue | Stop-Process -Force
    
    # Remove via installer
    $setupPaths = @(
        "$env:SystemRoot\SysWOW64\OneDriveSetup.exe",
        "$env:SystemRoot\System32\OneDriveSetup.exe"
    )
    
    foreach ($path in $setupPaths) {
        if (Test-Path $path) {
            Write-Host "  Found OneDrive at $path" -ForegroundColor Cyan
            Start-Process -FilePath $path -ArgumentList "/uninstall" -Wait -WindowStyle Hidden
        }
    }
    
    Write-Host "  OneDrive removed" -ForegroundColor Green
}

function Remove-EdgeShortcut {
    $desktopPath = [Environment]::GetFolderPath("CommonDesktopDirectory")
    $edgeShortcut = Join-Path $desktopPath "Microsoft Edge.lnk"
    
    if (Test-Path $edgeShortcut) {
        Remove-Item $edgeShortcut -Force
        Write-Host "  Microsoft Edge shortcut removed" -ForegroundColor Green
    } else {
        Write-Host "  Microsoft Edge shortcut not found" -ForegroundColor Gray
    }
}

function Ultimate-StartMenuCleanup {
    Write-Host "`n" + ("=" * 60) -ForegroundColor Cyan
    Write-Host "STARTING ULTIMATE START MENU CLEANUP" -ForegroundColor White
    Write-Host ("=" * 60) -ForegroundColor Cyan
    
    # 1. Nuke the entire CloudStore cache
    Write-Host "`n1. NUKING CloudStore cache..." -ForegroundColor White
    $cloudStorePath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore"
    
    if (Test-Path $cloudStorePath) {
        try {
            # Stop ShellExperienceHost and StartMenuExperienceHost first
            Stop-Process -Name "ShellExperienceHost" -Force -ErrorAction SilentlyContinue
            Stop-Process -Name "StartMenuExperienceHost" -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            
            # Take ownership recursively
            Write-Host "   Taking ownership..." -NoNewline
            Start-Process -FilePath "takeown" -ArgumentList "/f `"$cloudStorePath`" /r /d y" -Wait -WindowStyle Hidden
            Start-Process -FilePath "icacls" -ArgumentList "`"$cloudStorePath`" /grant administrators:F /t" -Wait -WindowStyle Hidden
            Write-Host " [DONE]" -ForegroundColor Green
            
            # Remove entire CloudStore
            Write-Host "   Removing CloudStore..." -NoNewline
            Remove-Item $cloudStorePath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host " [DONE]" -ForegroundColor Green
        } catch {
            Write-Host " [ERROR]" -ForegroundColor Red
        }
    }

    # 2. Remove all Start Menu related registry entries
    Write-Host "`n2. CLEANING Start Menu registry..." -ForegroundColor White
    $registryPaths = @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage2", 
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
    )
    
    foreach ($path in $registryPaths) {
        if (Test-Path $path) {
            Write-Host "   Cleaning $($path.Split('\')[-1])..." -NoNewline
            try {
                Remove-ItemProperty -Path $path -Name "Favorites" -ErrorAction SilentlyContinue -Force
                Remove-ItemProperty -Path $path -Name "FavoritesResolve" -ErrorAction SilentlyContinue -Force
                Remove-ItemProperty -Path $path -Name "Start_Layout" -ErrorAction SilentlyContinue -Force
                
                # For Taskband - remove Edge and Office pins
                if ($path -like "*Taskband*") {
                    Remove-ItemProperty -Path $path -Name "Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge" -ErrorAction SilentlyContinue -Force
                    Remove-ItemProperty -Path $path -Name "Microsoft.Office.WINWORD.EXE" -ErrorAction SilentlyContinue -Force
                    Remove-ItemProperty -Path $path -Name "Microsoft.Office.EXCEL.EXE" -ErrorAction SilentlyContinue -Force
                    Remove-ItemProperty -Path $path -Name "Microsoft.Office.POWERPOINT.EXE" -ErrorAction SilentlyContinue -Force
                }
                Write-Host " [DONE]" -ForegroundColor Green
            } catch {
                Write-Host " [ERROR]" -ForegroundColor Red
            }
        }
    }

    # 3. Delete all Start Menu layout files
    Write-Host "`n3. DELETING layout files..." -ForegroundColor White
    $layoutDir = "$env:LOCALAPPDATA\Microsoft\Windows\Shell"
    
    if (Test-Path $layoutDir) {
        Write-Host "   Cleaning Shell folder..." -NoNewline
        try {
            Get-ChildItem $layoutDir -Recurse -Include "*layout*", "*default*", "*start*" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            
            # Specifically target LayoutModification files
            $layoutFiles = @("LayoutModification.xml", "DefaultLayouts.xml", "StartLayout.xml")
            foreach ($file in $layoutFiles) {
                $filePath = Join-Path $layoutDir $file
                if (Test-Path $filePath) {
                    Remove-Item $filePath -Force -ErrorAction SilentlyContinue
                }
            }
            Write-Host " [DONE]" -ForegroundColor Green
        } catch {
            Write-Host " [ERROR]" -ForegroundColor Red
        }
    }

    # 4. Create EMPTY forced layout
    Write-Host "`n4. APPLYING EMPTY layout..." -ForegroundColor White
    $forcedLayoutPath = "$env:LOCALAPPDATA\Microsoft\Windows\Shell\LayoutModification.xml"
    
    $emptyLayout = @'
<LayoutModificationTemplate xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification" xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout" xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout" Version="1">
<LayoutOptions StartTileGroupCellWidth="6" StartTileGroupsColumnCount="1" />
<DefaultLayoutOverride>
    <StartLayoutCollection>
        <defaultlayout:StartLayout GroupCellWidth="6" />
    </StartLayoutCollection>
</DefaultLayoutOverride>
</LayoutModificationTemplate>
'@

    try {
        # Ensure directory exists
        $layoutDir = Split-Path $forcedLayoutPath -Parent
        if (!(Test-Path $layoutDir)) {
            New-Item -ItemType Directory -Path $layoutDir -Force
        }
        
        # Write empty layout
        $emptyLayout | Out-File -FilePath $forcedLayoutPath -Encoding UTF8 -Force
        Write-Host "   Empty layout applied [DONE]" -ForegroundColor Green
    } catch {
        Write-Host "   [ERROR applying layout]" -ForegroundColor Red
    }

    # 5. Restart Explorer to apply changes
    Write-Host "`n5. RESTARTING EXPLORER..." -ForegroundColor White
    
    # Stop Explorer and related processes
    $processes = @("ShellExperienceHost", "StartMenuExperienceHost", "SearchUI", "explorer")
    
    foreach ($process in $processes) {
        try {
            Stop-Process -Name $process -Force -ErrorAction SilentlyContinue
            Write-Host "   Stopped $process" -ForegroundColor Green
        } catch { }
    }
    
    # Wait a moment
    Start-Sleep -Seconds 2
    
    # Restart Explorer
    try {
        Start-Process "explorer.exe"
        Write-Host "   Explorer restarted [DONE]" -ForegroundColor Green
    } catch {
        Write-Host "   [ERROR restarting Explorer]" -ForegroundColor Red
    }
    
    Write-Host "`n" + ("=" * 60) -ForegroundColor Cyan
    Write-Host "ULTIMATE START MENU CLEANUP COMPLETED!" -ForegroundColor Green
    Write-Host ("=" * 60) -ForegroundColor Cyan
}

function Remove-AppxApplication {
    param($App)
    
    Write-Host "Removing $($App.Name)..." -NoNewline
    
    $success = $true
    foreach ($pattern in $App.Patterns) {
        try {
            $packages = Get-AppxPackage $pattern -ErrorAction SilentlyContinue
            if ($packages) {
                $packages | Remove-AppxPackage -ErrorAction SilentlyContinue | Out-Null
                Write-Host " [REMOVED]" -ForegroundColor Green
                $success = $true
            } else {
                Write-Host " [NOT FOUND]" -ForegroundColor Gray
            }
        } catch {
            $success = $false
            Write-Host " [ERROR]" -ForegroundColor Red
        }
    }
}

function Apply-RegistrySetting {
    param($Setting)
    
    Write-Host "Applying: $($Setting.Name)..." -NoNewline
    
    try {
        if (-not (Test-Path $Setting.Path)) {
            New-Item -Path $Setting.Path -Force | Out-Null
        }
        
        if ($Setting.Action -eq "Delete") {
            Remove-ItemProperty -Path $Setting.Path -Name $Setting.Key -ErrorAction SilentlyContinue -Force
        } else {
            Set-ItemProperty -Path $Setting.Path -Name $Setting.Key -Value $Setting.Value -Type $Setting.Type -Force
        }
        
        Write-Host " [OK]" -ForegroundColor Green
    } catch {
        Write-Host " [ERROR]" -ForegroundColor Red
    }
}

function Start-Execution {
    param($SelectedApps, $SelectedSettings, $StartMenuSelected)
    
    if (($SelectedApps.Count -eq 0) -and ($SelectedSettings.Count -eq 0) -and (-not $StartMenuSelected)) {
        Write-Host "No actions selected!" -ForegroundColor Red
        pause
        return $false
    }
    
    # Confirmation
    Clear-Host
    Show-Header
    
    Write-Host "ACTIONS TO EXECUTE:" -ForegroundColor White
    Write-Host "Applications: $($SelectedApps.Count)" -ForegroundColor White
    Write-Host "Settings: $($SelectedSettings.Count)" -ForegroundColor White
    if ($StartMenuSelected) {
        Write-Host "Start Menu: ULTIMATE Cleanup" -ForegroundColor Green
    } else {
        Write-Host "Start Menu: None" -ForegroundColor Gray
    }
    
    Write-Host "`nProceed with execution? (Y/N)" -ForegroundColor White -NoNewline
    $confirmation = Read-Host " "
    
    if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
        return $false
    }
    
    # Create log file
    $logPath = Join-Path $env:TEMP "windows_cleanup_log.txt"
    "=== Execution started ===" | Out-File $logPath
    "Date: $(Get-Date)" | Out-File $logPath -Append
    "" | Out-File $logPath -Append
    
    Clear-Host
    Show-Header
    Write-Host "EXECUTING SELECTED ACTIONS..." -ForegroundColor Green
    Write-Host "Log file: $logPath" -ForegroundColor Cyan
    Write-Host ""
    
    # Execute app actions
    if ($SelectedApps.Count -gt 0) {
        Write-Host "REMOVING APPLICATIONS:" -ForegroundColor White
        foreach ($index in $SelectedApps) {
            $app = $AppList[$index]
            "Processing: $($app.Name)" | Out-File $logPath -Append
            Remove-AppxApplication -App $app
            "$($app.Name) - completed" | Out-File $logPath -Append
            Start-Sleep -Milliseconds 300
        }
        Write-Host ""
    }
    
    # Apply registry settings
    if ($SelectedSettings.Count -gt 0) {
        Write-Host "APPLYING SETTINGS:" -ForegroundColor White
        foreach ($index in $SelectedSettings) {
            $setting = $RegistrySettings[$index]
            "Setting: $($setting.Name)" | Out-File $logPath -Append
            Apply-RegistrySetting -Setting $setting
            "$($setting.Name) - completed" | Out-File $logPath -Append
            Start-Sleep -Milliseconds 200
        }
        Write-Host ""
    }
    
    # Execute Start Menu cleanup
    if ($StartMenuSelected) {
        "Start Menu: Ultimate Cleanup" | Out-File $logPath -Append
        Ultimate-StartMenuCleanup
        "Start Menu - completed" | Out-File $logPath -Append
    }
    
    # Completion message
    Write-Host "`n" + ("=" * 70) -ForegroundColor Cyan
    Write-Host "EXECUTION COMPLETED SUCCESSFULLY!" -ForegroundColor Green
    Write-Host ("=" * 70) -ForegroundColor Cyan
    Write-Host "Log saved to: $logPath" -ForegroundColor White
    
    "`n=== Execution completed ===" | Out-File $logPath -Append
    
    Write-Host "`nPress any key to close..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    
    # Close PowerShell window
    Stop-Process -Id $PID
}

# Main program loop with arrow key navigation
$SelectedApps = 0..($AppList.Count-1)  # Select all by default
$SelectedSettings = 0..($RegistrySettings.Count-1)  # Select all by default
$StartMenuSelected = $true

$totalItems = $AppList.Count + $RegistrySettings.Count + 1  # +1 for Start Menu only (no actions)
$currentSelection = 0

do {
    Show-Menu -SelectedApps $SelectedApps -SelectedSettings $SelectedSettings -StartMenuSelected $StartMenuSelected -CurrentSelection $currentSelection
    
    $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    
    switch ($key.VirtualKeyCode) {
        38 { # Up arrow
            $currentSelection = if ($currentSelection -eq 0) { $totalItems - 1 } else { $currentSelection - 1 }
        }
        40 { # Down arrow
            $currentSelection = if ($currentSelection -eq $totalItems - 1) { 0 } else { $currentSelection + 1 }
        }
        32 { # Space bar - toggle selection
            if ($currentSelection -lt $AppList.Count) {
                # Toggle app selection
                if ($SelectedApps.Contains($currentSelection)) {
                    $SelectedApps = $SelectedApps | Where-Object { $_ -ne $currentSelection }
                } else {
                    $SelectedApps += $currentSelection
                    $SelectedApps = $SelectedApps | Sort-Object
                }
            } elseif ($currentSelection -lt ($AppList.Count + $RegistrySettings.Count)) {
                # Toggle setting selection
                $settingIndex = $currentSelection - $AppList.Count
                if ($SelectedSettings.Contains($settingIndex)) {
                    $SelectedSettings = $SelectedSettings | Where-Object { $_ -ne $settingIndex }
                } else {
                    $SelectedSettings += $settingIndex
                    $SelectedSettings = $SelectedSettings | Sort-Object
                }
            } elseif ($currentSelection -eq ($AppList.Count + $RegistrySettings.Count)) {
                # Toggle Start Menu cleanup
                $StartMenuSelected = -not $StartMenuSelected
            }
        }
        13 { # Enter key - start execution
            $executed = Start-Execution -SelectedApps $SelectedApps -SelectedSettings $SelectedSettings -StartMenuSelected $StartMenuSelected
            if (-not $executed) {
                # If execution was cancelled, stay in the menu
                continue
            }
        }
    }
} while ($true)