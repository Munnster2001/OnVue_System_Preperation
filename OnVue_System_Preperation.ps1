# Pearson VUE OnVUE System Preparation Script
# This script closes processes and stops services that may interfere with online proctoring
# Run as Administrator for full functionality

function Show-Menu {
    Clear-Host
    Write-Host "Pearson VUE OnVUE System Preparation Script" -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host "1. Close interfering processes"
    Write-Host "2. Stop interfering services"
    Write-Host "3. Disable Windows notifications"
    Write-Host "4. Clear temporary files"
    Write-Host "5. Check for high-risk software"
    Write-Host "6. Run all tasks"
    Write-Host "7. Exit"
    Write-Host "`nSelect an option (1-7): " -NoNewline -ForegroundColor Cyan
}

function Close-Processes {
    Write-Host "`nClosing potentially interfering processes..." -ForegroundColor Cyan
    foreach ($processName in $processesToClose) {
        try {
            $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
            if ($processes) {
                foreach ($process in $processes) {
                    if ($processName -like "powershell*" -and $process.Id -eq $PID) {
                        Write-Host "Skipping current PowerShell process (PID: $($process.Id))" -ForegroundColor Yellow
                        continue
                    }
                    Write-Host "Closing: $($process.ProcessName) (PID: $($process.Id))" -ForegroundColor Yellow
                    $process.CloseMainWindow() | Out-Null
                    Start-Sleep -Milliseconds 500
                    if (!$process.HasExited) {
                        $process.Kill()
                        Write-Host "  Force closed: $($process.ProcessName)" -ForegroundColor Red
                    }
                    # Special handling for persistent agent.exe (EaseUS Todo Backup)
                    if ($processName -eq "agent") {
                        Start-Sleep -Milliseconds 500
                        $stillRunning = Get-Process -Name "agent" -ErrorAction SilentlyContinue
                        if ($stillRunning) {
                            Write-Host "  Persistent agent.exe detected, attempting additional termination..." -ForegroundColor Yellow
                            Stop-Process -Name "agent" -Force -ErrorAction SilentlyContinue
                        }
                    }
                }
            }
        }
        catch {
            Write-Host "Could not close process: $processName - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

function Stop-Services {
    Write-Host "`nStopping potentially interfering services..." -ForegroundColor Cyan
    if ($isAdmin) {
        foreach ($serviceName in $servicesToStop) {
            try {
                $services = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
                foreach ($service in $services) {
                    if ($service.Status -eq 'Running') {
                        Write-Host "Stopping service: $($service.Name)" -ForegroundColor Yellow
                        Stop-Service -Name $service.Name -Force -ErrorAction SilentlyContinue
                        Write-Host "  Stopped: $($service.Name)" -ForegroundColor Green
                    }
                }
            }
            catch {
                Write-Host "Could not stop service: $serviceName - $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "Skipping service operations - Administrator privileges required" -ForegroundColor Yellow
    }
}

function Disable-Notifications {
    Write-Host "`nDisabling Windows notifications temporarily..." -ForegroundColor Cyan
    try {
        $registryPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount"
        if (Test-Path $registryPath) {
            Write-Host "Focus Assist settings adjusted" -ForegroundColor Green
        }
    } catch {
        Write-Host "Could not modify notification settings: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Clear-TempFiles {
    Write-Host "`nClearing temporary files..." -ForegroundColor Cyan
    try {
        Remove-Item -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Temporary files cleared" -ForegroundColor Green
    } catch {
        Write-Host "Some temporary files could not be cleared" -ForegroundColor Yellow
    }
}

function Check-HighRiskSoftware {
    Write-Host "`nChecking for virtualization and high-risk software..." -ForegroundColor Cyan
    $foundRiskyProcesses = @()
    foreach ($process in $highRiskProcesses) {
        $running = Get-Process -Name $process -ErrorAction SilentlyContinue
        if ($running) {
            $foundRiskyProcesses += $process
            Write-Host "WARNING: High-risk software detected: $process" -ForegroundColor Red
        }
    }
    if ($foundRiskyProcesses.Count -gt 0) {
        Write-Host "`nCRITICAL: The following applications MUST be completely closed:" -ForegroundColor Red
        foreach ($process in $foundRiskyProcesses) {
            Write-Host "  - $process" -ForegroundColor Red
        }
        Write-Host "These applications may cause exam termination if detected!" -ForegroundColor Red
    }
}

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Warning "This script should be run as Administrator for best results."
    Write-Host "Some operations may fail without elevated privileges." -ForegroundColor Yellow
}

# Define processes to close (common applications that may interfere)
$processesToClose = @(
    "chrome", "firefox", "edge", "msedge", "msedgewebview2", "opera", "safari", "brave", "vivaldi",
    "teams", "skype", "discord", "slack", "zoom", "webex", "gotomeeting", "bluejeans",
    "whatsapp", "telegram", "signal", "messenger",
    "spotify", "itunes", "vlc", "mediaplayer", "winamp", "foobar2000", "musicbee",
    "steam", "epicgameslauncher", "origin", "uplay", "battle.net", "gog", "bethesda", "plariumplay",
    "obs64", "obs32", "xsplit", "streamlabs", "bandicam", "camtasia", "snagit",
    "fraps", "shadowplay", "amdrelive", "nvidia*", "radeon*",
    "anydesk", "teamviewer", "remotedesktop", "vnc", "realvnc", "tightvnc", "ultravnc",
    "chrome-remote", "parsec", "splashtop",
    "notepad++", "sublimetext", "vscode", "atom", "brackets", "vim", "emacs",
    "visualstudio", "intellij", "eclipse", "netbeans", "pycharm", "webstorm",
    "dropbox", "googledrive", "onedrive", "box", "pcloud", "mega", "sync", "nextcloud",
    "vmware", "vmware-vmx", "vmware-hostd", "virtualbox", "vboxheadless", "vboxsvc",
    "hyper-v", "docker", "dockerdesktop",
    "utorrent", "bittorrent", "qbittorrent", "deluge", "transmission", "vuze",
    "wireshark", "fiddler", "charles", "burpsuite", "owasp*",
    "nordvpn", "expressvpn", "cyberghost", "openvpn", "cisco-vpn", "forticlient",
    "tailscale", "tailscaled", "zerotier", "hamachi", "tunnelbear", "windscribe",
    "ccleaner", "malwarebytes", "iobit", "glary", "wise*", "advanced*",
    "hwinfo64", "hwinfo32", "hwinfo", "cpu-z", "gpu-z", "msi-afterburner",
    "coretemp", "hwmonitor", "speccy", "aida64", "crystaldiskinfo",
    "process-explorer", "process-monitor", "autoruns", "regmon", "filemon",
    "unigetui", "winget-ui", "chocolatey*", "scoop", "ninite", "patch*my*pc",
    "windows*update*", "driver*booster", "driver*easy", "snappy*driver*",
    "driverpack*", "driver*genius", "3dp*chip", "double*driver",
    "streamdeck", "elgato*", "corsair*", "razer*", "steelseries*",
    "msi-mystic*", "asus-aura*", "gigabyte-rgb*", "asrock-polychrome*",
    "nzxt-cam", "alienware*", "dragon*center", "armory*crate",
    "winword", "excel", "powerpnt", "outlook", "onenote", "visio", "project",
    "acrobat", "acrord32", "acrobat-reader", "foxit*", "sumatra*",
    "libreoffice*", "openoffice*", "writer", "calc", "impress", "microsoft.cmdpal.ui",
    "carbonite", "backblaze", "crashplan", "acronis", "syncback", "easeus*", "agent", "edgegameassist", "trayprocess",
    "paragon*", "macrium*", "aomei*", "minitool*backup",
    "mysql", "postgresql", "mongodb", "redis", "sqlite",
    "powertoys", "wox", "launchy", "keypirinha", "everything", "listary",
    "rainmeter", "wallpaper*engine", "wallpaper64", "fences", "start10", "startallback", "translucenttb",
    "greenshot", "lightshot", "sharex", "puush", "gyazo", "flameshot",
    "pickpick", "faststone*", "irfanview", "paint.net", "gimp", "photoshop",
    "1password", "lastpass", "bitwarden", "keepass*", "dashlane", "roboform",
    "idm", "jdownloader*", "freedownloadmanager", "eagleget", "flashget",
    "audacity", "handbrake", "ffmpeg", "mediainfo", "mp3tag", "foobar*",
    "k-lite*", "potplayer", "mpc-hc", "gom*player",
    "winaerotweaker", "ultimate*tweaker", "winaero", "tweakui", "o&o*shutup*",
    "thisismywindows", "optimizer", "windows*10*tweaker", "win10*privacy*",
    "regedit", "registry*editor", "registry*workshop", "regcool", "regshot",
    "system*information*viewer", "msconfig", "task*scheduler", "event*viewer",
    "treesize", "spacesniffer", "windirstat", "filelight", "ridnacs",
    "duplicate*cleaner", "alldup", "auslogics*duplicate*", "clone*spy",
    "unlocker", "lockhunter", "whocrashed", "bluescreenview",
    "avast*", "avg*", "norton*", "mcafee*", "kaspersky*", "bitdefender*",
    "windows*defender*", "msmpeng", "antimalware*",
    "bleachbit", "cleanmaster", "disk*cleanup", "wise*care*",
    "totalcmd", "xplorer2", "freecommander", "directory*opus",
    "windows*terminal", "cmd", "conemu", "cmder", "hyper",
    "terminus", "alacritty", "kitty", "wezterm", "fluent*terminal",
    "notepad++", "sublime*", "vscode", "atom", "brackets", "vim", "emacs",
    "nano", "micro", "joe", "kate", "gedit", "mousepad",
    "msinfo32", "dxdiag", "system*information*viewer", "sisoftware*sandra",
    "pc*wizard", "everest", "hwinfo*", "cpu*id*", "gpu*caps*viewer",
    "powershell*"
)

# Define services to stop (temporary - they can be restarted later)
$servicesToStop = @(
    "TeamViewer*",
    "AnyDesk*",
    "RemoteDesktop*",
    "VNC*",
    "RealVNC*",
    "TightVNC*",
    "UltraVNC*",
    "Dropbox*",
    "DbxSvc",
    "GoogleUpdate*",
    "OneDrive*",
    "BoxSync*",
    "pCloud*",
    "MEGA*",
    "Nextcloud*",
    "VMware*",
    "VMAuthdService",
    "VMnetDHCP",
    "VMUSBArbService",
    "VirtualBox*",
    "vboxdrv",
    "HyperV*",
    "vmms",
    "vmcompute",
    "Docker*",
    "com.docker*",
    "Steam*",
    "Origin*",
    "Epic*",
    "Battle.net*",
    "Uplay*",
    "GOGGalaxy*",
    "Spotify*",
    "iTunes*",
    "PlariumPlay*",
    "OBS*",
    "XSplit*",
    "Streamlabs*",
    "Bandicam*",
    "Camtasia*",
    "NVIDIAShadowPlay*",
    "AMD*Record*",
    "NordVPN*",
    "ExpressVPN*",
    "CyberGhost*",
    "OpenVPN*",
    "Cisco*VPN*",
    "FortiClient*",
    "Tailscale*",
    "ZeroTier*",
    "Hamachi*",
    "TunnelBear*",
    "Windscribe*",
    "Docker*",
    "MSSQL*",
    "MySQL*",
    "Apache*",
    "nginx*",
    "IIS*",
    "Tomcat*",
    "Carbonite*",
    "Backblaze*",
    "CrashPlan*",
    "Acronis*",
    "EaseUS*",
    "Ease*Todo*",
    "EaseUS Agent",
    "Paragon*",
    "Macrium*",
    "AOMEI*",
    "MiniTool*Backup",
    "Malwarebytes*",
    "CCleaner*",
    "IObit*",
    "UniGetUI*",
    "WinGet*",
    "Chocolatey*",
    "Scoop*",
    "Ninite*",
    "PatchMyPC*",
    "DriverBooster*",
    "DriverEasy*",
    "SnappyDriver*",
    "DriverPack*",
    "3DP*",
    "DoubleDriver*",
    "Elgato*",
    "StreamDeck*",
    "Corsair*",
    "Logi*",
    "Razer*",
    "SteelSeries*",
    "MSI*",
    "ASUS*",
    "Gigabyte*",
    "ASRock*",
    "NZXT*",
    "Alienware*",
    "HWiNFO*",
    "CPUID*",
    "MSI Afterburner*",
    "Core Temp*",
    "AIDA64*",
    "Microsoft Office*",
    "Office*",
    "Adobe*",
    "Acrobat*",
    "PowerToys*",
    "Rainmeter*",
    "Wallpaper Engine*",
    "Everything*",
    "Listary*",
    "1Password*",
    "LastPass*",
    "Bitwarden*",
    "KeePass*",
    "Dashlane*",
    "IDM*",
    "JDownloader*",
    "FDM*"
)

# Define high-risk processes
$highRiskProcesses = @(
    "vmware", "vmware-vmx", "virtualbox", "vboxheadless", "hyper-v", "docker",
    "obs64", "obs32", "bandicam", "camtasia", "snagit", "fraps",
    "teamviewer", "anydesk", "vnc", "chrome-remote", "parsec",
    "wireshark", "fiddler", "charles", "burpsuite"
)

# Main menu loop
while ($true) {
    Show-Menu
    $choice = Read-Host
    switch ($choice) {
        "1" { Close-Processes }
        "2" { Stop-Services }
        "3" { Disable-Notifications }
        "4" { Clear-TempFiles }
        "5" { Check-HighRiskSoftware }
        "6" { 
            Close-Processes
            Stop-Services
            Disable-Notifications
            Clear-TempFiles
            Check-HighRiskSoftware
        }
        "7" { 
            Write-Host "`nSystem preparation completed!" -ForegroundColor Green
            Write-Host "`nRecommendations before starting your exam:" -ForegroundColor White
            Write-Host "1. Close any remaining applications manually" -ForegroundColor White
            Write-Host "2. Disable antivirus real-time scanning temporarily" -ForegroundColor White
            Write-Host "3. Disconnect from VPN if connected" -ForegroundColor White
            Write-Host "4. Ensure VMware/VirtualBox are completely shut down" -ForegroundColor White
            Write-Host "5. Close all cloud sync applications (Dropbox, OneDrive, etc.)" -ForegroundColor White
            Write-Host "6. Ensure stable internet connection" -ForegroundColor White
            Write-Host "7. Close this PowerShell window" -ForegroundColor White
            Write-Host "8. Restart your computer if issues persist" -ForegroundColor White
            Write-Host "`nExiting..." -ForegroundColor Gray
            exit
        }
        default { Write-Host "Invalid option, please select 1-7" -ForegroundColor Red }
    }
    Write-Host "`nPress Enter to continue..." -ForegroundColor Gray
    Read-Host
}
