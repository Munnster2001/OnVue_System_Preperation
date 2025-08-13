<xaiArtifact artifact_id="7d686514-433e-4d0a-8ac6-7d1b2daec13f" artifact_version_id="96fd82e3-a752-4abd-ac6f-a52ed86cbc05" title="README.md" contentType="text/markdown">

# OnVUE System Preparation Script

## Overview
This PowerShell script prepares a Windows system for Pearson VUE OnVUE online proctoring by closing potentially interfering processes, stopping services, disabling notifications, clearing temporary files, and checking for high-risk software. It features a menu-driven interface for user-friendly operation.

## Features
- **Close Processes**: Terminates applications like browsers, communication apps, and backup tools (e.g., EaseUS Todo Backup's `agent.exe`, `edgegameassist.exe`, `trayprocess.exe`).
- **Stop Services**: Stops services such as "EaseUS Agent," cloud sync, and VPNs that may disrupt proctoring.
- **Disable Notifications**: Temporarily adjusts Windows Focus Assist settings to suppress notifications.
- **Clear Temporary Files**: Removes files in the TEMP directory to optimize system performance.
- **Check High-Risk Software**: Detects virtualization, screen recording, and remote access tools that could cause exam termination.
- **Menu-Driven**: Interactive menu to select specific tasks or run all at once.
- **Persistent Process Handling**: Special handling for stubborn processes like `agent.exe`.

## Requirements
- Windows operating system
- PowerShell 5.1 or later
- Administrator privileges (recommended for full functionality)

## Usage
1. Download the script (`OnVUE_System_Preparation.ps1`).
2. Right-click the script and select "Run with PowerShell as Administrator."
3. Choose an option from the menu (1-7):
   - 1: Close interfering processes
   - 2: Stop interfering services
   - 3: Disable Windows notifications
   - 4: Clear temporary files
   - 5: Check for high-risk software
   - 6: Run all tasks
   - 7: Exit
4. Follow on-screen prompts and recommendations before starting your exam.

## Recommendations
- Run as Administrator to ensure all services and processes are properly managed.
- Manually close any remaining applications.
- Disable antivirus real-time scanning temporarily.
- Disconnect from VPNs.
- Ensure VMware/VirtualBox and cloud sync apps (e.g., Dropbox, OneDrive) are closed.
- Verify a stable internet connection.
- Restart the computer if issues persist.

## Notes
- The script avoids self-termination by skipping the current PowerShell process.
- Some operations (e.g., stopping services) require elevated privileges.
- The script targets specific processes like `nextcloud.exe`, `msedgewebview2.exe`, and EaseUS-related processes to ensure compliance with OnVUE requirements.

## License
MIT License. See [LICENSE](LICENSE) for details.

</xaiArtifact>
