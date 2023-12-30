 **# MS365ModuleInstall.ps1**

**## Easily install or uninstall essential Microsoft 365 PowerShell modules.**

**### Key Features:**

- **Installs or uninstalls multiple modules efficiently.**
- **Ensures current NuGet version for compatibility.**
- **Generates detailed log files for tracking actions.**
- **Moves installed modules to a common location for easy management.**
- **Supports the following modules:**
  - MicrosoftTeams
  - MSOnline
  - AzureAD
  - ExchangeOnlineManagement
  - Microsoft.Online.Sharepoint.PowerShell
  - SharePointPnPPowerShellOnline
  - ORCA
  - WhiteboardAdmin

**### Usage:**

1. **Open PowerShell as an Administrator.**

2. **To install modules:**

   ```powershell
   .\MS365ModuleInstall.ps1 -Mode Install
   ```

3. **To uninstall modules:**

   ```powershell
   .\MS365ModuleInstall.ps1 -Mode Uninstall
   ```

4. **Review the log file (optional):**

   - The log file is created at `$env:windir\Temp\Logs\MS365ModuleInstall.log`.

**### Additional Notes:**

- **Requires an active internet connection.**
- **May prompt for confirmations.**
- **Recommended to run as an administrator.**

**Author:** Warren Sherwen
**Version:** 1.0
