
# Fix Outlook OAB Download Issue

This PowerShell script helps resolve the Outlook Offline Address Book (OAB) download error (0x8004010F). The script clears existing OAB files and forces Outlook to download a fresh OAB. It also includes an optional feature to repair the Outlook profile.

## Features

- **Clear Existing OAB Files**: Automatically deletes outdated OAB files from the local Outlook data folder.
- **Force OAB Download**: Uses Outlook's COM interface to trigger a fresh download of the Offline Address Book.
- **Outlook Profile Repair**: Optionally repairs the Outlook profile to fix potential configuration issues.

## Usage

1. Clone or download the repository.
2. Open the PowerShell script (`Fix-OutlookOABIssue.ps1`) in an editor to review and customize if necessary.
3. Run the script with administrator privileges to clear OAB files and force a download.

### Running the Script

```bash
# Example of running the script in PowerShell
.\Fix-OutlookOABIssue.ps1
```

If you want to enable the Outlook profile repair, uncomment the corresponding line in the script:

```powershell
# Uncomment the following line to repair the Outlook profile
# Repair-OutlookProfile
```

## Disclaimer

This script is provided as-is without any warranty. Use it at your own risk. Ensure you back up important data before running the script.

## Author

**aviado1**  
[https://github.com/aviado1](https://github.com/aviado1)
