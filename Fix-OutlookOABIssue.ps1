# This PowerShell script is designed to fix the Outlook Offline Address Book (OAB) download issue
# (error 0x8004010F). It will clear existing OAB files and force Outlook to download a fresh OAB.
# Optionally, it can also repair the Outlook profile if needed.



# Function to clear existing OAB files
function Clear-OABFiles {
    Write-Host "Clearing existing OAB files..."
    $oabPath = "$env:LOCALAPPDATA\Microsoft\Outlook"
    if (Test-Path $oabPath) {
        Remove-Item "$oabPath\*.oab" -Force -Recurse
        Write-Host "OAB files cleared."
    } else {
        Write-Host "OAB path not found."
    }
}

# Function to force a new OAB download in Outlook
function Download-OAB {
    Write-Host "Forcing Outlook to download the Offline Address Book..."
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $syncObjects = $namespace.SyncObjects

    foreach ($syncObject in $syncObjects) {
        if ($syncObject.Name -like "*Offline Address Book*") {
            $syncObject.Start
            Write-Host "Forced OAB download started."
            return
        }
    }

    Write-Host "No OAB sync object found in Outlook."
}

# Function to repair the Outlook profile (optional)
function Repair-OutlookProfile {
    Write-Host "Repairing Outlook profile..."
    # This assumes that the user is using the default profile; modify as needed.
    $profileName = "Outlook"
    $profilePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$profileName"

    if (Test-Path $profilePath) {
        Write-Host "Repairing existing profile..."
        New-ItemProperty -Path $profilePath -Name "Repair" -Value 1
        Write-Host "Profile repair initiated."
    } else {
        Write-Host "Profile not found. Skipping repair."
    }
}

# Run the automation steps
Clear-OABFiles
Download-OAB

# Optional: Uncomment the next line if you want to try repairing the Outlook profile
# Repair-OutlookProfile

Write-Host "Process completed. Restart Outlook and check if the issue is resolved."
