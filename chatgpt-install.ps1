<#
    
File Name: chatgpt_install.ps1
Coder: ChatGPT-4
Designer: kyle@mytech.today
Created: 2023-03-29
Updated: 2023-03-29 - Version 0.5 - added support for JSON object with multiple shortcuts loaded from a URL

.SYNOPSIS
Installs the ChatGPT URL shortcut(s) on the user's desktop, with options to add an icon, by downloading the icon file and creating a shortcut to a URL defined in a JSON object.

.DESCRIPTION
This script installs the ChatGPT URL shortcut(s) on the user's desktop, by downloading the icon file and creating a shortcut to a URL defined in a JSON object. The JSON object can be loaded from a URL that can be set via the command line interface. The script checks if the shortcut(s) already exist on the desktop, start menu, and taskbar, and adds them if they are missing. The icon file is downloaded and stored in the %USERPROFILE%\myTechToday\icons directory, and assigned to the shortcut(s) if specified in the JSON object.

.PARAMETER ShortcutsUrl
The URL of the JSON object containing the shortcut(s) information. The JSON object should have the following structure:
{
    "Name": "ChatGPT",
    "URL": "https://somedomain.com/thingy",
    "Icon": "https://somedomain.com/icon.ico"
}

.EXAMPLE
Install-ChatGPTShortcut -ShortcutsUrl "https://somedomain.com/chatgpt_shortcuts.json"
This example installs the ChatGPT URL shortcut(s) defined in the JSON object located at https://somedomain.com/chatgpt_shortcuts.json.

.NOTES
- This script requires the user to have administrative privileges to create the shortcut(s) on the desktop, start menu, and taskbar, and to create the icons directory.
- The script uses the WScript.Shell COM object to create the shortcut(s).
- The script uses the Invoke-RestMethod cmdlet to load the JSON object from the specified URL.
- The script uses the Invoke-WebRequest cmdlet to download the icon file.
- The script uses the Test-Path cmdlet to check if the icons directory and the shortcut(s) already exist.
- The script uses the Write-Verbose and Write-Error cmdlets to provide detailed information on the script execution.

#>

function New-Shortcut {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetPath,
        [Parameter(Mandatory = $false)]
        [string]$Arguments,
        [Parameter(Mandatory = $false)]
        [string]$WorkingDirectory,
        [Parameter(Mandatory = $false)]
        [string]$IconLocation,
        [Parameter(Mandatory = $false)]
        [int]$IconIndex
    )

    try {
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\$Name.lnk")
        $Shortcut.TargetPath = $TargetPath
        if ($Arguments) {
            $Shortcut.Arguments = $Arguments
        }
        if ($WorkingDirectory) {
            $Shortcut.WorkingDirectory = $WorkingDirectory
        }
        if ($IconLocation) {
            $Shortcut.IconLocation = "$IconLocation,$IconIndex"
        }
        $Shortcut.Save()
        Write-Verbose "Shortcut created on desktop: $($Shortcut.FullName)"
    }
    catch {
        Write-Error "Failed to create shortcut on desktop: $_"
        return
    }
}

function Install-ChatGPTShortcut {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ShortcutsUrl
    )

    try {
        # Load the shortcuts JSON object from the URL
        $shortcuts = Invoke-RestMethod -Uri $ShortcutsUrl

        # Create the icons directory if it doesn't exist
        $iconsDirectory = "$env:USERPROFILE\myTechToday\icons"
        if (!(Test-Path $iconsDirectory)) {
            New-Item -ItemType Directory -Path $iconsDirectory -ErrorAction Stop | Out-Null
        }

        # Process each shortcut in the JSON object
        foreach ($shortcut in $shortcuts) {
            $name = $shortcut.Name
            $url = $shortcut.URL
            $iconUrl = $shortcut.Icon

            # Define the local icon path
            $iconPath = "$iconsDirectory\$name.ico"

            # Download the icon file
            if (!(Test-Path $iconPath)) {
                try {
                    Invoke-WebRequest -Uri $iconUrl -OutFile $iconPath -ErrorAction Stop
                }
                catch {
                    Write-Error "Failed to download icon file for shortcut '$name': $_"
                    continue
                }
            }

            # Check if shortcut exists on desktop, start menu, and taskbar
            $desktopShortcutPath = "$env
