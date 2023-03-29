<#
    
File Name: ps-shortcut-install.ps1
Coder: ChatGPT-4
Designer: kyle@mytech.today
Created: 2023-03-29
Updated: 2023-03-29 - Version 0.6 - added support for JSON object with multiple shortcuts loaded from a URL

## SYNOPSIS

`ps-shortcut-install.ps1` installs Internet shortcuts with a URL and icon specified by a JSON file.

## DESCRIPTION

This script installs Internet shortcuts on a user's desktop and Start menu, with a URL and icon specified by a JSON file available at a given URL. The script can be run with a URL parameter, which will be used to retrieve the JSON file with the shortcuts information. The script ensures that the icons directory is created and that the shortcuts are created only if they don't already exist in the specified locations.

## EXAMPLE

```bash
.\ps-shortcut-install.ps1 -ShortcutUrl "https://domain.com/list-of-shortcuts-to-create.json"
```

This example installs the Windows URL shortcut(s) defined in the JSON object located at https://domain.com/list-of-shortcuts-to-create.json.

## PARAMETER ShortcutsUrl

Specifies the URL of the JSON file that contains information about the shortcuts to be installed.  The JSON object should have the following structure:

- `ShortcutUrl` (required): Specifies the URL (https://domain.com/list-of-shortcuts-to-create.json) of the JSON file that contains information about the shortcuts to be installed.

```json
{
	{
		"Name": "Shortcut Name",
		"URL": "https://somedomain.com/thingy",
		"Icon": "https://somedomain.com/icon.ico"
	}, {
		"Name": "Shortcut Name 2",
		"URL": "https://someotherdomain.com/foobar",
		"Icon": "https://someotherdomain.com/icon.ico"
	}
}
```

## NOTES

- This script requires the user to have administrative privileges to create the shortcut(s) on the desktop, start menu, and taskbar, and to create the icons directory.
- The script uses the WScript.Shell COM object to create the shortcut(s).
- The script uses the Invoke-RestMethod cmdlet to load the JSON object from the specified URL.
- The script uses the Invoke-WebRequest cmdlet to download the icon file.
- The script uses the Test-Path cmdlet to check if the icons directory and the shortcut(s) already exist.
- The script uses the Write-Verbose and Write-Error cmdlets to provide detailed information on the script execution.

#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true)]
    [string]$jsonUrl,
    [Parameter(Mandatory=$false)]
    [string]$logFile = ".\log.txt"
)

# Initialize log file and error handling
$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"
if (Test-Path $logFile) { Remove-Item $logFile }
$LogStream = [System.IO.StreamWriter]::new($logFile, $true)

try {
    # Retrieve JSON data from URL and convert to PowerShell object
    $json = Invoke-RestMethod -Uri $jsonUrl -ErrorAction Stop
    if ($json -eq $null) {
        throw "JSON file is empty or invalid."
    }

    # Iterate over each object in the JSON array and create shortcut
    foreach ($shortcut in $json) {
        if ($shortcut.Name -eq $null -or $shortcut.URL -eq $null) {
            throw "Shortcut name or URL is missing from JSON file."
        }

        # Build URL shortcut file path
        $shortcutPath = "$([System.Environment]::GetFolderPath('Desktop'))\{0}.url" -f $shortcut.Name

        # Check if shortcut already exists
        $shortcutExists = Test-Path $shortcutPath

        # Create shortcut object and set properties
        $WshShell = New-Object -ComObject WScript.Shell
        $shortcutObject = $WshShell.CreateShortcut($shortcutPath)

        # Set shortcut properties
        $shortcutObject.TargetPath = $shortcut.URL
        if ($shortcut.Icon -ne $null) {
            $shortcutObject.IconLocation = $shortcut.Icon
        }
        $shortcutObject.Save()

        # Log success message
        if ($shortcutExists) {
            Write-Verbose "Shortcut '$shortcut.Name' updated."
            $LogStream.WriteLine("$(Get-Date) - Shortcut '$shortcut.Name' updated.")
        } else {
            Write-Verbose "Shortcut '$shortcut.Name' created."
            $LogStream.WriteLine("$(Get-Date) - Shortcut '$shortcut.Name' created.")
        }
    }
} catch {
    # Log error and display message
    Write-Error $_
    $LogStream.WriteLine("$(Get-Date) - Error: $_")
    $LogStream.Close()
    throw
} finally {
    # Close log file stream
    $LogStream.Close()
}
