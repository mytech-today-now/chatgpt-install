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
param (
    [Parameter(Mandatory=$true)]
    [string]$ShortcutUrl
)

function New-Shortcut {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [Parameter(Mandatory=$true)]
        [string]$TargetPath,
        [string]$Arguments = "",
        [string]$WorkingDirectory = "",
        [string]$IconLocation = "",
        [int]$IconIndex = 0
    )

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut("$([Environment]::GetFolderPath('Desktop'))\$Name.lnk")
    $shortcut.TargetPath = $TargetPath
    $shortcut.Arguments = $Arguments
    $shortcut.WorkingDirectory = $WorkingDirectory
    $shortcut.IconLocation = $IconLocation
    $shortcut.IconIndex = $IconIndex
    $shortcut.Save()
}

function Install-Shortcut {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ShortcutUrl
    )

    try {
        $shortcutJson = Invoke-RestMethod -Uri $ShortcutUrl -ErrorAction Stop
    } catch {
        Write-Error "Failed to load shortcut JSON from $ShortcutUrl. $_"
        return
    }

    $iconsDirectory = "$([Environment]::GetFolderPath('MyDocuments'))\icons"
    if (-not (Test-Path $iconsDirectory)) {
        try {
            New-Item -Path $iconsDirectory -ItemType Directory -ErrorAction Stop | Out-Null
        } catch {
            Write-Error "Failed to create directory $iconsDirectory. $_"
            return
        }
    }

    foreach ($shortcut in $shortcutJson) {
        $name = $shortcut.Name
        $targetPath = $shortcut.Url
        $iconLocation = "$iconsDirectory\$name.ico"

        if (-not (Test-Path "$([Environment]::GetFolderPath('Desktop'))\$name.lnk")) {
            try {
                Invoke-WebRequest -Uri $shortcut.Icon -OutFile $iconLocation -ErrorAction Stop
                New-Shortcut -Name $name -TargetPath $targetPath -IconLocation $iconLocation -IconIndex 0
            } catch {
                Write-Error "Failed to create shortcut '$name' on desktop. $_"
            }
        }

        if (-not (Test-Path "$([Environment]::GetFolderPath('Programs'))\$name.lnk")) {
            try {
                Invoke-WebRequest -Uri $shortcut.Icon -OutFile $iconLocation -ErrorAction Stop
                New-Shortcut -Name $name -TargetPath $targetPath -IconLocation $iconLocation -IconIndex 0 | Out-Null
                $startMenuDirectory = "$([Environment]::GetFolderPath('Programs'))\ChatGPT"
                if (-not (Test-Path $startMenuDirectory)) {
                    New-Item -Path $startMenuDirectory -ItemType Directory -ErrorAction Stop | Out-Null
                }
                Move-Item "$([Environment]::GetFolderPath('Programs'))\$name.lnk" "$startMenuDirectory\$name.lnk" -ErrorAction Stop | Out-Null
            } catch {
                Write-Error "Failed to create shortcut '$name' in start menu. $_"
            }
        }
    }
}

Install-Shortcut -ShortcutUrl $ShortcutUrl
