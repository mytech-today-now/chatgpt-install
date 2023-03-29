# Windows Shortcut Installer

ps-shortcut-install.ps1

## Overview

This PowerShell script displays a dialog box containing a checkbox list of PowerShell scripts imported from an external URL. The user can select one or more scripts to run, and then execute them in sequential order by clicking the "Run" button. An error message is displayed if a script fails to execute.

## Description

This script installs the Windows URL shortcut(s) on the user's desktop, by downloading the icon file and creating a shortcut to a URL defined in a JSON object. The JSON object can be loaded from a URL that can be set via the command line interface. The script checks if the shortcut(s) already exist on the desktop, start menu, and taskbar, and adds them if they are missing. The icon file is downloaded and stored in the %USERPROFILE%\myTechToday\icons directory, and assigned to the shortcut(s) if specified in the JSON object.

## Usage

To use the script, run it from the command line and specify the URL of the text file containing the list of PowerShell scripts to import:

```powershell
.\ps-shortcut-install -Url "http://example.com/scripts.txt"
```

## PARAMETERS

The URL of the JSON object containing the shortcut(s) information. The JSON object should have the following structure:

```json
{
    "Name": "ChatGPT",
    "URL": "https://somedomain.com/thingy",
    "Icon": "https://somedomain.com/icon.ico"
}
```

## Notes

**Designer**: kyle@mytech.today
**Coder**: ChatGPT-4
**Created**: 2023-03-29
**Updated**: 2023-03-29 - __added support for JSON object with multiple shortcuts loaded from a URL__

---

## Links
[GitHub Repository](https://github.com/mytech-today-now/ps-script-initiator)

---

## EXAMPLE

```bash
> ps-shortcut-install -ShortcutsUrl "https://somedomain.com/windows_shortcuts.json"
```
This example installs the specified Windows URL shortcut(s) defined in the JSON object located at https://somedomain.com/chatgpt_shortcuts.json

## NOTES

- This script requires the user to have administrative privileges to create the shortcut(s) on the desktop, start menu, and taskbar, and to create the icons directory.
- The script uses the WScript.Shell COM object to create the shortcut(s).
- The script uses the Invoke-RestMethod cmdlet to load the JSON object from the specified URL.
- The script uses the Invoke-WebRequest cmdlet to download the icon file.
- The script uses the Test-Path cmdlet to check if the icons directory and the shortcut(s) already exist.
- The script uses the Write-Verbose and Write-Error cmdlets to provide detailed information on the script execution.
