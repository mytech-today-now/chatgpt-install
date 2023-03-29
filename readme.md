# PowerShell Script Initiator
base-install.ps1

## Overview

This PowerShell script displays a dialog box containing a checkbox list of PowerShell scripts imported from an external URL. The user can select one or more scripts to run, and then execute them in sequential order by clicking the "Run" button. An error message is displayed if a script fails to execute.

## Usage

To use the script, run it from the command line and specify the URL of the text file containing the list of PowerShell scripts to import:

```powershell
.\base-install.ps1 -Url "http://example.com/scripts.txt"
```

## Parameters

Url: Specifies the URL of the text file containing the list of PowerShell scripts to import.

## Notes

**Designer**: kyle@mytech.today
**Coder**: ChatGPT-4
**Created**: 2023-03-29
**Updated**: 2023-04-05 - __Added progress bar, improved error handling, and accessibility__

## Links
[GitHub Repository](https://github.com/mytech-today-now/ps-script-initiator)

---