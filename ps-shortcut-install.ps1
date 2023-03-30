<#
    File Name: windows_install.ps1
    Coder: ChatGPT-4
    Designer: kyle@mytech.today
    Created: 2023-03-29
    Updated: 2023-03-29 - Version 0.6 - added missing Name and URL property checks for shortcuts in JSON object
                                        added support for JSON object with multiple shortcuts loaded from a URL
#>

<#  Example usage 
    
    ./windows_install.ps1 ShortcutsUrl "URL"
#>

function Get-Shortcuts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ShortcutsUrl
    )

    try {
        return Invoke-RestMethod -Uri $ShortcutsUrl
    }
    catch {
        Write-Error "Failed to load shortcuts JSON: $_"
        return $null
    }
}

function Test-ShortcutProperties {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [PSCustomObject]$Shortcut
    )

    process {
        if (-not ($Shortcut.Name) -or -not ($Shortcut.URL)) {
            Write-Error "Shortcut name or URL is missing from JSON file"
            return $false
        }
        return $true
    }
}

function Save-Icon {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$IconUrl,
        [Parameter(Mandatory=$true)]
        [string]$IconPath
    )

    try {
        Invoke-WebRequest -Uri $IconUrl -OutFile $IconPath -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to download icon file: $_"
        return $false
    }
    return $true
}

function New-WindowsShortcut {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [Parameter(Mandatory=$true)]
        [string]$Url,
        [Parameter(Mandatory=$true)]
        [string]$IconPath
    )

    $TargetPath = (Get-Command "iexplore.exe").Path

    try {
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\$Name.lnk")
        $Shortcut.TargetPath = $TargetPath
        $Shortcut.Arguments = "-k $Url"
        $Shortcut.IconLocation = "$IconPath,0"
        $Shortcut.Save()
        Write-Verbose "Shortcut created on desktop: $($Shortcut.FullName)"
    }
    catch {
        Write-Error "Failed to create shortcut on desktop: $_"
        return
    }
}

function Install-WindowsShortcuts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ShortcutsUrl
    )

    $IconsDirectory = "$env:USERPROFILE\myTechToday\icons"
    if (!(Test-Path $IconsDirectory)) {
        New-Item -ItemType Directory -Path $IconsDirectory -ErrorAction Stop | Out-Null
    }

    Get-Shortcuts -ShortcutsUrl $ShortcutsUrl |
    Where-Object { Test-ShortcutProperties -Shortcut $_ } |
    ForEach-Object {
        $IconPath = "$IconsDirectory\$($_.Name).ico"
        if (-not (Test-Path $IconPath)) {
            Save-Icon -IconUrl $_.Icon -IconPath $IconPath
        }
        New-WindowsShortcut -Name $_.Name -Url $_.URL -IconPath $IconPath
    }
}
