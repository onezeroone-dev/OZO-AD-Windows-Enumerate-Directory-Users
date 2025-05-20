# OZO AD Windows Enumerate Directory Users Installation and Usage
## Description
Enumerates the directories and files in a given path and produces an Excel report containing users, email addresses, groups, and last modified child date. This information can be useful in determining which folders are being actively used and by whom.

The resulting workbook contains six tabs:

|Tab|Description|
|---|-----------|
|Item&Detail|This tab provides a detailed relationship between the top-level subdirectory and the users and groups found within that particular folder; and also shows the total size and timestamp for the most recently modified child.|
|Users|A unique list of all of the users who have access to one or more of the files within the `Path`.|
|Groups|A unique list of all of the groups who are assigned to one or more of the files and folders within the `Path`.|
|Invalid&nbsp;SIDs|A list of SIDs that represent a disabled user; or a SID that cannot be resolved by the current domain controller (likely indicating a SID from a legacy domain).|
|Long&nbsp;Paths|A list of paths containing more than the number of characters specified with `LongLength`. Defaults to 256.|
|Problem&nbsp;Paths|A list of paths that generated an error. This may be due to invalid characters, a long path, or inadequate permissions. See the `ErrorMessage` column for more detail.|

## Prerequisites
This script requires the _ActiveDirectory_, _ImportExcel_, _OZO_, _OZOFiles_, and _OZOLogger_ PowerShell modules. The _ActiveDirectory_ PowerShell module is included with the [_Remote Server Administration Tools (RSAT) for Windows_](https://learn.microsoft.com/en-us/troubleshoot/windows-server/system-management-components/remote-server-administration-tools) feature installation. The remaining modules are published to [PowerShell Gallery](https://learn.microsoft.com/en-us/powershell/scripting/gallery/overview?view=powershell-5.1). Ensure your system is configured for this repository then execute the following in an _Administrator_ PowerShell:

```powershell
Install-Module ImportExcel,OZO,OZOFiles,OZOLogger
```

## Installation
This script is published to [PowerShell Gallery](https://learn.microsoft.com/en-us/powershell/scripting/gallery/overview?view=powershell-5.1). Ensure your system is configured for this repository then execute the following in an _Administrator_ PowerShell:

```powershell
Install-Script ozo-ad-windows-enumerate-directory-users
```

## Usage
```powershell
ozo-ad-windows-enumerate-directory-users
    -OutDir <String>
    -Path   <String>
```

## Parameters
|Parameter|Description|
|---------|-----------|
|`OutDir`|Directory for the Excel report. Defaults to the current directory.|
|`Path`|The path to inspect. Defaults to the current directory.|

## Acknowledgements
Special thanks to my employer, [Sonic Healthcare USA](https://sonichealthcareusa.com), who supports the growth of my PowerShell skillset and enables me to contribute portions of my work product to the PowerShell community.
