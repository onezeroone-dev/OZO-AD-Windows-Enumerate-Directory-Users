#Requires -Modules ActiveDirectory,ImportExcel,@{ModuleName="OZO";ModuleVersion="1.6.0"},OZOFiles,OZOLogger -Version 5.1

<#PSScriptInfo
    .VERSION 1.0.0
    .GUID 7465a5f8-752f-4b68-a91c-b681bc639e81
    .AUTHOR Andy Lievertz <alievertz@onezeroone.dev>
    .COMPANYNAME One Zero One
    .COPYRIGHT This script is released under the terms of the GNU General Public License ("GPL") version 2.0.
    .TAGS 
    .LICENSEURI https://github.com/onezeroone-dev/OZO-AD-Windows-Enumrate-Directory-Users/blob/main/LICENSE
    .PROJECTURI https://github.com/onezeroone-dev/OZO-AD-Windows-Enumrate-Directory-Users
    .ICONURI 
    .EXTERNALMODULEDEPENDENCIES ActiveDirectory,ImportExcel
    .REQUIREDSCRIPTS 
    .EXTERNALSCRIPTDEPENDENCIES 
    .RELEASENOTES https://github.com/onezeroone-dev/OZO-AD-Windows-Enumrate-Directory-Users/blob/main/CHANGELOG.md
#>

<# 
    .SYNOPSIS
    See description.
    .DESCRIPTION 
    Enumerates the directories and files in a given path and produces an Excel report containing users, email addresses, groups, and last modified child date.
    .PARAMETER OutDir
    Directory for the Excel report. Defaults to the current directory.
    .PARAMETER Path
    The path to inspect. Defaults to the current directory.
    .LINK
    https://github.com/onezeroone-dev/OZO-AD-Windows-Enumrate-Directory-Users/blob/main/README.md
#>

# PARAMETERS
[CmdletBinding(SupportsShouldProcess = $true)] Param (
    [Parameter(Mandatory=$false,HelpMessage="Path for the Excel report")][String]$OutDir = (Get-Location),
    [Parameter(Mandatory=$false,HelpMessage="Path to inspect")][String]$Path = (Get-Location)
)

# CLASSES
Class OZOMain {
    # PROPERTIES: Booleans, Strings
    [Boolean] $Validates = $true
    [String]  $excelPath = $null
    [String]  $outDir    = $null
    [String]  $Path      = $null
    # PROPERTIES: PSCustomObjects
    [PSCustomObject] $ozoLogger = $null
    # PROPERTIES: Lists
    [System.Collections.Generic.List[PSCustomObject]] $childItems = @()
    # METHODS
    # Constructor method
    OZOMain($OutDir,$Path) {
        # Set properties
        $this.outDir = $OutDir
        $this.Path   = $Path
        # Create a logger
        $this.ozoLogger = (New-OZOLogger)
        # Log a process start message
        $this.ozoLogger.Write("Starting process.","Information")
        # Determine if the configuation is valid
        If (($this.ValidateEnvironment()) -eq $true) {
            # Call GetItems to generate OZOItem objects
            $this.GetChildItems()
        }
        # Report
        $this.Report()
        # Log a process end message
        $this.ozoLogger.Write(("Process complete."),"Information")
    }
    # Environment validation method
    Hidden [Boolean] ValidateEnvironment() {
        # Control variable
        [Boolean] $Return = $true
        # Determine if the outDir exists
        If ([Boolean](Test-Path -Path $this.outDir) -eq $true) {
            # Output directory exists; set the Excel path
            $this.excelPath = (Join-Path -Path $this.outDir -ChildPath ((Get-OZO8601Date -Time) + "-ozo-ad-windows-directory-users.xlsx"))
        } Else {
            # Output directory does not exist; report
            $this.ozoLogger.Write(("Output directory is invalid or inaccessible."),"Error")
            $Return = $false
        }
        # Return
        return $Return
    }
    # GetItems method
    Hidden [Void] GetChildItems() {
        # Iterate through the child items
        ForEach ($childItem in (Get-ChildItem -Path $this.Path)) {
            # Add an OZOitem object to the Items list
            $this.childItems.Add(([OZOChildItem]::new($childItem)))
        }
    }
    # Report method
    Hidden [Void] Report() {
        # Determine if any Items were created
        If ($this.Items.Count -gt 0) {
            # At least one Item was processed; output selected object properties to Excel
            $this.Items | Select-Object -Property itemName,Validates,@{Name="Messages";Expression={$_.Messages -Join "; "}} | Export-Excel -WorksheetName $this.Json.ExcelWorksheetName -Path $this.excelPath
            $this.Logger.Write(("See " + $this.excelPath + " for results."),"Information")
        }
    }
}

Class OZOChildItem {
    # PROPERTIES: Booleans, Strings
    [Boolean] $Validates     = $true
    [String]  $childItemName = $null
    # PROPERTIES: Lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS
    # Constructor method
    OZOItem($ChildItem) {
        # Set properties
        $this.childItemName = $ChildItem
    }
}

# Create a Main object
[OZOMain]::new($OutDir,$Path) | Out-Null
