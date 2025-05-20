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
        # Determine if any ChildItems were processed
        If ($this.ChildItems.Count -gt 0) {
            # At least one ChildItem was processed; produce the Item Detail sheet
            $this.childItems | Select-Object -Property @{Name="Name";Expression={$_.childItem.Name}},
            @{Name="Path";Expression={$_.ozoDirectorySummary.Path}},
            @{Name="Last Modified";Expression={$_.ozoDirectorySummary.newestChildWriteTime}},
            @{Name="Total size (GB)";Expression={($_.ozoDirectorySummary.totalSize / 1073741824)}},
            @{Name="Long paths";Expression={[Boolean]$_.ozoDirectorySummary.longPaths}},
            @{Name="Problem paths";Expression={[Boolean]$_.ozoDirectorySummary.problemPaths}},
            @{Name="Messages";Expression={$_.Messages -Join "; "}} | Export-Excel -WorksheetName "Item Detail" -Path $this.excelPath
            # Produce the sheet detailing all unique users found for all child items
            $this.childItems.adObjects | Where-Object {$null -ne $_.adUser} | Select-Object -Unique -Property @{Name="Last Name";Expression={$_.adUser.Surname}},
            @{Name="First Name";Expression={$_.adUser.GivenName}},
            @{Name="Account";Expression={$_.adUser.SamAccountName}},
            @{Name="Email Address";Expression={$_.adUser.EmailAddress}} | Sort-Object -Property "Last Name","First Name","Account" | Export-Excel -WorksheetName "Users" -Path $this.excelPath
            # Produce the sheet detailing all unique groups found for all items
            $this.childItems.adObjects | Where-Object {$null -ne $_.adGroup} | Select-Object -Unique -Property @{Name="Group Name";Expression={$_.adGroup.Name}} | Export-Excel -WorksheetName "Groups" -Path $this.excelPath
            # Determine if there are any invalid SIDs
            If (($this.childItems.adObjects | Where-Object {$_.Validates -eq $false}).Count -gt 0) {
                # Found invalid SIDs; produce the sheet detailing all unique invalid SIDs
                $this.childItems.adObjects | Where-Object {$_.Validates -eq $false} | Select-Object -Unique -Property @{Name="SID";Expression={$_.objectSID}},@{Name="Messasges";Expression={$_.Messages -Join "; "}} | Export-Excel -WorksheetName "Invalid Objects" -Path $this.excelPath
            }
            # Determine if there are any long paths
            If (($this.childItems.ozoDirectorySummary.longPaths).Count -gt 0) {
                # Found long paths on one or more objects; produce sheet detailing all long paths
                $this.childItems.ozoDirectorySummary.longPaths | Select-Object -Property @{Name="Path";Expression={$_.FullName}},@{Name="Path Length";Expression={$_.FullName.Length}} | Export-Excel -WorksheetName "Long Paths" -Path $this.excelPath
            }
            # Determine if there are any problem paths
            If (($this.childItems.ozoDirectorySummary.problemPaths).Count -gt 0) {
                # Found problem paths on one or more objects; produce sheet detailing all problem paths
                $this.childItems.ozoDirectorySummary.problemPaths | Select-Object -Property @{Name="Path";Expression={$_.FullName}},@{Name="Error Message";Expression={$_.ErrorMessage}} | Export-Excel -WorksheetName "Problem Paths" -Path $this.excelPath
            }
        } Else {
            # No ChildItems were processed
            $this.ozoLogger.Write("No child items were processed.","Warning")
        }
    }
}

Class OZOChildItem {
    # PROPERTIES: FileSystemInfo
    [System.IO.FileSystemInfo] $childItem = $null
    # PROPERTIES: PSCustomObjects
    [PSCustomObject] $ozoDirectorySummary = $null
    # PROPERTIES: PSCustomObject Lists
    [System.Collections.Generic.List[PSCustomObject]] $adObjects = @()
    # PROPERTIES: String Lists
    [System.Collections.Generic.List[String]]         $Messages  = @()
    # METHODS
    # Constructor method
    OZOChildItem($ChildItem) {
        # Set properties
        $this.childItem = $ChildItem
        # Get directory summary
        $this.ozoDirectorySummary = (Get-OZODirectorySummary -Path $this.childItem.FullName)
        # Iterate through the SIDs in the directory summary
        ForEach ($objectSID in $this.ozoDirectorySummary.objectSIDs) {
            # Determine if identity reference is a a domain ID
            If ($objectSID -Like "S-1-5-21*") {
                # objectSID is a domain ID; add an OZOADObject to the adObjects list
                $this.adObjects.Add(([OZOADObject]::new($objectSID)))
            }
        }
        # Iterate through the valid adObjects that contain an AD group
        ForEach ($adObject in ($this.adObjects | Where-Object {$_.Validates -eq $true -And $null -ne $_.adGroup})) {
            # Iterate through the members of the group, recusrively
            ForEach ($groupMember in (Get-ADGroupMember -Identity $adObject.objectSID -Recursive)) {
                # Try to get the SID for this user
                Try {
                    $userSID = (Get-ADUser -Ideneity $groupMember -ErrorAction Stop).SID
                    # Success; determine if adObjects does not already contain this SID
                    If ($this.adObjects.objectSID -NotContains $userSID) {
                        # adObjects does not already contain an object with this SID; add it
                        $this.adObjects.Add(([OZOADObject]::new($userSID)))
                    }
                } Catch {
                    # Failure
                    $this.Messages.Add(("While processing the " + $adObject.adGroup.Name + " [recursive] group members, unable to get SID for user with DN " + $groupMember + ". Error message is: " + $_))
                }
            }
        }
    }
}

Class OZOADObject {
    # PROPERTIES: Booleans, Strings
    [Boolean] $Validates = $true
    [String]  $objectSID = $null
    # PROPERTIES: PSCustomObjects
    [PSCustomObject] $adGroup  = $null
    [PSCustomObject] $adObject = $null
    [PSCustomObject] $adUser   = $null
    # PROPERTIES: String Lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS
    # Constructor method
    OZOADObject($ObjectSID) {
        # Set properties
        $this.objectSID = $ObjectSID
        # Try to get the AD object
        Try {
            $this.adObject = (Get-ADObject -Filter {objectSid -eq $this.objectSID} -ErrorAction Stop)
            # Success; switch on objectClass
            Switch ($this.adObject.objectClass) {
                "user" {
                    $this.Validates = $this.GetADUser()
                }
                "group" {
                    $this.Validates = $this.GetADGroup()
                }
                default {
                    $this.Messages.Add(("ObjectClass " + $this.adObject.objectClass + " is not handled"))
                    $this.Validates -eq $false
                }
            }
        } Catch {
            # Failure
            $this.Messages.Add(("Unable to get AD object. Error message is: " + $_))
        }
    }
    # Get AD User method
    Hidden [Boolean] GetADUser() {
        # Control variable
        [Boolean] $Return = $true
        # Determine that the object is valid and not null and class is User
        If ($this.Validates -eq $true -And $null -ne $this.adObject -And $this.adObject.objectClass -eq "user") {
            # Valid, not null, and user class; try to get the AD user
            Try {
                $this.adUser = (Get-ADUser -Identity $this.objectSID -Properties SamAccountName,GivenName,Surname,EmailAddress -ErrorAction Stop)
                # Success; determine if user is disabled
                If ($this.adUser.Enabled -eq $false) {
                    # User is disabled
                    $this.Messages.Add("User is disabled")
                    $Return = $false
                }
            } Catch {
                # Failure
                $this.Messages.Add(("Unable to get AD user object. Error message is: " + $_))
                $Return = $false
            }
            # Make sure group is null
            $this.adGroup = $null
        }
        # Return
        return $Return
    }
    # Get AD Group method
    Hidden [Boolean] GetADGroup() {
        # Control variable
        [Boolean] $Return = $true
        # Determine that the object is valid and not null and class is Group
        If ($this.Validates -eq $true -And $null -ne $this.adObject -And $this.adObject.objectClass -eq "group") {
            # Valid, not null, and group class; try to get the group
            Try {
                $this.adGroup = (Get-ADGroup -Identity $this.objectSID -ErrorAction Stop)
                # Success
            } Catch {
                # Failure
                $this.Messages.Add(("Unable to get AD group object. Error message is: " + $_))
                $Return = $false
            }
            # Make sure user is null
            $this.adUser = $null
        }
        # Return
        return $Return
    }
}

# Create a Main object
[OZOMain]::new($OutDir,$Path) | Out-Null
