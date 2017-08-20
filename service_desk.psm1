#===========================================================================
# Service Desk Modules
# 
# .Synopsis
# Requirement to run service_desk.ps1
#
#===========================================================================

Function Convert-ListToDataGrid () {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, HelpMessage="List of values to bind to DataGrid",ValueFromPipeline=$true)][AllowEmptyCollection()][System.Collections.ArrayList] $List,
		[string[]] $SortParameter,
		[System.Windows.Controls.DataGrid] $DataGrid
    )
 
    PROCESS {
        try { $DataGrid.ItemsSource = $List | Sort-Object $SortParameter } 
        catch { $DataGrid.ItemsSource = @(New-Object PSCustomObject -Property @{}) }
    }
}

Function Convert-ObjectToRows () {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)][AllowEmptyCollection()][PSCustomObject] $Object
    )

	BEGIN {
		[System.Collections.ArrayList]$List = @()
		$HashProperty = @{}
	}
 
    PROCESS {
        $Object.PSObject.Properties | foreach-object {
			$HashProperty.Value = $_.Value
			$HashProperty.Name = $_.Name         
			$Item = New-Object -TypeName PSCustomObject -Property $HashProperty
			$List.Add($Item)
		}
    }

	END {
		return $List
	}
}

Function Update-DataGrid () {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [string[]] $Name,
        [int] $GetItem,
		[System.Windows.Controls.DataGrid] $DataGrid,
		[System.Windows.Controls.Grid] $NoResultsGrid
    )

	BEGIN{
		Clear-DataGrid -DataGrid $DataGrid
		$SortParameter = "Name"
		[System.Collections.ArrayList]$Items = @()
	}

    PROCESS {		
		# Realistically there needs to be a better way to handle the "GetItem" selector as it doesn't scale well
		# Computer Methods; 0 = Drives, 1 = Printers, 2 = Installed Programs, 3 = Details
		# Asset Methods; 100 = Details
		# User Methods; 200 = Details
        if ($GetItem -lt 2) { 
			$Items = Get-RemoteItems -ComputerName $Name -GetItem $GetItem
			$SortParameter = "UserName"
		}
        elseif ($GetItem -eq 2) {             
            Get-RemoteProgram -ComputerName $Name -ExcludeSimilar -SimilarWord 4 -Property Version, DisplayVersion | ForEach-Object {    
                # Try get the Version attribute (not every software has this) otherwise show DisplayVersion (i.e. ProductVersion) TODO Workaround; Incorporate into actual method
                $Items.Add((New-Object PSCustomObject -Property @{ProgramName = $_.ProgramName; DisplayVersion = $(if ($_.Version){ $_.Version } Else {$_.DisplayVersion});}))
			}	
			$SortParameter = "ProgramName"
		}
		elseif ($GetItem -eq 3) {
			$Items = (Convert-ObjectToRows -Object (Get-ComputerDetails -ComputerName $Name))
		}   		
		elseif ($GetItem -eq 100) {
			$csv = Import-Csv ($DirectoryPath + "\data\equipment.csv")
			$Items = (Convert-ObjectToRows -Object ($csv | where {$_.name -eq $Name}))
		} 
		elseif ($GetItem -eq 200) {
			# TODO
		}

		# Output the list of items obtained through the above methods
		if ($Items) { Convert-ListToDataGrid -List $Items -SortParameter $SortParameter -DataGrid $DataGrid}
    }

	END {			
		if ($Items) { $DataGrid.Visibility = 'Visible' }
		else { $NoResultsGrid.Visibility = 'Visible' }
	}
}

Function Clear-Grid () { 
	[CmdletBinding()]
    PARAM (
        [System.Windows.Controls.TextBox] $TextBox,
        [System.Windows.Controls.DataGrid] $DataGrid,
        [System.Windows.Controls.Grid] $NoResultsGrid
    )

    PROCESS {
		$TextBox.Text = ""
        Clear-DataGrid -DataGrid $DataGrid -NoResultsGrid $NoResultsGrid
    }
}

Function Clear-DataGrid () { 
	[CmdletBinding()]
    PARAM (
        [System.Windows.Controls.DataGrid] $DataGrid,
		[System.Windows.Controls.Grid] $NoResultsGrid
    )
    PROCESS {
        $DataGrid.ItemsSource = @(New-Object PSCustomObject -Property @{})
		$DataGrid.Visibility = 'Hidden' 
		try { $NoResultsGrid.Visibility = 'Hidden' } catch { }
    }
}

Function Clear-DataGrids () {
    [CmdletBinding()]
    PARAM (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [System.Windows.Controls.DataGrid[]] $DataGrids
    )
 
    PROCESS {
        foreach ($DataGrid in $DataGrids) { $DataGrid.ItemsSource = @(New-Object PSCustomObject -Property @{}) }
    }
}

Function Get-ComputerDetails {
    [CmdletBinding(SupportsShouldProcess=$true)]
    PARAM(
        [Parameter(ValueFromPipeline              =$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0
        )]
        [string[]] $ComputerName
    )

    PROCESS {
        foreach ($Computer in $ComputerName) {
			try {
				$computerSystem = Get-WmiObject Win32_ComputerSystem -Computer $Computer
				$computerBIOS = Get-WmiObject Win32_BIOS -Computer $Computer
				$computerOS = Get-WmiObject Win32_OperatingSystem -Computer $Computer
				$computerCPU = Get-WmiObject Win32_Processor -Computer $Computer
				#$computerHDD = Get-WmiObject Win32_LogicalDisk -ComputerName $Computer -Filter drivetype=3
				
				# Collate the useful information into an object
				$Item = New-Object PSCustomObject -Property @{
					Domain = $computerSystem.Domain
					Manufacturer = $computerSystem.Manufacturer 
					Model = $computerSystem.Model 
					Serial = $computerBIOS.SerialNumber
					CPU = $computerCPU.Name 
					OS = $computerOS.caption + ", Service Pack: " + $computerOS.ServicePackMajorVersion
					UserName = $computerSystem.UserName
					LastReboot = $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
				}			
				
				return $Item
			}
			catch { 
				return 
			}
        }
    }
}

Function Get-RemoteItems {
    [CmdletBinding(SupportsShouldProcess=$true)]
    PARAM(
        [Parameter(ValueFromPipeline              =$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0
        )]
        [string[]] $ComputerName,
        [Parameter(Position=0)]
        [int] $GetItem
    )

    BEGIN {
        $HashProperty = @{}
    }

    PROCESS {
        foreach ($Computer in $ComputerName) {
			try {$RegBase = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('Users', $Computer)}
			catch { 
				write-host No reg
				return
			}
            
            $UserKeys = $RegBase.GetSubKeyNames()
            
            $UserKeys | ForEach-Object {
                $CurrentKey = $_                
                
                if ($CurrentKey) {
                    $VolKey = $RegBase.OpenSubKey("$($CurrentKey)\\Volatile Environment")                    

                    # Probably a better way of handling types... switch or bool?
                    if ($GetItem -le 0){
                        $CurrentRegKey = $RegBase.OpenSubKey("$($CurrentKey)\\Network")
                    }
                    else {
                        $CurrentRegKey = $RegBase.OpenSubKey("$($CurrentKey)\\Printers\\Connections")
                    }

                    if ($CurrentRegKey) {
                        if ($VolKey) {
                            $UserName = $VolKey.GetValue("UserDomain") + "\" + $VolKey.GetValue("UserName")
                            $HashProperty.UserName = $UserName
                        }

                        $CurrentRegKey.GetSubKeyNames() | ForEach-Object {     
                            if ($GetItem -le 0){
                                $HashProperty.DriveLetter = $_
                                $HashProperty.DrivePath = $RegBase.OpenSubKey("$($CurrentKey)\\Network\\$($_)").GetValue('RemotePath')
                            }
                            else {
                                $HashProperty.PrinterName = $_
                                $HashProperty.PrinterServer = $RegBase.OpenSubKey("$($CurrentKey)\\Printers\\Connections\\$($_)").GetValue('Server')
                            }
                            
                            New-Object -TypeName PSCustomObject -Property $HashProperty
                        }
                    }
                }
            } | ForEach-Object -Begin {
                [System.Collections.ArrayList]$Array = @()
            } -Process {
                $null = $Array.Add($_)
            } -End {
                return $Array 
            }            
        }
    }
}

Function Get-RemoteProgram {
<#
.Synopsis
Generates a list of installed programs on a computer

.DESCRIPTION
This function generates a list by querying the registry and returning the installed programs of a local or remote computer.

.NOTES   
Name       : Get-RemoteProgram
Author     : Jaap Brasser
Version    : 1.3
DateCreated: 2013-08-23
DateUpdated: 2016-08-26
Blog       : http://www.jaapbrasser.com

.LINK
http://www.jaapbrasser.com

.PARAMETER ComputerName
The computer to which connectivity will be checked

.PARAMETER Property
Additional values to be loaded from the registry. Can contain a string or an array of string that will be attempted to retrieve from the registry for each program entry

.PARAMETER ExcludeSimilar
This will filter out similar programnames, the default value is to filter on the first 3 words in a program name. If a program only consists of less words it is excluded and it will not be filtered. For example if you Visual Studio 2015 installed it will list all the components individually, using -ExcludeSimilar will only display the first entry.

.PARAMETER SimilarWord
This parameter only works when ExcludeSimilar is specified, it changes the default of first 3 words to any desired value.

.EXAMPLE
Get-RemoteProgram

Description:
Will generate a list of installed programs on local machine

.EXAMPLE
Get-RemoteProgram -ComputerName server01,server02

Description:
Will generate a list of installed programs on server01 and server02

.EXAMPLE
Get-RemoteProgram -ComputerName Server01 -Property DisplayVersion,VersionMajor

Description:
Will gather the list of programs from Server01 and attempts to retrieve the displayversion and versionmajor subkeys from the registry for each installed program

.EXAMPLE
'server01','server02' | Get-RemoteProgram -Property Uninstallstring

Description
Will retrieve the installed programs on server01/02 that are passed on to the function through the pipeline and also retrieves the uninstall string for each program

.EXAMPLE
'server01','server02' | Get-RemoteProgram -Property Uninstallstring -ExcludeSimilar -SimilarWord 4

Description
Will retrieve the installed programs on server01/02 that are passed on to the function through the pipeline and also retrieves the uninstall string for each program. Will only display a single entry of a program of which the first four words are identical.
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(ValueFromPipeline              =$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0
        )]
        [string[]]
            $ComputerName,
        [Parameter(Position=0)]
        [string[]]
            $Property,
        [switch]
            $ExcludeSimilar,
        [int]
            $SimilarWord
    )

    begin {
        $RegistryLocation = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\',
                            'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
        $HashProperty = @{}
        $SelectProperty = @('ProgramName','ComputerName')
        if ($Property) {
            $SelectProperty += $Property
        }
    }

    process {
        foreach ($Computer in $ComputerName) {
            $RegBase = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$Computer)
            $RegistryLocation | ForEach-Object {
                $CurrentReg = $_
                if ($RegBase) {
                    $CurrentRegKey = $RegBase.OpenSubKey($CurrentReg)
                    if ($CurrentRegKey) {
                        $CurrentRegKey.GetSubKeyNames() | ForEach-Object {
                            if ($Property) {
                                foreach ($CurrentProperty in $Property) {
                                    $HashProperty.$CurrentProperty = ($RegBase.OpenSubKey("$CurrentReg$_")).GetValue($CurrentProperty)
                                }
                            }
                            $HashProperty.ComputerName = $Computer
                            $HashProperty.ProgramName = ($DisplayName = ($RegBase.OpenSubKey("$CurrentReg$_")).GetValue('DisplayName'))
                            $HashProperty.Version = ($Version = ($RegBase.OpenSubKey("$CurrentReg$_")).GetValue('Version'))
                            if ($DisplayName -Or $Version) {
                                New-Object -TypeName PSCustomObject -Property $HashProperty |
                                Select-Object -Property $SelectProperty
                            }
                        }
                    }
                }
            } | ForEach-Object -Begin {
                if ($SimilarWord) {
                    $Regex = [regex]"(^(.+?\s){$SimilarWord}).*$|(.*)"
                } else {
                    $Regex = [regex]"(^(.+?\s){3}).*$|(.*)"
                }
                [System.Collections.ArrayList]$Array = @()
            } -Process {
                if ($ExcludeSimilar) {
                    $null = $Array.Add($_)
                } else {
                    $_
                }
            } -End {
                if ($ExcludeSimilar) {
                    $Array | Select-Object -Property *,@{
                        name       = 'GroupedName'
                        expression = {
                            ($_.ProgramName -split $Regex)[1]
                        }
                    } |
                    Group-Object -Property 'GroupedName' | ForEach-Object {
                        $_.Group[0] | Select-Object -Property * -ExcludeProperty GroupedName
                    }
                }
            }
        }
    }
}
