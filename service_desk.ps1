#===========================================================================
# Service Desk Tool
#
# .Synopsis
# Performs niche assistance type tasks for Service Desk
#
# .NOTES   
# Name       : service_desk
# Author     : Matthew Kalnins
# Email      : servicedesk@matthewkalnins.com
# Version    : 0.0.2.0
#===========================================================================

#===========================================================================
# Global Variable Declaration
#===========================================================================

$Invocation = (Get-Variable MyInvocation).Value
$global:DirectoryPath = Split-Path $Invocation.MyCommand.Path
$PSM = $DirectoryPath + "\service_desk.psm1"
$DEBUG = $true

#===========================================================================
# Initialization
#===========================================================================

# In case this has run before, remove service desk module and add it again
try {Remove-Module service_desk} catch {}
Import-Module $PSM

# Initialize the Windows Presentation Framework
Add-Type -AssemblyName PresentationFramework

# Form initialization
$inputXML = (Get-Content ($DirectoryPath + “\MainWindow.xaml”))
[xml]$XAML = $inputXML 
$reader=(New-Object System.Xml.XmlNodeReader $XAML)
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
 
$XAML.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
 
Function Get-FormVariables{
    write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
    get-variable WPF*
}
 
Get-FormVariables

#===========================================================================
# Load Components
#===========================================================================

if ($DEBUG){ 
	$WPFComputerName.Text = $env:COMPUTERNAME
}

# Set logo source dynamically
$WPFLogo.Source = $DirectoryPath + "\images\logo.png"
$Form.Icon = $DirectoryPath + "\images\favicon.ico"

#===========================================================================
# Listeners
#===========================================================================

#===========================================================================
# Visibility Buttons @TODO write a less hacky way of visibility management
#===========================================================================
$WPFAboutButton.Add_Click({ 
	$WPFAssetGrid.Visibility = $WPFComputerGrid.Visibility = $WPFUserGrid.Visibility = $WPFDataGrid.Visibility = $WPFNoResultsGrid.Visibility ='Hidden'
	$WPFAboutGrid.Visibility = 'Visible'
})

$WPFComputerButton.Add_Click({ 
    $WPFAboutGrid.Visibility = $WPFAssetGrid.Visibility = $WPFUserGrid.Visibility = $WPFDataGrid.Visibility = $WPFNoResultsGrid.Visibility ='Hidden'
	$WPFComputerGrid.Visibility = 'Visible'    
})

$WPFUserButton.Add_Click({ 
    $WPFAboutGrid.Visibility = $WPFComputerGrid.Visibility = $WPFAssetGrid.Visibility = $WPFDataGrid.Visibility = $WPFNoResultsGrid.Visibility ='Hidden'
	$WPFUserGrid.Visibility = 'Visible'
})

$WPFAssetButton.Add_Click({ 	
    $WPFAboutGrid.Visibility = $WPFComputerGrid.Visibility = $WPFUserGrid.Visibility = $WPFDataGrid.Visibility = $WPFNoResultsGrid.Visibility ='Hidden'
	$WPFAssetGrid.Visibility = 'Visible'
})

#===========================================================================
# Function Buttons
#===========================================================================

$WPFDrivesButton.Add_Click({
    Update-DataGrid -Name $WPFComputerName.Text -GetItem 0 -DataGrid $WPFDataGrid -NoResultsGrid $WPFNoResultsGrid
})

$WPFPrintersButton.Add_Click({ 
    Update-DataGrid -Name $WPFComputerName.Text -GetItem 1 -DataGrid $WPFDataGrid -NoResultsGrid $WPFNoResultsGrid
})

$WPFProgramsButton.Add_Click({    
    Update-DataGrid -Name $WPFComputerName.Text -GetItem 2 -DataGrid $WPFDataGrid -NoResultsGrid $WPFNoResultsGrid
})

$WPFComputerDetailsButton.Add_Click({    
	Update-DataGrid -Name $WPFComputerName.Text -GetItem 3 -DataGrid $WPFDataGrid -NoResultsGrid $WPFNoResultsGrid
})

$WPFAssetDetailsButton.Add_Click({
    Update-DataGrid -Name $WPFAssetName.Text -GetItem 100 -DataGrid $WPFDataGrid -NoResultsGrid $WPFNoResultsGrid
})

$WPFUserDetailsButton.Add_Click({
    Update-DataGrid -Name $WPFUserName.Text -GetItem 200 -DataGrid $WPFDataGrid -NoResultsGrid $WPFNoResultsGrid
})

if($DEBUG){
	$WPFControlTab.Visibility = "Visible"

	$WPFPingButton.Add_Click({
		try{
			$Ping = (Test-Connection -ComputerName $WPFComputerName.Text -Quiet)
			[System.Windows.MessageBox]::Show('Device is online: ' + $Ping,'Ping Results','Ok','Information')
		} catch {
			[System.Windows.MessageBox]::Show('Unable to Ping device. Likely not connected.','Error','Ok','Error')
		}	
	})

	$WPFRestartButton.Add_Click({
		$msgBoxInput =  [System.Windows.MessageBox]::Show('Are you sure?','Attempting To Reboot ' + $WPFComputerName.Text,'YesNo','Warning')

		  switch  ($msgBoxInput) {
			  'Yes' { Restart-Computer -ComputerName $WPFComputerName.Text -Force  }
			}    
	})
}

#===========================================================================
# Clear Buttons
#===========================================================================

$WPFComputerClearButton.Add_Click({
    Clear-Grid -TextBox $WPFComputerName -DataGrid $WPFDataGrid -NoResultsGrid $WPFNoResultsGrid
})

$WPFAssetClearButton.Add_Click({
	Clear-Grid -TextBox $WPFAssetName -DataGrid $WPFDataGrid -NoResultsGrid $WPFNoResultsGrid
})

$WPFUserClearButton.Add_Click({
	Clear-Grid -TextBox $WPFUserName -DataGrid $WPFDataGrid -NoResultsGrid $WPFNoResultsGrid
})

$WPFComputerTabControl.Add_SelectionChanged({    
	Clear-DataGrid -DataGrid $WPFDataGrid -NoResultsGrid $WPFNoResultsGrid
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null