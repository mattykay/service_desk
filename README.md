# service_desk

## About

This tool is a GUI based script that retrieves Windows device information through Powershell commands (see screenshots below). It was created in my own time and on my own device whilst on an IT service desk role as part of a graduate rotation inside of a large enterprise.

This tool is intended as an example of allowing non-technical users ability to obtain information that otherwise may be unclear how to obtain or require a combination of tools - it is suggested to become familiar with Windows Administration tools instead of using this for production purposes.

## Usage

Simply run the service_desk.ps1 file to initiate the GUI. (A desktop shortcut link was created in real application to hide the Powershell window when in use but has not been included in this repo).

Enter a device asset number you have the rights to connect to in the top right and press the button of the action you wish to do. (Device can be name if within the same OU, or FQDN if within another unit. E.g. "computer1" or "computer2.domain.com")

"Asset" tab allows looking up a input host to retrieve data from a known CSV (i.e. CMDB dump) - an example /data/equipment.csv is provided.

"User" tab added, however disabled as no underlying code (challenge is need selector and logic to handle local or domain account lookups).

## Release Information

- Customer and company information have been removed
- Branding has been replaced with uglier stock images.

## Screenshots

### Device Mapped Drives

![Device Mapped Drives](/screenshots/mappeddrives.png?raw=true "Device Mapped Drives")

### Device Mapped Printers

![Device Mapped Printers](/screenshots/mappedprinters.png?raw=true "Device Mapped Printers")

### Device Installed Programs

![Device Installed Programs](/screenshots/installedprograms.png?raw=true "Device Installed Programs")

### Device Details

![Device Details](/screenshots/computerdetails.png?raw=true "Device Details")

### Asset Details

![Asset Details](/screenshots/assetdetails.png?raw=true "Asset Details")

### Reboot Device

![Reboot Device](/screenshots/rebootcomputer.png?raw=true "Reboot Device")

### Ping Device

![Ping Device](/screenshots/pingcomputer.png?raw=true "Ping Device")
