<#
.SYNOPSIS
This script connects to a vCenter Server, manages custom attributes on virtual machines based on a CSV file, and supports a report mode.

.DESCRIPTION
The script performs the following actions:
1. Sets PowerCLI configuration.
2. Imports credentials from an encrypted XML file.
3. Connects to a vCenter Server.
4. Creates custom attributes if they don't already exist.
5. Imports a CSV file and applies custom attributes to VMs as specified in the CSV.
6. Logs actions and errors to a timestamped log file.
7. Supports a report mode where no actual changes are made but a detailed report is generated.

.REQUIREMENTS
- PowerShell version 7.4.3 or later.
- .NET Framework version 4.8 or later.
- VMware PowerCLI module installed (version 13.2.1 or later).
- Necessary permissions to manage custom attributes and virtual machines in vCenter Server.
- Encrypted credentials XML file for authentication.
- CSV file prepared in advance with the following columns:
  - Server Name - Mandatory --> Virtual Machine Name
  - Rest Of Custom Attributes will be identified by the given csv file

.AUTHOR
Mehmet Hakan Can

.PARAMETER vCenterServer
The name or address of the vCenter Server to connect to.

.PARAMETER Report
If specified, the script will generate a report without making any actual changes.

.EXAMPLE
.\Custom_Attribute_Code.ps1 -vCenterServer "vcenter.example.com" -Report
Generates a report of the script without making any changes. Logs actions and outputs them to the console for review.

.EXAMPLE
.\Custom_Attribute_Code.ps1 -vCenterServer "vcenter.example.com"
Executes the script and applies custom attributes to VMs as specified in the CSV file. Logs actions and errors to a timestamped log file.

.NOTES
Running this script is at the user's own risk. The author cannot be held responsible for any problems caused by the script. 
Please ensure you have a valid backup and snapshot of the vCenter Server before running the script.
#>
