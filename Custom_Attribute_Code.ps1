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

param (
    [Parameter(Mandatory=$true)]
    [string]$vCenterServer,
    
    [switch]$Report
)

# Function to write log messages and print them to the console
function Write-Log {
    param (
        [string]$message,
        [string]$logFile
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $message"
    Add-Content -Path $logFile -Value $logEntry
    Write-Output $logEntry
}

# Function to select a CSV file using an Open File Dialog
function Select-CSVFile {
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $openFileDialog.Title = "Select a CSV file"
    $openFileDialog.ShowDialog() | Out-Null
    return $openFileDialog.FileName
}

# Check PowerShell version
function Check-PowerShellVersion {
    $requiredVersion = [Version]"7.4.3"
    if ($PSVersionTable.PSVersion -lt $requiredVersion) {
        throw "PowerShell version 7.4.3 or later is required. Installed version is $($PSVersionTable.PSVersion)."
    }
}

# Check .NET Framework version
function Check-DotNetVersion {
    $registryKey = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
    $releaseKey = Get-ItemProperty -Path $registryKey -Name Release -ErrorAction SilentlyContinue
    if ($null -eq $releaseKey) {
        throw ".NET Framework 4.8 or later is required. .NET Framework is not installed."
    }
    
    $release = $releaseKey.Release
    if ($release -lt 528040) { # .NET Framework 4.8 or later
        throw ".NET Framework 4.8 or later is required. Installed version release key is $release."
    }
}

# Check PowerCLI module version
function Check-PowerCLIVersion {
    $requiredVersion = [Version]"13.2.1.22851661" # Adjust this to the required version
    $installedModule = Get-Module -ListAvailable -Name VMware.PowerCLI
    if ($null -eq $installedModule) {
        throw "VMware.PowerCLI module is not installed."
    }

    if ($installedModule.Version -lt $requiredVersion) {
        throw "PowerCLI version $requiredVersion or later is required. Installed version is $($installedModule.Version)."
    }
}

# Check all requirements
Check-PowerShellVersion
Check-DotNetVersion
Check-PowerCLIVersion

# Select the CSV file
$csvFilePath = Select-CSVFile
if (-not (Test-Path $csvFilePath)) {
    throw "No CSV file selected or file does not exist."
}

# Get the folder of the selected CSV file
$csvFolderPath = Split-Path -Path $csvFilePath -Parent

# Define the paths for the log file and the credential file in the same folder
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path -Path $csvFolderPath -ChildPath "log_file_$timestamp.log"
$credentialFile = Join-Path -Path $csvFolderPath -ChildPath "credential_file.xml"

# Start logging
Write-Log "Script started." $logFile

try {
    # Disable CEIP participation and ignore certificate warnings for PowerCLI
    Write-Log "Setting PowerCLI configuration..." $logFile
    Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -Confirm:$false -Verbose
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -Verbose
    Write-Log "PowerCLI configuration set." $logFile

    # Import the credentials from the encrypted file
    Write-Log "Importing credentials..." $logFile
    $credential = Import-Clixml -Path $credentialFile
    Write-Log "Credentials imported." $logFile

    # Connect to vCenter Server
    Write-Log "Connecting to vCenter Server..." $logFile
    $connection = Connect-VIServer -Server $vCenterServer -Credential $credential -Verbose -ErrorAction Stop
    Write-Log "Connected to vCenter Server." $logFile

} catch {
    Write-Log "Error in initial setup: $_" $logFile
    Write-Error "Error in initial setup: $_"
    throw
}

try {
    # Import the CSV file
    Write-Log "Importing CSV file..." $logFile
    $csv = Import-Csv -Path $csvFilePath
    Write-Log "CSV file imported from path: $csvFilePath" $logFile
} catch {
    Write-Log "Error importing CSV file: $_" $logFile
    Write-Error "Error importing CSV file: $_"
    throw
}

# Extract custom attribute names from the CSV header
$customAttributes = $csv | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name | Where-Object { $_ -ne 'Server Name' }

# Get all existing custom attributes
$existingCustomAttributes = Get-CustomAttribute -ErrorAction SilentlyContinue
$existingCustomAttributeNames = $existingCustomAttributes | Select-Object -ExpandProperty Name

# Report all existing custom attributes
Write-Log "Existing custom attributes:" $logFile
foreach ($attribute in $existingCustomAttributeNames) {
    Write-Log $attribute $logFile
}

foreach ($attribute in $customAttributes) {
    try {
        # Check if the custom attribute already exists
        $existingAttribute = Get-CustomAttribute -Name $attribute -ErrorAction SilentlyContinue
        if ($Report) {
            if ($existingAttribute) {
                Write-Log "Report: Custom attribute '$attribute' already exists." $logFile
            } else {
                Write-Log "Report: Custom attribute '$attribute' does not exist and would be created." $logFile
            }
        } else {
            if (-not $existingAttribute) {
                New-CustomAttribute -Name $attribute -TargetType VirtualMachine -Confirm:$false -Verbose
                Write-Log "Created custom attribute '$attribute'." $logFile
            } else {
                Write-Log "Custom attribute '$attribute' already exists." $logFile
            }
        }
    } catch {
        Write-Log "Error creating custom attribute '$attribute': $_" $logFile
        Write-Error "Error creating custom attribute '$attribute': $_"
    }
}

# Generate report for each VM and their custom attributes
foreach ($row in $csv) {
    # Validate Server Name
    if ([string]::IsNullOrEmpty($row.'Server Name')) {
        Write-Log "Error: Server Name is null or empty for a row. Skipping this row." $logFile
        Write-Error "Error: Server Name is null or empty for a row. Skipping this row."
        continue
    }

    try {
        $vm = Get-VM -Name $row.'Server Name' -ErrorAction SilentlyContinue
        
        if ($null -ne $vm) {
            foreach ($attribute in $customAttributes) {
                $value = $row.$attribute
                $existingValue = Get-Annotation -Entity $vm -CustomAttribute $attribute -ErrorAction SilentlyContinue
                
                if ($Report) {
                    if ($existingValue -ne $value) {
                        Write-Log "Report: Custom attribute '$attribute' would be set to '$value' for VM '$($vm.Name)'." $logFile
                    } else {
                        Write-Log "Report: Custom attribute '$attribute' for VM '$($vm.Name)' already has the value '$value'." $logFile
                    }
                } else {
                    if ($existingValue -ne $value) {
                        Set-Annotation -Entity $vm -CustomAttribute $attribute -Value $value -ErrorAction Stop
                        Write-Log "Set custom attribute '$attribute' to '$value' for VM '$($vm.Name)'." $logFile
                    } else {
                        Write-Log "Custom attribute '$attribute' for VM '$($vm.Name)' already has the value '$value'." $logFile
                    }
                }
            }
        } else {
            Write-Log "Report: VM '$($row.'Server Name')' not found." $logFile
        }
    } catch {
        Write-Log "Error processing VM '$($row.'Server Name')': $_" $logFile
        Write-Error "Error processing VM '$($row.'Server Name')': $_"
    }
}

# Finish logging
Write-Log "Script finished." $logFile

# Disconnect from vCenter Server if not in report mode
if (-not $Report) {
    Disconnect-VIServer -Server $connection -Confirm:$false -Verbose
    Write-Log "Disconnected from vCenter Server." $logFile
} else {
    Write-Log "Report: Skipping disconnection from vCenter Server." $logFile
}
