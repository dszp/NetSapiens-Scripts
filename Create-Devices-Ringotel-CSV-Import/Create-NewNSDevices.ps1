<#
.SYNOPSIS
    Creates new devices for NetSapiens subscribers based on provided extensions or a CSV file. 
    Optionally, create CSV import files designed for use in creating new users in the 
    Ringotel VoIP app service that connect to the PBX devices.

.DESCRIPTION
    This script creates new devices for NetSapiens subscribers in a specified domain. It can process
    extensions provided as an array or from a CSV file. The script checks for existing devices and
    creates new ones as needed, generating a CSV file for active and inactive extensions.

    REQUIREMENTS/DEPENDENCIES:
    You must do this first: Place the NetSapiensAPI module folder in the same directory as this script
    The module is located at: https://github.com/dszp/NetSapiensAPI (requires version 0.1.1 or later).

    1Password CLI is required to fetch the NetSapiens API credentials. The 1Password CLI is documented at:
    https://developer.1password.com/docs/cli/get-started/ and you must have a 1Password account and save 
    your API credentials in 1Password and access them via the correct vault paths. Alternately, 
    you can modify the script to insert your own credentials directly or using your own secrets management 
    system in place of pulling them from 1Password.

.NOTES
    Version:         0.0.3
    Last Updated:    2025-02-11
    Author:          David Szpunar
    
    Version History:
    0.0.1 - 2025-02-05
        * Initial release
        * Basic functionality for creating new devices for import into Ringotel from NetSapiens.
        * Support for CSV input and direct extension array input
        * Multiple switch options to control behavior, including creating import files, reporting only, etc.
        * Command-line only (no GUI). Requires NetSapiensAPI module to be in the same directory as this script.
        * NetSapiens API credentials are required; the script pulls them from 1Password using the 1Password 
          CLI by default.
    0.0.2 - 2025-02-06
        * Added requirement for NetSapiensAPI module to be version 0.1.1 or later to resolve filtering bug 
          that prevented proper filtering out of system extensions.
    0.0.3 - 2025-02-11
        * Updated inactive device export to include username and authname fields, which should prevent duplicate 
          entries in Ringotel import so duplicate extensions don't get created for inactive display/directory.
        * Added Test-SubscriberBlocklist function to define an additional list of subscribers to be skipped based 
          on their properties that make them unsuited to Ringotel activation. Moved check of ServiceCode to indicate 
          SYSTEM extensions to this check. Feel free to customize this list to your needs.

.PARAMETER DomainName
    The NetSapiens domain name to operate on.

.PARAMETER Extensions
    An array of extension numbers to process. Use this to provide a set of extensions to activate 
    on the command line rather than by providing a CSV file using -CsvSource. Overridden if a CSV 
    file is specified using -CsvSource.

.PARAMETER CsvSource
    Path to a CSV file containing extension numbers. Use this to provide a set of extensions to 
    activate from a CSV file rather than by providing an array of extensions using -Extensions. 
    If specified, this parameter takes precedence over -Extensions. The file must be in CSV format 
    and must have column names on the first line. Any column name that starts with "ext" will be 
    used to determine the extension numbers column; all other columns will be ignored. An error 
    will be raised if the file is not found, is not in CSV format, or does not have exactly one
    header that starts with "ext".

.PARAMETER Suffix
    Optional suffix to append to the extension when creating new devices. Default is 'r'.

.PARAMETER UseCallerIdName
    Switch to use the Caller ID Name instead of the subscriber's full name in the output CSV files.

.PARAMETER CreateBillable
    Switch to allow creation of new billable devices (device added to an extension that has no devices already).

.PARAMETER ReportOnly
    Switch to prevent creation of new devices and only report on existing devices. An alert will be output 
    where a device would have been crated, but extensions will be added to the Inactive list/export 
    rather than being created.

.PARAMETER CreateImportFiles
    Switch to create import files for Ringotel (active and inactive), based on the various OutputFilePath 
    parameters that are available. Files will only be created if there is one or more device to list.

.PARAMETER OutputFilePath
    Path for the output CSV file for newly created active devices to import into Ringotel. 
    Default is './ringotel_import.csv' in the current folder. SIP credentials are included 
    in this file, keep it safe!

.PARAMETER OutputFilePathAlreadyActive
    Path for the output CSV file for existing Ringotel-active devices to re-import into Ringotel or review. 
    Default is './ringotel_alreadyactive.csv' in the current folder. SIP credentials are included 
    in this file, keep it safe!

.PARAMETER OutputFilePathAlreadyInactive
    Path for the output CSV file for all non-system PBX domain extensions for review or for import into 
    Ringotel as Inactive users for directory purposes (keep mind importing over top of existing inactive 
    users will DUPLICATE users in Ringotel!). 
    Default is './ringotel_inactive_import.csv' in the current folder. SIP credentials are NOT in this file.

.EXAMPLE
    .\Create-NewNSDevices.ps1 -DomainName "example.0000.service" -Extensions 1001,1002,1003 -CreateBillable -CreateImportFiles

.EXAMPLE
    .\Create-NewNSDevices.ps1 -DomainName "example.0000.service" -CsvPath "extensions.csv" -Suffix "x" -UseCallerIdName -CreateImportFiles
#>

[CmdletBinding(DefaultParameterSetName = 'Array')]
param (
    [Parameter(Mandatory = $true)]
    [string]$DomainName,

    [Parameter(
        ParameterSetName = 'Array',
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [string[]]$Extensions,

    [Parameter(
        ParameterSetName = 'File',
        Position = 0,
        ValueFromPipelineByPropertyName = $true
    )]
    [ValidateScript({
            if (-not (Test-Path -Path $_ -PathType Leaf)) {
                throw "File not found: $_"
            }
            if (-not ($_ -match '\.csv$')) {
                throw "File must be a CSV file"
            }
            $true
        })]
    [string]$CsvSource,

    [Parameter(Mandatory = $false)]
    [string]$Suffix = "r",

    [Parameter(Mandatory = $false)]
    [switch]$UseCallerIdName,

    [Parameter(Mandatory = $false)]
    [switch]$CreateBillable,

    [Parameter(Mandatory = $false)]
    [switch]$CreateImportFiles,

    [Parameter(Mandatory = $false)]
    [switch]$ReportOnly,

    [Parameter(Mandatory = $false)]
    [string]$OutputFilePath = './ringotel_import.csv',

    [Parameter(Mandatory = $false)]
    [string]$OutputFilePathInactive = './ringotel_inactive_import.csv',

    [Parameter(Mandatory = $false)]
    [string]$OutputFilePathAlreadyActive = './ringotel_alreadyactive.csv'
)

# MUST DO THIS FIRST: Place the NetSapiensAPI module in the same directory as this script
# The module is located at: https://github.com/dszp/NetSapiensAPI
# Ensure version 0.1.1 or later is used to resolve filtering bug from module that prevented 
# proper filtering out of system extensions.

# Get the current script's directory and construct module path
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulePath = Join-Path $scriptDir "NetSapiensAPI"

# Verify module file exists
if (-not (Test-Path $modulePath)) {
    throw "Module manifest not found at: $modulePath"
}

# Import the module
Import-Module $modulePath -Force -MinimumVersion 0.1.1

# Example configuration - replace with your values
$NsConfig = @{
    BaseUrl      = "https://api.ucaasnetwork.com"
    ClientId     = $(op read op://Employee/netsapiens-script/clientid)
    ClientSecret = $(op read op://Employee/netsapiens-script/clientsecret)
    Domain       = $DomainName
}

# Create credential object
$securePassword = ConvertTo-SecureString $(op read op://Employee/netsapiens-script/credential) -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($(op read op://Employee/netsapiens-script/username), $securePassword)

# Initialize arrays for CSV output
$csvActive = New-Object System.Collections.Generic.List[object]
$csvInactive = New-Object System.Collections.Generic.List[object]
$csvAlreadyActive = New-Object System.Collections.Generic.List[object]

# Connect to NetSapiens
Write-Host "Connecting to NetSapiens..." -ForegroundColor Cyan
Connect-NSServer -BaseUrl $NsConfig.BaseUrl -ClientId $NsConfig.ClientId -ClientSecret $NsConfig.ClientSecret -Credential $credential

##### FUNCTIONS #####

<#
.SYNOPSIS
Finds extensions from either an array or a CSV file.

.DESCRIPTION
The Find-Extensions function processes input extensions either directly from an array or from a CSV file. It supports two parameter sets: 'Array' for direct input and 'File' for CSV file input.

.PARAMETER Extensions
An array of extension strings to process. This parameter is part of the 'Array' parameter set.

.PARAMETER CsvPath
The path to a CSV file containing extensions. The file must have a header that starts with 'ext'. This parameter is part of the 'File' parameter set.

.EXAMPLE
Find-Extensions -Extensions "101", "102", "103"
Processes the extensions 101, 102, and 103 directly.

.EXAMPLE
Find-Extensions -CsvPath "path/to/your/file.csv"
Reads extensions from the specified CSV file and processes them.

.EXAMPLE
"101", "102", "103" | Find-Extensions
Demonstrates how to use pipeline input to process extensions.

.NOTES
- The function validates the CSV file to ensure it exists and has a .csv extension.
- When using a CSV file, the function looks for a header starting with 'ext' to identify the column containing extensions.
- The function returns an array of extensions, which can be empty if no valid extensions are found.

.OUTPUTS
System.String[]
Returns an array of extension strings.
#>
function Find-Extensions {
    [CmdletBinding(DefaultParameterSetName = 'Array')]
    param (
        [Parameter(
            ParameterSetName = 'Array',
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string[]]$Extensions,

        [Parameter(
            ParameterSetName = 'File',
            Position = 0,
            ValueFromPipelineByPropertyName = $true
        )]
        [ValidateScript({
                if (-not (Test-Path -Path $_ -PathType Leaf)) {
                    throw "File not found: $_"
                }
                if (-not ($_ -imatch '\.csv$')) {
                    throw "File must be a CSV file"
                }
                $true
            })]
        [string]$CsvPath
    )

    begin {
        $resultArray = @()
    }

    process {
        if ($PSCmdlet.ParameterSetName -eq 'Array') {
            $resultArray += $Extensions
        }
        else {
            try {
                $csvContent = Import-Csv -Path $CsvPath
                
                # Find header that starts with 'ext' (case-insensitive)
                $extHeader = $csvContent[0].PSObject.Properties.Name | Where-Object { $_ -imatch '^ext' }
                
                if ($null -eq $extHeader) {
                    Write-Warning "No header starting with 'ext' found in CSV file"
                    return @()
                }
                elseif ($extHeader.Count -gt 1) {
                    Write-Warning "Multiple headers starting with 'ext' found in CSV file"
                    return @()
                }
                
                # Extract values from the matching column
                $resultArray = $csvContent | ForEach-Object { $_.$extHeader } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            }
            catch {
                Write-Error "Error processing CSV file: $_"
                return @()
            }
        }
    }

    end {
        return $resultArray
    }
}

<#
.SYNOPSIS
Checks if a subscriber should be blocked/skipped from activation based on predefined criteria.

.DESCRIPTION
The Test-SubscriberBlocklist function evaluates a subscriber's properties against a set of predefined rules to determine if the subscriber should be blocked/skipped from activation.

.PARAMETER Subscriber
An object representing the subscriber from the NetSapiensAPI module.

.EXAMPLE
$result = Test-SubscriberBlocklist -Subscriber $subscriberObject
if ($result.IsBlocked) {
    Write-Host "Subscriber is blocked: $($result.Reason)"
} else {
    Write-Host "Subscriber is not blocked"
}

.NOTES
The function checks various conditions including:
- Service Code is not blank (indicates SYSTEM extension type)
- Extensions starting with '9'
- First names starting with 'Paging' or 'Routing Group'
- Names containing 'On-Call', 'Voicemail', 'Shared', 'Ringer', 'Pager', or 'Ring Group'
- Empty first and last namespace
- Domains containing '0000.#####.service'
- Extensions not starting with a digit
- Email field is blank (no way to send activation email)

.OUTPUTS
PSCustomObject
Returns an object with two properties: IsBlocked (bool) and Reason (string).
#>

function Test-SubscriberBlocklist {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        $Subscriber
    )
    $Extension = $Subscriber.User

    # Helper function to return result with reason
    function New-BlockResult {
        param([bool]$Blocked, [string]$Reason = '')
        if ($Blocked -and $Reason) {
            Write-Verbose "Extension $Extension blocked: $Reason"
        }
        return [PSCustomObject]@{
            IsBlocked = $Blocked
            Reason    = $Reason
        }
    }

    # Check each condition and return appropriate result
    if (-not [string]::IsNullOrEmpty($Subscriber.ServiceCode)) {
        return New-BlockResult $true "Service Code indicates SYSTEM extension type '$($subscriber.ServiceCode)'"
    }
    if ($Extension -like '9*') {
        return New-BlockResult $true "Extension starts with '9'"
    }
    if ($Subscriber.FirstName -like '*Paging*' -or $Subscriber.LastName -like '*Paging*') {
        return New-BlockResult $true "Name contains 'Paging'"
    }
    if ($Subscriber.FirstName -like '*Routing Group*' -or $Subscriber.LastName -like '*Routing Group*') {
        return New-BlockResult $true "Name contains 'Routing Group'"
    }
    if ($Subscriber.FirstName -like '*On-Call*' -or $Subscriber.LastName -like '*On-Call*') {
        return New-BlockResult $true "Name contains 'On-Call'"
    }
    if ($Subscriber.FirstName -like '*Voicemail*' -or $Subscriber.LastName -like '*Voicemail*') {
        return New-BlockResult $true "Name contains 'Voicemail'"
    }
    if ($Subscriber.FirstName -like '*Shared*' -or $Subscriber.LastName -like '*Shared*') {
        return New-BlockResult $true "Name contains 'Shared'"
    }
    if ($Subscriber.FirstName -like '*Ringer*' -or $Subscriber.LastName -like '*Ringer*') {
        return New-BlockResult $true "Name contains 'Ringer'"
    }
    if ($Subscriber.FirstName -like '*Pager*' -or $Subscriber.LastName -like '*Pager*') {
        return New-BlockResult $true "Name contains 'Pager'"
    }
    if ($Subscriber.FirstName -like '*Ring Group*' -or $Subscriber.LastName -like '*Ring Group*') {
        return New-BlockResult $true "Name contains 'Ring Group'"
    }
    if ($Subscriber.FirstName -like '*Fax*' -or $Subscriber.LastName -like '*Fax*') {
        return New-BlockResult $true "Name contains 'Fax'"
    }
    if ([string]::IsNullOrEmpty($Subscriber.Email)) {
        return New-BlockResult $true "Email field is empty"
    }
    if ($Subscriber.Domain -like '*0000.#####.service*') {
        return New-BlockResult $true "Domain contains '0000.#####.service'"
    }
    if (-not $Extension[0] -match '\d') {
        return New-BlockResult $true "Extension doesn't start with a digit"
    }
    if ([string]::IsNullOrEmpty($Subscriber.FirstName) -and [string]::IsNullOrEmpty($Subscriber.LastName)) {
        return New-BlockResult $true "Both first and last names are empty"
    }

    # If no conditions matched, return not blocked
    return New-BlockResult $false
}

##### BEGIN SCRIPT #####

if ($CsvSource) {
    Write-Host "Using CSV file: $CsvSource to gather extensions to activate."
    $Extensions = Find-Extensions -CsvPath $CsvSource
}
Write-Host "Extensions to activate: $Extensions" -ForegroundColor Green

$Subscribers = Get-NSSubscriber -Domain $DomainName

# Get subscriber information
Write-Host "`nProcessing subscribers..."
foreach ($Subscriber in $Subscribers) {
    $Extension = $Subscriber.User
    Write-Host "`nGetting subscriber information for extension $Extension..."

    $blockResult = Test-SubscriberBlocklist -Subscriber $Subscriber -Verbose
    if ($blockResult.IsBlocked) {
        # Write-Host "Subscriber $Extension is blocked because: $($blockResult.Reason)" -ForegroundColor Yellow
        continue
    }

    # Check to see if the extension has any devices and warn or error if it doesn't (and is thus non-billable)
    $extensionCount = Get-NSDeviceCount -Domain $DomainName -User $Subscriber.User
    if ($extensionCount -eq 0) {
        if (!$CreateBillable) {
            Write-Host "No devices found for extension $Extension. Use -CreateBillable switch to allow creation of new device on unbillable extension." -ForegroundColor Red
            $ringotel_data = [PSCustomObject]@{
                extension = $Subscriber.User
                name      = $subscriber.FullName
                email     = $subscriber.Email
                username  = "$Extension$Suffix"
                authname  = "$Extension$Suffix"
                password  = ""
            }
            $csvInactive.Add($ringotel_data)
            continue
        }
        else {
            if ($Extension -in $Extensions -and !$ReportOnly) {
                Write-Host "No devices found for extension $Extension. Creating new device on unbillable extension." -ForegroundColor Green
            }
            elseif ($Extension -in $Extensions -and $ReportOnly) {
                Write-Host "No devices found for extension $Extension. Would have created new device on unbillable extension if not in ReportOnly mode." -ForegroundColor Red
            }
            else {
                Write-Host "No devices found for extension $Extension and none requested. Adding to Inactive list." -ForegroundColor Green
            }
        }
    }

    $newDevice = "sip:$Extension$Suffix@$DomainName"
    $newUser = $Extension
    Write-Host "NewDevice: $newDevice"
    $deviceExists = Get-NSDeviceCount -Domain $DomainName -AOR $newDevice
    if ($deviceExists -gt 0) {
        # $existingDevice = Get-NSDevice -Domain $DomainName -User $newDevice
        Write-Host "Device already exists. Won't create, but will retrieve existing device and record in the Already Active list." -ForegroundColor Yellow
        $new_device = New-NSDevice -Domain $DomainName -Device $newDevice -User $newUser -Mac $newMac -Model $newModel
        $ringotel_data = [PSCustomObject]@{
            extension = $Extension
            name      = $subscriber.FullName
            email     = $subscriber.Email
            username  = $newDevice.Split(':')[1].Split('@')[0]
            authname  = $newDevice.Split(':')[1].Split('@')[0]
            password  = $new_device.AuthenticationKey
        }
        $csvAlreadyActive.Add($ringotel_data)
    }
    else {
        if ($Extension -in $Extensions -and !$ReportOnly) {
            # Only activate extensions that are in the Extensions array
            Write-Host "Device does not yet exist for requested extension $Extension. Creating new device and adding to the Active list." -ForegroundColor Green
            $new_device = New-NSDevice -Domain $DomainName -Device $newDevice -User $newUser -Mac $newMac -Model $newModel

            if ($new_device) {
                $ringotel_data = [PSCustomObject]@{
                    extension = $Extension
                    name      = $subscriber.FullName
                    email     = $subscriber.Email
                    username  = "$Extension$Suffix"
                    authname  = "$Extension$Suffix"
                    password  = $new_device.AuthenticationKey
                }
                $csvActive.Add($ringotel_data)
            }
        }
        else {
            if ($Extension -in $Extensions -and $ReportOnly) {
                Write-Host "In ReportOnly mode or WOULD have created new device for extension $Extension. Listing as Inactive instead." -ForegroundColor Red
            }
            $ringotel_data = [PSCustomObject]@{
                extension = $Extension
                name      = $subscriber.FullName
                email     = $subscriber.Email
                username  = "$Extension$Suffix"
                authname  = "$Extension$Suffix"
                password  = ""
            }
            $csvInactive.Add($ringotel_data)
        }
    }
}

# Output results summary
Write-Host "`nProcessing complete!" -ForegroundColor Green
Write-Host "Active extensions (new): $($csvActive.Count)" -ForegroundColor Green
Write-Host "Already active extensions: $($csvAlreadyActive.Count)" -ForegroundColor Cyan
Write-Host "Inactive extensions: $($csvInactive.Count)" -ForegroundColor Yellow

if ($CreateImportFiles) {
    if ($csvActive.Count -gt 0) {
        Write-Host "Outputting Import Data to $OutputFilePath" -ForegroundColor Green
        $csvActive | 
        Sort-Object -Property extension | 
        Select-Object extension, name, email, username, authname, password | 
        Export-Csv -Path $OutputFilePath -NoTypeInformation
    }
    if ($csvInactive.Count -gt 0) {
        Write-Host "Outputting Import Data to $OutputFilePathInactive" -ForegroundColor Yellow
        $csvInactive | 
        Sort-Object -Property extension | 
        Select-Object extension, name, email, username, authname, password | 
        Export-Csv -Path $OutputFilePathInactive -NoTypeInformation
    }
    if ($csvAlreadyActive.Count -gt 0) {
        Write-Host "Outputting Import Data to $OutputFilePathAlreadyActive" -ForegroundColor Cyan
        $csvAlreadyActive | 
        Sort-Object -Property extension | 
        Select-Object extension, name, email, username, authname, password | 
        Export-Csv -Path $OutputFilePathAlreadyActive -NoTypeInformation
    }
}
