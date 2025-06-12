#Requires -RunAsAdministrator
<#
.SYNOPSIS
    This script syncs helpdesk information with Entra groups, Active Directory, and Intune.
.DESCRIPTION
    This script performs the following operations:
        - Adds or removes computers to Entra groups for software distribution based on helpdesk data.
        - Removes the primary user from all Intune devices, converting them to shared devices.
        - Removes invalid computer objects from Active Directory.
        - Attempts to rename any misnamed computers in Active Directory to match helpdesk records.
        - Cleans up any duplicate device entries found in Entra ID and Intune.
        - Removes invalid serial numbers from Autopilot deployment profiles.
        - Generates a list of computers that cannot be found in Entra ID during the helpdesk group sync,
          for further evaluation by the helpdesk.
.NOTES
    This script creates a custom log in Event Viewer under "Application and Service Logs" -> "Intunesync".
    Transcription log location: %programdata%\IntuneSync\Intunesync.log
    List of invalid/missing computers location: %programdata%\IntuneSync\MissingDevices.txt
    Ensure all global variables are correctly configured and the syncing server has the required certificate installed.

    EventLogIDs (for structured logging in Event Viewer):
        000 - Script related events
            00 - Script Executed Successfully
            01 - Script Started
        100 - User related events (not extensively used in this script)
        200 - Device related events
            201 - Successfully Deleted Primary User From Device
            202 - Failed to delete primary user from device
            203 - Failed to Find device to delete primary user
            204 - Successfully Removed Device from AutoPilot
            205 - Failed to Removed Device from AutoPilot
            206 - Number of Duplicate Devices found in Intune
            207 - Duplicate Device Successfully Removed from Intune
            208 - Failed to Remove Duplicate Device from Intune
            209 - Couldn't Find Device in Entra
        300 - Group related events
            301 - Successfully Added Computer to group
            302 - Successfully Deleted Computer from group
            303 - Error Adding Device to Group
            304 - Error Deleting Device From Group
            305 - Error Couldn't find Building ID
        400 - Connection related events
            401 - Successfully Connected to Microsoft Graph API
            402 - Failed to Connect to Microsoft Graph API
        500 - Module/Log related events
            502 - Failed to Install Module
            503 - Removed Log File exceeded 50MB
        600 - Active Directory related events
            601 - Successfully Deleted computer from AD
            602 - Failed to Delete computer from AD
            603 - Successfully Renamed Computer
            604 - Failed to Renamed Computer
            605 - Couldn't connect to Remote Computer
            606 - Computer Object too new, trying next sync
            607 - No misnamed computer found in AD
#>

# Path to the credential XML file for Secret Management.
$Credxmlpath = "<path to credential XML file>"

#region Credential and Secret Management
try {
    # Import necessary modules for Secret Management.
    Import-Module Microsoft.PowerShell.SecretManagement, Microsoft.PowerShell.SecretStore -ErrorAction Stop
    # Import encrypted credentials from the XML file.
    $Credential = Import-Clixml $Credxmlpath -ErrorAction Stop
    # Unlock the Secret Store using the imported password.
    Unlock-SecretStore -Password $Credential.password -ErrorAction Stop
}
catch {
    # Output any errors encountered during credential import or Secret Store unlock.
    $_
    # Exit the script with an error code if an issue occurs.
    Exit 1
}
#endregion

#region Global Variables
####Global Variables#####
# Retrieve Schoolza API key from Secret Store.
$SchoolzaAPIkey = Get-secret -Name "SchoolzaAPIkey" -AsPlainText
# Retrieve Azure Tenant ID from Secret Store.
$azureTenantId = Get-secret -Name "azureTenantId" -AsPlainText
# Retrieve Azure Application ID from Secret Store.
$azureAplicationId = Get-secret -Name "azureAplicationId" -AsPlainText
# Read the Azure certificate thumbprint from a text file located in the script's root directory.
$azureCertThumbprint = Get-Content -Path "$PSScriptRoot/Thumbprint.txt"
# Define the exact names of the Entra (Azure AD) lab groups to be synced.
$LabGroupNames = @(
    'Intune-Lab-NHS-RM402'
    'Intune-Lab-NHS-RM406'
    'Intune-Lab-NHS-RM300'
    'Intune-Lab-NHS-RM301'
    'Intune-Lab-EHS-RMD140' 
    'Intune-Lab-EHS-RMD150'
    'Intune-Lab-EHS-RME160'
    'Intune-Lab-HC-RM145'
    'Intune-Lab-EHS-RMB200'
    'Intune-Lab-TMS-RM250'
    'Intune-Lab-TMS-RM249'
    'Intune-Lab-WMS-RM402'
    )
# Do not modify past here. This comment indicates a demarcation for user-configurable variables.
#endregion

#region Event Logging Functions
function Get-IntuneEventLog {
    <#
    .SYNOPSIS
        Ensures the custom 'IntuneSync' event log exists.
    .DESCRIPTION
        Checks if an event log named 'IntuneSync' exists. If not, it creates a new custom event log
        under "Application and Service Logs" with "IntuneSync" as the source.
    #>
    if([System.Diagnostics.EventLog]::Exists('IntuneSync')){}
    else {
        New-EventLog -source "IntuneSync" -LogName "IntuneSync"
    }
}

function Write-IntuneEventLog {
    <#
    .SYNOPSIS
        Writes a custom entry to the 'IntuneSync' event log.
    .DESCRIPTION
        This function takes an Event ID, Event Type (e.g., "Information", "Error", "Warning"),
        and a message string, and writes them to the custom 'IntuneSync' event log.
    .PARAMETER EventID
        The numerical Event ID for the log entry, as defined in the script's NOTES.
    .PARAMETER EventType
        The type of event to log (e.g., "Information", "Error", "Warning").
    .PARAMETER Message
        The detailed message to be written to the event log.
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $EventID,
        [Parameter(Position=1,mandatory=$true)]
        [string] $EventType,
        [Parameter(Position=2,mandatory=$true)]
        [String] $Message
    )
    
    Write-EventLog -LogName "IntuneSync" -Source "IntuneSync" -EventID $EventID -EntryType $EventType -Message $Message -Category 1 -RawData 10,20
}
#endregion

#region Helpdesk and Building Information Functions
function Get-D303LabMatch {
    <#
    .SYNOPSIS
        Matches a given building and room number to helpdesk room data.
    .DESCRIPTION
        Translates a short building code (e.g., "NHS") to its full name and then searches through
        provided helpdesk room data to find a matching active room. It handles cases with multiple matches
        by prioritizing exact room name matches.
    .PARAMETER Building
        The short building code (e.g., "NHS", "EHS").
    .PARAMETER RoomNumber
        The room number within the building.
    .PARAMETER RoomData
        An array of objects containing room information from the helpdesk system (e.g., building_name, room_name, room_status).
    .RETURNS
        An object representing the matched room data from the helpdesk.
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $Building,
        [Parameter(Position=1,mandatory=$true)]
        [string] $RoomNumber,
        [Parameter(Position=2,mandatory=$true)]
        [array] $RoomData
    )
    
    # Initialize the full building name variable.
    $SwBldName = ""
    # Use a switch statement to translate short building codes to full names.
    switch($Building) {  
        "MVC" {$SwBldName = "Mades Johnstone"}
        "ADM" {$SwBldName = "Administration Center"}
        "FOX" {$SwBldName = "Fox Ridge School"}
        "AND" {$SwBldName = "Anderson Elementary School"}
        "DAV" {$SwBldName = "Davis Elementary School"}
        "LIN" {$SwBldName = "Lincoln Elementary School"}
        "COR" {$SwBldName = "Corron Elementary School"}
        "MUN" {$SwBldName = "Munhall Elementary School"}
        "RIC" {$SwBldName = "Richmond Intermediate School"}
        "WAS" {$SwBldName = "Wasco Elementary School"}
        "WIL" {$SwBldName = "Wild Rose Elementary School"}
        "FER" {$SwBldName = "Ferson Creek Elementary School"}
        "BEL" {$SwBldName = "Bell Graham Elementary School"}
        "NOR" {$SwBldName = "Norton Creek Elementary School"}
        "HC" {$SwBldName = "Haines Center"}
        "EHS" {$SwBldName = "St. Charles East High School"}
        "NHS" {$SwBldName = "St. Charles North High School"}
        "TMS" {$SwBldName = "Thompson Campus"}
        "WMS" {$SwBldName = "Wredling Middle School"}
        "PECK" {$SwBldName = "Peck Road"}
    }

    # Filter room data for active rooms matching the building name and room number (partial or exact).
    $OutputRoomData = ""
    $OutputRoomData = $RoomData | Where-Object {($_.building_name -eq $SwBldName ) -and (($_.room_name -Like "*$RoomNumber *") -xor ($_.room_name -Like "*$RoomNumber" )) -and ($_.room_status -eq "active")}
    
    # If multiple rooms match, try to find an exact room number match to refine the result.
    if ($OutputRoomData.count -gt 1) {
        $OutputRoomData = $OutputRoomData | Where-Object {($_.building_name -eq $SwBldName ) -and ($_.room_name -eq "$RoomNumber") -and ($_.room_status -eq "active")}
    }
    # Return the matched room data.
    Return $OutputRoomData
}

function Get-HelpdeskDevicesByRoom {
    <#
    .SYNOPSIS
        Retrieves asset tags for devices deployed in a specific room from the helpdesk.
    .DESCRIPTION
        Queries the helpdesk API for devices (asset types 2 and 11) that are currently
        'deployed' within a given room ID.
    .PARAMETER RoomID
        The unique identifier of the room in the helpdesk system.
    .PARAMETER HelpdeskAPIkey
        The API key for authenticating with the helpdesk system.
    .RETURNS
        An array of asset tags (computer names) of devices found in the specified room.
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $RoomID,
        [Parameter(Position=1,mandatory=$true)]
        [string] $HelpdeskAPIkey
    )
    # Define headers for the API request, including the API key.
    $headers = @{
        'schoolza-api-key' = $HelpdeskAPIkey
    }
    Write-Verbose $RoomID
    # Invoke the REST method to get devices by room ID and filter for deployed status.
    $RoomOutput = Invoke-RestMethod -Uri "https://d303.k12it.app/api/devices?room_id=$($RoomID)&asset_types=[2,11]" -Method Get -Headers $headers 
    $RoomOutput = $RoomOutput | Where-Object {$_.asset_status -eq "deployed"}
    # Return only the asset tags (computer names).
    Return $RoomOutput.asset_tag
}

function Get-HelpdeskDevice {
    <#
    .SYNOPSIS
        Retrieves detailed information for a single device from the helpdesk by serial number.
    .DESCRIPTION
        Queries the helpdesk API for a specific device using its serial number.
    .PARAMETER serialnumber
        The serial number of the device to query.
    .PARAMETER HelpdeskAPIkey
        The API key for authenticating with the helpdesk system.
    .RETURNS
        An object containing detailed information about the device from the helpdesk.
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $serialnumber,
        [Parameter(Position=1,mandatory=$true)]
        [string] $HelpdeskAPIkey
    )
    # Define headers for the API request, including the API key.
    $headers = @{
        'schoolza-api-key' = $HelpdeskAPIkey
    }
    # Invoke the REST method to get device information by serial number.
    $HelpdeskAssetInfo = Invoke-RestMethod -Uri "https://d303.k12it.app/api/device?serial_number=$($serialnumber)" -Method Get -Headers $headers 
    # Return the device information.
    Return $HelpdeskAssetInfo
}

function Get-HelpdeskRoomIDs {
    <#
    .SYNOPSIS
        Retrieves a list of all rooms from the helpdesk system.
    .DESCRIPTION
        Queries the helpdesk API to get a comprehensive list of all defined rooms.
    .PARAMETER HelpdeskAPIkey
        The API key for authenticating with the helpdesk system.
    .RETURNS
        An array of objects, each representing a room with its details from the helpdesk.
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $HelpdeskAPIkey
    )
    # Define headers for the API request, including the API key.
    $headers = @{
        'schoolza-api-key' = $HelpdeskAPIkey
    }
    # Invoke the REST method to get all room information.
    $HelpdeskRoomsOutput = Invoke-RestMethod -Uri "https://d303.k12it.app/api/rooms" -Method Get -Headers $headers 
    # Return the room information.
    Return $HelpdeskRoomsOutput
}

function Get-HelpdeskRetiredDevice {
    <#
    .SYNOPSIS
        Retrieves a list of retired/disposed devices from the helpdesk system.
    .DESCRIPTION
        Queries the helpdesk API for devices that have a status indicating they are retired,
        deleted, scrapped, or lost. It specifically targets asset types 2 and 11.
    .PARAMETER HelpdeskAPIkey
        The API key for authenticating with the helpdesk system.
    .PARAMETER AssetStatus
        (Optional) A JSON array string specifying the asset statuses to filter by.
        Defaults to '["recycled","deleted","scrap","lost"]'.
    .PARAMETER Assettypes
        (Optional) A JSON array string specifying the asset types to filter by.
        Defaults to '[2,11]'.
    .RETURNS
        An array of objects, each representing a retired device from the helpdesk.
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $HelpdeskAPIkey,
        [Parameter(Position=1)]
        [string] $AssetStatus = '["recycled","deleted","scrap","lost"]',
        [Parameter(Position=2)]
        [string] $Assettypes = '[2,11]'

    )
    # Define headers for the API request, including the API key.
    $headers = @{
        'schoolza-api-key' = $HelpdeskAPIkey
    }
    # Invoke the REST method to get devices based on specified status and type.
    Invoke-RestMethod -Uri "https://d303.k12it.app/api/devices?asset_statuses=$($AssetStatus)&asset_types=$($Assettypes)" -Method Get -Headers $headers 
}

function Get-HelpdeskUnassignedDevices {
    <#
    .SYNOPSIS
        Retrieves a list of unassigned devices from the helpdesk system.
    .DESCRIPTION
        Queries the helpdesk API for devices with an "available" status and specified asset types,
        excluding devices whose asset tags start with "LM".
    .PARAMETER HelpdeskAPIkey
        The API key for authenticating with the helpdesk system.
    .PARAMETER Assettypes
        (Optional) A JSON array string specifying the asset types to filter by.
        Defaults to '[2,11]'.
    .RETURNS
        An array of asset tags (computer names) for unassigned devices.
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $HelpdeskAPIkey,
        [Parameter(Position=1)]
        [string] $Assettypes = '[2,11]'

    )
    # Define headers for the API request, including the API key.
    $headers = @{
        'schoolza-api-key' = $HelpdeskAPIkey
    }
    # Set the asset status to "available".
    $AssetStatus = '["available"]'
    # Invoke the REST method to get unassigned devices.
    $UnassignedDevices = Invoke-RestMethod -Uri "https://d303.k12it.app/api/devices?asset_statuses=$($AssetStatus)&asset_types=$($Assettypes)" -Method Get -Headers $headers 
    # Filter out devices whose asset tags start with "LM".
    $UnassignedDevices = $UnassignedDevices | Where-Object {$_.asset_tag -notlike "LM*"}
    # Return only the asset tags.
    Return $UnassignedDevices.asset_tag
}

function Get-HelpdeskBuildings {
    <#
    .SYNOPSIS
        Retrieves a list of all buildings from the helpdesk system.
    .DESCRIPTION
        Queries the helpdesk API to get a comprehensive list of all defined buildings.
    .PARAMETER HelpdeskAPIkey
        The API key for authenticating with the helpdesk system.
    .RETURNS
        An array of objects, each representing a building with its details from the helpdesk.
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $HelpdeskAPIkey
    )
    # Define headers for the API request, including the API key.
    $headers = @{
        'schoolza-api-key' = $HelpdeskAPIkey
    }
    # Invoke the REST method to get all building information.
    $HelpdeskBuildings = Invoke-RestMethod -Uri "https://d303.k12it.app/api/buildings" -Method Get -Headers $headers 
    # Return the building information.
    Return $HelpdeskBuildings
}

function Get-HelpdeskDevicesByBuilding {
    <#
    .SYNOPSIS
        Retrieves asset tags for devices within a specific building from the helpdesk.
    .DESCRIPTION
        First, it resolves the building ID from the helpdesk system based on the provided building name.
        Then, it queries the helpdesk API for devices (asset types 2 and 11) that are in the specified building
        and match the given status. It excludes devices whose asset tags start with "LM".
    .PARAMETER HelpdeskAPIkey
        The API key for authenticating with the helpdesk system.
    .PARAMETER Building
        The name or partial name of the building to search for.
    .PARAMETER Status
        (Optional) A JSON array string specifying the asset statuses to filter by.
        Defaults to "['available','deployed','Loaner','out for repair']".
    .PARAMETER AssetType
        (Optional) A JSON array string specifying the asset types to filter by.
        Defaults to "[2,11]".
    .RETURNS
        An array of asset tags (computer names) of devices found in the specified building,
        or "NotFound" if the building ID cannot be resolved.
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $HelpdeskAPIkey,
        [Parameter(Position=1,mandatory=$true)]
        [string] $Building,
        [Parameter(mandatory=$false)]
        [string] $Status = "['available','deployed','Loaner','out for repair']",
        [Parameter(mandatory=$false)]
        [string] $AssetType = "[2,11]"
    )
    # Get all buildings from the helpdesk.
    $Buildings = Get-HelpdeskBuildings -HelpdeskAPIkey $HelpdeskAPIkey
    
    # Resolve the building ID based on the provided building name.
    $BuildingID = ($Buildings | Where-Object {$_.building_name -like "*$Building*"}).building_id
    
    # If the building ID is not found, return "NotFound".
    If(!($BuildingID)){
        Return "NotFound"
    }
    # Define headers for the API request, including the API key.
    $headers = @{
        'schoolza-api-key' = $HelpdeskAPIkey
    }
    # Invoke the REST method to get devices by status and type.
    $HelpdeskAssetInfo = Invoke-RestMethod -Uri "https://d303.k12it.app/api/devices?asset_statuses=$Status&asset_types=$AssetType" -Method Get -Headers $headers 
    # Filter devices by the resolved building ID.
    $HelpdeskAssetInfo = $HelpdeskAssetInfo | Where-Object {$_.building_id -eq "$BuildingID"}
    # Exclude devices whose asset tags start with "LM".
    $HelpdeskAssetInfo = $HelpdeskAssetInfo | Where-Object {$_.asset_tag -notlike "LM*"}
    # Return only the asset tags.
    Return $HelpdeskAssetInfo.asset_tag
}
#endregion

#region Microsoft Graph API Functions
function Get-D303MsGraphGroupIDbyName {
    <#
    .SYNOPSIS
        Retrieves the Group ID (Object ID) of an Entra (Azure AD) group by its display name.
    .DESCRIPTION
        Queries Microsoft Graph API to find a group based on its display name and returns its unique ID.
    .PARAMETER GroupName
        The display name of the Entra group.
    .RETURNS
        The unique ID (object ID) of the matched Entra group.
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $GroupName
    )
    try {
        $GroupIDOutput = ""
        # Get all Entra groups and filter by display name.
        $GroupIDOutput = Get-MgGroup -all | Where-Object { $_.DisplayName -eq $GroupName }
    }
    catch {
        # Output error if the group cannot be found.
        Write-Host "Problem Finding Group"
        $_
    }
    # Return the ID of the found group.
    Return $GroupIDOutput.Id
}

function Get-D303MsGraphGroupMembers {
    <#
    .SYNOPSIS
        Retrieves the display names of members within a specified Entra (Azure AD) group.
    .DESCRIPTION
        Queries Microsoft Graph API to get all members of a given group ID and extracts their display names.
    .PARAMETER GroupID
        The unique ID (object ID) of the Entra group.
    .RETURNS
        An array of display names of the group members.
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $GroupID
    )
    try {
        $GroupMSGraphOutput = @()
        # Get all members of the specified group and expand their additional properties to get display name.
        $GroupMSGraphOutput = Get-MgGroupMember -GroupId $GroupID -all | Select-Object * -ExpandProperty additionalProperties | Select-Object {$_.AdditionalProperties["displayName"]}
    }
    catch {
        # Output error if group members cannot be retrieved.
        Write-Host "Problem Finding Group"
        $_
    }
    # Return the extracted display names.
    Return $GroupMSGraphOutput.'$_.AdditionalProperties["displayName"]'
}

function Get-AllD303MsGraphDevices {
    <#
    .SYNOPSIS
        Retrieves all devices registered in Microsoft Graph (Entra ID).
    .DESCRIPTION
        Queries Microsoft Graph API to get a comprehensive list of all registered devices.
    .RETURNS
        An array of objects, each representing a device from Microsoft Graph.
    #>
    try {
        # Get all devices from Microsoft Graph.
        $OutputAllMSGraphAssets = Get-MgDevice -all -ErrorAction Stop
    }
    catch {
        # Output any errors encountered.
        $_
    }
    # Return the list of devices.
    Return $OutputAllMSGraphAssets
}

function Remove-IntuneDevicePrimaryUser {
    <#
    .SYNOPSIS
        Removes the primary user from an Intune managed device.
    .DESCRIPTION
        Sends a DELETE request to the Microsoft Graph API to disassociate the primary user
        from a specified Intune managed device, effectively converting it to a shared device.
    .PARAMETER IntuneDeviceId
        The unique ID of the Intune managed device from which to remove the primary user.
    #>
    [cmdletbinding()]
    
    param
    (
    [parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    $IntuneDeviceId
    )   
        # Define the Graph API version and resource path.
        $graphApiVersion = "v1.0"
        $Resource = "deviceManagement/managedDevices('$IntuneDeviceId')/users/`$ref"
        try {
            # Construct the URI and send a DELETE request.
            $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)" 
            Invoke-MgGraphRequest -Uri $uri -Method Delete -ErrorAction Stop
        }   
        catch { 
            # Output any errors encountered.
            $_
        }   
}

function Get-WinIntuneManagedDevices {
    <#
    .SYNOPSIS
        Retrieves Intune managed devices, optionally filtered by device name.
    .DESCRIPTION
        Queries the Microsoft Graph API for managed devices, specifically targeting
        Windows devices. It can filter by a specific device name.
    .PARAMETER deviceName
        (Optional) The name of the device to search for.
    .RETURNS
        An array of objects representing Intune managed devices.
    #>
    [cmdletbinding()] 
    param
    (
    [parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$deviceName
    )
    $graphApiVersion = "v1.0"
    try {
        # Construct the resource URL, applying a filter if a device name is provided.
        $Resource = "deviceManagement/managedDevices"
        if ($deviceName) {
            $Resource += "?`$filter=deviceName eq '$deviceName'"
        }
        $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"   
        # Invoke the Graph API request and return the value property.
        (Invoke-MgGraphRequest -Uri $uri -Method Get).value
    }
    catch {
        # Output any errors encountered.
        $_
    }
}
#endregion

#region Script Setup and Connection
function Start-D303DeviceSyncPrereqs {
    <#
    .SYNOPSIS
        Performs prerequisite checks and setup for the device synchronization script.
    .DESCRIPTION
        This function checks if the 'Microsoft.Graph' module is installed and imports it.
        If the module is not found, it attempts to install it. It also creates a directory
        for logs and sets up initial content for log files.
    #>
    # Check if Microsoft.Graph module is installed.
    if(!(Get-InstalledModule Microsoft.Graph)){
        try {
            # Set a higher maximum function count to avoid import issues with large modules.
            Set-Variable -Name MaximumFunctionCount -Value 32768 
            # Import the Microsoft.Graph module.
            Import-Module Microsoft.Graph -Confirm:$false -ErrorAction Stop
        }
        catch {
            # Log an error if the module fails to import and exit the script.
            Write-IntuneEventLog -EventID 502 -EventType "Error" -Message "Failed to Import Microsoft Graph API. $($_.Exception.Message)"
            Exit
        }
    }
    # Check if the IntuneSync directory exists in ProgramData, create if not.
    if(!(Test-Path $ENV:ProgramData/IntuneSync)){
        mkdir $ENV:ProgramData/IntuneSync
    }
    # Get current date for log file headers.
    $Date = Get-Date    
    # Initialize the MissingDevices.txt file with a header.
    Set-Content $ENV:ProgramData/IntuneSync/MissingDevices.txt "Missing Devices in Entra/AD - Sync Time: $Date"
    # Initialize the UnassignedDevices.txt file with a header.
    Set-Content $ENV:ProgramData/IntuneSync/UnassignedDevices.txt "Unassigned Devices in the Helpdesk - Sync Time: $Date"
}

function Start-D303ScriptConnections {
    <#
    .SYNOPSIS
        Establishes a connection to the Microsoft Graph API.
    .DESCRIPTION
        Connects to the Microsoft Graph API using application credentials (Tenant ID, Application ID,
        and Certificate Thumbprint). It checks if a connection already exists and logs success or failure.
    .PARAMETER TenantId
        The Azure AD Tenant ID.
    .PARAMETER ApplicationId
        The Azure AD Application ID (client ID) for the service principal.
    .PARAMETER Thumbprint
        The thumbprint of the certificate used for authentication.
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $TenantId,
        [Parameter(Position=1,mandatory=$true)]
        [string] $ApplicationId,
        [Parameter(Position=2,mandatory=$true)]
        [String] $Thumbprint
    )
    # Check if a Graph connection is already active.
    if(!(Get-MgUser -ErrorAction SilentlyContinue)){
        try {
            # Connect to Microsoft Graph using certificate authentication.
            Connect-MgGraph -TenantId $TenantId -ApplicationId $ApplicationId -CertificateThumbprint $Thumbprint -NoWelcome
            # Log successful connection.
            Write-IntuneEventLog -EventID 401 -EventType "Information" -Message "Successfully connected to Microsoft Graph API"
        }
        catch {
            # Log connection failure and exit the script.
            Write-IntuneEventLog -EventID 402 -EventType "Error" -Message "Failed to connect to Microsoft Graph API"
            Exit
        }
    }
    # Pause for a few seconds after connection.
    Start-sleep -seconds 4
}
#endregion

#region Device Management Functions
function Remove-AllDevicesPrimaryUser {
    <#
    .SYNOPSIS
        Removes the primary user from all applicable Intune managed devices.
    .DESCRIPTION
        Retrieves a list of all Intune managed devices that are Azure AD registered Windows devices
        with a primary user assigned. It then iterates through this list and attempts to remove
        the primary user, converting the device to a shared device.
    #>
    # Get all Intune managed devices.
    $AllD303MSGraphDevices = Get-MgDeviceManagementManagedDevice -all
    # Filter for Azure AD registered Windows devices with a primary user and specific OS version.
    $RegisteredDevicesWithPrimaryUser = $AllD303MSGraphDevices | Where-Object { (($Null -ne $_.azureADRegistered) -and ($_.operatingSystem -eq "Windows" ) -and ( "" -ne $_.userId) -and ($_.OSVersion -like "10.0.2*")) }
    
    # If no primary users are found, log information and return.
    if (!($RegisteredDevicesWithPrimaryUser.DeviceName)) {
        Write-IntuneEventLog -EventID 200 -EventType "Information" -Message "No primary users to remove"
        Return
    }
    
    # Iterate through each device with a primary user.
    foreach ($Device in $RegisteredDevicesWithPrimaryUser){
        if($Device){
            try {
                # Attempt to remove the primary user.
                $DeleteIntuneDevicePrimaryUser = Remove-IntuneDevicePrimaryUser -IntuneDeviceId $Device.id -ErrorAction Stop
                # Log successful primary user removal.
                Write-IntuneEventLog -EventID 201 -EventType "Information" -Message "Successfully deleted primary user from $($Device.DeviceName) and converted to shared device"
            }
            catch {
                # Log failure to delete primary user.
                Write-IntuneEventLog -EventID 202 -EventType "Error" -Message "Failed to delete primary user from $($Device.DeviceName). $($_.Exception.Message)"
            }
            # This check is redundant as Remove-IntuneDevicePrimaryUser returns nothing on success.
            # if($DeleteIntuneDevicePrimaryUser -eq ""){
            #    #Write-Host "User deleted as Primary User from the device '$($Device.DeviceName)'..." -ForegroundColor Green
            # }
        }
        else {
            # Log warning if a device cannot be found to delete its primary user.
            Write-IntuneEventLog -EventID 203 -EventType "Warning" -Message "Unable to find $($Device.DeviceName) to delete primary user."
        }
    }
}

function Get-IntuneDevicePrimaryUser {
    <#
    .SYNOPSIS
        Retrieves the primary user ID of an Intune managed device.
    .DESCRIPTION
        This function is intended to retrieve the primary user associated with an Intune managed device.
        However, the current implementation attempts to retrieve users associated with the device,
        which might not directly correspond to the 'primary user' concept in Intune's simplified shared device model.
        The function returns the 'id' property of the first user found.
    .PARAMETER deviceId
        The unique ID of the Intune managed device.
    .RETURNS
        The ID of the primary user, or an empty string if not found or an error occurs.
    #>
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [string] $deviceId
    )
    $graphApiVersion = "v1.0"
    $Resource = "deviceManagement/managedDevices"
    $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)" + "/" + $deviceId + "/users"

    try {
        # Invoke Graph API to get users associated with the device.
        $primaryUser = Invoke-MgGraphRequest -uri $uri -Method GET
        # Extract the value property.
        $Ptest = $primaryUser.value
        # This line `$Ptest.count` is for debugging and doesn't affect the return.
        $Ptest.count
        # Return the ID of the first user found.
        return $primaryUser.value."id"
    } catch {
        # Output any errors encountered.
        $_
    }
}

function Remove-DuplicateComputersAD {
    <#
    .SYNOPSIS
        Removes a computer object from Active Directory.
    .DESCRIPTION
        Attempts to delete a specified computer object from Active Directory.
        This function is typically used for removing stale or duplicate entries.
    .PARAMETER DeviceName
        The name of the computer object to be removed from Active Directory.
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $DeviceName
    )
    try {
        # Get the Active Directory computer object.
        $ADComputerAccount = Get-ADComputer -Identity $DeviceName -ErrorAction Stop #-Credential $ADCreds
        # Remove the Active Directory object recursively and without confirmation.
        Remove-ADObject -Identity $ADComputerAccount -Recursive -Server "adds3.ad1.d303.org" -Confirm:$false -ErrorAction Stop #-Credential $ADCreds
        # Log successful deletion.
        Write-IntuneEventLog -EventID 601 -EventType "Information" -Message "Successfully deleted $DeviceName from Active Directory."
    }
    catch {
        # Log failure to delete the computer from Active Directory.
        Write-IntuneEventLog -EventID 602 -EventType "Error" -Message "Failed to delete $DeviceName from Active Directory. $($_.Exception.Message)"
    }
}

function Sync-ComputerNames {
    <#
    .SYNOPSIS
        Synchronizes computer names between Active Directory and Helpdesk/Intune.
    .DESCRIPTION
        This function identifies misnamed computers in a specific Active Directory OU (starting with "AP-"),
        checks their corresponding entries in Intune and the helpdesk, and attempts to:
        - Remove computers from AD if they are not found in Intune or are too old.
        - Rename computers in AD if their name does not match the helpdesk asset tag and they are reachable.
        It prioritizes renaming over deletion for existing devices and cleans up older duplicate AD entries.
    #>
    # Define the Active Directory OU where AutoPilot computers are located.
    $OUpath = 'OU=AutoPilot,OU=Computers,OU=SD303,DC=ad1,DC=d303,DC=org'
    # Get Active Directory computers in the specified OU that have names starting with "AP-".
    $APcomputers = Get-ADComputer -Filter * -SearchBase $OUpath -Properties * | Where-Object {$_.DNSHostName -like "AP-*"}
    
    # Get current date and subtract 2 hours for comparison with last modified timestamps.
    $DateTimeNow = Get-Date 
    $DateTimeNow = $DateTimeNow.AddHours(-2)
    $DateTimeNow = $DateTimeNow.ToString("yyyy/MM/dd HH:mm:ss")
    
    # If no misnamed computers are found, log and exit.
    if($Null -eq $APcomputers){
        Write-IntuneEventLog -EventID 607 -EventType "Information" -Message "No misnamed devices in Active Directory"
        return
    }
    
    # Initialize an array to store custom device information.
    [array]$DeviceTable = @()

    # Iterate through each "AP-" named device found in Active Directory.
    foreach ($Device in $APcomputers) {
        $DeviceName = $Device.DNSHostName
        # Extract the short computer name (remove domain suffix).
        $DeviceName = $DeviceName.Substring(0, $DeviceName.IndexOf('.'))
        
        # Get Intune device information based on the Active Directory device name.
        $IntuneDeviceInfo = Get-WinIntuneManagedDevices -deviceName $DeviceName
        
        # If the device is not found in Intune.
        if($Null -eq $IntuneDeviceInfo){
            Write-IntuneEventLog -EventID 209 -EventType "Warning" -Message "Could not find $DeviceName in Intune. Attempting to remove device from Active Directory"
            $CurrentTime = get-date -Format "MM/dd/yyyy hh:mm:ss tt"
            # Calculate time difference since last change in AD.
            $TimeDiff = New-TimeSpan -Start $Device.whenChanged -End $CurrentTime
            # If the device hasn't been changed in AD for at least 1 day, remove it.
            if($TimeDiff.Days -ge 1){
                Remove-DuplicateComputersAD -DeviceName $DeviceName
            }
            else {
                # Log if the device is too new to be removed.
                Write-IntuneEventLog -EventID 606 -EventType "Warning" -Message "$DeviceName has been modified too recently, will retry next sync."
            }
        }
        else{
            # If found in Intune, get helpdesk device information by serial number.
            $HelpdeskDeviceInfo = Get-HelpdeskDevice -HelpdeskAPIkey $SchoolzaAPIkey -serialnumber $IntuneDeviceInfo.serialNumber
            # Create a custom object with relevant device information.
            $CustomDevicesInfo = @{
                Name            = $DeviceName
                SerialNumber    = $HelpdeskDeviceInfo.serial_number
                HelpdeskName    = $HelpdeskDeviceInfo.asset_tag
                CreatedDate     = [DateTime]$Device.createTimeStamp # AD creation timestamp
            }
            # Add the custom object to the device table.
            [array]$DeviceTable += [pscustomobject]$CustomDevicesInfo
        }
    }
    
    # Get unique serial numbers from the collected device data.
    $UniqueSerial = $DeviceTable.SerialNumber | Select-Object -Unique
    
    # Iterate through each unique serial number to handle renaming and duplicate cleanup.
    foreach ($Serial in $UniqueSerial) {
        # Find all entries in the device table that match the current serial number.
        $MatchingSerials = $DeviceTable | Where-Object {$_.SerialNumber -eq $Serial}
        # Get the enrollment time of the matching serials (assuming it's consistent for duplicates).
        $EnrollTime = $MatchingSerials.CreatedDate.tostring("yyyy/MM/dd HH:mm:ss")
        
        # If there's only one entry for this serial number (no duplicates in AD, but might be misnamed).
        if ($MatchingSerials.name.count -lt 2) {
            # Check if the device was enrolled before the 2-hour threshold.
            if ($EnrollTime -lt $DateTimeNow){
                # If the AD name doesn't match the helpdesk name.
                if ($MatchingSerials.Name -ne $MatchingSerials.HelpdeskName) {
                    # Check if the computer is reachable.
                    if(Test-Connection -ComputerName $MatchingSerials.Name -ErrorAction SilentlyContinue){
                        try {
                            # Attempt to rename the computer in AD (using -WhatIf for testing during development).
                            Rename-Computer -ComputerName $MatchingSerials.Name -NewName $MatchingSerials.HelpdeskName -Confirm:$false -Force -Restart -ErrorAction Stop #-WhatIf
                            # Log successful renaming.
                            Write-IntuneEventLog -EventID 603 -EventType "Information" -Message "Successfully renamed computer from $($MatchingSerials.Name) to $($MatchingSerials.HelpdeskName)."
                        }
                        catch {
                            # Log failure to rename the computer.
                            Write-IntuneEventLog -EventID 604 -EventType "Error" -Message "Failed to rename computer from $($MatchingSerials.Name) to $($MatchingSerials.HelpdeskName). $($_.Exception.Message)"
                        }
                    }
                    else {
                        # Log if the computer is unreachable.
                        Write-IntuneEventLog -EventID 605 -EventType "Warning" -Message "Failed to connect to $($MatchingSerials.Name) will try again next cycle."
                    }
                }
            }
            else {
                # Log if the device was enrolled too recently to be processed for renaming.
                Write-IntuneEventLog -EventID 606 -EventType "Warning" -Message "$($MatchingSerials.Name) was enrolled too recently will try again next cycle."
            }
            continue # Move to the next unique serial number.
        }

        # If there are multiple entries for the same serial number (duplicates in AD).
        foreach ($device in $MatchingSerials) {
            if($MatchingSerials.Name.Count -gt 1){
                # Identify the oldest duplicate based on creation date.
                $OldAsset = $MatchingSerials | Foreach-Object {$_.CreatedDate; $_} | Group-Object -Property {$_.SerialNumber} | Foreach-Object {$_.group | Sort-Object CreatedDate | Select-Object -First 1}
                try {
                    # Attempt to remove the oldest duplicate from Active Directory.
                    Remove-DuplicateComputersAD -DeviceName $OldAsset.Name -ErrorAction Stop
                }
                catch {
                    # Output error if deletion fails.
                    $_
                    # Remove the failed deletion asset from the current matching serials to avoid re-processing.
                    $MatchingSerials = $MatchingSerials | Where-Object{$_.Name -ne $OldAsset.Name}
                }       
                # Remove the oldest asset from the current matching serials.
                $MatchingSerials = $MatchingSerials | Where-Object{$_.Name -ne $OldAsset.Name}
            }
            else { # This 'else' block seems to be a fallback for a single remaining entry after duplicate removal.
                # Check if the device was enrolled before the 2-hour threshold.
                if ($MatchingSerials.CreatedDate.datetime -lt $DateTimeNow.AddHours(-2).DateTime){
                    # If the AD name doesn't match the helpdesk name.
                    if ($MatchingSerials.Name -ne $MatchingSerials.HelpdeskName) {
                        # Check if the computer is reachable.
                        if(Test-Connection -ComputerName $MatchingSerials.Name -ErrorAction SilentlyContinue){
                            try {
                                # Attempt to rename the computer.
                                Rename-Computer -ComputerName $MatchingSerials.Name -NewName $MatchingSerials.HelpdeskName -Confirm:$false -Force -Restart -ErrorAction Stop #-WhatIf
                                # Log successful renaming.
                                Write-IntuneEventLog -EventID 603 -EventType "Information" -Message "Successfully renamed computer from $($MatchingSerials.Name) to $($MatchingSerials.HelpdeskName)."
                            }
                            catch {
                                # Log failure to rename.
                                Write-IntuneEventLog -EventID 604 -EventType "Error" -Message "Failed to rename computer from $($MatchingSerials.Name) to $($MatchingSerials.HelpdeskName). $($_.Exception.Message)" 
                            }
                        }
                        else {
                            # Log if the computer is unreachable.
                            Write-IntuneEventLog -EventID 605 -EventType "Warning" -Message "Failed to connect to $($MatchingSerials.Name) will try again next cycle."
                        }
                    }
                    else {
                        # Log if the computer is already named correctly.
                        Write-IntuneEventLog -EventID 607 -EventType "Information" -Message "Computer already named correctly, cleaned up artifact devices"
                    }
                }
                else {
                    # Log if the device was enrolled too recently.
                    Write-IntuneEventLog -EventID 606 -EventType "Warning" -Message "$($MatchingSerials.Name) was enrolled too recently will try again next cycle."
                }
            }
        }
    }   
}

function Remove-RetiredDevices {
    <#
    .SYNOPSIS
        Removes retired devices from Autopilot deployment profiles in Intune.
    .DESCRIPTION
        Retrieves a list of all devices registered in Autopilot profiles in Intune.
        It then compares this list with devices marked as "retired" in the helpdesk system.
        Any devices found in Autopilot that are also marked as retired in the helpdesk
        are then removed from their Autopilot deployment profiles.
    #>
    $graphApiVersion = "v1.0"
    try {
        $Resource = "deviceManagement/windowsAutopilotDeviceIdentities"
        $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)" 

        # Get the initial list of Autopilot devices.
        $DeviceList = Invoke-MgGraphRequest -Uri $uri -Method Get
    } catch {
        # Output any errors encountered during initial device retrieval.
        $_
    }
    # Initialize an array to store the full list of Autopilot devices.
    [array]$FullDeviceList += $DeviceList.value
    
    # Handle pagination to get all Autopilot devices if there are more than one page of results.
    while ($DeviceList.'@odata.nextLink') {
        try {
            $DeviceList = Invoke-MgGraphRequest -Uri $DeviceList.'@odata.nextLink' -Method Get
            [array]$FullDeviceList += $DeviceList.value
        } catch {
            $_
        }
    }
    # Get the list of retired devices from the helpdesk.
    $RetiredDevices = Get-HelpdeskRetiredDevice -HelpdeskAPIkey $SchoolzaAPIkey
    # Filter out retired devices that don't have a serial number.
    $RetiredDevices = $RetiredDevices | Where-Object {"" -ne $_.serial_number }
    
    # Create a copy of the full device list for comparison.
    $DeviceCompareList = $FullDeviceList
    
    # Compare the serial numbers of Autopilot devices with retired devices from the helpdesk.
    # The '==' SideIndicator means the serial number is present in both lists.
    $MatchedObjects = Compare-Object -ReferenceObject $DeviceCompareList.serialNumber -DifferenceObject $RetiredDevices.serial_number -IncludeEqual
    $DevicestoRemove = $MatchedObjects | Where-Object {$_.SideIndicator -eq '=='} 

    # Iterate through each serial number that needs to be removed from Autopilot.
    foreach ($SerialNumber in $DevicestoRemove.InputObject) {
        $Device = ""
        # Find the full device object in the Autopilot list using the serial number.
        $Device = $FullDeviceList | Where-Object {$_.SerialNumber -eq $SerialNumber}
        try {
            # Define the resource and URI for deleting the Autopilot device.
            $Resourcedelete = "/deviceManagement/windowsAutopilotDeviceIdentities/$($Device.id)"
            $uridelete = "https://graph.microsoft.com/$graphApiVersion/$($ResourceDelete)" 
            # Send a DELETE request to remove the device from Autopilot.
            Invoke-MgGraphRequest -Uri $uridelete -Method DELETE -ErrorAction Stop
            # Log successful removal.
            Write-IntuneEventLog -EventID 204 -EventType "Information" -Message "Successfully deleted serial number: $($Device.serialNumber) from Autopilot."
        } catch {
            # Log failure to remove the device from Autopilot.
            Write-IntuneEventLog -EventID 205 -EventType "Error" -Message "Failed to deleted serial number: $($Device.serialNumber) from Autopilot. $($_.Exception.Message)"
        }
    }
}

function Start-DeviceGroupSync {
    <#
    .SYNOPSIS
        Synchronizes device membership for predefined lab groups in Entra ID.
    .DESCRIPTION
        This function iterates through a list of specified lab group names.
        For each group, it retrieves the current members from Entra ID and the corresponding
        devices from the helpdesk system based on room information.
        It then adds any devices found in the helpdesk for that room but missing from the Entra group,
        and removes any devices from the Entra group that are no longer in the helpdesk for that room.
    #>
    # Get all room IDs from the helpdesk system.
    $HelpdeskRoomIDs = Get-HelpdeskRoomIDs -HelpdeskAPIkey $SchoolzaAPIkey
    
    # Iterate through each predefined lab group name.
    foreach ($group in $LabGroupNames) {   
        $GroupMSGraphAssets = @()
        # Get the Group ID for the current lab group.
        $LabID = Get-D303MsGraphGroupIDbyName -GroupName $group
        
        # Split the group name to extract building and room number information.
        $SplitGroupName = $group -split "-"
        $Building = $SplitGroupName[2]
        $RoomNumber = $SplitGroupName[3].trim(" ","R","M")
        
        # Match the building and room number to a helpdesk room ID.
        $MatchedRoomID = Get-D303LabMatch -Building $Building -RoomNumber $RoomNumber -RoomData $HelpdeskRoomIDs
        
        $RoomAssets = @()
        # Get devices from the helpdesk for the matched room.
        $RoomAssets = Get-HelpdeskDevicesByRoom -RoomID $MatchedRoomID.room_id -HelpdeskAPIkey $SchoolzaAPIkey
        
        # Get current members (display names) of the Entra group.
        $GroupMSGraphAssets = Get-D303MsGraphGroupMembers -GroupId $LabID
        # Sort and get unique asset names.
        $GroupMSGraphAssets = $GroupMSGraphAssets | Sort-Object | Get-Unique -AsString
        
        # If the Entra group currently has no members.
        if($Null -eq $GroupMSGraphAssets){
            # Add all devices from the helpdesk room to the Entra group.
            foreach ($asset in $RoomAssets) {
                # Get the Entra ID for the asset by display name.
                $EntraID = Get-MgDevice -Filter "startswith(displayName, '$asset')"
                try {
                    # Iterate through found Entra IDs (there might be multiple if duplicates exist, though less common for devices).
                    foreach ($id in $($EntraID.Id)) {
                        # Add the device to the Entra group.
                        New-MgGroupMember -GroupId $LabID -DirectoryObjectId $id -ErrorAction stop
                        # Log successful addition.
                        Write-IntuneEventLog -EventID 301 -EventType "Information" -Message "Added $asset to $group."
                    }
                }
                catch {           
                    # Log error if adding fails and add to MissingDevices.txt.
                    Write-IntuneEventLog -EventID 303 -EventType "Error" -Message "Failed to add $asset to $group. $($_.Exception.Message)"
                    Add-Content -Path "$ENV:ProgramData/IntuneSync/MissingDevices.txt" -Value "$asset - $group"
                }           
            }
        }
        else {
            # Find devices that are in the helpdesk room but missing from the Entra group.
            $MissingHelpdeskAssets = @()
            $MissingHelpdeskAssets = Compare-Object -ReferenceObject $RoomAssets -DifferenceObject $GroupMSGraphAssets | Select-Object -ExpandProperty InputObject
            
            # Add missing devices to the Entra group.
            foreach ($asset in $MissingHelpdeskAssets) {
                $EntraID = Get-MgDevice -Filter "startswith(displayName, '$asset')"
                try {
                    foreach ($mid in $($EntraID.Id)) {
                        New-MgGroupMember -GroupId $LabID -DirectoryObjectId $mid -ErrorAction stop
                        Write-IntuneEventLog -EventID 301 -EventType "Information" -Message "Added $asset to $group."
                    }
                }
                catch {           
                    Write-IntuneEventLog -EventID 303 -EventType "Error" -Message "Failed to add $asset to $group. $($_.Exception.Message)" 
                    Add-Content -Path "$ENV:ProgramData/IntuneSync/MissingDevices.txt" -Value "$asset - $group"
                }           
            }
            # Find devices that are in the Entra group but no longer in the helpdesk room.
            $InvaildAssets = @()
            $InvaildAssets = Compare-Object -ReferenceObject $GroupMSGraphAssets -DifferenceObject $RoomAssets | Where-Object {$_ -like "*<=*"} | Select-Object -ExpandProperty InputObject
            
            # Remove invalid devices from the Entra group.
            foreach ($asset in $InvaildAssets) {
                $EntraID = Get-MgDevice -Filter "startswith(displayName, '$asset')"
                try {
                    foreach ($rid in $($EntraID.Id)) {
                        Remove-MgGroupMemberByRef -GroupId $LabID -DirectoryObjectId $rid -ErrorAction stop
                        Write-IntuneEventLog -EventID 302 -EventType "Information" -Message "Deleted $asset from $group."
                    }
                }
                catch {           
                    Write-IntuneEventLog -EventID 304 -EventType "Error" -Message "Failed to delete $asset from $group. $($_.Exception.Message)" 
                }           
            }
        }
    }
}

function Invoke-IntuneCleanup {
    <#
    .SYNOPSIS
        Cleans up duplicate device entries in Intune.
    .DESCRIPTION
        This function identifies devices in Intune that have duplicate serial numbers.
        For each set of duplicates, it retains the device with the most recent last sync time
        and removes the older, duplicate entries from Intune.
    #>
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param()
    Begin {
    }
    Process {
        # Get all Intune managed devices.
        $devices = Get-MgDeviceManagementManagedDevice -All
        Write-Verbose "Found $($devices.Count) devices."
        
        # Group devices by serial number, excluding empty or "Defaultstring" serial numbers.
        $deviceGroups = $devices | Where-Object { -not [String]::IsNullOrWhiteSpace($_.serialNumber) -and ($_.serialNumber -ne "Defaultstring") } | Group-Object -Property serialNumber
        # Filter for groups that have more than one device (duplicates).
        $duplicatedDevices = $deviceGroups | Where-Object {$_.Count -gt 1 }
        Write-Verbose "Found $($duplicatedDevices.Count) serialNumbers with duplicated entries"
        # Log the number of duplicate serial numbers found.
        Write-IntuneEventLog -EventID 206 -EventType "Information" -Message "Found $($duplicatedDevices.Count) serial numbers with duplicate entries in Intune."
        
        # Iterate through each set of duplicated devices.
        foreach($duplicatedDevice in $duplicatedDevices){
            # Find the newest device in the group based on last sync time.
            $newestDevice = $duplicatedDevice.Group | Sort-Object -Property lastSyncDateTime -Descending | Select-Object -First 1
            Write-Verbose "Serial $($duplicatedDevice.Name)"
            Write-Verbose "Keep $($newestDevice.deviceName) $($newestDevice.lastSyncDateTime)"
            
            # Iterate through the older duplicate devices (all except the newest).
            foreach($oldDevice in ($duplicatedDevice.Group | Sort-Object -Property lastSyncDateTime -Descending | Select-Object -Skip 1)){
                Write-Verbose "Remove $($oldDevice.deviceName) $($oldDevice.lastSyncDateTime)"
                try {
                    # Remove the older duplicate device from Intune.
                    Remove-MgDeviceManagementManagedDevice -managedDeviceId $oldDevice.id -ErrorAction Stop
                    # Log successful removal of duplicate.
                    Write-IntuneEventLog -EventID 207 -EventType "Information" -Message "Successfully removed duplicate device $($oldDevice.deviceName) with the last sync time $($oldDevice.lastSyncDateTime) from Intune."
                }
                catch {
                    # Log failure to remove duplicate.
                    Write-IntuneEventLog -EventID 208 -EventType "Error" -Message "Failed removed duplicate device $($oldDevice.deviceName) with the last sync time $($oldDevice.lastSyncDateTime) from Intune. $($_.Exception.Message)"
                }
            }
        }
        # The commented-out section below appears to be an alternative or older approach for handling
        # duplicates based on display name rather than serial number. It's currently not active.
        # $MGDevices = Get-MgDevice -All
        # Write-Verbose "Found $($MGDevices.Count) devices."
        # $DeviceDuplicateNames = $MGDevices | Group-Object DisplayName | Where-Object {$_.Count -gt 1 }
        # Write-Verbose "Found $($DeviceDuplicateNames.Count) Devices with duplicate Name entries in Intune."
        # Write-IntuneEventLog -EventID 206 -EventType "Information" -Message "Found $($DeviceDuplicateNames.Count) Devices with duplicate Name entries in Intune."
        # foreach($duplicatedDevice in $DeviceDuplicateNames){
        #     # Find device which is the newest.
        #     $newestDevice = $duplicatedDevice.Group | Sort-Object -Property {$_.AdditionalProperties.createdDateTime} -Descending | Select-Object -First 1
        #     Write-Verbose "Keep $($newestDevice.DisplayName) $($newestDevice."AdditionalProperties"."createdDateTime")"
        #     foreach($oldDevice in ($duplicatedDevice.Group | Sort-Object -Property {$_.AdditionalProperties.createdDateTime} -Descending | Select-Object -Skip 1)){
        #         Write-Verbose "Remove $($oldDevice.DisplayName) $($oldDevice."AdditionalProperties"."createdDateTime")"
        #         try {
        #             Remove-MgDevice -DeviceId $oldDevice.id -ErrorAction Stop
        #             Write-IntuneEventLog -EventID 207 -EventType "Information" -Message "Successfully removed duplicate device $($oldDevice.DisplayName) with the last sync time $($oldDevice."AdditionalProperties"."createdDateTime") from Intune."
        #         }
        #         catch {
        #             Write-IntuneEventLog -EventID 208 -EventType "Error" -Message "Failed removed duplicate device $($oldDevice.DisplayName) with the last sync time $($oldDevice."AdditionalProperties"."createdDateTime") from Intune. $($_.Exception.Message)"
        #         }  
        #     }
        # }

    }
    End {
    }
}

function Start-DeviceGroupSyncUnassigned {
    <#
    .SYNOPSIS
        Synchronizes unassigned devices from the helpdesk to a specified Entra ID group.
    .DESCRIPTION
        This function identifies devices in the helpdesk marked as "available" (unassigned)
        and synchronizes their membership with a designated Entra ID group.
        It adds devices to the group if they are unassigned in the helpdesk but not in the group,
        and removes devices from the group if they are no longer marked as unassigned in the helpdesk.
    .PARAMETER GroupName
        The display name of the Entra ID group for unassigned devices (e.g., "Intune-Unassigned").
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $GroupName
    )
    $GroupMSGraphAssets = @()
    $Group = $GroupName
    # Get the Group ID for the specified unassigned group.
    $GroupID = Get-D303MsGraphGroupIDbyName -GroupName $GroupName
    
    # Get unassigned devices from the helpdesk.
    $UnassignedAssets = Get-HelpdeskUnassignedDevices -HelpdeskAPIkey $SchoolzaAPIkey
    
    # Get current members (display names) of the Entra group.
    $GroupMSGraphAssets = Get-D303MsGraphGroupMembers -GroupId $GroupID
    $GroupMSGraphAssets = $GroupMSGraphAssets | Sort-Object | Get-Unique -AsString

    # If the Entra group currently has no members.
    if($Null -eq $GroupMSGraphAssets){
        # Add all unassigned devices to the Entra group.
        foreach ($asset in $UnassignedAssets) {
            # Get the Entra ID for the asset.
            $EntraID = Get-MgDevice -Filter "startswith(displayName, '$asset')"
            try {
                foreach ($id in $($EntraID.Id)) {
                    New-MgGroupMember -GroupId $GroupID -DirectoryObjectId $id -ErrorAction stop
                    Write-IntuneEventLog -EventID 301 -EventType "Information" -Message "Added $asset to $group."
                    # Also log to a file for review.
                    Add-Content -Path "$ENV:ProgramData/IntuneSync/UnassignedDevices.txt" -Value "$asset - $group"
                }
            }
            catch {           
                Write-IntuneEventLog -EventID 303 -EventType "Error" -Message "Failed to add $asset to $group. $($_.Exception.Message)"
                Add-Content -Path "$ENV:ProgramData/IntuneSync/MissingDevices.txt" -Value "$asset - $group"
            }           
        }
    }
    else {
        # Find devices that are unassigned in the helpdesk but missing from the Entra group.
        $MissingHelpdeskAssets = Compare-Object -ReferenceObject $UnassignedAssets -DifferenceObject $GroupMSGraphAssets | Select-Object -ExpandProperty InputObject
        foreach ($asset in $MissingHelpdeskAssets) {
            $EntraID = Get-MgDevice -Filter "startswith(displayName, '$asset')"
            try {
                foreach ($mid in $($EntraID.Id)) {
                    New-MgGroupMember -GroupId $GroupID -DirectoryObjectId $mid -ErrorAction stop
                    Write-IntuneEventLog -EventID 301 -EventType "Information" -Message "Added $asset to $group."
                    Add-Content -Path "$ENV:ProgramData/IntuneSync/UnassignedDevices.txt" -Value "$asset - $group"
                }
            }
            catch {           
                Write-IntuneEventLog -EventID 303 -EventType "Error" -Message "Failed to add $asset to $group. $($_.Exception.Message)" 
                Add-Content -Path "$ENV:ProgramData/IntuneSync/MissingDevices.txt" -Value "$asset - $group"

            }           
        }
        # Find devices that are in the Entra group but no longer unassigned in the helpdesk.
        $InvaildAssets = Compare-Object -ReferenceObject $GroupMSGraphAssets -DifferenceObject $UnassignedAssets | Where-Object {$_ -like "*<=*"} | Select-Object -ExpandProperty InputObject
        foreach ($asset in $InvaildAssets) {
            $EntraID = Get-MgDevice -Filter "startswith(displayName, '$asset')"
            try {
                foreach ($rid in $($EntraID.Id)) {
                    Remove-MgGroupMemberByRef -GroupId $GroupID -DirectoryObjectId $rid -ErrorAction stop
                    Write-IntuneEventLog -EventID 302 -EventType "Information" -Message "Deleted $asset from $group."
                }
            }
            catch {           
                Write-IntuneEventLog -EventID 304 -EventType "Error" -Message "Failed to delete $asset from $group. $($_.Exception.Message)" 
            }           
        }
        # Clear temporary variables.
        $InvaildAssets = @()
        $MissingHelpdeskAssets = @()
    }
}

function Start-DeviceGroupSyncCompleted {
    <#
    .SYNOPSIS
        Synchronizes completed Autopilot devices to a specified Entra ID group.
    .DESCRIPTION
        This function identifies Windows 10/11 devices in Entra ID that are Intune-managed,
        are registered, have names matching a specific pattern (e.g., LDXX-XXXX),
        and have been registered for at least 3 hours.
        These "completed" devices are then synchronized with a designated Entra ID group.
        It adds devices to the group if they meet the criteria but are not in the group,
        and removes devices from the group if they no longer meet the criteria.
    .PARAMETER CompletedGroupName
        The display name of the Entra ID group for completed Autopilot devices (e.g., "AutoPilot-Completed").
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $CompletedGroupName
    )
    $GroupMSGraphAssets = @()
    $group = $CompletedGroupName
    # Get the Group ID for the specified completed group.
    $GroupID = Get-D303MsGraphGroupIDbyName -GroupName $CompletedGroupName 
    
    # Get all Entra devices filtered by OS version and specific management properties.
    $CompleteDev = Get-MgDevice -Filter "startswith(OperatingSystemVersion, '10.0.2')" -all
    $CompleteDev = $CompleteDev | where-object {($_.DisplayName -match '[lLdD]\d\d-\d\d\d\d$') -and ($_.IsManaged -eq "True") -and ((($_.AdditionalProperties.managementType -eq "MDM") -or ($_.managementType -eq "MDM")) -and ($_.ProfileType -eq "RegisteredDevice"))}
    
    $currentDate = Get-Date
    $CompletedAssets = @()
    # Filter for devices registered at least 3 hours ago.
    $CompletedAssets = $CompleteDev | Where-Object { ($currentDate - $_.RegistrationDateTime).TotalHours -ge 3 }
    $CompletedAssets = $CompletedAssets.DisplayName 
    
    # Get current members (display names) of the Entra group.
    $GroupMSGraphAssets = Get-D303MsGraphGroupMembers -GroupId $GroupID
    $GroupMSGraphAssets = $GroupMSGraphAssets | Sort-Object | Get-Unique -AsString
    
    # If the Entra group currently has no members.
    if($Null -eq $GroupMSGraphAssets){
        # Add all completed devices to the Entra group.
        foreach ($asset in $CompletedAssets) {
            $EntraID = Get-MgDevice -Filter "startswith(displayName, '$asset')"
            try {
                foreach ($id in $($EntraID.Id)) {
                    New-MgGroupMember -GroupId $GroupID -DirectoryObjectId $id -ErrorAction stop
                    Write-IntuneEventLog -EventID 301 -EventType "Information" -Message "Added $asset to $group."
                }
            }
            catch {           
                Write-IntuneEventLog -EventID 303 -EventType "Error" -Message "Failed to add $asset to $group. $($_.Exception.Message)"
            }           
        }
    }
    else {
        # Find devices that meet the "completed" criteria but are missing from the Entra group.
        $MissingCompletedAssets = Compare-Object -ReferenceObject $CompletedAssets -DifferenceObject $GroupMSGraphAssets | Select-Object -ExpandProperty InputObject
        foreach ($asset in $MissingCompletedAssets) {
            $EntraID = Get-MgDevice -Filter "startswith(displayName, '$asset')"
            try {
                foreach ($mid in $($EntraID.Id)) {
                    New-MgGroupMember -GroupId $GroupID -DirectoryObjectId $mid -ErrorAction stop
                    Write-IntuneEventLog -EventID 301 -EventType "Information" -Message "Added $asset to $group."
                }
            }
            catch {           
                Write-IntuneEventLog -EventID 303 -EventType "Error" -Message "Failed to add $asset to $group. $($_.Exception.Message)" 
            }           
        }
        # Find devices that are in the Entra group but no longer meet the "completed" criteria.
        $InvaildAssets = Compare-Object -ReferenceObject $GroupMSGraphAssets -DifferenceObject $CompletedAssets | Where-Object {$_ -like "*<=*"} | Select-Object -ExpandProperty InputObject
        foreach ($asset in $InvaildAssets) {
            $EntraID = Get-MgDevice -Filter "startswith(displayName, '$asset')"
            try {
                foreach ($rid in $($EntraID.Id)) {
                    Remove-MgGroupMemberByRef -GroupId $GroupID -DirectoryObjectId $rid -ErrorAction stop
                    Write-IntuneEventLog -EventID 302 -EventType "Information" -Message "Deleted $asset from $group."
                }
            }
            catch {           
                Write-IntuneEventLog -EventID 304 -EventType "Error" -Message "Failed to delete $asset from $group. $($_.Exception.Message)" 
            }           
        }
        # Clear temporary variables.
        $InvaildAssets = @()
        $MissingCompletedAssets = @()
    }
}

function Start-DeviceGroupSyncBuilding {
    <#
    .SYNOPSIS
        Synchronizes devices within a specific building from the helpdesk to an Entra ID group.
    .DESCRIPTION
        This function retrieves devices associated with a given building from the helpdesk system.
        It then synchronizes their membership with a designated Entra ID group.
        It adds devices to the group if they are in the specified building but not in the group,
        and removes devices from the group if they are no longer in that building according to the helpdesk.
    .PARAMETER GroupName
        The display name of the Entra ID group to synchronize (e.g., "AutoPatch-Pilot").
    .PARAMETER Building
        The name or partial name of the building to get devices from (e.g., "ITS Peck Road").
    #>
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string] $GroupName,
        [Parameter(Position=1,mandatory=$true)]
        [string] $Building
    )
    $GroupMSGraphAssets = @()
    $Group = $GroupName
    # Get the Group ID for the specified group.
    $GroupID = Get-D303MsGraphGroupIDbyName -GroupName $GroupName
    
    # Get devices by building from the helpdesk.
    $BuildingAssets = Get-HelpdeskDevicesByBuilding -HelpdeskAPIkey $SchoolzaAPIkey -Building $Building
    
    # If the building was not found in the helpdesk, log an error and return.
    if ($BuildingAssets -eq "NotFound") {
        Write-IntuneEventLog -EventID 305 -EventType "Error" -Message "Couldn't find building by name, Please check building name."
        Return 404
    }
    
    # Get current members (display names) of the Entra group.
    $GroupMSGraphAssets = Get-D303MsGraphGroupMembers -GroupId $GroupID
    $GroupMSGraphAssets = $GroupMSGraphAssets | Sort-Object | Get-Unique -AsString

    # If the Entra group currently has no members.
    if($Null -eq $GroupMSGraphAssets){
        # Add all devices from the building to the Entra group.
        foreach ($asset in $BuildingAssets) {
            $EntraID = Get-MgDevice -Filter "startswith(displayName, '$asset')"
            try {
                foreach ($id in $($EntraID.Id)) {
                    New-MgGroupMember -GroupId $GroupID -DirectoryObjectId $id -ErrorAction stop
                    Write-IntuneEventLog -EventID 301 -EventType "Information" -Message "Added $asset to $group."
                }
            }
            catch {           
                Write-IntuneEventLog -EventID 303 -EventType "Error" -Message "Failed to add $asset to $group. $($_.Exception.Message)"
            }           
        }
    }
    else {
        # Find devices that are in the building but missing from the Entra group.
        $MissingHelpdeskAssets = Compare-Object -ReferenceObject $BuildingAssets -DifferenceObject $GroupMSGraphAssets | Select-Object -ExpandProperty InputObject
        foreach ($asset in $MissingHelpdeskAssets) {
            $EntraID = Get-MgDevice -Filter "startswith(displayName, '$asset')"
            try {
                foreach ($mid in $($EntraID.Id)) {
                    New-MgGroupMember -GroupId $GroupID -DirectoryObjectId $mid -ErrorAction stop
                    Write-IntuneEventLog -EventID 301 -EventType "Information" -Message "Added $asset to $group."
                }
            }
            catch {           
                Write-IntuneEventLog -EventID 303 -EventType "Error" -Message "Failed to add $asset to $group. $($_.Exception.Message)" 
            }           
        }
        # Find devices that are in the Entra group but no longer in the specified building.
        $InvaildAssets = Compare-Object -ReferenceObject $GroupMSGraphAssets -DifferenceObject $BuildingAssets | Where-Object {$_ -like "*<=*"} | Select-Object -ExpandProperty InputObject
        foreach ($asset in $InvaildAssets) {
            $EntraID = Get-MgDevice -Filter "startswith(displayName, '$asset')"
            try {
                foreach ($rid in $($EntraID.Id)) {
                    Remove-MgGroupMemberByRef -GroupId $GroupID -DirectoryObjectId $rid -ErrorAction stop
                    Write-IntuneEventLog -EventID 302 -EventType "Information" -Message "Deleted $asset from $group."
                }
            }
            catch {           
                Write-IntuneEventLog -EventID 304 -EventType "Error" -Message "Failed to delete $asset from $group. $($_.Exception.Message)" 
            }           
        }
        # Clear temporary variables.
        $InvaildAssets = @()
        $MissingHelpdeskAssets = @()
    }
}
#endregion

########MAIN EXECUTION BLOCK############

# Check the size of the transcription log file. If it exceeds 50MB, delete it to start a new one.
$LogFileSizeMB = (Get-Item "$ENV:ProgramData/IntuneSync/Intunesync.log" -ErrorAction SilentlyContinue).Length/1MB
if($LogFileSizeMB -gt 50){
    Remove-Item "$ENV:ProgramData/IntuneSync/Intunesync.log"
    # Log that the log file was rotated.
    Write-IntuneEventLog -EventID 503 -EventType "Warning" -Message "Log file exceeded 50MB rolling over to new logfile."
}

# Start transcription to capture all console output to a log file.
Start-Transcript -Path "$ENV:ProgramData/IntuneSync/Intunesync.log" -Append
# Record the script start timestamp.
$StartDate = Get-date
Write-Output "Start sync timestamp: $StartDate"

# Ensure the custom Event Log for IntuneSync exists.
Get-IntuneEventLog
# Log the start of the Intune Sync process to the Event Log.
Write-IntuneEventLog -EventID 01 -EventType "Information" -Message "Intune Sync Started.`nTimestamp: $StartDate"

# Perform prerequisite checks and setup (e.g., module import, directory creation).
Start-D303DeviceSyncPrereqs
# Establish connection to Microsoft Graph API.
Start-D303ScriptConnections -TenantId $azureTenantId -ApplicationId $azureAplicationId -Thumbprint $azureCertThumbprint

# Remove primary users from Intune devices, converting them to shared devices.
Remove-AllDevicesPrimaryUser

# Synchronize devices for the predefined lab groups.
Start-DeviceGroupSync
# Synchronize unassigned devices to the "Intune-Unassigned" group.
Start-DeviceGroupSyncUnassigned -GroupName "Intune-Unassigned"
# Synchronize completed Autopilot devices to the "AutoPilot-Completed" group.
Start-DeviceGroupSyncCompleted -CompletedGroupName "AutoPilot-Completed"
# Synchronize devices for the "AutoPatch-Pilot" group based on the "ITS Peck Road" building.
Start-DeviceGroupSyncBuilding -GroupName "AutoPatch-Pilot" -Building "ITS Peck Road"

# Remove retired devices from Autopilot deployment profiles.
Remove-RetiredDevices
# Clean up duplicate device entries in Intune.
Invoke-IntuneCleanup
# Synchronize computer names between Active Directory and helpdesk/Intune.
Sync-ComputerNames

# Disconnect from Microsoft Graph (currently commented out in the original script).
#Disconnect-Graph

# Record the script end timestamp.
$EndDate = get-date
# Calculate the total execution time of the script.
$ExecutionTime = New-TimeSpan -Start $StartDate -End $EndDate
Write-Output "Stop sync timestamp: $EndDate"
Write-Output "Script exectuion time: $ExecutionTime"
# Log the completion of the Intune Sync process and its execution time to the Event Log.
Write-IntuneEventLog -EventID 00 -EventType "Information" -Message "Intune Sync Finished. `nTimestamp: $EndDate`nExecution time: $ExecutionTime"

# Stop transcription, saving the log.
Stop-Transcript