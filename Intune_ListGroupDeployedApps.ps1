# PowerShell Script to Find All Apps Assigned to a Specific Group in Intune
# Author: Gemini
# Description: This script connects to the Microsoft Graph API to retrieve a list of all
# applications assigned to a specific Microsoft Entra ID group. It prompts the user
# for the group name and outputs the app name and assignment intent.
#
# Prerequisites:
# - Microsoft Graph PowerShell SDK module must be installed.
#   Run: Install-Module Microsoft.Graph.Groups -Scope AllUsers -Force
#   Run: Install-Module Microsoft.Graph.DeviceManagement.Apps -Scope AllUsers -Force
# - An Intune administrator account with the necessary permissions (e.g., DeviceManagementApps.Read.All, Group.Read.All)
#   is required to run the script.

#region Helper Functions
function Connect-IntuneGraph {
    # This function handles the connection to Microsoft Graph with the required scopes.
    # It will prompt for an interactive login.
    try {
        # Define the scopes needed for this script.
        $scopes = "Group.Read.All", "DeviceManagementApps.Read.All"
        
        # Connect to Microsoft Graph. The command will open a browser window for authentication.
        Write-Host "Connecting to Microsoft Graph. A browser window will open for authentication..." -ForegroundColor Green
        Connect-MgGraph -Scopes $scopes -ErrorAction Stop

        Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to connect to Microsoft Graph. Please check your permissions and try again." -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}
#endregion

# Main Script Logic
try {
    # Step 1: Connect to Intune/Microsoft Graph
    if (-not (Connect-IntuneGraph)) {
        throw "Connection failed. Exiting script."
    }

    # Step 2: Prompt the user for the group name
    $groupName = Read-Host "Enter the exact name of the group you want to check"
    if ([string]::IsNullOrWhiteSpace($groupName)) {
        Write-Host "Group name cannot be empty. Exiting script." -ForegroundColor Red
        return
    }

    Write-Host "Searching for group '$groupName'..." -ForegroundColor Yellow

    # Step 3: Find the group object based on the provided name
    # The -ConsistencyLevel parameter is required for the -Search parameter to work.
    $group = Get-MgGroup -ConsistencyLevel eventual -Search "DisplayName:$($groupName)" -ErrorAction SilentlyContinue | Select-Object -First 1
    
    if (-not $group) {
        Write-Host "Group not found. Please check the group name and try again. Exiting script." -ForegroundColor Red
        return
    }

    Write-Host "Found group: $($group.DisplayName) (ID: $($group.Id))" -ForegroundColor Green
    
    # Step 4: Get all mobile apps from Intune
    Write-Host "Fetching all mobile applications from Intune. This may take a few moments..." -ForegroundColor Yellow
    $allApps = Get-MgDeviceAppManagementMobileApp -All -ExpandProperty 'assignments' -ErrorAction Stop
    
    Write-Host "Found $($allApps.Count) applications." -ForegroundColor Green
    
    # Step 5: Filter apps by group assignment
    $assignedApps = @() # Create an empty array to store the results.

    foreach ($app in $allApps) {
        # Check if the app has any assignments
        if ($app.Assignments) {
            # Check each assignment for a match with the target group's ID.
            $match = $app.Assignments | Where-Object { $_.Id -like "$($group.Id)*" }
            
            if ($match) {
                # If a match is found, add the app details to our results array.
                $assignedApps += [PSCustomObject]@{
                    AppName = $app.DisplayName
                    AssignmentIntent = $match.Intent # Intent can be 'Required', 'Available', 'Uninstall'
                    GroupDisplayName = $group.DisplayName
                }
            }
        }
    }

    # Step 6: Output the results
    Write-Host " "
    Write-Host "---------------------------------------------------------" -ForegroundColor DarkYellow
    Write-Host "  Apps Assigned to Group: $($group.DisplayName)" -ForegroundColor Cyan
    Write-Host "---------------------------------------------------------" -ForegroundColor DarkYellow
    Write-Host " "

    if ($assignedApps.Count -gt 0) {
        $assignedApps | Format-Table -AutoSize
    } else {
        Write-Host "No applications found deployed to this group." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "An error occurred during script execution:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
finally {
    # Disconnect from the Graph session to clean up.
    #Disconnect-MgGraph
    Write-Host "Script execution complete. Disconnected from Microsoft Graph." -ForegroundColor Green
}