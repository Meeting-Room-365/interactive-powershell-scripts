#####################################################################
##  Meeting Room 365 Interactive Room List Manager                 ##
##  -------------------------------------------------------------  ##
##  (c) Copyright 2024 Meeting Room 365 llc. All Rights Reserved.  ##
##  Visit www.meetingroom365.com for more details.                 ##
#####################################################################

# Function to display the ASCII welcome message
function Show-WelcomeMessage {
    Clear-Host  # Clears the screen
    $asciiArt = @"
::::    ::::  :::::::::: :::::::::: ::::::::::: ::::::::::: ::::    :::  ::::::::   :::::::::   ::::::::   ::::::::  ::::    ::::    ::::::::   ::::::::  ::::::::::
+:+:+: :+:+:+ :+:        :+:            :+:         :+:     :+:+:   :+: :+:    :+:  :+:    :+: :+:    :+: :+:    :+: +:+:+: :+:+:+  :+:    :+: :+:    :+: :+:    :+:
+:+ +:+:+ +:+ +:+        +:+            +:+         +:+     :+:+:+  +:+ +:+         +:+    +:+ +:+    +:+ +:+    +:+ +:+ +:+:+ +:+         +:+ +:+        +:+
+#+  +:+  +#+ +#++:++#   +#++:++#       +#+         +#+     +#+ +:+ +#+ :#:         +#++:++#:  +#+    +:+ +#+    +:+ +#+  +:+  +#+      +#++:  +#++:++#+  +#++:++#+
+#+       +#+ +#+        +#+            +#+         +#+     +#+  +#+#+# +#+   +#+#  +#+    +#+ +#+    +#+ +#+    +#+ +#+       +#+         +#+ +#+    +#+        +#+
#+#       #+# #+#        #+#            #+#         #+#     #+#   #+#+# #+#    #+#  #+#    #+# #+#    #+# #+#    #+# #+#       #+#  #+#    #+# #+#    #+# #+#    #+#
###       ### ########## ##########     ###     ########### ###    ####  ########   ###    ###  ########   ########  ###       ###   ########   ########   ########
"@

    Write-Host $asciiArt
    Write-Host ""
    Write-Host "---------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Write-Host "Welcome to the Meeting Room 365 Room List Manager. Feel free to share this script as long as the copyright and welcome message is in-tact."
    Write-Host "(c) Copyright 2024 Meeting Room 365 llc. All Rights Reserved."
    Write-Host "Visit www.meetingroom365.com for more details."
    Write-Host "---------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Write-Host ""
    Write-Host "Tip: resize this window until you can see the logo above. This will help display large tables correctly." -ForegroundColor Yellow
    Write-Host ""
}

# Function to check and install a module if it's missing
function Ensure-ModuleInstalled {
    param (
        [string]$ModuleName
    )

    if (!(Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Installing module $ModuleName..."
        Install-Module -Name $ModuleName -Force -Confirm:$false
    } else {
        Write-Host "Module $ModuleName is already installed."
    }
}

# Function to get interactive input
function Get-UserInput {
    param (
        [string]$Prompt
    )
    Write-Host ""  # Add a blank line for spacing
    Write-Host $Prompt
    Read-Host
}

# Function to display options and get a choice from the user
function Show-Menu {
    Write-Host ""  # Add a blank line for spacing
    Write-Host "Please choose an option:"
    Write-Host "1. Show Room Lists"
    Write-Host "2. Create a New Room List"
    Write-Host "3. List Resource Mailboxes"
    Write-Host "4. Fix Subjects for Room Mailboxes"
    Write-Host "5. Reset Password for Resource Mailboxes"
    Write-Host "6. Create a New Service User"
    Write-Host "7. Exit"
    return (Get-UserInput -Prompt "Enter your choice (1, 2, 3, 4, 5, 6, or 7)")
}

# Function to generate a secure password with special characters
function Generate-SecurePassword {
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()"
    $securePassword = -join ((65..90) + (97..122) + (48..57) + (33..47) | Get-Random -Count 10 | ForEach-Object { [char]$_ })
    return $securePassword
}

# Function to create a new service user and assign management scope
function Create-NewServiceUser {
    # Prompt for the email address
    $email = Get-UserInput -Prompt "Enter the email address for the new service user:"

    # Ensure the email is not blank and contains an '@' symbol
    if ([string]::IsNullOrWhiteSpace($email) -or -not $email.Contains('@')) {
        Write-Host "Invalid email address. Please provide a valid email."
        return
    }

    # Extract the alias from the email address (the part before the @ symbol)
    $alias = $email.Split('@')[0]

    # Generate a secure password
    $newPassword = Generate-SecurePassword
    $securePassword = $newPassword | ConvertTo-SecureString -AsPlainText -Force

    # Create the new service user mailbox
    try {
        $mailbox = New-Mailbox -Alias $alias -MicrosoftOnlineServicesID $email -Password $securePassword -Name "Service User $alias"
        Write-Host "Service user mailbox created successfully."
    } catch {
        Write-Host "Failed to create service user mailbox. Error: $_"
        return
    }

    # Create the new management scope for RoomMailboxes, Workspaces, and Equipment mailboxes
    try {
        New-ManagementScope -Name "RoomWorkspacesAndEquipment" -RecipientRestrictionFilter { RecipientTypeDetails -eq "RoomMailbox" -or RecipientTypeDetails -eq "Workspace" -or RecipientTypeDetails -eq "EquipmentMailbox" }
        Write-Host "Management scope for room mailboxes, workspaces, and equipment created successfully."
    } catch {
        Write-Host "Failed to create management scope. Error: $_"
    }

    # Assign the ApplicationImpersonation role and the management role to the service user
    try {
        New-ManagementRoleAssignment -Name "RoomMailboxManager" -Role "ApplicationImpersonation" -User $email -CustomRecipientWriteScope "RoomWorkspacesAndEquipment"
        Write-Host "ApplicationImpersonation role and management role assignment for the service user created successfully."
    } catch {
        Write-Host "Failed to assign management role. Error: $_"
    }

    # Display the email and password for the new service user
    Write-Host "Service user created successfully."
    Write-Host "Email: $email"
    Write-Host "Password: $newPassword"
}

# Function to reset password for resource mailboxes (rooms and workspaces)
function Reset-PasswordForResourceMailboxes {
    Write-Host "Fetching all resource mailboxes (rooms and workspaces)..."

    # Get both RoomMailboxes and workspaces
    $roomsAndWorkspaces = Get-Mailbox -RecipientTypeDetails RoomMailbox
    $workspaces = Get-Mailbox -Filter { ResourceType -eq 'Workspace' }

    $allResources = $roomsAndWorkspaces + $workspaces

    if ($allResources.Count -eq 0) {
        Write-Host "No resource mailboxes found."
        return
    }

    # Prepare an array for the result
    $resourceListResults = @()
    $i = 1

    # Collect resource mailbox information for table output
    foreach ($resource in $allResources) {
        $resourceListResults += [PSCustomObject]@{
            Number = $i
            Name   = $resource.Name
            Email  = $resource.PrimarySmtpAddress
        }
        $i++
    }

    # Display resource mailboxes in a table format with numbers
    $resourceListResults | Format-Table Number, Name, Email -AutoSize

    # Prompt user to select a resource mailbox by number
    $resourceChoice = Get-UserInput -Prompt "Enter the number of the resource mailbox to reset the password, or leave blank to return to the main menu"

    if ([string]::IsNullOrWhiteSpace($resourceChoice) -eq $false -and ($resourceChoice -as [int]) -le $resourceListResults.Count) {
        $selectedResource = $resourceListResults | Where-Object { $_.Number -eq $resourceChoice }

        # Confirm if the user wants to reset the password
        $confirmReset = Get-UserInput -Prompt "Are you sure you want to reset the password for $($selectedResource.Name)? (Y/N)"
        if ($confirmReset -eq "Y") {
            # Generate a secure password with special characters
            $newPassword = Generate-SecurePassword
            $securePassword = $newPassword | ConvertTo-SecureString -AsPlainText -Force

            try {
                # Enable room mailbox account and reset the password using Set-Mailbox
                Set-Mailbox -Identity $selectedResource.Email -EnableRoomMailboxAccount $true -RoomMailboxPassword $securePassword
                Write-Host "Password reset successfully for $($selectedResource.Name)."
                Write-Host "Email: $($selectedResource.Email)"
                Write-Host "New password: $newPassword"
            } catch {
                Write-Host "Failed to reset the password. Error: $_"
            }
        } else {
            Write-Host "Password reset cancelled."
        }
    } else {
        Write-Host "Returning to the main menu..."
    }
}

# Function to fix subjects for room mailboxes and list again
function Fix-SubjectsForRoomMailboxes {
    do {
        Write-Host "Fetching all room mailboxes..."

        # Get all room mailboxes
        $roomMailboxes = Get-Mailbox -RecipientTypeDetails RoomMailbox

        if ($roomMailboxes.Count -eq 0) {
            Write-Host "No room mailboxes found."
            return
        }

        # Prepare an array for the result
        $roomMailboxResults = @()

        # Collect room mailbox information for table output, including the status of AddOrganizerToSubject
        $i = 1
        foreach ($room in $roomMailboxes) {
            $calendarProcessing = Get-CalendarProcessing -Identity $room.PrimarySmtpAddress
            $addOrganizerToSubject = $calendarProcessing.AddOrganizerToSubject

            $roomMailboxResults += [PSCustomObject]@{
                Number              = $i
                Name                = $room.Name
                Email               = $room.PrimarySmtpAddress
                AddOrganizerToSubject = $addOrganizerToSubject
            }
            $i++
        }

        # Display room mailboxes in a table format with AddOrganizerToSubject status
        $roomMailboxResults | Format-Table Number, Name, Email, AddOrganizerToSubject -AutoSize

        # Prompt user to select a room mailbox by number
        $roomChoice = Get-UserInput -Prompt "Enter the number of the room mailbox to fix the subject handling, or leave blank to return to the main menu"

        if ([string]::IsNullOrWhiteSpace($roomChoice) -eq $false -and ($roomChoice -as [int]) -le $roomMailboxResults.Count) {
            $selectedRoom = $roomMailboxResults | Where-Object { $_.Number -eq $roomChoice }

            try {
                # Run the Set-CalendarProcessing command on the selected room mailbox
                Set-CalendarProcessing -Identity $selectedRoom.Email -AddOrganizerToSubject $false -DeleteSubject $false -DeleteComments $false -RemovePrivateProperty $false
                Write-Host "Successfully fixed the subject handling for room mailbox: $($selectedRoom.Name)"
            } catch {
                Write-Host "Failed to fix subject handling for room mailbox. Error: $_"
            }

            # Ask if the user wants to continue fixing more room mailboxes
            $continueFixing = Get-UserInput -Prompt "Do you want to fix subject handling for another room mailbox? (Y/N)"
        } else {
            Write-Host "Returning to the main menu..."
            break
        }
    } while ($continueFixing -eq "Y")
}

# Function to show room lists, view members, and perform add/remove actions
function Show-RoomLists {
    Write-Host "Fetching all room lists..."

    # Get all room lists
    $roomLists = Get-DistributionGroup -RecipientTypeDetails RoomList

    if ($roomLists.Count -eq 0) {
        Write-Host "No room lists found."
        return
    }

    # Prepare an array for the result
    $roomListResults = @()

    # Collect room list information for table output
    $i = 1
    foreach ($list in $roomLists) {
        $roomListResults += [PSCustomObject]@{
            Number = $i
            Name   = $list.Name
            Email  = $list.PrimarySmtpAddress
        }
        $i++
    }

    # Display room lists in a table format with numbers
    $roomListResults | Format-Table Number, Name, Email -AutoSize

    # Prompt user to select a room list by number
    $listChoice = Get-UserInput -Prompt "Enter the number of the room list to view resources, or leave blank to return to the main menu"

    if ([string]::IsNullOrWhiteSpace($listChoice) -eq $false -and ($listChoice -as [int]) -le $roomListResults.Count) {
        $selectedRoomList = $roomListResults | Where-Object { $_.Number -eq $listChoice }
        if ($selectedRoomList) {
            Display-RoomListResources $selectedRoomList
        } else {
            Write-Host "Invalid selection. Returning to the main menu."
        }
    } else {
        Write-Host "Returning to the main menu..."
    }
}

# Function to display resources in a room list and perform actions
function Display-RoomListResources {
    param (
        [PSCustomObject]$selectedRoomList
    )

    Write-Host "Fetching resources in the room list: $($selectedRoomList.Name)..."
    $members = Get-DistributionGroupMember -Identity $selectedRoomList.Email

    if ($members.Count -eq 0) {
        Write-Host "No resources found in the selected room list."
    } else {
        # Display resources with numbers and additional details from Get-Place
        $memberListResults = @()
        $j = 1
        foreach ($member in $members) {
            $placeDetails = Get-Place -Identity $member.PrimarySmtpAddress
            $building = $placeDetails.Building
            $capacity = $placeDetails.Capacity
            $location = $placeDetails.City
            $floor = $placeDetails.Floor
            $countryOrRegion = $placeDetails.CountryOrRegion
            $postalCode = $placeDetails.PostalCode
            $state = $placeDetails.State
            $label = $placeDetails.Label
            $type = $placeDetails.Type

            $memberListResults += [PSCustomObject]@{
                Number          = $j
                Name            = $member.Name
                Email           = $member.PrimarySmtpAddress
                Type            = $type
                Capacity        = $capacity
                Building        = $building
                Floor           = $floor
                City            = $location
                State           = $state
                PostalCode      = $postalCode
                CountryOrRegion = $countryOrRegion
                Label           = $label
            }
            $j++
        }

        Write-Host "Resources in $($selectedRoomList.Name):"
        $memberListResults | Format-Table Number, Name, Email, Type, Capacity, Building, Floor, City, State, PostalCode, CountryOrRegion, Label -AutoSize
    }

    # Prompt to add, remove, or modify a resource
    $action = Get-UserInput -Prompt "Do you want to add, remove, or modify a resource? (Enter 'add', 'remove', or 'modify', or leave blank to return to the main menu)"

    switch ($action.ToLower()) {
        'add' {
            Add-Resource $selectedRoomList.Email
            Display-RoomListResources $selectedRoomList  # Refresh the list after adding
            return
        }

        'remove' {
            if ($members.Count -gt 0) {
                $resourceToRemove = Get-UserInput -Prompt "Enter the number of the resource to remove from the list"
                if ([int]$resourceToRemove -le $members.Count -and [int]$resourceToRemove -ge 1) {
                    $selectedResource = $memberListResults | Where-Object { $_.Number -eq $resourceToRemove }
                    Remove-ResourceFromRoomList -roomListAlias $selectedRoomList.Email -resourceEmail $selectedResource.Email
                    Display-RoomListResources $selectedRoomList  # Refresh the list after removing
                    return
                } else {
                    Write-Host "Invalid selection. Returning to room list..."
                    return
                }
            } else {
                Write-Host "No resources to remove. Returning to room list..."
                return
            }
        }

        'modify' {
            if ($members.Count -gt 0) {
                $resourceToUpdate = Get-UserInput -Prompt "Enter the number of the resource to update properties"
                if ([int]$resourceToUpdate -le $members.Count -and [int]$resourceToUpdate -ge 1) {
                    $selectedResource = $memberListResults | Where-Object { $_.Number -eq $resourceToUpdate }
                    Update-ResourceProperties -resourceEmail $selectedResource.Email
                    Write-Host "Refreshing the list of resources in the room list after updating..."
                    Display-RoomListResources $selectedRoomList  # Refresh the list after updating
                    return
                } else {
                    Write-Host "Invalid selection. Returning to room list..."
                    return
                }
            } else {
                Write-Host "No resources to modify. Returning to room list..."
                return
            }
        }

        default {
            Write-Host "Returning to the main menu..."
            return
        }
    }
}

# Function to add a resource and automatically generate an email address
function Add-Resource {
    param (
        [string]$roomListAlias
    )

    # Ask if it's a room or workspace first
    $typeChoice = (Get-UserInput -Prompt "Do you want to add a room or workspace? (Enter 'room' or 'workspace')").ToLower()

    # Ensure valid choice
    if ($typeChoice -ne "room" -and $typeChoice -ne "workspace") {
        Write-Host "Invalid input. Please enter 'room' or 'workspace'. Returning to the main menu..."
        return
    }

    # Ask for the name of the resource
    $resourceName = Get-UserInput -Prompt "Enter the name of the $typeChoice"

    try {
        # Create room or workspace mailbox without PrimarySmtpAddress
        if ($typeChoice -eq "room") {
            New-Mailbox -Name $resourceName -Room
        } elseif ($typeChoice -eq "workspace") {
            New-Mailbox -Name $resourceName -Room
            Set-Mailbox -Identity $resourceName -Type Workspace
        }

        Add-DistributionGroupMember -Identity $roomListAlias -Member $resourceName
        Write-Host "$typeChoice '$resourceName' added to the room list."
    } catch {
        Write-Host "Failed to add $typeChoice. Please check the details and try again."
    }
}

# Function to remove a resource from a room list
function Remove-ResourceFromRoomList {
    param (
        [string]$roomListAlias,
        [string]$resourceEmail
    )

    try {
        Remove-DistributionGroupMember -Identity $roomListAlias -Member $resourceEmail
        Write-Host "Resource '$resourceEmail' removed from the room list '$roomListAlias'."
    } catch {
        Write-Host "Failed to remove the resource from the room list. Please check the details and try again."
    }
}

# Function to add a resource to a room list
function Add-ResourceToRoomList {
    param (
        [PSCustomObject]$selectedResource
    )

    # Display room lists as a numbered list
    Write-Host "Fetching all room lists..."
    $roomLists = Get-DistributionGroup -RecipientTypeDetails RoomList

    if ($roomLists.Count -eq 0) {
        Write-Host "No room lists found."
        return
    }

    $roomListResults = @()
    $j = 1

    foreach ($list in $roomLists) {
        $roomListResults += [PSCustomObject]@{
            Number = $j
            Name   = $list.Name
            Email  = $list.PrimarySmtpAddress
        }
        $j++
    }

    # Display room lists
    $roomListResults | Format-Table Number, Name, Email -AutoSize

    # Prompt user to select a room list by number to add the resource to
    $roomListChoice = Get-UserInput -Prompt "Enter the number of the room list to add the resource to, or leave blank to return to the main menu"

    if ([string]::IsNullOrWhiteSpace($roomListChoice) -eq $false -and ($roomListChoice -as [int]) -le $roomListResults.Count) {
        $selectedRoomList = $roomListResults | Where-Object { $_.Number -eq $roomListChoice }
        try {
            Add-DistributionGroupMember -Identity $selectedRoomList.Email -Member $selectedResource.Email
            Write-Host "Resource '$($selectedResource.Name)' added to the room list '$($selectedRoomList.Name)'."
        } catch {
            Write-Host "Failed to add the resource to the room list. Please check the details and try again."
        }
    } else {
        Write-Host "Returning to the main menu..."
    }
}

# Function to create a new room list with auto-generated alias
function Create-NewRoomList {
    $roomListName = Get-UserInput -Prompt "Enter the name of the room list you want to create:"
    $roomListAlias = Get-UserInput -Prompt "Enter the alias for the room list (leave blank to auto-generate):"

    # If no alias is provided, generate one based on the room list name
    if ([string]::IsNullOrWhiteSpace($roomListAlias)) {
        $roomListAlias = $roomListName.ToLower().Replace(" ", "-")
        Write-Host "No alias provided. Generated alias: $roomListAlias"
    }

    $organizationalUnit = Get-UserInput -Prompt "Enter the organizational unit (OU), or leave blank to use the default OU:"

    if ([string]::IsNullOrWhiteSpace($organizationalUnit)) {
        New-DistributionGroup -Name $roomListName -Alias $roomListAlias -RoomList
    } else {
        New-DistributionGroup -Name $roomListName -Alias $roomListAlias -RoomList -OrganizationalUnit $organizationalUnit
    }

    Write-Host "Room list '$roomListName' created successfully."
}

# Function to update properties for a resource using Set-Place and Set-Mailbox
function Update-ResourceProperties {
    param (
        [string]$resourceEmail
    )

    Write-Host "Updating properties for resource: $resourceEmail"

    # Get user input for each property; leave blank to skip
    $newName = Get-UserInput -Prompt "Enter new name for the resource (leave blank to skip):"
    $newCapacity = Get-UserInput -Prompt "Enter new capacity for the resource (leave blank to skip):"
    $newBuilding = Get-UserInput -Prompt "Enter new building for the resource (leave blank to skip):"
    $newFloor = Get-UserInput -Prompt "Enter new floor for the resource (leave blank to skip):"
    $newCity = Get-UserInput -Prompt "Enter new city for the resource (leave blank to skip):"
    $newState = Get-UserInput -Prompt "Enter new state for the resource (leave blank to skip):"
    $newPostalCode = Get-UserInput -Prompt "Enter new postal code for the resource (leave blank to skip):"
    $newCountryOrRegion = Get-UserInput -Prompt "Enter new country/region for the resource (leave blank to skip):"
    $newLabel = Get-UserInput -Prompt "Enter new label for the resource (leave blank to skip):"

    # Prepare parameters for Set-Place
    $params = @{}

    if (-not [string]::IsNullOrEmpty($newCapacity)) { $params['Capacity'] = $newCapacity }
    if (-not [string]::IsNullOrEmpty($newBuilding)) { $params['Building'] = $newBuilding }
    if (-not [string]::IsNullOrEmpty($newFloor)) { $params['Floor'] = $newFloor }
    if (-not [string]::IsNullOrEmpty($newCity)) { $params['City'] = $newCity }
    if (-not [string]::IsNullOrEmpty($newState)) { $params['State'] = $newState }
    if (-not [string]::IsNullOrEmpty($newPostalCode)) { $params['PostalCode'] = $newPostalCode }
    if (-not [string]::IsNullOrEmpty($newCountryOrRegion)) { $params['CountryOrRegion'] = $newCountryOrRegion }
    if (-not [string]::IsNullOrEmpty($newLabel)) { $params['Label'] = $newLabel }

    # If a new name is provided, update it with Set-Mailbox
    if (-not [string]::IsNullOrEmpty($newName)) {
        try {
            Set-Mailbox -Identity $resourceEmail -Name $newName
            Write-Host "Successfully updated the resource name to '$newName'."
        } catch {
            Write-Host "Failed to update resource name. Error: $_"
        }
    }

    # If any parameters were set, apply them using Set-Place
    if ($params.Count -gt 0) {
        try {
            Set-Place -Identity $resourceEmail @params
            Write-Host "Successfully updated the resource properties."
        } catch {
            Write-Host "Failed to update resource properties. Error: $_"
        }
    } else {
        Write-Host "No properties were updated."
    }
}

# Function to list all resource mailboxes (rooms & workspaces), show room list membership, and allow resource selection by number
function List-ResourceMailboxes {
    Write-Host "Fetching all resource mailboxes (rooms and workspaces)..."

    # Get both RoomMailboxes and mailboxes that are workspaces
    $roomsAndWorkspaces = Get-Mailbox -RecipientTypeDetails RoomMailbox
    $workspaces = Get-Mailbox -Filter { ResourceType -eq 'Workspace' }

    $allResources = $roomsAndWorkspaces + $workspaces

    if ($allResources.Count -eq 0) {
        Write-Host "No resource mailboxes found."
        return
    }

    # Get all distribution groups and store their members in a hash table for faster lookup
    Write-Host "Fetching distribution groups and caching memberships..."
    $distributionGroups = Get-DistributionGroup
    $groupMembers = @{}

    foreach ($group in $distributionGroups) {
        $groupMembers[$group.Alias] = Get-DistributionGroupMember -Identity $group.Alias
    }

    # Prepare an array for the result
    $resourceListResults = @()
    $i = 1

    # Display resources with their respective room list memberships
    foreach ($resource in $allResources) {
        $membership = @()

        # Check if the resource is a member of each cached distribution group
        foreach ($group in $groupMembers.Keys) {
            $members = $groupMembers[$group]
            if ($members | Where-Object { $_.PrimarySmtpAddress -eq $resource.PrimarySmtpAddress }) {
                $membership += $distributionGroups | Where-Object { $_.Alias -eq $group } | Select-Object -ExpandProperty Name
            }
        }

        if ($membership.Count -eq 0) {
            $membership = "None"
        } else {
            $membership = $membership -join ", "
        }

        $placeDetails = Get-Place -Identity $resource.PrimarySmtpAddress
        $building = $placeDetails.Building
        $capacity = $placeDetails.Capacity
        $location = $placeDetails.City
        $floor = $placeDetails.Floor
        $countryOrRegion = $placeDetails.CountryOrRegion
        $postalCode = $placeDetails.PostalCode
        $state = $placeDetails.State
        $label = $placeDetails.Label
        $type = $placeDetails.Type

        $resourceListResults += [PSCustomObject]@{
            Number          = $i
            Name            = $resource.Name
            Email           = $resource.PrimarySmtpAddress
            MemberOfGroups  = $membership
            Type            = $type
            Capacity        = $capacity
            Building        = $building
            Floor           = $floor
            City            = $location
            State           = $state
            PostalCode      = $postalCode
            CountryOrRegion = $countryOrRegion
            Label           = $label
        }
        $i++
    }

    # Display resources as a numbered list with memberships and additional details
    $resourceListResults | Format-Table Number, Name, Email, Type, Capacity, Building, Floor, City, State, PostalCode, CountryOrRegion, Label, MemberOfGroups -AutoSize

    # Prompt user to select a resource by number
    $resourceChoice = Get-UserInput -Prompt "Enter the number of the resource to add to or remove from a room list, or leave blank to return to the main menu"

    if ([string]::IsNullOrWhiteSpace($resourceChoice) -eq $false -and ($resourceChoice -as [int]) -le $resourceListResults.Count) {
        $selectedResource = $resourceListResults | Where-Object { $_.Number -eq $resourceChoice }

        if ($selectedResource.MemberOfGroups -eq "None") {
            Write-Host "The selected resource is not a member of any room list."
            Add-ResourceToRoomList $selectedResource
        } else {
            Write-Host "The selected resource is a member of the following room list(s): $($selectedResource.MemberOfGroups)"
            $actionChoice = Get-UserInput -Prompt "Do you want to add this resource to another room list or remove it from a list? (Enter 'add' or 'remove')"

            if ($actionChoice -eq "add") {
                Add-ResourceToRoomList $selectedResource
            } elseif ($actionChoice -eq "remove") {
                Remove-ResourceFromRoomList $selectedResource
            } else {
                Write-Host "Invalid input. Returning to the main menu..."
            }
        }
    } else {
        Write-Host "Returning to the main menu..."
    }
}

# Main script logic
try {
    Show-WelcomeMessage  # Display the welcome message and clear the screen
    Ensure-ModuleInstalled -ModuleName "ExchangeOnlineManagement" # Ensure Exchange Online Management module is installed
    Write-Host "Connecting to Exchange Online..."
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -ShowProgress $true -ShowBanner:$false
}
catch {
    Write-Host "Failed to connect to Exchange Online."
    exit
}

# Loop to show menu and process user choices until the user decides to exit
do {
    $choice = Show-Menu
    switch ($choice) {
        1 { Show-RoomLists }
        2 { Create-NewRoomList }
        3 { List-ResourceMailboxes }
        4 { Fix-SubjectsForRoomMailboxes }
        5 { Reset-PasswordForResourceMailboxes }
        6 { Create-NewServiceUser }
        7 {
            Write-Host "Exiting script..."
            Write-Host "Please note, it may take up to 24 hours for room lists to sync and appear in Outlook."
            break
        }
        default { Write-Host "Invalid choice. Returning to the main menu..." }
    }
} while ($choice -ne 7)
