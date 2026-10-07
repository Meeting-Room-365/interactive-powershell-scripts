#####################################################################
##  Meeting Room 365 Convert to Resource Mailbox                   ##
##  -------------------------------------------------------------  ##
##  (c) Copyright 2024 Meeting Room 365 llc. All Rights Reserved.  ##
##  Visit www.meetingroom365.com for more details.                 ##
#####################################################################

<#
.SYNOPSIS
  Converts an existing Microsoft 365 mailbox to a Room or Workspace
  resource mailbox and configures its calendar processing.

.DESCRIPTION
  Prompts for any value not passed as a parameter.

.EXAMPLE
  ./ConvertToResourceMailbox.ps1

.EXAMPLE
  ./ConvertToResourceMailbox.ps1 -Identity "boardroom@contoso.com" -Type Room -Capacity 12 -RoomList "rooms-hq@contoso.com"
#>
param(
    [string] $Identity,
    [ValidateSet('Room', 'Workspace')] [string] $Type,
    [int]    $Capacity,
    [string] $RoomList,     # existing room list (distribution group of type RoomList)
    [switch] $BlockSignIn   # disable interactive sign-in (Microsoft Graph)
)

$ErrorActionPreference = "Stop"

# -------------------------
# Helpers
# -------------------------
function Ask-YesNo([string]$prompt, [string]$default = "n") {
    while ($true) {
        $ans = Read-Host "$prompt [$default]"
        if (-not $ans) { $ans = $default }
        if ($ans -in @("y", "n")) { return ($ans -eq "y") }
    }
}

function Get-CalendarItemCount([string]$id) {
    (Get-MailboxFolderStatistics -Identity $id -FolderScope Calendar |
        Where-Object FolderType -eq "Calendar").ItemsInFolder
}

function Ensure-Module([string]$name) {
    if (-not (Get-Module -ListAvailable -Name $name)) {
        Write-Host "Installing module: $name" -ForegroundColor Yellow
        Install-Module $name -Scope CurrentUser -Force
    }
}

Clear-Host
Write-Host "Convert Mailbox to Resource Mailbox" -ForegroundColor Cyan
Write-Host "-----------------------------------"
Write-Host ""

# -------------------------
# Connect
# -------------------------
Ensure-Module "ExchangeOnlineManagement"
Import-Module ExchangeOnlineManagement

if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
}

# -------------------------
# Inputs
# -------------------------
while (-not $Identity) { $Identity = Read-Host "Mailbox email address to convert" }

$mbx = Get-Mailbox -Identity $Identity -ErrorAction SilentlyContinue
if (-not $mbx) {
    Write-Host "Mailbox not found: $Identity" -ForegroundColor Red
    exit 1
}

Write-Host "Found:" -NoNewline
Write-Host " $($mbx.DisplayName) <$($mbx.PrimarySmtpAddress)> ($($mbx.RecipientTypeDetails))" -ForegroundColor Green
Write-Host ""

while (-not $Type) {
    $choice = Read-Host "Convert to (1) Room or (2) Workspace? [1]"
    switch ($choice) {
        { $_ -in @("", "1") } { $Type = "Room" }
        "2"                   { $Type = "Workspace" }
    }
}

if (-not $PSBoundParameters.ContainsKey('Capacity')) {
    $ans = Read-Host "Capacity (blank to skip)"
    if ($ans) { $Capacity = [int]$ans }
}

if (-not $PSBoundParameters.ContainsKey('RoomList')) {
    $RoomList = Read-Host "Room list to add it to (blank to skip)"
}

if (-not $PSBoundParameters.ContainsKey('BlockSignIn')) {
    $BlockSignIn = Ask-YesNo "Block interactive sign-in for this account? (y/n)"
}

Write-Host ""
if (-not (Ask-YesNo "Convert $($mbx.PrimarySmtpAddress) to a $Type mailbox? (y/n)" "y")) {
    Write-Host "Cancelled."
    exit 0
}
Write-Host ""

# -------------------------
# 1) Convert mailbox type
#    Existing calendar items are kept. Calendar processing below only
#    applies to requests (new or updated) received after this point.
# -------------------------
$itemsBefore = Get-CalendarItemCount $Identity
Write-Host "Calendar items before: $itemsBefore"

$targetType = "$($Type)Mailbox"
if ($mbx.RecipientTypeDetails -ne $targetType) {
    Set-Mailbox -Identity $Identity -Type $Type
    Write-Host "Converted from $($mbx.RecipientTypeDetails) to $targetType"
} else {
    Write-Host "Already a $targetType"
}

if ($Capacity) {
    Set-Mailbox -Identity $Identity -ResourceCapacity $Capacity
    Write-Host "Capacity set to $Capacity"
}

# -------------------------
# 2) Calendar processing: auto-accept, keep real subjects/organizer visible
# -------------------------
Set-CalendarProcessing -Identity $Identity `
    -AutomateProcessing AutoAccept `
    -AllowConflicts $false `
    -BookingWindowInDays 180 `
    -MaximumDurationInMinutes 1440 `
    -DeleteSubject $false `
    -AddOrganizerToSubject $false `
    -DeleteComments $false `
    -RemovePrivateProperty $false `
    -ProcessExternalMeetingMessages $false
Write-Host "Calendar processing configured"

# -------------------------
# 3) Room list (shows in Outlook Room Finder)
# -------------------------
if ($RoomList) {
    try {
        Add-DistributionGroupMember -Identity $RoomList -Member $Identity
        Write-Host "Added to room list $RoomList"
    } catch {
        Write-Host "Warning: could not add to room list ($($_.Exception.Message))" -ForegroundColor Yellow
    }
}

# -------------------------
# 4) Block interactive sign-in
# -------------------------
if ($BlockSignIn) {
    try {
        Ensure-Module "Microsoft.Graph.Users"
        Import-Module Microsoft.Graph.Users
        Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome
        Update-MgUser -UserId $mbx.ExternalDirectoryObjectId -AccountEnabled:$false
        Write-Host "Interactive sign-in blocked"
    } catch {
        Write-Host "Warning: could not block sign-in ($($_.Exception.Message))" -ForegroundColor Yellow
    }
}

# -------------------------
# Verify
# -------------------------
$itemsAfter = Get-CalendarItemCount $Identity
if ($itemsAfter -lt $itemsBefore) {
    Write-Host "Warning: calendar items dropped from $itemsBefore to $itemsAfter" -ForegroundColor Yellow
} else {
    Write-Host "Calendar items after: $itemsAfter (existing meetings preserved)"
}

Write-Host ""
Write-Host "✔ Conversion complete" -ForegroundColor Green
Get-Mailbox -Identity $Identity | Format-List DisplayName, PrimarySmtpAddress, RecipientTypeDetails, ResourceCapacity
Get-CalendarProcessing -Identity $Identity | Format-List AutomateProcessing, DeleteSubject, AddOrganizerToSubject, RemovePrivateProperty
