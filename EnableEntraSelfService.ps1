#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Entra ID Self-Service Profile Setup Wizard

.DESCRIPTION
  Simple interactive wizard to configure tenant-level settings
  that affect user self-service profile editing.

  NOTE:
  Some attributes (displayName, jobTitle, department, etc.)
  are always admin-managed in Entra ID and cannot be enabled
  for self-service editing. This script does not prompt for them.
#>

$ErrorActionPreference = "Stop"

# -------------------------
# Helpers
# -------------------------
function Ask-YesNo([string]$prompt, [string]$default = "y") {
  $suffix = if ($default) { " [$default]" } else { "" }
  while ($true) {
    $ans = Read-Host "$prompt$suffix"
    if (-not $ans -and $default) { $ans = $default }
    if ($ans -in @("y","n")) { return ($ans -eq "y") }
  }
}

function Ensure-Module([string]$name) {
  if (-not (Get-Module -ListAvailable -Name $name)) {
    Write-Host "Installing module: $name" -ForegroundColor Yellow
    Install-Module $name -Scope CurrentUser -Force
  }
}

# -------------------------
# Init
# -------------------------
$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProfilePath = Join-Path $ScriptDir "entra-selfservice-profile.json"

Clear-Host
Write-Host "Entra ID Self-Service Profile Setup" -ForegroundColor Cyan
Write-Host "----------------------------------"
Write-Host ""

# -------------------------
# Graph connection
# -------------------------
Ensure-Module "Microsoft.Graph.Authentication"
Ensure-Module "Microsoft.Graph.Identity.SignIns"

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Identity.SignIns

$scopes = @(
  "Directory.Read.All",
  "PeopleSettings.ReadWrite.All",
  "Policy.ReadWrite.Authorization"
)

Write-Host "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes $scopes | Out-Null

$org = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization"
$tenantName = $org.value[0].displayName
$tenantId   = $org.value[0].id

Write-Host "Tenant:" -NoNewline
Write-Host " $tenantName" -ForegroundColor Green
Write-Host ""

# -------------------------
# Wizard questions (ONLY real tenant knobs)
# -------------------------
$allowSelfEdit = Ask-YesNo "Allow users to edit their own profile information?"
$allowPhoto    = Ask-YesNo "Allow users to change their profile photo?"
$allowContact  = Ask-YesNo "Allow users to edit contact info (phone, language)?"
$allowBio      = Ask-YesNo "Allow users to edit bio / about-me fields?"

Write-Host ""

# -------------------------
# Save intent (for UI + audit)
# -------------------------
$profile = @{
  tenant = @{
    displayName = $tenantName
    id = $tenantId
  }
  createdAt = (Get-Date).ToString("s")

  selfService = @{
    profileEditingEnabled = $allowSelfEdit
    contactInfoEditable   = $allowContact
    bioEditable           = $allowBio
    photoEditable         = $allowPhoto
  }

  notes = @{
    adminOnlyAttributes = @(
      "displayName",
      "givenName",
      "surname",
      "jobTitle",
      "department",
      "companyName",
      "userPrincipalName"
    )
  }
}

$profile | ConvertTo-Json -Depth 6 | Out-File $ProfilePath -Encoding UTF8
Write-Host "Saved configuration to $ProfilePath"
Write-Host ""

# -------------------------
# Apply tenant settings
# -------------------------

# 1) Authorization policy: allow self-profile editing
if ($allowSelfEdit) {
  Write-Host "Enabling self-service profile editing..."
  Update-MgPolicyAuthorizationPolicy `
    -BodyParameter @{ allowedToEditSelfProfile = $true } | Out-Null
} else {
  Write-Host "Disabling self-service profile editing..."
  Update-MgPolicyAuthorizationPolicy `
    -BodyParameter @{ allowedToEditSelfProfile = $false } | Out-Null
}

# 2) Profile photo policy (Graph beta)
#    This is the ONLY profile attribute with a real tenant switch
$ROLE_GLOBAL_ADMIN = "62e90394-69f5-4237-9190-012177145e10"
$ROLE_USER_ADMIN   = "fe930be7-5e62-47db-91af-98c3a49a38b1"
$ROLE_PEOPLE_ADMIN = "024906de-61e5-49c8-8572-40335f1e0e10"

try {
  if ($allowPhoto) {
    Write-Host "Allowing users to update profile photos..."
    Invoke-MgGraphRequest `
      -Method PATCH `
      -Uri "https://graph.microsoft.com/beta/admin/people/photoUpdateSettings" `
      -ContentType "application/json" `
      -Body (@{
        source = "cloud"
        allowedRoles = @()
      } | ConvertTo-Json) | Out-Null
  } else {
    Write-Host "Restricting profile photo updates to admins..."
    Invoke-MgGraphRequest `
      -Method PATCH `
      -Uri "https://graph.microsoft.com/beta/admin/people/photoUpdateSettings" `
      -ContentType "application/json" `
      -Body (@{
        source = "cloud"
        allowedRoles = @(
          $ROLE_GLOBAL_ADMIN,
          $ROLE_USER_ADMIN,
          $ROLE_PEOPLE_ADMIN
        )
      } | ConvertTo-Json) | Out-Null
  }
} catch {
  Write-Host "Warning: Could not update photo policy ($($_.Exception.Message))" -ForegroundColor Yellow
}

# -------------------------
# Done
# -------------------------
Write-Host ""
Write-Host "✔ Self-service profile configuration complete" -ForegroundColor Green
Write-Host ""
Write-Host "Summary:"
Write-Host "  Users can edit profile:   $allowSelfEdit"
Write-Host "  Users can edit contact:   $allowContact"
Write-Host "  Users can edit bio:       $allowBio"
Write-Host "  Users can change photo:   $allowPhoto"
Write-Host ""
Write-Host "Admin-only attributes remain admin-managed (by design)."
Write-Host "You can safely re-run this script at any time."
