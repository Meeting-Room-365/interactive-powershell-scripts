#Requires -Modules ExchangeOnlineManagement
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Interactive Microsoft 365 Anti-Spam Setup Wizard

.DESCRIPTION
  Sets up sane anti-spam and anti-phishing policies in Microsoft 365
  using Exchange Online / Defender cmdlets.

  - Single-file interactive wizard
  - Tenant-wide or group-based
  - Idempotent (safe to re-run)
  - Writes a JSON profile next to the script

.NOTES
  This intentionally avoids Preset Security Policies.
  Safe to test on tenants with default / poor configurations.
#>

$ErrorActionPreference = "Stop"

# -------------------------
# Initialization
# -------------------------
$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProfilePath = Join-Path $ScriptDir "m365-antispam-profile.json"

Clear-Host
Write-Host "Microsoft 365 Anti-Spam Setup Wizard" -ForegroundColor Cyan
Write-Host "===================================" -ForegroundColor Cyan
Write-Host ""

# -------------------------
# Connect to Exchange Online
# -------------------------
Write-Host "Connecting to Exchange Online..."
Connect-ExchangeOnline -ShowBanner:$false

$org = Get-OrganizationConfig
$tenantName = $org.Name

Write-Host "Connected to tenant:" -NoNewline
Write-Host " $tenantName" -ForegroundColor Green
Write-Host ""

# -------------------------
# Preset policy warning (best-effort)
# -------------------------
try {
  $presets = Get-PolicyConfig | Where-Object Name -Like "*Preset*"
  if ($presets) {
    Write-Warning "Preset security policies detected. Custom policies may be overridden."
    Write-Host ""
  }
} catch {
  Write-Verbose "Preset detection skipped"
}

# -------------------------
# Check for existing profile
# -------------------------
$useProfile = $false
$savedProfile = $null

if (Test-Path $ProfilePath) {
  try {
    $savedProfile = Get-Content $ProfilePath -Raw | ConvertFrom-Json
    Write-Host "Found existing profile:" -ForegroundColor Yellow
    Write-Host "  Created: $($savedProfile.createdAt)"
    Write-Host "  Scope:   $($savedProfile.scope.mode)"
    if ($savedProfile.scope.groupName) {
      Write-Host "  Group:   $($savedProfile.scope.groupName)"
    }
    Write-Host "  Bulk threshold: $($savedProfile.antiSpam.bulkThreshold)"
    Write-Host "  Impersonation:  $($savedProfile.antiPhish.enableImpersonation)"
    Write-Host ""
    $useProfileResponse = Read-Host "Use this profile? (y/n) [y]"
    if (-not $useProfileResponse -or $useProfileResponse -eq "y") {
      $useProfile = $true
    }
    Write-Host ""
  } catch {
    Write-Warning "Could not load profile from $ProfilePath. Starting fresh."
    Write-Host ""
  }
}

# -------------------------
# Wizard: Scope
# -------------------------
if ($useProfile) {
  $scopeMode = $savedProfile.scope.mode
  $groupName = $savedProfile.scope.groupName
  Write-Host "Using profile settings: Scope = $scopeMode" -ForegroundColor Cyan
  if ($scopeMode -eq "group" -and $groupName) {
    Write-Host "  Group: $groupName" -ForegroundColor Cyan
  }
} else {
  Write-Host "Policy scope"
  Write-Host "------------"
  $scopeMode = Read-Host "Apply policy to (tenant/group) [tenant]"
  if (-not $scopeMode) { $scopeMode = "tenant" }

  $groupName = $null
  if ($scopeMode -eq "group") {
    $groupName = Read-Host "Enter mail-enabled security or distribution group name"
  }
}

Write-Host ""

# -------------------------
# Wizard: Sensitivity
# -------------------------
if ($useProfile) {
  $bulkLevel = [int]$savedProfile.antiSpam.bulkThreshold
  $fpTolerance = if ($savedProfile.quarantine.userCanReleaseSpam) { "medium" } else { "low" }
  $impersonation = if ($savedProfile.antiPhish.enableImpersonation) { "y" } else { "n" }
  Write-Host "Using profile settings:" -ForegroundColor Cyan
  Write-Host "  Bulk threshold: $bulkLevel" -ForegroundColor Cyan
  Write-Host "  False-positive tolerance: $fpTolerance" -ForegroundColor Cyan
  Write-Host "  Impersonation protection: $impersonation" -ForegroundColor Cyan
} else {
  Write-Host "Spam sensitivity"
  Write-Host "---------------"
  $bulkLevel = Read-Host "Bulk email strictness (1=strict, 7=loose) [5]"
  if (-not $bulkLevel) { $bulkLevel = 5 }
  $bulkLevel = [int]$bulkLevel

  $fpTolerance = Read-Host "False-positive tolerance (low/medium/high) [medium]"
  if (-not $fpTolerance) { $fpTolerance = "medium" }

  $impersonation = Read-Host "Enable impersonation protection? (y/n) [y]"
  if (-not $impersonation) { $impersonation = "y" }
}

Write-Host ""

# -------------------------
# Build profile object
# -------------------------
$profile = @{
  name       = "Wizard Anti-Spam Policy"
  tenant     = $tenantName
  createdAt = (Get-Date).ToString("s")

  scope = @{
    mode      = $scopeMode
    groupName = $groupName
  }

  antiSpam = @{
    spamAction                = "Quarantine"
    highConfidenceSpamAction  = "Quarantine"
    phishAction               = "Quarantine"
    highConfidencePhishAction = "Quarantine"
    bulkThreshold             = [int]$bulkLevel
  }

  antiPhish = @{
    enableImpersonation = ($impersonation -eq "y")
    action              = "Quarantine"
  }

  quarantine = @{
    userCanReleaseSpam  = ($fpTolerance -ne "low")
    userCanReleasePhish = $false
  }
}

# -------------------------
# Save profile locally
# -------------------------
$profile | ConvertTo-Json -Depth 6 | Out-File $ProfilePath -Encoding UTF8
Write-Host "Saved profile to $ProfilePath"
Write-Host ""

# -------------------------
# Anti-Spam Policy (Hosted Content Filter)
# -------------------------
$policyName = $profile.name

$existingPolicy = Get-HostedContentFilterPolicy `
  -Identity $policyName `
  -ErrorAction SilentlyContinue

if ($existingPolicy) {
  Write-Host "Updating existing anti-spam policy..."
  Set-HostedContentFilterPolicy `
    -Identity $policyName `
    -SpamAction Quarantine `
    -HighConfidenceSpamAction Quarantine `
    -PhishSpamAction Quarantine `
    -HighConfidencePhishAction Quarantine `
    -BulkThreshold $profile.antiSpam.bulkThreshold
} else {
  Write-Host "Creating anti-spam policy..."
  New-HostedContentFilterPolicy `
    -Name $policyName `
    -SpamAction Quarantine `
    -HighConfidenceSpamAction Quarantine `
    -PhishSpamAction Quarantine `
    -HighConfidencePhishAction Quarantine `
    -BulkThreshold $profile.antiSpam.bulkThreshold
}

# -------------------------
# Scope Rule (tenant or group)
# -------------------------
$ruleName = "$policyName Rule"

$existingRule = Get-HostedContentFilterRule `
  -Identity $ruleName `
  -ErrorAction SilentlyContinue

if ($existingRule) {
  Write-Host "Updating policy scope rule..."
  if ($profile.scope.mode -eq "group") {
    $group = Get-DistributionGroup $profile.scope.groupName
    Set-HostedContentFilterRule `
      -Identity $ruleName `
      -SentToMemberOf $group.Name
  } else {
    # For tenant-wide, scope to all accepted domains
    $domains = (Get-AcceptedDomain).DomainName
    Set-HostedContentFilterRule `
      -Identity $ruleName `
      -RecipientDomainIs $domains
  }
} else {
  Write-Host "Creating policy scope rule..."
  if ($profile.scope.mode -eq "group") {
    $group = Get-DistributionGroup $profile.scope.groupName
    New-HostedContentFilterRule `
      -Name $ruleName `
      -HostedContentFilterPolicy $policyName `
      -SentToMemberOf $group.Name
  } else {
    # For tenant-wide, create rule scoped to all accepted domains
    $domains = (Get-AcceptedDomain).DomainName
    New-HostedContentFilterRule `
      -Name $ruleName `
      -HostedContentFilterPolicy $policyName `
      -RecipientDomainIs $domains
  }
}

# -------------------------
# Anti-Phish Policy (Impersonation)
# -------------------------
if ($profile.antiPhish.enableImpersonation) {
  $existingPhish = Get-AntiPhishPolicy `
    -Identity $policyName `
    -ErrorAction SilentlyContinue

  if (-not $existingPhish) {
    Write-Host "Enabling anti-phishing protections..."
    New-AntiPhishPolicy `
      -Name $policyName
  } else {
    Write-Host "Anti-phishing policy already exists."
  }
}

# -------------------------
# Final summary
# -------------------------
Write-Host ""
Write-Host "✔ Anti-spam configuration complete" -ForegroundColor Green
Write-Host "----------------------------------"
Write-Host "Tenant:          $tenantName"
Write-Host "Scope:           $($profile.scope.mode)"
if ($profile.scope.mode -eq "group") {
  Write-Host "Group:           $($profile.scope.groupName)"
}
Write-Host "Bulk threshold:  $bulkLevel"
Write-Host "Impersonation:   $($profile.antiPhish.enableImpersonation)"
Write-Host ""
Write-Host "You can safely re-run this script at any time."
