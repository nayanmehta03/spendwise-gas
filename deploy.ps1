<#
.SYNOPSIS
  Push local Apps Script sources with clasp and redeploy Spendwise to its
  FIXED web-app deployment, so the live /exec URL never changes.

.DESCRIPTION
  Bare `clasp deploy` creates a BRAND NEW deployment with a BRAND NEW URL.
  This script always targets an existing deployment id, so bookmarks, the
  Google Chat app connection, and any email "Log it" links keep working.

  Steps: preflight -> auth probe -> file-list check -> push -> version -> redeploy.

.PARAMETER Description
  Deployment description. Defaults to "<git-short-sha> <branch> - <timestamp>".

.PARAMETER DeploymentId
  Override the deployment id. Normally read from deploy.config.json or the
  SPENDWISE_DEPLOYMENT_ID environment variable.

.PARAMETER DryRun
  Run every check and print the file list, but push and deploy nothing.

.PARAMETER SkipPush
  Redeploy the code already on script.google.com; do not push local files.

.PARAMETER Force
  Do not stop for the "you have uncommitted changes" confirmation.

.PARAMETER Open
  Open the live web app in the browser once the deploy succeeds.

.EXAMPLE
  .\deploy.ps1

.EXAMPLE
  .\deploy.ps1 -Description "analytics filter fix" -Open

.EXAMPLE
  .\deploy.ps1 -DryRun
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
  [string] $Description,
  [string] $DeploymentId,
  [switch] $DryRun,
  [switch] $SkipPush,
  [switch] $Force,
  [switch] $Open
)

# 'Continue', not 'Stop': in Windows PowerShell 5.1 a native exe writing to stderr
# is surfaced as a NativeCommandError, which under 'Stop' aborts the script before
# it can print a useful message. Every clasp/git call below checks $LASTEXITCODE
# explicitly instead, and cmdlet calls that can fail are wrapped in try/catch.
$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

# ----------------------------------------------------------------- output ---
function Write-Step  ($m) { Write-Host ""; Write-Host "==> $m" -ForegroundColor Cyan }
function Write-Ok    ($m) { Write-Host "    OK  $m" -ForegroundColor Green }
function Write-Note  ($m) { Write-Host "    $m"     -ForegroundColor DarkGray }
function Write-Warn2 ($m) { Write-Host "    !   $m" -ForegroundColor Yellow }
function Fail ($m, $hint) {
  Write-Host ""
  Write-Host "FAILED: $m" -ForegroundColor Red
  if ($hint) { Write-Host "        $hint" -ForegroundColor Yellow }
  Write-Host ""
  exit 1
}

Write-Host ""
Write-Host "  Spendwise deploy" -ForegroundColor White
Write-Host "  ----------------" -ForegroundColor DarkGray

# -------------------------------------------------------------- preflight ---
Write-Step "Preflight"

if (-not (Get-Command clasp -ErrorAction SilentlyContinue)) {
  Fail "clasp is not on PATH." "Install it with:  npm install -g @google/clasp"
}
# Note: no `| Select-Object -First 1` here - early pipeline termination kills the
# clasp shim and reports a bogus non-zero exit code.
$claspVersionLines = @(& clasp --version)
if ($LASTEXITCODE -ne 0) { Fail "Could not run clasp." }
$claspVersion = ($claspVersionLines | Where-Object { $_ } | Select-Object -Last 1)
Write-Ok "clasp $claspVersion"

if (-not (Test-Path '.clasp.json')) {
  Fail ".clasp.json not found in $Root." "It is gitignored on purpose (it holds the private scriptId). Restore it, or run: clasp clone-script <scriptId>"
}
try { $claspCfg = Get-Content '.clasp.json' -Raw | ConvertFrom-Json }
catch { Fail ".clasp.json is not valid JSON." $_.Exception.Message }
if (-not $claspCfg.scriptId) { Fail ".clasp.json has no scriptId." }
Write-Ok "scriptId $($claspCfg.scriptId.Substring(0, 12))..."

if (-not (Test-Path '.claspignore')) {
  Write-Warn2 ".claspignore is missing - clasp would push README, screenshots and legacy files."
}

# ------------------------------------------------ resolve the deployment ----
$webAppUrl = $null
if (-not $DeploymentId) { $DeploymentId = $env:SPENDWISE_DEPLOYMENT_ID }
if (-not $DeploymentId -and (Test-Path 'deploy.config.json')) {
  try { $deployCfg = Get-Content 'deploy.config.json' -Raw | ConvertFrom-Json }
  catch { Fail "deploy.config.json is not valid JSON." $_.Exception.Message }
  $DeploymentId = $deployCfg.deploymentId
  $webAppUrl    = $deployCfg.webAppUrl
}
if (-not $DeploymentId) {
  Fail "No deployment id." "Copy deploy.config.example.json to deploy.config.json and fill in deploymentId (it is gitignored), or pass -DeploymentId."
}
if ($DeploymentId -notmatch '^AKfyc[A-Za-z0-9_-]{20,}$') {
  Fail "'$DeploymentId' does not look like a deployment id." "Use the AKfyc... segment of the /macros/s/<id>/exec URL - not the scriptId, not the whole URL."
}
if (-not $webAppUrl) { $webAppUrl = "https://script.google.com/macros/s/$DeploymentId/exec" }
Write-Ok "deployment $($DeploymentId.Substring(0, 16))..."

# -------------------------------------------------------------- git state ---
$sha = 'nogit'; $branch = 'nogit'; $dirty = @()
if (Test-Path '.git') {
  $sha    = (& git rev-parse --short HEAD)
  $branch = (& git rev-parse --abbrev-ref HEAD)
  $dirty  = @(& git status --porcelain | Where-Object { $_ })

  if ($dirty.Count -gt 0) {
    Write-Warn2 "$($dirty.Count) uncommitted change(s) on '$branch' - these WILL go live:"
    $dirty | Select-Object -First 12 | ForEach-Object { Write-Note "  $_" }
    if ($dirty.Count -gt 12) { Write-Note "  ... and $($dirty.Count - 12) more" }

    if (-not $Force -and -not $DryRun) {
      $answer = Read-Host "    Deploy uncommitted changes anyway? [y/N]"
      if ($answer -notmatch '^(y|yes)$') {
        Write-Host "    Aborted." -ForegroundColor DarkGray
        exit 130
      }
    }
  }
  else {
    Write-Ok "working tree clean ($branch @ $sha)"
  }
}

if (-not $Description) {
  $Description = "$sha $branch - " + (Get-Date -Format 'yyyy-MM-dd HH:mm')
}
if ($Description.Length -gt 200) { $Description = $Description.Substring(0, 200) }

# ------------------------------------------------------------- auth probe ---
Write-Step "Checking Google authorization"
$errFile = Join-Path ([System.IO.Path]::GetTempPath()) "spendwise-clasp-auth.err"
$null = & clasp list-deployments 2>$errFile
$authExit = $LASTEXITCODE
$authErr  = ''
if (Test-Path $errFile) {
  $authErr = (Get-Content $errFile -Raw -ErrorAction SilentlyContinue)
  Remove-Item $errFile -ErrorAction SilentlyContinue
}

if ($authExit -ne 0) {
  # Pull the one useful line out of clasp's stderr, dropping the PowerShell stack frames.
  $detail = $null
  $jm = [regex]::Match($authErr, '"error_description"\s*:\s*"([^"]+)"')
  if ($jm.Success) {
    $detail = $jm.Groups[1].Value
  }
  elseif ($authErr) {
    $detail = ($authErr -split "`n" |
      Where-Object { $_.Trim() -and $_ -notmatch '^\s*(At |\+|~)' } |
      Select-Object -First 1)
  }
  if ($detail) { Write-Note $detail.Trim() }

  if ($authErr -match 'invalid_grant|invalid_rapt|unauthorized|401') {
    Fail "clasp is not authorized - the saved token has expired." "Run:  clasp login    then re-run this script. (Only you can complete that browser flow.)"
  }
  Fail "clasp could not reach the Apps Script API." "Check your network, then try:  clasp login"
}
Write-Ok "authorized"

# ---------------------------------------------------------- files to push ---
Write-Step "Files clasp will push"
$status = @(& clasp show-file-status)
if ($LASTEXITCODE -ne 0) { Fail "clasp show-file-status failed." }

# clasp prints a "Tracked files:" block followed by an "Untracked files:" block.
# Only the tracked block is pushed - scanning the whole output would flag every
# gitignored/ignored file as a leak.
$tracked   = @()
$inTracked = $false
foreach ($line in $status) {
  if ($line -match '^\s*Tracked files:')   { $inTracked = $true;  continue }
  if ($line -match '^\s*Untracked files:') { $inTracked = $false; continue }
  if ($inTracked -and $line.Trim()) {
    $tracked += ($line -replace '^[^A-Za-z0-9_.]+', '').Trim()   # strip the tree glyphs
  }
}
if ($tracked.Count -eq 0) {
  Write-Warn2 "Could not parse the tracked-file list; showing raw output instead."
  $status | ForEach-Object { Write-Note $_ }
  $tracked = @($status | Where-Object { $_.Trim() })
}
else {
  $tracked | ForEach-Object { Write-Note $_ }
  Write-Note "($($tracked.Count) files)"
}

# Guard 1: files that must never reach the script project.
#   shard-functions.js re-declares functions that live in AdminOps.js, which is a
#   project-wide duplicate-declaration error the moment it lands.
foreach ($leak in @('shard-functions.js', 'README.md', 'DESIGN_SPEC.md', 'CLAUDE.md', 'deploy.ps1', 'deploy.sh', 'deploy.config.json')) {
  if ($tracked -contains $leak) {
    Fail "'$leak' is in the push set." "Restore the .claspignore whitelist - pushing it breaks the script project."
  }
}

# Guard 2: .claspignore is a whitelist, so a new source file silently does not
# deploy until it is listed there. Catch that before it becomes a mystery.
$knownExcluded = @('shard-functions.js')
$orphans = @(Get-ChildItem -File -Path (Join-Path $Root '*') -Include '*.js', '*.html' -ErrorAction SilentlyContinue |
  Where-Object { $tracked -notcontains $_.Name -and $knownExcluded -notcontains $_.Name } |
  ForEach-Object { $_.Name })
if ($orphans.Count -gt 0) {
  Write-Warn2 "Not in the .claspignore whitelist - these will NOT deploy:"
  $orphans | ForEach-Object { Write-Note "  $_" }
  Write-Note "Add '!$($orphans[0])' to .claspignore if it is meant to ship."
}

if ($DryRun) {
  Write-Host ""
  Write-Host "  Dry run - nothing pushed, nothing deployed." -ForegroundColor Yellow
  Write-Note "Would deploy to: $webAppUrl"
  Write-Note "Description:     $Description"
  Write-Host ""
  exit 0
}

# -------------------------------------------------------------------- push ---
if ($SkipPush) {
  Write-Step "Skipping push (-SkipPush)"
}
else {
  Write-Step "Pushing sources"
  & clasp push --force
  if ($LASTEXITCODE -ne 0) { Fail "clasp push failed - nothing was deployed." "Fix the error above; the live app is untouched." }
  Write-Ok "pushed"
}

# ----------------------------------------------------------------- version ---
Write-Step "Creating version"
$versionOut = & clasp create-version $Description
if ($LASTEXITCODE -ne 0) { Fail "clasp create-version failed - the live app is untouched." }
$versionOut | ForEach-Object { Write-Note $_ }

$versionText   = ($versionOut | Out-String)
$versionNumber = $null
$m = [regex]::Match($versionText, '(?i)version\s+(\d+)')
if ($m.Success) {
  $versionNumber = $m.Groups[1].Value
}
else {
  $all = [regex]::Matches($versionText, '\b(\d+)\b')
  if ($all.Count -gt 0) { $versionNumber = $all[$all.Count - 1].Groups[1].Value }
}

# ---------------------------------------------------------------- redeploy ---
Write-Step "Redeploying to the existing deployment"
if ($versionNumber) {
  Write-Note "version $versionNumber -> $($DeploymentId.Substring(0, 16))..."
  & clasp update-deployment -V $versionNumber -d $Description $DeploymentId
}
else {
  Write-Warn2 "Could not read the version number; falling back to create-deployment -i."
  & clasp create-deployment -i $DeploymentId -d $Description
}
if ($LASTEXITCODE -ne 0) {
  Fail "Redeploy failed." "Code IS pushed, but the live deployment still serves the old version. Re-run with -SkipPush once fixed."
}

# -------------------------------------------------------------------- done ---
$stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
"$stamp`tv$versionNumber`t$sha`t$branch`t$Description" | Add-Content -Path '.deploy-history.log' -Encoding utf8

Write-Host ""
Write-Host "  Deployed." -ForegroundColor Green
Write-Host "  version : $versionNumber" -ForegroundColor Gray
Write-Host "  desc    : $Description"   -ForegroundColor Gray
Write-Host "  live    : $webAppUrl"     -ForegroundColor White
Write-Host ""
Write-Note "Hard-reload the tab (Ctrl+Shift+R) - Apps Script caches the HTML shell."
Write-Host ""

if ($Open) { Start-Process $webAppUrl }
