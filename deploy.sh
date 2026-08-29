#!/usr/bin/env bash
# Thin wrapper so `./deploy.sh` works from Git Bash / WSL-with-Windows-PowerShell.
# All the logic lives in deploy.ps1 - there is deliberately only one implementation.
#
#   ./deploy.sh
#   ./deploy.sh -DryRun
#   ./deploy.sh -Description "analytics filter fix" -Open
set -euo pipefail
here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if command -v pwsh >/dev/null 2>&1; then
  shell=pwsh
elif command -v powershell >/dev/null 2>&1; then
  shell=powershell
else
  echo "Neither pwsh nor powershell found on PATH. Run deploy.ps1 from a PowerShell prompt." >&2
  exit 1
fi

exec "$shell" -NoProfile -ExecutionPolicy Bypass -File "$here/deploy.ps1" "$@"
