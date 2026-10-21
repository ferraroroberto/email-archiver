#Requires -Version 5.1
<#
.SYNOPSIS
    Pre-ship verification gate for email-archiver. One pass/fail pipeline.

.DESCRIPTION
    Runs, fail-fast:
      1. pytest        - the unit suite (tests/); it never sends mail or touches a mailbox
      2. static ratchet - scripts/static_ratchet.py through scripts/static_ratchet_gate.py:
                         fails only on a duplicate block (jscpd via npx) or an unused
                         import/variable (ruff F401,F841) that is not in
                         .static-ratchet-baseline.json. A detector that cannot run
                         (exit 3) fails as NOT CONFIRMED, never a pass or a skip.
                         Needs Node and ruff in the .venv. Not path-routed: duplication
                         needs the whole tree.

    Anchors to the repo root, so run it from anywhere:  & .\scripts\verify-before-ship.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-Location (Join-Path $PSScriptRoot "..")

$py = ".\.venv\Scripts\python.exe"
if (-not (Test-Path $py)) {
    Write-Host "[FAIL] .venv not found at $py" -ForegroundColor Red
    Write-Host "       Create it and run: $py -m pip install -r requirements.txt" -ForegroundColor Red
    exit 1
}

function Invoke-Stage {
    param([string]$Name, [scriptblock]$Body)
    Write-Host ""
    Write-Host ">> $Name" -ForegroundColor Cyan
    & $Body
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[FAIL] $Name (exit $LASTEXITCODE)" -ForegroundColor Red
        exit 1
    }
    Write-Host "[PASS] $Name" -ForegroundColor Green
}

Invoke-Stage "pytest"         { & $py -m pytest tests/ }
Invoke-Stage "static ratchet" { & $py -m scripts.static_ratchet_gate }

Write-Host ""
Write-Host "[PASS] verify-before-ship" -ForegroundColor Green
