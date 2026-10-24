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
      3. docs ratchet  - scripts/docs_ratchet.py through scripts/docs_ratchet_gate.py:
                         fails only on README.md past its size budget or baselined size,
                         or a new history marker (issue reference, dated heading,
                         changelog heading) in README.md / docs/**/*.md that is not in
                         .docs-ratchet-baseline.json. An unreadable file (exit 3) fails
                         as NOT CONFIRMED. Not path-routed: a count needs the whole tree.
      4. structure ratchet - scripts/structure_ratchet.py through
                         scripts/structure_ratchet_gate.py: fails only on a file past 1,000
                         lines, a function past 100 lines or a directory past 40 modules that
                         is new or above its entry in .structure-ratchet-baseline.json, or a
                         root folder outside the .fleet.toml shape + extra_roots. A tree it
                         could not measure (exit 3) fails as NOT CONFIRMED. Not path-routed.
      5. instruction ratchet - scripts/instruction_ratchet.py through
                         scripts/instruction_ratchet_gate.py: fails only on a prompt-audit
                         rule count in CLAUDE.md / AGENTS.md above its entry in
                         .instruction-ratchet-baseline.json, or a file past its size cap. An
                         instruction file it could not measure (exit 3) fails as NOT
                         CONFIRMED. Not path-routed: one file's reader depends on another.
      6. doc refs ratchet - scripts/doc_refs_ratchet.py through
                         scripts/doc_refs_ratchet_gate.py: fails only on a repo path, `python -m`
                         module or relative link in README.md / CLAUDE.md / docs/**/*.md that
                         does not resolve against the tracked tree and is not in
                         .doc-refs-ratchet-baseline.json. A doc it could not read (exit 3)
                         fails as NOT CONFIRMED. Not path-routed: a code change can strand a
                         reference without touching the doc.

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
Invoke-Stage "docs ratchet"   { & $py -m scripts.docs_ratchet_gate }
Invoke-Stage "structure ratchet" { & $py -m scripts.structure_ratchet_gate }
Invoke-Stage "instruction ratchet" { & $py -m scripts.instruction_ratchet_gate }
Invoke-Stage "doc refs ratchet" { & $py -m scripts.doc_refs_ratchet_gate }

Write-Host ""
Write-Host "[PASS] verify-before-ship" -ForegroundColor Green
