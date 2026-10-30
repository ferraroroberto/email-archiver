#Requires -Version 5.1
<#
.SYNOPSIS
    Pre-ship verification gate for email-archiver. One pass/fail pipeline.

.DESCRIPTION
    Runs, fail-fast:
      1. pytest        - the unit suite (tests/); it never sends mail or touches a mailbox
      2. static ratchet - scripts/static_ratchet.py through scripts/static_ratchet_gate.py:
                         fails only on a duplicate block (jscpd via npx), an unused
                         import/variable (ruff F401,F841) or a ruff bug (B) / security
                         (S) finding (S113 strict) that is not in
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
      7. dead code ratchet - scripts/dead_code_ratchet.py through
                         scripts/dead_code_ratchet_gate.py: fails only on a function, class or
                         method that nothing in the tracked tree refers to and that is not in
                         .dead-code-ratchet-baseline.json. A Python file it could not parse
                         (exit 3) fails as NOT CONFIRMED. Not path-routed: deleting the last
                         caller in one file strands a definition in another.
      8. conventions ratchet - scripts/conventions_ratchet.py through
                         scripts/conventions_ratchet_gate.py: fails only on a no-window,
                         logging-not-print or no-drive-path hit count in a file above its
                         entry in .conventions-ratchet-baseline.json. A Python file it could
                         not parse (exit 3) fails as NOT CONFIRMED. Not path-routed: a rule
                         added to .fleet.toml must see the whole tree.
      9. comment ratchet - scripts/comment_ratchet.py through
                         scripts/comment_ratchet_gate.py: fails only on a file whose count of
                         `#NNN` issue references inside comments rises above its entry in
                         .comment-ratchet-baseline.json. A file it could not scan (exit 3)
                         fails as NOT CONFIRMED. Not path-routed: a count needs the whole tree.

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

. (Join-Path $PSScriptRoot "box_lock_gate.ps1"); Enter-BoxLockGate

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
Invoke-Stage "dead code ratchet" { & $py -m scripts.dead_code_ratchet_gate }
Invoke-Stage "conventions ratchet" { & $py -m scripts.conventions_ratchet_gate }
Invoke-Stage "comment ratchet" { & $py -m scripts.comment_ratchet_gate }

Write-Host ""
Write-Host "[PASS] verify-before-ship" -ForegroundColor Green
