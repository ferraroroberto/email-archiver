"""Compatibility shim for callers that only have the repo root on `sys.path`.

The package lives in `src/email_archiver/`. Another local app imports
`email_archiver.config` with `python -c` from this folder, and the README's
`python -m email_archiver <subcommand>` runs from here, so this module puts
`src/` first on the path and hands the import (or the `-m` run) to the real
package. New code imports the package from `src/`, never through this file.
"""
import importlib
import runpy
import sys
from pathlib import Path

# First on the path, so the lookups below find the package, not this file again.
sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

if __name__ == "__main__":
    runpy.run_module("email_archiver", run_name="__main__", alter_sys=True)
else:
    # The import system returns whatever sits in sys.modules once this module
    # finishes, so every importer gets the real package, never this shim.
    del sys.modules[__name__]
    sys.modules[__name__] = importlib.import_module(__name__)
