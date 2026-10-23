"""Entry point: ``python -m libreoffice_mcp.worker`` or ``python /path/to/worker/__main__.py``.

The second form is what the server uses, because LibreOffice's bundled Python
cannot be assumed to have this package installed (it runs on the system Python,
not the server's virtualenv). We therefore put the package root on ``sys.path`` ourselves.
"""

import os
import sys

if not __package__:
    here = os.path.dirname(os.path.abspath(__file__))
    # Drop the worker dir itself so our module names can never shadow the stdlib.
    sys.path[:] = [p for p in sys.path if os.path.abspath(p or ".") != here]
    sys.path.insert(0, os.path.dirname(os.path.dirname(here)))

from libreoffice_mcp.worker.main import main

if __name__ == "__main__":
    sys.exit(main())
