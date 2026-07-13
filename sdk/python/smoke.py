"""CI smoke test (not shipped — pyproject ships only officecli.py).
Requires officecli on PATH and verifies the SDK pipe round-trip."""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import officecli  # noqa: E402

f = os.path.join(tempfile.gettempdir(), f"officecli-smoke-{os.getpid()}.xlsx")
d = officecli.create(f, "--force")
d.send({"command": "set", "path": "/Sheet1/A1", "props": {"text": "smoke-ok"}})
g = d.send({"command": "get", "path": "/Sheet1/A1"})
d.close()
try:
    os.unlink(f)
except OSError:
    pass

if "smoke-ok" not in str(g):
    print("python SDK smoke FAIL: A1 mismatch", g)
    sys.exit(1)
print("python SDK smoke PASS")
