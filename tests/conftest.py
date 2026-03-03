import sys
from pathlib import Path

# Ensure the project root (where main.py, config_loader.py etc. live)
# is on sys.path when pytest imports modules from tests.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

