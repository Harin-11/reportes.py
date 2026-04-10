import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from reporting.rendering import main, render_docx

__all__ = ["render_docx", "main"]


if __name__ == "__main__":
    main()
