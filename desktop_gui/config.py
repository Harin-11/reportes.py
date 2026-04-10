from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT_DIR / "src"
DATA_DIR = ROOT_DIR / "data"
OUTPUT_DIR = ROOT_DIR / "output"
TEMPLATE_DIR = ROOT_DIR / "templates"

PROFILE_PATH = DATA_DIR / "gui_profile.json"
DRAFT_PATH = DATA_DIR / "gui_draft.json"
GUI_REPORT_PATH = DATA_DIR / "gui_report_data.json"
GUI_CONTEXT_PATH = DATA_DIR / "gui_report_context.json"
PEA_PATH = DATA_DIR / "pea_piads_iv.json"
PEA_PROGRESS_PATH = DATA_DIR / "pea_progress_state.json"
ROTATION_HISTORY_PATH = DATA_DIR / "rotation_history.json"

LEGACY_TEMPLATE = OUTPUT_DIR / "plantilla_informe_docxtpl_pulida.docx"
DEFAULT_TEMPLATE = TEMPLATE_DIR / "plantilla_informe_fpe.docx"

WEEKDAY_NAMES = ["LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO"]

BG = "#F4F6FB"
PANEL = "#FFFFFF"
PRIMARY = "#174A8B"
TEXT = "#1F2937"
MUTED = "#6B7280"

FONT = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 14, "bold")
FONT_SMALL = ("Segoe UI", 9)


def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
