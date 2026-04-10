from pathlib import Path

from desktop_gui.config import DEFAULT_TEMPLATE, LEGACY_TEMPLATE
from reporting.shared import (
    format_hours,
    load_json,
    normalize_short_date,
    parse_date,
    parse_hours,
    progress_rank,
    save_json,
    split_lines,
)


def resolve_template_path(path_text: str | None) -> Path:
    if path_text:
        path = Path(path_text)
    else:
        path = DEFAULT_TEMPLATE

    if path.exists():
        return path

    if path.name == LEGACY_TEMPLATE.name and DEFAULT_TEMPLATE.exists():
        return DEFAULT_TEMPLATE

    return path
