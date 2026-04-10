from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from desktop_gui.config import GUI_CONTEXT_PATH, GUI_REPORT_PATH, OUTPUT_DIR, ROTATION_HISTORY_PATH
from desktop_gui.utils import load_json, save_json
from reporting.context_builder import build_context
from reporting.rendering import render_docx


@dataclass(slots=True)
class GenerationResult:
    docx_path: Path
    pdf_path: Path | None
    context_path: Path
    report_path: Path


def _rotation_identity(item: dict) -> tuple[str, str, str, str]:
    return (
        str(item.get("area", "")).strip(),
        str(item.get("desde", "")).strip(),
        str(item.get("hasta", "")).strip(),
        str(item.get("semana", "")).strip(),
    )


def merge_rotation_history(history: list[dict], current: list[dict]) -> list[dict]:
    merged = []
    seen = set()

    for item in history + current:
        identity = _rotation_identity(item)
        if not any(identity):
            continue
        if identity in seen:
            continue
        seen.add(identity)
        merged.append(
            {
                "area": item.get("area", ""),
                "desde": item.get("desde", ""),
                "hasta": item.get("hasta", ""),
                "semana": item.get("semana", ""),
            }
        )

    return merged


def generate_report_files(
    *,
    template_path: Path,
    report_data: dict,
    pea_master: list[dict],
    output_name: str,
    export_pdf: bool,
    status_callback: Callable[[str], None] | None = None,
) -> GenerationResult:
    status_callback = status_callback or (lambda _: None)

    rotation_history = load_json(ROTATION_HISTORY_PATH, [])
    merged_rotations = merge_rotation_history(rotation_history, report_data.get("rotaciones", []))

    report_payload = dict(report_data)
    report_payload["rotaciones"] = merged_rotations

    save_json(GUI_REPORT_PATH, report_payload)
    context = build_context(report_payload, pea_master)
    save_json(GUI_CONTEXT_PATH, context)

    docx_path = OUTPUT_DIR / f"{output_name}.docx"
    status_callback("Generando DOCX...")
    render_docx(template_path, GUI_CONTEXT_PATH, docx_path)
    save_json(ROTATION_HISTORY_PATH, merged_rotations)

    pdf_path = None
    if export_pdf:
        from docx2pdf import convert

        pdf_path = OUTPUT_DIR / f"{output_name}.pdf"
        status_callback("Exportando PDF...")
        convert(str(docx_path), str(pdf_path))

    return GenerationResult(
        docx_path=docx_path,
        pdf_path=pdf_path,
        context_path=GUI_CONTEXT_PATH,
        report_path=GUI_REPORT_PATH,
    )
