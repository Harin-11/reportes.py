from __future__ import annotations

import unicodedata
from dataclasses import dataclass
from pathlib import Path


def _normalize_text(value: str) -> str:
    simplified = unicodedata.normalize("NFKD", str(value or ""))
    simplified = "".join(ch for ch in simplified if not unicodedata.combining(ch))
    return " ".join(simplified.upper().split())


def _as_text_lines(value) -> list[str]:
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    if value is None:
        return []
    return [line.strip() for line in str(value).splitlines() if line.strip()]


def _hours_label_to_clock(text: str) -> str:
    raw = str(text or "").strip().lower()
    if not raw:
        return ""

    hours = 0
    minutes = 0
    parts = raw.replace(",", " ").split()
    for index, token in enumerate(parts):
        if token.isdigit() and index + 1 < len(parts):
            label = parts[index + 1]
            if label.startswith("hora"):
                hours = int(token)
            elif label.startswith("minuto"):
                minutes = int(token)

    if hours == 0 and minutes == 0:
        return ""
    return f"{hours}:{minutes:02d}"


def _extract_section_hours(line: str, section: str) -> str:
    raw = str(line or "").strip()
    if not raw:
        return ""
    if section == "senati":
        _, _, suffix = raw.partition(":")
        return _hours_label_to_clock(suffix)
    return _hours_label_to_clock(raw.rsplit(":", 1)[-1])


def _lines_to_entries(lines: list[str], section: str) -> list[dict]:
    descriptions: list[str] = []
    total_hours = ""
    current_section = None

    for line in lines:
        normalized = _normalize_text(line)
        if normalized.startswith("SENATI:"):
            current_section = "senati"
            if section == current_section:
                total_hours = _extract_section_hours(line, current_section)
            continue
        if normalized.startswith("EMPRESA:"):
            current_section = "empresa"
            if section == current_section:
                total_hours = _extract_section_hours(line, current_section)
            continue
        if set(line.strip()) == {"-"}:
            continue
        if line.strip().startswith("-") and current_section == section:
            descriptions.append(line.strip()[1:].strip())
        elif current_section == section and line.strip():
            descriptions.append(line.strip())

    if not descriptions and not total_hours:
        return []

    description = " | ".join(descriptions) if descriptions else ("Actividad SENATI" if section == "senati" else "Actividad empresa")
    return [{"descripcion": description, "horas": total_hours}]


def _report_day_to_gui_day(day: dict) -> dict:
    lines = _as_text_lines(day.get("actividades"))
    return {
        "dia": day.get("dia", ""),
        "fecha": day.get("fecha", ""),
        "senati_entries": _lines_to_entries(lines, "senati"),
        "empresa_entries": _lines_to_entries(lines, "empresa"),
    }


def _report_week_to_gui_week(week: dict) -> dict:
    days = [_report_day_to_gui_day(day) for day in week.get("dias", [])]
    dated_days = [day.get("fecha", "") for day in days if day.get("fecha", "")]
    return {
        "numero_semana": week.get("numero_semana", ""),
        "desde": dated_days[0] if dated_days else "",
        "hasta": dated_days[-1] if dated_days else "",
        "dias": days,
    }


def _report_data_to_draft(payload: dict, template_path: str) -> dict:
    rotaciones = payload.get("rotaciones", [])
    first_rotation = rotaciones[0] if rotaciones else {}
    area_text = first_rotation.get("area", payload.get("area_empresa", ""))
    return {
        "profile": {
            "nombre_estudiante": payload.get("nombre_estudiante", ""),
            "id_estudiante": payload.get("id_estudiante", ""),
            "bloque": payload.get("bloque", ""),
            "carrera": payload.get("carrera", ""),
            "instructor": payload.get("instructor", ""),
            "semestre": payload.get("semestre", ""),
            "fecha_inicio_semestre": payload.get("fecha_inicio_semestre", ""),
            "fecha_fin_semestre": payload.get("fecha_fin_semestre", ""),
            "escuela_segmentada": payload.get("escuela", ""),
            "nombre_empresa": payload.get("nombre_empresa", ""),
            "area_empresa_segmentada": area_text,
            "area_empresa": area_text,
            "escuela": payload.get("escuela", ""),
        },
        "output_name": "informe_fpe",
        "export_pdf": False,
        "template_path": template_path,
        "weeks": [_report_week_to_gui_week(week) for week in payload.get("semanas", [])[:3]],
        "pea_avances": payload.get("pea_avances", {}),
        "report": {
            "tarea_significativa": payload.get("tarea_significativa", ""),
            "porque_tarea_significativa": payload.get("porque_tarea_significativa", ""),
            "descripcion_tarea_significativa": payload.get("descripcion_tarea_significativa", ""),
            "maquinas_usadas": payload.get("maquinas_usadas", ""),
            "equipos_usados": payload.get("equipos_usados", ""),
            "herramientas_usadas": payload.get("herramientas_usadas", ""),
            "materiales_usados": payload.get("materiales_usados", ""),
            "charlas_seguridad": payload.get("charlas_seguridad", ""),
            "explicacion_actividades_y_medidas_seguras_usadas": payload.get(
                "explicacion_actividades_y_medidas_seguras_usadas", ""
            ),
            "resultado_ejecucion": payload.get("resultado_ejecucion", ""),
            "objetivo_logrado_o_no": payload.get("objetivo_logrado_o_no", ""),
            "recomendaciones_resultado": payload.get("recomendaciones_resultado", ""),
            "recomendaciones_monitor": payload.get("recomendaciones_monitor", ""),
            "diagrama_path": payload.get("diagrama_path", ""),
            "firma_estudiante_path": payload.get("firma_estudiante_path", ""),
            "firma_monitor_path": payload.get("firma_monitor_path", ""),
        },
    }


def _has_meaningful_report_content(payload: dict) -> bool:
    scalar_keys = (
        "nombre_estudiante",
        "id_estudiante",
        "bloque",
        "carrera",
        "instructor",
        "semestre",
        "escuela",
        "nombre_empresa",
        "tarea_significativa",
        "recomendaciones_monitor",
    )
    list_keys = (
        "rotaciones",
        "semanas",
        "porque_tarea_significativa",
        "descripcion_tarea_significativa",
        "maquinas_usadas",
        "equipos_usados",
        "herramientas_usadas",
        "materiales_usados",
        "charlas_seguridad",
        "explicacion_actividades_y_medidas_seguras_usadas",
        "resultado_ejecucion",
        "objetivo_logrado_o_no",
        "recomendaciones_resultado",
    )

    if any(str(payload.get(key, "")).strip() for key in scalar_keys):
        return True
    return any(bool(payload.get(key)) for key in list_keys)


def _profile_to_draft(payload: dict, template_path: str) -> dict:
    return {
        "profile": dict(payload),
        "output_name": "informe_fpe",
        "export_pdf": False,
        "template_path": template_path,
        "weeks": [],
        "pea_avances": {},
        "report": {},
    }


def _pea_state_to_draft(payload: dict, template_path: str) -> dict:
    return {
        "profile": {},
        "output_name": "informe_fpe",
        "export_pdf": False,
        "template_path": template_path,
        "weeks": [],
        "pea_avances": payload,
        "report": {},
    }


@dataclass(slots=True)
class ImportedPayload:
    draft: dict
    kind: str
    message: str


def import_json_payload(payload: dict, *, template_path: str) -> ImportedPayload:
    if not isinstance(payload, dict):
        raise ValueError("El archivo JSON debe contener un objeto.")

    if {"profile", "weeks", "report"}.issubset(payload.keys()):
        return ImportedPayload(payload, "draft", "Borrador de GUI cargado.")

    if "semanas" in payload and "tarea_significativa" in payload:
        message = "Informe JSON cargado como base editable en la GUI."
        if not _has_meaningful_report_content(payload):
            message = "Plantilla base vacía cargada."
        return ImportedPayload(
            _report_data_to_draft(payload, template_path),
            "report",
            message,
        )

    profile_keys = {
        "nombre_estudiante",
        "id_estudiante",
        "bloque",
        "carrera",
        "instructor",
        "semestre",
        "fecha_inicio_semestre",
        "fecha_fin_semestre",
        "escuela",
        "escuela_segmentada",
        "nombre_empresa",
        "area_empresa",
        "area_empresa_segmentada",
    }
    if payload and set(payload.keys()).issubset(profile_keys):
        return ImportedPayload(_profile_to_draft(payload, template_path), "profile", "Perfil JSON cargado.")

    if payload and all(isinstance(value, dict) for value in payload.values()):
        return ImportedPayload(_pea_state_to_draft(payload, template_path), "pea", "Estado del PEA cargado.")

    raise ValueError(
        "Formato JSON no reconocido. Usa un borrador GUI, un informe JSON, un perfil o un estado del PEA."
    )
