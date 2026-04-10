from __future__ import annotations

import argparse
from pathlib import Path

from reporting.shared import get_image_pixel_size, load_json


def normalize_lines(value) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    text = str(value).strip()
    if not text:
        return []
    return [line.strip() for line in text.splitlines() if line.strip()]


def warn(message: str) -> None:
    print(f"[WARN] {message}")


def info(message: str) -> None:
    print(f"[OK] {message}")


def check_weeks(data: dict) -> None:
    weeks = data.get("semanas", [])
    for week in weeks[:3]:
        numero = week.get("numero_semana", "?")
        for day in week.get("dias", []):
            dia = day.get("dia", "DIA")
            lines = normalize_lines(day.get("actividades"))
            dia_normalized = dia.upper()
            if len(lines) > 8:
                warn(f"Semana {numero} {dia}: {len(lines)} líneas de actividades. Riesgo medio de que la celda crezca demasiado.")
            elif not lines and dia_normalized not in {"SÁBADO", "SABADO"}:
                warn(f"Semana {numero} {dia}: no tiene actividades.")


def check_rotaciones(data: dict) -> None:
    rotaciones = data.get("rotaciones", [])
    if len(rotaciones) > 10:
        warn(f"Hay {len(rotaciones)} rotaciones. Puede que la tabla pase a otra página.")
    else:
        info(f"Rotaciones dentro de un rango normal: {len(rotaciones)} filas.")


def check_narrative_field(name: str, value, max_lines: int, max_chars: int) -> None:
    lines = normalize_lines(value)
    char_count = sum(len(line) for line in lines)

    if len(lines) > max_lines or char_count > max_chars:
        warn(
            f"{name}: {len(lines)} líneas y {char_count} caracteres. Conviene resumir para evitar desbordes visuales."
        )
    else:
        info(f"{name}: tamaño razonable.")


def check_diagram_image(path_text: str) -> None:
    raw = str(path_text or "").strip()
    if not raw:
        info("diagrama_path: no se adjuntó imagen.")
        return

    path = Path(raw)
    if not path.exists():
        warn("diagrama_path: la imagen indicada no existe.")
        return

    size = get_image_pixel_size(path)
    if not size:
        warn("diagrama_path: no se pudo leer el tamaño de la imagen; usa PNG o JPG.")
        return

    width_px, height_px = size
    ratio = height_px / width_px if width_px else 0
    if ratio > 1.1:
        warn(
            f"diagrama_path: imagen vertical ({width_px}x{height_px}). Conviene usar una captura más horizontal para no descuadrar la tabla."
        )
    elif width_px > 2400 or height_px > 1800:
        warn(
            f"diagrama_path: imagen grande ({width_px}x{height_px}). El render la reducirá, pero conviene recortarla antes."
        )
    else:
        info(f"diagrama_path: tamaño razonable ({width_px}x{height_px}).")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Revisa riesgos visuales comunes antes de renderizar el informe."
    )
    parser.add_argument("--data", required=True, type=Path)
    args = parser.parse_args()

    data = load_json(args.data, {})
    check_rotaciones(data)
    check_weeks(data)
    check_narrative_field("porque_tarea_significativa", data.get("porque_tarea_significativa"), 4, 420)
    check_narrative_field("descripcion_tarea_significativa", data.get("descripcion_tarea_significativa"), 8, 1300)
    check_narrative_field("charlas_seguridad", data.get("charlas_seguridad"), 6, 300)
    check_narrative_field(
        "explicacion_actividades_y_medidas_seguras_usadas",
        data.get("explicacion_actividades_y_medidas_seguras_usadas"),
        6,
        420,
    )
    check_narrative_field("resultado_ejecucion", data.get("resultado_ejecucion"), 5, 450)
    check_narrative_field("recomendaciones_resultado", data.get("recomendaciones_resultado"), 6, 400)
    check_diagram_image(data.get("diagrama_path", ""))


if __name__ == "__main__":
    main()
