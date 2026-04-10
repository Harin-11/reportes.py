from __future__ import annotations

import argparse
from pathlib import Path

from reporting.shared import load_json, save_json


def merge_pea_items(pea_items: list[dict], avances: dict) -> list[dict]:
    merged = []
    include_all = bool(avances.get("__include_all__")) if isinstance(avances, dict) else False

    for item in pea_items:
        numero = item.get("numero", "")
        avance = avances.get(numero, {}) if isinstance(avances, dict) else {}
        merged_item = dict(item)
        merged_item["seminario"] = ""

        for key in ("op1", "op2", "op3", "op4"):
            if key in avance:
                merged_item[key] = avance[key]

        has_progress = any(str(merged_item.get(key, "")).strip() for key in ("op1", "op2", "op3", "op4"))
        if include_all or has_progress:
            merged.append(merged_item)

    return merged


def build_context(report_data: dict, pea_items: list[dict]) -> dict:
    context = dict(report_data)
    avances = context.pop("pea_avances", {})
    context["pea_items"] = merge_pea_items(pea_items, avances)
    return context


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Construye el contexto final de un informe uniendo datos generales con el PEA maestro."
    )
    parser.add_argument("--report-data", required=True, type=Path)
    parser.add_argument("--pea", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()

    report_data = load_json(args.report_data, {})
    pea_items = load_json(args.pea, [])
    context = build_context(report_data, pea_items)
    save_json(args.output, context)
    print(f"Contexto generado en: {args.output}")


if __name__ == "__main__":
    main()
