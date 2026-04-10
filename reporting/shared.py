from __future__ import annotations

import json
import struct
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any

PROGRESS_KEYS = ("op1", "op2", "op3", "op4")


def load_json(path: Path, default: Any = None) -> Any:
    if not path.exists():
        return default
    with path.open("r", encoding="utf-8-sig") as fh:
        return json.load(fh)


def normalize_saved_data(value: Any) -> Any:
    if isinstance(value, dict):
        return {key: normalize_saved_data(child) for key, child in value.items()}
    if isinstance(value, list):
        return [normalize_saved_data(item) for item in value]
    if isinstance(value, str):
        return unicodedata.normalize("NFC", value).replace("\r\n", "\n").replace("\r", "\n")
    return value


def save_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(normalize_saved_data(data), fh, ensure_ascii=False, indent=2)


def parse_hours(value: str) -> float:
    raw = (value or "").strip().replace(",", ".")
    if not raw:
        return 0.0
    if ":" in raw:
        hour_text, minute_text = raw.split(":", 1)
        hours = int(hour_text or 0)
        minutes = int(minute_text or 0)
        return max(hours, 0) + max(minutes, 0) / 60.0
    return max(float(raw), 0.0)


def format_hours(value: float) -> str:
    total_minutes = round(value * 60)
    hours = total_minutes // 60
    minutes = total_minutes % 60
    parts = []
    if hours:
        parts.append(f"{hours} hora" if hours == 1 else f"{hours} horas")
    if minutes:
        parts.append(f"{minutes} minuto" if minutes == 1 else f"{minutes} minutos")
    return " ".join(parts)


def split_lines(text: str) -> list[str]:
    return [line.strip() for line in str(text or "").splitlines() if line.strip()]


def parse_date(text: str) -> datetime:
    return datetime.strptime(text.strip(), "%d/%m/%Y")


def normalize_short_date(text: str) -> str:
    if not text:
        return ""
    try:
        return parse_date(text).strftime("%d/%m")
    except ValueError:
        parts = text.strip().split("/")
        if len(parts) >= 2:
            return f"{parts[0].zfill(2)}/{parts[1].zfill(2)}"
        return text.strip()


def progress_rank(progress: dict | None) -> int:
    current = progress or {}
    for rank, key in enumerate(PROGRESS_KEYS, start=1):
        if str(current.get(key, "")).strip():
            return rank
    return 0


def join_text_lines(value: Any) -> str:
    if isinstance(value, list):
        return "\n".join(str(item) for item in value if str(item).strip())
    return str(value or "")


def get_image_pixel_size(path: Path) -> tuple[int, int] | None:
    try:
        with path.open("rb") as fh:
            header = fh.read(32)
            if header.startswith(b"\x89PNG\r\n\x1a\n") and len(header) >= 24:
                width, height = struct.unpack(">II", header[16:24])
                return width, height

            if header[:2] == b"\xff\xd8":
                fh.seek(2)
                while True:
                    marker_prefix = fh.read(1)
                    if not marker_prefix:
                        return None
                    if marker_prefix != b"\xff":
                        continue

                    marker_type = fh.read(1)
                    while marker_type == b"\xff":
                        marker_type = fh.read(1)
                    if not marker_type or marker_type in {b"\xd8", b"\xd9"}:
                        continue

                    size_bytes = fh.read(2)
                    if len(size_bytes) != 2:
                        return None
                    segment_size = struct.unpack(">H", size_bytes)[0]
                    if segment_size < 2:
                        return None

                    sof_markers = {
                        b"\xc0", b"\xc1", b"\xc2", b"\xc3",
                        b"\xc5", b"\xc6", b"\xc7",
                        b"\xc9", b"\xca", b"\xcb",
                        b"\xcd", b"\xce", b"\xcf",
                    }
                    if marker_type in sof_markers:
                        payload = fh.read(segment_size - 2)
                        if len(payload) < 5:
                            return None
                        height, width = struct.unpack(">HH", payload[1:5])
                        return width, height

                    fh.seek(segment_size - 2, 1)
    except OSError:
        return None

    return None
