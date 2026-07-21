from __future__ import annotations

import hashlib
import zipfile
from dataclasses import dataclass
from io import BytesIO
from typing import Any

import pandas as pd
from PIL import Image


PRESENTATION_REQUIRED_TOKENS = {
    "{{COMPANY}}",
    "{{IT_SCORE}}",
    "{{SUMMARY_1}}",
    "{{RISK_1_TITLE}}",
    "{{RISK_6_TITLE}}",
    "{{STRENGTH_1}}",
    "{{THREAT_1_LABEL}}",
    "{{THREAT_1_VALUE}}",
    "{{THREAT_6_LABEL}}",
    "{{THREAT_6_VALUE}}",
    "{{COVERAGE_AVERAGE}}",
    "{{COVERAGE_INSIGHT}}",
    "{{REG_TITLE}}",
    "{{REG_APPLICABILITY}}",
    "{{REG_EXPECTATIONS}}",
    "{{REG_IMPLEMENTATION}}",
    "{{REG_ANCHORS}}",
    "{{FRAMEWORKS}}",
    "{{REC_1_TITLE}}",
    "{{REC_1_ACTION}}",
    "{{REC_1_EVIDENCE}}",
    "{{REC_1_LEGAL}}",
    "{{REC_1_METRIC}}",
    "{{REC_1_SOLUTION}}",
    "{{REC_1_VENDORS}}",
    "{{REC_8_TITLE}}",
    "{{REC_8_ACTION}}",
    "{{ROADMAP_1_1}}",
    "{{ROADMAP_1_1_RESULT}}",
    "{{ROADMAP_1_2_RESULT}}",
    "{{DECISION_1}}",
}
PORTFOLIO_REQUIRED_COLUMNS = {
    "Vendor",
    "Distributor KZ",
    "Distributor Status",
}


@dataclass(frozen=True)
class AssetValidation:
    ok: bool
    message: str
    content_type: str
    details: dict[str, Any]


def sha256_hex(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def validate_logo(data: bytes, filename: str) -> AssetValidation:
    if not data or len(data) > 5 * 1024 * 1024:
        return AssetValidation(False, "Логотип должен быть не больше 5 МБ.", "", {})
    try:
        with Image.open(BytesIO(data)) as image:
            image.verify()
        with Image.open(BytesIO(data)) as image:
            width, height = image.size
            image_format = str(image.format or "").upper()
    except Exception:
        return AssetValidation(False, "Файл не является корректным PNG или JPEG.", "", {})
    if image_format not in {"PNG", "JPEG"}:
        return AssetValidation(False, "Поддерживаются только PNG и JPEG.", "", {})
    if width < 200 or height < 80:
        return AssetValidation(False, "Рекомендуемый минимальный размер логотипа: 200 x 80 px.", "", {})
    content_type = "image/png" if image_format == "PNG" else "image/jpeg"
    return AssetValidation(
        True,
        f"Логотип проверен: {width} x {height} px.",
        content_type,
        {"width": width, "height": height, "format": image_format, "filename": filename},
    )


def validate_presentation_template(data: bytes, filename: str) -> AssetValidation:
    if not data or len(data) > 25 * 1024 * 1024:
        return AssetValidation(False, "Шаблон презентации должен быть не больше 25 МБ.", "", {})
    try:
        with zipfile.ZipFile(BytesIO(data), "r") as archive:
            bad_file = archive.testzip()
            if bad_file:
                return AssetValidation(False, f"Повреждён файл внутри PPTX: {bad_file}.", "", {})
            slide_names = [
                name for name in archive.namelist()
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            ]
            slide_xml = "\n".join(
                archive.read(name).decode("utf-8", errors="replace") for name in slide_names
            )
    except (OSError, zipfile.BadZipFile):
        return AssetValidation(False, "Файл не является корректной презентацией PPTX.", "", {})
    missing = sorted(token for token in PRESENTATION_REQUIRED_TOKENS if token not in slide_xml)
    if missing:
        preview = ", ".join(missing[:5])
        return AssetValidation(
            False,
            f"Шаблон несовместим: отсутствуют служебные поля {preview}.",
            "",
            {"missing_tokens": missing},
        )
    return AssetValidation(
        True,
        f"Шаблон проверен: слайдов {len(slide_names)}, служебные поля на месте.",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        {"slides": len(slide_names), "filename": filename},
    )


def validate_vendor_matrix(data: bytes, filename: str) -> AssetValidation:
    if not data or len(data) > 15 * 1024 * 1024:
        return AssetValidation(False, "Файл портфеля должен быть не больше 15 МБ.", "", {})
    try:
        frame = pd.read_excel(BytesIO(data))
    except Exception as exc:
        return AssetValidation(False, f"Excel-файл не читается: {exc}", "", {})
    missing = sorted(PORTFOLIO_REQUIRED_COLUMNS.difference(frame.columns))
    if missing:
        return AssetValidation(
            False,
            f"Не хватает обязательных столбцов: {', '.join(missing)}.",
            "",
            {"missing_columns": missing},
        )
    category_columns = [
        column for column in frame.columns
        if column not in PORTFOLIO_REQUIRED_COLUMNS
        and column not in {"Distributor Source", "Notes"}
    ]
    if frame.empty or not category_columns:
        return AssetValidation(False, "В портфеле нет производителей или категорий решений.", "", {})
    vendors = frame["Vendor"].dropna().astype(str).str.strip()
    if vendors.empty or vendors.duplicated().any():
        return AssetValidation(False, "Поле Vendor пустое или содержит дубли производителей.", "", {})
    return AssetValidation(
        True,
        f"Портфель проверен: производителей {len(vendors)}, категорий {len(category_columns)}.",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        {"vendors": len(vendors), "categories": len(category_columns), "filename": filename},
    )
