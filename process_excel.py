#!/usr/bin/env python3
"""Build browser-safe factory dashboard shards from a large ERP source.

XLSX workbooks and Google Visualization CSV sources are read in streaming mode.
Processed rows are written to monthly JSONL files first, then packed as
deterministic gzip JSON payloads. Factory identity and source settings live in
factory-config.json; the calculation engine stays identical for every plant.
"""

from __future__ import annotations

import argparse
import csv
import gzip
import hashlib
import json
import math
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
from contextlib import contextmanager
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, BinaryIO, Iterable
from urllib.parse import quote
from urllib.request import Request, urlopen
from xml.etree import ElementTree as ET
from zoneinfo import ZoneInfo

FILE_ID = ""
FILE_NAME = "P3_Tong_Hop_LTT.xlsx"
SHEET_NAME = "P3.Tổng hợp lệnh thao tác"
SOURCE_RANGE = "A9:CT"
PLANT_CODE = ""
PLANT_NAME = ""
SOURCE_TYPE = "drive_xlsx"
SEGMENT_MODE = "auto_value_stream"
SCHEMA_VERSION = "gsp-factory-static-shards-v1"
MAX_COLUMNS = 98  # A:CT, zero-based indices 0..97.

GS1_AF_SEGMENT_CATALOG = {
    "HOC": "01_Nhóm hàng Hộp cứng",
    "HOT": "02_Nhóm hàng Hộp thường (hộp mềm)",
    "HBD": "03_Nhóm hàng Hộp bồi duplex",
    "HBL": "04_Nhóm hàng Hộp bồi label",
    "FLC": "05_Nhóm hàng Hộp Flexo carton",
    "FLP": "06_Nhóm hàng Hộp Flexo proces",
    "FPK": "07_Nhóm hàng PK phôi carton",
    "SHD": "08_Nhóm hàng Sách hướng dẫn",
    "PLL": "09_Nhóm hàng Pallet",
    "PHOI": "10_Nhóm hàng Phôi sóng",
    "GCI": "11_Nhóm hàng Gia công in",
    "KHA": "12_Nhóm hàng Khay giấy",
    "TUI": "13_Nhóm hàng Túi giấy",
    "NVLC": "14_Nhóm Nguyên vật liệu chính",
    "NVLP": "15_Nhóm Nguyên vật liệu phụ",
    "PTRO": "16_Nhóm CCDC, VTB, VIT, VPP",
    "KHC": "17_Nhóm hàng khác",
    "LE": "18_Nhóm Lề, phế",
    "TCKT": "19_Nhóm hàng thương mại",
}
AF_MISSING_LABEL = "00_Chưa khai báo dòng hàng mẹ"
AF_CONFLICT_LABEL = "98_Xung đột AF theo LTT/phiếu"
AF_UNMAPPED_LABEL = "99_AF chưa ánh xạ"

PROCESSED_COLS = [
    "date",
    "month",
    "week",
    "segment",
    "machine",
    "operator",
    "ltt",
    "stat",
    "mat_code",
    "mat_name",
    "process",
    "process_code",
    "line",
    "line_name",
    "unit",
    "shift",
    "qty",
    "ok",
    "ng",
    "ng_rate",
    "allow_qty",
    "allow_rate",
    "over_qty",
    "over_pos",
    "over_rate",
    "downtime",
    "prep_h",
    "nvl_h",
    "machine_h",
    "file_h",
    "reason",
    "capacity",
    "oee",
    "A",
    "P",
    "Q",
    "actual_prod",
    "achv_tech",
    "confidence",
    "flag_count",
    "flags",
    "rag",
    "ltt_req_qty",
    "ltt_allow_ng_qty",
    "converted_qty",
    "setup_std_h",
    "run_std_h",
    "total_std_h",
    "k_lot",
    "job_size_class",
    "oee_weight",
    "oee_lot_adjusted",
    "oee_weight_valid",
    "capa_time_std",
    "capa_actual",
    "capa_rate_std",
    "capa_rate_tech",
]


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_af(value: Any) -> str:
    return clean_text(value).upper()


def finite(value: float) -> float:
    return value if math.isfinite(value) else 0.0


def to_num(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return finite(float(value))
    text = str(value).strip()
    if not text:
        return 0.0
    is_percent = "%" in text
    text = text.replace("%", "").replace(" ", "").replace("\u00a0", "")
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    has_comma = "," in text
    has_dot = "." in text
    if has_comma and has_dot:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif has_comma:
        parts = text.split(",")
        if len(parts) > 2:
            text = text.replace(",", "")
        elif len(parts) == 2 and len(parts[1]) == 3 and parts[0] != "0":
            text = "".join(parts)
        else:
            text = text.replace(",", ".")
    elif has_dot:
        parts = text.split(".")
        if len(parts) > 2:
            text = text.replace(".", "")
        elif len(parts) == 2 and len(parts[1]) == 3 and parts[0] != "0":
            text = "".join(parts)
    text = re.sub(r"[^0-9.\-]", "", text)
    try:
        result = float(text)
    except (TypeError, ValueError):
        return 0.0
    if negative:
        result = -result
    if is_percent:
        result /= 100.0
    return finite(result)


def excel_serial_datetime(value: float) -> datetime | None:
    if not math.isfinite(value) or value <= 0:
        return None
    try:
        return datetime(1899, 12, 30, tzinfo=timezone.utc) + timedelta(days=value)
    except (OverflowError, ValueError):
        return None


def parse_datetime_value(value: Any) -> datetime | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.replace(tzinfo=value.tzinfo or timezone.utc).astimezone(timezone.utc)
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day, tzinfo=timezone.utc)
    if isinstance(value, (int, float)):
        number = float(value)
        if 20000 < number < 80000:
            return excel_serial_datetime(number)
    text = str(value).strip()
    if re.fullmatch(r"\d+(?:\.\d+)?", text):
        number = float(text)
        if 20000 < number < 80000:
            return excel_serial_datetime(number)
    match = re.search(
        r"(?:new\s+)?Date\s*\(\s*(\d{4})\s*,\s*(\d{1,2})\s*,\s*(\d{1,2})"
        r"(?:\s*,\s*(\d{1,2}))?(?:\s*,\s*(\d{1,2}))?(?:\s*,\s*(\d{1,2}))?",
        text,
        re.I,
    )
    if match:
        year, month0, day, hour, minute, second = [
            int(item or 0) for item in match.groups()
        ]
        try:
            return datetime(
                year, month0 + 1, day, hour, minute, second, tzinfo=timezone.utc
            )
        except ValueError:
            return None
    candidates = (
        ("%Y-%m-%dT%H:%M:%S", text[:19]),
        ("%Y-%m-%d %H:%M:%S", text[:19]),
        ("%Y-%m-%d", text[:10]),
        ("%d/%m/%Y %H:%M:%S", text[:19]),
        ("%d/%m/%Y", text[:10]),
        ("%d/%m/%y", text[:8]),
        ("%d-%m-%Y", text[:10]),
    )
    for fmt, candidate in candidates:
        try:
            return datetime.strptime(candidate, fmt).replace(tzinfo=timezone.utc)
        except ValueError:
            pass
    return None


def iso_date(value: datetime | None) -> str:
    return value.date().isoformat() if value else ""


def month_key(value: datetime | None) -> str:
    return value.strftime("%Y-%m") if value else "Không tháng"


def week_key(value: datetime | None) -> str:
    if not value:
        return "Không tuần"
    iso = value.isocalendar()
    return f"{iso.year}-W{iso.week:02d}"


def raw_at(raw: list[Any], index: int) -> Any:
    return raw[index] if index < len(raw) else ""


def classify_segment(
    raw: list[Any],
    other_label: str,
    segment_mode: str,
    af_conflict_ltts: set[str] | None = None,
    af_conflict_stats: set[str] | None = None,
) -> str:
    line_mother = normalize_af(raw_at(raw, 31))
    ltt = clean_text(raw_at(raw, 2))
    stat = clean_text(raw_at(raw, 45))
    if segment_mode in {"gs1_parent_line_af", "gs1_value_stream"}:
        if (
            (ltt and af_conflict_ltts and ltt in af_conflict_ltts)
            or (stat and af_conflict_stats and stat in af_conflict_stats)
        ):
            return AF_CONFLICT_LABEL
        if not line_mother:
            return AF_MISSING_LABEL
        return GS1_AF_SEGMENT_CATALOG.get(line_mother, AF_UNMAPPED_LABEL)
    line = clean_text(raw_at(raw, 32)).upper()
    line_name = clean_text(raw_at(raw, 33)).upper()
    process_code = clean_text(raw_at(raw, 28)).upper()
    process_name = clean_text(raw_at(raw, 29)).upper()
    machine = clean_text(raw_at(raw, 47) or raw_at(raw, 27)).upper()
    material = " ".join(
        clean_text(raw_at(raw, index)) for index in (5, 6, 17)
    ).upper()
    text = " ".join(
        [
            line_mother,
            line,
            line_name,
            process_code,
            process_name,
            machine,
            material,
        ]
    )
    if segment_mode == "auto_value_stream":
        display_line_name = clean_text(raw_at(raw, 33))
        display_line = clean_text(raw_at(raw, 32))
        display_mother = clean_text(raw_at(raw, 31))
        if display_line_name:
            return re.sub(
                r"^\s*Dòng\s+hàng\s*",
                "",
                display_line_name,
                flags=re.I,
            ) or display_line_name
        return display_line or display_mother or other_label
    if (
        "KHAY" in text
        or line_mother == "KHAY"
        or line.startswith("KH")
        or "KHAY" in process_code
        or "KHAY" in process_name
        or "KHAY" in machine
    ):
        return "Khay giấy"
    flexo_process = (
        "INFLEXOCUON" in process_code
        or "FLEXO" in process_name
        or " IN CUỘN " in text
        or "IN CUON" in text
    )
    flexo_line = (
        line_mother == "FLC"
        or line.startswith("FS")
        or "FLEXO" in line_name
        or "FLEXO SÓNG" in material
        or "FLEXO SONG" in material
    )
    if flexo_process and flexo_line:
        return "Flexo cuộn 6 màu"
    if (
        line_mother == "TUI"
        or line.startswith("T")
        or "TUI" in process_code
        or "TÚI" in process_name
        or "TUI" in process_name
        or machine.startswith(("QD", "QL", "QX"))
    ):
        return "Túi giấy"
    if "INFLEXOCUON" in process_code or "YRK" in machine:
        return "Flexo cuộn 6 màu"
    return other_label


def rag_for(obj: dict[str, Any]) -> str:
    threshold = max(0.08, to_num(obj["allow_rate"]) * 2 + 0.02)
    if (
        to_num(obj["over_pos"]) > 1000
        or to_num(obj["downtime"]) > 4
        or to_num(obj["ng_rate"]) > threshold
        or obj["confidence"] == "Đỏ"
    ):
        return "Đỏ"
    if (
        to_num(obj["over_pos"]) > 0
        or to_num(obj["downtime"]) > 0
        or obj["confidence"] == "Vàng"
    ):
        return "Vàng"
    return "Xanh"


def process_raw_row(
    raw: list[Any],
    other_label: str,
    segment_mode: str,
    af_conflict_ltts: set[str] | None = None,
    af_conflict_stats: set[str] | None = None,
) -> list[Any]:
    while len(raw) < 100:
        raw.append("")
    stat_date = parse_datetime_value(raw[43])  # AR
    ltt_req_qty = to_num(raw[18])  # S
    ltt_allow_ng_qty = to_num(raw[20])  # U
    ltt_plan_h = to_num(raw[25])  # Z
    qty = to_num(raw[58])  # BG
    ok_qty = to_num(raw[59])  # BH
    ng_raw = to_num(raw[60])  # BI
    ng_qty = ng_raw if ng_raw != 0 else max(qty - ok_qty, 0)
    ng_rate_raw = to_num(raw[61])
    ng_rate = ng_rate_raw if ng_rate_raw != 0 else (ng_qty / qty if qty else 0)
    allow_qty = to_num(raw[85])  # CH
    allow_rate_raw = to_num(raw[86])  # CI
    allow_rate = (
        allow_rate_raw if allow_rate_raw != 0 else (allow_qty / qty if qty else 0)
    )
    over_raw = to_num(raw[93])  # CP
    over_qty = ng_qty - allow_qty if raw[93] in (None, "") else over_raw
    over_pos = max(over_qty, 0)
    over_rate_raw = to_num(raw[92])  # CO
    over_rate = (
        over_rate_raw if over_rate_raw != 0 else (over_qty / qty if qty else 0)
    )
    prep_h = to_num(raw[73]) / 60
    nvl_h = to_num(raw[74]) / 60
    machine_h = to_num(raw[75]) / 60
    file_h = to_num(raw[76]) / 60
    downtime_raw = to_num(raw[77])
    downtime = (
        downtime_raw
        if downtime_raw != 0
        else prep_h + nvl_h + machine_h + file_h
    )
    capacity = to_num(raw[79])  # CB
    converted_qty = to_num(raw[81])  # CD
    setup_std_h = (to_num(raw[82]) + to_num(raw[83])) / 60
    run_std_h = converted_qty / capacity if capacity > 0 and converted_qty > 0 else 0
    total_std_source = to_num(raw[87])  # CJ
    total_std_h = total_std_source if total_std_source > 0 else setup_std_h + run_std_h
    k_lot_raw = run_std_h / total_std_h if total_std_h > 0 and run_std_h > 0 else 0
    k_lot = max(0, min(1, k_lot_raw))
    if k_lot >= 0.75:
        job_size_class = "Job dài"
    elif k_lot >= 0.45:
        job_size_class = "Job trung bình"
    elif k_lot > 0:
        job_size_class = "Job ngắn"
    else:
        job_size_class = "Thiếu dữ liệu"
    availability = to_num(raw[88])
    performance = to_num(raw[89])
    quality = to_num(raw[90])
    oee_raw = to_num(raw[91])
    oee = (
        oee_raw
        if oee_raw != 0
        else (
            availability * performance * quality
            if availability and performance and quality
            else 0
        )
    )
    oee_weight_valid = 0 < oee <= 1.2 and total_std_h > 0
    oee_weight = oee * total_std_h if oee_weight_valid else 0
    oee_lot_adjusted = min(oee / k_lot, 1) if oee > 0 and k_lot > 0 else 0
    capa_time_source = to_num(raw[94])  # CQ
    if capa_time_source != 0:
        capa_time_std = capa_time_source
    elif ltt_req_qty > 0 and ltt_plan_h > 0:
        capa_time_std = ltt_req_qty / ltt_plan_h
    else:
        capa_time_std = capacity
    capa_actual = to_num(raw[95])  # CR
    capa_rate_source = to_num(raw[96])  # CS
    if capa_rate_source != 0:
        capa_rate_std = capa_rate_source
    elif capa_actual > 0 and capa_time_std > 0:
        capa_rate_std = capa_actual / capa_time_std
    elif capa_actual > 0 and capacity > 0:
        capa_rate_std = capa_actual / capacity
    else:
        capa_rate_std = 0
    capa_rate_tech = to_num(raw[97])  # CT

    start = parse_datetime_value(raw[55])
    finish = parse_datetime_value(raw[56])
    flags: list[str] = []
    if segment_mode in {"gs1_parent_line_af", "gs1_value_stream"}:
        line_mother = normalize_af(raw[31])
        ltt = clean_text(raw[2])
        stat = clean_text(raw[45])
        if not line_mother:
            flags.append("AF trống")
        elif line_mother not in GS1_AF_SEGMENT_CATALOG:
            flags.append(f"AF ngoài danh mục: {line_mother}")
        if ltt and af_conflict_ltts and ltt in af_conflict_ltts:
            flags.append("Xung đột AF theo LTT")
        if stat and af_conflict_stats and stat in af_conflict_stats:
            flags.append("Xung đột AF theo phiếu thống kê")
    if (
        stat_date
        and stat_date.date()
        > datetime.now(ZoneInfo("Asia/Bangkok")).date() + timedelta(days=1)
    ):
        flags.append("Ngày TK AR ở tương lai")
    if downtime > 0 and not clean_text(raw[78]):
        flags.append("Downtime thiếu nguyên nhân")
    if start and finish and finish < start:
        flags.append("Thiếu/sai thời gian thống kê")
    if start and not 2025 <= start.year <= 2027:
        flags.append("TG bắt đầu TK năm bất thường")
    if finish and not 2025 <= finish.year <= 2027:
        flags.append("TG kết thúc TK năm bất thường")
    if (
        availability < 0
        or performance < 0
        or quality < 0
        or oee < 0
        or availability > 1.2
        or performance > 1.2
        or quality > 1.05
        or oee > 1.2
        or (0 < oee < 0.05)
    ):
        flags.append("OEE/A/P/Q bất thường")
    confidence = "Xanh"
    if any(
        "Downtime thiếu" in item
        or "năm bất thường" in item
        or "Ngày TK AR" in item
        or "AF trống" in item
        or "AF ngoài danh mục" in item
        or "Xung đột AF" in item
        for item in flags
    ):
        confidence = "Đỏ"
    elif flags:
        confidence = "Vàng"

    obj: dict[str, Any] = {
        "date": iso_date(stat_date),
        "month": month_key(stat_date),
        "week": week_key(stat_date),
        "segment": classify_segment(
            raw,
            other_label,
            segment_mode,
            af_conflict_ltts,
            af_conflict_stats,
        ),
        "machine": clean_text(raw[47] or raw[27]) or "Chưa có",
        "operator": clean_text(raw[49]) or "Chưa có",
        "ltt": clean_text(raw[2]),
        "stat": clean_text(raw[45]),
        "mat_code": clean_text(raw[5]),
        "mat_name": clean_text(raw[6]),
        "process": clean_text(raw[29]),
        "process_code": clean_text(raw[28]),
        "line": clean_text(raw[32]),
        "line_name": clean_text(raw[33]),
        "unit": clean_text(raw[62] or raw[7]),
        "shift": clean_text(raw[54] or raw[22]),
        "qty": qty,
        "ok": ok_qty,
        "ng": ng_qty,
        "ng_rate": ng_rate,
        "allow_qty": allow_qty,
        "allow_rate": allow_rate,
        "over_qty": over_qty,
        "over_pos": over_pos,
        "over_rate": over_rate,
        "downtime": downtime,
        "prep_h": prep_h,
        "nvl_h": nvl_h,
        "machine_h": machine_h,
        "file_h": file_h,
        "reason": clean_text(raw[78]) or "Chưa có",
        "capacity": capacity,
        "oee": oee,
        "A": availability,
        "P": performance,
        "Q": quality,
        "actual_prod": capa_actual,
        "achv_tech": capa_rate_tech,
        "confidence": confidence,
        "flag_count": len(flags),
        "flags": "; ".join(flags),
        "ltt_req_qty": ltt_req_qty,
        "ltt_allow_ng_qty": ltt_allow_ng_qty,
        "converted_qty": converted_qty,
        "setup_std_h": setup_std_h,
        "run_std_h": run_std_h,
        "total_std_h": total_std_h,
        "k_lot": k_lot,
        "job_size_class": job_size_class,
        "oee_weight": oee_weight,
        "oee_lot_adjusted": oee_lot_adjusted,
        "oee_weight_valid": oee_weight_valid,
        "capa_time_std": capa_time_std,
        "capa_actual": capa_actual,
        "capa_rate_std": capa_rate_std,
        "capa_rate_tech": capa_rate_tech,
    }
    obj["rag"] = rag_for(obj)
    return [obj[column] for column in PROCESSED_COLS]


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def column_index(reference: str) -> int:
    match = re.match(r"([A-Z]+)", reference or "")
    if not match:
        return -1
    result = 0
    for char in match.group(1):
        result = result * 26 + ord(char) - 64
    return result - 1


def shared_text(element: ET.Element) -> str:
    return "".join(
        child.text or "" for child in element.iter() if local_name(child.tag) == "t"
    )


def read_shared_strings(workbook: zipfile.ZipFile) -> list[str]:
    path = "xl/sharedStrings.xml"
    if path not in workbook.namelist():
        return []
    strings: list[str] = []
    with workbook.open(path) as stream:
        for event, element in ET.iterparse(stream, events=("end",)):
            if local_name(element.tag) == "si":
                strings.append(shared_text(element))
                element.clear()
    return strings


def resolve_sheet_path(workbook: zipfile.ZipFile, wanted_name: str) -> str:
    namespaces = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "pkg": "http://schemas.openxmlformats.org/package/2006/relationships",
    }
    workbook_xml = ET.fromstring(workbook.read("xl/workbook.xml"))
    relation_id = ""
    for sheet in workbook_xml.findall(".//main:sheet", namespaces):
        if sheet.attrib.get("name") == wanted_name:
            relation_id = sheet.attrib.get(f"{{{namespaces['rel']}}}id", "")
            break
    if not relation_id:
        available = [
            sheet.attrib.get("name", "")
            for sheet in workbook_xml.findall(".//main:sheet", namespaces)
        ]
        raise RuntimeError(
            f"Không tìm thấy sheet {wanted_name!r}. Có các sheet: {available}"
        )
    rels_xml = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
    for relation in rels_xml.findall("pkg:Relationship", namespaces):
        if relation.attrib.get("Id") == relation_id:
            target = relation.attrib.get("Target", "")
            if target.startswith("/"):
                return target.lstrip("/")
            return str(Path("xl") / target).replace("\\", "/")
    raise RuntimeError(f"Không tìm thấy quan hệ XML cho sheet {wanted_name!r}.")


def cell_value(cell: ET.Element, shared_strings: list[str]) -> Any:
    cell_type = cell.attrib.get("t", "")
    if cell_type == "inlineStr":
        return shared_text(cell)
    value: str | None = None
    for child in cell:
        if local_name(child.tag) == "v":
            value = child.text or ""
            break
    if value is None:
        return ""
    if cell_type == "s":
        try:
            return shared_strings[int(value)]
        except (ValueError, IndexError):
            return ""
    if cell_type == "b":
        return value == "1"
    if cell_type in ("", "n"):
        try:
            number = float(value)
            return int(number) if number.is_integer() else number
        except ValueError:
            return value
    return value


def iter_sheet_rows(
    workbook: zipfile.ZipFile,
    sheet_path: str,
    shared_strings: list[str],
    max_columns: int,
) -> Iterable[tuple[int, list[Any]]]:
    with workbook.open(sheet_path) as stream:
        for event, element in ET.iterparse(stream, events=("end",)):
            if local_name(element.tag) != "row":
                continue
            row_number = int(element.attrib.get("r", "0") or 0)
            row: list[Any] = [""] * max_columns
            for cell in element:
                if local_name(cell.tag) != "c":
                    continue
                index = column_index(cell.attrib.get("r", ""))
                if 0 <= index < max_columns:
                    row[index] = cell_value(cell, shared_strings)
            yield row_number, row
            element.clear()


@contextmanager
def source_rows(
    source_path: Path,
    source_type: str,
    sheet_name: str,
) -> Iterable[Iterable[tuple[int, list[Any]]]]:
    if source_type == "google_sheet_csv":
        with source_path.open(
            "r", encoding="utf-8-sig", newline=""
        ) as stream:
            reader = csv.reader(stream)

            def csv_rows() -> Iterable[tuple[int, list[Any]]]:
                for offset, raw in enumerate(reader):
                    row = list(raw[:MAX_COLUMNS])
                    if len(row) < MAX_COLUMNS:
                        row.extend([""] * (MAX_COLUMNS - len(row)))
                    yield 9 + offset, row

            print(f"Google Sheet CSV: {sheet_name} · {SOURCE_RANGE}", flush=True)
            yield csv_rows()
        return

    if source_type != "drive_xlsx":
        raise RuntimeError(f"Kiểu nguồn chưa hỗ trợ: {source_type!r}.")
    with zipfile.ZipFile(source_path) as workbook:
        shared_strings = read_shared_strings(workbook)
        print(f"Shared strings: {len(shared_strings):,}", flush=True)
        sheet_path = resolve_sheet_path(workbook, sheet_name)
        print(f"Sheet XML: {sheet_path}", flush=True)
        yield iter_sheet_rows(
            workbook, sheet_path, shared_strings, MAX_COLUMNS
        )


def collect_af_quality(
    source_path: Path,
    source_type: str,
    sheet_name: str,
    plant_code: str,
) -> dict[str, Any]:
    ltt_values: dict[str, set[str]] = defaultdict(set)
    stat_values: dict[str, set[str]] = defaultdict(set)
    code_rows: dict[str, int] = defaultdict(int)
    accepted_keys: list[tuple[str, str, str]] = []
    accepted_rows = 0
    with source_rows(source_path, source_type, sheet_name) as rows:
        header_ok = False
        for row_number, raw in rows:
            if row_number < 9:
                continue
            if row_number == 9:
                header = " | ".join(clean_text(item) for item in raw)
                header_ok = (
                    "Lệnh công đoạn" in header and "Số thống kê" in header
                )
                if not header_ok:
                    raise RuntimeError(
                        "Không nhận diện được header dòng 9 khi kiểm tra AF."
                    )
                continue
            plant = clean_text(raw_at(raw, 4)) or "(trống)"
            if plant != plant_code:
                continue
            if not (
                clean_text(raw_at(raw, 2))
                or clean_text(raw_at(raw, 45))
                or clean_text(raw_at(raw, 5))
            ):
                continue
            accepted_rows += 1
            af = normalize_af(raw_at(raw, 31))
            code_rows[af or "__MISSING__"] += 1
            ltt = clean_text(raw_at(raw, 2))
            stat = clean_text(raw_at(raw, 45))
            accepted_keys.append((af, ltt, stat))
            if ltt:
                ltt_values[ltt].add(af or "__MISSING__")
            if stat:
                stat_values[stat].add(af or "__MISSING__")
        if not header_ok:
            raise RuntimeError("Không đọc được dòng header 9 khi kiểm tra AF.")

    conflict_ltts = {
        key for key, values in ltt_values.items() if len(values) > 1
    }
    conflict_stats = {
        key for key, values in stat_values.items() if len(values) > 1
    }
    raw_missing_rows = code_rows.get("__MISSING__", 0)
    unmapped_codes = {
        code: count
        for code, count in code_rows.items()
        if code != "__MISSING__" and code not in GS1_AF_SEGMENT_CATALOG
    }
    valid_rows = 0
    missing_rows = 0
    unmapped_rows = 0
    conflict_rows = 0
    for af, ltt, stat in accepted_keys:
        if (ltt and ltt in conflict_ltts) or (stat and stat in conflict_stats):
            conflict_rows += 1
        elif not af:
            missing_rows += 1
        elif af not in GS1_AF_SEGMENT_CATALOG:
            unmapped_rows += 1
        else:
            valid_rows += 1
    return {
        "accepted_rows": accepted_rows,
        "valid_rows": valid_rows,
        "valid_rate": valid_rows / accepted_rows if accepted_rows else 0,
        "missing_rows": missing_rows,
        "raw_missing_rows": raw_missing_rows,
        "unmapped_rows": unmapped_rows,
        "raw_unmapped_rows": sum(unmapped_codes.values()),
        "conflict_rows": conflict_rows,
        "unmapped_codes": dict(
            sorted(unmapped_codes.items(), key=lambda item: (-item[1], item[0]))
        ),
        "conflict_ltts": conflict_ltts,
        "conflict_stats": conflict_stats,
        "code_rows": dict(
            sorted(
                (
                    ("" if code == "__MISSING__" else code, count)
                    for code, count in code_rows.items()
                ),
                key=lambda item: item[0],
            )
        ),
    }


@dataclass
class ShardStats:
    key: str
    file_name: str
    rows: int = 0
    ltts: set[str] = field(default_factory=set)
    stats: set[str] = field(default_factory=set)
    machines: set[str] = field(default_factory=set)
    operators: set[str] = field(default_factory=set)
    date_min: str = ""
    date_max: str = ""
    ltt_required_qty: float = 0
    ltt_allow_ng_qty: float = 0
    actual_qty: float = 0
    ng: float = 0
    allow_qty: float = 0
    over_pos: float = 0
    downtime: float = 0
    oee_sum: float = 0
    oee_count: int = 0
    red: int = 0
    yellow: int = 0

    def add(self, processed: list[Any]) -> None:
        row = dict(zip(PROCESSED_COLS, processed))
        self.rows += 1
        ltt = str(row["ltt"] or "")
        stat = str(row["stat"] or "")
        if ltt and ltt not in self.ltts:
            self.ltts.add(ltt)
            self.ltt_required_qty += to_num(row["ltt_req_qty"])
            self.ltt_allow_ng_qty += to_num(row["ltt_allow_ng_qty"])
        if stat and stat not in self.stats:
            self.stats.add(stat)
            self.actual_qty += to_num(row["qty"])
            self.ng += to_num(row["ng"])
            self.allow_qty += to_num(row["allow_qty"])
            self.over_pos += max(to_num(row["ng"]) - to_num(row["allow_qty"]), 0)
        if row["machine"]:
            self.machines.add(str(row["machine"]))
        if row["operator"]:
            self.operators.add(str(row["operator"]))
        self.downtime += to_num(row["downtime"])
        self.oee_sum += to_num(row["oee"])
        self.oee_count += 1
        if row["confidence"] == "Đỏ":
            self.red += 1
        if row["confidence"] == "Vàng":
            self.yellow += 1
        day = str(row["date"] or "")
        if day:
            if not self.date_min or day < self.date_min:
                self.date_min = day
            if not self.date_max or day > self.date_max:
                self.date_max = day

    def meta(self, source_label: str, generated: str) -> dict[str, Any]:
        return {
            "source": source_label,
            "generated": generated,
            "period": self.key,
            "rows": self.rows,
            "ltt": len(self.ltts),
            "stats": len(self.stats),
            "machines": len(self.machines),
            "operators": len(self.operators),
            "date_min": self.date_min,
            "date_max": self.date_max,
        }

    def exec_summary(self) -> dict[str, Any]:
        defect_rate = self.ng / self.actual_qty if self.actual_qty else 0
        allow_rate = (
            self.ltt_allow_ng_qty / self.ltt_required_qty
            if self.ltt_required_qty
            else 0
        )
        completion_rate = (
            self.actual_qty / self.ltt_required_qty if self.ltt_required_qty else 0
        )
        return {
            "month": self.key,
            "rows": self.rows,
            "ltt": len(self.ltts),
            "stats": len(self.stats),
            "machines": len(self.machines),
            "ops": len(self.operators),
            "ltt_required_qty": self.ltt_required_qty,
            "ltt_allow_ng_qty": self.ltt_allow_ng_qty,
            "actual_qty": self.actual_qty,
            "ng": self.ng,
            "allow_qty": self.allow_qty,
            "over_pos": self.over_pos,
            "dt": self.downtime,
            "oee": self.oee_sum / self.oee_count if self.oee_count else 0,
            "defect_rate": defect_rate,
            "allow_rate": allow_rate,
            "completion_rate": completion_rate,
            "red": self.red,
            "yellow": self.yellow,
        }


def compact_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), allow_nan=False)


def write_gzip_json(
    output_path: Path, prefix: str, jsonl_path: Path | None, suffix: str
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("wb") as raw:
        with gzip.GzipFile(fileobj=raw, mode="wb", compresslevel=9, mtime=0) as zipped:
            zipped.write(prefix.encode("utf-8"))
            if jsonl_path and jsonl_path.exists():
                first = True
                with jsonl_path.open("rb") as source:
                    for line in source:
                        line = line.strip()
                        if not line:
                            continue
                        if not first:
                            zipped.write(b",")
                        zipped.write(line)
                        first = False
            zipped.write(suffix.encode("utf-8"))


def format_source_signature(headers: dict[str, str], file_id: str) -> str:
    parts = [
        SCHEMA_VERSION,
        file_id,
        headers.get("content-length", ""),
        headers.get("last-modified", ""),
        headers.get("etag", ""),
    ]
    return hashlib.sha256("|".join(parts).encode("utf-8")).hexdigest()[:24]


def source_url(source_type: str, file_id: str, sheet_name: str) -> str:
    if source_type == "google_sheet_csv":
        return (
            "https://docs.google.com/spreadsheets/d/"
            f"{file_id}/gviz/tq?tqx=out:csv&headers=1"
            f"&sheet={quote(sheet_name)}&range={quote(SOURCE_RANGE)}"
        )
    return (
        "https://drive.usercontent.google.com/download"
        f"?id={file_id}&export=download&confirm=t"
    )


def probe_source(
    url: str, file_id: str, source_type: str, plant_code: str
) -> dict[str, Any]:
    if source_type == "google_sheet_csv":
        window = int(time.time() // 900)
        signature = hashlib.sha256(
            f"{SCHEMA_VERSION}|{file_id}|{window}".encode("utf-8")
        ).hexdigest()[:24]
        return {
            "signature": signature,
            "size": 0,
            "modified": f"15-minute window {window}",
            "etag": "",
            "resolved_url": url,
        }
    request = Request(
        url,
        method="HEAD",
        headers={"User-Agent": f"{plant_code}-Factory-Dashboard/1.0"},
    )
    with urlopen(request, timeout=60) as response:
        headers = {key.lower(): value for key, value in response.headers.items()}
        resolved_url = response.geturl()
    size = int(headers.get("content-length", "0") or 0)
    if size <= 0:
        raise RuntimeError("Nguồn Excel không trả Content-Length hợp lệ.")
    return {
        "signature": format_source_signature(headers, file_id),
        "size": size,
        "modified": headers.get("last-modified", ""),
        "etag": headers.get("etag", ""),
        "resolved_url": resolved_url,
    }


def download_source(
    url: str,
    target: Path,
    expected_size: int = 0,
    source_type: str = "drive_xlsx",
    plant_code: str = "GSP",
) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    request = Request(
        url, headers={"User-Agent": f"{plant_code}-Factory-Dashboard/1.0"}
    )
    with urlopen(request, timeout=300) as response:
        content_type = response.headers.get("content-type", "")
        with target.open("wb") as output:
            downloaded = 0
            last_report = 0
            while chunk := response.read(4 * 1024 * 1024):
                if not chunk:
                    continue
                output.write(chunk)
                downloaded += len(chunk)
                if downloaded - last_report >= 32 * 1024 * 1024:
                    print(
                        f"Đã tải {downloaded / 1024 / 1024:.1f} MB...",
                        flush=True,
                    )
                    last_report = downloaded
    if expected_size and target.stat().st_size != expected_size:
        raise RuntimeError(
            f"Dung lượng tải về {target.stat().st_size} khác metadata {expected_size}."
        )
    if source_type == "drive_xlsx" and target.read_bytes()[:2] != b"PK":
        raise RuntimeError(
            f"Nguồn tải về không phải XLSX/ZIP. Content-Type: {content_type}"
        )


def schedule_window() -> tuple[datetime, datetime]:
    now = datetime.now(ZoneInfo("Asia/Bangkok"))
    today = datetime(now.year, now.month, now.day, tzinfo=timezone.utc)
    monday = today - timedelta(days=today.isoweekday() - 1)
    return monday, monday + timedelta(days=14)


def schedule_variant_overlaps(raw: list[Any], start: datetime, end: datetime) -> bool:
    begin = parse_datetime_value(raw_at(raw, 23))
    finish = parse_datetime_value(raw_at(raw, 24))
    return bool(begin and finish and finish >= begin and finish >= start and begin < end)


def safe_shard_name(key: str) -> str:
    return "period_undated.json.gz" if key == "Không tháng" else f"period_{key.replace('-', '_')}.json.gz"


def month_label(key: str) -> str:
    if key == "Không tháng":
        return "Thiếu ngày thống kê"
    year, month = key.split("-")
    return f"Tháng {month}/{year}"


def build_data(
    workbook_path: Path,
    out_dir: Path,
    plant_code: str,
    plant_name: str,
    source_meta: dict[str, Any],
    source_type: str,
    segment_mode: str,
    sheet_name: str,
    file_name: str,
    file_id: str,
) -> dict[str, Any]:
    started = time.monotonic()
    generated_dt = datetime.now(ZoneInfo("Asia/Bangkok"))
    generated = generated_dt.strftime("%H:%M:%S %d/%m/%Y")
    generated_iso = generated_dt.isoformat(timespec="seconds")
    source_label = (
        f"{'Google Sheet' if source_type == 'google_sheet_csv' else 'Excel Drive'} "
        f"near-live · {file_name} · {plant_code}"
    )
    other_label = f"{plant_code} khác"
    af_quality: dict[str, Any] = {
        "accepted_rows": 0,
        "valid_rows": 0,
        "valid_rate": 0,
        "missing_rows": 0,
        "raw_missing_rows": 0,
        "unmapped_rows": 0,
        "raw_unmapped_rows": 0,
        "conflict_rows": 0,
        "unmapped_codes": {},
        "conflict_ltts": set(),
        "conflict_stats": set(),
        "code_rows": {},
    }
    if segment_mode in {"gs1_parent_line_af", "gs1_value_stream"}:
        print("Đang kiểm tra tính nhất quán cột AF theo LTT/phiếu…", flush=True)
        af_quality = collect_af_quality(
            workbook_path,
            source_type,
            sheet_name,
            plant_code,
        )
        print(
            "AF quality: "
            f"{af_quality['valid_rate']:.1%} hợp lệ · "
            f"{af_quality['missing_rows']:,} dòng trống · "
            f"{af_quality['unmapped_rows']:,} dòng ngoài danh mục · "
            f"{len(af_quality['conflict_ltts']):,} LTT xung đột · "
            f"{len(af_quality['conflict_stats']):,} phiếu xung đột.",
            flush=True,
        )

    if out_dir.exists():
        for pattern in ("*.json.gz", "manifest.json"):
            for path in out_dir.glob(pattern):
                path.unlink()
    out_dir.mkdir(parents=True, exist_ok=True)

    temp_root = Path(
        tempfile.mkdtemp(prefix=f"{plant_code.lower()}-source-jsonl-")
    )
    handles: dict[str, Any] = {}
    shard_stats: dict[str, ShardStats] = {}
    scanned = 0
    accepted = 0
    plant_counts: dict[str, int] = defaultdict(int)
    global_ltts: set[str] = set()
    global_stats: set[str] = set()
    global_machines: set[str] = set()
    global_operators: set[str] = set()
    global_date_min = ""
    global_date_max = ""
    future_date_rows = 0

    actual_by_ltt: dict[str, float] = defaultdict(float)
    plan_by_ltt: dict[str, float] = {}
    plan_conflict: set[str] = set()
    schedule_candidates: dict[str, dict[str, Any]] = {}
    schedule_start, schedule_end = schedule_window()

    try:
        with source_rows(
            workbook_path, source_type, sheet_name
        ) as rows:
            header_ok = False
            for row_number, raw in rows:
                if row_number < 9:
                    continue
                if row_number == 9:
                    header = " | ".join(clean_text(item) for item in raw)
                    header_ok = (
                        "Lệnh công đoạn" in header and "Số thống kê" in header
                    )
                    if not header_ok:
                        raise RuntimeError(
                            "Không nhận diện được header dòng 9: thiếu "
                            "'Lệnh công đoạn' hoặc 'Số thống kê'."
                        )
                    continue
                scanned += 1
                plant = clean_text(raw_at(raw, 4)) or "(trống)"
                plant_counts[plant] += 1
                if plant != plant_code:
                    continue
                if not (
                    clean_text(raw_at(raw, 2))
                    or clean_text(raw_at(raw, 45))
                    or clean_text(raw_at(raw, 5))
                ):
                    continue

                accepted += 1
                processed = process_raw_row(
                    raw,
                    other_label,
                    segment_mode,
                    af_quality["conflict_ltts"],
                    af_quality["conflict_stats"],
                )
                processed_obj = dict(zip(PROCESSED_COLS, processed))
                key = str(processed_obj["month"])
                if key not in handles:
                    jsonl_path = temp_root / f"{len(handles):02d}.jsonl"
                    handles[key] = jsonl_path.open("wb")
                    shard_stats[key] = ShardStats(
                        key=key, file_name=safe_shard_name(key)
                    )
                handles[key].write(compact_json(processed).encode("utf-8") + b"\n")
                shard_stats[key].add(processed)

                ltt = str(processed_obj["ltt"] or "")
                stat = str(processed_obj["stat"] or "")
                machine = str(processed_obj["machine"] or "")
                operator = str(processed_obj["operator"] or "")
                day = str(processed_obj["date"] or "")
                if ltt:
                    global_ltts.add(ltt)
                    actual_by_ltt[ltt] += to_num(raw_at(raw, 58))
                    plan = to_num(raw_at(raw, 18))
                    if plan > 0:
                        if ltt in plan_by_ltt and plan_by_ltt[ltt] != plan:
                            plan_conflict.add(ltt)
                        else:
                            plan_by_ltt.setdefault(ltt, plan)
                if stat:
                    global_stats.add(stat)
                if machine:
                    global_machines.add(machine)
                if operator:
                    global_operators.add(operator)
                if day:
                    if day > generated_dt.date().isoformat():
                        future_date_rows += 1
                    if not global_date_min or day < global_date_min:
                        global_date_min = day
                    if not global_date_max or day > global_date_max:
                        global_date_max = day

                if ltt and schedule_variant_overlaps(raw, schedule_start, schedule_end):
                    group = schedule_candidates.setdefault(
                        ltt,
                        {
                            "l": ltt,
                            "f": clean_text(raw_at(raw, 5)),
                            "g": clean_text(raw_at(raw, 6)),
                            "k": clean_text(raw_at(raw, 10)),
                            "af_values": set(),
                            "variants": {},
                            "filters": set(),
                        },
                    )
                    group["af_values"].add(
                        normalize_af(raw_at(raw, 31)) or "(trống)"
                    )
                    ab = clean_text(raw_at(raw, 27))
                    machine_filter = clean_text(raw_at(raw, 47) or raw_at(raw, 27))
                    if ab:
                        group["filters"].add(ab)
                    if machine_filter:
                        group["filters"].add(machine_filter)
                    variant_key = (
                        ab,
                        str(raw_at(raw, 23)),
                        str(raw_at(raw, 24)),
                    )
                    group["variants"].setdefault(
                        variant_key,
                        {
                            "ab": ab,
                            "x": raw_at(raw, 23),
                            "y": raw_at(raw, 24),
                            "seg": processed_obj["segment"],
                        },
                    )

                if accepted % 25_000 == 0:
                    print(
                        f"Đã xử lý {accepted:,} dòng {plant_code} "
                        f"(đã quét {scanned:,})...",
                        flush=True,
                    )
            if not header_ok:
                raise RuntimeError("Không đọc được dòng header 9.")
    finally:
        for handle in handles.values():
            handle.close()

    if not accepted:
        raise RuntimeError(f"Không tìm thấy dữ liệu nhà máy {plant_code}.")

    periods: list[dict[str, Any]] = []
    sorted_keys = sorted(
        shard_stats,
        key=lambda item: ("9999-99" if item == "Không tháng" else item),
    )
    for key in sorted_keys:
        stats = shard_stats[key]
        payload_meta = stats.meta(source_label, generated)
        prefix = (
            '{"meta":'
            + compact_json(payload_meta)
            + ',"cols":'
            + compact_json(PROCESSED_COLS)
            + ',"rows":['
        )
        output_path = out_dir / stats.file_name
        jsonl_path = temp_root / f"{list(handles).index(key):02d}.jsonl"
        write_gzip_json(output_path, prefix, jsonl_path, "]}")
        periods.append(
            {
                "value": key,
                "label": month_label(key),
                "file": stats.file_name,
                "rows": stats.rows,
                "ltt": len(stats.ltts),
                "stats": len(stats.stats),
                "machines": len(stats.machines),
                "operators": len(stats.operators),
                "date_min": stats.date_min,
                "date_max": stats.date_max,
                "exec_summary": stats.exec_summary(),
                "bytes_gzip": output_path.stat().st_size,
            }
        )

    schedule_rows: list[dict[str, Any]] = []
    for ltt, group in schedule_candidates.items():
        variants = list(group["variants"].values())
        af_values = sorted(group["af_values"])
        af_conflict = (
            len(af_values) > 1
            or ltt in af_quality["conflict_ltts"]
        )
        conflict = len(variants) > 1 or ltt in plan_conflict or af_conflict
        for variant in variants:
            schedule_rows.append(
                {
                    "l": ltt,
                    "f": group["f"],
                    "g": group["g"],
                    "k": group["k"],
                    "af": " / ".join(af_values),
                    "s": plan_by_ltt.get(ltt, 0),
                    "bg": actual_by_ltt.get(ltt, 0),
                    "ab": variant["ab"],
                    "x": variant["x"],
                    "y": variant["y"],
                    "seg": variant["seg"],
                    "mf": sorted(group["filters"]),
                    "t": 1,
                    "sourceConflict": conflict,
                    "afConflict": af_conflict,
                }
            )
    schedule_rows.sort(
        key=lambda row: (
            clean_text(row["af"]),
            clean_text(row["ab"]),
            str(row["x"]),
            clean_text(row["l"]),
        )
    )
    schedule_meta = {
        "source": source_label,
        "generated": generated,
        "plant": plant_code,
        "plant_name": plant_name,
        "window_start": schedule_start.isoformat(),
        "window_end": schedule_end.isoformat(),
        "rows": len(schedule_rows),
        "ltt": len({row["l"] for row in schedule_rows}),
    }
    schedule_payload = {
        "meta": schedule_meta,
        "rows": schedule_rows,
    }
    schedule_path = out_dir / "schedule_current.json.gz"
    write_gzip_json(
        schedule_path,
        compact_json(schedule_payload),
        None,
        "",
    )

    dated_periods = [
        item["value"]
        for item in periods
        if item["value"] != "Không tháng"
    ]
    current_period = generated_dt.strftime("%Y-%m")
    valid_default_periods = [
        period for period in dated_periods if period <= current_period
    ]
    latest_period = (
        valid_default_periods[-1]
        if valid_default_periods
        else (dated_periods[-1] if dated_periods else "Không tháng")
    )
    manifest = {
        "schema": SCHEMA_VERSION,
        "schema_version": 1,
        "plant": plant_code,
        "plant_name": plant_name,
        "segment_mode": segment_mode,
        "generated_at": generated_iso,
        "generated_display": generated,
        "latest_period": latest_period,
        "latest_detected_period": (
            dated_periods[-1] if dated_periods else "Không tháng"
        ),
        "load_mode": "single-period",
        "load_mode_note": (
            "Dashboard chỉ nạp một tháng mỗi lần để tránh giữ 279 nghìn dòng "
            "trong RAM trình duyệt."
        ),
        "source": {
            "type": source_type,
            "file_id": file_id,
            "file_name": file_name,
            "sheet_name": sheet_name,
            "range": SOURCE_RANGE,
            "size": source_meta.get("size", workbook_path.stat().st_size),
            "modified": source_meta.get("modified", ""),
            "etag": source_meta.get("etag", ""),
            "signature": source_meta["signature"],
        },
        "global": {
            "scanned_rows": scanned,
            "accepted_rows": accepted,
            "plant_counts": dict(
                sorted(plant_counts.items(), key=lambda item: item[1], reverse=True)
            ),
            "ltt": len(global_ltts),
            "stats": len(global_stats),
            "machines": len(global_machines),
            "operators": len(global_operators),
            "date_min": global_date_min,
            "date_max": global_date_max,
            "future_date_rows": future_date_rows,
            "undated_rows": shard_stats.get(
                "Không tháng", ShardStats("", "")
            ).rows,
            "af_quality": {
                "catalog_size": len(GS1_AF_SEGMENT_CATALOG),
                "valid_rows": af_quality["valid_rows"],
                "valid_rate": af_quality["valid_rate"],
                "missing_rows": af_quality["missing_rows"],
                "raw_missing_rows": af_quality["raw_missing_rows"],
                "unmapped_rows": af_quality["unmapped_rows"],
                "raw_unmapped_rows": af_quality["raw_unmapped_rows"],
                "conflict_rows": af_quality["conflict_rows"],
                "unmapped_codes": af_quality["unmapped_codes"],
                "conflict_ltts": len(af_quality["conflict_ltts"]),
                "conflict_stats": len(af_quality["conflict_stats"]),
                "code_rows": af_quality["code_rows"],
            },
        },
        "periods": periods,
        "schedule": {
            "file": schedule_path.name,
            **schedule_meta,
            "bytes_gzip": schedule_path.stat().st_size,
        },
        "processing_seconds": round(time.monotonic() - started, 1),
    }
    (out_dir / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    shutil.rmtree(temp_root, ignore_errors=True)
    return manifest


def local_source_meta(path: Path) -> dict[str, Any]:
    stat = path.stat()
    digest = hashlib.sha256()
    with path.open("rb") as source:
        while chunk := source.read(4 * 1024 * 1024):
            digest.update(chunk)
    return {
        "signature": hashlib.sha256(
            f"{SCHEMA_VERSION}|{stat.st_size}|{digest.hexdigest()}".encode("utf-8")
        ).hexdigest()[:24],
        "size": stat.st_size,
        "modified": datetime.fromtimestamp(stat.st_mtime, timezone.utc).isoformat(),
        "etag": digest.hexdigest(),
    }


def load_factory_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise RuntimeError(f"Không tìm thấy cấu hình nhà máy: {path}.")
    config = json.loads(path.read_text(encoding="utf-8"))
    factory = config.get("factory") or {}
    source = config.get("source") or {}
    for key, label in (
        (factory.get("code"), "factory.code"),
        (factory.get("name"), "factory.name"),
        (source.get("file_id"), "source.file_id"),
        (source.get("file_name"), "source.file_name"),
        (source.get("sheet_name"), "source.sheet_name"),
    ):
        if not clean_text(key):
            raise RuntimeError(f"Cấu hình thiếu {label}.")
    return config


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--config", type=Path, default=Path("factory-config.json")
    )
    parser.add_argument("--probe", action="store_true")
    parser.add_argument("--github-output", type=Path)
    parser.add_argument("--input", type=Path)
    parser.add_argument("--out", type=Path, default=Path("site/data"))
    parser.add_argument("--plant")
    parser.add_argument("--plant-name")
    parser.add_argument("--file-id")
    parser.add_argument("--file-name")
    parser.add_argument("--sheet-name")
    parser.add_argument(
        "--source-type",
        choices=("drive_xlsx", "google_sheet_csv"),
    )
    parser.add_argument("--segment-mode")
    parser.add_argument("--source-url")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config = load_factory_config(args.config)
    factory_cfg = config["factory"]
    source_cfg = config["source"]
    plant_code = args.plant or clean_text(factory_cfg["code"])
    plant_name = args.plant_name or clean_text(factory_cfg["name"])
    file_id = args.file_id or clean_text(source_cfg["file_id"])
    file_name = args.file_name or clean_text(source_cfg["file_name"])
    sheet_name = args.sheet_name or clean_text(source_cfg["sheet_name"])
    source_type = (
        args.source_type
        or clean_text(source_cfg.get("type"))
        or SOURCE_TYPE
    )
    segment_mode = (
        args.segment_mode
        or clean_text(config.get("dashboard", {}).get("segment_mode"))
        or SEGMENT_MODE
    )
    resolved_source_url = (
        args.source_url
        or clean_text(source_cfg.get("download_url"))
        or source_url(source_type, file_id, sheet_name)
    )
    if args.input:
        source_meta = local_source_meta(args.input)
    else:
        source_meta = probe_source(
            resolved_source_url, file_id, source_type, plant_code
        )
    if args.probe:
        output = {
            "signature": source_meta["signature"],
            "size": source_meta["size"],
            "modified": source_meta.get("modified", ""),
        }
        print(compact_json(output))
        if args.github_output:
            with args.github_output.open("a", encoding="utf-8") as stream:
                for key, value in output.items():
                    stream.write(f"{key}={value}\n")
        return 0

    temporary_download: Path | None = None
    workbook_path = args.input
    if not workbook_path:
        temp_dir = Path(
            tempfile.mkdtemp(prefix=f"{plant_code.lower()}-source-")
        )
        temporary_download = temp_dir / file_name
        print(
            f"Tải nguồn {file_name}"
            + (
                f" · {source_meta['size'] / 1024 / 1024:.1f} MB..."
                if source_meta["size"]
                else "..."
            ),
            flush=True,
        )
        download_source(
            resolved_source_url,
            temporary_download,
            int(source_meta["size"]),
            source_type,
            plant_code,
        )
        workbook_path = temporary_download
    assert workbook_path is not None
    try:
        manifest = build_data(
            workbook_path=workbook_path,
            out_dir=args.out,
            plant_code=plant_code,
            plant_name=plant_name,
            source_meta=source_meta,
            source_type=source_type,
            segment_mode=segment_mode,
            sheet_name=sheet_name,
            file_name=file_name,
            file_id=file_id,
        )
    finally:
        if temporary_download:
            shutil.rmtree(temporary_download.parent, ignore_errors=True)
    print(
        "Hoàn tất: "
        f"{manifest['global']['accepted_rows']:,} dòng, "
        f"{len(manifest['periods'])} phân vùng, "
        f"{manifest['processing_seconds']} giây.",
        flush=True,
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"LỖI: {exc}", file=sys.stderr)
        raise
