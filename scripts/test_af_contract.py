#!/usr/bin/env python3
"""End-to-end contract test for the GS1 parent-line AF filter."""

from __future__ import annotations

import argparse
import gzip
import importlib.util
import json
import shutil
import sys
from contextlib import contextmanager
from datetime import timedelta
from pathlib import Path
from typing import Any, Iterable


SCRIPT_DIR = Path(__file__).resolve().parent
SITE_DIR = SCRIPT_DIR.parent
PROCESSOR_PATH = SCRIPT_DIR / "process_excel.py"


def load_processor() -> Any:
    spec = importlib.util.spec_from_file_location("gs1_process_excel", PROCESSOR_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError("Không thể nạp process_excel.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def blank_row() -> list[Any]:
    return [""] * 98


def header_row() -> list[Any]:
    row = blank_row()
    row[2] = "Lệnh công đoạn"
    row[45] = "Số thống kê"
    return row


def data_row(
    *,
    ltt: str,
    stat: str,
    af: str,
    sequence: int,
    schedule_start: str = "",
    schedule_end: str = "",
    machine: str = "",
) -> list[Any]:
    row = blank_row()
    row[2] = ltt
    row[4] = "GS1"
    row[5] = f"MAT-{sequence:03d}"
    row[6] = f"Sản phẩm kiểm thử {sequence:03d}"
    row[18] = 100
    row[20] = 1
    row[23] = schedule_start
    row[24] = schedule_end
    row[27] = machine or f"M{sequence:02d}"
    row[29] = "Công đoạn kiểm thử"
    row[31] = af
    row[43] = "2026-07-01"
    row[45] = stat
    row[47] = machine or f"M{sequence:02d}"
    row[49] = "NV kiểm thử"
    row[58] = 100
    row[59] = 98
    row[60] = 2
    row[62] = "SP"
    row[79] = 100
    row[81] = 100
    row[85] = 1
    row[87] = 1
    row[88] = 0.9
    row[89] = 0.9
    row[90] = 0.99
    row[91] = 0.8019
    return row


def fixture_rows(processor: Any) -> list[tuple[int, list[Any]]]:
    rows: list[tuple[int, list[Any]]] = [(9, header_row())]
    sequence = 1
    window_start, _ = processor.schedule_window()
    schedule_start = window_start.isoformat()
    schedule_end = (window_start + timedelta(days=1)).isoformat()
    for code in processor.GS1_AF_SEGMENT_CATALOG:
        rows.append(
            (
                9 + sequence,
                data_row(
                    ltt=f"LTT-{code}",
                    stat=f"STAT-{code}",
                    af=code.lower() if code == "HBD" else f" {code} ",
                    sequence=sequence,
                ),
            )
        )
        sequence += 1
    for code in processor.GS1_AF_ALIAS:
        rows.append(
            (
                9 + sequence,
                data_row(
                    ltt=f"LTT-ALIAS-{code}",
                    stat=f"STAT-ALIAS-{code}",
                    af=f" {code.lower()} ",
                    sequence=sequence,
                ),
            )
        )
        sequence += 1
    for af in ("SOB", "SOE"):
        rows.append(
            (
                9 + sequence,
                data_row(
                    ltt="LTT-ALIAS-PHOI-NO-CONFLICT",
                    stat=f"STAT-ALIAS-PHOI-{af}",
                    af=af,
                    sequence=sequence,
                    schedule_start=schedule_start,
                    schedule_end=schedule_end,
                    machine="M-ALIAS-PHOI",
                ),
            )
        )
        sequence += 1
    for af in ("DUP", "HBL"):
        rows.append(
            (
                9 + sequence,
                data_row(
                    ltt=f"LTT-ALIAS-HBL-{af}",
                    stat="STAT-ALIAS-HBL-NO-CONFLICT",
                    af=af,
                    sequence=sequence,
                    schedule_start=schedule_start,
                    schedule_end=schedule_end,
                ),
            )
        )
        sequence += 1
    rows.append(
        (
            9 + sequence,
            data_row(
                ltt="LTT-MISSING",
                stat="STAT-MISSING",
                af="",
                sequence=sequence,
            ),
        )
    )
    sequence += 1
    rows.append(
        (
            9 + sequence,
            data_row(
                ltt="LTT-UNMAPPED",
                stat="STAT-UNMAPPED",
                af="NEWCODE",
                sequence=sequence,
            ),
        )
    )
    sequence += 1
    for af in ("HOC", "HOT"):
        rows.append(
            (
                9 + sequence,
                data_row(
                    ltt="LTT-CONFLICT",
                    stat=f"STAT-LTT-CONFLICT-{af}",
                    af=af,
                    sequence=sequence,
                ),
            )
        )
        sequence += 1
    for af in ("HBD", "HBL"):
        rows.append(
            (
                9 + sequence,
                data_row(
                    ltt=f"LTT-STAT-CONFLICT-{af}",
                    stat="STAT-CONFLICT",
                    af=af,
                    sequence=sequence,
                ),
            )
        )
        sequence += 1
    for af in ("SOB", "DUP"):
        rows.append(
            (
                9 + sequence,
                data_row(
                    ltt="LTT-ALIAS-TRUE-CONFLICT",
                    stat=f"STAT-ALIAS-TRUE-CONFLICT-{af}",
                    af=af,
                    sequence=sequence,
                    schedule_start=schedule_start,
                    schedule_end=schedule_end,
                    machine="M-ALIAS-TRUE-CONFLICT",
                ),
            )
        )
        sequence += 1
    for af in ("SEC", "HBD"):
        rows.append(
            (
                9 + sequence,
                data_row(
                    ltt=f"LTT-ALIAS-STAT-CONFLICT-{af}",
                    stat="STAT-ALIAS-TRUE-CONFLICT",
                    af=af,
                    sequence=sequence,
                    schedule_start=schedule_start,
                    schedule_end=schedule_end,
                ),
            )
        )
        sequence += 1
    return rows


def read_all_segments(out_dir: Path) -> list[str]:
    manifest = json.loads((out_dir / "manifest.json").read_text(encoding="utf-8"))
    segments: list[str] = []
    for period in manifest["periods"]:
        with gzip.open(out_dir / period["file"], "rt", encoding="utf-8") as stream:
            payload = json.load(stream)
        segment_index = payload["cols"].index("segment")
        segments.extend(str(row[segment_index]) for row in payload["rows"])
    return segments


def read_schedule(out_dir: Path) -> dict[str, Any]:
    manifest = json.loads((out_dir / "manifest.json").read_text(encoding="utf-8"))
    with gzip.open(
        out_dir / manifest["schedule"]["file"], "rt", encoding="utf-8"
    ) as stream:
        return json.load(stream)


def assert_frontend_contract(processor: Any) -> None:
    index_text = (SITE_DIR / "index.html").read_text(encoding="utf-8")
    cursor = -1
    for label in processor.GS1_AF_SEGMENT_CATALOG.values():
        next_cursor = index_text.find(label, cursor + 1)
        if next_cursor < 0:
            raise AssertionError(f"Thiếu nhãn AF trong index.html: {label}")
        if next_cursor <= cursor:
            raise AssertionError("Thứ tự nhãn AF trong index.html không đúng 01–19")
        cursor = next_cursor
    for label in (
        processor.AF_MISSING_LABEL,
        processor.AF_CONFLICT_LABEL,
        processor.AF_UNMAPPED_LABEL,
    ):
        if label not in index_text:
            raise AssertionError(f"Thiếu nhãn kiểm soát trong index.html: {label}")
    for alias, canonical in processor.GS1_AF_ALIAS.items():
        if f"{alias}:'{canonical}'" not in index_text:
            raise AssertionError(
                f"Frontend thiếu alias AF {alias} → {canonical}"
            )
    config = json.loads((SITE_DIR / "factory-config.json").read_text(encoding="utf-8"))
    if config["dashboard"]["segment_mode"] != "gs1_parent_line_af":
        raise AssertionError("factory-config.json chưa dùng gs1_parent_line_af")


def run(out_dir: Path) -> None:
    processor = load_processor()
    rows = fixture_rows(processor)

    @contextmanager
    def fake_source_rows(
        source_path: Path,
        source_type: str,
        sheet_name: str,
    ) -> Iterable[Iterable[tuple[int, list[Any]]]]:
        del source_path, source_type, sheet_name
        yield iter(rows)

    processor.source_rows = fake_source_rows
    shutil.rmtree(out_dir, ignore_errors=True)
    out_dir.mkdir(parents=True, exist_ok=True)
    synthetic_source = out_dir / "synthetic-gs1.xlsx"
    synthetic_source.touch()
    manifest = processor.build_data(
        workbook_path=synthetic_source,
        out_dir=out_dir,
        plant_code="GS1",
        plant_name="GoldSun Hà Nội",
        source_meta={
            "signature": "synthetic-af-contract",
            "size": 0,
            "modified": "synthetic",
            "etag": "",
        },
        source_type="drive_xlsx",
        segment_mode="gs1_parent_line_af",
        sheet_name="P3.Tổng hợp lệnh thao tác",
        file_name="synthetic-gs1.xlsx",
        file_id="synthetic",
    )
    af_quality = manifest["global"]["af_quality"]
    if manifest["global"]["accepted_rows"] != 44:
        raise AssertionError(
            "accepted_rows: kỳ vọng 44, thực tế "
            f"{manifest['global']['accepted_rows']}"
        )
    expected = {
        "valid_rows": 34,
        "missing_rows": 1,
        "unmapped_rows": 1,
        "conflict_rows": 8,
        "alias_rows": 17,
    }
    for key, value in expected.items():
        if af_quality[key] != value:
            raise AssertionError(f"{key}: kỳ vọng {value}, thực tế {af_quality[key]}")
    if af_quality["conflict_ltts"] != 2:
        raise AssertionError("Kỳ vọng đúng 2 LTT xung đột AF chuẩn")
    if af_quality["conflict_stats"] != 2:
        raise AssertionError("Kỳ vọng đúng 2 phiếu xung đột AF chuẩn")
    segments = read_all_segments(out_dir)
    expected_segments = list(processor.GS1_AF_SEGMENT_CATALOG.values())
    for label in expected_segments:
        expected_count = 13 if label.startswith("10_") else 4 if label.startswith("04_") else 1
        if segments.count(label) != expected_count:
            raise AssertionError(
                f"Nhãn {label}: kỳ vọng {expected_count}, "
                f"thực tế {segments.count(label)}"
            )
    if segments.count(processor.AF_MISSING_LABEL) != 1:
        raise AssertionError("Nhóm AF trống phải có đúng 1 dòng")
    if segments.count(processor.AF_UNMAPPED_LABEL) != 1:
        raise AssertionError("Nhóm AF chưa ánh xạ phải có đúng 1 dòng")
    if segments.count(processor.AF_CONFLICT_LABEL) != 8:
        raise AssertionError("Nhóm xung đột AF phải có đúng 8 dòng")
    schedule = read_schedule(out_dir)
    if len(schedule["rows"]) != 6:
        raise AssertionError(
            f"Lịch alias: kỳ vọng 6 dòng, thực tế {len(schedule['rows'])}"
        )
    schedule_by_ltt = {row["l"]: row for row in schedule["rows"]}
    phoi_row = schedule_by_ltt["LTT-ALIAS-PHOI-NO-CONFLICT"]
    if phoi_row["seg"] != processor.GS1_AF_SEGMENT_CATALOG["PHOI"]:
        raise AssertionError("Lịch máy chưa quy SOB/SOE về PHOI")
    if phoi_row["afConflict"]:
        raise AssertionError("Lịch máy báo xung đột giả cho SOB + SOE")
    for ltt in ("LTT-ALIAS-HBL-DUP", "LTT-ALIAS-HBL-HBL"):
        row = schedule_by_ltt[ltt]
        if row["seg"] != processor.GS1_AF_SEGMENT_CATALOG["HBL"]:
            raise AssertionError(f"Lịch máy chưa quy đúng HBL: {ltt}")
        if row["afConflict"]:
            raise AssertionError(f"Lịch máy báo xung đột giả: {ltt}")
    for ltt in (
        "LTT-ALIAS-TRUE-CONFLICT",
        "LTT-ALIAS-STAT-CONFLICT-SEC",
        "LTT-ALIAS-STAT-CONFLICT-HBD",
    ):
        row = schedule_by_ltt[ltt]
        if row["seg"] != processor.AF_CONFLICT_LABEL or not row["afConflict"]:
            raise AssertionError(f"Lịch máy bỏ sót xung đột thật: {ltt}")
    if processor.normalize_af(" hbd ") != "HBD":
        raise AssertionError("Chuẩn hóa AF chưa trim và upper-case đúng")
    for alias, canonical in processor.GS1_AF_ALIAS.items():
        if processor.canonical_af(f" {alias.lower()} ") != canonical:
            raise AssertionError(f"Ánh xạ alias sai: {alias} → {canonical}")
    assert_frontend_contract(processor)
    print(
        "AF contract OK: 19 mã chuẩn + 11 alias + 00/98/99 · "
        "44/44 dòng đối soát · frontend/backend đồng bộ."
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--out",
        type=Path,
        default=SITE_DIR / ".qa-af-contract",
    )
    args = parser.parse_args()
    run(args.out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
