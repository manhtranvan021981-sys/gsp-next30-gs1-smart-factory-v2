#!/usr/bin/env python3
"""Fail deployment when a generated factory data package is incomplete."""

from __future__ import annotations

import argparse
import gzip
import json
from pathlib import Path


EXPECTED_COLS = [
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

GS1_AF_SEGMENTS = {
    "01_Nhóm hàng Hộp cứng",
    "02_Nhóm hàng Hộp thường (hộp mềm)",
    "03_Nhóm hàng Hộp bồi duplex",
    "04_Nhóm hàng Hộp bồi label",
    "05_Nhóm hàng Hộp Flexo carton",
    "06_Nhóm hàng Hộp Flexo process",
    "07_Nhóm hàng PK phôi carton",
    "08_Nhóm hàng Sách hướng dẫn",
    "09_Nhóm hàng Pallet",
    "10_Nhóm hàng Phôi sóng",
    "11_Nhóm hàng Gia công in",
    "12_Nhóm hàng Khay giấy",
    "13_Nhóm hàng Túi giấy",
    "14_Nhóm Nguyên vật liệu chính",
    "15_Nhóm Nguyên vật liệu phụ",
    "16_Nhóm CCDC, VTB, VIT, VPP",
    "17_Nhóm hàng khác",
    "18_Nhóm Lề, phế",
    "19_Nhóm hàng thương mại",
    "00_Chưa khai báo dòng hàng mẹ",
    "98_Xung đột AF theo LTT/phiếu",
    "99_AF chưa ánh xạ",
}


def read_gzip_json(path: Path):
    with gzip.open(path, "rt", encoding="utf-8") as stream:
        return json.load(stream)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", type=Path, default=Path("site/data"))
    parser.add_argument(
        "--config", type=Path, default=Path("factory-config.json")
    )
    parser.add_argument("--plant")
    args = parser.parse_args()
    config = json.loads(args.config.read_text(encoding="utf-8"))
    expected_plant = args.plant or config["factory"]["code"]
    assert config["dashboard"]["segment_mode"] == "gs1_parent_line_af"
    manifest_path = args.data / "manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["schema"] == "gsp-factory-static-shards-v2-af-alias"
    assert manifest["plant"] == expected_plant, (
        f"Sai nhà máy: manifest={manifest['plant']}, "
        f"yêu cầu={expected_plant}."
    )
    assert manifest["global"]["accepted_rows"] > 0
    assert manifest["periods"], "Không có phân vùng tháng."
    total = 0
    segment_counts: dict[str, int] = {}
    for period in manifest["periods"]:
        payload = read_gzip_json(args.data / period["file"])
        assert payload["cols"] == EXPECTED_COLS
        assert len(payload["rows"]) == period["rows"]
        assert payload["meta"]["rows"] == period["rows"]
        assert period["rows"] <= 100_000, (
            f"Phân vùng {period['value']} có {period['rows']} dòng; "
            "quá ngưỡng an toàn trình duyệt."
        )
        assert all(len(row) == len(EXPECTED_COLS) for row in payload["rows"])
        segment_index = EXPECTED_COLS.index("segment")
        for row in payload["rows"]:
            segment = row[segment_index]
            assert segment in GS1_AF_SEGMENTS, f"Mảng AF không hợp lệ: {segment!r}"
            segment_counts[segment] = segment_counts.get(segment, 0) + 1
        total += period["rows"]
    assert total == manifest["global"]["accepted_rows"]
    af_quality = manifest["global"]["af_quality"]
    assert af_quality["catalog_size"] == 19
    assert af_quality["aliases"] == {
        "SOB": "PHOI",
        "SOE": "PHOI",
        "SOA": "PHOI",
        "SBA": "PHOI",
        "SBC": "PHOI",
        "SBE": "PHOI",
        "SOC": "PHOI",
        "SOG": "PHOI",
        "SEE": "PHOI",
        "SEC": "PHOI",
        "DUP": "HBL",
    }
    assert (
        af_quality["valid_rows"]
        + af_quality["missing_rows"]
        + af_quality["unmapped_rows"]
        + af_quality["conflict_rows"]
        == total
    ), "Không khớp tổng phân loại AF hợp lệ/00/98/99."
    assert segment_counts.get("00_Chưa khai báo dòng hàng mẹ", 0) == af_quality["missing_rows"]
    assert segment_counts.get("98_Xung đột AF theo LTT/phiếu", 0) == af_quality["conflict_rows"]
    assert segment_counts.get("99_AF chưa ánh xạ", 0) == af_quality["unmapped_rows"]
    schedule = read_gzip_json(args.data / manifest["schedule"]["file"])
    assert schedule["meta"]["plant"] == expected_plant
    assert len(schedule["rows"]) == manifest["schedule"]["rows"]
    assert all(row["seg"] in GS1_AF_SEGMENTS for row in schedule["rows"])
    print(
        f"VERIFY OK: {total:,} dòng, {len(manifest['periods'])} phân vùng, "
        f"{len(schedule['rows']):,} dòng lịch hiện hành."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
