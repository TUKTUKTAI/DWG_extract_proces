#!/usr/bin/env python3
"""
Prototype DXF -> extract .xlsx generator (voldoende schema-compatibel voor P22_0002_Main.java-tests).

Dit repliceert Mosaic DataExtract NIET volledig, maar maakt wel een Excel-bestand met de
kolomkoppen die de Java-naverwerkingsstap verwacht.

Benodigdheden:
  pip install ezdxf openpyxl
"""

from __future__ import annotations

import argparse
import csv
import math
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, List

try:
    import ezdxf
except ImportError:  # pragma: no cover
    print("Missing dependency: ezdxf (pip install ezdxf)", file=sys.stderr)
    raise

try:
    from openpyxl import Workbook
except ImportError:  # pragma: no cover
    print("Missing dependency: openpyxl (pip install openpyxl)", file=sys.stderr)
    raise


HEADERS: List[str] = [
    "File Name",
    "File Location",
    "Name",
    "Position X",
    "Position Y",
    "Position Z",
    "Start X",
    "Start Y",
    "Start Z",
    "End X",
    "End Y",
    "End Z",
    "DATUM",
    "PLAATDEEL",
    "GETEKEND",
    "GEZIEN",
    "REVISIE",
    "SOORT",
    "SOORT_HANDW",
    "Value",
    "Layer",
    "Length",
    "Rotation",
    "File Modified",
    "CAMTAG",
    "Closed",
    "_\\U+002E\\U+002E\\U+002E",
    "Prompt",
    "Tag",
    "CADFIL",
    "EXTRAINFO",
    "EXTRAINFO(1)",
    "UPPERCASE",
    "LOWERCASE",
    "LTYPE",
    "BW",
    "BH",
    "Contents",
]


ATTR_TO_COL = {
    "CAMTAG": "CAMTAG",
    "CADFIL": "CADFIL",
    "EXTRAINFO": "EXTRAINFO",
    "EXTRAINFO(1)": "EXTRAINFO(1)",
    "UPPERCASE": "UPPERCASE",
    "LOWERCASE": "LOWERCASE",
    "LTYPE": "LTYPE",
    "BW": "BW",
    "BH": "BH",
}


def fmt_num(value) -> str:
    if value is None:
        return ""
    try:
        num = float(value)
    except Exception:
        return str(value)
    text = f"{num:.4f}"
    return text


def mtime_string(path: Path) -> str:
    try:
        return datetime.fromtimestamp(path.stat().st_mtime).strftime("%d-%m-%Y")
    except Exception:
        return ""


def base_row(source: Path, entity_name: str, layer: str = "") -> dict:
    return {
        "File Name": source.name,
        "File Location": str(source.parent),
        "Name": entity_name,
        "Position X": "",
        "Position Y": "",
        "Position Z": "",
        "Start X": "",
        "Start Y": "",
        "Start Z": "",
        "End X": "",
        "End Y": "",
        "End Z": "",
        "DATUM": "",
        "PLAATDEEL": "",
        "GETEKEND": "",
        "GEZIEN": "",
        "REVISIE": "",
        "SOORT": "",
        "SOORT_HANDW": "",
        "Value": "",
        "Layer": layer or "",
        "Length": "",
        "Rotation": "",
        "File Modified": mtime_string(source),
        "CAMTAG": "",
        "Closed": "",
        "_\\U+002E\\U+002E\\U+002E": "",
        "Prompt": "",
        "Tag": "",
        "CADFIL": "",
        "EXTRAINFO": "",
        "EXTRAINFO(1)": "",
        "UPPERCASE": "",
        "LOWERCASE": "",
        "LTYPE": "",
        "BW": "",
        "BH": "",
        "Contents": "",
    }


def set_pos(row: dict, x=None, y=None, z=None) -> None:
    row["Position X"] = fmt_num(x) if x is not None else ""
    row["Position Y"] = fmt_num(y) if y is not None else ""
    row["Position Z"] = fmt_num(z) if z is not None else ""


def set_start_end(row: dict, start, end) -> None:
    sx, sy, sz = (list(start) + [0, 0, 0])[:3]
    ex, ey, ez = (list(end) + [0, 0, 0])[:3]
    row["Start X"] = fmt_num(sx)
    row["Start Y"] = fmt_num(sy)
    row["Start Z"] = fmt_num(sz)
    row["End X"] = fmt_num(ex)
    row["End Y"] = fmt_num(ey)
    row["End Z"] = fmt_num(ez)
    row["Length"] = fmt_num(math.dist((sx, sy, sz), (ex, ey, ez)))


def plain_mtext(entity) -> str:
    try:
        return entity.plain_text()
    except Exception:
        try:
            return entity.text
        except Exception:
            return ""


def insert_row(source: Path, entity) -> dict:
    row = base_row(source, entity.dxf.name, entity.dxf.layer)
    ins = entity.dxf.insert
    set_pos(row, ins.x, ins.y, getattr(ins, "z", 0))
    row["Rotation"] = fmt_num(getattr(entity.dxf, "rotation", 0))

    attrs = {}
    for att in getattr(entity, "attribs", []):
        tag = str(getattr(att.dxf, "tag", "") or "").strip()
        txt = str(getattr(att.dxf, "text", "") or "")
        if tag:
            attrs[tag.upper()] = txt

    for tag, col in ATTR_TO_COL.items():
        if tag.upper() in attrs:
            row[col] = attrs[tag.upper()]

    # Als deze block-attributen onder eenvoudigere tags voorkomen, map ze ook.
    for fallback in ("EXTRAINFO1", "EXTRAINFO_1"):
        if not row["EXTRAINFO(1)"] and fallback in attrs:
            row["EXTRAINFO(1)"] = attrs[fallback]

    # Zet iets in "Contents" voor zichtbaarheid tijdens prototypetests.
    if attrs:
        row["Contents"] = " | ".join([f"{k}={v}" for k, v in sorted(attrs.items())])

    return row


def entity_rows(source: Path) -> Iterable[dict]:
    doc = ezdxf.readfile(source)
    msp = doc.modelspace()

    for entity in msp:
        dxftype = entity.dxftype()
        layer = getattr(entity.dxf, "layer", "")

        try:
            if dxftype == "TEXT":
                row = base_row(source, "Text", layer)
                ins = entity.dxf.insert
                set_pos(row, ins.x, ins.y, getattr(ins, "z", 0))
                row["Rotation"] = fmt_num(getattr(entity.dxf, "rotation", 0))
                row["Value"] = str(getattr(entity.dxf, "text", "") or "")
                yield row

            elif dxftype == "MTEXT":
                row = base_row(source, "MText", layer)
                ins = entity.dxf.insert
                set_pos(row, ins.x, ins.y, getattr(ins, "z", 0))
                row["Rotation"] = fmt_num(getattr(entity.dxf, "rotation", 0))
                raw = str(getattr(entity, "text", "") or "")
                row["Value"] = raw
                row["Contents"] = plain_mtext(entity)
                yield row

            elif dxftype == "LINE":
                row = base_row(source, "Line", layer)
                start = entity.dxf.start
                end = entity.dxf.end
                set_start_end(row, start, end)
                yield row

            elif dxftype == "INSERT":
                yield insert_row(source, entity)

            elif dxftype == "ATTRIB":
                # Optioneel: toon als tekstachtige rij zodat Java attribuuttekst nog kan "zien" indien nodig.
                row = base_row(source, "Text", layer)
                ins = entity.dxf.insert
                set_pos(row, ins.x, ins.y, getattr(ins, "z", 0))
                row["Value"] = str(getattr(entity.dxf, "text", "") or "")
                row["Tag"] = str(getattr(entity.dxf, "tag", "") or "")
                yield row

            elif dxftype in ("LWPOLYLINE", "POLYLINE"):
                row = base_row(source, dxftype, layer)
                row["Closed"] = "1" if bool(entity.closed) else "0"
                try:
                    if dxftype == "LWPOLYLINE":
                        pts = list(entity.get_points("xy"))
                        if pts:
                            set_pos(row, pts[0][0], pts[0][1], 0)
                    else:
                        pts = [tuple(v.dxf.location) for v in entity.vertices]
                        if pts:
                            set_pos(row, pts[0][0], pts[0][1], pts[0][2] if len(pts[0]) > 2 else 0)
                    if len(pts) > 1:
                        length = 0.0
                        for i in range(1, len(pts)):
                            p1 = pts[i - 1]
                            p2 = pts[i]
                            length += math.dist(
                                (p1[0], p1[1], p1[2] if len(p1) > 2 else 0),
                                (p2[0], p2[1], p2[2] if len(p2) > 2 else 0),
                            )
                        row["Length"] = fmt_num(length)
                except Exception:
                    pass
                yield row

        except Exception:
            # Sla foutieve entities over zodat de bestandsverwerking door kan gaan.
            continue


def write_xlsx(rows: Iterable[dict], out_path: Path) -> int:
    wb = Workbook()
    ws = wb.active
    ws.title = "DataExtract"
    ws.append(HEADERS)
    count = 0
    for row in rows:
        ws.append([str(row.get(h, "")) if row.get(h, "") is not None else "" for h in HEADERS])
        count += 1
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return count


def iter_input_files(input_path: Path, recursive: bool) -> List[Path]:
    if input_path.is_file():
        return [input_path]
    pattern = "**/*.dxf" if recursive else "*.dxf"
    return sorted(p for p in input_path.glob(pattern) if p.is_file())


def has_nbd_signals(rows: List[dict]) -> bool:
    """Heuristiek voor NBD-bordspecificatie-tekeningen: zoek bekende block/entity-namen of gemapte attributen."""
    signal_names = {
        "BORDDATA",
        "BORDDATA1",
        "BPS_FILENAME",
        "PADRI",
        "FRSPEC",
        "FLSPECI",
        "SPEC01",
        "SPEC02",
        "SPEC03",
    }
    for row in rows:
        name = str(row.get("Name", "") or "").upper()
        value = str(row.get("Value", "") or "").upper()
        contents = str(row.get("Contents", "") or "").upper()

        if name in signal_names or any(name.startswith(s) for s in ("SPEC",)):
            return True
        if any(token in value for token in ("BORDDATA", "SPEC", "STANDPLAATS", "BPS_")):
            return True
        if any(token in contents for token in ("CADFIL=", "CAMTAG=", "EXTRAINFO=", "UPPERCASE=", "LOWERCASE=")):
            return True
        if any(str(row.get(k, "") or "").strip() for k in ("CADFIL", "CAMTAG", "UPPERCASE", "LOWERCASE", "LTYPE", "BW", "BH")):
            return True
    return False


@dataclass
class ErrorLogRow:
    timestamp: str
    source: str
    error: str


def append_error_log(error_log: Path, row: ErrorLogRow) -> None:
    error_log.parent.mkdir(parents=True, exist_ok=True)
    write_header = not error_log.exists()
    with error_log.open("a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        if write_header:
            writer.writerow(["timestamp", "source", "error"])
        writer.writerow([row.timestamp, row.source, row.error])


def derive_object_nummer_from_stem(stem: str) -> str:
    m = re.match(r"^(.+?)([A-Za-z]+)$", stem or "")
    return m.group(1) if m else (stem or "")


def append_niet_verwerkt_csv(csv_path: Path, *, type_value: str, source: str, object_nummer: str, reason: str) -> None:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    write_header = not csv_path.exists()
    with csv_path.open("a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        if write_header:
            writer.writerow(["tijd", "type", "bestand", "object_key", "object_nummer", "reden"])
        writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            type_value,
            source,
            "",
            object_nummer,
            reason,
        ])


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert DXF files to DataExtract-like .xlsx files.")
    parser.add_argument("--input", required=True, help="DXF file or folder")
    parser.add_argument("--output-dir", required=True, help="Output folder for generated extract .xlsx files")
    parser.add_argument("--recursive", action="store_true", help="Recurse into subfolders when input is a folder")
    parser.add_argument("--strict-nbd", action="store_true", help="Reject DXFs that do not look like NBD sign-spec drawings")
    parser.add_argument("--min-rows", type=int, default=1, help="Minimum extracted rows required to accept a DXF")
    parser.add_argument(
        "--error-log",
        default="",
        help="Optional error log CSV path; failed DXFs will be logged and processing continues",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)

    if not input_path.exists():
        print(f"Input does not exist: {input_path}", file=sys.stderr)
        return 2

    files = iter_input_files(input_path, args.recursive)
    if not files:
        print("No DXF files found.", file=sys.stderr)
        return 1

    ok = 0
    niet_verwerkt_csv = None
    if args.error_log:
        niet_verwerkt_csv = Path(args.error_log).with_name("niet_verwerkt.csv")
    else:
        niet_verwerkt_csv = output_dir.parent / "Doel" / "niet_verwerkt.csv"
    for dxf_file in files:
        try:
            out_path = output_dir / f"{dxf_file.stem}.xlsx"
            rows = list(entity_rows(dxf_file))
            if len(rows) < max(0, args.min_rows):
                raise RuntimeError(f"Too few rows extracted ({len(rows)} < {args.min_rows})")
            if args.strict_nbd and not has_nbd_signals(rows):
                raise RuntimeError("DXF parsed, but no expected NBD sign-spec markers were found")
            row_count = write_xlsx(rows, out_path)
            ok += 1
            print(f"[OK] {dxf_file.name} -> {out_path.name} ({row_count} rows)")
        except Exception as exc:
            print(f"[ERR] {dxf_file.name}: {exc}", file=sys.stderr)
            if args.error_log:
                append_error_log(
                    Path(args.error_log),
                    ErrorLogRow(
                        timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        source=str(dxf_file),
                        error=f"{exc.__class__.__name__}: {exc}",
                    ),
                )
            if niet_verwerkt_csv:
                append_niet_verwerkt_csv(
                    niet_verwerkt_csv,
                    type_value="dxf_extract",
                    source=str(dxf_file),
                    object_nummer=derive_object_nummer_from_stem(dxf_file.stem),
                    reason=f"{exc.__class__.__name__}: {exc}",
                )

    print(f"Done. {ok}/{len(files)} DXF files converted.")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
