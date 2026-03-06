#!/usr/bin/env python3
"""Multimetre kalibrasyon mastarı için .xlsx dosyası üretir (ek bağımlılık olmadan)."""

from argparse import ArgumentParser
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


def parse_ranges(raw_ranges: str) -> list[float]:
    ranges: list[float] = []
    for item in raw_ranges.split(","):
        cleaned = item.strip().replace("V", "").replace("A", "").replace("Ω", "")
        if cleaned:
            ranges.append(float(cleaned))
    if not ranges:
        raise ValueError("En az bir ölçüm aralığı girilmelidir.")
    return ranges


def build_rows(ranges: list[float]) -> list[dict[str, float]]:
    return [
        {
            "range": rng,
            "value_10": rng * 0.10,
            "value_50": rng * 0.50,
            "value_90": rng * 0.90,
        }
        for rng in ranges
    ]


def inline_str_cell(ref: str, value: str) -> str:
    return f'<c r="{ref}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'


def number_cell(ref: str, value: float, style_id: int = 0) -> str:
    style_attr = f' s="{style_id}"' if style_id else ""
    return f'<c r="{ref}"{style_attr}><v>{value:.3f}</v></c>'


def int_cell(ref: str, value: int) -> str:
    return f'<c r="{ref}"><v>{value}</v></c>'


def row_xml(row_idx: int, cells: list[str]) -> str:
    return f'<row r="{row_idx}">{"".join(cells)}</row>'


def sheet_xml(brand: str, model: str, rows: list[dict[str, float]]) -> str:
    xml_rows: list[str] = []

    xml_rows.append(
        row_xml(
            1,
            [inline_str_cell("A1", "MULTİMETRE KALİBRASYON MASTARI")],
        )
    )
    xml_rows.append(
        row_xml(
            2,
            [
                inline_str_cell("A2", "Marka"),
                inline_str_cell("B2", brand),
                inline_str_cell("D2", "Model"),
                inline_str_cell("E2", model),
            ],
        )
    )
    xml_rows.append(
        row_xml(
            3,
            [
                inline_str_cell("A3", "Tarih"),
                inline_str_cell("B3", datetime.now().strftime("%Y-%m-%d")),
            ],
        )
    )

    headers = ["No", "Ölçüm Aralığı", "%10 Noktası", "%50 Noktası", "%90 Noktası", "Birim"]
    xml_rows.append(
        row_xml(5, [inline_str_cell(f"{chr(65 + i)}5", header) for i, header in enumerate(headers)])
    )

    for idx, item in enumerate(rows, start=1):
        excel_row = idx + 5
        xml_rows.append(
            row_xml(
                excel_row,
                [
                    int_cell(f"A{excel_row}", idx),
                    number_cell(f"B{excel_row}", item["range"], style_id=1),
                    number_cell(f"C{excel_row}", item["value_10"], style_id=1),
                    number_cell(f"D{excel_row}", item["value_50"], style_id=1),
                    number_cell(f"E{excel_row}", item["value_90"], style_id=1),
                    inline_str_cell(f"F{excel_row}", "A/V/Ω"),
                ],
            )
        )

    return f"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <dimension ref=\"A1:F{max(6, len(rows) + 5)}\"/>
  <sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>
  <sheetFormatPr defaultRowHeight=\"15\"/>
  <cols>
    <col min=\"1\" max=\"6\" width=\"18\" customWidth=\"1\"/>
  </cols>
  <sheetData>
    {''.join(xml_rows)}
  </sheetData>
  <mergeCells count=\"1\"><mergeCell ref=\"A1:F1\"/></mergeCells>
</worksheet>
"""


def workbook_xml() -> str:
    return """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <sheets>
    <sheet name=\"Mastar\" sheetId=\"1\" r:id=\"rId1\"/>
  </sheets>
</workbook>
"""


def styles_xml() -> str:
    return """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
  <numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"0.000\"/></numFmts>
  <fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>
  <fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>
  <borders count=\"1\"><border/></borders>
  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>
  <cellXfs count=\"2\">
    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>
    <xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>
  </cellXfs>
  <cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>
</styleSheet>
"""


def content_types_xml() -> str:
    return """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>
</Types>
"""


def root_rels_xml() -> str:
    return """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>
</Relationships>
"""


def workbook_rels_xml() -> str:
    return """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>
</Relationships>
"""


def create_workbook(brand: str, model: str, ranges: list[float], output_file: Path) -> None:
    data_rows = build_rows(ranges)
    with ZipFile(output_file, "w", compression=ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml())
        zf.writestr("_rels/.rels", root_rels_xml())
        zf.writestr("xl/workbook.xml", workbook_xml())
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml())
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml(brand, model, data_rows))
        zf.writestr("xl/styles.xml", styles_xml())


def main() -> None:
    parser = ArgumentParser(description="Multimetre kalibrasyon mastarı için Excel dosyası oluşturur.")
    parser.add_argument("--brand", required=True, help="Multimetre markası")
    parser.add_argument("--model", required=True, help="Multimetre modeli")
    parser.add_argument(
        "--ranges",
        required=True,
        help="Virgülle ayrılmış ölçüm aralıkları. Örn: 0.2,2,20,200",
    )
    parser.add_argument("--output", default="multimetre_mastar.xlsx", help="Excel dosyası adı")

    args = parser.parse_args()
    create_workbook(args.brand, args.model, parse_ranges(args.ranges), Path(args.output))
    print(f"Excel dosyası oluşturuldu: {args.output}")


if __name__ == "__main__":
    main()
