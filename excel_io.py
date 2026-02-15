import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)


def _col_to_idx(col: str) -> int:
    total = 0
    for c in col:
        total = total * 26 + ord(c) - 64
    return total - 1


def _idx_to_col(idx: int) -> str:
    idx += 1
    result = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result


def _cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    typ = cell.attrib.get("t")
    if typ == "s":
        v = cell.find(f"{{{NS_MAIN}}}v")
        return "" if v is None else shared_strings[int(v.text)]
    if typ == "inlineStr":
        t = cell.find(f"{{{NS_MAIN}}}is/{{{NS_MAIN}}}t")
        return "" if t is None else (t.text or "")
    v = cell.find(f"{{{NS_MAIN}}}v")
    return "" if v is None else (v.text or "")


@dataclass
class WorkbookData:
    sheet_paths: dict[str, str]
    sheet_xml: dict[str, ET.Element]


def load_sheet_rows(path: str) -> WorkbookData:
    with zipfile.ZipFile(path, "r") as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

        rel_map = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels.findall(f"{{{NS_PKG_REL}}}Relationship")
        }

        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sst = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sst.findall(f"{{{NS_MAIN}}}si"):
                shared_strings.append(
                    "".join(t.text or "" for t in si.findall(f".//{{{NS_MAIN}}}t"))
                )

        sheet_paths: dict[str, str] = {}
        sheet_xml: dict[str, ET.Element] = {}
        for sheet in workbook.findall(f"{{{NS_MAIN}}}sheets/{{{NS_MAIN}}}sheet"):
            rid = sheet.attrib[f"{{{NS_REL}}}id"]
            target = rel_map[rid]
            full_path = f"xl/{target}"
            sheet_name = sheet.attrib["name"]
            sheet_paths[sheet_name] = full_path
            sheet_xml[sheet_name] = ET.fromstring(zf.read(full_path))

        rows_by_sheet = {}
        for sheet_name, xml_root in sheet_xml.items():
            rows = []
            for row in xml_root.findall(f"{{{NS_MAIN}}}sheetData/{{{NS_MAIN}}}row"):
                parsed = {}
                for cell in row.findall(f"{{{NS_MAIN}}}c"):
                    ref = cell.attrib.get("r", "")
                    m = re.match(r"([A-Z]+)(\d+)", ref)
                    if not m:
                        continue
                    parsed[_col_to_idx(m.group(1))] = _cell_value(cell, shared_strings)
                if parsed:
                    max_col = max(parsed)
                    values = [""] * (max_col + 1)
                    for i, val in parsed.items():
                        values[i] = val
                    rows.append(values)
            rows_by_sheet[sheet_name] = rows

        wb = WorkbookData(sheet_paths=sheet_paths, sheet_xml=sheet_xml)
        wb.rows = rows_by_sheet
        return wb


def save_sheet_rows(path: str, workbook_data: WorkbookData, updates: dict[str, list[list]]):
    for sheet_name, rows in updates.items():
        root = workbook_data.sheet_xml[sheet_name]
        sheet_data = root.find(f"{{{NS_MAIN}}}sheetData")
        if sheet_data is None:
            sheet_data = ET.SubElement(root, f"{{{NS_MAIN}}}sheetData")
        sheet_data.clear()

        for r_idx, row_values in enumerate(rows, 1):
            row_elem = ET.SubElement(sheet_data, f"{{{NS_MAIN}}}row", {"r": str(r_idx)})
            for c_idx, value in enumerate(row_values):
                if value in (None, ""):
                    continue
                ref = f"{_idx_to_col(c_idx)}{r_idx}"
                svalue = str(value)
                is_num = re.fullmatch(r"-?\d+(\.\d+)?", svalue) is not None
                cell = ET.SubElement(row_elem, f"{{{NS_MAIN}}}c", {"r": ref})
                if is_num:
                    v = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
                    v.text = svalue
                else:
                    cell.attrib["t"] = "inlineStr"
                    is_elem = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
                    t = ET.SubElement(is_elem, f"{{{NS_MAIN}}}t")
                    t.text = svalue

        dim = root.find(f"{{{NS_MAIN}}}dimension")
        if dim is not None:
            max_cols = max((len(r) for r in rows), default=1)
            max_rows = max(len(rows), 1)
            dim.attrib["ref"] = f"A1:{_idx_to_col(max_cols - 1)}{max_rows}"

    with zipfile.ZipFile(path, "r") as src:
        existing = {name: src.read(name) for name in src.namelist()}

    for sheet_name, rows in updates.items():
        _ = rows
        sheet_path = workbook_data.sheet_paths[sheet_name]
        existing[sheet_path] = ET.tostring(
            workbook_data.sheet_xml[sheet_name], encoding="utf-8", xml_declaration=True
        )

    existing.pop("xl/sharedStrings.xml", None)

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as dst:
        for name, content in existing.items():
            dst.writestr(name, content)
