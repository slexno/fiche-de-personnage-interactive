from __future__ import annotations

import math
import re
import uuid
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass
class WorkbookData:
    path: Path
    sheets: dict[str, list[dict]]
    headers: dict[str, list[str]]


class XlsxMini:
    @staticmethod
    def load(path: Path) -> WorkbookData:
        zf = zipfile.ZipFile(path)
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {r.attrib["Id"]: r.attrib["Target"] for r in rels}

        shared = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sst = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sst.findall(f"{{{NS_MAIN}}}si"):
                text = "".join((t.text or "") for t in si.findall(f".//{{{NS_MAIN}}}t"))
                shared.append(text)

        sheets: dict[str, list[dict]] = {}
        headers: dict[str, list[str]] = {}
        for sheet in wb.find(f"{{{NS_MAIN}}}sheets"):
            name = sheet.attrib["name"]
            rid = sheet.attrib[f"{{{NS_REL}}}id"]
            target = "xl/" + rid_to_target[rid]
            xml = ET.fromstring(zf.read(target))
            rows = xml.findall(f"{{{NS_MAIN}}}sheetData/{{{NS_MAIN}}}row")
            if not rows:
                sheets[name] = []
                headers[name] = []
                continue
            header_values = XlsxMini._row_values(rows[0], shared)
            headers[name] = header_values
            sheet_rows = []
            for row in rows[1:]:
                vals = XlsxMini._row_values(row, shared)
                if not any(v != "" for v in vals):
                    continue
                row_data = {header_values[i]: (vals[i] if i < len(vals) else "") for i in range(len(header_values)) if header_values[i] != ""}
                sheet_rows.append(row_data)
            sheets[name] = sheet_rows

        return WorkbookData(path=path, sheets=sheets, headers=headers)

    @staticmethod
    def _row_values(row, shared: list[str]) -> list[str]:
        cells = {}
        for c in row.findall(f"{{{NS_MAIN}}}c"):
            ref = c.attrib.get("r", "A1")
            col = re.match(r"([A-Z]+)", ref).group(1)
            idx = XlsxMini._col_to_idx(col)
            t = c.attrib.get("t")
            val = ""
            v = c.find(f"{{{NS_MAIN}}}v")
            if v is not None and v.text is not None:
                val = shared[int(v.text)] if t == "s" else v.text
            else:
                isel = c.find(f"{{{NS_MAIN}}}is/{{{NS_MAIN}}}t")
                if isel is not None and isel.text is not None:
                    val = isel.text
            cells[idx] = val
        if not cells:
            return []
        max_idx = max(cells)
        return [cells.get(i, "") for i in range(max_idx + 1)]

    @staticmethod
    def _col_to_idx(col: str) -> int:
        n = 0
        for ch in col:
            n = n * 26 + ord(ch) - 64
        return n - 1

    @staticmethod
    def _idx_to_col(idx: int) -> str:
        idx += 1
        out = ""
        while idx:
            idx, r = divmod(idx - 1, 26)
            out = chr(65 + r) + out
        return out

    @staticmethod
    def save(data: WorkbookData):
        sheets = list(data.sheets.keys())
        with zipfile.ZipFile(data.path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", XlsxMini._content_types(len(sheets)))
            zf.writestr("_rels/.rels", XlsxMini._root_rels())
            zf.writestr("xl/workbook.xml", XlsxMini._workbook_xml(sheets))
            zf.writestr("xl/_rels/workbook.xml.rels", XlsxMini._workbook_rels(len(sheets)))
            zf.writestr("xl/styles.xml", XlsxMini._styles())
            for i, name in enumerate(sheets, start=1):
                headers = data.headers.get(name, [])
                rows = data.sheets[name]
                zf.writestr(f"xl/worksheets/sheet{i}.xml", XlsxMini._sheet_xml(headers, rows))

    @staticmethod
    def _content_types(sheet_count: int) -> str:
        overrides = "".join(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(1, sheet_count + 1)
        )
        return (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f"{overrides}"
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            "</Types>"
        )

    @staticmethod
    def _root_rels() -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>"
        )

    @staticmethod
    def _workbook_xml(sheets: list[str]) -> str:
        body = "".join(
            f'<sheet name="{name}" sheetId="{i}" r:id="rId{i}"/>' for i, name in enumerate(sheets, start=1)
        )
        return (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f"<sheets>{body}</sheets></workbook>"
        )

    @staticmethod
    def _workbook_rels(sheet_count: int) -> str:
        sheet_rels = "".join(
            f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
            for i in range(1, sheet_count + 1)
        )
        return (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{sheet_rels}"
            f'<Relationship Id="rId{sheet_count+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            "</Relationships>"
        )

    @staticmethod
    def _styles() -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border/></borders>'
            '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
            '<cellXfs count="1"><xf/></cellXfs>'
            '</styleSheet>'
        )

    @staticmethod
    def _sheet_xml(headers: list[str], rows: list[dict]) -> str:
        all_rows = [headers] + [[str(r.get(h, "")) for h in headers] for r in rows]
        row_xml = []
        for r_idx, vals in enumerate(all_rows, start=1):
            cells = []
            for c_idx, val in enumerate(vals):
                if val == "":
                    continue
                ref = f"{XlsxMini._idx_to_col(c_idx)}{r_idx}"
                if re.fullmatch(r"-?\d+(\.\d+)?", val):
                    cells.append(f'<c r="{ref}"><v>{val}</v></c>')
                else:
                    safe = (
                        val.replace("&", "&amp;")
                        .replace("<", "&lt;")
                        .replace(">", "&gt;")
                    )
                    cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{safe}</t></is></c>')
            row_xml.append(f'<row r="{r_idx}">{"".join(cells)}</row>')
        return (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<sheetData>{"".join(row_xml)}</sheetData></worksheet>'
        )


class CharacterAppStore:
    def __init__(self, root: Path):
        self.char = XlsxMini.load(root / "caracteristique.xlsx")
        self.inv = XlsxMini.load(root / "inventaire.xlsx")
        self.shop = XlsxMini.load(root / "magasin.xlsx")
        self._normalize_inventory()

    def _normalize_inventory(self):
        for bucket in ["sac à dos", "coffre"]:
            for item in self.inv.sheets.get(bucket, []):
                item.setdefault("id", str(uuid.uuid4()))
                item.setdefault("type", "item")
                item.setdefault("equiped", "0")
        weapon_map = {w.get("Armes", "").strip().lower(): w for w in self.inv.sheets.get("armes", []) if w.get("Armes")}
        equip_map = {e.get("Equipement", "").strip().lower(): e for e in self.inv.sheets.get("equipement", []) if e.get("Equipement")}

        for bucket in ["sac à dos", "coffre"]:
            for item in self.inv.sheets.get(bucket, []):
                name = item.get("Objet", "").strip().lower()
                if name in weapon_map:
                    item["type"] = "arme"
                    row = weapon_map[name]
                    item["Range (ft)"] = row.get("Range (ft)", item.get("Range (ft)", ""))
                    item["Hit"] = row.get("Hit", item.get("Hit", ""))
                    item["Damage"] = row.get("Damage", item.get("Damage", ""))
                    item["description"] = item.get("description") or row.get("description", "")
                if name in equip_map:
                    item["type"] = "equipement"
                    row = equip_map[name]
                    item["bonus Armor class"] = row.get("bonus Armor class", item.get("bonus Armor class", "0"))
                    item["effet(optionel)"] = row.get("effet(optionel)", item.get("effet(optionel)", "ho le nul il a pas d'effets"))
                    item["description"] = item.get("description") or row.get("description", "")
        if not any(i.get("Objet", "").lower() == "crédits" for i in self.inv.sheets["sac à dos"]):
            self.inv.sheets["sac à dos"].append(
                {
                    "Objet": "Crédits",
                    "Valeur (en crédit)": "500",
                    "Poid (kg)": "0",
                    "poid unitaire (kg)": "0",
                    "Quantité": "1",
                    "Prix unitaire (en crédit)": "500",
                    "description": "Monnaie du personnage",
                    "id": str(uuid.uuid4()),
                    "type": "currency",
                    "equiped": "0",
                }
            )
        for sheet in ["sac à dos", "coffre"]:
            headers = self.inv.headers[sheet]
            for extra in ["id", "type", "equiped", "Range (ft)", "Hit", "Damage", "bonus Armor class", "effet(optionel)"]:
                if extra not in headers:
                    headers.append(extra)

    def build_state(self):
        stats = self._build_stats()
        inventory = self._build_inventory()
        shop = self.shop.sheets
        return {"stats": stats, "inventory": inventory, "shop": shop}

    def _build_stats(self):
        rows = self.char.sheets.get("Feuil1", [])
        stats = []
        skills = []
        for r in rows:
            if r.get("Statistiques"):
                score = int(float(r.get("Score", "10") or 10))
                bonus = math.floor((score - 10) / 2)
                stats.append({"name": r.get("Statistiques"), "score": score, "bonus": bonus})
            if r.get("Competence"):
                specialized = str(r.get("Spécialisation", "0")) in {"1", "true", "True"}
                base = int(float(r.get("Bonus", "0") or 0))
                skills.append(
                    {
                        "name": r.get("Competence"),
                        "mod": r.get("Modificateur", ""),
                        "bonus": base + (2 if specialized else 0),
                        "specialized": specialized,
                    }
                )
        dex_bonus = next((s["bonus"] for s in stats if s["name"].lower().startswith("dex")), 0)
        if self._bag_weight() > 50:
            dex_bonus -= 1
        ac_bonus = sum(int(float(i.get("bonus Armor class", "0") or 0)) for i in self._equipped("equipement"))
        armor_class = 9 + ac_bonus + dex_bonus
        return {"stats": stats, "skills": skills, "armor_class": armor_class}

    def _equipped(self, typ: str):
        return [i for i in self.inv.sheets["sac à dos"] if i.get("type") == typ and i.get("equiped") == "1"]

    def _bag_weight(self):
        return sum(float(i.get("Poid (kg)", "0") or 0) for i in self.inv.sheets["sac à dos"] if i.get("type") != "currency")

    def _build_inventory(self):
        bag = self.inv.sheets["sac à dos"]
        chest = self.inv.sheets["coffre"]
        weapons = [i for i in bag if i.get("type") == "arme"]
        equipments = [i for i in bag if i.get("type") == "equipement"]
        credits = self._credits()
        return {
            "bag": bag,
            "chest": chest,
            "weapons": weapons,
            "equipments": equipments,
            "bag_weight": self._bag_weight(),
            "overweight": self._bag_weight() > 50,
            "credits": credits,
        }

    def _credits(self):
        c = next((i for i in self.inv.sheets["sac à dos"] if i.get("type") == "currency"), None)
        if not c:
            return 0
        return float(c.get("Valeur (en crédit)", "0") or 0)

    def apply_action(self, payload: dict):
        action = payload.get("action")
        if action == "update_stat":
            self._update_stat(payload)
        elif action == "toggle_skill":
            self._toggle_skill(payload)
        elif action == "add_item":
            self._add_item(payload)
        elif action == "transfer_item":
            self._transfer(payload)
        elif action == "assign_type":
            self._assign_type(payload)
        elif action == "toggle_equip":
            self._toggle_equip(payload)
        elif action == "buy":
            self._buy(payload)
        elif action == "sell":
            self._sell(payload)
        elif action == "sort":
            self._sort(payload)
        elif action == "update_item":
            self._update_item(payload)

        self._sync_derived_tables()
        XlsxMini.save(self.char)
        XlsxMini.save(self.inv)
        return {"ok": True, "state": self.build_state()}

    def _find_stat_row(self, name: str):
        return next((r for r in self.char.sheets["Feuil1"] if r.get("Statistiques") == name), None)

    def _update_stat(self, payload):
        row = self._find_stat_row(payload["name"])
        if row:
            val = max(1, min(20, int(payload["score"])))
            row["Score"] = str(val)
            row["Bonus"] = str(math.floor((val - 10) / 2))

    def _toggle_skill(self, payload):
        for r in self.char.sheets["Feuil1"]:
            if r.get("Competence") == payload["name"]:
                r["Spécialisation"] = "1" if payload.get("specialized") else "0"

    def _add_item(self, payload):
        item = payload["item"]
        item.setdefault("id", str(uuid.uuid4()))
        item.setdefault("type", "item")
        item.setdefault("equiped", "0")
        qty = float(item.get("Quantité", 1) or 1)
        pu = float(item.get("Prix unitaire (en crédit)", 0) or 0)
        wt = float(item.get("poid unitaire (kg)", 0) or 0)
        item["Valeur (en crédit)"] = str(pu * qty)
        item["Poid (kg)"] = str(wt * qty)
        self.inv.sheets["sac à dos"].append(item)

    def _transfer(self, payload):
        src = self.inv.sheets[payload["from"]]
        dst = self.inv.sheets[payload["to"]]
        idx = next((i for i, it in enumerate(src) if it.get("id") == payload["id"]), None)
        if idx is None:
            return

        item = src[idx]
        if item.get("type") == "currency":
            return

        qty = max(1, int(float(payload.get("qty", 1) or 1)))
        stock = int(float(item.get("Quantité", "1") or 1))
        qty = min(qty, stock)

        if qty == stock:
            dst.append(src.pop(idx))
            return

        moved = dict(item)
        moved["id"] = str(uuid.uuid4())
        moved["Quantité"] = str(qty)

        left = stock - qty
        item["Quantité"] = str(left)

        for target in (item, moved):
            q = float(target.get("Quantité", 1) or 1)
            pu = float(target.get("Prix unitaire (en crédit)", 0) or 0)
            wu = float(target.get("poid unitaire (kg)", 0) or 0)
            if target.get("type") != "currency":
                target["Valeur (en crédit)"] = str(round(q * pu, 2))
            target["Poid (kg)"] = str(round(q * wu, 2))

        dst.append(moved)

    def _assign_type(self, payload):
        item = self._find_item(payload["id"])
        if not item:
            return
        typ = payload["type"]
        item["type"] = typ
        if typ == "arme":
            for key in ["Range (ft)", "Hit", "Damage"]:
                item[key] = str(payload.get(key, item.get(key, "")))
        if typ == "equipement":
            item["bonus Armor class"] = str(payload.get("bonus Armor class", item.get("bonus Armor class", 0)))
            item["effet(optionel)"] = payload.get("effet(optionel)", item.get("effet(optionel)", "ho le nul il a pas d'effets"))

    def _toggle_equip(self, payload):
        item = self._find_item(payload["id"])
        if not item:
            return
        typ = item.get("type")
        if payload.get("equiped"):
            if typ == "arme" and len(self._equipped("arme")) >= 4:
                return
            if typ == "equipement" and len(self._equipped("equipement")) >= 3:
                return
        item["equiped"] = "1" if payload.get("equiped") else "0"

    def _find_item(self, item_id: str):
        for sheet in ["sac à dos", "coffre"]:
            for it in self.inv.sheets[sheet]:
                if it.get("id") == item_id:
                    return it
        return None

    def _set_credits(self, value: float):
        c = next((i for i in self.inv.sheets["sac à dos"] if i.get("type") == "currency"), None)
        if c:
            c["Valeur (en crédit)"] = str(round(value, 2))
            c["Prix unitaire (en crédit)"] = c["Valeur (en crédit)"]

    def _buy(self, payload):
        sheet = payload["sheet"]
        name = payload["name"]
        qty = max(1, int(payload.get("qty", 1)))
        row = next((r for r in self.shop.sheets[sheet] if r.get("nom de l'objet") == name), None)
        if not row:
            return
        price = float(row.get("prix unitaire (crédit)", 0) or 0)
        total = price * qty
        credits = self._credits()
        if credits < total:
            return
        self._set_credits(credits - total)
        item = {
            "Objet": name,
            "Quantité": str(qty),
            "Prix unitaire (en crédit)": str(price),
            "description": row.get("description", ""),
            "poid unitaire (kg)": row.get("poid unitaire(kg)", "0"),
            "Range (ft)": row.get("Range (ft)", ""),
            "Hit": row.get("Hit", ""),
            "Damage": row.get("Damage", ""),
            "bonus Armor class": row.get("bonus armor class", "0"),
            "effet(optionel)": row.get("effet", "ho le nul il a pas d'effets"),
        }
        self._add_item({"item": item})

    def _sell(self, payload):
        item = self._find_item(payload["id"])
        if not item or item.get("type") == "currency":
            return
        qty = max(1, int(payload.get("qty", 1)))
        stock = int(float(item.get("Quantité", "1") or 1))
        qty = min(qty, stock)
        price = float(item.get("Prix unitaire (en crédit)", 0) or 0)
        self._set_credits(self._credits() + price * qty)
        left = stock - qty
        if left <= 0:
            for sheet in ["sac à dos", "coffre"]:
                self.inv.sheets[sheet] = [i for i in self.inv.sheets[sheet] if i.get("id") != item.get("id")]
        else:
            item["Quantité"] = str(left)
            item["Valeur (en crédit)"] = str(left * price)
            unit_w = float(item.get("poid unitaire (kg)", 0) or 0)
            item["Poid (kg)"] = str(left * unit_w)

    def _update_item(self, payload):
        item = self._find_item(payload["id"])
        if not item:
            return
        allowed = [
            "description",
            "effet(optionel)",
            "poid unitaire (kg)",
            "Prix unitaire (en crédit)",
            "Quantité",
            "Range (ft)",
            "Hit",
            "Damage",
            "bonus Armor class",
            "type",
        ]
        for key in allowed:
            if key in payload:
                item[key] = str(payload[key])

    def _sort(self, payload):
        key = payload["key"]
        source = self.inv.sheets[payload.get("source", "sac à dos")]
        if key == "alpha":
            source.sort(key=lambda x: x.get("Objet", "").lower())
        elif key == "prix":
            source.sort(key=lambda x: float(x.get("Prix unitaire (en crédit)", 0) or 0))
        elif key == "poids":
            source.sort(key=lambda x: float(x.get("Poid (kg)", 0) or 0))

    def _sync_derived_tables(self):
        for sheet in ["sac à dos", "coffre"]:
            for item in self.inv.sheets[sheet]:
                q = float(item.get("Quantité", 1) or 1)
                pu = float(item.get("Prix unitaire (en crédit)", 0) or 0)
                wu = float(item.get("poid unitaire (kg)", 0) or 0)
                if item.get("type") != "currency":
                    item["Valeur (en crédit)"] = str(round(q * pu, 2))
                item["Poid (kg)"] = str(round(q * wu, 2))

        self.inv.sheets["armes"] = [
            {
                "Armes": i.get("Objet", ""),
                "Range (ft)": i.get("Range (ft)", ""),
                "Hit": i.get("Hit", ""),
                "Damage": i.get("Damage", ""),
                "description": i.get("description", ""),
            }
            for i in self.inv.sheets["sac à dos"]
            if i.get("type") == "arme"
        ]
        self.inv.sheets["equipement"] = [
            {
                "Equipement": i.get("Objet", ""),
                "bonus Armor class": i.get("bonus Armor class", "0"),
                "effet(optionel)": i.get("effet(optionel)", "ho le nul il a pas d'effets"),
                "description": i.get("description", ""),
            }
            for i in self.inv.sheets["sac à dos"]
            if i.get("type") == "equipement"
        ]
        XlsxMini.save(self.shop)
