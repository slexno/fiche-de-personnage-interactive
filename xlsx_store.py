from __future__ import annotations

import math
import unicodedata
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
        self.root = root
        self.char = XlsxMini.load(root / "caracteristique.xlsx")
        self.inv = XlsxMini.load(root / "inventaire.xlsx")
        self.shop = XlsxMini.load(root / "magasin.xlsx")
        self._enrich_shop_images()
        self._normalize_inventory()

    @staticmethod
    def _slug(value: str) -> str:
        txt = unicodedata.normalize("NFKD", value or "").encode("ascii", "ignore").decode("ascii")
        txt = txt.lower()
        return " ".join("".join(ch if ch.isalnum() else " " for ch in txt).split())

    @staticmethod
    def _truthy(v) -> bool:
        return str(v).strip().lower() in {"1", "true", "yes", "oui", "x"}

    @staticmethod
    def _to_float(v, default=0.0):
        try:
            return float(v)
        except Exception:
            return default

    @staticmethod
    def _normalize_key(v) -> str:
        text = str(v or "").strip().lower()
        text = unicodedata.normalize("NFKD", text)
        text = "".join(ch for ch in text if not unicodedata.combining(ch))
        return re.sub(r"[^a-z]", "", text)

    @staticmethod
    def _canonical_stat_key(v) -> str:
        key = CharacterAppStore._normalize_key(v)
        aliases = {
            "for": "for",
            "force": "for",
            "str": "for",
            "strength": "for",
            "dex": "dex",
            "dexterite": "dex",
            "dexterity": "dex",
            "agi": "dex",
            "agilite": "dex",
            "agility": "dex",
            "con": "con",
            "constitution": "con",
            "int": "int",
            "intelligence": "int",
            "inteligence": "int",
            "sag": "sag",
            "sagesse": "sag",
            "wis": "sag",
            "wisdom": "sag",
            "cha": "cha",
            "charisme": "cha",
            "charisma": "cha",
        }
        return aliases.get(key, key)


    @staticmethod
    def _display_stat_name(v) -> str:
        canonical = CharacterAppStore._canonical_stat_key(v)
        labels = {
            "for": "Force",
            "dex": "Dextérité",
            "con": "Constitution",
            "int": "Intelligence",
            "sag": "Sagesse",
            "cha": "Charisme",
        }
        raw = str(v or "").strip()
        return labels.get(canonical, raw)

    def _find_stat_bonus(self, effective_map: dict[str, float], stat_name: str) -> float:
        target = self._canonical_stat_key(stat_name)
        if not target:
            return 0
        if target in effective_map:
            return effective_map[target]
        short = target[:3]
        return next((v for k, v in effective_map.items() if k.startswith(short)), 0)

    def _enrich_shop_images(self):
        image_dir = self.root / "image"
        files = [f for f in image_dir.iterdir() if f.is_file()] if image_dir.exists() else []
        by_slug = {self._slug(f.stem): f"image/{f.name}" for f in files}

        for sheet_rows in self.shop.sheets.values():
            for row in sheet_rows:
                raw = str(row.get("image", "") or "").strip()
                resolved = ""
                if raw and raw not in {"#VALUE!", "#N/A"}:
                    resolved = raw if raw.startswith("http") or raw.startswith("image/") else f"image/{raw}"
                if not resolved:
                    resolved = by_slug.get(self._slug(row.get("nom de l'objet", "")), "")
                row["resolved_image"] = resolved
                row["resolved_hit_modifier"] = (
                    row.get("Modificateur")
                    or row.get("Hit Mod")
                    or row.get("Hit Stat")
                    or (row.get("Hit") if not re.fullmatch(r"-?\d+(\.\d+)?", str(row.get("Hit", "")).strip()) else "")
                    or row.get("modificateur")
                    or ""
                )

    def _normalize_inventory(self):
        for bucket in ["sac à dos", "coffre"]:
            for item in self.inv.sheets.get(bucket, []):
                item.setdefault("id", str(uuid.uuid4()))
                item.setdefault("type", "item")
                item.setdefault("equiped", "0")

        if not any(i.get("Objet", "").lower() == "crédits" for i in self.inv.sheets["sac à dos"]):
            self.inv.sheets["sac à dos"].append({
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
            })

        for sheet in ["sac à dos", "coffre"]:
            headers = self.inv.headers[sheet]
            for extra in ["id", "type", "equiped", "Range (ft)", "Hit", "Damage", "Hit Stat", "Hit Specialized", "bonus Armor class", "effet(optionel)"]:
                if extra not in headers:
                    headers.append(extra)
        if "Expertise" not in self.char.headers["Feuil1"]:
            self.char.headers["Feuil1"].append("Expertise")

    def _bag_weight(self):
        return sum(self._to_float(i.get("Poid (kg)", 0)) for i in self.inv.sheets["sac à dos"] if i.get("type") != "currency")

    def _credits(self):
        c = next((i for i in self.inv.sheets["sac à dos"] if i.get("type") == "currency"), None)
        return self._to_float(c.get("Valeur (en crédit)", 0)) if c else 0.0

    def _set_credits(self, value: float):
        c = next((i for i in self.inv.sheets["sac à dos"] if i.get("type") == "currency"), None)
        if c:
            c["Valeur (en crédit)"] = str(round(value, 2))
            c["Prix unitaire (en crédit)"] = c["Valeur (en crédit)"]

    def _compute_stats_context(self):
        stats = []
        by_name = {}
        for r in self.char.sheets.get("Feuil1", []):
            if r.get("Statistiques"):
                score = int(self._to_float(r.get("Score", 10), 10))
                raw_bonus = math.floor((score - 10) / 2)
                name = str(r.get("Statistiques"))
                key = self._canonical_stat_key(name)
                by_name[key] = {"score": score, "raw_bonus": raw_bonus}
                stats.append({"name": name, "score": score, "raw_bonus": raw_bonus, "bonus": raw_bonus})

        force_bonus = by_name.get("for", {}).get("raw_bonus", 0)
        max_carry = 50 + 10 * force_bonus
        overweight = max(0.0, self._bag_weight() - max_carry)
        dex_penalty = math.ceil(overweight / 10) if overweight > 0 else 0

        for st in stats:
            if self._canonical_stat_key(st["name"]) == "dex":
                st["bonus"] = st["raw_bonus"] - dex_penalty

        effective = {self._canonical_stat_key(st["name"]): st["bonus"] for st in stats}
        return stats, effective, max_carry, dex_penalty

    def _weapon_hit_display(self, item: dict, effective_map: dict[str, float]) -> str:
        raw_hit = str(item.get("Hit", "") or "").strip()
        base = self._to_float(raw_hit, 0) if re.fullmatch(r"-?\d+(\.\d+)?", raw_hit) else 0
        mod_name = str(item.get("Hit Stat", "") or item.get("Modificateur", "") or (raw_hit if base == 0 else ""))
        stat_bonus = self._find_stat_bonus(effective_map, mod_name)
        spec = 2 if self._truthy(item.get("Hit Specialized", 0)) else 0
        return str(int(base + stat_bonus + spec))

    def _build_stats(self):
        stats, effective, max_carry, dex_penalty = self._compute_stats_context()
        skills = []
        for r in self.char.sheets.get("Feuil1", []):
            if not r.get("Competence"):
                continue
            specialized = self._truthy(r.get("Spécialisation", "0"))
            expertise = self._truthy(r.get("Expertise", "0"))
            mod = str(r.get("Modificateur", ""))
            stat_mod_bonus = self._find_stat_bonus(effective, mod)
            skills.append({
                "name": r.get("Competence"),
                "mod": self._display_stat_name(mod) if mod else mod,
                "bonus": stat_mod_bonus + (2 if specialized else 0) + (2 if expertise else 0),
                "specialized": specialized,
                "expertise": expertise,
            })

        dex_bonus = next((s["bonus"] for s in stats if self._canonical_stat_key(s["name"]) == "dex"), 0)
        ac_bonus = sum(int(self._to_float(i.get("bonus Armor class", "0"), 0)) for i in self.inv.sheets["sac à dos"] if i.get("type") == "equipement" and i.get("equiped") == "1")
        armor_class = 9 + ac_bonus + dex_bonus
        return {
            "stats": [{"name": self._display_stat_name(s["name"]), "score": s["score"], "bonus": s["bonus"]} for s in stats],
            "skills": skills,
            "armor_class": armor_class,
            "max_carry": max_carry,
            "dex_penalty": dex_penalty,
        }

    def _item_stack_key(self, i: dict):
        fields = ["Objet", "type", "Prix unitaire (en crédit)", "poid unitaire (kg)"]
        return tuple(str(i.get(f, "")).strip().lower() for f in fields)

    def _stack_into(self, sheet: str, item: dict):
        if item.get("type") == "currency":
            self.inv.sheets[sheet].append(item)
            return
        key = self._item_stack_key(item)
        for existing in self.inv.sheets[sheet]:
            if existing.get("type") == "currency":
                continue
            if self._item_stack_key(existing) == key:
                q = self._to_float(existing.get("Quantité", 0)) + self._to_float(item.get("Quantité", 0))
                existing["Quantité"] = str(q)
                return
        self.inv.sheets[sheet].append(item)

    def _build_inventory(self):
        bag = [i for i in self.inv.sheets["sac à dos"] if i.get("type") != "currency"]
        chest = self.inv.sheets["coffre"]
        _, effective, max_carry, dex_penalty = self._compute_stats_context()
        weapons = []
        for i in bag:
            if i.get("type") == "arme":
                w = dict(i)
                w["display_hit"] = self._weapon_hit_display(i, effective)
                weapons.append(w)
        equipments = [i for i in bag if i.get("type") == "equipement"]
        return {
            "bag": bag,
            "chest": chest,
            "weapons": weapons,
            "equipments": equipments,
            "bag_weight": self._bag_weight(),
            "max_carry": max_carry,
            "overweight": self._bag_weight() > max_carry,
            "dex_penalty": dex_penalty,
            "credits": self._credits(),
        }

    def build_state(self):
        return {"stats": self._build_stats(), "inventory": self._build_inventory(), "shop": self.shop.sheets}

    def apply_action(self, payload: dict):
        action = payload.get("action")
        feedback = {"ok": True}
        if action == "update_stat": self._update_stat(payload)
        elif action == "toggle_skill": self._toggle_skill(payload)
        elif action == "toggle_expertise": self._toggle_expertise(payload)
        elif action == "add_item": feedback = self._add_item(payload)
        elif action == "transfer_item": self._transfer(payload)
        elif action == "assign_type": self._assign_type(payload)
        elif action == "toggle_equip": self._toggle_equip(payload)
        elif action == "buy": feedback = self._buy(payload)
        elif action == "sell": self._sell(payload)
        elif action == "sort": self._sort(payload)
        elif action == "update_item": self._update_item(payload)
        elif action == "update_credits": self._set_credits(self._to_float(payload.get("credits", 0)))

        self._sync_derived_tables()
        XlsxMini.save(self.char)
        XlsxMini.save(self.inv)
        return {**feedback, "state": self.build_state()}

    def _update_stat(self, payload):
        for r in self.char.sheets["Feuil1"]:
            if self._canonical_stat_key(r.get("Statistiques")) == self._canonical_stat_key(payload.get("name")):
                val = max(1, min(20, int(float(payload["score"]))))
                r["Score"] = str(val)
                r["Bonus"] = str(math.floor((val - 10) / 2))

    def _toggle_skill(self, payload):
        for r in self.char.sheets["Feuil1"]:
            if r.get("Competence") == payload["name"]:
                r["Spécialisation"] = "1" if payload.get("specialized") else "0"

    def _toggle_expertise(self, payload):
        for r in self.char.sheets["Feuil1"]:
            if r.get("Competence") == payload["name"]:
                r["Expertise"] = "1" if payload.get("expertise") else "0"

    def _add_item(self, payload):
        item = dict(payload.get("item", {}))
        name = str(item.get("Objet", "")).strip()
        raw_price = str(item.get("Prix unitaire (en crédit)", "")).strip()
        if not name:
            return {"ok": False, "error": "Le nom de l'objet est obligatoire."}
        if raw_price == "" or not re.fullmatch(r"-?\d+(\.\d+)?", raw_price):
            return {"ok": False, "error": "Le prix unitaire est obligatoire et doit être un nombre."}

        item["Objet"] = name
        item["Prix unitaire (en crédit)"] = raw_price
        item.setdefault("id", str(uuid.uuid4()))
        item.setdefault("equiped", "0")
        if item.get("type") not in {"arme", "equipement", "item", "currency"}:
            if str(item.get("Range (ft)", "")).strip() or str(item.get("Hit", "")).strip() or str(item.get("Damage", "")).strip():
                item["type"] = "arme"
            elif str(item.get("bonus Armor class", "")).strip() or str(item.get("effet(optionel)", "")).strip():
                item["type"] = "equipement"
            else:
                item["type"] = "item"
        self._stack_into("sac à dos", item)
        return {"ok": True}

    def _transfer(self, payload):
        src = self.inv.sheets[payload["from"]]
        dst_name = payload["to"]
        idx = next((i for i, it in enumerate(src) if it.get("id") == payload["id"]), None)
        if idx is None:
            return
        item = src[idx]
        if item.get("type") == "currency":
            return
        qty = max(1, int(self._to_float(payload.get("qty", 1), 1)))
        stock = int(self._to_float(item.get("Quantité", 1), 1))
        qty = min(qty, stock)
        if qty == stock:
            self._stack_into(dst_name, src.pop(idx))
            return
        moved = dict(item)
        moved["id"] = str(uuid.uuid4())
        moved["Quantité"] = str(qty)
        item["Quantité"] = str(stock - qty)
        self._stack_into(dst_name, moved)

    def _assign_type(self, payload):
        item = self._find_item(payload["id"])
        if not item:
            return
        typ = payload["type"]
        item["type"] = typ
        if typ == "arme":
            for key in ["Range (ft)", "Hit", "Damage", "Hit Stat", "Hit Specialized"]:
                if key in payload:
                    item[key] = str(payload.get(key, item.get(key, "")))
        if typ == "equipement":
            if "bonus Armor class" in payload:
                item["bonus Armor class"] = str(payload.get("bonus Armor class", item.get("bonus Armor class", 0)))
            if "effet(optionel)" in payload:
                item["effet(optionel)"] = payload.get("effet(optionel)", item.get("effet(optionel)", "ho le nul il a pas d'effets"))

    def _toggle_equip(self, payload):
        item = self._find_item(payload["id"])
        if not item:
            return
        typ = item.get("type")
        if payload.get("equiped"):
            if typ == "arme" and len([i for i in self.inv.sheets["sac à dos"] if i.get("type") == "arme" and i.get("equiped") == "1"]) >= 4:
                return
            if typ == "equipement" and len([i for i in self.inv.sheets["sac à dos"] if i.get("type") == "equipement" and i.get("equiped") == "1"]) >= 3:
                return
        item["equiped"] = "1" if payload.get("equiped") else "0"

    def _find_item(self, item_id: str):
        for sheet in ["sac à dos", "coffre"]:
            for it in self.inv.sheets[sheet]:
                if it.get("id") == item_id:
                    return it
        return None

    def _buy(self, payload):
        sheet, name = payload["sheet"], payload["name"]
        qty = max(1, int(payload.get("qty", 1)))
        row = next((r for r in self.shop.sheets[sheet] if r.get("nom de l'objet") == name), None)
        if not row:
            return {"ok": False, "error": "Objet introuvable"}
        price = self._to_float(row.get("prix unitaire (crédit)", 0))
        total = price * qty
        credits = self._credits()
        if credits < total:
            return {"ok": False, "error": "Fonds insuffisants", "missing_credits": round(total - credits, 2)}
        raw_hit = str(row.get("Hit", "") or "").strip()
        hit_is_number = bool(re.fullmatch(r"-?\d+(\.\d+)?", raw_hit))
        hit_stat = (
            row.get("resolved_hit_modifier")
            or row.get("Hit Stat")
            or ("" if hit_is_number else raw_hit)
        )
        self._set_credits(credits - total)
        self._add_item({"item": {
            "Objet": name,
            "Quantité": str(qty),
            "Prix unitaire (en crédit)": str(price),
            "description": row.get("description", ""),
            "poid unitaire (kg)": row.get("poid unitaire(kg)", "0"),
            "Range (ft)": row.get("Range (ft)", ""),
            "Hit": raw_hit if hit_is_number else "0",
            "Damage": row.get("Damage", ""),
            "Hit Stat": hit_stat,
            "bonus Armor class": row.get("bonus armor class", "0"),
            "effet(optionel)": row.get("effet", ""),
        }})
        return {"ok": True}

    def _sell(self, payload):
        item = self._find_item(payload["id"])
        if not item or item.get("type") == "currency":
            return
        qty = max(1, int(payload.get("qty", 1)))
        stock = int(self._to_float(item.get("Quantité", 1), 1))
        qty = min(qty, stock)
        self._set_credits(self._credits() + self._to_float(item.get("Prix unitaire (en crédit)", 0)) * qty)
        left = stock - qty
        if left <= 0:
            for sheet in ["sac à dos", "coffre"]:
                self.inv.sheets[sheet] = [i for i in self.inv.sheets[sheet] if i.get("id") != item.get("id")]
        else:
            item["Quantité"] = str(left)

    def _update_item(self, payload):
        item = self._find_item(payload["id"])
        if not item:
            return
        allowed = ["description", "effet(optionel)", "poid unitaire (kg)", "Prix unitaire (en crédit)", "Quantité", "Range (ft)", "Hit", "Damage", "Hit Stat", "Hit Specialized", "bonus Armor class", "type"]
        for key in allowed:
            if key in payload:
                item[key] = str(payload[key])

    def _sort(self, payload):
        key = payload["key"]
        source = self.inv.sheets[payload.get("source", "sac à dos")]
        if key == "alpha":
            source.sort(key=lambda x: x.get("Objet", "").lower())
        elif key == "prix":
            source.sort(key=lambda x: self._to_float(
                x.get("Valeur (en crédit)", self._to_float(x.get("Prix unitaire (en crédit)", 0)) * self._to_float(x.get("Quantité", 1), 1))
            ))
        elif key == "poids":
            source.sort(key=lambda x: self._to_float(
                x.get("Poid (kg)", self._to_float(x.get("poid unitaire (kg)", 0)) * self._to_float(x.get("Quantité", 1), 1))
            ))

    def _sync_derived_tables(self):
        for sheet in ["sac à dos", "coffre"]:
            for item in self.inv.sheets[sheet]:
                q = self._to_float(item.get("Quantité", 1), 1)
                pu = self._to_float(item.get("Prix unitaire (en crédit)", 0), 0)
                wu = self._to_float(item.get("poid unitaire (kg)", 0), 0)
                if item.get("type") != "currency":
                    item["Valeur (en crédit)"] = str(round(q * pu, 2))
                item["Poid (kg)"] = str(round(q * wu, 2))

        _, effective, _, _ = self._compute_stats_context()
        self.inv.sheets["armes"] = [
            {"Armes": i.get("Objet", ""), "Range (ft)": i.get("Range (ft)", ""), "Hit": self._weapon_hit_display(i, effective), "Damage": i.get("Damage", ""), "description": i.get("description", "")}
            for i in self.inv.sheets["sac à dos"] if i.get("type") == "arme"
        ]
        self.inv.sheets["equipement"] = [
            {"Equipement": i.get("Objet", ""), "bonus Armor class": i.get("bonus Armor class", "0"), "effet(optionel)": i.get("effet(optionel)", "ho le nul il a pas d'effets"), "description": i.get("description", "")}
            for i in self.inv.sheets["sac à dos"] if i.get("type") == "equipement"
        ]
        XlsxMini.save(self.shop)
