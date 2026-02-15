import json
from copy import deepcopy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from excel_io import load_sheet_rows, save_sheet_rows

BASE = Path(__file__).parent

STAT_MODIFIERS = {"Force": "force", "Dextérité": "dex", "Constitution": "con", "Intéligence": "int", "Sagesse": "wis", "Charisme": "cha"}
SKILL_MOD = {"Force": "force", "Dextérité": "dex", "Constitution": "con", "Intéligence": "int", "Sagesse": "wis", "Charisme": "cha"}


class DataStore:
    def __init__(self):
        self.stats_xlsx = str(BASE / "caracteristique.xlsx")
        self.inventory_xlsx = str(BASE / "inventaire.xlsx")
        self.shop_xlsx = str(BASE / "magasin.xlsx")
        self.load_all()

    def load_all(self):
        self.load_stats()
        self.load_inventory()
        self.load_shop()

    def load_stats(self):
        wb = load_sheet_rows(self.stats_xlsx)
        rows = wb.rows["Feuil1"]
        self.stats_wb = wb
        self.stats = {}
        self.skills = []
        for row in rows[1:]:
            if len(row) >= 2 and row[0]:
                self.stats[row[0]] = max(0, min(20, int(float(row[1] or 0))))
            if len(row) >= 7 and row[3]:
                self.skills.append({"name": row[3], "modifier": row[4], "specialized": str(row[6]) in {"1", "True", "true"}})

    def load_inventory(self):
        wb = load_sheet_rows(self.inventory_xlsx)
        self.inventory_wb = wb

        def parse_items(sheet):
            rows = wb.rows[sheet]
            out = []
            for row in rows[1:]:
                if not row or not row[0]:
                    continue
                out.append({
                    "name": row[0],
                    "total_value": float(row[1] or 0),
                    "total_weight": float(row[2] or 0),
                    "unit_weight": float(row[3] or 0),
                    "qty": int(float(row[4] or 0)),
                    "unit_price": float(row[5] or 0),
                    "description": row[6] if len(row) > 6 else "",
                    "type": "item",
                })
            return out

        self.bag = parse_items("sac à dos")
        self.chest = parse_items("coffre")

        self.weapons = []
        for row in wb.rows["armes"][1:]:
            if row and row[0]:
                self.weapons.append({"name": row[0], "range": row[1] if len(row) > 1 else "", "hit": row[2] if len(row) > 2 else "", "damage": row[3] if len(row) > 3 else "", "description": row[4] if len(row) > 4 else "", "equipped": False})

        self.equipments = []
        for row in wb.rows["equipement"][1:]:
            if row and row[0]:
                self.equipments.append({"name": row[0], "ac_bonus": float(row[1] or 0) if len(row) > 1 else 0, "effect": row[2] if len(row) > 2 else "", "description": row[3] if len(row) > 3 else "", "equipped": False})

    def load_shop(self):
        wb = load_sheet_rows(self.shop_xlsx)
        self.shop_wb = wb
        self.shop = {}
        for sheet, rows in wb.rows.items():
            if not rows:
                continue
            headers = rows[0]
            items = []
            for row in rows[1:]:
                if not row or not row[0]:
                    continue
                item = {}
                for i, h in enumerate(headers):
                    if h:
                        item[h] = row[i] if i < len(row) else ""
                items.append(item)
            self.shop[sheet] = items

    def _ability_mod(self, score):
        return (score - 10) // 2

    def _dex_penalty(self):
        return -1 if self.total_bag_weight() > 50 else 0

    def stat_modifiers(self):
        out = {}
        for stat, score in self.stats.items():
            mod = self._ability_mod(score)
            if stat == "Dextérité":
                mod += self._dex_penalty()
            out[stat] = mod
        return out

    def skill_bonus(self, skill):
        mods = self.stat_modifiers()
        base = mods.get(skill["modifier"], 0)
        return base + (2 if skill["specialized"] else 0)

    def total_bag_weight(self):
        return sum(i["total_weight"] for i in self.bag)

    def credits(self):
        for i in self.bag:
            if i["name"].lower() == "crédit":
                return i
        credit = {"name": "Crédit", "total_value": 0, "total_weight": 0, "unit_weight": 0, "qty": 0, "unit_price": 1, "description": "Crédits", "type": "item"}
        self.bag.append(credit)
        return credit

    def _sync_rows(self, items):
        rows = [["Objet", "Valeur (en crédit)", "Poid (kg)", "poid unitaire (kg)", "Quantité", "Prix unitaire (en crédit)", "description"]]
        for i in items:
            i["total_value"] = round(i["unit_price"] * i["qty"], 2)
            i["total_weight"] = round(i["unit_weight"] * i["qty"], 2)
            rows.append([i["name"], i["total_value"], i["total_weight"], i["unit_weight"], i["qty"], i["unit_price"], i.get("description", "")])
        return rows

    def save_inventory(self):
        updates = {
            "sac à dos": self._sync_rows(self.bag),
            "coffre": self._sync_rows(self.chest),
            "armes": [["Armes", "Range (ft)", "Hit", "Damage", "description"]] + [[w["name"], w["range"], w["hit"], w["damage"], w["description"]] for w in self.weapons],
            "equipement": [["Equipement", "bonus Armor class", "effet(optionel)", "description"]] + [[e["name"], e["ac_bonus"], e["effect"] or "ho le nul il a pas d'effets", e["description"]] for e in self.equipments],
        }
        save_sheet_rows(self.inventory_xlsx, self.inventory_wb, updates)

    def save_stats(self):
        rows = [["Statistiques", "Score", "Bonus", "Competence", "Modificateur", "Bonus", "Spécialisation"]]
        stats_items = list(self.stats.items())
        max_rows = max(len(stats_items), len(self.skills))
        mods = self.stat_modifiers()
        for idx in range(max_rows):
            row = ["", "", "", "", "", "", ""]
            if idx < len(stats_items):
                name, score = stats_items[idx]
                row[0], row[1], row[2] = name, score, mods.get(name, 0)
            if idx < len(self.skills):
                sk = self.skills[idx]
                row[3], row[4], row[5], row[6] = sk["name"], sk["modifier"], self.skill_bonus(sk), 1 if sk["specialized"] else 0
            rows.append(row)
        save_sheet_rows(self.stats_xlsx, self.stats_wb, {"Feuil1": rows})

    def _find(self, where, name):
        items = self.bag if where == "bag" else self.chest
        for item in items:
            if item["name"] == name:
                return item
        return None

    def add_item(self, target, payload):
        items = self.bag if target == "bag" else self.chest
        existing = self._find(target, payload["name"])
        if existing:
            existing["qty"] += int(payload["qty"])
        else:
            payload["qty"] = int(payload["qty"])
            payload["unit_weight"] = float(payload["unit_weight"])
            payload["unit_price"] = float(payload["unit_price"])
            payload["description"] = payload.get("description", "")
            payload["type"] = payload.get("type", "item")
            items.append(payload)
        self.save_inventory()

    def transfer(self, source, name, qty):
        src = self.bag if source == "bag" else self.chest
        dst = self.chest if source == "bag" else self.bag
        item = next((i for i in src if i["name"] == name), None)
        if not item:
            return
        qty = min(item["qty"], int(qty))
        item["qty"] -= qty
        if item["qty"] <= 0:
            src.remove(item)
        target = next((i for i in dst if i["name"] == name), None)
        if target:
            target["qty"] += qty
        else:
            copied = deepcopy(item)
            copied["qty"] = qty
            dst.append(copied)
        self.save_inventory()

    def set_stat(self, stat, value):
        self.stats[stat] = max(0, min(20, int(value)))
        self.save_stats()

    def toggle_skill(self, name, specialized):
        for sk in self.skills:
            if sk["name"] == name:
                sk["specialized"] = bool(specialized)
        self.save_stats()

    def assign_weapon(self, payload):
        w = next((x for x in self.weapons if x["name"] == payload["name"]), None)
        if not w:
            self.weapons.append({**payload, "equipped": False})
        else:
            w.update(payload)
        self.save_inventory()

    def assign_equipment(self, payload):
        e = next((x for x in self.equipments if x["name"] == payload["name"]), None)
        if not e:
            self.equipments.append({**payload, "equipped": False})
        else:
            e.update(payload)
        self.save_inventory()

    def toggle_equipped(self, kind, name, equipped):
        items = self.weapons if kind == "weapon" else self.equipments
        limit = 4 if kind == "weapon" else 3
        current = sum(1 for i in items if i.get("equipped"))
        for item in items:
            if item["name"] == name:
                if equipped and not item.get("equipped") and current >= limit:
                    return False
                item["equipped"] = equipped
        return True

    def buy(self, category, name, qty):
        qty = int(qty)
        item = next((i for i in self.shop.get(category, []) if i.get("nom de l'objet") == name), None)
        if not item:
            return False, "Objet introuvable"
        price = float(item.get("prix unitaire (crédit)", 0))
        cost = price * qty
        credit = self.credits()
        if credit["qty"] < cost:
            return False, "Pas assez de crédits"
        credit["qty"] -= cost
        self.add_item("bag", {"name": name, "unit_weight": float(item.get("poid unitaire(kg)") or 0), "unit_price": price, "qty": qty, "description": item.get("description", "")})
        return True, "Achat validé"

    def sell(self, source, name, qty):
        item = self._find(source, name)
        if not item:
            return
        qty = min(int(qty), item["qty"])
        item["qty"] -= qty
        gains = qty * item["unit_price"]
        if item["qty"] <= 0:
            (self.bag if source == "bag" else self.chest).remove(item)
        self.credits()["qty"] += gains
        self.save_inventory()

    def snapshot(self):
        mods = self.stat_modifiers()
        ac_bonus = sum(e["ac_bonus"] for e in self.equipments if e.get("equipped"))
        ac = 9 + ac_bonus + mods.get("Dextérité", 0)
        return {
            "stats": [{"name": k, "score": v, "bonus": mods[k]} for k, v in self.stats.items()],
            "skills": [{"name": s["name"], "modifier": s["modifier"], "specialized": s["specialized"], "bonus": self.skill_bonus(s)} for s in self.skills],
            "armor_class": ac,
            "weight_warning": self.total_bag_weight() > 50,
            "bag": self.bag,
            "chest": self.chest,
            "weapons": self.weapons,
            "equipments": self.equipments,
            "shop": self.shop,
            "credits": self.credits()["qty"],
        }


STORE = DataStore()


class Handler(BaseHTTPRequestHandler):
    def _send_json(self, payload, code=200):
        body = json.dumps(payload).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path, ctype):
        content = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            return self._send_file(BASE / "web" / "index.html", "text/html; charset=utf-8")
        if parsed.path == "/app.js":
            return self._send_file(BASE / "web" / "app.js", "text/javascript; charset=utf-8")
        if parsed.path == "/styles.css":
            return self._send_file(BASE / "web" / "styles.css", "text/css; charset=utf-8")
        if parsed.path.startswith("/image/"):
            image = BASE / parsed.path.lstrip("/")
            if image.exists():
                return self._send_file(image, "image/jpeg")
        if parsed.path == "/api/state":
            return self._send_json(STORE.snapshot())
        self._send_json({"error": "Not found"}, 404)

    def do_POST(self):
        parsed = urlparse(self.path)
        size = int(self.headers.get("Content-Length", "0"))
        payload = json.loads(self.rfile.read(size) or "{}")

        if parsed.path == "/api/stat":
            STORE.set_stat(payload["name"], payload["value"])
        elif parsed.path == "/api/skill":
            STORE.toggle_skill(payload["name"], payload["specialized"])
        elif parsed.path == "/api/item":
            STORE.add_item(payload.get("target", "bag"), payload)
        elif parsed.path == "/api/transfer":
            STORE.transfer(payload["source"], payload["name"], payload["qty"])
        elif parsed.path == "/api/weapon":
            STORE.assign_weapon(payload)
        elif parsed.path == "/api/equipment":
            STORE.assign_equipment(payload)
        elif parsed.path == "/api/equip":
            ok = STORE.toggle_equipped(payload["kind"], payload["name"], payload["equipped"])
            if not ok:
                return self._send_json({"ok": False, "message": "Limite atteinte"}, 400)
        elif parsed.path == "/api/buy":
            ok, msg = STORE.buy(payload["category"], payload["name"], payload["qty"])
            if not ok:
                return self._send_json({"ok": False, "message": msg}, 400)
        elif parsed.path == "/api/sell":
            STORE.sell(payload["source"], payload["name"], payload["qty"])
        else:
            return self._send_json({"error": "Not found"}, 404)

        self._send_json({"ok": True, "state": STORE.snapshot()})


if __name__ == "__main__":
    server = ThreadingHTTPServer(("0.0.0.0", 8000), Handler)
    print("Server running at http://0.0.0.0:8000")
    server.serve_forever()
