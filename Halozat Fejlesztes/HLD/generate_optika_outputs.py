import csv
import html
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path


BASE = Path(r"C:\Projects\SOTE\Halozat Fejlesztes\HLD")
VSDX_PATH = BASE / "SOTE Optika-V2.vsdx"

NS = {"v": "http://schemas.microsoft.com/office/visio/2012/main"}


def clean_text(value: str) -> str:
    value = " ".join((value or "").split())
    replacements = {
        "Rendezõk": "Rendezők",
        "Rendezõ": "Rendező",
        "Nõi": "Női",
        "Bõr": "Bőr",
        "Mûszaki": "Műszaki",
        "Fõ.": "Fő.",
        "Tömõ": "Tömő",
        "Urulógia": "Urológia",
        "Naygvárad": "Nagyvárad",
        "technika": "technika",
        "Össesen": "Összesen",
        "Szerves Vegytan": "Szerves Vegytan",
        "Schöffmérei": "Schöpf-Merei",
        "Reszõ": "Rezső",
        "Fül-Orr-Gége": "Fül-Orr-Gége",
    }
    for src, dst in replacements.items():
        value = value.replace(src, dst)
    return value.strip()


def parse_vsdx():
    with zipfile.ZipFile(VSDX_PATH) as zf:
        root = ET.fromstring(zf.read("visio/pages/page1.xml"))

    shapes = {}
    for shape in root.find("v:Shapes", NS).findall("v:Shape", NS):
        cells = {c.attrib.get("N"): c.attrib for c in shape.findall("v:Cell", NS)}
        sid = shape.attrib["ID"]
        shapes[sid] = {
            "id": sid,
            "nameu": shape.attrib.get("NameU", ""),
            "text": clean_text("".join(shape.itertext())),
            "pinx": float(cells.get("PinX", {}).get("V", 0) or 0),
            "piny": float(cells.get("PinY", {}).get("V", 0) or 0),
            "width": float(cells.get("Width", {}).get("V", 0) or 0),
            "height": float(cells.get("Height", {}).get("V", 0) or 0),
            "fill": cells.get("FillForegnd", {}).get("V"),
            "cells": cells,
        }

    links = []
    for shape in shapes.values():
        if "Dynamic connector" not in shape["nameu"]:
            continue
        begin_f = shape["cells"].get("BeginX", {}).get("F", "")
        end_f = shape["cells"].get("EndX", {}).get("F", "")
        m1 = re.search(r"Sheet\.(\d+)!", begin_f)
        m2 = re.search(r"Sheet\.(\d+)!", end_f)
        links.append(
            {
                "id": shape["id"],
                "from": m1.group(1) if m1 else None,
                "to": m2.group(1) if m2 else None,
                "label": shape["text"],
            }
        )

    return shapes, links


def parse_building(text: str):
    if text == "BKT Gépterem":
        return {"name": "BKT Gépterem", "rendezok": None, "switches": None}
    if text == "KKT Gépterem":
        return {"name": "KKT Gépterem", "rendezok": None, "switches": None}
    if text == "PRO-M":
        return {"name": "PRO-M", "rendezok": None, "switches": None}

    match = re.search(
        r"Épület:\s*(.*?)\s*Rendezők:\s*([0-9?]+)?\s*Access Switch:\s*([0-9?]+)?",
        text,
    )
    if match:
        rend_raw = (match.group(2) or "").strip()
        sw_raw = (match.group(3) or "").strip()
        return {
            "name": match.group(1).strip(),
            "rendezok": int(rend_raw) if rend_raw.isdigit() else None,
            "rendezok_raw": rend_raw or "?",
            "switches": int(sw_raw) if sw_raw.isdigit() else None,
            "switches_raw": sw_raw or "?",
        }

    if text:
        return {"name": text, "rendezok": None, "switches": None}
    return None


SITE_OVERRIDES = {
    "1": {"role": "core", "owner": "BKT", "display_name": "BKT Gépterem"},
    "70": {"role": "core", "owner": "KKT", "display_name": "KKT Gépterem"},
    "3": {"role": "access", "owner": "BKT", "decision": "Közvetlen core-access", "note": "1 rendező, a 8x SMF bőven elegendő a 2x uplinkhez (4 szál)."},
    "7": {"role": "access", "owner": "BKT", "decision": "Közvetlen core-access", "note": "2 rendező, a Pathológia felőli 12x MMF elegendő a 8 szál igényhez."},
    "14": {"role": "distribution", "owner": "BKT", "decision": "Lokális disztribúció", "note": "4 rendező mellett a meglévő 4x MMF csak disztribúciós uplinkre elegendő."},
    "20": {"role": "distribution", "owner": "BKT", "decision": "Lokális disztribúció", "note": "9 rendezőnél a meglévő 4x MMF csak disztribúciós felkötést tesz reálissá."},
    "22": {"role": "distribution", "owner": "BKT", "decision": "Lokális disztribúció", "note": "10 rendező, a 4x SMF közvetlen accesshez kevés, disztribúcióhoz megfelelő."},
    "26": {"role": "distribution", "owner": "BKT", "decision": "Könyvtár mint disztribúció", "note": "A Gazdasági és Műszaki Fő. innen lóg tovább, ezért ez legyen az aggregációs pont."},
    "29": {"role": "access", "owner": "BKT", "decision": "Access a Könyvtár mögött", "note": "1 rendező, a meglévő 4x SMF elegendő, külön disztribúció nem szükséges."},
    "32": {"role": "distribution", "owner": "BKT", "decision": "Lokális disztribúció", "note": "8 rendező, a 4x SMF csak disztribúciós uplinkre alkalmas."},
    "40": {"role": "distribution", "owner": "BKT", "decision": "Lokális disztribúció", "note": "11 rendező, a BKT felé meglévő 30x SMF jól használható disztribúciós felkötésre."},
    "43": {"role": "distribution", "owner": "BKT", "decision": "Lokális disztribúció", "note": "12 rendező, a 12x SMF közvetlen 2x uplinkes accesshez kevés lenne."},
    "46": {"role": "distribution", "owner": "BKT", "decision": "Lokális disztribúció", "note": "8 rendező, de a BKT felőli optika ismeretlen; minimum 4x OS2 dedikált uplink szükséges."},
    "48": {"role": "access", "owner": "BKT", "decision": "Access a FOK OC mögött", "note": "1 rendező, a FOK OC felé lévő 4x MMF elég 1 uplinkes kiszolgálásra."},
    "53": {"role": "distribution", "owner": "BKT", "decision": "Disztribúció a Transzplantáció ágához", "note": "Saját 3 rendezője mellett egy teljes alágat szolgál ki."},
    "55": {"role": "distribution", "owner": "BKT", "decision": "Disztribúció a Rektori ágához", "note": "Innen indul tovább a Rektori/Bőr ág."},
    "57": {"role": "access", "owner": "BKT", "decision": "Access a II. Belgyógy B mögött", "note": "A disztribúciót a II. Belgyógy B-ben tartom, így itt külön disztribúció nem szükséges."},
    "59": {"role": "access", "owner": "BKT", "decision": "Access a Rektori mögött", "note": "A meglévő 4x MMF kevés 4 rendezőhöz, bővíteni kell 8 szál összkapacitásig."},
    "61": {"role": "distribution", "owner": "BKT", "decision": "Lokális disztribúció", "note": "9? rendező, részben ismeretlen környezet; a 12x OS2 közvetlenül is inkább disztribúcióra megfelelő."},
    "63": {"role": "survey", "owner": "BKT", "decision": "Felmérés után véglegesíthető", "note": "A rendező- és switchszám hiányzik, ezért jelenleg csak előzetes helyfoglalás adható."},
    "64": {"role": "survey", "owner": "BKT", "decision": "Felmérés után véglegesíthető", "note": "A rendező- és switchszám hiányzik, ezért jelenleg csak előzetes helyfoglalás adható."},
    "67": {"role": "distribution", "owner": "BKT", "decision": "Disztribúció az EOK/Ferenc tér ág számára", "note": "6 rendező plusz külső alág, ezért itt érdemes aggregálni."},
    "1000": {"role": "distribution", "owner": "KKT", "decision": "Lokális disztribúció", "note": "7 rendező, a KKT felé meglévő OS2/OM3 készlet disztribúciós uplinkre megfelelő."},
    "1002": {"role": "survey", "owner": "KKT", "decision": "Felmérés után véglegesíthető", "note": "Urológia szerepe és pontos optikai készlete részben bizonytalan."},
    "1003": {"role": "distribution", "owner": "KKT", "decision": "Disztribúció a Tömő ág számára", "note": "7 rendező mellett innen megy tovább a Tömő felé is kapcsolat."},
    "1007": {"role": "distribution", "owner": "KKT", "decision": "Lokális disztribúció", "note": "8 rendező, a Neurológia felé lévő 4x OS2 inkább disztribúciós uplinknek alkalmas."},
    "1009": {"role": "survey", "owner": "KKT", "decision": "Felmérés után véglegesíthető", "note": "A Balassa irány jelen feladatban csak érintőlegesen látszik."},
    "1014": {"role": "survey", "owner": "KKT", "decision": "Felmérés után véglegesíthető", "note": "A Kálvária előtti optikai elosztó felőli szakasz ismeretlen."},
    "1016": {"role": "distribution", "owner": "KKT", "decision": "Lokális disztribúció", "note": "8 rendező; a 12x OS2 elegendő disztribúciós uplinkhez."},
    "1019": {"role": "distribution", "owner": "KKT", "decision": "Lokális disztribúció", "note": "6 rendező, közvetlen access helyett disztribúció célszerű."},
    "1022": {"role": "distribution", "owner": "KKT", "decision": "Lokális disztribúció", "note": "12 rendező; a hozzá vezető optika bizonytalan, de a szerepe egyértelműen disztribúciós."},
    "1024": {"role": "access", "owner": "KKT", "decision": "Közvetlen core-access", "note": "1 rendező, a 8x MMF elegendő a 4 szál igényhez."},
    "1026": {"role": "access", "owner": "KKT", "decision": "Közvetlen core-access Mosoda útvonalon", "note": "1 rendező, a Mosodából kijövő 4x OM3 elegendő a 4 szál igényhez."},
    "1027": {"role": "access", "owner": "KKT", "decision": "Közvetlen core-access", "note": "1 rendező, de a Visió alapján csak 2x OM3 látszik; +2 szál bővítés kell."},
    "1030": {"role": "distribution", "owner": "KKT", "decision": "KKT oldali fő disztribúció", "note": "8 rendező, valamint több alág és másik core felőli redundáns optika is itt csatlakozik."},
    "1033": {"role": "access", "owner": "KKT", "decision": "Access a KKT Sebészet mögött", "note": "1 rendező, a 2x OM4 elegendő 1 uplinkhez."},
    "1035": {"role": "distribution", "owner": "KKT", "decision": "Lokális disztribúció", "note": "3 rendező és jelentős meglévő OS2 készlet."},
    "1037": {"role": "distribution", "owner": "KKT", "decision": "Lokális disztribúció", "note": "15 rendező; közvetlen accessre a látható 2x OS2 nem elég."},
    "1042": {"role": "distribution", "owner": "KKT", "decision": "Kiemelt disztribúciós campuspont", "note": "33 rendező és több további végpont indul innen."},
    "1045": {"role": "survey", "owner": "KKT", "decision": "Felmérés után véglegesíthető", "note": "A Rezső Kollégium rendezőszáma nem ismert, a link 4 vagy 8 szálas bizonytalansággal szerepel."},
    "1051": {"role": "passive", "owner": "KKT", "display_name": "Biztonságtechnika optikai elosztó"},
    "1052": {"role": "mpls", "owner": "PRO-M", "decision": "MPLS L3 végpont", "note": "Nem campus optikai access/disztribúciós pontként kezelendő."},
    "1053": {"role": "mpls", "owner": "PRO-M", "decision": "MPLS L3 végpont", "note": "Nem campus optikai access/disztribúciós pontként kezelendő."},
    "1054": {"role": "mpls", "owner": "PRO-M", "decision": "MPLS L3 végpont", "note": "Nem campus optikai access/disztribúciós pontként kezelendő."},
    "1055": {"role": "mpls", "owner": "PRO-M", "decision": "MPLS L3 végpont", "note": "Nem campus optikai access/disztribúciós pontként kezelendő."},
    "1056": {"role": "mpls", "owner": "PRO-M", "display_name": "PRO-M", "decision": "MPLS L3 aggregáció", "note": "Külön kezelt MPLS L3 ág."},
    "1077": {"role": "distribution", "owner": "BKT", "decision": "Előzetesen disztribúció", "note": "24 rendező alapján biztosan nem célszerű közvetlen core-accessként kezelni."},
    "1078": {"role": "survey", "owner": "BKT", "decision": "Felmérés után véglegesíthető", "note": "A hálózati optikai környezet a rajzból nem fejthető vissza teljesen."},
    "1083": {"role": "distribution", "owner": "BKT", "decision": "Előzetesen disztribúció", "note": "8 rendező és ismeretlen optikai útvonal, ezért minimum disztribúciós helyfoglalás szükséges."},
    "1085": {"role": "distribution", "owner": "BKT", "decision": "Előzetesen disztribúció", "note": "8 rendező; a jelenlegi optika ismeretlen."},
    "1111": {"role": "access", "owner": "BKT", "decision": "Access a II. Gyermek Klinika mögött", "note": "Kis végpont, 8x OS2 feliratú összeköttetéssel."},
    "1115": {"role": "access", "owner": "BKT", "decision": "Access a II. Belgyógy A mögött", "note": "A 12x OS2 összeköttetés elegendő 1 uplinkes kiszolgálásra."},
}


MANUAL_LINKS = [
    {"from": "1024", "to": "1027", "label": "2x OM3 50/125", "status": "build", "action": "Kazán: +2 szál OM3/OM4 szükséges a 4 szál eléréséhez."},
    {"from": "46", "to": "48", "label": "4x MMF 62,5/125", "status": "ok", "action": "Sellye Koll: a FOK OC felé meglévő kapcsolat elegendő kis access végpontnak."},
    {"from": "70", "to": "1019", "label": "4x OS2 9/125", "status": "ok", "action": "Radiológia: a meglévő 4x OS2 alkalmas disztribúciós uplinkre."},
    {"from": "70", "to": "1035", "label": "12x OS2 9/125", "status": "ok", "action": "Fül-Orr-Gége: a meglévő 12x OS2 bőséges disztribúciós kapacitást ad."},
    {"from": "70", "to": "1002", "label": "OS2 9/125 + OM2 50/125 (részben ismert)", "status": "survey", "action": "Urológia: a panelkiosztás ismert, de a teljes optikai készletet és felhasználhatóságát ellenőrizni kell."},
    {"from": "70", "to": "1009", "label": "8x OS2 9/125", "status": "ok", "action": "Balassa Kollégium: a meglévő 8x OS2 legalább egy disztribúciós uplinkre elegendő."},
    {"from": "1051", "to": "70", "label": "ismeretlen", "status": "survey", "action": "Kálvária előtti biztonságtechnikai elosztó upstream optikája nem ismert, helyszíni survey kell."},
]


LINK_STATUS = {
    frozenset(("1", "46")): ("survey", "FOK OC: a BKT felőli optika ismeretlen, minimum 4x OS2 9/125 dedikált szálpár szükséges."),
    frozenset(("57", "59")): ("build", "Bőrgyógyászat: a 4x MMF kevés 4 rendező 1-uplikes kiszolgálásához, 8 szál összkapacitásig bővíteni kell (+4 szál)."),
    frozenset(("61", "63")): ("survey", "Szerves Vegytan ág: rendezőszám és optikai készlet felmérése szükséges."),
    frozenset(("63", "64")): ("survey", "Schöpf-Merei ág: rendezőszám és optikai készlet felmérése szükséges."),
    frozenset(("1016", "1022")): ("survey", "KBE: ismeretlen optika, minimum 4x OS2 9/125 szükséges a disztribúciós uplinkhez."),
    frozenset(("1030", "70")): ("survey", "KKT Sebészet: a core felőli optika részben ismeretlen, legalább 4x OS2 9/125 dedikáltan szükséges."),
    frozenset(("1037", "70")): ("build", "I. Gyermekklinika: disztribúciós uplinkhez legalább 4 OS2 szál kell; a látható 2x OS2 mellé +2 szál OS2 javasolt."),
    frozenset(("1042", "1045")): ("survey", "Rezső Kollégium: 4 vagy 8 szál látszik, rendezőszám nélkül csak felmérés után véglegesíthető."),
    frozenset(("1075", "1079")): ("survey", "PAK ág: optikai típus és darabszám nem ismert."),
    frozenset(("1079", "1083")): ("survey", "PAK-KÚT ág: optikai típus és darabszám nem ismert."),
    frozenset(("1085", "1077")): ("survey", "V70–Érsebészet ág: optikai típus és darabszám nem ismert."),
    frozenset(("1056", "1052")): ("mpls", "MPLS L3 kapcsolat."),
    frozenset(("1056", "1053")): ("mpls", "MPLS L3 kapcsolat."),
    frozenset(("1056", "1054")): ("mpls", "MPLS L3 kapcsolat."),
    frozenset(("1056", "1055")): ("mpls", "MPLS L3 kapcsolat."),
}


ROLE_COLORS = {
    "core": "#177245",
    "distribution": "#1756a9",
    "access": "#7a4fb0",
    "passive": "#d9a300",
    "mpls": "#666666",
    "survey": "#b94a48",
}

LINK_COLORS = {
    "ok": "#3a7a3d",
    "build": "#d97706",
    "survey": "#c53030",
    "mpls": "#6b7280",
}


def infer_status(link):
    key = frozenset((link["from"], link["to"]))
    if key in LINK_STATUS:
        status, action = LINK_STATUS[key]
        return status, action
    if "?" in (link.get("label") or ""):
        return "survey", "A felirat alapján az optikai típus vagy mennyiség nem ismert."
    return "ok", "A meglévő optika a választott szerepkiosztással várhatóan felhasználható."


def build_dataset():
    shapes, links = parse_vsdx()
    nodes = {}
    for sid, shape in shapes.items():
        parsed = parse_building(shape["text"])
        if not parsed:
            continue
        if sid not in SITE_OVERRIDES and not (
            "Épület:" in shape["text"]
            or "Gépterem" in shape["text"]
            or shape["fill"] == "#fee599"
            or shape["text"] in {"PRO-M"}
        ):
            continue

        node = {
            "id": sid,
            "name": SITE_OVERRIDES.get(sid, {}).get("display_name", parsed["name"]),
            "rendezok": parsed.get("rendezok"),
            "rendezok_raw": parsed.get("rendezok_raw", ""),
            "access_switches": parsed.get("switches"),
            "access_switches_raw": parsed.get("switches_raw", ""),
            "pinx": shape["pinx"],
            "piny": shape["piny"],
            "width": shape["width"],
            "height": shape["height"],
            "fill": shape["fill"],
            "role": "passive" if shape["fill"] == "#fee599" else "survey",
            "owner": "",
            "decision": "",
            "note": "",
            "raw_text": shape["text"],
        }
        node.update(SITE_OVERRIDES.get(sid, {}))
        nodes[sid] = node

    visible_links = []
    for link in links:
        if not link["from"] or not link["to"]:
            continue
        if link["from"] not in nodes or link["to"] not in nodes:
            continue
        status, action = infer_status(link)
        visible_links.append(
            {
                "from": link["from"],
                "to": link["to"],
                "label": clean_text(link["label"]),
                "status": status,
                "action": action,
            }
        )

    for link in MANUAL_LINKS:
        visible_links.append(link)

    current_links = defaultdict(list)
    for link in visible_links:
        current_links[link["from"]].append(f"{nodes[link['to']]['name']}: {link['label'] or 'nincs felirat'}")
        current_links[link["to"]].append(f"{nodes[link['from']]['name']}: {link['label'] or 'nincs felirat'}")

    rows = []
    for sid, node in sorted(nodes.items(), key=lambda item: (item[1]["owner"], item[1]["name"])):
        current_state = "; ".join(sorted(current_links.get(sid, [])))
        target_state = node.get("decision", "")
        rows.append(
            {
                "azonosito": sid,
                "helyszin": node["name"],
                "oldal": node.get("owner", ""),
                "rendezok": node["rendezok_raw"] or (node["rendezok"] if node["rendezok"] is not None else ""),
                "access_switch": node["access_switches_raw"] or (node["access_switches"] if node["access_switches"] is not None else ""),
                "cel_szerep": node["role"],
                "jelenlegi_kapcsolatok": current_state,
                "celallapot": target_state,
                "megjegyzes": node.get("note", ""),
            }
        )

    return nodes, visible_links, rows


def infer_section_owner(node_id, nodes, links):
    owner = nodes[node_id].get("owner")
    if owner:
        return owner

    neighbors = []
    for link in links:
        if link["from"] == node_id and link["to"] in nodes:
            neighbors.append(nodes[link["to"]].get("owner"))
        if link["to"] == node_id and link["from"] in nodes:
            neighbors.append(nodes[link["from"]].get("owner"))

    neighbors = [item for item in neighbors if item]
    if neighbors:
        return neighbors[0]
    return "BKT"


def card_meta(node):
    rend = node["rendezok_raw"] or (str(node["rendezok"]) if node["rendezok"] is not None else "-")
    sw = node["access_switches_raw"] or (str(node["access_switches"]) if node["access_switches"] is not None else "-")
    return f"rendező: {rend} | access sw: {sw}"


def generate_html(nodes, links):
    section_order = ["BKT", "KKT", "PRO-M"]
    role_order = {
        "core": 0,
        "passive": 1,
        "distribution": 2,
        "survey": 3,
        "access": 4,
        "mpls": 4,
    }
    column_x = {
        "core": 88,
        "passive": 350,
        "distribution": 660,
        "survey": 660,
        "access": 970,
        "mpls": 970,
    }
    column_label = {
        "core": "Core",
        "passive": "Passzív optika",
        "distribution": "Disztribúció",
        "survey": "Survey / bizonytalan",
        "access": "Access",
        "mpls": "MPLS végpont",
    }
    card_width = {
        "core": 220,
        "passive": 220,
        "distribution": 240,
        "survey": 240,
        "access": 240,
        "mpls": 240,
    }
    section_meta = {
        "BKT": {"title": "BKT oldal", "subtitle": "BKT core-ból kiszolgált campus-ágak"},
        "KKT": {"title": "KKT oldal", "subtitle": "KKT core-ból kiszolgált campus-ágak"},
        "PRO-M": {"title": "MPLS / külső helyszínek", "subtitle": "Pontozott MPLS L3 logika külön kezelve"},
    }

    sections = {key: {"nodes": [], "links": []} for key in section_order}
    for node_id, node in nodes.items():
        section = infer_section_owner(node_id, nodes, links)
        if section not in sections:
            section = "PRO-M" if node["role"] == "mpls" else "BKT"
        sections[section]["nodes"].append(node_id)

    for link in links:
        a_owner = infer_section_owner(link["from"], nodes, links)
        b_owner = infer_section_owner(link["to"], nodes, links)
        if a_owner == b_owner:
            sections.setdefault(a_owner, {"nodes": [], "links": []})["links"].append(link)

    section_html = []
    for section_key in section_order:
        node_ids = sections[section_key]["nodes"]
        if not node_ids:
            continue

        sorted_nodes = sorted(
            node_ids,
            key=lambda nid: (
                role_order.get(nodes[nid]["role"], 9),
                0 if nodes[nid]["role"] == "core" else nodes[nid]["pinx"],
                nodes[nid]["piny"],
                nodes[nid]["name"],
            ),
        )

        y_cursor = 110
        role_counters = defaultdict(int)
        positions = {}
        node_cards = []
        for nid in sorted_nodes:
            node = nodes[nid]
            role = node["role"]
            x = column_x.get(role, 660)
            width = card_width.get(role, 240)
            height = 88 if role in {"distribution", "survey"} else 76
            y = y_cursor
            y_cursor += height + 26
            role_counters[role] += 1
            positions[nid] = {
                "x": x,
                "y": y,
                "w": width,
                "h": height,
                "cx": x + width / 2,
                "cy": y + height / 2,
            }
            tooltip = html.escape((node.get("decision", "") + " | " + node.get("note", "")).strip(" |"))
            note = html.escape(node.get("decision") or "")
            meta = html.escape(card_meta(node))
            node_cards.append(
                f"""
                <div class="node role-{role}" style="left:{x}px;top:{y}px;width:{width}px;height:{height}px;border-color:{ROLE_COLORS[role]};" title="{tooltip}">
                  <div class="node-title">{html.escape(node['name'])}</div>
                  <div class="node-role">{html.escape(column_label.get(role, role.title()))}</div>
                  <div class="node-meta">{meta}</div>
                  <div class="node-note">{note}</div>
                </div>
                """
            )

        line_elems = []
        label_offsets = defaultdict(int)
        for idx, link in enumerate(sections[section_key]["links"]):
            if link["from"] not in positions or link["to"] not in positions:
                continue
            a = positions[link["from"]]
            b = positions[link["to"]]
            start_x = a["x"] + a["w"]
            start_y = a["cy"]
            end_x = b["x"]
            end_y = b["cy"]
            if a["x"] > b["x"]:
                start_x = a["x"]
                end_x = b["x"] + b["w"]
            elbow_x = int((start_x + end_x) / 2)
            color = LINK_COLORS[link["status"]]
            dash = ' stroke-dasharray="7 6"' if link["status"] == "mpls" else ""
            label_key = (elbow_x, int((start_y + end_y) / 2 / 30))
            label_offsets[label_key] += 1
            label_y = int((start_y + end_y) / 2) + (label_offsets[label_key] - 1) * 20 - 10
            label_x = elbow_x
            line_elems.append(
                f"""
                <path d="M {start_x} {start_y} L {elbow_x} {start_y} L {elbow_x} {end_y} L {end_x} {end_y}" stroke="{color}" stroke-width="3" fill="none"{dash}>
                  <title>{html.escape(link['label'] or 'nincs felirat')} | {html.escape(link['action'])}</title>
                </path>
                <rect x="{label_x - 120}" y="{label_y - 11}" width="240" height="22" rx="7" class="edge-label-bg edge-{link['status']}"></rect>
                <text x="{label_x}" y="{label_y + 4}" text-anchor="middle" class="edge-label">{html.escape(link['label'] or 'nincs felirat')}</text>
                """
            )

        canvas_height = max((pos["y"] + pos["h"] for pos in positions.values()), default=420) + 70
        section_html.append(
            f"""
            <section class="section">
              <div class="section-head">
                <div>
                  <h2>{section_meta[section_key]['title']}</h2>
                  <p>{section_meta[section_key]['subtitle']}</p>
                </div>
                <div class="section-stats">
                  <span>{sum(1 for nid in node_ids if nodes[nid]['role'] == 'distribution')} disztribúció</span>
                  <span>{sum(1 for nid in node_ids if nodes[nid]['role'] == 'access')} access</span>
                  <span>{sum(1 for nid in node_ids if nodes[nid]['role'] == 'survey')} survey</span>
                </div>
              </div>
              <div class="topology" style="height:{canvas_height}px;">
                <div class="column column-core" style="left:70px;">Core</div>
                <div class="column column-passive" style="left:338px;">Passzív optika</div>
                <div class="column column-distribution" style="left:648px;">Disztribúció / survey</div>
                <div class="column column-access" style="left:958px;">Access / végpont</div>
                <svg class="wires" viewBox="0 0 1260 {canvas_height}">
                  {''.join(line_elems)}
                </svg>
                {''.join(node_cards)}
              </div>
            </section>
            """
        )

    cross_core_links = []
    for link in links:
        pair = {link["from"], link["to"]}
        if pair == {"1", "70"}:
            cross_core_links.append(link)

    cross_summary = ""
    if cross_core_links:
        link = cross_core_links[0]
        cross_summary = f"""
        <div class="cross-core">
          <div class="cross-core-card">
            <div class="cross-core-title">Core összeköttetés</div>
            <div class="cross-core-line">
              <span class="core-box">BKT core</span>
              <span class="core-link">{html.escape(link['label'])}</span>
              <span class="core-box">KKT core</span>
            </div>
            <div class="cross-core-note">{html.escape(link['action'])}</div>
          </div>
        </div>
        """

    html_text = f"""<!doctype html>
<html lang="hu">
<head>
  <meta charset="utf-8">
  <title>SOTE Optikai Célállapot</title>
  <style>
    body {{
      margin: 0;
      font-family: "Segoe UI", Arial, sans-serif;
      background:
        radial-gradient(circle at top left, rgba(58, 124, 165, 0.12), transparent 26%),
        linear-gradient(180deg, #edf4fb 0%, #f8fafc 55%, #eef3f8 100%);
      color: #14212b;
    }}
    .wrap {{
      max-width: 1420px;
      margin: 0 auto;
      padding: 28px 24px 56px;
    }}
    .intro {{
      display: grid;
      grid-template-columns: 1.3fr .7fr;
      gap: 18px;
      margin-bottom: 24px;
    }}
    .hero, .summary, .section {{
      background: rgba(255,255,255,0.9);
      border: 1px solid rgba(209, 219, 230, 0.9);
      border-radius: 24px;
      box-shadow: 0 20px 50px rgba(13, 31, 45, 0.08);
    }}
    .hero {{
      padding: 24px 26px;
    }}
    h1 {{
      margin: 0 0 10px;
      font-size: 30px;
      line-height: 1.1;
    }}
    .hero p {{
      margin: 0;
      line-height: 1.5;
      color: #415364;
      font-size: 15px;
    }}
    .summary {{
      padding: 22px;
      display: grid;
      gap: 12px;
      align-content: start;
    }}
    .summary strong {{
      display: block;
      font-size: 14px;
      margin-bottom: 4px;
    }}
    .legend {{
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      margin-bottom: 24px;
    }}
    .pill {{
      background: rgba(255,255,255,0.88);
      border-radius: 999px;
      padding: 9px 13px;
      border: 1px solid #d8e0eb;
      font-size: 13px;
      box-shadow: 0 8px 22px rgba(14, 27, 41, 0.05);
    }}
    .section {{
      padding: 22px;
      margin-bottom: 22px;
    }}
    .section-head {{
      display: flex;
      justify-content: space-between;
      gap: 18px;
      align-items: start;
      margin-bottom: 18px;
    }}
    .section-head h2 {{
      margin: 0 0 6px;
      font-size: 22px;
    }}
    .section-head p {{
      margin: 0;
      color: #516273;
      font-size: 14px;
    }}
    .section-stats {{
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      justify-content: end;
    }}
    .section-stats span {{
      font-size: 12px;
      padding: 7px 10px;
      border-radius: 999px;
      background: #f5f8fb;
      border: 1px solid #dbe5ee;
      color: #435466;
    }}
    .topology {{
      position: relative;
      overflow: auto;
      background:
        linear-gradient(90deg, rgba(18, 48, 69, 0.03) 1px, transparent 1px) 0 0 / 22px 22px,
        linear-gradient(rgba(18, 48, 69, 0.03) 1px, transparent 1px) 0 0 / 22px 22px,
        linear-gradient(180deg, #fbfdff 0%, #f4f8fb 100%);
      border: 1px solid #d8e0eb;
      border-radius: 22px;
      min-height: 480px;
      min-width: 1240px;
    }}
    .column {{
      position: absolute;
      top: 18px;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: #708191;
      font-weight: 700;
      z-index: 2;
    }}
    .wires {{
      position: absolute;
      inset: 0;
      width: 1260px;
      height: 100%;
      z-index: 1;
    }}
    .node {{
      position: absolute;
      background: rgba(255,255,255,0.98);
      border: 3px solid;
      border-radius: 18px;
      padding: 12px 14px;
      box-sizing: border-box;
      box-shadow: 0 18px 36px rgba(12, 27, 42, 0.11);
      z-index: 2;
    }}
    .node-title {{
      font-weight: 700;
      font-size: 14px;
      margin-bottom: 4px;
      line-height: 1.25;
    }}
    .node-role {{
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: #6a7b8c;
      margin-bottom: 8px;
    }}
    .node-meta {{
      font-size: 12px;
      color: #4a5d6f;
      margin-bottom: 8px;
    }}
    .node-note {{
      font-size: 12px;
      line-height: 1.35;
      color: #26333f;
    }}
    .edge-label {{
      font-size: 11px;
      fill: #1f2937;
      font-weight: 600;
    }}
    .edge-label-bg {{
      fill: rgba(255,255,255,0.92);
      stroke: rgba(199, 210, 220, 0.9);
    }}
    .edge-build {{
      fill: rgba(255, 247, 237, 0.95);
      stroke: rgba(217, 119, 6, 0.35);
    }}
    .edge-survey {{
      fill: rgba(254, 242, 242, 0.95);
      stroke: rgba(197, 48, 48, 0.30);
    }}
    .cross-core {{
      margin-bottom: 22px;
    }}
    .cross-core-card {{
      background: linear-gradient(135deg, #123b56 0%, #1a5d7e 100%);
      color: #fff;
      border-radius: 22px;
      padding: 18px 22px;
      box-shadow: 0 20px 45px rgba(13, 43, 64, 0.18);
    }}
    .cross-core-title {{
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.1em;
      opacity: 0.8;
      margin-bottom: 12px;
    }}
    .cross-core-line {{
      display: flex;
      align-items: center;
      gap: 14px;
      flex-wrap: wrap;
      font-weight: 700;
    }}
    .core-box {{
      padding: 10px 14px;
      border-radius: 12px;
      background: rgba(255,255,255,0.12);
      border: 1px solid rgba(255,255,255,0.16);
    }}
    .core-link {{
      padding: 8px 12px;
      border-radius: 999px;
      background: rgba(255,255,255,0.16);
      font-size: 13px;
    }}
    .cross-core-note {{
      margin-top: 10px;
      color: rgba(255,255,255,0.82);
      font-size: 13px;
    }}
    @media (max-width: 1100px) {{
      .intro {{
        grid-template-columns: 1fr;
      }}
      .section-head {{
        flex-direction: column;
      }}
    }}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="intro">
      <div class="hero">
        <h1>SOTE optikai célállapot topológia</h1>
        <p>Az ábra már nem a Visio koordinátáit másolja, hanem olvasható hálózati nézetben mutatja a javasolt vegyes <code>core-distribution-access</code> és <code>core-access</code> struktúrát. A helyszínek külön BKT, KKT és MPLS panelekbe vannak rendezve, így a disztribúciós pontok, access végpontok, passzív optikai elemek és a bővítendő szakaszok azonnal átláthatók.</p>
      </div>
      <div class="summary">
        <div><strong>Tervezési elv</strong> disztribúció minimalizálása, ahol a meglévő optika még reálisan engedi.</div>
        <div><strong>Kék oldali szabály</strong> rendezőnként 2 uplink, azaz 4 szál.</div>
        <div><strong>Lila oldali szabály</strong> rendezőnként 1 uplink, azaz 2 szál.</div>
        <div><strong>Színkód</strong> zöld: használható, narancs: bővítendő, piros: survey szükséges.</div>
      </div>
    </div>
    <div class="legend">
      <div class="pill">Zöld keret: core</div>
      <div class="pill">Kék keret: disztribúció</div>
      <div class="pill">Lila keret: access</div>
      <div class="pill">Sárga keret: passzív optikai pont</div>
      <div class="pill">Piros keret: felmérés szükséges</div>
      <div class="pill">Narancs vonal: új optika kell</div>
      <div class="pill">Piros vonal: bizonytalan / survey</div>
      <div class="pill">Szaggatott szürke: MPLS L3</div>
    </div>
    {cross_summary}
    {''.join(section_html)}
  </div>
</body>
</html>
"""
    (BASE / "SOTE_optikai_celallapot.html").write_text(html_text, encoding="utf-8")


def generate_report(rows, links):
    assumptions = [
        "Ahol a Visio színinformációja nem volt XML-ből egyértelműen kiolvasható, ott a kék/lila uplink-szabályt a topológiai szerep és az alárendeltség alapján alkalmaztam.",
        "A kék épületeknél rendezőnként 2 fizikai uplinkkel, vagyis 4 szállal számoltam.",
        "A lila, alárendelt access pontoknál rendezőnként 1 uplinkkel, vagyis 2 szállal számoltam.",
        "Az ismeretlen `?` jelölésű optikai szakaszoknál csak minimumigényt adtam meg; ezeket helyszíni felméréssel kell véglegesíteni.",
        "A PRO-M ág és a hozzá tartozó külső helyszínek MPLS L3 végpontként szerepelnek, nem campus optikai access/disztribúciós pontként.",
    ]

    upgrades = [link for link in links if link["status"] in {"build", "survey"}]

    md = []
    md.append("# SOTE optikai célállapot összefoglaló")
    md.append("")
    md.append("## Alapfeltevések")
    for item in assumptions:
        md.append(f"- {item}")
    md.append("")
    md.append("## Célállapot logika")
    md.append("- A kis, 1 rendezős végpontok ahol elegendő optika látszik, közvetlen `core-access` kapcsolattal maradnak.")
    md.append("- A nagy, sok rendezős vagy további épületeket kiszolgáló helyszínek lokális disztribúciós pontok lesznek.")
    md.append("- Ahol a teljesen lapos core-access megoldáshoz aránytalanul sok új szálat kellene építeni, ott a disztribúciót megtartottam.")
    md.append("")
    md.append("## Optikai bővítések és survey pontok")
    for link in upgrades:
        a = next(row["helyszin"] for row in rows if row["azonosito"] == link["from"])
        b = next(row["helyszin"] for row in rows if row["azonosito"] == link["to"])
        prefix = "Bővítés" if link["status"] == "build" else "Survey"
        md.append(f"- {prefix}: {a} <-> {b} | meglévő felirat: `{link['label'] or 'nincs'}` | {link['action']}")
    md.append("")
    md.append("## Helyszínszintű táblázat")
    md.append("")
    md.append("| Helyszín | Oldal | Rendezők | Cél szerep | Jelenlegi állapot röviden | Elvárt állapot / megjegyzés |")
    md.append("| --- | --- | ---: | --- | --- | --- |")
    for row in rows:
        md.append(
            f"| {row['helyszin']} | {row['oldal']} | {row['rendezok'] or '-'} | {row['cel_szerep']} | "
            f"{row['jelenlegi_kapcsolatok'] or '-'} | {row['celallapot']} {row['megjegyzes']} |"
        )
    (BASE / "SOTE_optikai_jelenlegi_vs_celallapot.md").write_text("\n".join(md), encoding="utf-8")


def generate_csv(rows):
    out_path = BASE / "SOTE_optikai_jelenlegi_vs_celallapot.csv"
    with out_path.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(
            fh,
            fieldnames=[
                "azonosito",
                "helyszin",
                "oldal",
                "rendezok",
                "access_switch",
                "cel_szerep",
                "jelenlegi_kapcsolatok",
                "celallapot",
                "megjegyzes",
            ],
            delimiter=";",
        )
        writer.writeheader()
        writer.writerows(rows)


def main():
    nodes, links, rows = build_dataset()
    generate_html(nodes, links)
    generate_report(rows, links)
    generate_csv(rows)
    summary = {
        "nodes": len(nodes),
        "links": len(links),
        "outputs": [
            str(BASE / "SOTE_optikai_celallapot.html"),
            str(BASE / "SOTE_optikai_jelenlegi_vs_celallapot.md"),
            str(BASE / "SOTE_optikai_jelenlegi_vs_celallapot.csv"),
        ],
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
