#!/usr/bin/env python3
"""
build_data.py — extrait la BDD « Cartographie Outils IT » et injecte le modèle
de données dans index.html, entre les marqueurs DONNÉES BDD.

    python3 data/build_data.py --xlsx data/BDD_Cartographie_Outils_IT.xlsx

Trois axes sont repris de la BDD :
  · Nature      : colonne « Nature » (Socle de données / Application métiers)
  · Interne/Ext : colonne « Interne / Externe »
  · Alimentation: colonne « Alimentation (Manuel/Auto) »
Les flux sont reconstruits depuis « Alimente quel outil » et « Alimenté par qui ».
"""

import argparse
import collections
import json
import pathlib
import re
import sys

import openpyxl

ROOT = pathlib.Path(__file__).resolve().parent.parent
SHEET = "📋 Collecte Entretiens"
MARK_A = "/* ==== DONNÉES BDD — début (généré par data/build_data.py) ==== */"
MARK_B = "/* ==== DONNÉES BDD — fin ==== */"

# --------------------------------------------------------------------------- #
#  Normalisation des libellés de la BDD (saisie libre lors des entretiens)     #
# --------------------------------------------------------------------------- #
ALIAS = {
    "quickconnect": "Quick Connect", "trimbleconnect": "Trimble Connect",
    "sharepoint": "SharePoint", "e-checking": "E-Checking", "idcapture": "IDCapture",
    "ms project": "MS Project", "power bi": "Power BI", "powerbi": "Power BI",
    "cabine de pilotage(power bi)": "Power BI", "wastemarket place (wakio)": "Wastemarket Place",
    "bip /baps": "BIP / BAPS", "bip": "BIP / BAPS", "harmony": "Harmony", "harmonie": "Harmony",
    "autocad": "AutoCAD", "civil 3d": "Civil 3D", "simple bim": "Simple BIM",
    "bim vision": "BIM Vision", "sofistik bridge": "Sofistik Bridge", "e-paraph": "E-Paraph",
    "e-project": "E-Project", "by my site": "By My Site", "h&r for you": "H&R For You",
    "cyclone 3dr": "Cyclone 3DR", "cyclone 3d": "Cyclone 3DR", "la scene": "La Scene",
    "liveobject(orange)": "Live Objects (Orange)", "sixsense monitoring": "Sixense Monitoring",
    "lafarge +": "Lafarge+", "power point": "PowerPoint", "baseware": "Basware",
    "docusign": "DocuSign", "notebooklm": "NotebookLM", "puma": "Puma", "covadis": "Covadis",
}

SERVICE = {
    "Direction": "dir", "Méthode/BIM": "met",
    "Qualité": "qse", "Environnement": "qse",
    "Contrat Manager": "daf", "Comptabilité/gestion": "daf", "Assistant RH": "daf",
    "Reponsable Travaux": "trv", "Ingé travaux": "trv",
    "TOPO": "top", "TUNNEL": "tun",
}

SERVICES = [
    {"id": "met", "name": "Méthodes & BIM",          "kicker": "Ingénierie · BIM",                 "col": "left"},
    {"id": "top", "name": "Topographie",             "kicker": "Relevés · Implantation",           "col": "left"},
    {"id": "tun", "name": "Tunnel",                  "kicker": "Ouvrage souterrain",               "col": "left"},
    {"id": "trv", "name": "Travaux",                 "kicker": "Exécution chantier",               "col": "left"},
    {"id": "qse", "name": "Qualité & Environnement", "kicker": "Contrôle · Réserves · Nuisances",  "col": "right"},
    {"id": "dir", "name": "Direction",               "kicker": "Pilotage · Sous-traitance",        "col": "right"},
    {"id": "daf", "name": "DAF",                     "kicker": "Contrats · Compta · Achats · RH",  "col": "right"},
]

# Missions par pôle : (libellé, valeurs FLUX absorbées, outils rattachés à la main, proposée)
#
# La colonne FLUX de la BDD n'est pas exploitable telle quelle : libellés fragmentés
# (« GESTION CHANTIER » / « GESTION DES CHANTIER »), et vide pour Topographie, Tunnel et
# les outils Environnement. Les libellés retenus sont ceux de la note de cadrage ; ceux
# marqués « proposée » ne s'appuient sur aucune ligne d'entretien et restent à valider.
MISSIONS = {
    "met": [
        ("Production de maquettes",              ["PRODUCTION DE MAQUETTES"], []),
        ("Gestion des maquettes",                ["GESTION DE MAQUETTES"], []),
    ],
    "top": [
        ("Relevés et nuages de points",          [], ["La Scene", "Cyclone 3DR", "Trimble Connect"], True),
        ("Plans et calculs",                     [], ["AutoCAD", "Covadis"], True),
        ("Suivi des levers",                     [], ["Pixis", "Excel"], True),
    ],
    "tun": [
        ("Planification et avancement",          [], ["MS Project", "YELLOW", "Power BI"], True),
        ("Plans et maquettes",                   [], ["AutoCAD", "Trimble Connect"], True),
        ("Réserves et contrôles",                [], ["IDCapture", "Quick Connect", "Excel"], True),
        ("Gestion du stock",                     [], ["GMAO"], True),
    ],
    "trv": [
        ("Jalons de chantier",                   ["AVANCEMENTS JALON CHANTIER"], []),
        ("Engins et matériels",                  ["COMMANDE CHANTIER"], []),
        ("Réserves et MaD",                      ["MAD"], []),
        ("Conduite de chantier",                 ["GESTION CHANTIER", "GESTION DES CHANTIER",
                                                  "GESTION IMPRÉVU CHANTIER"], []),
        ("Validation et contrats",               ["CIRCUIT DE VALIDATION", "GESTION DES CONTRATS"], []),
    ],
    "qse": [
        ("Contrôle qualité",                     ["CONTRÔLE QUALITÉ",
                                                  "CONTRÔLE QUALITÉ (VISUALISATION)"], ["CEMEX"]),
        ("MaD et levées de réserve",             ["MAD"], ["Trimble Connect"]),
        ("Suivi environnemental",                ["CONTRÔLE ENVIRONNEMENT",
                                                  "CONTRÔLE QUALITÉ /ENVIRONNEMENT"], ["PowerPoint"]),
        ("Suivi des consommations",              [], ["Live Objects (Orange)", "Sixense Monitoring",
                                                      "Lucee TP", "Lafarge+"], True),
        ("Gestion des déchets",                  [], ["Wastemarket Place"], True),
        ("Commande de matériel",                 ["COMMANDE DE MATÉRIEL"], []),
    ],
    "dir": [
        ("Récolte des KPI",                      ["AVANCEMENTS JALON CHANTIER"], []),
        ("Validation contrats et dépenses",      ["CIRCUIT DE VALIDATION", "GESTION DES CONTRATS"], []),
        ("Gestion des ST",                       ["GESTION DES ST", "GESTION TRAVAUX"], []),
        ("Matériels",                            ["GESTION DES MATÉRIELS"], []),
        ("Services généraux",                    [], ["CWT", "Tableau", "Word"], True),
    ],
    "daf": [
        ("Engagement des dépenses",               ["ENGAGEMENT DES DÉPENSES"], []),
        ("Facturation",                           ["GESTION DE FACTURATION", "FACTURATION PARTENAIRE"], []),
        ("Encaissements",                         ["GESTION DES ENCAISSEMENTS"], []),
        ("Commandes et achats",                   [], ["Achat +", "Pablo"], True),
        ("Gestion contractuelle",                 ["GESTION DE CONTRATS", "AVANCEMENTS JALON CHANTIER"], []),
        ("Personnel et formations",                ["GESTION DU PERSONNEL", "ON BOARDING RH",
                                                    "GESTION DES FORMATIONS"], ["Neoaccès"]),
    ],
}

# Socle commun : outils tagués « Socle de données » dans la BDD, complétés par les
# plateformes transverses que la cartographie source plaçait déjà en logiciel commun.
# Colonne de gauche = outils échangeant avec les services de gauche, et
# inversement : l'ordre des listes commande le placement dans la grille 2 colonnes,
# ce qui évite aux connecteurs de traverser le socle.
# Les groupes qui alimentent le socle de données (Terrain, Bureautique) sont placés
# juste au-dessus de celui-ci : les connecteurs internes restent courts.
SOCLE = [
    ("GED & collaboration",      ["Trimble Connect", "Mezzoteam"]),
    ("Validation & signature",   ["E-Paraph", "DocuSign"]),
    ("Achats, finance & tiers",  ["Pablo", "Harmony", "BATIS", "TIPS", "E-Checking"]),
    ("RH & main-d'œuvre",        ["Puma", "BIP / BAPS"]),
    ("Terrain & intégration",    ["Quick Connect", "YELLOW", "Traktor"]),
    ("Bureautique",              ["Excel", "Word"]),
]
CORE = ["SharePoint", "Power BI"]

# Outils cités dans la BDD sans ligne d'entretien qualifiée : rattachement proposé.
A_QUALIFIER = {
    "GMAO":      ("tun", "Gestion du stock tunnel"),
    "Achat +":   ("daf", "Achats opérationnels (amont Pablo)"),
    "Neoaccès":  ("daf", "Création des badges et gestion des accès"),
    "CEMEX":     ("qse", "Portail fournisseur béton"),
}

# Flux reconstruits depuis « Alimente quel outil » / « Alimenté par qui ».
# Le caractère automatique ou manuel d'un flux est celui du mode d'alimentation
# relevé pour l'outil de destination.
FLOWS = [
    ("s:Excel",             "s:SharePoint",  "Dépôt des tableaux de suivi"),
    ("s:Excel",             "s:Power BI",    "Alimentation de la cabine de pilotage"),
    ("s:Quick Connect",     "s:SharePoint",  "Fiches de contrôle vers le DOE"),
    ("s:Quick Connect",     "s:Power BI",    "Indicateurs qualité et sécurité"),
    ("s:Pablo",             "s:Harmony",     "Commandes vers l'engagement de dépense"),
    ("daf:Basware",         "s:Harmony",     "Factures fournisseurs harmonisées"),
    ("s:Harmony",           "daf:Basware",   "Retour de saisie et litiges"),
    ("s:BIP / BAPS",        "s:Puma",        "Heures chantier vers la paie"),
    ("trv:ERP MAT",         "s:Power BI",    "Inventaire matériel"),
    ("qse:Wastemarket Place", "s:Power BI",  "Suivi des déchets"),
    ("s:Trimble Connect",   "top:AutoCAD",   "Maquettes vers le dessin 2D"),
    ("s:Trimble Connect",   "top:Cyclone 3DR", "Maquettes vers le nuage de points"),
    ("s:Trimble Connect",   "qse:IDCapture", "Repérage des constats — flux indirect, par captures d'écran", False),
    ("top:La Scene",        "top:Cyclone 3DR", "Scans vers le retraitement"),
    ("top:AutoCAD",         "top:Covadis",   "Plans vers les calculs topo"),
]


def norm(v):
    s = "" if v is None else str(v).strip()
    return ALIAS.get(s.lower(), s)


def txt(v):
    return "" if v is None else str(v).replace("\n", " ").strip()


PAIRS = set()          # (pôle, valeur FLUX, outil) — sert à ventiler les outils par mission


def read(xlsx):
    """Agrège la feuille d'entretiens : un enregistrement par outil."""
    ws = openpyxl.load_workbook(xlsx, data_only=True)[SHEET]
    rows = list(ws.iter_rows(values_only=True))[3:]

    T = collections.defaultdict(lambda: {
        "svc": collections.Counter(), "nature": collections.Counter(),
        "ie": collections.Counter(), "alim": collections.Counter(),
        "ed": collections.Counter(), "proc": collections.Counter(),
        "usage": [], "mat": collections.Counter(),
    })
    for r in rows:
        name = norm(r[6])
        if not name or name.isdigit():
            continue
        t = T[name]
        sid = SERVICE.get(txt(r[2]))
        if sid:
            t["svc"][sid] += 1
        if txt(r[4]) and txt(r[4]).upper() not in ("N.A", "NA"):
            t["proc"][txt(r[4]).upper()] += 1

        nat = txt(r[5]).lower()
        if "socle" in nat and "métier" in nat:
            t["nature"]["mixte"] += 1
        elif "socle" in nat:
            t["nature"]["socle"] += 1
        elif "métier" in nat:
            t["nature"]["metier"] += 1

        ie = txt(r[8]).lower()
        if ie.startswith("int"):
            t["ie"]["interne"] += 1
        elif ie.startswith("ext"):
            t["ie"]["externe"] += 1

        al = txt(r[10]).lower()
        if "manuel" in al and "auto" in al:
            t["alim"]["mixte"] += 1
        elif "auto" in al:
            t["alim"]["auto"] += 1
        elif "manuel" in al:
            t["alim"]["manuel"] += 1

        if sid:
            PAIRS.add((sid, txt(r[4]).upper(), name))
        if txt(r[9]):
            t["ed"][txt(r[9])] += 1
        if txt(r[15]):
            t["mat"][re.sub(r"^\d+\s*[–-]\s*", "", txt(r[15]))] += 1
        if txt(r[7]):
            t["usage"].append(txt(r[7]))
    return T


def top(counter, default="?"):
    return counter.most_common(1)[0][0] if counter else default


def build_model(T):
    socle_names = [n for _, group in SOCLE for n in group] + CORE

    tools = {}
    for name, t in T.items():
        tools[name] = {
            "ie": top(t["ie"]),
            "alim": top(t["alim"]),
            "ed": re.sub(r"\s*\(.*?\)?$", "", top(t["ed"], "")).strip(),
            "mat": top(t["mat"], ""),
            "svc": sorted(t["svc"]),
            "nb": len(t["svc"]),
            "usage": (t["usage"][0][:110] if t["usage"] else ""),
            "socle": name in socle_names,
        }
    for name, (sid, usage) in A_QUALIFIER.items():
        tools.setdefault(name, {"ie": "?", "alim": "?", "ed": "", "mat": "",
                                "svc": [sid], "nb": 1, "usage": usage,
                                "socle": False, "todo": True})

    services = []
    for s in SERVICES:
        pid = s["id"]
        own = sorted(n for n, d in tools.items() if pid in d["svc"])

        def order(names):
            """Outils métier d'abord, socle ensuite : la mission se lit de ses outils
            propres vers ce qu'elle emprunte au socle commun."""
            uniq = sorted(set(n for n in names if n in tools))
            return (sorted(n for n in uniq if not tools[n]["socle"]) +
                    sorted(n for n in uniq if tools[n]["socle"]))

        missions, placed = [], set()
        for entry in MISSIONS.get(pid, []):
            label, fluxes, extra = entry[:3]
            todo = entry[3] if len(entry) > 3 else False
            names = {t for (sid, f, t) in PAIRS if sid == pid and f in fluxes}
            names.update(extra)
            ordered = order(names)
            placed.update(ordered)
            missions.append({"label": label, "todo": todo, "tools": ordered})

        rest = [n for n in own if n not in placed]
        if rest:
            missions.append({"label": "Autres outils", "todo": True, "tools": order(rest)})

        services.append(dict(s, tools=own, missions=missions))

    missing = [e for f in FLOWS for e in f[:2]
               if e.split(":", 1)[1] not in tools]

    if missing:
        print("ATTENTION — flux vers un outil inconnu :", set(missing), file=sys.stderr)

    flows = []
    for entry in FLOWS:
        src, dst, obj = entry[:3]
        on_map = entry[3] if len(entry) > 3 else True
        dst_tool = dst.split(":", 1)[1]
        flows.append({"from": src, "to": dst, "obj": obj, "map": on_map,
                      "alim": tools.get(dst_tool, {}).get("alim", "?")})

    return {"socle": [{"group": g, "tools": ts} for g, ts in SOCLE],
            "core": CORE, "services": services, "tools": tools, "flows": flows}


def inject(model, html_path):
    html = pathlib.Path(html_path)
    src = html.read_text(encoding="utf-8")
    a, b = src.find(MARK_A), src.find(MARK_B)
    if a < 0 or b < 0:
        sys.exit(f"marqueurs DONNÉES BDD absents de {html_path}")
    literal = "const BDD = " + json.dumps(model, ensure_ascii=False, indent=1) + ";"
    html.write_text(src[:a] + MARK_A + "\n" + literal + "\n" + src[b:], encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--xlsx", default=str(ROOT / "data" / "BDD_Cartographie_Outils_IT.xlsx"))
    ap.add_argument("--html", default=str(ROOT / "index.html"))
    args = ap.parse_args()

    model = build_model(read(args.xlsx))
    inject(model, args.html)

    n_miss = sum(len(s["missions"]) for s in model["services"])
    orphans = [(s["name"], m["tools"]) for s in model["services"]
               for m in s["missions"] if m["label"] == "Autres outils"]
    for name, outils in orphans:
        print(f"  ! {name} : outils sans mission — {', '.join(outils)}", file=sys.stderr)
    n_int = sum(1 for d in model["tools"].values() if d["ie"] == "interne")
    n_ext = sum(1 for d in model["tools"].values() if d["ie"] == "externe")
    n_auto = sum(1 for d in model["tools"].values() if d["alim"] in ("auto", "mixte"))
    print(f'{len(model["tools"])} outils · {len(model["socle"])} groupes socle · '
          f'{len(model["services"])} pôles · {n_miss} missions · {len(model["flows"])} flux · '
          f'{n_int} internes / {n_ext} externes · {n_auto} alimentés en automatique')


if __name__ == "__main__":
    sys.exit(main())
