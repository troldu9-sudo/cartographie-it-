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

Deux modèles de mission coexistent, et c'est voulu :
  · MISSIONS — l'inventaire : un libellé, les outils qui le servent. Utilisé par les
    pôles dont le processus n'a pas été recueilli (Direction, Méthodes & BIM,
    Qualité & Environnement).
  · PROCESS  — la chaîne : entrée de la donnée, étapes ordonnées, finalité, et les
    points où la mission alimente un autre pôle. Saisi depuis le brief métier — la
    BDD ne porte pas l'ordre des étapes. Un pôle présent dans PROCESS ignore
    MISSIONS et ne retient que les outils cités dans ses chaînes.
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

# Le contrat management était rattaché à la DAF ; il devient un pôle à part entière,
# la chaîne de facturation partenaire circulant désormais explicitement entre les deux.
SERVICE = {
    "Direction": "dir", "Méthode/BIM": "met",
    "Qualité": "qse", "Environnement": "qse",
    "Contrat Manager": "ctr", "Comptabilité/gestion": "daf", "Assistant RH": "daf",
    "Reponsable Travaux": "trv", "Ingé travaux": "trv",
    "TOPO": "top", "TUNNEL": "tun",
}

SERVICES = [
    {"id": "met", "name": "Méthodes & BIM",          "kicker": "Maquettes et plans"},
    {"id": "top", "name": "Topographie",             "kicker": "Relevés · Implantation"},
    {"id": "tun", "name": "Travaux tunnel",          "kicker": "Creusement et tunnelier"},
    {"id": "trv", "name": "Travaux",                 "kicker": "Conduite de travaux"},
    {"id": "ctr", "name": "Contrat",                 "kicker": "Contract management"},
    {"id": "dir", "name": "Direction",               "kicker": "Pilotage du chantier"},
    {"id": "daf", "name": "DAF",                     "kicker": "RH · Comptabilité · Gestion"},
    {"id": "qse", "name": "Qualité & Environnement", "kicker": "Contrôle et conformité"},
    {"id": "sec", "name": "Sécurité",                "kicker": "Prévention · Habilitations"},
]

# Répartition sur la planche 1 : cinq pôles techniques en haut, quatre pôles
# support en bas, socle en bandeau au centre.
BOARD_TOP = ["met", "top", "tun", "trv", "ctr"]
BOARD_BOTTOM = ["dir", "daf", "qse", "sec"]

# Grille de la planche « Flux inter-services ». L'ordre n'est pas cosmétique :
# un fil ne sait tracer qu'une courbe de Bézier, sans contournement. Il ne reste
# lisible qu'entre deux cartes voisines — même rangée mitoyenne, ou rangées
# consécutives. build_model() refuse silencieusement le reste : il le signale.
XBOARD_ROWS = [
    ["met", "top", "tun"],
    ["dir", "ctr", "daf"],
    ["qse", "trv", "sec"],
]

# --------------------------------------------------------------------------- #
#  Modèle « inventaire » — pôles dont le processus n'a pas été recueilli       #
# --------------------------------------------------------------------------- #
# (libellé, valeurs FLUX absorbées, outils rattachés à la main, proposée ?)
MISSIONS = {
    "met": [
        ("Production de maquettes",          ["PRODUCTION DE MAQUETTES"], []),
        ("Gestion des maquettes",            ["GESTION DE MAQUETTES"], []),
    ],
    "dir": [
        ("Récolte des KPI",                  ["AVANCEMENTS JALON CHANTIER"], []),
        ("Validation contrats et dépenses",  ["CIRCUIT DE VALIDATION", "GESTION DES CONTRATS"], []),
        ("Gestion des ST",                   ["GESTION DES ST", "GESTION TRAVAUX"], []),
        ("Matériels",                        ["GESTION DES MATÉRIELS"], []),
        ("Services généraux",                [], ["CWT", "Tableau", "Word"], True),
    ],
    "qse": [
        ("Contrôle qualité",                 ["CONTRÔLE QUALITÉ",
                                              "CONTRÔLE QUALITÉ (VISUALISATION)"], ["CEMEX"]),
        ("MaD et levées de réserve",         ["MAD"], ["Trimble Connect"]),
        ("Suivi environnemental",            ["CONTRÔLE ENVIRONNEMENT",
                                              "CONTRÔLE QUALITÉ /ENVIRONNEMENT"], ["PowerPoint"]),
        ("Suivi des consommations",          [], ["Live Objects (Orange)", "Sixense Monitoring",
                                                  "Lucee TP", "Lafarge+"], True),
        ("Gestion des déchets",              [], ["Wastemarket Place"], True),
        ("Commande de matériel",             ["COMMANDE DE MATÉRIEL"], []),
    ],
}

# --------------------------------------------------------------------------- #
#  Modèle « chaîne » — entrée de la donnée, étapes ordonnées, finalité         #
# --------------------------------------------------------------------------- #
# Chaque mission : {label, note?, todo?, in[], steps[], out, links[]}
#   in    : sources de la donnée      — {lab, tool?}
#   steps : étapes ordonnées          — {act, tool?}
#   out   : finalité                  — {lab, tool?}
#   links : points d'interconnexion   — {at, to, mission, obj?}
#           at = "in" (la mission est alimentée par le pôle cible)
#              | "out" (la mission alimente le pôle cible)
#              | index d'étape (l'échange se fait à cette étape)
# Ces liens produisent à la fois le badge posé sur l'affiche et les fils tracés
# sur la planche « Flux inter-services » : une seule saisie, deux rendus.
PROCESS = {
    "daf": [
        {"label": "Contrôle budgétaire",
         "note": "conjoint avec les Travaux",
         "in": [{"lab": "VALO main-d'œuvre"},
                {"lab": "Extraction matériel", "tool": "BYMAT"},
                {"lab": "Dépenses diverses", "tool": "Basware"}],
         "steps": [{"act": "Constitution de la PC100", "tool": "Excel"},
                   {"act": "Ventilation par centre d'imputation", "tool": "Excel"},
                   {"act": "Ventilation par atelier", "tool": "Excel"},
                   {"act": "Contrôle des écarts au budget", "tool": "Excel"}],
         "out": {"lab": "Budget par atelier arbitré"},
         "links": [{"at": "in", "to": "trv", "mission": "Commande matériel chantier",
                    "obj": "Extraction BYMAT — matériel mobilisé"},
                   {"at": "out", "to": "trv", "mission": "Contrôle budgétaire",
                    "obj": "PC100 — dépenses ventilées par atelier"}]},

        {"label": "Préfacturation partenaire",
         "in": [{"lab": "Factures partenaires", "tool": "Basware"},
                {"lab": "Pointage main-d'œuvre", "tool": "Puma"},
                {"lab": "Heures intérim", "tool": "BIP / BAPS"}],
         "steps": [{"act": "Extraction du pointage", "tool": "Excel"},
                   {"act": "Préfacturation mensuelle", "tool": "Excel"},
                   {"act": "Circuit de validation", "tool": "E-Paraph"}],
         "out": {"lab": "Valorisation de la main-d'œuvre (VALO MO)"},
         "links": [{"at": "out", "to": "ctr", "mission": "Facturation partenaire",
                    "obj": "Valorisation MO à refacturer"},
                   {"at": "out", "to": "daf", "mission": "Contrôle budgétaire",
                    "obj": "VALO MO — coût de la main-d'œuvre"}]},

        {"label": "Engagement des dépenses et facturation",
         "in": [{"lab": "Besoin chantier"}],
         "steps": [{"act": "Bon de commande", "tool": "Pablo"},
                   {"act": "Engagement de la dépense", "tool": "Harmony"},
                   {"act": "Facture fournisseur", "tool": "Basware"},
                   {"act": "Validation interne", "tool": "E-Paraph"},
                   {"act": "Signature", "tool": "DocuSign"},
                   {"act": "Mise en paiement", "tool": "TIPS"}],
         "out": {"lab": "Dépense engagée et réglée"},
         "links": [{"at": "out", "to": "daf", "mission": "Contrôle budgétaire",
                    "obj": "Dépenses diverses — factures et commandes"}]},

        {"label": "Onboarding RH et formations",
         "in": [{"lab": "Nouvel arrivant"}],
         "steps": [{"act": "Vérification travail illégal", "tool": "E-Checking"},
                   {"act": "Badge et accès site", "tool": "Neoaccès"},
                   {"act": "Plan de formation", "tool": "By My Site"},
                   {"act": "Commande de formation", "tool": "Pablo"}],
         "out": {"lab": "Collaborateur habilité et sur site"},
         "links": [{"at": "out", "to": "sec", "mission": "Accueil, formation et habilitations",
                    "obj": "Nouvel arrivant à accueillir sur site"}]},

        {"label": "Gestion du personnel et paye",
         "note": "BYCN et intérim",
         "in": [{"lab": "Pointage chantier", "tool": "Puma"}],
         "steps": [{"act": "Heures intérim", "tool": "BIP / BAPS"},
                   {"act": "Compagnons Bouygues", "tool": "BYCN"},
                   {"act": "Congés et attestations", "tool": "H&R For You"},
                   {"act": "Contrôle des heures", "tool": "Excel"}],
         "out": {"lab": "Paye émise, heures valorisées"},
         "links": [{"at": "in", "to": "trv", "mission": "Commande de main-d'œuvre",
                    "obj": "Besoin en personnel intérimaire"},
                   {"at": "out", "to": "daf", "mission": "Préfacturation partenaire",
                    "obj": "Heures pointées à valoriser"}]},
    ],

    "ctr": [
        {"label": "Gestion des contrats",
         "in": [{"lab": "Marché et avenants"}],
         "steps": [{"act": "Dépôt documentaire", "tool": "SharePoint"},
                   {"act": "Mémoires et réclamations", "tool": "NotebookLM"},
                   {"act": "Rédaction", "tool": "Word"},
                   {"act": "Suivi contractuel", "tool": "Excel"}],
         "out": {"lab": "Dossier contractuel à jour"}},

        {"label": "Avancement et jalons chantier",
         "in": [{"lab": "Avancement chantier"}],
         "steps": [{"act": "Consolidation", "tool": "Excel"},
                   {"act": "Note de situation", "tool": "Word"},
                   {"act": "Tableau de bord", "tool": "Power BI"}],
         "out": {"lab": "Situation mensuelle établie"},
         "links": [{"at": "in", "to": "trv", "mission": "Avancement et jalons chantier",
                    "obj": "Avancement et jalons relevés au chantier"},
                   {"at": "out", "to": "ctr", "mission": "Gestion des encaissements",
                    "obj": "Situation mensuelle à encaisser"}]},

        {"label": "Gestion des encaissements",
         "in": [{"lab": "Suivi des avancements et jalons", "tool": "Excel"}],
         "steps": [{"act": "Situation envoyée au maître d'œuvre", "tool": "Outlook"},
                   {"act": "Circuit de validation", "tool": "E-Project"},
                   {"act": "Validation et paiement du maître d'ouvrage"},
                   {"act": "Facture maître d'ouvrage", "tool": "Chorus"},
                   {"act": "Envoi par mail", "tool": "Outlook"}],
         "out": {"lab": "Dispatch entre partenaires", "tool": "Appli bancaire"}},

        {"label": "Facturation partenaire",
         "in": [{"lab": "Valorisation de la main-d'œuvre"}],
         "steps": [{"act": "Facture partenaire", "tool": "Basware"},
                   {"act": "Activation de la facturation", "tool": "E-Project"}],
         "out": {"lab": "Partenaire facturé"},
         "links": [{"at": "in", "to": "daf", "mission": "Préfacturation partenaire",
                    "obj": "Valorisation MO à refacturer"}]},
    ],

    "trv": [
        {"label": "Gestion de chantier et imprévus",
         "in": [{"lab": "Relevé terrain", "tool": "Quick Connect"}],
         "steps": [{"act": "Constats et réserves", "tool": "IDCapture"},
                   {"act": "Plans et maquettes", "tool": "Trimble Connect"},
                   {"act": "Diffusion documentaire", "tool": "Mezzoteam"},
                   {"act": "Suivi d'exécution", "tool": "Excel"}],
         "out": {"lab": "Chantier piloté, imprévus tracés"}},

        {"label": "Avancement et jalons chantier",
         "in": [{"lab": "Avancement terrain"}],
         "steps": [{"act": "Planification", "tool": "MS Project"},
                   {"act": "Adhérence au planning", "tool": "MMS"},
                   {"act": "Tableau de bord", "tool": "Power BI"}],
         "out": {"lab": "Jalons tenus"},
         "links": [{"at": "out", "to": "ctr", "mission": "Avancement et jalons chantier",
                    "obj": "Avancement et jalons relevés au chantier"},
                   {"at": "out", "to": "dir", "mission": "Récolte des KPI",
                    "obj": "Avancement consolidé pour le pilotage"}]},

        {"label": "Commande matériel chantier",
         "in": [{"lab": "Besoin matériel"}],
         "steps": [{"act": "Demande de matériel", "tool": "Traktor"},
                   {"act": "Coffrage, grue, engins", "tool": "ERP MAT"},
                   {"act": "Bon de commande", "tool": "Pablo"}],
         "out": {"lab": "Matériel mobilisé sur site"},
         "links": [{"at": "out", "to": "daf", "mission": "Contrôle budgétaire",
                    "obj": "Extraction BYMAT — matériel mobilisé"}]},

        {"label": "Commande de main-d'œuvre",
         "in": [{"lab": "Besoin en main-d'œuvre"}],
         "steps": [{"act": "Demande par mail", "tool": "Outlook"},
                   {"act": "Contact de la boîte d'intérim"}],
         "out": {"lab": "Personnel affecté au chantier"},
         "links": [{"at": "out", "to": "daf", "mission": "Gestion du personnel et paye",
                    "obj": "Besoin en personnel intérimaire"}]},

        {"label": "Contrôle budgétaire",
         "note": "conjoint avec la DAF",
         "in": [{"lab": "Dépenses chantier (PC100)"}],
         "steps": [{"act": "Ventilation par atelier", "tool": "Excel"},
                   {"act": "Analyse des écarts", "tool": "Excel"}],
         "out": {"lab": "Arbitrage budgétaire chantier"},
         "links": [{"at": "in", "to": "daf", "mission": "Contrôle budgétaire",
                    "obj": "PC100 — dépenses ventilées par atelier"}]},

        {"label": "Gestion des contrats et avenants",
         "in": [{"lab": "Avenant"}],
         "steps": [{"act": "Rédaction", "tool": "Word"},
                   {"act": "Validation interne", "tool": "E-Paraph"}],
         "out": {"lab": "Avenant signé"}},
    ],

    "tun": [
        {"label": "Gestion de chantier", "todo": True,
         "in": [{"lab": "Relevé terrain tunnel", "tool": "Quick Connect"}],
         "steps": [{"act": "Réserves et contrôles", "tool": "IDCapture"},
                   {"act": "Plans d'exécution", "tool": "AutoCAD"},
                   {"act": "Maquettes", "tool": "Trimble Connect"},
                   {"act": "Suivi d'exécution", "tool": "Excel"}],
         "out": {"lab": "Creusement piloté"}},

        {"label": "Avancement et jalons chantier", "todo": True,
         "in": [{"lab": "Avancement tunnelier", "tool": "YELLOW"}],
         "steps": [{"act": "Planification", "tool": "MS Project"},
                   {"act": "Tableau de bord", "tool": "Power BI"}],
         "out": {"lab": "Jalons tunnel tenus"},
         "links": [{"at": "out", "to": "dir", "mission": "Récolte des KPI",
                    "obj": "Avancement du creusement"},
                   {"at": "out", "to": "ctr", "mission": "Avancement et jalons chantier",
                    "obj": "Jalons tunnel pour la situation"}]},

        {"label": "Commande chantier", "todo": True,
         "in": [{"lab": "Besoin — stock tunnel", "tool": "GMAO"}],
         "steps": [{"act": "Bon de commande", "tool": "Pablo"},
                   {"act": "Matériel et coffrage", "tool": "ERP MAT"}],
         "out": {"lab": "Consommables et matériel livrés"},
         "links": [{"at": "out", "to": "daf", "mission": "Contrôle budgétaire",
                    "obj": "Dépenses matériel tunnel"}]},
    ],

    "top": [
        {"label": "Acquisition et traitement des nuages de points",
         "in": [{"lab": "Acquisition des plans", "tool": "SharePoint"}],
         "steps": [{"act": "Exploitation des plans", "tool": "Trimble Connect"},
                   {"act": "Assemblage des scans chantier", "tool": "La Scene"},
                   {"act": "Création des nuages de points", "tool": "Covadis"}],
         "out": {"lab": "Nuage de points exploitable"}},

        {"label": "Reporting et diffusion des plans",
         "in": [{"lab": "Ouverture des plans", "tool": "Trimble Connect"}],
         "steps": [{"act": "Assemblage et vérification de conformité", "tool": "Cyclone 3DR"},
                   {"act": "Rapport de contrôle", "tool": "Excel"}],
         "out": {"lab": "Rapport de conformité diffusé"},
         "links": [{"at": "out", "to": "met", "mission": "Gestion des maquettes",
                    "obj": "Plans contrôlés et nuages de points"},
                   {"at": "out", "to": "tun", "mission": "Gestion de chantier",
                    "obj": "Plans d'implantation du tunnel"}]},

        {"label": "Guidage du tunnelier",
         "in": [{"lab": "Position du tunnelier"}],
         "steps": [{"act": "Guidage temps réel", "tool": "Pixis"}],
         "out": {"lab": "Tracé guidé en temps réel"},
         "links": [{"at": "out", "to": "tun", "mission": "Avancement et jalons chantier",
                    "obj": "Position et tracé du tunnelier"}]},
    ],

    "sec": [
        {"label": "Accueil, formation et habilitations",
         "in": [{"lab": "Nouvel arrivant"}],
         "steps": [{"act": "Accueil de site et formation au poste", "tool": "Quick Connect"},
                   {"act": "Habilitations élec. et autorisation de conduite", "tool": "Quick Connect"},
                   {"act": "Suivi des formations", "tool": "By My Site"},
                   {"act": "Dossier collaborateur", "tool": "HRMYOU / Global HR"}],
         "out": {"lab": "Collaborateur habilité au poste"},
         "links": [{"at": "in", "to": "daf", "mission": "Onboarding RH et formations",
                    "obj": "Nouvel arrivant à accueillir sur site"}]},

        {"label": "Prévention et suivi terrain",
         "in": [{"lab": "Terrain · ¼ heure sécurité"}],
         "steps": [{"act": "Visite sécurité et ¼ h sécurité", "tool": "Quick Connect"},
                   {"act": "Suivi IDV, addictions, organismes", "tool": "Quick Connect Sécurité"},
                   {"act": "Comptage tunnel", "tool": "Lotus"},
                   {"act": "Diffusion documentaire", "tool": "Mezzoteam"},
                   {"act": "Commande EPI et prestations", "tool": "Pablo"}],
         "out": {"lab": "Risques maîtrisés sur site"}},

        {"label": "Événements et indicateurs sécurité",
         "in": [{"lab": "Événement ATB, AT, PAT, HIPO"}],
         "steps": [{"act": "Reporting événement", "tool": "Cority"},
                   {"act": "Volume d'heures exposées", "tool": "Heures Travaillées"},
                   {"act": "Indicateurs sécurité", "tool": "Power BI"}],
         "out": {"lab": "Taux de fréquence et de gravité publiés"},
         "links": [{"at": "out", "to": "dir", "mission": "Récolte des KPI",
                    "obj": "Indicateurs sécurité du chantier"}]},
    ],
}

# --------------------------------------------------------------------------- #
#  Socle commun et outils hors BDD                                            #
# --------------------------------------------------------------------------- #
# Socle commun : outils tagués « Socle de données » dans la BDD, complétés par les
# plateformes transverses que la cartographie source plaçait déjà en logiciel commun.
# Outlook rejoint la bureautique : le brief le fait porter la commande de main-d'œuvre
# (Travaux) et l'envoi des situations (Contrat) — c'est un outil transverse.
SOCLE = [
    ("GED & collaboration",      ["Trimble Connect", "Mezzoteam"]),
    ("Validation & signature",   ["E-Paraph", "DocuSign"]),
    ("Achats, finance & tiers",  ["Pablo", "Harmony", "BATIS", "TIPS", "E-Checking"]),
    ("RH & main-d'œuvre",        ["Puma", "BIP / BAPS"]),
    ("Terrain & intégration",    ["Quick Connect", "YELLOW", "Traktor"]),
    ("Bureautique",              ["Excel", "Word", "Outlook"]),
]
CORE = ["SharePoint", "Power BI"]

# Outils cités dans la BDD sans ligne d'entretien qualifiée : rattachement proposé.
A_QUALIFIER = {
    "GMAO":      ("tun", "Gestion du stock tunnel"),
    "Achat +":   ("daf", "Achats opérationnels (amont Pablo)"),
    "Neoaccès":  ("daf", "Création des badges et gestion des accès"),
    "CEMEX":     ("qse", "Portail fournisseur béton"),
}

# Outils absents de la BDD, introduits par le brief métier. Ils sont marqués comme
# tels sur les planches et dans le référentiel : la base d'entretiens ne les
# documente pas, leur nature et leur alimentation restent à confirmer.
HORS_BDD = {
    "Outlook":                 ("externe", "manuel", "Microsoft",
                                "Messagerie : demandes de MO et envoi des situations"),
    "Chorus":                  ("externe", "manuel", "AIFE",
                                "Facturation par le maître d'ouvrage"),
    "Appli bancaire":          ("externe", "manuel", "",
                                "Dispatch des paiements entre partenaires"),
    "BYCN":                    ("interne", "?", "Bouygues Construction",
                                "Paye des compagnons Bouygues"),
    "BYMAT":                   ("interne", "auto", "Bouygues Construction",
                                "Extraction des coûts matériel pour le contrôle budgétaire"),
    "Cority":                  ("externe", "manuel", "Cority",
                                "Reporting des événements sécurité : ATB, AT, PAT, HIPO"),
    "Quick Connect Sécurité":  ("interne", "manuel", "Bouygues Construction",
                                "Suivi IDV, addictions et organismes"),
    "Lotus":                   ("?", "manuel", "",
                                "Comptage des personnes présentes en tunnel"),
    "Heures Travaillées":      ("interne", "manuel", "",
                                "Volume d'heures exposées, dénominateur des taux sécurité"),
    "HRMYOU / Global HR":      ("interne", "manuel", "Bouygues Construction",
                                "Outils RH groupe : dossier collaborateur"),
}

# Flux reconstruits depuis « Alimente quel outil » / « Alimenté par qui ».
# Le caractère automatique ou manuel d'un flux est celui du mode d'alimentation
# relevé pour l'outil de destination. Ils restent strictement adossés à la BDD :
# les échanges décrits par le brief sont portés par PROCESS, pas ajoutés ici.
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
    ("s:Trimble Connect",   "top:La Scene",  "Plans vers l'assemblage des scans"),
    ("s:Trimble Connect",   "top:Cyclone 3DR", "Maquettes vers le contrôle de conformité"),
    ("s:Trimble Connect",   "qse:IDCapture", "Repérage des constats — flux indirect, par captures d'écran", False),
    ("top:La Scene",        "top:Covadis",   "Scans assemblés vers les nuages de points"),
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


def chain_tools(mission):
    """Outils cités par une chaîne, dans l'ordre de lecture, sans doublon."""
    refs = ([e.get("tool") for e in mission.get("in", [])] +
            [s.get("tool") for s in mission.get("steps", [])] +
            [mission.get("out", {}).get("tool")])
    out = []
    for n in refs:
        if n and n not in out:
            out.append(n)
    return out


def mission_labels(pid):
    """Libellés de mission d'un pôle, quel que soit son modèle."""
    if pid in PROCESS:
        return [m["label"] for m in PROCESS[pid]]
    return [e[0] for e in MISSIONS.get(pid, [])]


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
                                "socle": name in socle_names, "todo": True})
    for name, (ie, alim, ed, usage) in HORS_BDD.items():
        tools.setdefault(name, {"ie": ie, "alim": alim, "ed": ed, "mat": "",
                                "svc": [], "nb": 0, "usage": usage,
                                "socle": name in socle_names, "brief": True})

    # Un pôle décrit par PROCESS ne retient que les outils cités dans ses chaînes :
    # c'est la chaîne qui fait foi, pas la déclaration d'entretien. C'est ce qui
    # retire AutoCAD de la Topographie et rattache Pablo au tunnel.
    for pid, procs in PROCESS.items():
        cited = {n for m in procs for n in chain_tools(m)}
        unknown = cited - set(tools)
        if unknown:
            print(f"ATTENTION — {pid} : étape citant un outil inconnu : "
                  f"{', '.join(sorted(unknown))}", file=sys.stderr)
        for name, d in tools.items():
            svc = set(d["svc"])
            svc.add(pid) if name in cited else svc.discard(pid)
            d["svc"] = sorted(svc)
    for d in tools.values():
        d["nb"] = len(d["svc"])

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

        if pid in PROCESS:
            missions = [{"label": m["label"], "todo": m.get("todo", False),
                         "tools": chain_tools(m)} for m in PROCESS[pid]]
        else:
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

        services.append(dict(s, tools=own, missions=missions,
                             mode="process" if pid in PROCESS else "inventory"))

    missing = [e for f in FLOWS for e in f[:2] if e.split(":", 1)[1] not in tools]
    if missing:
        print("ATTENTION — flux vers un outil inconnu :", set(missing), file=sys.stderr)

    flows = []
    for entry in FLOWS:
        src, dst, obj = entry[:3]
        on_map = entry[3] if len(entry) > 3 else True
        dst_tool = dst.split(":", 1)[1]
        flows.append({"from": src, "to": dst, "obj": obj, "map": on_map,
                      "alim": tools.get(dst_tool, {}).get("alim", "?")})

    # Les liens saisis dans PROCESS produisent les fils de la planche inter-services.
    # Une seule saisie, deux rendus : le badge sur l'affiche et le fil sur la planche.
    xflows, seen = [], set()
    for pid, procs in PROCESS.items():
        for m in procs:
            for lk in m.get("links", []):
                target = mission_labels(lk["to"])
                if lk["mission"] not in target:
                    print(f"ATTENTION — lien {pid}/{m['label']} → {lk['to']} : "
                          f"mission « {lk['mission']} » inconnue", file=sys.stderr)
                    continue
                here, there = f"{pid}:{m['label']}", f"{lk['to']}:{lk['mission']}"
                src, dst = (here, there) if lk["at"] == "out" else (there, here)
                if (src, dst) in seen:
                    continue
                seen.add((src, dst))
                # Un enchaînement interne à un pôle n'est pas un flux inter-services :
                # il reste en badge sur l'affiche et en ligne dans la matrice, mais
                # n'est pas tracé — le fil repasserait par-dessus les missions.
                xflows.append({"from": src, "to": dst, "obj": lk.get("obj", ""),
                               "map": pid != lk["to"], "alim": "manuel"})

    pos = {p: (r, c) for r, row in enumerate(XBOARD_ROWS) for c, p in enumerate(row)}
    for f in xflows:
        a, b = f["from"].split(":", 1)[0], f["to"].split(":", 1)[0]
        if a == b or a not in pos or b not in pos:
            continue
        (ra, ca), (rb, cb) = pos[a], pos[b]
        if abs(ra - rb) > 1 or (ra == rb and abs(ca - cb) > 1):
            print(f"ATTENTION — planche inter-services : {a} et {b} ne sont pas voisins "
                  f"dans XBOARD_ROWS, le fil traversera une carte", file=sys.stderr)

    return {"socle": [{"group": g, "tools": ts} for g, ts in SOCLE],
            "core": CORE, "services": services, "tools": tools, "flows": flows,
            "process": PROCESS, "xflows": xflows,
            "board": {"top": BOARD_TOP, "bottom": BOARD_BOTTOM, "xrows": XBOARD_ROWS}}


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
    n_steps = sum(len(m["steps"]) for procs in PROCESS.values() for m in procs)
    for s in model["services"]:
        for m in s["missions"]:
            if m["label"] == "Autres outils":
                print(f"  ! {s['name']} : outils sans mission — "
                      f"{', '.join(m['tools'])}", file=sys.stderr)
    for n, d in sorted(model["tools"].items()):
        if not d["svc"]:
            print(f"  ! {n} : outil rattaché à aucun pôle", file=sys.stderr)

    n_brief = sum(1 for d in model["tools"].values() if d.get("brief"))
    n_int = sum(1 for d in model["tools"].values() if d["ie"] == "interne")
    n_ext = sum(1 for d in model["tools"].values() if d["ie"] == "externe")
    n_auto = sum(1 for d in model["tools"].values() if d["alim"] in ("auto", "mixte"))
    n_proc = sum(1 for s in model["services"] if s["mode"] == "process")
    print(f'{len(model["tools"])} outils ({n_brief} hors BDD) · '
          f'{len(model["socle"])} groupes socle · {len(model["services"])} pôles '
          f'({n_proc} en chaînes) · {n_miss} missions · {n_steps} étapes · '
          f'{len(model["flows"])} flux applicatifs · {len(model["xflows"])} flux inter-services · '
          f'{n_int} internes / {n_ext} externes · {n_auto} alimentés en automatique')


if __name__ == "__main__":
    sys.exit(main())
