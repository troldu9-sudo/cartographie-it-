#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extraction de la BDD d'entretiens vers la source unique data/carto.json.

    python3 data/extract_bdd.py

Écrit deux fichiers :
  data/carto.json      — source unique consommée par les deux générateurs
  data/gap-report.md   — tout ce que la BDD ne dit pas, ou dit deux fois

Le script ne décide rien tout seul : chaque arbitrage est une constante nommée
en tête de fichier, et chaque écart est reporté dans le gap-report avec le
numéro de ligne de la BDD qui le porte.
"""

import json
import re
import sys
import unicodedata
from collections import Counter, OrderedDict, defaultdict
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

RACINE = Path(__file__).resolve().parent.parent
BDD = RACINE / "data" / "BDD_Cartographie_Outils_IT.xlsx"
SORTIE_JSON = RACINE / "data" / "carto.json"
SORTIE_GAP = RACINE / "data" / "gap-report.md"

FEUILLE_COLLECTE = "📋 Collecte Entretiens"
FEUILLE_BORD = "📊 Tableau de Bord"
LIGNE_ENTETE = 3
PREMIERE_LIGNE = 4
DERNIERE_LIGNE = 102  # 103 = ligne de totaux, hors périmètre

# Colonnes exploitées. « Fréquence d'utilisation » (L) et « Niveau Maturité » (P)
# sont volontairement absentes : elles ne sont ni extraites ni affichées.
COL = {
    "n": "A", "service": "C", "interlocuteur": "D", "flux": "E", "nature": "F",
    "outil": "G", "usage": "H", "ie": "I", "editeur": "J", "alim": "K",
    "alimente_par": "M", "alimente_vers": "N",
}

# ─────────────────────────────────────────────────────────────────────────────
# Charte — tokens uniques, seuls autorisés dans les livrables
# ─────────────────────────────────────────────────────────────────────────────
CHARTE = OrderedDict([
    ("primary", "#E85A0C"),
    ("accent", "#C1121F"),
    ("encre", "#1C1C1C"),
    ("gris", "#6B6B6B"),
    ("fond", "#F4F4F4"),
    ("graph_a", "#3A6B8A"),   # graphiques de la synthèse uniquement
    ("graph_b", "#4A7C59"),   # graphiques de la synthèse uniquement
    ("blanc", "#FFFFFF"),
    ("trait", "#D8D8D8"),
    ("police", "Segoe UI, Calibri, DejaVu Sans, Arial, sans-serif"),
    # PowerPoint ne comprend pas une pile CSS : les largeurs de texte du
    # module layout sont celles de Calibri, c'est donc Calibri qui est posé.
    ("police_pptx", "Calibri"),
])

# ─────────────────────────────────────────────────────────────────────────────
# Normalisation des libellés d'outils (validée en phase 0)
# ─────────────────────────────────────────────────────────────────────────────
CANON_OUTIL = {
    "Baseware": "Basware",
    "POWER BI": "Power BI",
    "Cabine de pilotage(Power BI)": "Power BI",   # rapport Power BI, fusionné
    "IDCAPTURE": "IDCapture",
    "Trimbleconnect": "Trimble Connect",
    "Sharepoint": "SharePoint",
    "Docusign": "DocuSign",
    "Autocad": "AutoCAD",
    "Puma": "PUMA",
    "Pablo": "PABLO",
    "MS PROJECT": "MS Project",
    "BIP /BAPS": "BIP/BAPS",
    "Achat +": "HA+",
    "Neoacces": "Visioaccès",
    "Liveobject(ORANGE)": "Live Objects (Orange)",
    "Wastemarket place (wakio)": "Waste Marketplace (Wakio)",
    "E-paraph": "E-Paraph",
    "E-checking": "E-Checking",
    "E-project": "E-Project",
    "La scene": "La Scene",
    "Cyclone 3dr": "Cyclone 3DR",
    "Simple BIM": "Simple BIM",
    "BIM vision": "BIM Vision",
    "Sofistik bridge": "Sofistik Bridge",
    "Lucee TP": "Lucee TP",
    "Lafarge +": "Lafarge +",
    "Sixsense monitoring": "Sixsense Monitoring",
    "BY MY SITE": "By My Site",
    "H&R For you": "H&R For You",
    "Power Point": "PowerPoint",
    "MIRO": "Miro",
    "Oasis": "Oasis",
    "Quickconnect": "QuickConnect",
    "Tableau": "Tableau",
    "Traktor": "Traktor",
    "ERP MAT": "ERP MAT",
    "NotebookLM": "NotebookLM",
    "Anaconda": "Anaconda",
    "QGIS": "QGIS",
    "Revit": "Revit",
    "Inventor": "Inventor",
    "Civil 3D": "Civil 3D",
    "Covadis": "Covadis",
    "Pixis": "Pixis",
    "TIPS": "TIPS",
    "MMS": "MMS",
    "YELLOW": "Yellow",
    "BATIS": "BATIS",
    "CWT": "CWT",
    "Harmony": "Harmony",
    "Mezzoteam": "Mezzoteam",
    "Excel": "Excel",
    "Word": "Word",
}

# ─────────────────────────────────────────────────────────────────────────────
# Les 9 services de la cartographie
# ─────────────────────────────────────────────────────────────────────────────
# « court » sert aux axes de graphique : un libellé qui se replie sur deux
# lignes fait sauter une étiquette sur deux dans le classement.
SERVICES = OrderedDict([
    ("direction",     {"nom": "Direction", "court": "Direction",
                       "kicker": "Pilotage du chantier"}),
    ("daf",           {"nom": "DAF", "court": "DAF",
                       "kicker": "RH · Comptabilité · Gestion"}),
    ("contrat",       {"nom": "Contrat", "court": "Contrat",
                       "kicker": "Contract management"}),
    ("methode-bim",   {"nom": "Méthode & BIM", "court": "Méthode & BIM",
                       "kicker": "Maquettes et plans"}),
    ("travaux",       {"nom": "Travaux", "court": "Travaux",
                       "kicker": "Conduite de travaux"}),
    ("travaux-tunnel", {"nom": "Travaux tunnel", "court": "Trav. tunnel",
                        "kicker": "Creusement et tunnelier"}),
    ("topo",          {"nom": "Topo", "court": "Topo",
                       "kicker": "Topographie et implantation"}),
    ("qualite-env",   {"nom": "Qualité & Environnement", "court": "Qualité & Env.",
                       "kicker": "Contrôle et conformité"}),
    ("securite",      {"nom": "Sécurité", "court": "Sécurité",
                       "kicker": "À documenter"}),
])

# Service vide sur consigne : aucun outil prélevé, blocs éditables
SERVICES_VIDES = {"securite"}

# Rattachement des libellés bruts de la BDD aux 9 services
RATTACHEMENT = {
    "Direction": "direction",
    "Comptabilité/gestion": "daf",
    "Assistant RH": "daf",
    "Contrat Manager": "contrat",
    "Méthode/BIM": "methode-bim",
    "Ingé travaux": "travaux",
    "Reponsable Travaux": "travaux",
    "TUNNEL": "travaux-tunnel",
    "TOPO": "topo",
    "Qualité": "qualite-env",
    "Environnement": "qualite-env",
}

# ─────────────────────────────────────────────────────────────────────────────
# Harmonisation des libellés de processus (colonne FLUX)
# ─────────────────────────────────────────────────────────────────────────────
CANON_PROCESSUS = {
    "GESTION DES ST": "Gestion des sous-traitants",
    "CIRCUIT DE VALIDATION": "Circuit de validation",
    "GESTION DES MATÉRIELS": "Gestion des matériels",
    "GESTION DES CONTRATS": "Gestion des contrats",
    "GESTION DE CONTRATS": "Gestion des contrats",
    "Contrôle Qualité": "Contrôle qualité",
    "Contrôle Qualité (visualisation)": "Contrôle qualité",
    "Contrôle Qualité /Environnement": "Contrôle qualité et environnement",
    "Contrôle Environnement": "Contrôle environnement",
    "Commande de matériel": "Commande de matériel",
    "MAD": "Mise à disposition (MAD)",
    "GESTION DE MAQUETTES": "Gestion des maquettes",
    "PRODUCTION DE MAQUETTES": "Production des maquettes",
    "GESTION DES ENCAISSEMENTS": "Gestion des encaissements",
    "AVANCEMENTS JALON CHANTIER": "Avancement et jalons chantier",
    "FACTURATION PARTENAIRE": "Facturation partenaire",
    "ENGAGEMENT DES DÉPENSES": "Engagement des dépenses",
    "GESTION DE FACTURATION": "Gestion de la facturation",
    "GESTION DES CHANTIER": "Gestion de chantier",
    "GESTION CHANTIER": "Gestion de chantier",
    "GESTION TRAVAUX": "Gestion de chantier",
    "GESTION IMPRÉVU CHANTIER": "Gestion des imprévus chantier",
    "GESTION DU PERSONNEL": "Gestion du personnel",
    "GESTION DES FORMATIONS": "Gestion des formations",
    "ON BOARDING RH": "Onboarding RH",
    "COMMANDE CHANTIER": "Commande chantier",
}
# Valeurs de FLUX qui ne désignent aucun processus
FLUX_NON_RENSEIGNE = {"", "N.A", "NA"}

# ─────────────────────────────────────────────────────────────────────────────
# Fonction de l'outil, en une ligne, condensée depuis la colonne « Usage »
# de la BDD. Aucune fonction n'est inventée : un outil absent de cette table
# reçoit un TODO visible dans le référentiel.
# ─────────────────────────────────────────────────────────────────────────────
FONCTION = {
    "Anaconda": "Écriture de scripts (code)",
    "AutoCAD": "Production et modification des plans 2D",
    "BATIS": "Déclaration des sous-traitants en paiement direct",
    "BIM Vision": "Ouverture et visualisation des maquettes",
    "BIP/BAPS": "Gestion de l'intérim : heures, renouvellements, vérifications",
    "By My Site": "Centralisation des formations",
    "Basware": "Enregistrement et validation des factures fournisseurs",
    "CWT": "Réservation de voyages selon la politique groupe",
    "Civil 3D": "Dessin des terrassements",
    "Covadis": "Génération des points topographiques",
    "Cyclone 3DR": "Traitement des nuages de points et analyse des scans",
    "DocuSign": "Notification et signature des nouveaux contrats",
    "E-Checking": "Lutte contre le travail illégal et vérification des autorisations",
    "E-Paraph": "Circuit de validation interne avant signature",
    "E-Project": "Activation de la facturation et suivi de situation",
    "ERP MAT": "Commande de matériel : coffrage, grue à tour, sécurité",
    "Excel": "Tableurs de suivi : factures, chronofacture, rapprochements",
    "H&R For You": "Demandes RH : congés, attestations, épargne",
    "Harmony": "Engagement des dépenses et suivi des commandes",
    "IDCapture": "Levée de réserves et contrôle qualité",
    "Inventor": "Conception 3D de pièces mécaniques",
    "La Scene": "Assemblage des positions de scan",
    "Lafarge +": "Suivi des bons de livraison et des m³ coulés",
    "Live Objects (Orange)": "Remontée des indicateurs des compteurs de pompage",
    "Lucee TP": "Identification des espèces envahissantes",
    "Miro": "Animation des obeya",
    "MMS": "Suivi de l'adhérence au planning et des jalons",
    "MS Project": "Planification des travaux",
    "Mezzoteam": "GED : diffusion et suivi des plans",
    "NotebookLM": "Rédaction de mémoires et de réclamations",
    "Oasis": "Déclaration des incidents",
    "PABLO": "Achats et bons de commande",
    "Pixis": "Guidage du tunnelier",
    "Power BI": "Tableaux de bord : KPI, consommations, avancement",
    "PowerPoint": "Déclaration des presqu'accidents (HIPO)",
    "PUMA": "Pointage des compagnons : heures et primes de poste",
    "QGIS": "SIG : repérage sur plan",
    "QuickConnect": "Relevés terrain : sensibilisation, consommations, prélèvements",
    "Revit": "Production des maquettes 3D",
    "SharePoint": "Accès et partage des documents",
    "Simple BIM": "Fusion des maquettes",
    "Sixsense Monitoring": "Mesures acoustiques (sonomètre)",
    "Sofistik Bridge": "Modélisation de la maquette tunnel",
    "TIPS": "Pointage Soletanche Bachy",
    "Tableau": "Commande traiteur",
    "Traktor": "Commande de gros engins",
    "Trimble Connect": "Dépôt et consultation des maquettes et plans",
    "Waste Marketplace (Wakio)": "Gestion des déchets",
    "Word": "Rédaction : contrats, avenants, courriers",
    "Yellow": "Suivi temps réel de l'avancement du tunnelier",
    # outils de l'onglet Tableau de Bord, hors table d'entretiens
    "CEMEX": None,
    "GMAO": "Gestion du stock (tunnel)",
    "HA+": "Achats opérationnels",
    "Visioaccès": "Création des badges et gestion des accès",
    # outils introduits par les corrections métier
    "OneDrive": None,
    "Probity": None,
    "Chorus": "Facturation par le maître d'ouvrage",
    "Sogelink/DICT": "Déclaration préalable avant terrassement",
    "RNDTS": "Traçabilité des déblais",
    "Remind": "Archivage des documents à 10 ans",
    "BYCN": None,
}

# ─────────────────────────────────────────────────────────────────────────────
# Colonne « Alimente quel outil » : ce qui désigne réellement un outil.
# Tout ce qui n'est pas ici est rejeté et listé dans le gap-report.
# ─────────────────────────────────────────────────────────────────────────────
CIBLES_VALIDES = {
    "Sharepoint": ["SharePoint"],
    "HARMONIE": ["Harmony"],
    "Harmonie": ["Harmony"],
    "Harmony": ["Harmony"],
    "Baseware": ["Basware"],
    "Powerbi": ["Power BI"],
    "Power BI": ["Power BI"],
    "Covadis": ["Covadis"],
    "cyclone 3d": ["Cyclone 3DR"],
    "cyclone 3d, Autocad": ["Cyclone 3DR", "AutoCAD"],
    "Inventaire via power bi": ["Power BI"],
    "Idcapture indirectement par des screenshot en 3d": ["IDCapture"],
}
# Cibles explicitement écartées, avec le motif porté au gap-report
CIBLES_REJETEES = {
    "Aucun": "cellule « Aucun »",
    "Tout": "cible générique, aucun outil nommé",
    "Tous": "cible générique, aucun outil nommé",
    "Non": "cellule « Non »",
    "Martin": "désigne une personne, pas un outil",
    "Relier à un autre logica bouyguesiel de compt": "cible non nommée",
    "Export manuel des pulsations pour les réinjecter dasn des tab internes":
        "cible non nommée (tableaux internes)",
    "Ponctuellement en fonction des demandes": "décrit une fréquence, pas une cible",
    "Avant POWER BI,": "formulation ambiguë, sens du flux indéterminé",
    "pixys tool(rapport de lever d'anneau, position machine) , mobidic":
        "cibles inconnues de la BDD (pixys tool, mobidic)",
    "BIP(chef de chantier qui deverse dans PUMA)":
        "sens inversé et repris par la correction « processus paye »",
}
# Flux documentés mais non tracés sur la carte macro (relation indirecte)
FLUX_HORS_CARTE = {("Trimble Connect", "IDCapture")}

# ─────────────────────────────────────────────────────────────────────────────
# CORRECTIONS MÉTIER — ajouts hors BDD, tracés « origine » dans carto.json
# ─────────────────────────────────────────────────────────────────────────────
HUB = ["SharePoint", "OneDrive", "Probity", "Mezzoteam"]

OUTILS_AJOUTES = [
    # (nom, origine, services, nature, interne/externe, alimentation, todo)
    ("OneDrive", "brief", [], "socle", None, None,
     "absent de la BDD — ajouté au titre du hub de données"),
    ("Probity", "brief", [], "socle", None, None,
     "absent de la BDD — ajouté au titre du hub de données"),
    ("Chorus", "correction", ["contrat"], "application", "externe", "manuel", None),
    ("Sogelink/DICT", "correction", ["qualite-env"], "application", "externe", "manuel", None),
    ("RNDTS", "correction", ["qualite-env"], "application", "externe", "manuel",
     "valider avec Camille"),
    ("Remind", "correction", ["methode-bim"], "socle", "interne", "manuel", None),
    ("BYCN", "correction", ["daf"], "application", "interne", None,
     "entité ou applicatif ? cible du pointage à préciser"),
    ("GMAO", "bdd-annexe", ["travaux", "travaux-tunnel"], "application", None, None, None),
    ("Visioaccès", "bdd-annexe", ["daf"], "application", None, "manuel", None),
    ("HA+", "bdd-annexe", ["travaux"], "application", None, None,
     "service déduit de la note « achat pablo ⇒ opérationnel », à confirmer"),
    ("CEMEX", "bdd-annexe", [], "application", "externe", None,
     "aucun service ni usage renseigné dans la BDD"),
]

FLUX_AJOUTES = [
    # (de, vers, mode, objet, todo)
    ("Harmony", "Basware", "auto", "Engagement des dépenses", None),
    ("Basware", "Harmony", "auto", "Retour de facturation", None),
    ("PABLO", "GMAO", "manuel", "Commande vers gestion de stock", None),
    ("GMAO", "PABLO", "manuel", "Réapprovisionnement", None),
    ("PUMA", "BIP/BAPS", "manuel", "Pointage des intérimaires", None),
    ("PUMA", "BYCN", "manuel", "Pointage des compagnons Bouygues", None),
    ("AutoCAD", "SharePoint", "manuel", "Dépôt des plans produits", None),
    ("SharePoint", "Mezzoteam", "manuel", "Diffusion au client", None),
    ("Mezzoteam", "Remind", "manuel", "Archivage 10 ans", None),
    ("Chorus", "Basware", "manuel", "Facturation du maître d'ouvrage", "flux à confirmer"),
]

# Processus ajoutés par les corrections, avec début et fin explicites
PROCESSUS_AJOUTES = {
    "daf": [{
        "nom": "Paye",
        "debut": "Pointage PUMA",
        "fin": "BIP (intérimaires) / BYCN (compagnons Bouygues)",
        "etapes": ["PUMA", "BIP/BAPS", "BYCN"],
        "origine": "correction",
        "ordre_atteste": True,
        "todo": None,
    }],
    "methode-bim": [{
        "nom": "Chaîne documentaire des plans",
        "debut": "AutoCAD (production du plan)",
        "fin": "Remind (archivage 10 ans)",
        "etapes": ["AutoCAD", "SharePoint", "Mezzoteam", "Remind"],
        "origine": "correction",
        "ordre_atteste": True,
        "todo": "l'étape « usage Travaux » entre SharePoint et Mezzoteam n'est portée par aucun outil de la BDD",
    }],
    "travaux": [{
        "nom": "Commande de MO",
        "debut": None,
        "fin": None,
        "etapes": [],
        "origine": "correction",
        "ordre_atteste": False,
        "todo": "mission ajoutée par correction — aucun outil rattaché dans la BDD",
    }],
    "contrat": [{
        "nom": "Facturation maître d'ouvrage",
        "debut": None,
        "fin": None,
        "etapes": ["Chorus"],
        "origine": "correction",
        "ordre_atteste": False,
        "todo": "début et fin non renseignés",
    }],
    "qualite-env": [{
        "nom": "Déclaration préalable avant terrassement",
        "debut": None,
        "fin": None,
        "etapes": ["Sogelink/DICT"],
        "origine": "correction",
        "ordre_atteste": False,
        "todo": "début et fin non renseignés",
    }, {
        "nom": "Traçabilité des déblais",
        "debut": None,
        "fin": None,
        "etapes": ["RNDTS"],
        "origine": "correction",
        "ordre_atteste": False,
        "todo": "valider avec Camille",
    }],
}

# Nombre de blocs processus vierges sur les affiches de service vides
BLOCS_VIERGES = 3


# ═════════════════════════════════════════════════════════════════════════════
# Utilitaires
# ═════════════════════════════════════════════════════════════════════════════

def slug(texte):
    """Identifiant DOM stable : accents dépouillés, minuscules, tirets.

    La plage de marques combinantes doit rester écrite en séquences
    d'échappement : avec les caractères combinants littéraux, la regex ne
    survit pas à une réécriture du fichier.
    """
    base = unicodedata.normalize("NFD", texte)
    base = re.sub("[̀-ͯ]", "", base)
    base = re.sub(r"[^A-Za-z0-9]+", "-", base).strip("-").lower()
    return base or "x"


def txt(valeur):
    """Cellule → chaîne nettoyée ; None et cellules vides deviennent ''."""
    if valeur is None:
        return ""
    return re.sub(r"\s+", " ", str(valeur)).strip()


def canon_outil(brut):
    """Libellé brut de la BDD → libellé normalisé de la cartographie."""
    brut = txt(brut)
    return CANON_OUTIL.get(brut, brut)


def majorite(votes, egalite, vide=None):
    """Dépouille un compteur de votes. Renvoie (valeur, a_tranche)."""
    if not votes:
        return vide, False
    ordre = votes.most_common()
    if len(ordre) > 1 and ordre[0][1] == ordre[1][1]:
        return egalite, False
    return ordre[0][0], True


# ═════════════════════════════════════════════════════════════════════════════
# Lecture de la BDD
# ═════════════════════════════════════════════════════════════════════════════

def lire_collecte(classeur):
    """Renvoie la liste des lignes d'entretien, sous forme de dictionnaires."""
    feuille = classeur[FEUILLE_COLLECTE]
    lignes = []
    for numero in range(PREMIERE_LIGNE, DERNIERE_LIGNE + 1):
        ligne = {cle: txt(feuille["%s%d" % (col, numero)].value)
                 for cle, col in COL.items()}
        if not ligne["outil"]:
            continue
        ligne["ligne"] = numero
        ligne["outil_brut"] = ligne["outil"]
        ligne["outil"] = canon_outil(ligne["outil"])
        lignes.append(ligne)
    return lignes


def lire_annexe(classeur, connus):
    """Outils de l'onglet Tableau de Bord absents de la table d'entretiens.

    Les libellés y portent parfois leur note entre parenthèses
    (« GMAO(gestion du stock => tunnel) ») : le nom et la note sont séparés.
    """
    feuille = classeur[FEUILLE_BORD]
    annexe = []
    for numero in range(4, feuille.max_row + 1):
        brut = txt(feuille["I%d" % numero].value)
        if not brut:
            continue
        note = txt(feuille["J%d" % numero].value)
        canon = canon_outil(brut)
        if canon in connus:
            continue
        # le libellé ne correspond à rien : il porte peut-être sa note en incise
        decoupe = re.match(r"^([^(]+?)\s*\((.+)\)\s*$", brut)
        if decoupe:
            canon = canon_outil(decoupe.group(1))
            note = note or decoupe.group(2)
            if canon in connus:
                continue
        annexe.append((numero, brut, note, canon))
    return annexe


# ═════════════════════════════════════════════════════════════════════════════
# Construction du modèle
# ═════════════════════════════════════════════════════════════════════════════

def construire_outils(lignes, ecarts):
    """Un enregistrement par outil distinct, avec ses arbitrages tracés."""
    votes_nature = defaultdict(Counter)
    votes_ie = defaultdict(Counter)
    votes_alim = defaultdict(Counter)
    services_par_outil = defaultdict(list)
    editeurs = defaultdict(Counter)
    lignes_par_outil = defaultdict(list)
    usages = defaultdict(list)

    for ligne in lignes:
        outil = ligne["outil"]
        lignes_par_outil[outil].append(ligne["ligne"])

        nature = ligne["nature"]
        if nature:
            if "Socle" in nature and "Application" in nature:
                votes_nature[outil]["socle"] += 1
                votes_nature[outil]["application"] += 1
            elif "Socle" in nature:
                votes_nature[outil]["socle"] += 1
            else:
                votes_nature[outil]["application"] += 1

        ie = ligne["ie"].lower()
        if ie in ("interne", "externe"):
            votes_ie[outil][ie] += 1

        alim = ligne["alim"]
        if alim:
            bas = alim.lower()
            auto = "automatique" in bas
            manuel = "manuel" in bas
            if auto:
                votes_alim[outil]["auto"] += 1
            if manuel:
                votes_alim[outil]["manuel"] += 1

        service = RATTACHEMENT.get(ligne["service"])
        if service is None:
            ecarts["service_inconnu"].append((ligne["ligne"], ligne["service"], outil))
        elif service not in services_par_outil[outil]:
            services_par_outil[outil].append(service)

        if ligne["editeur"]:
            editeurs[outil][ligne["editeur"]] += 1
        if ligne["usage"]:
            usages[outil].append(ligne["usage"])

    outils = OrderedDict()
    for nom in sorted(lignes_par_outil, key=lambda n: n.lower()):
        nature, tranche_nature = majorite(votes_nature[nom], "socle", "application")
        if nom in HUB:
            nature = "socle"
        ie, tranche_ie = majorite(votes_ie[nom], None, None)
        alim, _ = majorite(votes_alim[nom], "manuel", "manuel")

        todo = []
        if ie is None:
            todo.append("interne/externe non tranché par la BDD")
        fonction = FONCTION.get(nom)
        if not fonction:
            todo.append("fonction à renseigner")

        if len(votes_nature[nom]) > 1:
            ecarts["nature_contradictoire"].append(
                (nom, dict(votes_nature[nom]), nature))
        if len(votes_ie[nom]) > 1 or ie is None:
            ecarts["ie_contradictoire"].append((nom, dict(votes_ie[nom]), ie))

        outils[nom] = OrderedDict([
            ("nom", nom),
            ("id", slug(nom)),
            ("fonction", fonction),
            ("nature", nature),
            ("ie", ie),
            ("alim", alim),
            ("editeur", editeurs[nom].most_common(1)[0][0] if editeurs[nom] else None),
            ("services", services_par_outil[nom]),
            ("hub", nom in HUB),
            ("origine", "bdd"),
            ("lignes", lignes_par_outil[nom]),
            ("usages", sorted(set(usages[nom]))),
            ("todo", todo),
        ])
    return outils


def ajouter_outils_hors_bdd(outils, ecarts):
    """Applique OUTILS_AJOUTES : hub du brief, corrections métier, annexe."""
    for nom, origine, services, nature, ie, alim, todo in OUTILS_AJOUTES:
        if nom in outils:
            # l'outil existe déjà dans la table d'entretiens : on enrichit
            for service in services:
                if service not in outils[nom]["services"]:
                    outils[nom]["services"].append(service)
            continue
        entree = OrderedDict([
            ("nom", nom),
            ("id", slug(nom)),
            ("fonction", FONCTION.get(nom)),
            ("nature", nature),
            ("ie", ie),
            ("alim", alim),
            ("editeur", None),
            ("services", list(services)),
            ("hub", nom in HUB),
            ("origine", origine),
            ("lignes", []),
            ("usages", []),
            ("todo", []),
        ])
        if todo:
            entree["todo"].append(todo)
        if ie is None:
            entree["todo"].append("interne/externe non renseigné")
        if alim is None:
            entree["todo"].append("mode d'alimentation non renseigné")
        if not entree["fonction"]:
            entree["todo"].append("fonction à renseigner")
        if not services:
            entree["todo"].append("aucun service de rattachement")
        outils[nom] = entree
        ecarts["outils_hors_table"].append((nom, origine, todo))
    return OrderedDict(sorted(outils.items(), key=lambda kv: kv[0].lower()))


def construire_services(lignes, outils, ecarts):
    """Un enregistrement par service, avec ses processus et leurs étapes.

    La BDD ne porte aucune notion d'étape ordonnée : l'ordre des étapes est
    celui des lignes d'entretien, et chaque processus issu de la BDD est
    marqué « ordre non attesté ».
    """
    par_service = defaultdict(lambda: OrderedDict())
    for ligne in lignes:
        service = RATTACHEMENT.get(ligne["service"])
        if service is None or service in SERVICES_VIDES:
            continue
        brut = ligne["flux"]
        if brut in FLUX_NON_RENSEIGNE:
            ecarts["flux_vide"].append((ligne["ligne"], ligne["service"], ligne["outil"]))
            nom = "Processus non renseigné"
            deduit = True
        else:
            nom = CANON_PROCESSUS.get(brut)
            if nom is None:
                nom = brut.capitalize()
                ecarts["processus_non_harmonise"].append((ligne["ligne"], brut))
            deduit = False
        processus = par_service[service].setdefault(nom, {
            "nom": nom, "debut": None, "fin": None, "etapes": [],
            "origine": "bdd", "ordre_atteste": False,
            "deduit": deduit, "lignes": [], "todo": None,
        })
        processus["lignes"].append(ligne["ligne"])
        if ligne["outil"] not in [e["outil"] for e in processus["etapes"]]:
            processus["etapes"].append({
                "outil": ligne["outil"],
                "alim": outils[ligne["outil"]]["alim"],
                "ie": outils[ligne["outil"]]["ie"],
                "usage": ligne["usage"] or None,
            })

    services = []
    for sid, meta in SERVICES.items():
        liste = []
        for nom, processus in par_service.get(sid, {}).items():
            if processus["deduit"]:
                processus["todo"] = (
                    "processus déduit de la colonne Usage — la colonne FLUX "
                    "est vide sur ces lignes")
            liste.append(processus)
        for ajout in PROCESSUS_AJOUTES.get(sid, []):
            liste.append({
                "nom": ajout["nom"], "debut": ajout["debut"], "fin": ajout["fin"],
                "etapes": [{"outil": o,
                            "alim": outils[o]["alim"] if o in outils else None,
                            "ie": outils[o]["ie"] if o in outils else None,
                            "usage": None} for o in ajout["etapes"]],
                "origine": ajout["origine"],
                "ordre_atteste": ajout["ordre_atteste"],
                "deduit": False, "lignes": [], "todo": ajout["todo"],
            })
        vide = sid in SERVICES_VIDES
        if vide:
            liste = [{"nom": None, "debut": None, "fin": None, "etapes": [],
                      "origine": "vierge", "ordre_atteste": False, "deduit": False,
                      "lignes": [], "todo": "bloc vierge à compléter"}
                     for _ in range(BLOCS_VIERGES)]
        outils_service = [n for n, o in outils.items() if sid in o["services"]]
        services.append(OrderedDict([
            ("id", sid),
            ("nom", meta["nom"]),
            ("court", meta["court"]),
            ("kicker", meta["kicker"]),
            ("vide", vide),
            ("outils", sorted(outils_service, key=str.lower)),
            ("processus", liste),
        ]))
    return services


def construire_flux(lignes, outils, ecarts):
    """Colonne « Alimente quel outil » → flux orientés, plus les corrections.

    La BDD ne qualifie jamais l'automatisation d'un flux. La règle retenue,
    reprise du projet d'origine : un flux est automatisé quand son outil de
    destination est lui-même alimenté automatiquement.
    """
    flux = []
    vus = set()
    for ligne in lignes:
        brut = ligne["alimente_vers"]
        if not brut:
            ecarts["cible_vide"].append((ligne["ligne"], ligne["outil"]))
            continue
        if brut in CIBLES_REJETEES:
            ecarts["cible_rejetee"].append((ligne["ligne"], ligne["outil"], brut,
                                            CIBLES_REJETEES[brut]))
            continue
        cibles = CIBLES_VALIDES.get(brut)
        if cibles is None:
            ecarts["cible_rejetee"].append((ligne["ligne"], ligne["outil"], brut,
                                            "libellé non reconnu comme un outil"))
            continue
        for cible in cibles:
            if cible == ligne["outil"]:
                continue
            cle = (ligne["outil"], cible)
            if cle in vus:
                continue
            vus.add(cle)
            mode = "auto" if outils.get(cible, {}).get("alim") == "auto" else "manuel"
            flux.append(OrderedDict([
                ("de", ligne["outil"]),
                ("vers", cible),
                ("mode", mode),
                ("objet", ligne["flux"] and CANON_PROCESSUS.get(ligne["flux"], ligne["flux"]) or None),
                ("carte", cle not in FLUX_HORS_CARTE),
                ("origine", "bdd"),
                ("ligne", ligne["ligne"]),
                ("todo", None),
            ]))

    for de, vers, mode, objet, todo in FLUX_AJOUTES:
        cle = (de, vers)
        if cle in vus:
            for existant in flux:
                if (existant["de"], existant["vers"]) == cle:
                    existant["objet"] = objet
                    existant["mode"] = mode
                    existant["origine"] = "bdd+correction"
            continue
        vus.add(cle)
        flux.append(OrderedDict([
            ("de", de), ("vers", vers), ("mode", mode), ("objet", objet),
            ("carte", cle not in FLUX_HORS_CARTE),
            ("origine", "correction"), ("ligne", None), ("todo", todo),
        ]))

    for f in flux:
        for bout in ("de", "vers"):
            if f[bout] not in outils:
                ecarts["flux_orphelin"].append((f["de"], f["vers"], f[bout]))
    return flux


# ═════════════════════════════════════════════════════════════════════════════
# Gap-report
# ═════════════════════════════════════════════════════════════════════════════

def ecrire_gap_report(lignes, outils, services, flux, annexe, ecarts, natures_brutes):
    """Tout ce que la BDD ne dit pas, ou dit deux fois."""
    L = []
    a = L.append
    a("# Gap-report — cartographie IT chantier")
    a("")
    a("Généré par `data/extract_bdd.py` le %s, depuis `%s`."
      % (date.today().isoformat(), BDD.name))
    a("Ce rapport ne signale pas des défauts de génération : il liste ce que la")
    a("base d'entretiens ne permet pas de trancher, et ce qui a été ajouté sur")
    a("décision métier hors de la base.")
    a("")
    a("> Les numéros de ligne sont ceux de la feuille Excel (l'en-tête est en")
    a("> ligne 3, la première ligne de données en ligne 4), et non la colonne `N°`.")
    a("")

    a("## 1. Colonne « socle commun »")
    a("")
    a("La BDD **ne porte pas** de champ booléen `socle_commun`. La seule colonne")
    a("qui qualifie la nature d'un outil est **`Nature` (colonne F)** :")
    a("")
    a("| Valeur | Lignes |")
    a("|---|---:|")
    for valeur, nombre in natures_brutes.most_common():
        a("| `%s` | %d |" % (valeur or "*(vide)*", nombre))
    a("")
    a("Elle est exploitable mais **contradictoire d'un entretien à l'autre**")
    a("(section 8). Règle d'arbitrage retenue : *majorité des entretiens ;")
    a("égalité → socle ; les 4 outils du hub forcés socle*.")
    a("")

    a("## 2. Outils sans service")
    a("")
    if ecarts["service_inconnu"]:
        a("| Ligne | Libellé service | Outil |")
        a("|---:|---|---|")
        for numero, libelle, outil in ecarts["service_inconnu"]:
            a("| %d | `%s` | %s |" % (numero, libelle or "*(vide)*", outil))
    else:
        a("Aucune ligne de la table d'entretiens n'est privée de service : les")
        a("11 libellés bruts sont tous rattachés à l'un des 9 services.")
    a("")
    a("En revanche, l'onglet `📊 Tableau de Bord` liste **%d outils hors de la"
      % len(annexe))
    a("table d'entretiens**, sans service, sans usage et sans flux :")
    a("")
    a("| Ligne | Libellé brut | Note | Retenu sous | Services |")
    a("|---:|---|---|---|---|")
    for numero, brut, note, canon in annexe:
        entree = outils.get(canon)
        rattachement = "—"
        if entree:
            rattachement = ", ".join(SERVICES[s]["nom"] for s in entree["services"]) or "**aucun**"
        a("| %d | `%s` | %s | %s | %s |"
          % (numero, brut, note or "—", canon, rattachement))
    a("")

    a("## 3. Flux sans source ou cible")
    a("")
    a("Sur %d lignes d'entretien, la colonne « Alimente quel outil » donne :" % len(lignes))
    a("")
    a("- **%d cellules vides** — aucun flux déclaré ;" % len(ecarts["cible_vide"]))
    rejets = Counter(motif for _, _, _, motif in ecarts["cible_rejetee"])
    for motif, nombre in rejets.most_common():
        a("- **%d** — %s ;" % (nombre, motif))
    a("")
    a("Détail des cibles écartées qui ne sont ni `Aucun` ni vides :")
    a("")
    a("| Ligne | Outil source | Cible déclarée | Motif |")
    a("|---:|---|---|---|")
    for numero, outil, brut, motif in ecarts["cible_rejetee"]:
        if brut in ("Aucun", "Non"):
            continue
        a("| %d | %s | `%s` | %s |" % (numero, outil, brut, motif))
    a("")
    retenus = [f for f in flux if f["origine"] != "correction"]
    a("**%d flux retenus depuis la BDD**, %d ajoutés par les corrections métier, "
      "soit %d au total." % (len(retenus), len(flux) - len(retenus), len(flux)))
    a("")
    a("La BDD **ne qualifie jamais l'automatisation d'un flux**. Règle retenue :")
    a("un flux est automatisé quand son outil de destination est lui-même")
    a("alimenté automatiquement.")
    a("")

    a("## 4. Sens de flux corrigé")
    a("")
    a("Ligne 52 (`N° 49`) — la cellule `BIP(chef de chantier qui deverse dans PUMA)` porte")
    a("PUMA comme source alors que le texte décrit **BIP → PUMA**. La correction")
    a("métier « processus paye » tranche dans l'autre sens encore : *pointage")
    a("PUMA → BIP (intérimaires) / BYCN (compagnons Bouygues)*. **C'est la")
    a("correction qui a été appliquée**, la ligne 52 est écartée.")
    a("")

    a("## 5. Processus sans début ni fin")
    a("")
    a("La BDD ne porte **aucune** notion d'étape ordonnée : la colonne `FLUX`")
    a("nomme des thèmes, pas des chaînes. Conséquences sur le livrable 2 :")
    a("")
    sans_bornes = [(s["nom"], p["nom"]) for s in services for p in s["processus"]
                   if p["origine"] == "bdd" and not p["debut"] and not p["fin"]]
    a("- **%d processus** issus de la BDD n'ont ni début ni fin → bloc vide"
      % len(sans_bornes))
    a("  dimensionné et éditable, marqué `TODO` ;")
    a("- l'**ordre des étapes** de ces processus est celui des lignes")
    a("  d'entretien, il n'est attesté par rien ;")
    a("- seuls les processus issus des corrections (*Paye*, *Chaîne")
    a("  documentaire des plans*) portent un ordre et des bornes explicites.")
    a("")

    a("## 6. Colonne FLUX vide")
    a("")
    par_service = Counter(libelle for _, libelle, _ in ecarts["flux_vide"])
    a("**%d lignes** n'ont aucune valeur de FLUX :" % len(ecarts["flux_vide"]))
    a("")
    a("| Service | Lignes |")
    a("|---|---:|")
    for libelle, nombre in par_service.most_common():
        a("| %s | %d |" % (libelle, nombre))
    a("")
    a("Leurs outils sont regroupés sous un processus « Processus non renseigné »")
    a("marqué `TODO`, dont les étapes portent l'usage déclaré en BDD.")
    a("")

    a("## 7. Doublons de nommage")
    a("")
    reecrits = sum(1 for brut, retenu in CANON_OUTIL.items() if brut != retenu)
    a("Aucun doublon strict. **%d libellés réécrits**, dont cinq qui créaient"
      % reecrits)
    a("de vrais écarts, pas seulement de la casse :")
    a("")
    a("| Libellé BDD | Retenu | Effet |")
    a("|---|---|---|")
    a("| `Baseware` | **Basware** | orthographe de l'éditeur |")
    a("| `HARMONIE` / `Harmonie` (colonne cible) | **Harmony** | 6 flux pointaient vers un outil inexistant |")
    a("| `Cabine de pilotage(Power BI)` | **Power BI** | rapport fusionné dans l'outil |")
    a("| `Achat +` | **HA+** | correction métier |")
    a("| `Neoacces` | **Visioaccès** | correction métier |")
    a("")
    a("Le reste n'est que de la casse et des espaces (`POWER BI`, `IDCAPTURE`,")
    a("`Trimbleconnect`, `Sharepoint`, `Docusign`, `Autocad`, `Puma`, `Pablo`…).")
    a("")

    a("## 8. Contradictions sur la colonne `Nature`")
    a("")
    a("**%d outils** sont qualifiés différemment selon l'entretien :"
      % len(ecarts["nature_contradictoire"]))
    a("")
    a("| Outil | Votes | Retenu |")
    a("|---|---|---|")
    for nom, votes, retenu in sorted(ecarts["nature_contradictoire"]):
        detail = " · ".join("%s ×%d" % (k, v) for k, v in sorted(votes.items()))
        a("| %s | %s | **%s** |" % (nom, detail, retenu))
    a("")

    a("## 9. Contradictions sur `Interne / Externe`")
    a("")
    interne = sum(1 for o in outils.values() if o["ie"] == "interne")
    externe = sum(1 for o in outils.values() if o["ie"] == "externe")
    inconnu = sum(1 for o in outils.values() if o["ie"] is None)
    a("Après arbitrage : **%d internes · %d externes · %d `TODO`**."
      % (interne, externe, inconnu))
    a("Les `TODO` forment une troisième part dans le graphique de la synthèse ;")
    a("aucune valeur n'est devinée.")
    a("")
    a("| Outil | Votes | Retenu |")
    a("|---|---|---|")
    for nom, votes, retenu in sorted(ecarts["ie_contradictoire"]):
        detail = " · ".join("%s ×%d" % (k, v) for k, v in sorted(votes.items())) or "*aucun*"
        a("| %s | %s | **%s** |" % (nom, detail, retenu or "TODO"))
    a("")

    a("## 10. Colonnes non exploitées")
    a("")
    a("- **`Interfaces / Connexions` (O)** est polluée : des niveaux de maturité")
    a("  (`3 – Mature`, `4 – Optimisé`) et 10 valeurs `Fluide` y ont été saisis à")
    a("  la place de la colonne P. Inexploitable.")
    a("- **`Niveau Maturité` (P)** et **`Fréquence d'utilisation` (L)** sont")
    a("  écartées sur consigne : ni extraites, ni affichées nulle part.")
    a("- **`Alimenté par qui` (M)** nomme des personnes et des équipes, pas des")
    a("  outils : elle ne peut pas servir à tracer des flux entrants.")
    a("- **`Colonne32` (R)** est renseignée sur 1 ligne sur %d." % len(lignes))
    a("")

    a("## 11. Éléments ajoutés hors BDD")
    a("")
    a("Chacun porte un champ `origine` dans `carto.json` et un badge dans le")
    a("référentiel. `brief` = imposé par le cahier des charges ; `correction` =")
    a("section CORRECTIONS MÉTIER ; `bdd-annexe` = onglet Tableau de Bord.")
    a("")
    a("| Outil | Origine | Services | TODO |")
    a("|---|---|---|---|")
    for nom, origine, todo in ecarts["outils_hors_table"]:
        entree = outils[nom]
        noms = ", ".join(SERVICES[s]["nom"] for s in entree["services"]) or "—"
        a("| %s | `%s` | %s | %s |" % (nom, origine, noms, todo or "—"))
    a("")
    ajoutes = [f for f in flux if f["origine"] in ("correction", "bdd+correction")]
    a("Flux ajoutés ou requalifiés par les corrections : **%d**." % len(ajoutes))
    a("")
    a("| De | Vers | Mode | Objet |")
    a("|---|---|---|---|")
    for f in ajoutes:
        a("| %s | %s | %s | %s |" % (f["de"], f["vers"], f["mode"], f["objet"] or "—"))
    a("")

    a("## 12. Reste à faire (`TODO` visibles dans les livrables)")
    a("")
    total = sum(len(o["todo"]) for o in outils.values())
    a("- **%d `TODO` d'outils** (fonction, interne/externe, alimentation, service) ;" % total)
    a("- **%d processus sans début ni fin** ;" % len(sans_bornes))
    a("- le service **Sécurité** est volontairement vide : %d blocs processus"
      % BLOCS_VIERGES)
    a("  vierges, éditables, aucun outil prélevé aux autres services ;")
    a("- **RNDTS** : traçabilité des déblais — *valider avec Camille* ;")
    a("- **BYCN** : entité ou applicatif ? cible du pointage à préciser ;")
    a("- **CEMEX** : aucun service ni usage renseigné.")
    a("")
    SORTIE_GAP.write_text("\n".join(L) + "\n", encoding="utf-8")


# ═════════════════════════════════════════════════════════════════════════════
# Point d'entrée
# ═════════════════════════════════════════════════════════════════════════════

LEGENDE = OrderedDict([
    ("nature", [
        {"cle": "socle", "libelle": "Socle de données",
         "aide": "Plateforme transverse, partagée par plusieurs services"},
        {"cle": "application", "libelle": "Application métier",
         "aide": "Outil au service d'un métier précis"},
    ]),
    ("ie", [
        {"cle": "interne", "libelle": "Interne", "aide": "Développé par le groupe"},
        {"cle": "externe", "libelle": "Externe", "aide": "Fourni par un éditeur tiers"},
        {"cle": None, "libelle": "TODO", "aide": "Non tranché par la BDD"},
    ]),
    ("alim", [
        {"cle": "auto", "libelle": "Automatique", "aide": "Trait plein"},
        {"cle": "manuel", "libelle": "Manuel", "aide": "Trait pointillé"},
    ]),
    ("origine", [
        {"cle": "bdd", "libelle": "BDD", "aide": "Table d'entretiens"},
        {"cle": "bdd-annexe", "libelle": "BDD annexe", "aide": "Onglet Tableau de Bord"},
        {"cle": "correction", "libelle": "Correction", "aide": "Correction métier hors BDD"},
        {"cle": "brief", "libelle": "Brief", "aide": "Imposé par le cahier des charges"},
    ]),
])


def main():
    if not BDD.exists():
        sys.exit("BDD introuvable : %s" % BDD)

    classeur = load_workbook(BDD, data_only=True, read_only=True)
    lignes = lire_collecte(classeur)
    annexe = lire_annexe(classeur, {l["outil"] for l in lignes})
    natures_brutes = Counter(l["nature"] for l in lignes)

    ecarts = defaultdict(list)
    outils = construire_outils(lignes, ecarts)
    outils = ajouter_outils_hors_bdd(outils, ecarts)
    services = construire_services(lignes, outils, ecarts)
    flux = construire_flux(lignes, outils, ecarts)

    modele = OrderedDict([
        ("meta", OrderedDict([
            ("titre", "Cartographie des outils IT — chantier Bouygues Construction"),
            ("source", BDD.name),
            ("lignes_bdd", len(lignes)),
            ("genere_le", date.today().isoformat()),
            ("format_pt", [960, 540]),
        ])),
        ("charte", CHARTE),
        ("hub", HUB),
        ("services", services),
        ("outils", outils),
        ("flux", flux),
        ("legende", LEGENDE),
    ])
    SORTIE_JSON.write_text(
        json.dumps(modele, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    ecrire_gap_report(lignes, outils, services, flux, annexe, ecarts, natures_brutes)

    socle = sum(1 for o in outils.values() if o["nature"] == "socle")
    processus = sum(len(s["processus"]) for s in services)
    todos = sum(len(o["todo"]) for o in outils.values())
    print("%s : %d lignes lues" % (BDD.name, len(lignes)))
    print("%s : %d services, %d processus, %d outils (%d socle), %d flux"
          % (SORTIE_JSON.name, len(services), processus, len(outils), socle, len(flux)))
    print("%s : %d TODO d'outils, %d cibles de flux écartées"
          % (SORTIE_GAP.name, todos, len(ecarts["cible_rejetee"])))
    if ecarts["flux_orphelin"]:
        for de, vers, manquant in ecarts["flux_orphelin"]:
            print("  flux orphelin : %s → %s (%s inconnu)" % (de, vers, manquant),
                  file=sys.stderr)
    if ecarts["service_inconnu"]:
        for numero, libelle, outil in ecarts["service_inconnu"]:
            print("  ligne %d : service « %s » non rattaché (%s)"
                  % (numero, libelle, outil), file=sys.stderr)


if __name__ == "__main__":
    main()
