#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Chargement de data/carto.json et calcul des indicateurs de la synthèse.

Aucun contenu métier n'est écrit ici : le module ne fait que lire la source
unique et en dériver des compteurs. Changer carto.json suffit à changer les
chiffres, sur les deux livrables à la fois.
"""

import json
from collections import Counter, OrderedDict
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent
SOURCE = RACINE / "data" / "carto.json"

TOUS = "__tous__"          # clé du périmètre « Tous services »
LIBELLE_TOUS = "Tous services"


def charger(chemin=SOURCE):
    """Lit carto.json et l'enrichit d'index pratiques."""
    with open(chemin, encoding="utf-8") as flux:
        carto = json.load(flux, object_pairs_hook=OrderedDict)
    carto["_services_par_id"] = {s["id"]: s for s in carto["services"]}
    return carto


def outils_du_service(carto, sid):
    """Noms des outils rattachés à un service, dans l'ordre alphabétique."""
    return carto["_services_par_id"][sid]["outils"]


def flux_du_service(carto, sid):
    """Flux dont au moins une extrémité est un outil du service."""
    perimetre = set(outils_du_service(carto, sid))
    return [f for f in carto["flux"]
            if f["de"] in perimetre or f["vers"] in perimetre]


def _repartition(valeurs, cles):
    """Compte les valeurs et renvoie un dictionnaire ordonné complet."""
    compte = Counter(valeurs)
    return OrderedDict((cle, compte.get(cle, 0)) for cle in cles)


def indicateurs(carto, sid=TOUS):
    """Les six métriques de la page de synthèse, pour un périmètre donné.

    `sid` vaut TOUS pour la vue globale, ou l'identifiant d'un service.
    """
    if sid == TOUS:
        noms = list(carto["outils"])
        flux = list(carto["flux"])
        libelle = LIBELLE_TOUS
    else:
        noms = outils_du_service(carto, sid)
        flux = flux_du_service(carto, sid)
        libelle = carto["_services_par_id"][sid]["nom"]

    outils = [carto["outils"][n] for n in noms]
    hub = set(carto["hub"])
    socle = [o for o in outils if o["nature"] == "socle" or o["nom"] in hub]

    modes = _repartition((f["mode"] for f in flux), ["auto", "manuel"])
    total_modes = sum(modes.values())
    natures = _repartition((o["ie"] or "todo" for o in outils),
                           ["interne", "externe", "todo"])

    return OrderedDict([
        ("id", sid),
        ("libelle", libelle),
        ("nb_flux", len(flux)),
        ("nb_outils", len(outils)),
        ("nb_socle", len(socle)),
        ("flux_mode", modes),
        ("flux_mode_pct", OrderedDict(
            (cle, round(100.0 * valeur / total_modes, 1) if total_modes else 0.0)
            for cle, valeur in modes.items())),
        ("outils_ie", natures),
        ("classement", classement(carto)),
    ])


def classement(carto):
    """Services triés par nombre de flux décroissant, puis par nom."""
    lignes = []
    for service in carto["services"]:
        lignes.append(OrderedDict([
            ("id", service["id"]),
            ("nom", service["nom"]),
            ("nb_flux", len(flux_du_service(carto, service["id"]))),
            ("nb_outils", len(service["outils"])),
        ]))
    lignes.sort(key=lambda l: (-l["nb_flux"], l["nom"]))
    return lignes


def tous_les_indicateurs(carto):
    """Vue globale + une vue par service, prêtes à être sérialisées."""
    vues = OrderedDict()
    vues[TOUS] = indicateurs(carto, TOUS)
    for service in carto["services"]:
        vues[service["id"]] = indicateurs(carto, service["id"])
    return vues


def outils_tries(carto):
    """Référentiel : tous les outils, triés sans tenir compte de la casse."""
    return [carto["outils"][n] for n in sorted(carto["outils"], key=str.lower)]


if __name__ == "__main__":
    carto = charger()
    global_ = indicateurs(carto)
    print("flux %d · outils %d · socle %d"
          % (global_["nb_flux"], global_["nb_outils"], global_["nb_socle"]))
    print("auto/manuel : %s" % dict(global_["flux_mode_pct"]))
    print("interne/externe/todo : %s" % dict(global_["outils_ie"]))
    for ligne in global_["classement"]:
        print("  %-22s %2d flux  %2d outils"
              % (ligne["nom"], ligne["nb_flux"], ligne["nb_outils"]))
