#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Contrôles automatiques sur la source et sur les deux livrables.

    python3 build/verifier.py

Reprend les critères d'acceptation qui peuvent se vérifier sans regarder :
cohérence de carto.json, absence d'appel réseau dans le HTML, formes natives
et liens hypertexte dans le PPTX, et surtout le fait qu'aucun outil affiché
n'échappe à la source unique.

Ce qui ne se vérifie qu'à l'œil — débordements, chevauchements, métriques de
police — n'est pas ici : il faut regarder les captures des 12 pages HTML et
les 12 pages du PDF converti.

Sortie : une ligne par contrôle, code de retour non nul si l'un échoue.
"""

import re
import sys
import zipfile
from pathlib import Path

import layout
import model

RACINE = Path(__file__).resolve().parent.parent
HTML = RACINE / "out" / "carto.html"
PPTX = RACINE / "out" / "carto.pptx"
VENDOR = RACINE / "out" / "vendor" / "chart.min.js"

ORIGINES = {"bdd", "bdd-annexe", "correction", "brief"}
MODES = {"auto", "manuel"}
PAGES_ATTENDUES = 12


class Journal(object):
    def __init__(self):
        self.echecs = 0

    def controle(self, titre, ok, detail=""):
        marque = "ok  " if ok else "ÉCHEC"
        print("%-5s %s%s" % (marque, titre, (" — " + detail) if detail else ""))
        if not ok:
            self.echecs += 1


def verifier_source(journal, carto):
    outils = carto["outils"]

    inconnus = set()
    sans_mode = []
    reflexifs = []
    for flux in carto["flux"]:
        for bout in ("de", "vers"):
            if flux[bout] not in outils:
                inconnus.add(flux[bout])
        if flux["mode"] not in MODES:
            sans_mode.append("%s→%s" % (flux["de"], flux["vers"]))
        if flux["de"] == flux["vers"]:
            reflexifs.append(flux["de"])
    journal.controle("Chaque extrémité de flux est un outil connu",
                     not inconnus, ", ".join(sorted(inconnus)))
    journal.controle("Chaque flux porte un mode auto/manuel",
                     not sans_mode, ", ".join(sans_mode))
    journal.controle("Aucun flux ne boucle sur lui-même",
                     not reflexifs, ", ".join(reflexifs))

    mauvaises = sorted({o["origine"] for o in outils.values()} - ORIGINES)
    journal.controle("Origine d'outil dans le jeu attendu",
                     not mauvaises, ", ".join(mauvaises))

    tracables = [n for n, o in outils.items()
                 if o["origine"] == "bdd" and not o["lignes"]]
    journal.controle("Chaque outil « bdd » cite ses lignes d'entretien",
                     not tracables, ", ".join(tracables))

    orphelins = []
    for service in carto["services"]:
        for nom in service["outils"]:
            if nom not in outils:
                orphelins.append("%s/%s" % (service["id"], nom))
        for processus in service["processus"]:
            for etape in processus["etapes"]:
                if etape["outil"] not in outils:
                    orphelins.append("%s/%s" % (service["id"], etape["outil"]))
    journal.controle("Chaque outil cité par un service existe",
                     not orphelins, ", ".join(orphelins))

    hub_absent = [n for n in carto["hub"] if n not in outils]
    journal.controle("Les outils du hub existent", not hub_absent,
                     ", ".join(hub_absent))


def verifier_scene(journal, carto):
    """Aucun outil affiché ne doit échapper à carto.json."""
    pages = layout.construire(carto)
    journal.controle("Le graphe de scène compte %d pages" % PAGES_ATTENDUES,
                     len(pages) == PAGES_ATTENDUES, "%d trouvées" % len(pages))

    connus = {o["id"] for o in carto["outils"].values()}
    inconnus = set()
    for page in pages:
        for forme in page["formes"]:
            classe = forme.get("classe") or ""
            if "pastille" in classe or "etape" in classe:
                cle = forme.get("cle")
                if cle and cle not in connus:
                    inconnus.add(cle)
    journal.controle("Aucune pastille affichée hors de carto.json",
                     not inconnus, ", ".join(sorted(inconnus)))

    liens = {f["lien"] for page in pages for f in page["formes"]
             if f.get("lien") is not None}
    hors_bornes = {l for l in liens if not 0 <= l < len(pages)}
    journal.controle("Tous les liens visent une page existante",
                     not hors_bornes, str(sorted(hors_bornes)))
    return pages


def verifier_html(journal):
    if not HTML.exists():
        journal.controle("out/carto.html existe", False)
        return
    contenu = HTML.read_text(encoding="utf-8")
    urls = sorted(set(re.findall(r"https?://[^\"'\s<>)]+", contenu)))
    journal.controle("Aucune URL réseau dans le HTML", not urls,
                     ", ".join(urls[:3]))
    journal.controle("Chart.js est vendorisé", VENDOR.exists(),
                     "" if VENDOR.exists() else "out/vendor/chart.min.js absent")
    journal.controle("Le HTML référence le Chart.js local",
                     'src="vendor/chart.min.js"' in contenu)
    pages = contenu.count('<section class="page"')
    journal.controle("Le HTML compte %d pages" % PAGES_ATTENDUES,
                     pages == PAGES_ATTENDUES, "%d trouvées" % pages)


def verifier_pptx(journal):
    if not PPTX.exists():
        journal.controle("out/carto.pptx existe", False)
        return
    archive = zipfile.ZipFile(PPTX)
    diapos = sorted(
        (n for n in archive.namelist()
         if re.match(r"ppt/slides/slide\d+\.xml$", n)),
        key=lambda n: int(re.search(r"\d+", n.rsplit("/", 1)[1]).group()))
    journal.controle("Le PPTX compte %d diapositives" % PAGES_ATTENDUES,
                     len(diapos) == PAGES_ATTENDUES, "%d trouvées" % len(diapos))

    images = [n for n in archive.namelist() if n.startswith("ppt/media/")]
    journal.controle("Aucune image collée dans le PPTX", not images,
                     ", ".join(images[:3]))

    sauts, sans_retour = 0, []
    for numero, nom in enumerate(diapos, start=1):
        xml = archive.read(nom).decode("utf-8")
        compte = xml.count("ppaction://hlinksldjump")
        sauts += compte
        if numero > 1 and compte == 0:
            sans_retour.append(nom.rsplit("/", 1)[1])
    journal.controle("Chaque diapositive hors macro a un lien de retour",
                     not sans_retour, ", ".join(sans_retour))
    journal.controle("Le PPTX porte des liens de saut natifs", sauts > 0,
                     "%d liens" % sauts)

    premiere = archive.read(diapos[0]).decode("utf-8")
    journal.controle("La vue macro trace les flux en géométrie personnalisée",
                     "<a:custGeom>" in premiere,
                     "%d courbes" % premiere.count("<a:custGeom>"))
    ombres = sum(archive.read(n).decode("utf-8").count("<a:outerShdw")
                 for n in diapos)
    journal.controle("Aucune ombre portée", ombres == 0, "%d trouvées" % ombres)

    graphiques = [n for n in archive.namelist()
                  if re.match(r"ppt/charts/chart\d+\.xml$", n)]
    journal.controle("Les graphiques de la synthèse sont natifs",
                     len(graphiques) == 4, "%d trouvés" % len(graphiques))
    classeurs = [n for n in archive.namelist()
                 if n.startswith("ppt/embeddings/")]
    journal.controle("Les données des graphiques sont éditables",
                     len(classeurs) == len(graphiques),
                     "%d classeurs intégrés" % len(classeurs))


def main():
    journal = Journal()
    carto = model.charger()
    print("Source : %s, généré le %s"
          % (carto["meta"]["source"], carto["meta"]["genere_le"]))
    print()
    verifier_source(journal, carto)
    verifier_scene(journal, carto)
    verifier_html(journal)
    verifier_pptx(journal)

    global_ = model.indicateurs(carto)
    print()
    print("Décompte : %d services, %d processus, %d outils (%d socle), %d flux"
          % (len(carto["services"]),
             sum(len(s["processus"]) for s in carto["services"]),
             global_["nb_outils"], global_["nb_socle"], global_["nb_flux"]))
    print()
    if journal.echecs:
        print("%d contrôle(s) en échec." % journal.echecs)
        return 1
    print("Tous les contrôles automatiques passent.")
    print("Reste à regarder : les 12 captures HTML et les 12 pages du PDF "
          "converti depuis le PPTX.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
