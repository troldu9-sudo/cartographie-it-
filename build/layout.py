#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Géométrie commune aux deux livrables.

Ce module ne dessine rien : il transforme carto.json en un graphe de scène,
c'est-à-dire une liste de pages contenant des formes primitives positionnées
au point près. generate_html.py et generate_pptx.py se contentent ensuite de
rendre ces primitives dans leur format respectif.

C'est ce qui garantit que le PPTX ne peut pas diverger du HTML : il n'y a
qu'une seule mise en page, écrite ici.

Repère : 960 × 540 points (16:9). 1 pt = 1,3333 px à l'écran, 12700 EMU en
PowerPoint. L'origine est en haut à gauche.

Primitives émises
    rect     rectangle, éventuellement à coins arrondis, avec ou sans texte
    texte    bloc de texte sans fond
    bezier   courbe de Bézier cubique unique, avec ou sans flèche
    graphe   emplacement d'un graphique (canvas en HTML, chart natif en PPTX)
    table    tableau du référentiel, rendu ligne à ligne par chaque générateur
"""

import math
from collections import OrderedDict

import model

PAGE_W, PAGE_H = 960.0, 540.0
MARGE = 22.0
HAUT_BANDEAU = 46.0
HAUT_PIED = 24.0

# Chromium compose en DejaVu Sans, PowerPoint en Calibri : tout texte exporté
# peut être plus large que mesuré. Les largeurs ci-dessous sont celles de
# Calibri ; le facteur les ramène au pire cas.
MARGE_POLICE = 1.22

# Avances de Calibri, en millièmes de cadratin.
_AVANCES = {
    " ": 226, "!": 296, '"': 400, "#": 507, "$": 507, "%": 750, "&": 655,
    "'": 216, "(": 303, ")": 303, "*": 423, "+": 498, ",": 250, "-": 306,
    ".": 252, "/": 386, ":": 268, ";": 268, "<": 498, "=": 498, ">": 498,
    "?": 445, "@": 850, "[": 303, "\\": 386, "]": 303, "^": 498, "_": 498,
    "`": 306, "{": 303, "|": 227, "}": 303, "~": 498, "·": 252, "—": 1000,
    "–": 507, "«": 456, "»": 456, "→": 838, "▸": 600, "◀": 600, "×": 498,
    "A": 579, "B": 544, "C": 533, "D": 615, "E": 488, "F": 459, "G": 631,
    "H": 623, "I": 252, "J": 319, "K": 520, "L": 420, "M": 855, "N": 646,
    "O": 662, "P": 517, "Q": 673, "R": 543, "S": 459, "T": 487, "U": 642,
    "V": 567, "W": 890, "X": 519, "Y": 487, "Z": 468,
    "a": 479, "b": 525, "c": 423, "d": 525, "e": 498, "f": 305, "g": 471,
    "h": 525, "i": 230, "j": 239, "k": 455, "l": 230, "m": 799, "n": 525,
    "o": 527, "p": 525, "q": 525, "r": 349, "s": 391, "t": 335, "u": 525,
    "v": 452, "w": 715, "x": 433, "y": 453, "z": 395,
}
for _chiffre in "0123456789":
    _AVANCES[_chiffre] = 507
_AVANCE_DEFAUT = 520


def largeur(texte, taille):
    """Largeur d'un texte, en points, au pire cas des deux polices."""
    total = 0
    for caractere in texte:
        total += _AVANCES.get(caractere, _AVANCE_DEFAUT)
    return total / 1000.0 * taille * MARGE_POLICE


def tronquer(texte, taille, disponible):
    """Coupe un texte à la largeur disponible, en ajoutant une ellipse."""
    if largeur(texte, taille) <= disponible:
        return texte
    coupe = texte
    while coupe and largeur(coupe + "…", taille) > disponible:
        coupe = coupe[:-1]
    return (coupe.rstrip() + "…") if coupe else ""


# ═════════════════════════════════════════════════════════════════════════════
# Graphe de scène
# ═════════════════════════════════════════════════════════════════════════════

class Scene(object):
    """Accumulateur de formes pour une page."""

    def __init__(self, carto, identifiant, titre, genre, indice):
        self.carto = carto
        self.c = carto["charte"]
        self.id = identifiant
        self.titre = titre
        self.genre = genre
        self.indice = indice
        self.formes = []
        self.ancres = {}      # nom d'outil → (x, y, w, h) de sa pastille de référence

    def rect(self, x, y, w, h, **options):
        forme = OrderedDict([
            ("type", "rect"), ("x", x), ("y", y), ("w", w), ("h", h),
            ("fill", options.get("fill")),
            ("line", options.get("line")),
            ("lw", options.get("lw", 0.75)),
            ("dash", options.get("dash")),
            ("radius", options.get("radius", 0.0)),
            ("texte", options.get("texte")),
            ("taille", options.get("taille", 8.0)),
            ("gras", options.get("gras", False)),
            ("couleur", options.get("couleur", self.c["encre"])),
            ("align", options.get("align", "center")),
            ("valign", options.get("valign", "middle")),
            ("pad", options.get("pad", 3.0)),
            ("nowrap", options.get("nowrap", True)),
            ("lien", options.get("lien")),
            ("classe", options.get("classe")),
            ("cle", options.get("cle")),
            ("info", options.get("info")),
        ])
        self.formes.append(forme)
        return forme

    def texte(self, x, y, w, h, contenu, **options):
        forme = OrderedDict([
            ("type", "texte"), ("x", x), ("y", y), ("w", w), ("h", h),
            ("texte", contenu),
            ("taille", options.get("taille", 8.0)),
            ("gras", options.get("gras", False)),
            ("italique", options.get("italique", False)),
            ("couleur", options.get("couleur", self.c["encre"])),
            ("align", options.get("align", "left")),
            ("valign", options.get("valign", "middle")),
            ("interligne", options.get("interligne", 1.15)),
            ("lien", options.get("lien")),
            ("classe", options.get("classe")),
            ("cle", options.get("cle")),
        ])
        self.formes.append(forme)
        return forme

    def bezier(self, points, **options):
        forme = OrderedDict([
            ("type", "bezier"), ("pts", [list(p) for p in points]),
            ("couleur", options.get("couleur", self.c["gris"])),
            ("lw", options.get("lw", 1.1)),
            ("dash", options.get("dash")),
            ("fleche", options.get("fleche", True)),
            ("classe", options.get("classe")),
            ("cle", options.get("cle")),
            ("info", options.get("info")),
        ])
        self.formes.append(forme)
        return forme

    def graphe(self, x, y, w, h, spec):
        self.formes.append(OrderedDict([
            ("type", "graphe"), ("x", x), ("y", y), ("w", w), ("h", h),
            ("spec", spec),
        ]))

    def table(self, x, y, w, h, spec):
        self.formes.append(OrderedDict([
            ("type", "table"), ("x", x), ("y", y), ("w", w), ("h", h),
            ("spec", spec),
        ]))

    def en_dict(self):
        return OrderedDict([
            ("id", self.id), ("titre", self.titre), ("genre", self.genre),
            ("indice", self.indice), ("charte", self.c), ("formes", self.formes),
        ])


def bandeau(scene, sur_titre, titre, retour=None):
    """Bandeau de tête commun à toutes les pages."""
    c = scene.c
    scene.rect(0, 0, PAGE_W, HAUT_BANDEAU, fill=c["primary"], lw=0)
    scene.texte(MARGE, 8, 620, 15, titre, taille=15, gras=True,
                couleur=c["blanc"], align="left", valign="middle")
    scene.texte(MARGE, 25, 620, 12, sur_titre, taille=8,
                couleur=c["blanc"], align="left", valign="middle")
    if retour is not None:
        w = 118.0
        scene.rect(PAGE_W - MARGE - w, 13, w, 20, fill=c["blanc"], lw=0,
                   radius=3, texte="◀  Vue macro", taille=8, gras=True,
                   couleur=c["primary"], lien=retour, classe="retour")


def pied(scene, mention):
    """Pied de page : mention de source et pagination."""
    c = scene.c
    y = PAGE_H - HAUT_PIED
    scene.rect(0, y, PAGE_W, HAUT_PIED, fill=c["fond"], lw=0)
    scene.rect(0, y, PAGE_W, 0.6, fill=c["trait"], lw=0)
    scene.texte(MARGE, y + 6, 700, 12, mention, taille=6.5, couleur=c["gris"],
                align="left")
    scene.texte(PAGE_W - MARGE - 200, y + 6, 200, 12,
                "%d / %d" % (scene.indice + 1, scene.total), taille=6.5,
                couleur=c["gris"], align="right")


# ═════════════════════════════════════════════════════════════════════════════
# Pastilles d'outil
# ═════════════════════════════════════════════════════════════════════════════

def style_pastille(scene, outil):
    """Remplissage et contour d'une pastille, selon la nature de l'outil.

    Quatre états, tous composés des seuls tokens de la charte :
      hub          fond primaire, texte blanc
      socle        fond gris clair, contour primaire
      application  fond blanc, contour fin
      hors BDD     contour accent en pointillé, quelle que soit la nature
    """
    c = scene.c
    if outil["hub"]:
        style = {"fill": c["primary"], "line": c["primary"], "couleur": c["blanc"]}
    elif outil["nature"] == "socle":
        style = {"fill": c["fond"], "line": c["primary"], "couleur": c["encre"]}
    else:
        style = {"fill": c["blanc"], "line": c["trait"], "couleur": c["encre"]}
    if outil["origine"] != "bdd":
        style["line"] = c["accent"]
        style["dash"] = "dash"
    return style


def info_outil(outil):
    """Contenu de l'infobulle d'une pastille."""
    parties = [outil["nom"]]
    if outil["fonction"]:
        parties.append(outil["fonction"])
    parties.append("%s · %s · alimentation %s" % (
        "socle de données" if outil["nature"] == "socle" else "application métier",
        {"interne": "interne", "externe": "externe"}.get(outil["ie"], "TODO interne/externe"),
        {"auto": "automatique", "manuel": "manuelle"}.get(outil["alim"], "TODO")))
    if outil["todo"]:
        parties.append("TODO : " + " ; ".join(outil["todo"]))
    return "\n".join(parties)


def _placer_pastilles(x, y, largeur_zone, noms, taille, hauteur, gap, interligne):
    """Calcule la position de chaque pastille. Ne dessine rien."""
    places = []
    curseur_x, curseur_y = x, y
    for nom in noms:
        w = min(largeur_zone, largeur(nom, taille) + 11.0)
        if curseur_x > x and curseur_x + w > x + largeur_zone + 0.01:
            curseur_x = x
            curseur_y += hauteur + interligne
        places.append((nom, curseur_x, curseur_y, w))
        curseur_x += w + gap
    return places, (curseur_y + hauteur - y if places else 0.0)


def hauteur_pastilles(largeur_zone, noms, taille, hauteur, gap=4.0, interligne=3.0):
    """Hauteur qu'occuperaient ces pastilles, sans rien dessiner."""
    return _placer_pastilles(0, 0, largeur_zone, noms, taille, hauteur,
                             gap, interligne)[1]


def taille_pastilles(cartes, disponible, tailles=(8.0, 7.5, 7.0, 6.5, 6.0, 5.6)):
    """Plus grande taille de police qui tienne dans toutes les cartes d'une rangée.

    `cartes` est une liste de couples (largeur de zone, noms d'outils).
    """
    for taille in tailles:
        hauteur = round(taille * 2.0, 1)
        if all(hauteur_pastilles(zone, noms, taille, hauteur) <= disponible
               for zone, noms in cartes):
            return taille, hauteur
    taille = tailles[-1]
    return taille, round(taille * 2.0, 1)


def poser_pastilles(scene, x, y, largeur_zone, noms, taille=6.5, hauteur=13.0,
                    gap=4.0, interligne=3.0, ancrer=True):
    """Dispose des pastilles d'outil en lignes, avec retour à la ligne.

    Chaque occurrence est enregistrée comme ancre possible : un outil déclaré
    par plusieurs services en a plusieurs, et le routage choisit ensuite la
    plus proche de l'autre extrémité du flux.
    """
    places, totale = _placer_pastilles(x, y, largeur_zone, noms, taille,
                                       hauteur, gap, interligne)
    for nom, px_, py_, w in places:
        outil = scene.carto["outils"][nom]
        style = style_pastille(scene, outil)
        forme = scene.rect(px_, py_, w, hauteur, radius=hauteur / 2.0,
                           lw=0.7, taille=taille, texte=tronquer(nom, taille, w - 8),
                           classe="pastille", cle=outil["id"],
                           info=info_outil(outil), **style)
        if ancrer:
            scene.ancres.setdefault(nom, []).append(forme)
    return totale


# ═════════════════════════════════════════════════════════════════════════════
# Routage des connecteurs
# ═════════════════════════════════════════════════════════════════════════════

ARC_VOISIN = 34.0   # en deçà, le flux est tracé en arc par-dessus les pastilles


def _centre(forme):
    return forme["x"] + forme["w"] / 2.0, forme["y"] + forme["h"] / 2.0


def _cote(depart, arrivee):
    """Bord par lequel sortir : vertical si l'écart vertical domine."""
    ax, ay = _centre(depart)
    bx, by = _centre(arrivee)
    if abs(by - ay) >= abs(bx - ax) * 0.75:
        return "bas" if by > ay else "haut"
    return "droite" if bx > ax else "gauche"


def _point_bord(forme, cote, rang, total):
    """Point de sortie sur un bord, réparti quand plusieurs flux le partagent."""
    fraction = (rang + 1.0) / (total + 1.0)
    if cote == "haut":
        return forme["x"] + forme["w"] * fraction, forme["y"]
    if cote == "bas":
        return forme["x"] + forme["w"] * fraction, forme["y"] + forme["h"]
    if cote == "gauche":
        return forme["x"], forme["y"] + forme["h"] * fraction
    return forme["x"] + forme["w"], forme["y"] + forme["h"] * fraction


def _controle(point, cote, amplitude):
    x, y = point
    if cote == "haut":
        return x, y - amplitude
    if cote == "bas":
        return x, y + amplitude
    if cote == "gauche":
        return x - amplitude, y
    return x + amplitude, y


def _est_hub(forme):
    return "hub" in (forme.get("classe") or "")


def _cotes(depart, arrivee):
    """Bords de sortie et d'entrée d'un flux.

    Trois cas, comme dans la cartographie d'origine :
      hub ↔ hub        arc sous le bandeau, pour ne pas traverser les pastilles
                       intermédiaires alignées avec lui ;
      hub ↔ rangée     sortie verticale, la courbe reste dans le couloir entre
                       les deux zones au lieu de couper une carte ;
      sinon            bord le plus proche de la cible.
    """
    _, ay = _centre(depart)
    _, by = _centre(arrivee)
    if _est_hub(depart) and _est_hub(arrivee):
        return "bas", "bas"
    if _est_hub(depart) != _est_hub(arrivee):
        return ("bas", "haut") if by > ay else ("haut", "bas")
    return _cote(depart, arrivee), _cote(arrivee, depart)


def _candidats(scene, nom):
    """Pastilles où un flux peut s'accrocher.

    Un outil du hub n'a qu'une ancre légitime, celle du bandeau central :
    c'est là que le lecteur attend la relation « logiciel ↔ hub ».
    """
    ancres = scene.ancres.get(nom) or []
    if ancres and scene.carto["outils"][nom]["hub"]:
        return ancres[:1]
    return ancres


def tracer_flux(scene, flux):
    """Trace les flux en une seule courbe de Bézier cubique par flux.

    Les extrémités partagées sont réparties le long du bord, triées selon la
    position de l'extrémité opposée : c'est ce qui évite les croisements.
    """
    c = scene.c
    aretes = []
    for f in flux:
        candidats_a = _candidats(scene, f["de"])
        candidats_b = _candidats(scene, f["vers"])
        if not candidats_a or not candidats_b:
            continue
        # un outil déclaré par plusieurs services a plusieurs pastilles : on
        # retient le couple le plus court, sinon les fils traversent la planche
        depart, arrivee, meilleure = None, None, None
        for a in candidats_a:
            ax, ay = _centre(a)
            for b in candidats_b:
                bx, by = _centre(b)
                distance = math.hypot(bx - ax, by - ay)
                if meilleure is None or distance < meilleure:
                    depart, arrivee, meilleure = a, b, distance
        aretes.append((f, depart, arrivee) + _cotes(depart, arrivee))

    # regroupement par (pastille, bord) pour répartir les points de sortie
    groupes = {}
    for indice, (f, depart, arrivee, cote_a, cote_b) in enumerate(aretes):
        groupes.setdefault((id(depart), cote_a), []).append((indice, "a"))
        groupes.setdefault((id(arrivee), cote_b), []).append((indice, "b"))

    positions = {}
    for (_, cote), membres in groupes.items():
        def tri(membre):
            indice, bout = membre
            _, depart, arrivee, _, _ = aretes[indice]
            oppose = arrivee if bout == "a" else depart
            x, y = _centre(oppose)
            return x if cote in ("haut", "bas") else y
        membres.sort(key=tri)
        for rang, membre in enumerate(membres):
            positions[membre] = (rang, len(membres))

    for indice, (f, depart, arrivee, cote_a, cote_b) in enumerate(aretes):
        rang_a, total_a = positions[(indice, "a")]
        rang_b, total_b = positions[(indice, "b")]
        point_a = _point_bord(depart, cote_a, rang_a, total_a)
        point_b = _point_bord(arrivee, cote_b, rang_b, total_b)
        distance = math.hypot(point_b[0] - point_a[0], point_b[1] - point_a[1])
        amplitude = max(16.0, min(76.0, distance * 0.42))
        if _est_hub(depart) and _est_hub(arrivee):
            amplitude = 30.0
        elif distance < ARC_VOISIN:
            # deux pastilles côte à côte : sans cet arc, le flux se réduit à un
            # point coincé entre elles et personne ne le voit
            cote_a = cote_b = "haut"
            point_a = _point_bord(depart, "haut", rang_a, total_a)
            point_b = _point_bord(arrivee, "haut", rang_b, total_b)
            amplitude = 21.0
        auto = f["mode"] == "auto"
        scene.bezier(
            [point_a, _controle(point_a, cote_a, amplitude),
             _controle(point_b, cote_b, amplitude), point_b],
            couleur=c["accent"] if auto else c["gris"],
            lw=1.3 if auto else 1.0,
            dash=None if auto else "dash",
            classe="flux",
            cle="%s>%s" % (scene.carto["outils"][f["de"]]["id"],
                           scene.carto["outils"][f["vers"]]["id"]),
            info="%s → %s\nMode : %s\n%s" % (
                f["de"], f["vers"],
                "automatique" if auto else "manuel",
                f["objet"] or "objet non renseigné"))


# ═════════════════════════════════════════════════════════════════════════════
# Page 1 — vue macro
# ═════════════════════════════════════════════════════════════════════════════

# Quatre pôles techniques en haut, hub au centre, quatre pôles support en bas.
# Sécurité, vide, ferme la rangée du bas.
RANGEE_HAUT = ["methode-bim", "travaux", "travaux-tunnel", "topo", "qualite-env"]
RANGEE_BAS = ["direction", "daf", "contrat", "securite"]


def carte_service(scene, service, x, y, w, h, couleur_rule, lien,
                  taille=6.5, hauteur_pastille=13.0):
    """Une carte de service : bandeau cliquable puis pastilles de ses outils."""
    c = scene.c
    vide = service["vide"]
    scene.rect(x, y, w, h, fill=c["blanc"], line=c["trait"], lw=0.8, radius=4,
               dash="dash" if vide else None,
               lien=lien, classe="carte", cle=service["id"],
               info="%s — %d outil(s), %d processus\nCliquer pour ouvrir l'affiche"
                    % (service["nom"], len(service["outils"]),
                       len(service["processus"])))
    scene.rect(x, y, w, 2.6, fill=couleur_rule, lw=0, radius=0)
    scene.texte(x + 8, y + 6, w - 16, 12, service["nom"], taille=9.5, gras=True,
                couleur=c["encre"], align="left", classe="titre-carte")
    scene.texte(x + 8, y + 17, w - 46, 9, service["kicker"], taille=6.2,
                couleur=c["gris"], align="left")
    compteur = "TODO" if vide else "%d outils" % len(service["outils"])
    scene.texte(x + w - 46, y + 17, 38, 9, compteur, taille=6.2,
                couleur=c["accent"] if vide else c["gris"], align="right")
    scene.rect(x + 8, y + 28, w - 16, 0.5, fill=c["trait"], lw=0)

    if vide:
        scene.rect(x + 8, y + 34, w - 16, max(24.0, h - 44), fill=c["fond"],
                   line=c["accent"], lw=0.7, dash="dash", radius=3,
                   texte="TODO — service à documenter", taille=6.8,
                   couleur=c["accent"])
        return
    poser_pastilles(scene, x + 8, y + 34, w - 16, service["outils"],
                    taille=taille, hauteur=hauteur_pastille)


def page_macro(carto, indice, total, liens):
    c = carto["charte"]
    scene = Scene(carto, "macro", "Vue macro", "macro", indice)
    scene.total = total
    bandeau(scene, carto["meta"]["titre"],
            "Cartographie applicative et flux de données")

    utile = PAGE_W - 2 * MARGE
    gap = 9.0
    w_haut = (utile - gap * (len(RANGEE_HAUT) - 1)) / len(RANGEE_HAUT)
    w_bas = (utile - gap * (len(RANGEE_BAS) - 1)) / len(RANGEE_BAS)

    # les cartes s'ajustent à leur contenu : on cherche la plus grande taille
    # de pastille qui tienne dans la rangée, puis la rangée prend cette hauteur
    haut_y, hub_h, marge_bas = 56.0, 66.0, 92.0
    espace = PAGE_H - HAUT_PIED - marge_bas - hub_h - haut_y - 3 * 14.0
    plafond = espace / 2.0
    taille_haut, ph_haut = taille_pastilles(
        [(w_haut - 16, carto["_services_par_id"][s]["outils"]) for s in RANGEE_HAUT],
        plafond - 42.0)
    taille_bas, ph_bas = taille_pastilles(
        [(w_bas - 16, carto["_services_par_id"][s]["outils"]) for s in RANGEE_BAS],
        plafond - 42.0)
    haut_h = 42.0 + max(hauteur_pastilles(w_haut - 16,
                                          carto["_services_par_id"][s]["outils"],
                                          taille_haut, ph_haut)
                        for s in RANGEE_HAUT)
    bas_h = 42.0 + max(hauteur_pastilles(w_bas - 16,
                                         carto["_services_par_id"][s]["outils"],
                                         taille_bas, ph_bas)
                       for s in RANGEE_BAS)
    hauteur_legende = 34.0
    reste = (PAGE_H - HAUT_PIED - 10.0 - haut_y - haut_h - hub_h - bas_h
             - hauteur_legende)
    ecart = max(14.0, min(46.0, reste / 3.0))
    hub_y = haut_y + haut_h + ecart
    bas_y = hub_y + hub_h + ecart
    legende_y = PAGE_H - HAUT_PIED - 10.0 - hauteur_legende

    # ── hub de données, posé en premier : ses outils servent d'ancre aux flux
    scene.rect(MARGE, hub_y, utile, hub_h, fill=c["fond"], line=c["primary"],
               lw=1.0, radius=4)
    scene.texte(MARGE + 12, hub_y + 8, 300, 11, "Hub de données", taille=9.5,
                gras=True, couleur=c["primary"], align="left")
    scene.texte(MARGE + 12, hub_y + 20, 300, 9,
                "Gestion électronique documentaire", taille=6.4,
                couleur=c["gris"], align="left")
    largeur_hub = 0.0
    pastilles = []
    for nom in carto["hub"]:
        pastilles.append((nom, largeur(nom, 9.0) + 26.0))
        largeur_hub += pastilles[-1][1] + 10.0
    curseur = MARGE + utile / 2.0 - (largeur_hub - 10.0) / 2.0
    curseur = max(curseur, MARGE + 200.0)
    for nom, w in pastilles:
        outil = carto["outils"][nom]
        absent = outil["origine"] != "bdd"
        forme = scene.rect(
            curseur, hub_y + 20, w, 24, radius=12, lw=1.0,
            fill=c["blanc"] if absent else c["primary"],
            line=c["accent"] if absent else c["primary"],
            dash="dash" if absent else None,
            couleur=c["accent"] if absent else c["blanc"],
            texte=nom, taille=9.0, gras=True,
            classe="pastille hub", cle=outil["id"], info=info_outil(outil))
        scene.ancres.setdefault(nom, []).insert(0, forme)
        if absent:
            scene.texte(curseur, hub_y + 45, w, 8, "TODO : absent de la BDD",
                        taille=5.4, couleur=c["accent"], align="center")
        curseur += w + 10.0

    # ── cartes de service
    for rang, sid in enumerate(RANGEE_HAUT):
        service = carto["_services_par_id"][sid]
        carte_service(scene, service, MARGE + rang * (w_haut + gap), haut_y,
                      w_haut, haut_h, c["primary"], liens[sid],
                      taille_haut, ph_haut)
    for rang, sid in enumerate(RANGEE_BAS):
        service = carto["_services_par_id"][sid]
        carte_service(scene, service, MARGE + rang * (w_bas + gap), bas_y,
                      w_bas, bas_h, c["accent"], liens[sid],
                      taille_bas, ph_bas)

    tracer_flux(scene, [f for f in carto["flux"] if f["carte"]])

    # ── légende
    y = legende_y
    scene.rect(MARGE, y, utile, 34, fill=c["fond"], lw=0, radius=3)
    scene.texte(MARGE + 10, y + 4, 120, 10, "Légende", taille=7.5, gras=True,
                couleur=c["encre"], align="left")
    entrees = [
        ("bezier-auto", "Flux automatique"),
        ("bezier-manuel", "Flux manuel"),
        ("pastille-hub", "Hub de données"),
        ("pastille-socle", "Socle de données"),
        ("pastille-appli", "Application métier"),
        ("pastille-hors", "Hors BDD (correction / brief)"),
    ]
    x = MARGE + 10.0
    for genre, libelle in entrees:
        if genre.startswith("bezier"):
            auto = genre.endswith("auto")
            scene.bezier([(x, y + 22), (x + 6, y + 22), (x + 14, y + 22),
                          (x + 20, y + 22)],
                         couleur=c["accent"] if auto else c["gris"],
                         lw=1.3 if auto else 1.0,
                         dash=None if auto else "dash", fleche=True)
        else:
            style = {"pastille-hub": {"fill": c["primary"], "line": c["primary"]},
                     "pastille-socle": {"fill": c["fond"], "line": c["primary"]},
                     "pastille-appli": {"fill": c["blanc"], "line": c["trait"]},
                     "pastille-hors": {"fill": c["blanc"], "line": c["accent"],
                                       "dash": "dash"}}[genre]
            scene.rect(x, y + 16, 20, 11, radius=5.5, lw=0.7, **style)
        w = largeur(libelle, 6.4) + 6
        scene.texte(x + 24, y + 16, w, 11, libelle, taille=6.4,
                    couleur=c["gris"], align="left")
        x += 24 + w + 14

    pied(scene, "Source : %s — %d lignes d'entretien. "
                "Les flux non traçables sont détaillés dans data/gap-report.md."
         % (carto["meta"]["source"], carto["meta"]["lignes_bdd"]))
    return scene.en_dict()


# ═════════════════════════════════════════════════════════════════════════════
# Pages 2 à 10 — une affiche par service
# ═════════════════════════════════════════════════════════════════════════════

LARGEUR_NOM = 112.0        # colonne du libellé de processus
LARGEUR_BORNE = 82.0       # blocs début et fin
HAUTEUR_ETAPE = 27.0
HAUTEUR_ETAPE_MAX = 40.0
GAP_ETAPE = 6.0
GAP_BANDE = 6.0
HAUTEUR_BANDE_MAX = 112.0
TAILLE_ETAPE = 7.2
TAILLE_BADGE = 5.4

LIBELLE_ALIM = {"auto": "auto", "manuel": "manuel", None: "alim. TODO"}
LIBELLE_IE = {"interne": "interne", "externe": "externe", None: "int/ext TODO"}


def _lignes_etapes(scene, etapes, largeur_zone):
    """Découpe les étapes en lignes, en renvoyant chaque étape mesurée."""
    mesurees = []
    for etape in etapes:
        outil = scene.carto["outils"].get(etape["outil"])
        badge = "%s · %s" % (LIBELLE_ALIM.get(etape["alim"], "alim. TODO"),
                             LIBELLE_IE.get(etape["ie"], "int/ext TODO"))
        w = max(largeur(etape["outil"], TAILLE_ETAPE),
                largeur(badge, TAILLE_BADGE)) + 14.0
        mesurees.append((etape, outil, badge, min(w, largeur_zone)))
    lignes, courante, reste = [], [], largeur_zone
    for mesure in mesurees:
        besoin = mesure[3] + (GAP_ETAPE if courante else 0.0)
        if courante and besoin > reste + 0.01:
            lignes.append(courante)
            courante, reste = [], largeur_zone
            besoin = mesure[3]
        courante.append(mesure)
        reste -= besoin
    if courante:
        lignes.append(courante)
    return lignes or [[]]


HAUTEUR_BORNE_MAX = 56.0


def _borne(scene, x, y, w, h, contenu, role):
    """Bloc début ou fin. Vide, dimensionné et éditable quand la BDD ne dit rien."""
    c = scene.c
    hauteur = min(h, HAUTEUR_BORNE_MAX)
    y = y + (h - hauteur) / 2.0
    if contenu:
        scene.rect(x, y, w, hauteur, fill=c["primary"], line=c["primary"], lw=0.8,
                   radius=3, texte=contenu, taille=6.4, gras=True, pad=4.0,
                   nowrap=False, couleur=c["blanc"], classe="borne")
    else:
        scene.rect(x, y, w, hauteur, fill=c["blanc"], line=c["accent"], lw=0.8,
                   dash="dash", radius=3, classe="borne borne-vide")
        scene.texte(x, y + hauteur / 2.0 - 9, w, 8, role, taille=5.4,
                    couleur=c["gris"], align="center")
        scene.texte(x, y + hauteur / 2.0 - 1, w, 9, "TODO", taille=6.4, gras=True,
                    couleur=c["accent"], align="center")


def _fleche(scene, x, y, hauteur):
    scene.texte(x, y + hauteur / 2.0 - 5, 12, 10, "▸", taille=9,
                couleur=scene.c["gris"], align="center")


def page_service(carto, service, indice, total, lien_macro):
    c = carto["charte"]
    scene = Scene(carto, "service-" + service["id"], service["nom"], "service", indice)
    scene.total = total
    bandeau(scene, "Affiche service — processus, alimentation et nature des outils",
            service["nom"], retour=lien_macro)

    utile = PAGE_W - 2 * MARGE
    y = HAUT_BANDEAU + 10.0

    nb_flux = len(model.flux_du_service(carto, service["id"]))
    resume = "%d processus · %d outils · %d flux touchant le service" % (
        len(service["processus"]), len(service["outils"]), nb_flux)
    scene.texte(MARGE, y, utile - 220, 12, resume, taille=7.4, couleur=c["gris"],
                align="left")
    scene.texte(MARGE + utile - 220, y, 220, 12, service["kicker"], taille=7.4,
                couleur=c["primary"], gras=True, align="right")
    y += 16.0
    scene.rect(MARGE, y, utile, 0.6, fill=c["trait"], lw=0)
    y += 8.0

    zone_x = MARGE + LARGEUR_NOM + 8.0 + LARGEUR_BORNE + 12.0
    zone_fin_x = PAGE_W - MARGE - LARGEUR_BORNE
    zone_w = zone_fin_x - 12.0 - zone_x

    bandes = []
    for processus in service["processus"]:
        lignes = _lignes_etapes(scene, processus["etapes"], zone_w)
        naturelle = 14.0 + len(lignes) * HAUTEUR_ETAPE + (len(lignes) - 1) * 5.0
        bandes.append([processus, lignes, max(38.0, naturelle)])

    disponible = PAGE_H - HAUT_PIED - 10.0 - y
    total_naturel = sum(b[2] for b in bandes) + GAP_BANDE * (len(bandes) - 1)
    gap = GAP_BANDE
    if total_naturel < disponible and bandes:
        rab = (disponible - total_naturel) / len(bandes)
        for bande in bandes:
            supplement = min(rab, HAUTEUR_BANDE_MAX - bande[2])
            bande[2] += max(0.0, supplement)
        reste = disponible - sum(b[2] for b in bandes)
        if len(bandes) > 1:
            gap = min(42.0, max(GAP_BANDE, reste / (len(bandes) - 1)))
    elif total_naturel > disponible and len(bandes) > 1:
        facteur = (disponible - gap * (len(bandes) - 1)) / sum(b[2] for b in bandes)
        for bande in bandes:
            bande[2] *= facteur

    for processus, lignes, hauteur in bandes:
        vierge = processus["origine"] == "vierge"
        scene.rect(MARGE, y, utile, hauteur, fill=c["blanc"], line=c["trait"],
                   lw=0.7, radius=3, dash="dash" if vierge else None,
                   classe="bande")
        scene.rect(MARGE, y, 2.4, hauteur,
                   fill=c["accent"] if (vierge or processus["todo"]) else c["primary"],
                   lw=0)
        nom = processus["nom"] or "Processus à documenter"
        scene.texte(MARGE + 9, y + 7, LARGEUR_NOM - 4, 22, nom, taille=7.6,
                    gras=True, couleur=c["encre"], align="left", valign="top")
        sous = []
        if processus["origine"] in ("correction", "vierge"):
            sous.append("correction métier" if processus["origine"] == "correction"
                        else "bloc vierge")
        if not processus["ordre_atteste"] and processus["etapes"]:
            sous.append("ordre non attesté")
        if sous:
            scene.texte(MARGE + 9, y + 30, LARGEUR_NOM - 4, 16,
                        " · ".join(sous), taille=5.2, couleur=c["gris"],
                        align="left", valign="top")

        corps_y = y + 7.0
        corps_h = hauteur - 14.0 - (9.0 if processus["todo"] else 0.0)
        _borne(scene, MARGE + LARGEUR_NOM + 8, corps_y, LARGEUR_BORNE, corps_h,
               processus["debut"], "DÉBUT")
        _fleche(scene, MARGE + LARGEUR_NOM + 8 + LARGEUR_BORNE, corps_y, corps_h)
        _borne(scene, zone_fin_x, corps_y, LARGEUR_BORNE, corps_h,
               processus["fin"], "FIN")
        _fleche(scene, zone_fin_x - 12, corps_y, corps_h)

        if not processus["etapes"]:
            scene.rect(zone_x, corps_y, zone_w, corps_h, fill=c["fond"],
                       line=c["trait"], lw=0.6, dash="dash", radius=3,
                       texte="TODO — étapes à compléter", taille=6.4,
                       couleur=c["gris"], classe="zone-vide")
        else:
            hauteur_ligne = min(HAUTEUR_ETAPE_MAX,
                                (corps_h - 5.0 * (len(lignes) - 1)) / len(lignes))
            ligne_y = corps_y + (corps_h - (hauteur_ligne * len(lignes)
                                            + 5.0 * (len(lignes) - 1))) / 2.0
            for ligne in lignes:
                x = zone_x
                for etape, outil, badge, w in ligne:
                    style = (style_pastille(scene, outil) if outil else
                             {"fill": c["blanc"], "line": c["accent"], "dash": "dash",
                              "couleur": c["accent"]})
                    scene.rect(x, ligne_y, w, hauteur_ligne, radius=3, lw=0.8,
                               classe="etape", cle=outil["id"] if outil else None,
                               info=info_outil(outil) if outil else None, **style)
                    milieu = ligne_y + hauteur_ligne / 2.0
                    scene.texte(x, milieu - 10, w, 10, etape["outil"],
                                taille=TAILLE_ETAPE, gras=True,
                                couleur=style["couleur"], align="center")
                    scene.texte(x, milieu + 1, w, 8, badge,
                                taille=TAILLE_BADGE, couleur=style["couleur"],
                                align="center")
                    x += w + GAP_ETAPE
                ligne_y += hauteur_ligne + 5.0

        if processus["todo"]:
            scene.texte(MARGE + LARGEUR_NOM + 8, y + hauteur - 9,
                        utile - LARGEUR_NOM - 16, 8,
                        "TODO : " + processus["todo"], taille=5.2,
                        couleur=c["accent"], align="left")
        y += hauteur + gap

    pied(scene, "Alimentation et nature reprises de la BDD. « TODO » = donnée "
                "absente de la base, à compléter. Détail dans data/gap-report.md.")
    return scene.en_dict()


# ═════════════════════════════════════════════════════════════════════════════
# Page 11 — référentiel des outils
# ═════════════════════════════════════════════════════════════════════════════

CODE_SERVICE = OrderedDict([
    ("direction", "DIR"), ("daf", "DAF"), ("contrat", "CTR"),
    ("methode-bim", "M&B"), ("travaux", "TRV"), ("travaux-tunnel", "TUN"),
    ("topo", "TOP"), ("qualite-env", "Q&E"), ("securite", "SEC"),
])
COLONNES = [("Outil", 86.0, "left"), ("Fonction", 158.0, "left"),
            ("Services", 92.0, "left"), ("Nature", 24.0, "center"),
            ("Int/Ext", 24.0, "center"), ("Alim.", 28.0, "center"),
            ("Origine", 32.0, "center")]
ABREGE_NATURE = {"socle": "Socle", "application": "Appli"}
ABREGE_IE = {"interne": "Int", "externe": "Ext", None: "TODO"}
ABREGE_ALIM = {"auto": "Auto", "manuel": "Manuel", None: "TODO"}
ABREGE_ORIGINE = {"bdd": "BDD", "bdd-annexe": "Annexe", "correction": "Corr.",
                  "brief": "Brief"}


def page_referentiel(carto, indice, total, lien_macro):
    c = carto["charte"]
    scene = Scene(carto, "referentiel", "Référentiel des outils", "referentiel", indice)
    scene.total = total
    bandeau(scene, "Tous les outils recensés, leur fonction et leur rattachement",
            "Référentiel des outils", retour=lien_macro)

    outils = model.outils_tries(carto)
    utile = PAGE_W - 2 * MARGE
    scene.texte(MARGE, 54, utile, 11,
                "%d outils distincts · %d socle de données · %d issus d'une "
                "correction ou du brief" % (
                    len(outils),
                    sum(1 for o in outils if o["nature"] == "socle"),
                    sum(1 for o in outils if o["origine"] != "bdd")),
                taille=7.4, couleur=c["gris"], align="left")

    haut_table = 72.0
    bas_table = 466.0
    colonne_w = (utile - 16.0) / 2.0
    par_colonne = (len(outils) + 1) // 2
    pas = (bas_table - haut_table) / (par_colonne + 1)
    pas = min(pas, 14.0)

    for bloc in range(2):
        x0 = MARGE + bloc * (colonne_w + 16.0)
        lot = outils[bloc * par_colonne:(bloc + 1) * par_colonne]
        y = haut_table
        x = x0 + 4.0
        for libelle, w, align in COLONNES:
            scene.texte(x, y, w, 10, libelle, taille=5.8, gras=True,
                        couleur=c["gris"], align=align)
            x += w
        y += 11.0
        scene.rect(x0, y, colonne_w, 0.7, fill=c["encre"], lw=0)
        y += 1.5
        for rang, outil in enumerate(lot):
            if rang % 2 == 1:
                scene.rect(x0, y, colonne_w, pas, fill=c["fond"], lw=0)
            codes = " ".join(CODE_SERVICE[s] for s in CODE_SERVICE
                             if s in outil["services"])
            valeurs = [
                (outil["nom"], c["encre"], True),
                (outil["fonction"] or "TODO : fonction à renseigner",
                 c["gris"] if outil["fonction"] else c["accent"], False),
                (codes or "TODO", c["gris"] if codes else c["accent"], False),
                (ABREGE_NATURE.get(outil["nature"], "?"),
                 c["primary"] if outil["nature"] == "socle" else c["gris"], False),
                (ABREGE_IE.get(outil["ie"]),
                 c["accent"] if outil["ie"] is None else c["gris"], False),
                (ABREGE_ALIM.get(outil["alim"]),
                 c["accent"] if outil["alim"] is None else c["gris"], False),
                (ABREGE_ORIGINE.get(outil["origine"], outil["origine"]),
                 c["gris"] if outil["origine"] == "bdd" else c["accent"], False),
            ]
            x = x0 + 4.0
            for (libelle, w, align), (valeur, couleur, gras) in zip(COLONNES, valeurs):
                taille = 6.2 if gras else 5.8
                scene.texte(x, y + (pas - 8) / 2.0, w - 4, 8,
                            tronquer(valeur, taille, w - 5), taille=taille,
                            gras=gras, couleur=couleur, align=align,
                            classe="cellule")
                x += w
            y += pas

    # ── légende des couleurs
    y = 474.0
    scene.rect(MARGE, y, utile, 34, fill=c["fond"], lw=0, radius=3)
    scene.texte(MARGE + 10, y + 4, 160, 10, "Légende des couleurs", taille=7.5,
                gras=True, couleur=c["encre"], align="left")
    entrees = [
        (c["primary"], "Socle de données et hub"),
        (c["accent"], "TODO ou élément hors BDD"),
        (c["gris"], "Donnée renseignée par la BDD"),
        (c["encre"], "Nom d'outil"),
    ]
    x = MARGE + 10.0
    for couleur, libelle in entrees:
        scene.rect(x, y + 17, 10, 10, fill=couleur, lw=0, radius=2)
        w = largeur(libelle, 6.4) + 6
        scene.texte(x + 14, y + 17, w, 10, libelle, taille=6.4, couleur=c["gris"],
                    align="left")
        x += 14 + w + 16
    codes = " · ".join("%s = %s" % (code, carto["_services_par_id"][sid]["nom"])
                       for sid, code in CODE_SERVICE.items())
    scene.texte(MARGE + 10, y + 27, utile - 20, 8, codes, taille=5.4,
                couleur=c["gris"], align="left")

    pied(scene, "Origine — BDD : table d'entretiens · Annexe : onglet Tableau de "
                "Bord · Corr. : correction métier · Brief : cahier des charges.")
    return scene.en_dict()


# ═════════════════════════════════════════════════════════════════════════════
# Page 12 — synthèse
# ═════════════════════════════════════════════════════════════════════════════

# Chaque tuile et chaque graphique nomme la clé d'indicateur qu'il consomme :
# le HTML les relit au changement d'onglet, le PPTX prend la vue globale.
TUILES = [
    ("nb_flux", "Flux recensés", "flux tracés ou documentés"),
    ("nb_outils", "Outils distincts", "après normalisation des libellés"),
    ("nb_socle", "Outils socle commun", "hub compris"),
    ("pct_auto", "Flux automatiques", "part des flux du périmètre"),
    ("pct_interne", "Outils internes", "développés par le groupe"),
    ("nb_todo_ie", "Interne/externe TODO", "non tranché par la BDD"),
]


def page_synthese(carto, indice, total, lien_macro, liens):
    c = carto["charte"]
    scene = Scene(carto, "synthese", "Synthèse", "synthese", indice)
    scene.total = total
    bandeau(scene, "Toutes les valeurs sont calculées depuis carto.json",
            "Synthèse chiffrée", retour=lien_macro)

    utile = PAGE_W - 2 * MARGE

    # ── sélecteur de service (interactif en HTML, statique en PPTX)
    y = 54.0
    onglets = [(model.TOUS, model.LIBELLE_TOUS)] + [
        (s["id"], s["nom"]) for s in carto["services"]]
    x = MARGE
    for sid, libelle in onglets:
        w = largeur(libelle, 6.6) + 16.0
        actif = sid == model.TOUS
        scene.rect(x, y, w, 18, radius=9, lw=0.8,
                   fill=c["primary"] if actif else c["blanc"],
                   line=c["primary"] if actif else c["trait"],
                   couleur=c["blanc"] if actif else c["gris"],
                   texte=libelle, taille=6.6, gras=actif,
                   classe="onglet", cle=sid)
        x += w + 5.0

    # ── six tuiles d'indicateurs
    y = 80.0
    largeur_tuile = (utile - 5 * 7.0) / 6.0
    for rang, (cle, titre, aide) in enumerate(TUILES):
        tx = MARGE + rang * (largeur_tuile + 7.0)
        scene.rect(tx, y, largeur_tuile, 62, fill=c["blanc"], line=c["trait"],
                   lw=0.8, radius=4, classe="tuile")
        scene.rect(tx, y, largeur_tuile, 2.4, fill=c["primary"], lw=0)
        scene.texte(tx + 8, y + 12, largeur_tuile - 16, 24, "—", taille=21,
                    gras=True, couleur=c["primary"], align="left",
                    classe="tuile-valeur", cle=cle)
        scene.texte(tx + 8, y + 38, largeur_tuile - 16, 10, titre, taille=6.8,
                    gras=True, couleur=c["encre"], align="left")
        scene.texte(tx + 8, y + 48, largeur_tuile - 16, 9, aide, taille=5.4,
                    couleur=c["gris"], align="left")

    # ── quatre graphiques
    y = 154.0
    largeur_graphe = (utile - 16.0) / 2.0
    hauteur_graphe = 168.0
    graphiques = [
        {"id": "flux-mode", "genre": "doughnut",
         "titre": "Répartition des flux : automatique / manuel",
         "donnees": "flux_mode",
         "libelles": ["Automatique", "Manuel"],
         "couleurs": [c["primary"], c["graph_a"]],
         "note": "Un flux est dit automatisé quand son outil de destination "
                 "est alimenté automatiquement."},
        {"id": "outils-ie", "genre": "doughnut",
         "titre": "Répartition des outils : interne / externe",
         "donnees": "outils_ie",
         "libelles": ["Interne", "Externe", "À documenter"],
         "couleurs": [c["primary"], c["graph_a"], c["accent"]],
         "note": "« À documenter » = la BDD ne tranche pas."},
        {"id": "classement-flux", "genre": "barres",
         "titre": "Classement des services par nombre de flux",
         "donnees": "classement", "champ": "nb_flux",
         "couleurs": [c["primary"]], "cliquable": True,
         "note": "Vue toujours globale ; le service sélectionné est mis en avant. "
                 "Cliquer une barre ouvre l'affiche du service."},
        {"id": "classement-outils", "genre": "barres",
         "titre": "Nombre d'outils par service",
         "donnees": "classement", "champ": "nb_outils",
         "couleurs": [c["graph_b"]], "cliquable": True,
         "note": "Un outil déclaré par plusieurs services y est compté "
                 "plusieurs fois."},
    ]
    for rang, spec in enumerate(graphiques):
        gx = MARGE + (rang % 2) * (largeur_graphe + 16.0)
        gy = y + (rang // 2) * (hauteur_graphe + 12.0)
        scene.rect(gx, gy, largeur_graphe, hauteur_graphe, fill=c["blanc"],
                   line=c["trait"], lw=0.8, radius=4)
        scene.texte(gx + 10, gy + 7, largeur_graphe - 20, 11, spec["titre"],
                    taille=7.6, gras=True, couleur=c["encre"], align="left")
        scene.texte(gx + 10, gy + hauteur_graphe - 16, largeur_graphe - 20, 12,
                    spec["note"], taille=5.2, couleur=c["gris"], align="left")
        spec["liens"] = liens
        scene.graphe(gx + 10, gy + 22, largeur_graphe - 20,
                     hauteur_graphe - 42, spec)

    pied(scene, "Sélecteur de service actif dans la page HTML ; le PPTX présente "
                "la vue « Tous services », les données des graphiques restant "
                "éditables dans leur classeur intégré.")
    return scene.en_dict()


# ═════════════════════════════════════════════════════════════════════════════
# Assemblage des 12 pages
# ═════════════════════════════════════════════════════════════════════════════

def construire(carto):
    """Renvoie la liste ordonnée des pages, prête à être rendue."""
    services = carto["services"]
    total = 1 + len(services) + 2
    liens = {s["id"]: 1 + rang for rang, s in enumerate(services)}

    pages = [page_macro(carto, 0, total, liens)]
    for rang, service in enumerate(services):
        pages.append(page_service(carto, service, 1 + rang, total, 0))
    pages.append(page_referentiel(carto, 1 + len(services), total, 0))
    pages.append(page_synthese(carto, 2 + len(services), total, 0, liens))
    return pages


if __name__ == "__main__":
    carto = model.charger()
    pages = construire(carto)
    print("%d pages" % len(pages))
    for page in pages:
        genres = {}
        for forme in page["formes"]:
            genres[forme["type"]] = genres.get(forme["type"], 0) + 1
        print("  %-22s %-12s %s" % (page["id"], page["genre"], genres))
