#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rendu HTML du graphe de scène.

    python3 build/generate_html.py [--out out/carto.html]

Produit une page unique, autonome et hors ligne : CSS et JS sont inlinés, la
seule ressource externe est vendor/chart.min.js, en chemin relatif. Aucun
appel réseau à l'exécution.

Ce script ne connaît aucun contenu métier : il ne fait que rendre les
primitives produites par layout.py.
"""

import argparse
import json
from collections import OrderedDict
from pathlib import Path

import layout
import model

RACINE = Path(__file__).resolve().parent.parent
SORTIE = RACINE / "out" / "carto.html"
PX = 4.0 / 3.0          # 1 point = 1,3333 pixel CSS


def px(valeur):
    return round(valeur * PX, 2)


def echapper(texte):
    return (str(texte).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def attribut(nom, valeur):
    return "" if valeur in (None, "") else ' %s="%s"' % (nom, echapper(valeur))


# ═════════════════════════════════════════════════════════════════════════════
# Rendu des primitives
# ═════════════════════════════════════════════════════════════════════════════

def style_texte(forme):
    """Alignement, police et couleur d'un bloc de texte."""
    justification = {"left": "flex-start", "center": "center",
                     "right": "flex-end"}[forme["align"]]
    alignement = {"top": "flex-start", "middle": "center",
                  "bottom": "flex-end"}[forme.get("valign", "middle")]
    regles = [
        "justify-content:%s" % justification,
        "align-items:%s" % alignement,
        "text-align:%s" % forme["align"],
        "font-size:%gpt" % forme["taille"],
        "color:%s" % forme["couleur"],
        "line-height:%g" % forme.get("interligne", 1.15),
    ]
    if forme.get("gras"):
        regles.append("font-weight:700")
    if forme.get("italique"):
        regles.append("font-style:italic")
    # un bloc assez haut pour deux lignes a le droit de revenir à la ligne ;
    # les libellés de pastille, jamais — ils déborderaient de leur forme
    if forme.get("nowrap") or forme["h"] < forme["taille"] * 1.9:
        regles.append("white-space:nowrap")
    return ";".join(regles)


def rendre_rect(forme):
    regles = [
        "left:%gpx" % px(forme["x"]), "top:%gpx" % px(forme["y"]),
        "width:%gpx" % px(forme["w"]), "height:%gpx" % px(forme["h"]),
    ]
    if forme["fill"]:
        regles.append("background:%s" % forme["fill"])
    if forme["line"] and forme["lw"]:
        regles.append("border:%gpx %s %s" % (
            max(1.0, px(forme["lw"])),
            "dashed" if forme["dash"] else "solid", forme["line"]))
    if forme["radius"]:
        regles.append("border-radius:%gpx" % px(forme["radius"]))
    if forme["texte"]:
        regles.append(style_texte(forme))
        regles.append("padding:0 %gpx" % px(forme["pad"]))
        regles.append("overflow:hidden")
    classes = ["s", "rect"]
    if forme["classe"]:
        classes.extend(forme["classe"].split())
    if forme["lien"] is not None:
        classes.append("lien")
    contenu = echapper(forme["texte"]) if forme["texte"] else ""
    balise = "div"
    ouvrant = '<%s class="%s" style="%s"%s%s%s>' % (
        balise, " ".join(classes), ";".join(regles),
        attribut("data-cle", forme["cle"]),
        attribut("data-info", forme["info"]),
        attribut("data-lien", forme["lien"]))
    return ouvrant + contenu + "</%s>" % balise


def rendre_texte(forme):
    regles = [
        "left:%gpx" % px(forme["x"]), "top:%gpx" % px(forme["y"]),
        "width:%gpx" % px(forme["w"]), "height:%gpx" % px(forme["h"]),
        style_texte(forme),
    ]
    classes = ["s", "txt"]
    if forme["classe"]:
        classes.extend(forme["classe"].split())
    if forme["lien"] is not None:
        classes.append("lien")
    return '<div class="%s" style="%s"%s%s>%s</div>' % (
        " ".join(classes), ";".join(regles),
        attribut("data-cle", forme.get("cle")),
        attribut("data-lien", forme["lien"]),
        echapper(forme["texte"]))


def rendre_bezier(forme, index):
    (x0, y0), (x1, y1), (x2, y2), (x3, y3) = forme["pts"]
    chemin = "M %g %g C %g %g, %g %g, %g %g" % (
        px(x0), px(y0), px(x1), px(y1), px(x2), px(y2), px(x3), px(y3))
    classes = ["fil"]
    if forme["classe"]:
        classes.extend(forme["classe"].split())
    marqueur = "fleche-auto" if forme["dash"] is None else "fleche-manuel"
    return ('<path class="%s" d="%s" stroke="%s" stroke-width="%g" fill="none"%s%s%s%s/>'
            % (" ".join(classes), chemin, forme["couleur"], px(forme["lw"]),
               ' stroke-dasharray="4 3"' if forme["dash"] else "",
               ' marker-end="url(#%s)"' % marqueur if forme["fleche"] else "",
               attribut("data-cle", forme["cle"]),
               attribut("data-info", forme["info"])))


def rendre_graphe(forme, index):
    spec = forme["spec"]
    return ('<div class="s graphe" style="left:%gpx;top:%gpx;width:%gpx;height:%gpx">'
            '<canvas id="g-%s" data-spec="%s"></canvas></div>'
            % (px(forme["x"]), px(forme["y"]), px(forme["w"]), px(forme["h"]),
               spec["id"], echapper(json.dumps(spec, ensure_ascii=False))))


MARQUEUR = ('<marker id="%s" viewBox="0 0 8 8" refX="6.4" refY="4" '
            'markerWidth="5" markerHeight="5" orient="auto-start-reverse">'
            '<path d="M0 0.7 L7.4 4 L0 7.3 Z" fill="%s"/></marker>')


def marqueurs(charte):
    """Têtes de flèche : le sens du flux doit se lire sans infobulle."""
    return "<defs>%s%s</defs>" % (
        MARQUEUR % ("fleche-auto", charte["accent"]),
        MARQUEUR % ("fleche-manuel", charte["gris"]))


def rendre_page(page, pages):
    corps = ['<section class="page" id="%s" data-genre="%s">'
             % (page["id"], page["genre"])]
    fils = []
    for index, forme in enumerate(page["formes"]):
        if forme["type"] == "rect":
            corps.append(rendre_rect(forme))
        elif forme["type"] == "texte":
            corps.append(rendre_texte(forme))
        elif forme["type"] == "graphe":
            corps.append(rendre_graphe(forme, index))
        elif forme["type"] == "bezier":
            fils.append(rendre_bezier(forme, index))
    if fils:
        corps.append('<svg class="fils" viewBox="0 0 %g %g">%s%s</svg>'
                     % (px(layout.PAGE_W), px(layout.PAGE_H),
                        marqueurs(page["charte"]), "".join(fils)))
    corps.append("</section>")
    # les liens sont résolus en ancres réelles
    resultat = "\n".join(corps)
    for indice, cible in enumerate(pages):
        resultat = resultat.replace('data-lien="%d"' % indice,
                                    'data-lien="%s"' % cible["id"])
    return resultat


# ═════════════════════════════════════════════════════════════════════════════
# Feuille de style
# ═════════════════════════════════════════════════════════════════════════════

CSS = """
*,*::before,*::after{box-sizing:border-box}
html,body{margin:0;padding:0}
body{background:%(fond)s;color:%(encre)s;font-family:%(police)s;
     -webkit-font-smoothing:antialiased}
.page{position:relative;width:%(pw)gpx;height:%(ph)gpx;margin:0 auto 22px;
      background:#fff;overflow:hidden;box-shadow:0 0 0 1px %(trait)s;
      scroll-margin-top:46px}
.s{position:absolute;overflow:hidden}
/* les formes sont posées à plat : sans cela un libellé posé sur une carte
   intercepterait le survol destiné à la carte elle-même */
.txt{pointer-events:none}
.rect,.txt{display:flex}
.rect{align-items:center;justify-content:center}
.lien{cursor:pointer}
.lien:hover{outline:1.5px solid %(primary)s;outline-offset:1px}
.fils{position:absolute;left:0;top:0;width:100%%;height:100%%;
      pointer-events:none;z-index:60;overflow:visible}
.fil{pointer-events:stroke;cursor:help;transition:opacity .12s}
.pastille,.carte,.etape,.tuile,.bande{transition:opacity .12s}
.attenue{opacity:.13}
.vif{opacity:1}
.souligne{outline:2px solid %(accent)s;outline-offset:1px}
#bulle{position:fixed;z-index:999;max-width:320px;padding:7px 10px;
       background:%(encre)s;color:#fff;font-size:11px;line-height:1.45;
       border-radius:4px;pointer-events:none;opacity:0;transition:opacity .1s;
       white-space:pre-line}
#bulle.on{opacity:1}
.onglet{cursor:pointer}
.graphe canvas{width:100%%!important;height:100%%!important}
#barre{position:sticky;top:0;z-index:200;display:flex;gap:6px;flex-wrap:wrap;
       align-items:center;padding:8px 14px;background:%(encre)s;color:#fff;
       font-size:11px}
#barre a{color:#fff;text-decoration:none;padding:3px 8px;border-radius:3px;
         border:1px solid rgba(255,255,255,.28)}
#barre a:hover{background:%(primary)s;border-color:%(primary)s}
#barre .titre{font-weight:700;margin-right:10px}
@media print{
  @page{size:A3 landscape;margin:0}
  body{background:#fff}
  #barre,#bulle{display:none}
  .page{margin:0;box-shadow:none;break-after:page;page-break-after:always}
  .page:last-child{break-after:auto;page-break-after:auto}
}
"""


def feuille_de_style(charte):
    valeurs = dict(charte)
    valeurs["pw"] = px(layout.PAGE_W)
    valeurs["ph"] = px(layout.PAGE_H)
    return CSS % valeurs


# ═════════════════════════════════════════════════════════════════════════════
# Comportements : infobulles, surbrillance, navigation, filtre de la synthèse
# ═════════════════════════════════════════════════════════════════════════════

JS = r"""
(function () {
  "use strict";
  var D = window.DONNEES;
  var bulle = document.getElementById("bulle");

  // ── infobulles ────────────────────────────────────────────────────────────
  function poser(evenement) {
    var texte = evenement.currentTarget.getAttribute("data-info");
    if (!texte) { return; }
    bulle.textContent = texte;
    bulle.classList.add("on");
    deplacer(evenement);
  }
  function deplacer(evenement) {
    var x = evenement.clientX + 14, y = evenement.clientY + 16;
    var boite = bulle.getBoundingClientRect();
    if (x + boite.width > window.innerWidth - 8) { x = evenement.clientX - boite.width - 14; }
    if (y + boite.height > window.innerHeight - 8) { y = evenement.clientY - boite.height - 14; }
    bulle.style.left = Math.max(4, x) + "px";
    bulle.style.top = Math.max(4, y) + "px";
  }
  function retirer() { bulle.classList.remove("on"); }

  Array.prototype.forEach.call(document.querySelectorAll("[data-info]"), function (noeud) {
    noeud.addEventListener("mouseenter", poser);
    noeud.addEventListener("mousemove", deplacer);
    noeud.addEventListener("mouseleave", retirer);
  });

  // ── navigation par clic ───────────────────────────────────────────────────
  Array.prototype.forEach.call(document.querySelectorAll("[data-lien]"), function (noeud) {
    noeud.addEventListener("click", function () {
      var cible = document.getElementById(noeud.getAttribute("data-lien"));
      if (cible) { cible.scrollIntoView({ behavior: "smooth", block: "start" }); }
    });
  });

  // ── surbrillance des flux sur la vue macro ────────────────────────────────
  var macro = document.getElementById("macro");
  if (macro) {
    var fils = Array.prototype.slice.call(macro.querySelectorAll(".fil.flux"));
    var pastilles = Array.prototype.slice.call(macro.querySelectorAll(".pastille"));
    var cartes = Array.prototype.slice.call(macro.querySelectorAll(".carte"));
    var extremites = fils.map(function (fil) {
      return (fil.getAttribute("data-cle") || "").split(">");
    });

    function eteindre() {
      fils.concat(pastilles, cartes).forEach(function (n) {
        n.classList.remove("attenue", "souligne");
      });
    }
    function allumer(outilsVises) {
      var vises = {};
      outilsVises.forEach(function (identifiant) { vises[identifiant] = true; });
      var touches = {};
      fils.forEach(function (fil, rang) {
        var bouts = extremites[rang];
        var actif = vises[bouts[0]] || vises[bouts[1]];
        fil.classList.toggle("attenue", !actif);
        if (actif) { touches[bouts[0]] = true; touches[bouts[1]] = true; }
      });
      pastilles.forEach(function (p) {
        var identifiant = p.getAttribute("data-cle");
        p.classList.toggle("attenue", !(vises[identifiant] || touches[identifiant]));
      });
      cartes.forEach(function (c) {
        var sid = c.getAttribute("data-cle");
        var liste = (D.services[sid] || {}).outils || [];
        var actif = liste.some(function (o) { return vises[o] || touches[o]; });
        c.classList.toggle("attenue", !actif);
      });
    }

    pastilles.forEach(function (p) {
      p.addEventListener("mouseenter", function () {
        allumer([p.getAttribute("data-cle")]);
        p.classList.add("souligne");
      });
      p.addEventListener("mouseleave", eteindre);
    });
    cartes.forEach(function (c) {
      c.addEventListener("mouseenter", function () {
        allumer((D.services[c.getAttribute("data-cle")] || {}).outils || []);
        c.classList.add("souligne");
      });
      c.addEventListener("mouseleave", eteindre);
    });
    fils.forEach(function (fil, rang) {
      fil.addEventListener("mouseenter", function () {
        allumer(extremites[rang]);
        fil.classList.remove("attenue");
        fil.classList.add("souligne");
      });
      fil.addEventListener("mouseleave", eteindre);
    });
  }

  // ── synthèse : sélecteur de service, tuiles et graphiques ─────────────────
  var synthese = document.getElementById("synthese");
  if (!synthese || typeof Chart === "undefined") { return; }

  var charte = D.charte;
  Chart.defaults.font.family = charte.police;
  Chart.defaults.font.size = 10;
  Chart.defaults.color = charte.gris;
  Chart.defaults.animation = false;
  Chart.defaults.plugins.legend.labels.boxWidth = 10;
  Chart.defaults.plugins.legend.labels.boxHeight = 10;

  var onglets = Array.prototype.slice.call(synthese.querySelectorAll(".onglet"));
  var tuiles = Array.prototype.slice.call(synthese.querySelectorAll(".tuile-valeur"));
  var toiles = Array.prototype.slice.call(synthese.querySelectorAll("canvas"));
  var graphiques = {};
  var courant = D.tous;

  function pourcent(partie, total) {
    return total ? Math.round((100 * partie) / total) + " %" : "—";
  }
  function valeurTuile(vue, cle) {
    if (cle === "pct_auto") { return pourcent(vue.flux_mode.auto, vue.nb_flux); }
    if (cle === "pct_interne") { return pourcent(vue.outils_ie.interne, vue.nb_outils); }
    if (cle === "nb_todo_ie") { return String(vue.outils_ie.todo); }
    return String(vue[cle]);
  }

  function construire(toile, spec, vue) {
    var contexte = toile.getContext("2d");
    if (spec.genre === "doughnut") {
      var valeurs = Object.keys(vue[spec.donnees]).map(function (cle) {
        return vue[spec.donnees][cle];
      });
      return new Chart(contexte, {
        type: "doughnut",
        data: {
          labels: spec.libelles,
          datasets: [{ data: valeurs, backgroundColor: spec.couleurs,
                       borderColor: "#fff", borderWidth: 2 }]
        },
        options: {
          responsive: true, maintainAspectRatio: false, cutout: "58%",
          plugins: {
            legend: { position: "right" },
            tooltip: { callbacks: { label: function (ctx) {
              var total = ctx.dataset.data.reduce(function (a, b) { return a + b; }, 0);
              return ctx.label + " : " + ctx.raw + " (" + pourcent(ctx.raw, total) + ")";
            } } }
          }
        }
      });
    }
    var classement = vue.classement;
    return new Chart(contexte, {
      type: "bar",
      data: {
        labels: classement.map(function (l) { return l.nom; }),
        datasets: [{
          data: classement.map(function (l) { return l[spec.champ]; }),
          backgroundColor: classement.map(function (l) {
            return (courant !== D.tous && l.id === courant) ? charte.accent : spec.couleurs[0];
          }),
          borderWidth: 0, barPercentage: 0.78
        }]
      },
      options: {
        indexAxis: "y", responsive: true, maintainAspectRatio: false,
        onClick: function (evenement, elements) {
          if (!spec.cliquable || !elements.length) { return; }
          var cible = document.getElementById("service-" + classement[elements[0].index].id);
          if (cible) { cible.scrollIntoView({ behavior: "smooth", block: "start" }); }
        },
        plugins: { legend: { display: false } },
        scales: {
          x: { beginAtZero: true, ticks: { precision: 0 },
               grid: { color: charte.trait } },
          y: { grid: { display: false },
               ticks: { autoSkip: false, font: { size: 9 } } }
        }
      }
    });
  }

  function rendre() {
    var vue = D.vues[courant];
    tuiles.forEach(function (noeud) {
      noeud.textContent = valeurTuile(vue, noeud.getAttribute("data-cle"));
    });
    onglets.forEach(function (onglet) {
      var actif = onglet.getAttribute("data-cle") === courant;
      onglet.style.background = actif ? charte.primary : charte.blanc;
      onglet.style.borderColor = actif ? charte.primary : charte.trait;
      onglet.style.color = actif ? charte.blanc : charte.gris;
      onglet.style.fontWeight = actif ? "700" : "400";
    });
    toiles.forEach(function (toile) {
      var spec = JSON.parse(toile.getAttribute("data-spec"));
      if (graphiques[spec.id]) { graphiques[spec.id].destroy(); }
      graphiques[spec.id] = construire(toile, spec, vue);
    });
  }

  onglets.forEach(function (onglet) {
    onglet.addEventListener("click", function () {
      courant = onglet.getAttribute("data-cle");
      rendre();
    });
  });
  rendre();
}());
"""


# ═════════════════════════════════════════════════════════════════════════════
# Assemblage
# ═════════════════════════════════════════════════════════════════════════════

def donnees_js(carto):
    """Modèle sérialisé pour le navigateur : surbrillance et filtre de synthèse."""
    services = OrderedDict()
    for service in carto["services"]:
        services[service["id"]] = OrderedDict([
            ("nom", service["nom"]),
            ("outils", [carto["outils"][n]["id"] for n in service["outils"]]),
        ])
    return OrderedDict([
        ("charte", carto["charte"]),
        ("tous", model.TOUS),
        ("services", services),
        ("vues", model.tous_les_indicateurs(carto)),
    ])


def barre_navigation(pages):
    liens = ['<span class="titre">Cartographie IT — chantier</span>']
    for page in pages:
        liens.append('<a href="#%s">%s</a>' % (page["id"], echapper(page["titre"])))
    return '<nav id="barre">%s</nav>' % "".join(liens)


def generer(carto, chemin):
    pages = layout.construire(carto)
    corps = [rendre_page(page, pages) for page in pages]
    html = [
        "<!doctype html>",
        '<html lang="fr">',
        "<head>",
        '<meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width,initial-scale=1">',
        "<title>%s</title>" % echapper(carto["meta"]["titre"]),
        "<style>%s</style>" % feuille_de_style(carto["charte"]),
        "</head>",
        "<body>",
        barre_navigation(pages),
        '<div id="bulle"></div>',
        "\n".join(corps),
        '<script src="vendor/chart.min.js"></script>',
        "<script>window.DONNEES=%s;</script>" % json.dumps(
            donnees_js(carto), ensure_ascii=False),
        "<script>%s</script>" % JS,
        "</body>",
        "</html>",
    ]
    chemin.parent.mkdir(parents=True, exist_ok=True)
    chemin.write_text("\n".join(html), encoding="utf-8")
    return pages


def main():
    analyseur = argparse.ArgumentParser(description=__doc__)
    analyseur.add_argument("--out", default=str(SORTIE),
                           help="fichier HTML à écrire")
    arguments = analyseur.parse_args()

    carto = model.charger()
    chemin = Path(arguments.out)
    pages = generer(carto, chemin)
    poids = chemin.stat().st_size / 1024.0
    print("%s : %d pages, %.0f Ko" % (chemin, len(pages), poids))
    vendor = chemin.parent / "vendor" / "chart.min.js"
    if not vendor.exists():
        print("  attention : %s est absent, les graphiques ne s'afficheront pas"
              % vendor)


if __name__ == "__main__":
    main()
