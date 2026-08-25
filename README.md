# Cartographie des outils IT — chantier Bouygues Construction

Deux livrables, une seule source : `data/carto.json`.

| Livrable | Fichier | Contenu |
|---|---|---|
| Page web | `out/carto.html` | 12 pages, interactive, 100 % hors ligne |
| Présentation | `out/carto.pptx` | 12 diapositives, formes natives éditables |

Les 12 pages : 1 vue macro · 9 affiches de service · 1 référentiel des outils ·
1 synthèse chiffrée.

## Chaîne de génération

```
data/BDD_Cartographie_Outils_IT.xlsx      base d'entretiens, lecture seule
        │  data/extract_bdd.py
        ▼
data/carto.json  ★ source unique          + data/gap-report.md
        │  build/model.py    indicateurs dérivés
        │  build/layout.py   graphe de scène, 960 × 540 pt
        ├──────────────┬─────────────────┐
        ▼              ▼
build/generate_html.py  build/generate_pptx.py
        ▼              ▼
out/carto.html     out/carto.pptx
```

**Toute modification de contenu passe par `data/carto.json`**, jamais par le
code de rendu : ni `generate_html.py` ni `generate_pptx.py` ne contiennent de
nom d'outil, de service, de flux ou de couleur.

`build/layout.py` est le pivot : il calcule une géométrie unique, en points,
que les deux générateurs se contentent de rendre. Le PPTX ne peut donc pas
diverger du HTML — il n'y a pas deux mises en page à maintenir.

## Commandes

```bash
pip install python-pptx openpyxl                    # génération
pip install playwright pypdfium2                    # vérification (facultatif)

python3 data/extract_bdd.py                         # xlsx → carto.json + gap-report
python3 build/generate_html.py                      # carto.json → out/carto.html
python3 build/generate_pptx.py                      # carto.json → out/carto.pptx
```

Modifier `data/carto.json` puis relancer les deux générateurs suffit à tout
mettre à jour, synthèse comprise. Relancer `extract_bdd.py` écrase
`carto.json` : les retouches manuelles doivent être reportées dans les
constantes en tête du script.

## Hors ligne

`out/carto.html` s'ouvre en double-clic, sans serveur ni réseau. La seule
ressource externe est `out/vendor/chart.min.js` (Chart.js 4.4.1, vendorisé,
chemin relatif). Le HTML généré ne contient aucune URL :

```bash
grep -nE 'https?://' out/carto.html      # doit ne rien renvoyer
```

`out/vendor/` doit accompagner `carto.html` lors d'une copie ou d'un envoi.

## Ce que fait chaque page

**Vue macro.** Neuf cartes de service autour du hub de données (SharePoint,
OneDrive, Probity, Mezzoteam). Trait plein rouge = flux automatique, pointillé
gris = flux manuel, flèche pour le sens. Survol d'une carte, d'une pastille ou
d'un fil : les flux liés restent vifs, le reste s'estompe. Clic sur une carte :
affiche du service. En PPTX, le clic est un lien hypertexte natif.

**Affiches de service.** Une bande par processus : début ▸ étapes ▸ fin. Chaque
étape porte son mode d'alimentation et la nature interne/externe de l'outil.
La BDD ne renseigne aucun début ni aucune fin : ces blocs sont vides,
dimensionnés et éditables, marqués `TODO`.

**Référentiel.** Les 61 outils, leur fonction en une ligne, leurs services,
nature, interne/externe, alimentation et origine, avec la légende des couleurs.

**Synthèse.** Six indicateurs et quatre graphiques, tous calculés depuis
`carto.json`. En HTML, le sélecteur de service recalcule compteurs et
graphiques sur le périmètre choisi ; le PPTX présente la vue « Tous services »,
les données des graphiques restant éditables dans leur classeur intégré.

## Lire `data/carto.json`

| Clé | Contenu |
|---|---|
| `charte` | tokens de couleur et polices — seules couleurs autorisées |
| `hub` | les 4 outils du hub de données |
| `services[]` | `{id, nom, court, kicker, vide, outils[], processus[]}` |
| `services[].processus[]` | `{nom, debut, fin, etapes[], origine, ordre_atteste, todo}` |
| `outils{}` | `{nom, id, fonction, nature, ie, alim, editeur, services[], hub, origine, lignes[], usages[], todo[]}` |
| `flux[]` | `{de, vers, mode, objet, carte, origine, ligne, todo}` |

`origine` trace la provenance : `bdd` (table d'entretiens), `bdd-annexe`
(onglet Tableau de Bord), `correction` (correction métier), `brief` (cahier
des charges). Tout ce qui n'est pas `bdd` est badgé dans le référentiel.

`carte: false` documente un flux sans le tracer sur la vue macro — utilisé
pour la relation indirecte Trimble Connect → IDCapture, qui passe par des
captures d'écran.

`debut` ou `fin` à `null` déclenche le bloc vide éditable du livrable 2.

## Ce que la BDD ne dit pas

`data/gap-report.md` liste en douze sections tout ce que la base ne permet pas
de trancher et tout ce qui a été ajouté hors d'elle. À lire avant de conclure
quoi que ce soit d'un chiffre. En résumé :

- la BDD **ne porte pas** de champ `socle_commun` ; la colonne `Nature` est
  contradictoire d'un entretien à l'autre pour 6 outils ;
- 20 flux seulement sont exploitables sur 99 lignes — 39 cellules vides,
  27 « Aucun », 7 cibles génériques, 6 cibles qui ne sont pas des outils ;
- aucun processus n'a de début ni de fin : les blocs `TODO` du livrable 2 sont
  attendus, pas un défaut de génération ;
- l'ordre des étapes d'un processus issu de la BDD n'est attesté par rien ;
- `Fréquence d'utilisation` et `Niveau Maturité` ne sont ni extraits ni
  affichés ;
- **Sécurité** est un service voulu vide : trois blocs vierges éditables,
  aucun outil prélevé aux autres services ;
- **RNDTS** reste à valider avec Camille ; **BYCN** et **CEMEX** portent un
  `TODO` visible.

## Vérifier une modification

```bash
python3 data/extract_bdd.py            # 9 services, 43 processus, 61 outils, 20 flux
python3 build/generate_html.py
python3 build/generate_pptx.py
grep -nE 'https?://' out/carto.html    # aucun résultat attendu
python3 /mnt/skills/public/pptx/scripts/office/validate.py out/carto.pptx
```

Puis, pour ce qui ne se voit qu'à l'œil : capturer les 12 pages HTML avec
Playwright et les regarder, convertir le PPTX en PDF avec LibreOffice
(`libreoffice-impress` requis) et relire chaque page. Les défauts de métriques
de police n'apparaissent que dans le PDF converti, jamais dans le rendu HTML.

## Pièges vérifiés — ne pas les redécouvrir

**Métriques de police.** Chromium compose en DejaVu Sans, PowerPoint en
Calibri, le poste de travail en Segoe UI. `layout.largeur()` mesure en Calibri
et applique un facteur de sécurité ; les libellés d'une ligne ne se replient
jamais, ni en HTML (`white-space:nowrap`) ni en PPTX (`word_wrap = False`),
sans quoi ils débordent de leur forme dans un rendu et pas dans l'autre.

**DOM plate.** Les formes du HTML sont posées à plat, en position absolue :
un libellé posé sur une carte intercepterait le survol destiné à la carte,
d'où `pointer-events: none` sur tous les blocs de texte.

**Ancrage des flux.** Un outil déclaré par plusieurs services a plusieurs
pastilles ; le routage retient le couple le plus court, sinon les fils
traversent la planche. Les outils du hub font exception : leur seule ancre
légitime est celle du bandeau central.

**Routage.** Hub ↔ hub passe en arc sous le bandeau pour ne pas traverser les
pastilles alignées avec lui ; hub ↔ rangée sort verticalement pour rester dans
le couloir entre les zones ; deux pastilles voisines sont reliées par un arc
par-dessus, faute de quoi le flux se réduit à un point invisible.

**PPTX.** `no_shadow()` sur chaque forme, sinon le thème applique une ombre par
défaut. Remplissage et contour renseignés **avant** l'ombre : `effectLst` doit
rester le dernier enfant de `spPr`. Les points d'un `custGeom` sont relatifs à
la boîte englobante de la courbe, pas à la planche. PowerPoint trace la
première catégorie d'un graphique en barres **en bas** : l'ordre est inversé à
l'écriture. Les libellés d'axe sont forcés à une étiquette par catégorie, sinon
une sur deux disparaît.

**Regex de slug.** La plage de marques combinantes doit rester écrite en
séquences d'échappement `[̀-ͯ]`. Avec les caractères combinants
littéraux, la regex ne survit pas à une réécriture du fichier et `slug()`
cesse de dépouiller les accents.

## Décisions à ne pas défaire sans arbitrage

- **Neuf services**, dont Sécurité volontairement vide. Les 11 libellés bruts
  de la BDD y sont rattachés par la table `RATTACHEMENT`.
- **Socle commun** = majorité des entretiens, égalité → socle, les 4 outils du
  hub forcés socle. OneDrive et Probity sont absents de la BDD : ils figurent
  au hub sur consigne, grisés et marqués `TODO`.
- **Interne / externe** = majorité ; égalité ou colonne vide → `TODO` visible,
  jamais une valeur devinée. Cela fait 14 `TODO`, et une troisième part dans le
  graphique de la synthèse.
- **Nature des flux** : la BDD ne qualifie jamais l'automatisation flux par
  flux. Un flux est dit automatisé quand son outil de **destination** est
  alimenté automatiquement.
- **Sens du pointage** : la ligne 52 de la BDD porte PUMA comme source d'un
  texte qui décrit l'inverse. C'est la correction métier « processus paye »
  qui a été appliquée — pointage PUMA → BIP / BYCN — et la ligne écartée.
- **Répétitions assumées** : un outil apparaît dans plusieurs services quand
  plusieurs l'ont déclaré. C'est voulu, la carte est organisée par service.
