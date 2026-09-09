# Cartographie des outils IT — contexte projet

Cartographie du SI d'un chantier Bouygues Construction, produite depuis une base
d'entretiens Excel. Deux livrables issus d'une source unique : une page HTML
autonome et un PowerPoint **en formes natives éditables** (jamais une image).

Tout le contenu visible est en **français** ; les commentaires du code aussi.

## Chaîne de génération

```
data/BDD_Cartographie_Outils_IT.xlsx
        │  data/build_data.py     (agrège, normalise, calcule missions et chaînes)
        ▼
index.html   ← const BDD = {...} injecté entre les marqueurs « DONNÉES BDD »
        │  pptx/build.py + pptx/extract.js   (Playwright mesure le DOM réel)
        ▼
dist/Cartographie-Outils-IT.pptx
```

**`index.html` est la source de vérité géométrique.** `build.py` ouvre la page
dans Chromium, relève la position réelle de chaque élément via `extract.js`, et la
rejoue en formes natives. Ne jamais coder de coordonnées en dur dans `build.py` :
si la mise en page bouge, le PPTX suit tout seul. `scrape()` ramasse **un `.slide`
par diapo, dans l'ordre du DOM** — ajouter une planche ne demande aucune
modification de `build.py`.

## Commandes

```bash
pip install python-pptx playwright openpyxl && playwright install chromium

python3 data/build_data.py                                   # xlsx → modèle dans index.html
python3 pptx/build.py --out dist/Cartographie-Outils-IT.pptx # index.html → pptx
python3 pptx/build.py --chromium /chemin/vers/chrome         # si Playwright n'a pas son binaire
```

`build_data.py` signale sur stderr, sans jamais casser la sortie : les outils d'un
pôle non rattachés à une mission, les outils rattachés à aucun pôle, les étapes
citant un outil inconnu, les liens visant une mission inexistante, et les paires de
pôles non voisines dans la grille inter-services. Un signalement veut dire qu'une
constante est à compléter.

## Les deux modèles de mission — la distinction structurante

| Modèle | Constante | Pôles | Rendu |
|---|---|---|---|
| **chaîne** | `PROCESS` | daf, ctr, trv, tun, top, sec | entrée → étapes ordonnées → finalité |
| **inventaire** | `MISSIONS` | dir, met, qse | « Début à documenter » → outils → « Finalité à documenter » |

La BDD ne porte **pas** l'ordre des étapes : `PROCESS` est saisi à la main depuis le
brief métier, comme `FLOWS`. Un pôle présent dans `PROCESS` ignore `MISSIONS` et ne
retient **que** les outils cités dans ses chaînes — c'est ce mécanisme qui a retiré
AutoCAD de la Topographie et rattaché Pablo au tunnel. Retirer un pôle de `PROCESS`
le fait retomber sur `MISSIONS` sans autre modification.

Forme d'une mission de `PROCESS` :

```python
{"label": "Contrôle budgétaire",
 "note": "conjoint avec les Travaux",      # sous-titre ; "todo": True le remplace
 "in":    [{"lab": "Extraction matériel", "tool": "BYMAT"}],   # tool facultatif
 "steps": [{"act": "Constitution de la PC100", "tool": "Excel"}],
 "out":   {"lab": "Budget par atelier arbitré"},
 "links": [{"at": "out", "to": "trv", "mission": "Contrôle budgétaire", "obj": "PC100…"}]}
```

`at` vaut `"in"`, `"out"` ou l'index d'une étape. **Un `link` produit deux choses à
la fois** : le badge posé sur l'affiche du pôle, et le fil tracé sur la planche
« Flux inter-services ». Une seule saisie, deux rendus — ne jamais les désynchroniser.
`mission` doit correspondre exactement à un libellé du pôle cible, `PROCESS` ou
`MISSIONS` : `build_data.py` refuse et signale le reste.

## Modèle de données (`const BDD` dans index.html)

| Clé | Contenu |
|---|---|
| `services[]` | les 9 pôles : `{id, name, kicker, tools[], missions[], mode}` |
| `services[].mode` | `"process"` ou `"inventory"` — commande le rendu de l'affiche |
| `services[].missions[]` | `{label, todo, tools[]}` — vue à plat, quel que soit le modèle |
| `process{}` | par pôle en chaînes : les missions détaillées ci-dessus |
| `socle[]` | groupes du socle commun : `{group, tools[]}` |
| `core[]` | SharePoint, Power BI (bloc « socle de données ») |
| `tools{}` | par outil : `{ie, alim, ed, svc[], nb, socle, fonction, todo?, brief?}` |
| `flows[]` | flux applicatifs d'outil à outil : `{from, to, obj, alim, map}` |
| `xflows[]` | échanges de mission à mission, dérivés des `links` |
| `board{}` | `{top[], bottom[]}` — répartition des cartes de la planche 1 |
| `bornes{}` | `"pôle|libellé"` → début et finalité d'une mission d'inventaire |
| `chantier` | nom du chantier, repris au bandeau de la planche 1 |

`brief: true` marque un outil venu du brief et absent de la BDD (`HORS_BDD`) : il
s'affiche en pointillés, porte « hors BDD · brief » et son nom est coloré dans le
référentiel.

Extrémités de `flows` : `"s:Excel"` pour le socle, `"daf:Basware"` pour un outil de
pôle. Extrémités de `xflows` : `"daf:Contrôle budgétaire"` — **une mission**, pas un
outil. `map: false` documente sans tracer : pour `flows`, le flux indirect
Trimble Connect → IDCapture ; pour `xflows`, les enchaînements internes à un pôle.

Identifiants DOM : `{bid}-s-{slug}` (socle), `{bid}-t-{pôle}-{slug}` (outil ou
mission). `bid` vaut `b1` (cartographie), `x` (flux inter-services), `f-{pôle}`
(affiches). **Un `bid` distinct par planche est obligatoire** : sinon
`getElementById` renvoie l'élément de la planche 1 et les fils partent au mauvais
endroit.

## Où intervenir

| Besoin | Endroit |
|---|---|
| Chaînes de flux, étapes, finalités, interconnexions | `PROCESS` dans `data/build_data.py` |
| Missions des pôles sans processus recueilli | `MISSIONS` |
| Phrase de fonction d'un outil (référentiel) | `FONCTIONS` |
| Début / finalité d'une mission d'inventaire | `BORNES` |
| Outil volontairement rattaché à aucun pôle | `HORS_POLE` |
| Nom du chantier | `CHANTIER` |
| Composition des pôles | `SERVICE` + `SERVICES` |
| Outils du brief absents de la BDD | `HORS_BDD` |
| Composition du socle | `SOCLE` + `CORE` |
| Flux applicatifs | `FLOWS` |
| Libellés d'outils saisis en vrac | `ALIAS` |
| Répartition des cartes, planche 1 | `BOARD_TOP` + `BOARD_BOTTOM` |
| Ordre des affiches | `AFFICHES` dans `index.html` |
| Couleurs | `--met --top --tun --trv --ctr --qse --dir --daf --sec --core` en tête du CSS |

## Les 12 planches

1. Cartographie applicative — 9 pôles, socle en bandeau, flux d'outil à outil
2-10. Affiches services, dans l'ordre `AFFICHES` : dir, daf, ctr, met, trv, tun, top, qse, sec
11. Matrice des flux — les deux tableaux, applicatif et inter-services
12. Référentiel des outils, avec la fonction de chaque outil

La planche « Flux inter-services » a été retirée sur arbitrage métier. Les `xflows`
restent au modèle : ils alimentent le second tableau de l'annexe 1 et les badges
posés sur les affiches. Rien à reconstruire pour la rétablir, sinon la planche.

### Mise en page

Format natif 1600 × 900 px = 13,333 × 7,5 po (1 px = 0,6 pt = 7620 EMU).

`.board` et `.affiche` occupent la même zone. `.board` répartit ses rangées en
`space-between` avec des rangées à hauteur de contenu : **les couloirs entre rangées
sont l'espace de routage des fils**, les resserrer ramène les croisements.

Une affiche en chaînes : une `.pline` par mission, `flex:1 1 0` plafonné à 210 px,
`space-evenly`. Les blocs `.pin` / `.pstep` / `.pout` se partagent la largeur
restante à parts égales — c'est ce qui **aligne verticalement les colonnes entrée et
finalité d'une bande à l'autre**, et rend l'affiche scannable. Jusqu'à 8 blocs par
chaîne (`137 + 7×(20+137) = 1236 px` sur 1326 px utiles) ; la plus longue aujourd'hui
en compte 7 (Contrat, gestion des encaissements). Six bandes au maximum par affiche.

Sur l'affiche « inventaire », `.pin` et `.pout` sont à largeur fixe (152 px) : en
`flex:1 1 0` ils avalaient toute la bande dès qu'une mission ne comptait que deux
outils.

Les deux annexes portent la légende des couleurs de pôle (`svcLegendHTML()`), sans
laquelle elles ne se lisent pas seules : elles encodent le pôle par la couleur mais
ne le nomment nulle part. La planche 1 n'en porte plus, sur arbitrage métier. **Trois entrées par colonne au maximum** — `.masthead` a
une hauteur figée à 84 px et `.wrap` commence à 134 px, un masthead plus haut vient
buter sur le tableau. L'annexe 1 inclut l'entrée « Socle commun » (sa matrice a des
extrémités `s:` rendues en `--core`), l'annexe 2 non (sa colonne `svcdots` ne montre
que des pôles).

## Connecteurs

Tracés en Bézier cubique unique (une seule courbe par flux : le générateur PPTX
n'émet qu'un `cubicBezTo`). Le routage choisit parmi trois cas :

1. **inter-zones ou inter-rangées** → vertical, du bord horizontal le plus proche
   de la cible ;
2. **pastilles alignées verticalement** → arc par le couloir central du conteneur,
   pour ne pas traverser les pastilles intermédiaires ;
3. **sinon** → S horizontal ; en dessous de 26 px d'écart, trait droit.

Les extrémités partagées sont réparties verticalement par `spread()`, triées selon
la position de l'extrémité opposée : c'est ce qui évite les croisements.

## Pièges vérifiés — ne pas les redécouvrir

**Métriques de police.** Chromium compose ici en DejaVu Sans, PowerPoint en
Calibri. Tout texte exporté peut être plus large que mesuré. Ces défauts
n'apparaissent **jamais** dans le rendu HTML : seule la conversion en PDF les
révèle.
- Centrer le texte des pastilles et des blocs (`text-align:center`) : l'écart
  devient symétrique au lieu de déborder à droite.
- Ne jamais poser deux fragments de texte côte à côte en comptant sur la largeur
  mesurée : ils se chevauchent. Soit un seul paragraphe multi-runs, soit une
  gouttière large. Les titres de section de la planche 12 étaient en `display:flex`
  avec leur compteur : les deux fragments se sont chevauchés à l'export. Repassés en
  `display:block` avec un `<em>` inline, ils fusionnent en un paragraphe.
- Les badges d'interconnexion tiennent sur **deux lignes** (`<br>` entre le pôle et
  la mission) : sur une seule, « → Sécurité · Accueil, formation et habilitations »
  débordait de son bloc.
- Les libellés de missions sur deux colonnes ont déjà débordé d'une colonne à
  l'autre : d'où les libellés courts et la gouttière à 26 px.
- **Tronquer en JS, jamais par `text-overflow: ellipsis`.** `extract.js` relève le
  texte du DOM, pas le texte visuellement coupé par le CSS : une cellule raccourcie
  à l'écran ressort en pleine largeur dans le PPTX et chevauche la colonne suivante.
  D'où le helper `cut()` et le budget de 46 caractères de la colonne « Fonction »,
  relevé à la règle sur la largeur réelle puis diminué de 12 % pour Calibri. Mesurer
  la largeur d'un texte comme le fait `extract.js` — un `Range` sur le contenu — et
  non par `scrollWidth`, plafonné sur une cellule de tableau.

**Extraction (`extract.js`).**
- `color-mix()` est calculé par Chromium en `color(srgb r g b / a)`, pas en
  `rgba()` — les deux formes sont gérées, ne pas simplifier.
- `inlineRuns()` fusionne les fragments inline en un seul paragraphe, mais
  **abandonne** si l'élément contient un `<br>` ou un descendant peint (pastille,
  puce) : ces éléments sont exportés comme formes à leur position propre, le texte
  ne peut pas être fusionné avec eux. Le repli ligne par ligne est correct — c'est
  même lui qu'on recherche en posant un `<br>`.
- La casse CSS (`text-transform`) doit être appliquée au texte propre de l'élément
  autant qu'à celui de ses enfants ; sinon un titre en capitales ressort en
  minuscules dans le PPTX.
- Le trait pointillé se lit sur la classe `manuel` posée par `index.html`, pas sur
  une classe `dashed` : la chercher a longtemps fait sortir tous les flux en trait
  plein.
- Le texte enrichi est positionné sur la **boîte de contenu** (hors padding),
  sinon il se décale de la valeur du padding.

**Génération PPTX (`build.py`).**
- Appeler `no_shadow()` sur chaque forme : sans cela le thème applique une ombre
  par défaut.
- `custGeom` : le `<a:path>` reçoit ses dimensions en EMU et les points sont
  relatifs à la boîte englobante de la courbe, pas à la planche.
- Renseigner remplissage et contour **avant** l'ombre : `effectLst` doit rester le
  dernier enfant de `spPr`.
- `add_text` pose `word_wrap=False` avec 90 px de marge : un texte trop long
  déborde silencieusement au lieu de se replier. C'est la contrainte qui impose des
  libellés courts.

**Environnement.** Le répertoire `pptx/` du dépôt masque le paquet `python-pptx`
lorsqu'on teste `import pptx` depuis la racine : la vérification passe alors qu'il
n'est pas installé. Tester depuis un autre répertoire.

**Regex de slug** : la plage de marques combinantes doit rester écrite `[\u0300-\u036f]`. Avec les caractères combinants littéraux, la regex ne
survit pas à une réécriture du fichier et `slug()` cesse de dépouiller les
accents — les identifiants DOM ne correspondent alors plus aux flux.

## Vérification avant de livrer

1. `python3 data/build_data.py` — vérifier le décompte (9 pôles, 35 missions,
   64 outils, 14 flux applicatifs, 18 flux inter-services) et les avertissements.
2. Capturer les **12 planches** avec Playwright et **les regarder**. Contrôler en
   plus par script deux choses que l'œil rate : qu'aucun élément ne sorte de son
   conteneur (`.wrap`, `.affiche`, `.board` — pas seulement de la planche), et
   qu'aucune erreur ne remonte en console.
3. `python3 pptx/build.py` puis `validate.py` du skill pptx.
4. Convertir le PPTX en PDF (LibreOffice, paquet `libreoffice-impress` requis) et
   inspecter chaque page : c'est là et nulle part ailleurs que les défauts de
   métriques de police apparaissent.

## Décisions à ne pas défaire sans arbitrage

- **Composition du socle commun** : la colonne `Nature` de la BDD est
  contradictoire d'un entretien à l'autre. Le socle retenu = outils
  majoritairement tagués socle + plateformes transverses de la cartographie
  d'origine (Excel, Word, SharePoint, Power BI, E-Paraph, DocuSign), plus Outlook,
  que le brief fait porter la commande de MO et l'envoi des situations.
- **Contrat est un pôle** : le contract management était rattaché à la DAF ; il est
  autonome depuis que la chaîne de facturation partenaire circule explicitement
  entre les deux.
- **Nature des flux** : la BDD ne qualifie pas l'automatisation flux par flux. Un
  flux est dit automatisé quand son outil de **destination** est alimenté
  automatiquement. Les `xflows`, issus du brief, sont tous déclarés manuels.
- **Marqueurs de réserve non affichés** : « ordre non attesté », « mission
  proposée », « chaîne à valider » et les notes de bas de planche ont été retirés du
  rendu sur arbitrage métier. Le champ `todo` reste au modèle — l'information n'est
  pas perdue, seulement plus montrée. Conséquence assumée : rien ne distingue plus à
  l'œil un processus recueilli en entretien d'un processus déduit, dont les missions
  du tunnel décalquées de celles des Travaux.
- **Répétitions assumées** : un outil apparaît dans plusieurs pôles ou missions
  quand plusieurs l'ont déclaré. C'est voulu, la carte est organisée par pôle.
- **Enchaînements internes non tracés** : un `link` d'un pôle vers lui-même reste
  en badge et en ligne de tableau, mais n'est pas dessiné sur la planche
  inter-services — le fil repasserait par-dessus les missions de la carte.
- `archive/` conserve la première version, bâtie sur le seul PDF initial. Ne pas
  la régénérer, elle n'est plus alimentée par la BDD.
