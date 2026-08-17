# Cartographie des outils IT — contexte projet

Cartographie du SI d'un chantier Bouygues Construction, produite depuis une base
d'entretiens Excel. Deux livrables issus d'une source unique : une page HTML
autonome et un PowerPoint **en formes natives éditables** (jamais une image).

Tout le contenu visible est en **français** ; les commentaires du code aussi.

## Chaîne de génération

```
data/BDD_Cartographie_Outils_IT.xlsx
        │  data/build_data.py     (agrège, normalise, calcule les missions)
        ▼
index.html   ← const BDD = {...} injecté entre les marqueurs « DONNÉES BDD »
        │  pptx/build.py + pptx/extract.js   (Playwright mesure le DOM réel)
        ▼
dist/Cartographie-Outils-IT.pptx
```

**`index.html` est la source de vérité géométrique.** `build.py` ouvre la page
dans Chromium, relève la position réelle de chaque élément via `extract.js`, et la
rejoue en formes natives. Ne jamais coder de coordonnées en dur dans `build.py` :
si la mise en page bouge, le PPTX suit tout seul.

## Commandes

```bash
pip install python-pptx playwright openpyxl && playwright install chromium

python3 data/build_data.py                                   # xlsx → modèle dans index.html
python3 pptx/build.py --out dist/Cartographie-Outils-IT.pptx # index.html → pptx
python3 pptx/build.py --chromium /chemin/vers/chrome         # si Playwright n'a pas son binaire
```

`build_data.py` signale sur stderr tout outil d'un pôle non rattaché à une
mission (« outils sans mission ») — un signalement veut dire que `MISSIONS` est à
compléter, pas que la sortie est cassée.

## Modèle de données (`const BDD` dans index.html)

| Clé | Contenu |
|---|---|
| `services[]` | les 7 pôles : `{id, name, kicker, missions[]}` |
| `services[].missions[]` | `{label, todo, tools[]}` — `todo` = mission proposée, non documentée en BDD |
| `socle[]` | groupes du socle commun : `{group, tools[]}` |
| `core[]` | SharePoint, Power BI (bloc « socle de données ») |
| `tools{}` | par outil : `{ie, alim, ed, svc[], nb, socle, todo}` |
| `flows[]` | `{from, to, obj, alim, map}` |

Extrémités de flux : `"s:Excel"` pour le socle, `"daf:Basware"` pour un outil de
pôle. `map: false` documente le flux dans la matrice sans le tracer sur la carte
(utilisé pour le flux indirect Trimble Connect → IDCapture, par captures d'écran).

Identifiants DOM : `b1-s-{slug}` (socle) et `b1-t-{pôle}-{slug}` (outil de pôle).
Un outil servant plusieurs missions d'un même pôle voit ses occurrences suivantes
suffixées `-2`, `-3` — la première garde l'identifiant visé par les flux.

## Où intervenir

| Besoin | Endroit |
|---|---|
| Missions, libellés, rattachements | `MISSIONS` dans `data/build_data.py` |
| Composition des pôles | `SERVICE` + `SERVICES` |
| Composition du socle | `SOCLE` + `CORE` |
| Flux | `FLOWS` |
| Libellés d'outils saisis en vrac | `ALIAS` |
| Répartition haut/bas des pôles | `BOARD` en tête du script de `index.html` |
| Couleurs | `--met --top --tun --trv --qse --dir --daf --core` en tête du CSS |

## Mise en page de la planche 1

Format natif 1600 × 900 px = 13,333 × 7,5 po (1 px = 0,6 pt = 7620 EMU).

Quatre pôles techniques en haut, socle en **bandeau horizontal** au centre, trois
pôles support en bas. Ce n'est pas un choix esthétique : 31 missions et ~110
pastilles ne tiennent pas dans une mise en page à socle vertical et deux colonnes.
Deux leviers ont rendu la page unique possible, à préserver si le volume grandit :

- les outils du socle mobilisés par une mission sont **cités en ligne**
  (« via Excel · Power BI ») au lieu d'être répétés en pastilles ;
- les cartes de la rangée du bas (larges) affichent leurs missions **sur deux
  colonnes** (`.card.wide`).

Si le contenu déborde à nouveau, le repli déjà éprouvé est de scinder la
cartographie en deux planches (pôles techniques / pôles support), le socle étant
repris au centre de chacune — voir l'historique git.

## Connecteurs

Tracés en Bézier cubique unique (une seule courbe par flux : le générateur PPTX
n'émet qu'un `cubicBezTo`). Le routage choisit parmi trois cas :

1. **inter-zones** (rangée de pôles ↔ bandeau) → vertical, du bord horizontal le
   plus proche de la cible ;
2. **pastilles alignées verticalement** → arc par le couloir central du conteneur,
   pour ne pas traverser les pastilles intermédiaires ;
3. **sinon** → S horizontal ; en dessous de 26 px d'écart, trait droit.

Les extrémités partagées sont réparties verticalement par `spread()`, triées selon
la position de l'extrémité opposée : c'est ce qui évite les croisements.

## Pièges vérifiés — ne pas les redécouvrir

**Métriques de police.** Chromium compose ici en DejaVu Sans, PowerPoint en
Calibri. Tout texte exporté peut être plus large que mesuré.
- Centrer le texte des pastilles (`text-align:center`) : l'écart devient
  symétrique au lieu de déborder à droite.
- Ne jamais poser deux fragments de texte côte à côte en comptant sur la largeur
  mesurée : ils se chevauchent. Soit un seul paragraphe multi-runs, soit une
  gouttière large.
- Les libellés de missions sur deux colonnes ont déjà débordé d'une colonne à
  l'autre : d'où les libellés courts et la gouttière à 26 px.

**Extraction (`extract.js`).**
- `color-mix()` est calculé par Chromium en `color(srgb r g b / a)`, pas en
  `rgba()` — les deux formes sont gérées, ne pas simplifier.
- `inlineRuns()` fusionne les fragments inline en un seul paragraphe, mais
  **abandonne** si l'élément contient un `<br>` ou un descendant peint (pastille,
  puce) : ces éléments sont exportés comme formes à leur position propre, le texte
  ne peut pas être fusionné avec eux.
- Le texte enrichi est positionné sur la **boîte de contenu** (hors padding),
  sinon il se décale de la valeur du padding.

**Génération PPTX (`build.py`).**
- Appeler `no_shadow()` sur chaque forme : sans cela le thème applique une ombre
  par défaut.
- `custGeom` : le `<a:path>` reçoit ses dimensions en EMU et les points sont
  relatifs à la boîte englobante de la courbe, pas à la planche.
- Renseigner remplissage et contour **avant** l'ombre : `effectLst` doit rester le
  dernier enfant de `spPr`.

**Regex de slug** : la plage de marques combinantes doit rester écrite
`[\u0300-\u036f]`. Avec les caractères combinants littéraux, la regex ne
survit pas à une réécriture du fichier et `slug()` cesse de dépouiller les
accents — les identifiants DOM ne correspondent alors plus aux flux.

## Vérification avant de livrer

1. `python3 data/build_data.py` — vérifier le décompte (7 pôles, 31 missions,
   54 outils, 15 flux) et l'absence d'outil sans mission.
2. Capturer les 3 planches avec Playwright et **les regarder** : débordement de
   carte, chevauchement libellé/pastille, fil qui traverse une carte, erreurs JS
   en console.
3. `python3 pptx/build.py` puis `validate.py` du skill pptx.
4. Convertir le PPTX en PDF (LibreOffice, paquet `libreoffice-impress` requis) et
   inspecter chaque page : c'est là que les défauts de métriques de police
   apparaissent, jamais dans le rendu HTML.

## Décisions à ne pas défaire sans arbitrage

- **Composition du socle commun** : la colonne `Nature` de la BDD est
  contradictoire d'un entretien à l'autre. Le socle retenu = outils
  majoritairement tagués socle + plateformes transverses de la cartographie
  d'origine (Excel, Word, SharePoint, Power BI, E-Paraph, DocuSign).
- **Nature des flux** : la BDD ne qualifie pas l'automatisation flux par flux. Un
  flux est dit automatisé quand son outil de **destination** est alimenté
  automatiquement.
- **Missions proposées** : Topographie, Tunnel et les outils Environnement n'ont
  aucune valeur `FLUX` en BDD. Leurs missions sont déduites de la colonne *Usage*
  et marquées `todo` — ne pas les présenter comme documentées.
- **Répétitions assumées** : un outil apparaît dans plusieurs pôles ou missions
  quand plusieurs l'ont déclaré. C'est voulu, la carte est organisée par pôle.
- `archive/` conserve la première version, bâtie sur le seul PDF initial. Ne pas
  la régénérer, elle n'est plus alimentée par la BDD.
