# Cartographie des outils IT — contexte projet

Cartographie du SI d'un chantier Bouygues Construction, produite depuis une base
d'entretiens Excel. Deux livrables issus d'une source unique : une page HTML
autonome et un PowerPoint **en formes natives éditables** (jamais une image).

Tout le contenu visible est en **français** ; les commentaires du code aussi.

## Règles non négociables

Ces six règles définissent ce que la cartographie doit dire. Une modification qui
en casse une est un défaut, quelle que soit sa qualité par ailleurs.

1. **Un pôle, une carte, une seule planche.** Les huit pôles tiennent sur la
   planche 1. Un pôle n'est jamais scindé entre deux planches, jamais dupliqué,
   jamais renvoyé en annexe. Si le contenu déborde, on comprime — voir les leviers
   plus bas — on ne coupe pas.
2. **Chaque mission est dans la carte de son pôle**, avec les outils qu'elle
   mobilise : pastilles pour les outils propres au pôle, citation en ligne
   (« via Excel · Power BI ») pour ceux du socle.
3. **Tout flux tracé dit son mode.** Automatisé : trait plein, couleur du pôle.
   Manuel : trait **gris pointillé**. Les deux figurent en légende de la planche.
4. **Toute pastille dit la nature de son outil** : badge `INT`, `EXT`, ou `?`
   quand la BDD ne tranche pas — cartes de pôle **et** bandeau socle.
5. **Un outil alimenté automatiquement porte le badge `AUTO`**.
6. **Chaque carte porte un bandeau de titre à la couleur de son pôle.** C'est lui
   qui rend la frontière entre services indiscutable, la teinte seule n'y suffit
   pas : Méthodes & BIM, Direction et DAF sont à ΔE 13 à 24 les unes des autres.

Structure de référence, reprise du tableau blanc et des notes d'entretien :
**conteneur de pôle → mission → pastille d'outil**, et les flux se tracent d'outil
à outil, au travers des conteneurs — jamais de pôle à pôle.

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
| `services[]` | les 8 pôles : `{id, name, kicker, missions[]}` |
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

`ie` vaut `interne`, `externe` ou `?` ; `alim` vaut `auto`, `mixte`, `manuel` ou
`?`. Les badges de la carte sont dérivés de ces deux champs et de rien d'autre :
`IE_TAG` et `isAuto()` en tête du script de `index.html`.

## Où intervenir

| Besoin | Endroit |
|---|---|
| Missions, libellés, rattachements | `MISSIONS` dans `data/build_data.py` |
| Mission qui puise dans un autre pôle | valeur `"dir:CIRCUIT DE VALIDATION"` dans `MISSIONS` + `ABSORBED` |
| Composition des pôles | `SERVICE` + `SERVICES` |
| Composition du socle | `SOCLE` + `CORE` |
| Flux | `FLOWS` |
| Libellés d'outils saisis en vrac | `ALIAS` |
| Répartition haut/bas, cartes à deux colonnes | `BOARD` (`top`, `bottom`, `wide`) en tête du script de `index.html` |
| Hauteur des deux rangées | `.prow.tech` / `.prow.support` |
| Badges de nature et d'alimentation | `IE_TAG`, `ieRun()`, `alRun()` |
| Bandeau de titre des cartes | `.card-band` |
| Couleur et style des flux | `wireColor()` + `.wire.manuel` |
| Couleurs | `--met --top --tun --trv --qse --dir --daf --ctr --core` en tête du CSS |

**Missions transverses.** Une valeur `FLUX` nue est lue dans les lignes du pôle
lui-même ; préfixée — `"dir:CIRCUIT DE VALIDATION"` — elle est lue chez un autre
pôle. C'est ce qui permet à Contrat Manager de rassembler le contractuel déclaré
par la Direction et les Travaux. Le garde-fou `ABSORBED` compte ces outils comme
placés dans leur pôle d'origine : sans lui ils retomberaient dans la mission
fourre-tout « Autres outils » de Direction (`Basware, DocuSign, E-Paraph`) et de
Travaux (`E-Paraph, Word`).

## Mise en page de la planche 1

Format natif 1600 × 900 px = 13,333 × 7,5 po (1 px = 0,6 pt = 7620 EMU).

Quatre pôles techniques en haut, socle en **bandeau horizontal** au centre, quatre
pôles support en bas — huit cartes de 361 px. Ce n'est pas un choix esthétique :
32 missions, 45 pastilles d'outils métier et 18 logiciels de socle ne tiennent pas
dans une mise en page à socle vertical et deux colonnes. Trois leviers rendent la
page unique possible, à préserver :

- les outils du socle mobilisés par une mission sont **cités en ligne**
  (« via Excel · Power BI ») au lieu d'être répétés en pastilles ;
- les deux rangées n'ont pas la même hauteur : la rangée technique porte 2 à 4
  missions par carte, la rangée support 4 à 6, d'où `flex:85` contre `flex:100`
  (`.prow.tech` / `.prow.support`), soit 266 px et 313 px ;
- la carte la plus dense affiche ses missions **sur deux colonnes**
  (`.card.wide`, liste `BOARD.wide`). Aujourd'hui Qualité & Environnement seule :
  6 missions dont 3 avec ligne « via » font 359 px de contenu pour 318 disponibles.

**Si le contenu déborde, ne pas scinder la planche** (règle 1). Les leviers, dans
cet ordre :

1. **raccourcir les libellés d'outils** dans `ALIAS` — restent longs
   `Wastemarket Place` (17 car.), `Sofistik Bridge` (15). Attention, ce levier ne
   gagne quelque chose que s'il **supprime une ligne de pastilles** : abréger
   `Live Objects (Orange)` et `Sixense Monitoring` n'a rien changé au débordement
   de Qualité & Environnement, dont les pastilles tenaient déjà sur deux lignes ;
2. **ajouter la carte à `BOARD.wide`** — c'est ce qui a réglé les 41 px de
   Qualité & Environnement ; à 361 px les colonnes tombent à 156 px, vérifier que
   les libellés de mission n'y bavent pas (`mission.scrollWidth > clientWidth`) ;
3. rééquilibrer les deux rangées via `.prow.tech` / `.prow.support` — 3 px
   suffisaient à Tunnel, `flex:81` est passé à `flex:85` ;
4. resserrer les pastilles (`padding`, `gap`, corps de 10,8 px).

Un nom canonique est **aussi une extrémité de flux** : le renommer oblige à
reprendre `FLOWS`, les listes `extra` de `MISSIONS` et `A_QUALIFIER`. Le
signalement « ATTENTION — flux vers un outil inconnu » attrape un oubli.

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

**Couleur et style portent le mode d'alimentation** (règle 3) : `wireColor()`
renvoie `--muted` dès que le flux n'est pas automatisé, et `.wire.manuel` ajoute le
pointillé. Sur les 14 flux tracés, 8 sont manuels — donc gris — et 6 automatisés,
à la couleur de leur pôle. La pastille d'origine (`circle.port`) reprend la même
couleur, ne pas les désynchroniser.

## Pièges vérifiés — ne pas les redécouvrir

**Métriques de police.** Chromium compose ici en DejaVu Sans, PowerPoint en
Calibri. Tout texte exporté peut être plus large que mesuré.
- Centrer le texte des pastilles (`text-align:center`) : l'écart devient
  symétrique au lieu de déborder à droite.
- Ne jamais poser deux fragments de texte côte à côte en comptant sur la largeur
  mesurée : ils se chevauchent. Soit un seul paragraphe multi-runs, soit une
  gouttière large (la légende tient par sa gouttière flex de 8 px).
- Les libellés de missions sur deux colonnes ont déjà débordé d'une colonne à
  l'autre : d'où les libellés courts et la gouttière à 26 px.
- Contrôle sans PowerPoint : installer `fonts-crosextra-carlito` et rejouer la
  page avec `--font:Carlito`. Carlito a les métriques de Calibri, c'est le seul
  moyen fiable de mesurer le rendu final depuis Chromium.

**Badges de pastille.** `INT` / `EXT` / `?` / `AUTO` sont des **runs de texte non
peints** (`.tool em.ie`, `.tool em.al`), et c'est structurel :
- un marqueur peint — point coloré, badge à fond ou à bordure — fait renoncer
  `inlineRuns()`, et le marqueur repart alors en zone de texte séparée, posée sur
  le nom de l'outil ;
- pour la même raison, `.tool` et `.app` **ne sont pas des conteneurs flex** : un
  conteneur flex « blockifie » ses enfants, `inlineRuns()` exige des enfants
  `display:inline`. C'est `inline-block` + `text-align:center`.
- les marges (`margin-right:4px`) ne sont pas décoratives : `inlineRuns()` en
  déduit les espaces entre runs. Les supprimer colle `INT` au nom à l'export.
- le socle de données (`.core-app`) porte sa nature dans le sous-titre existant,
  pas dans un élément de plus, sinon le bloc passe à trois lignes et le bandeau
  grandit.

**Bandeau de titre.** `.card-band` n'a **pas** de `border-radius` : `rounded_adj()`
plafonne l'ajustement d'un rectangle arrondi à 0,5, et un bandeau de 5 px arrondi
sortirait en pilule dans PowerPoint. Ce sont les coins de la carte
(`.card{overflow:hidden}`) qui le rognent en HTML ; à l'export il reste un
rectangle franc au ras du bord, l'écart se limite à un filet aux extrémités.

**Extraction (`extract.js`).**
- `color-mix()` est calculé par Chromium en `color(srgb r g b / a)`, pas en
  `rgba()` — les deux formes sont gérées, ne pas simplifier.
- `inlineRuns()` fusionne les fragments inline en un seul paragraphe, mais
  **abandonne** si l'élément contient un `<br>` ou un descendant peint (pastille,
  puce) : ces éléments sont exportés comme formes à leur position propre, le texte
  ne peut pas être fusionné avec eux.
- Le texte enrichi est positionné sur la **boîte de contenu** (hors padding),
  sinon il se décale de la valeur du padding.
- Le pointillé se détecte sur `strokeDasharray`, **jamais sur un nom de classe** :
  la version précédente testait une classe `dashed` que `index.html` n'écrivait
  pas, et les huit flux manuels sortaient en trait plein sous une légende qui
  annonçait le contraire. Le CSS est la source unique.
- La couleur des flèches de légende est relevée sur le `<svg>` lui-même : `color`
  étant héritée, une flèche sans couleur propre rend comme avant, et une flèche
  teintée reste enfant direct de `.item`. L'en sortir casserait le garde-fou
  « svg » d'`inlineRuns()` qui protège le texte de la légende.

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

**Binaire Chromium.** Un `pip install playwright` récent réclame une révision de
navigateur que l'environnement n'a pas forcément (`Executable doesn't exist at
…chromium_headless_shell-1234`). Ne pas relancer `playwright install` à l'aveugle :
passer le binaire présent via `--chromium`.

## Vérification avant de livrer

1. `python3 data/build_data.py` — attendu : `54 outils · 6 groupes socle · 8 pôles
   · 32 missions · 15 flux · 16 internes / 28 externes · 8 alimentés en
   automatique`, et aucun « outils sans mission » — c'est le test du garde-fou
   `ABSORBED`. Vérifier aussi qu'aucune mission n'est vide : retirer une valeur
   `FLUX` d'un pôle sans retirer la mission laisse un libellé sans outil.
2. Capturer les 3 planches avec Playwright et **les regarder** : les 8 bandeaux
   colorés, débordement de carte, badge collé au nom, chevauchement
   libellé/pastille, fil qui traverse une carte, erreurs JS en console. Mesurer
   aussi en Carlito (voir les pièges) : `scrollHeight > clientHeight` sur `.card`
   et `.hub-band`, `scrollWidth > clientWidth` sur les `.mission` des cartes à deux
   colonnes, et texte de pastille plus large que sa boîte de contenu — les trois à
   zéro dans les deux polices.
3. Compter les flux dans le DOM : 14 tracés, dont **8 en pointillé** gris
   `rgb(120, 135, 159)`.
4. `python3 pptx/build.py` puis `validate.py` du skill pptx. Attendu sur
   `ppt/slides/slide1.xml` : 16 `custGeom` (14 flux + 2 flèches de légende) dont
   **9 `prstDash`** — les 8 flux manuels plus la flèche manuelle de la légende —
   tous en `78879F`. Un seul `prstDash` signifie que la détection du pointillé est
   de nouveau cassée.
5. Vérifier que les pastilles sortent en paragraphe unique : 61 paragraphes
   multi-runs commençant par `INT`, `EXT` ou `?`. Un badge exporté en zone de
   texte isolée (hors les 3 items de légende) veut dire qu'un élément peint s'est
   glissé dans la pastille. Contrôler aussi les **8 bandeaux de titre** : un
   rectangle de 366 × 5 px par pôle, à la couleur du pôle.
6. Convertir le PPTX en PDF (LibreOffice, paquet `libreoffice-impress`) et
   inspecter chaque page : c'est là que les défauts de métriques de police
   apparaissent, jamais dans le rendu HTML.

## Décisions à ne pas défaire sans arbitrage

- **Une seule planche de cartographie.** Le repli « scinder en deux planches
  (techniques / support) » a été explicitement écarté : il violerait la règle 1.
  Ne pas le réintroduire au premier débordement, comprimer.
- **Les flux manuels perdent la couleur de leur pôle.** Huit des quatorze flux
  tracés passent donc en gris : c'est un choix assumé, la ressaisie est le point
  de fragilité qu'on veut voir en premier.
- **Badge sur toutes les pastilles**, y compris les ~2/3 d'outils externes, et
  malgré le coût en largeur. La lecture ne doit pas dépendre de la comparaison de
  deux teintes voisines. Les citations en ligne « via … » restent en texte nu :
  l'outil est déjà badgé dans le bandeau socle.
- **Contrat Manager est un pôle transverse.** Il rassemble ses 8 lignes
  d'entretien *et* le contractuel que la Direction et les Travaux déclaraient
  chacun de leur côté : Direction a cédé « Validation contrats et dépenses »,
  Travaux « Validation et contrats ». Ne pas leur rendre ces missions sans
  arbitrage, la carte Contrat Manager perdrait sa raison d'être.
- **Une seule teinte ajoutée, pas de palette refaite.** `--ctr:#5F7A1B` (olive)
  est la teinte libre la plus éloignée des huit autres — ΔE minimal 41,8, contre
  36,7 pour la moutarde et 19,0 pour le bordeaux. Méthodes & BIM, Direction et DAF
  restent proches entre elles (ΔE 13 à 24) : c'est le bandeau de titre, pas la
  teinte, qui porte la distinction entre cartes (règle 6).
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
