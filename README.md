# Cartographie des outils IT — chantier

Cartographie du système d'information chantier, construite depuis la **BDD
Cartographie Outils IT** (16 entretiens, 54 outils recensés). Livrée en page HTML
autonome et en PowerPoint **entièrement éditable** (formes natives, pas d'image).

| Livrable | Chemin |
|---|---|
| Planches HTML (3 planches 16:9) | `index.html` |
| Présentation PowerPoint | `dist/Cartographie-Outils-IT.pptx` |
| Base de données source | `data/BDD_Cartographie_Outils_IT.xlsx` |
| Extraction de la BDD | `data/build_data.py` |
| Générateur PowerPoint | `pptx/build.py` + `pptx/extract.js` |
| Version issue du seul PDF initial | `archive/` |

## Les trois planches

**Planche 1 — Cartographie.** Toute la cartographie sur une seule page : les
quatre pôles techniques en haut, le socle applicatif commun en bandeau central,
les trois pôles support et pilotage en bas. Chaque pôle est décomposé en
**missions**, et chaque mission liste les outils qu'elle mobilise. Les connecteurs
relient deux outils dans le sens réel du flux.

Pour que les 31 missions tiennent sur une page, les outils du socle mobilisés par
une mission sont cités en ligne (« via Excel · Power BI ») plutôt qu'en pastilles :
ils figurent déjà dans le bandeau central. Seuls les outils propres au pôle
apparaissent en pastille.

**Planche 2 — Flux et modes d'alimentation.** Les 15 flux en tableau, la
répartition du parc, et les points d'attention.

**Planche 3 — Référentiel.** Les 54 outils avec éditeur, pôles utilisateurs,
nature et mode d'alimentation.

## Les sept pôles

| Pôle | Rangée | Origine |
|---|---|---|
| Méthodes & BIM | haut | service `Méthode/BIM` |
| Topographie | haut | service `TOPO` |
| Tunnel | haut | service `TUNNEL` |
| Travaux | haut | `Responsable Travaux` + `Ingé travaux` |
| Direction | bas | service `Direction` |
| DAF | bas | `Contrat Manager` + `Comptabilité/gestion` + `Assistant RH` |
| Qualité & Environnement | bas | `Qualité` + `Environnement` |

## Les axes de lecture

Trois dimensions de la BDD sont portées visuellement :

| Dimension | Colonne BDD | Encodage |
|---|---|---|
| Socle commun / application métier | `Nature` | Position : centre ou carte de pôle |
| Interne / externe | `Interne / Externe` | Badge `INT` / `EXT` sur chaque pastille — `?` si la BDD ne tranche pas — doublé d'une pastille teintée pour les internes |
| Alimentation | `Alimentation (Manuel/Auto)` | Badge `AUTO` sur l'outil ; sur les flux, trait plein à la couleur du pôle (automatisé) ou **gris pointillé** (manuel) |

Dans les cartes de pôle, les outils du socle mobilisés par une mission sont cités
en ligne après ses pastilles : on lit d'un coup d'œil ce que la mission emprunte au
socle commun.

Un outil au contour en pointillés, ou un libellé de mission en gris, est **à
valider** : il est cité dans la BDD mais aucun entretien ne le documente.

## Utilisation

### En HTML

Ouvrir `index.html` dans un navigateur. La page propose un basculement
clair / sombre, une mise en évidence interactive (survoler un pôle ou un
outil isole ses flux) et un export PDF 16:9 par l'impression.

### Dans PowerPoint

Ouvrir `dist/Cartographie-Outils-IT.pptx`. Tout est modifiable : textes,
couleurs, positions, courbes des connecteurs. Format 16:9, police Calibri.

## Régénérer après mise à jour de la BDD

```bash
pip install python-pptx playwright openpyxl && playwright install chromium

python3 data/build_data.py      # BDD Excel  → modèle injecté dans index.html
python3 pptx/build.py --out dist/Cartographie-Outils-IT.pptx
```

`build_data.py` agrège la feuille « Collecte Entretiens » (un enregistrement par
outil, valeur majoritaire pour les colonnes en conflit), normalise les libellés
saisis librement en entretien, et écrit le modèle entre les marqueurs
`DONNÉES BDD` de `index.html`.

`build.py` ouvre ensuite `index.html` dans Chromium, relève la géométrie réelle
de chaque planche via `extract.js`, et la rejoue en formes natives PowerPoint :
rectangles arrondis, ovales, courbes de Bézier en géométrie personnalisée, zones
de texte. Le HTML est la source unique de vérité.

## Où intervenir

- **Rattacher un outil au socle ou à un pôle** : listes `SOCLE`, `CORE` et
  `SERVICES` dans `data/build_data.py`. L'ordre des listes du socle commande le
  placement dans la grille à deux colonnes, ce qui évite aux connecteurs de
  traverser le socle.
- **Modifier une mission** : dictionnaire `MISSIONS` dans `data/build_data.py`.
  Chaque entrée vaut `(libellé, [valeurs FLUX absorbées], [outils ajoutés à la
  main], proposée)`. Tout outil d'un pôle non rattaché à une mission est signalé
  en sortie de script et regroupé sous « Autres outils ».
- **Répartir les pôles entre les deux rangées** : constante `BOARD` en tête du
  script de `index.html` (`top` et `bottom`). L'ordre commande aussi la longueur
  des connecteurs : un pôle placé sous la zone du socle qu'il alimente évite les
  fils qui traversent la planche.
- **Ajouter un flux** : liste `FLOWS` dans `data/build_data.py`, sous la forme
  `("service:Outil", "s:OutilSocle", "objet du flux")`. Un quatrième élément
  `False` documente le flux dans la matrice sans le tracer sur la carte. Aucune
  coordonnée à saisir : le tracé est calculé depuis la position réelle des
  éléments.
- **Normaliser un libellé d'outil** : dictionnaire `ALIAS`.
- **Changer les couleurs** : variables `--met`, `--top`, `--tun`, `--trv`,
  `--qse`, `--dir`, `--daf`, `--core` en tête de la feuille de styles de
  `index.html`.

## Choix et écarts assumés

- **Composition du socle commun.** La colonne `Nature` de la BDD est renseignée
  de façon inégale : un même outil est tantôt « Socle de données », tantôt
  « Application métiers ». Le socle retenu réunit les outils majoritairement
  tagués socle et les plateformes transverses que la cartographie initiale
  plaçait déjà en logiciel commun (Excel, Word, SharePoint, Power BI, E-Paraph,
  DocuSign). À arbitrer avec vous.
- **Regroupement en pôles.** « Contrat Manager », « Comptabilité / gestion » et
  « Assistant RH » sont réunis en *DAF* ; « Qualité » et « Environnement » en
  *Qualité & Environnement* ; « Responsable Travaux » et « Ingé travaux » en
  *Travaux*.
- **Libellés de missions.** Ils viennent de la note de cadrage, et absorbent les
  valeurs fragmentées de la colonne `FLUX` (« GESTION CHANTIER » et « GESTION DES
  CHANTIER » désignent la même mission). La BDD ne renseigne aucune mission pour
  Topographie, Tunnel et les outils Environnement : celles-ci sont proposées
  d'après la colonne *Usage* et restent à valider.
- **Nature des flux.** La BDD ne qualifie pas l'automatisation flux par flux. Un
  flux est présenté comme automatisé lorsque son outil de **destination** est
  alimenté automatiquement.
- **Rattachement des outils à qualifier.** GMAO, Achat +, Neoaccès et CEMEX sont
  cités sans ligne d'entretien : leur pôle est proposé d'après le commentaire de
  la BDD, et ils sont marqués comme à valider.
- **Un outil apparaît dans plusieurs cartes ou plusieurs missions** lorsque
  plusieurs pôles le déclarent (AutoCAD, IDCapture, MS Project…). C'est
  volontaire : la carte est organisée par pôle puis par mission.
