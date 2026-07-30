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

**Planche 1 — Cartographie.** Le socle applicatif commun au centre, les neuf
services autour. Chaque connecteur relie deux outils, dans le sens réel du flux.

**Planche 2 — Flux et modes d'alimentation.** Les 15 flux en tableau, la
répartition du parc, et les points d'attention.

**Planche 3 — Référentiel.** Les 54 outils avec éditeur, services utilisateurs,
nature et mode d'alimentation.

## Les axes de lecture

Trois dimensions de la BDD sont portées visuellement :

| Dimension | Colonne BDD | Encodage |
|---|---|---|
| Socle commun / application métier | `Nature` | Position : centre ou carte de service |
| Interne / externe | `Interne / Externe` | Pastille teintée (interne) ou neutre (externe) |
| Alimentation | `Alimentation (Manuel/Auto)` | Point cyan sur l'outil ; trait plein (automatisé) ou pointillé (manuel) sur les flux |

Un outil au contour en pointillés est **à qualifier** : il est cité dans la BDD
mais aucun entretien ne documente sa nature ni son alimentation.

## Utilisation

### En HTML

Ouvrir `index.html` dans un navigateur. La page propose un basculement
clair / sombre, une mise en évidence interactive (survoler un service ou un
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

- **Rattacher un outil au socle ou à un service** : listes `SOCLE`, `CORE` et
  `SERVICES` dans `data/build_data.py`. L'ordre des listes du socle commande le
  placement dans la grille à deux colonnes, ce qui évite aux connecteurs de
  traverser le socle.
- **Ajouter un flux** : liste `FLOWS` dans `data/build_data.py`, sous la forme
  `("service:Outil", "s:OutilSocle", "objet du flux")`. Un quatrième élément
  `False` documente le flux dans la matrice sans le tracer sur la carte. Aucune
  coordonnée à saisir : le tracé est calculé depuis la position réelle des
  éléments.
- **Normaliser un libellé d'outil** : dictionnaire `ALIAS`.
- **Changer les couleurs** : variables `--met`, `--top`, `--tun`, `--trv`,
  `--qua`, `--dir`, `--ges`, `--rh`, `--env`, `--core` en tête de la feuille de
  styles de `index.html`.

## Choix et écarts assumés

- **Composition du socle commun.** La colonne `Nature` de la BDD est renseignée
  de façon inégale : un même outil est tantôt « Socle de données », tantôt
  « Application métiers ». Le socle retenu réunit les outils majoritairement
  tagués socle et les plateformes transverses que la cartographie initiale
  plaçait déjà en logiciel commun (Excel, Word, SharePoint, Power BI, E-Paraph,
  DocuSign). À arbitrer avec vous.
- **Regroupement de services.** « Contrat Manager » et « Comptabilité / gestion »
  sont réunis en *Contrats & Gestion* ; « Responsable Travaux » et
  « Ingé travaux » en *Travaux*. Les neuf autres services de la BDD sont conservés
  tels quels.
- **Nature des flux.** La BDD ne qualifie pas l'automatisation flux par flux. Un
  flux est présenté comme automatisé lorsque son outil de **destination** est
  alimenté automatiquement.
- **Rattachement des outils à qualifier.** GMAO, Achat +, Neoaccès et CEMEX sont
  cités sans ligne d'entretien : leur service est proposé d'après le commentaire
  de la BDD, et ils sont marqués comme à qualifier.
- **Un outil apparaît dans plusieurs cartes** lorsque plusieurs services le
  déclarent (AutoCAD, IDCapture, MS Project…). C'est volontaire : la carte est
  organisée par service.
