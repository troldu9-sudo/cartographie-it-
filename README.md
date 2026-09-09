# Cartographie des outils IT — chantier

Cartographie du système d'information chantier, construite depuis la **BDD
Cartographie Outils IT** (16 entretiens) et complétée par le brief métier.
64 outils, 9 pôles, 37 missions. Livrée en page HTML autonome et en PowerPoint
**entièrement éditable** (formes natives, pas d'image).

| Livrable | Chemin |
|---|---|
| Planches HTML (13 planches 16:9) | `index.html` |
| Présentation PowerPoint | `dist/Cartographie-Outils-IT.pptx` |
| Base de données source | `data/BDD_Cartographie_Outils_IT.xlsx` |
| Extraction de la BDD | `data/build_data.py` |
| Générateur PowerPoint | `pptx/build.py` + `pptx/extract.js` |
| Version issue du seul PDF initial | `archive/` |

## Les treize planches

**Planche 1 — Cartographie applicative.** Les neuf pôles et leurs outils autour du
socle applicatif commun : cinq pôles techniques en haut, le socle en bandeau
central, quatre pôles support et pilotage en bas. Les connecteurs relient deux
outils dans le sens réel du flux.

**Planche 2 — Flux inter-services.** Le même chantier vu à la maille mission : ce
qu'un pôle produit et ce qu'un autre en fait. Les nœuds ne sont plus les outils
mais les missions, et chaque fil reprend un lien saisi dans les chaînes.

**Planches 3 à 11 — Une affiche par pôle.** Six pôles (DAF, Contrat, Travaux,
Travaux tunnel, Topographie, Sécurité) présentent une **chaîne de flux par
mission** : d'où vient la donnée, par quels outils elle passe, à quoi elle aboutit.
Un badge coloré marque l'étape où la mission alimente un autre pôle — ou en est
alimentée.

Les trois pôles dont le processus n'a pas été recueilli (Direction, Méthodes & BIM,
Qualité & Environnement) gardent l'affiche « inventaire » : les outils de chaque
mission entre deux blocs « à documenter ». La distinction est visible d'un coup
d'œil, et c'est voulu : rien ne laisse croire qu'un processus est attesté quand il
ne l'est pas.

**Planche 12 — Matrice des flux.** Deux tableaux : les 14 échanges d'outil à outil
relevés en entretien, et les 18 échanges de mission à mission décrits par le brief.
Plus la répartition du parc et les points d'attention.

**Planche 13 — Référentiel.** Les 64 outils avec éditeur, pôles utilisateurs,
nature et mode d'alimentation. Les outils venus du brief et absents de la base
d'entretiens sont signalés par la couleur de leur nom.

## Les neuf pôles

| Pôle | Affiche | Origine |
|---|---|---|
| Méthodes & BIM | inventaire | service `Méthode/BIM` |
| Topographie | chaînes | service `TOPO` |
| Travaux tunnel | chaînes | service `TUNNEL` |
| Travaux | chaînes | `Responsable Travaux` + `Ingé travaux` |
| Contrat | chaînes | service `Contrat Manager` |
| Direction | inventaire | service `Direction` |
| DAF | chaînes | `Comptabilité/gestion` + `Assistant RH` |
| Qualité & Environnement | inventaire | `Qualité` + `Environnement` |
| Sécurité | chaînes | brief métier — aucune ligne d'entretien |

## Les axes de lecture

Trois dimensions de la BDD sont portées visuellement :

| Dimension | Colonne BDD | Encodage |
|---|---|---|
| Socle commun / application métier | `Nature` | Position : centre ou carte de pôle |
| Interne / externe | `Interne / Externe` | Pastille teintée (interne) ou neutre (externe) |
| Alimentation | `Alimentation (Manuel/Auto)` | Point cyan sur l'outil ; trait plein (automatisé) ou pointillé (manuel) sur les flux |

Un outil au contour en pointillés est **hors base d'entretiens** : il vient du
brief métier, et sa nature comme son alimentation restent à confirmer. Un libellé
de mission en gris, accompagné de « chaîne à valider », est proposé et non attesté.

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

- **Modifier une chaîne de flux** : dictionnaire `PROCESS` dans
  `data/build_data.py`. Chaque mission porte ses sources (`in`), ses étapes
  ordonnées (`steps`), sa finalité (`out`) et ses interconnexions (`links`). Un
  `links` produit **à la fois** le badge sur l'affiche et le fil sur la planche
  inter-services : une seule saisie, deux rendus.
- **Un pôle sans processus recueilli** : dictionnaire `MISSIONS`, qui ne sert plus
  qu'aux pôles absents de `PROCESS`. Chaque entrée vaut `(libellé, [valeurs FLUX
  absorbées], [outils ajoutés à la main], proposée)`.
- **Ajouter un outil venu du brief** : dictionnaire `HORS_BDD`. Il sera marqué
  comme tel sur les planches et dans le référentiel.
- **Rattacher un outil au socle ou à un pôle** : listes `SOCLE`, `CORE` et
  `SERVICES`. Un pôle décrit par `PROCESS` ne retient que les outils cités dans ses
  chaînes — inutile de l'y déclarer ailleurs.
- **Répartir les cartes** : `BOARD_TOP` et `BOARD_BOTTOM` pour la cartographie,
  `XBOARD_ROWS` pour la planche inter-services. Cette dernière n'est pas
  cosmétique : un fil ne reste lisible qu'entre deux cartes voisines, et le script
  signale toute paire trop éloignée.
- **Ajouter un flux applicatif** : liste `FLOWS`, sous la forme
  `("service:Outil", "s:OutilSocle", "objet du flux")`. Un quatrième élément
  `False` documente le flux dans la matrice sans le tracer. Aucune coordonnée à
  saisir : le tracé est calculé depuis la position réelle des éléments.
- **Normaliser un libellé d'outil** : dictionnaire `ALIAS`.
- **Changer les couleurs** : variables `--met`, `--top`, `--tun`, `--trv`, `--ctr`,
  `--qse`, `--dir`, `--daf`, `--sec`, `--core` en tête de la feuille de styles de
  `index.html`.

## Choix et écarts assumés

- **Composition du socle commun.** La colonne `Nature` de la BDD est renseignée
  de façon inégale : un même outil est tantôt « Socle de données », tantôt
  « Application métiers ». Le socle retenu réunit les outils majoritairement
  tagués socle et les plateformes transverses que la cartographie initiale
  plaçait déjà en logiciel commun (Excel, Word, SharePoint, Power BI, E-Paraph,
  DocuSign). À arbitrer avec vous.
- **Regroupement en pôles.** « Comptabilité / gestion » et « Assistant RH » sont
  réunis en *DAF* ; « Qualité » et « Environnement » en *Qualité & Environnement* ;
  « Responsable Travaux » et « Ingé travaux » en *Travaux*. « Contrat Manager »
  était rattaché à la DAF : il devient un pôle autonome, la chaîne de facturation
  partenaire circulant désormais explicitement entre les deux.
- **L'ordre des étapes vient du brief, pas de la BDD.** La base d'entretiens ne
  documente aucun enchaînement : les chaînes des six pôles concernés sont saisies
  à la main depuis le brief métier, et les outils qu'elles citent font foi. C'est
  ce qui retire AutoCAD de la Topographie et rattache Pablo au tunnel.
- **Dix outils viennent du brief et non de la BDD** : Outlook, Chorus, appli
  bancaire, BYCN, BYMAT, Cority, Quick Connect Sécurité, Lotus, Heures Travaillées,
  HRMYOU / Global HR. Ils sont signalés comme tels partout où ils apparaissent.
- **Libellés de missions.** Ils viennent de la note de cadrage, et absorbent les
  valeurs fragmentées de la colonne `FLUX` (« GESTION CHANTIER » et « GESTION DES
  CHANTIER » désignent la même mission). La BDD ne renseigne aucune mission pour
  Topographie, Tunnel et les outils Environnement : celles-ci sont proposées
  d'après la colonne *Usage* et restent à valider. Les missions du tunnel sont
  décalquées de celles des Travaux et marquées comme telles.
- **Nature des flux.** La BDD ne qualifie pas l'automatisation flux par flux. Un
  flux est présenté comme automatisé lorsque son outil de **destination** est
  alimenté automatiquement.
- **Rattachement des outils à qualifier.** GMAO, Achat +, Neoaccès et CEMEX sont
  cités sans ligne d'entretien : leur pôle est proposé d'après le commentaire de
  la BDD, et ils sont marqués comme à valider.
- **Un outil apparaît dans plusieurs cartes ou plusieurs missions** lorsque
  plusieurs pôles le déclarent (AutoCAD, IDCapture, MS Project…). C'est
  volontaire : la carte est organisée par pôle puis par mission.
