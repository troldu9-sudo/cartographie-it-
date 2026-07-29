# Cartographie des flux de données SI

Refonte graphique de la cartographie « Flux de données » : une page HTML autonome
et un fichier PowerPoint **entièrement éditable** (formes natives, pas d'image).

| Livrable | Chemin |
|---|---|
| Planches HTML (2 planches 16:9) | `index.html` |
| Présentation PowerPoint | `dist/Cartographie-flux-de-donnees.pptx` |
| Générateur PowerPoint | `pptx/build.py` + `pptx/extract.js` |

## Contenu

**Planche 1 — Cartographie.** Les quatre directions (Technique & Méthodes, DAF,
Qualité & Environnement, Travaux) entourent le socle applicatif commun. Chaque
connecteur relie un processus métier à l'application du socle qu'il alimente ou
qui l'alimente, avec le sens réel du flux.

**Planche 2 — Matrice des flux.** Les 11 flux inter-applicatifs en tableau, les
traitements internes aux directions, et les points de lecture de la cartographie.

## Utilisation

### En HTML

Ouvrir `index.html` dans un navigateur. La page propose :

- un **basculement clair / sombre** ;
- une **mise en évidence interactive** : survoler une direction ou une application
  du socle isole ses flux ;
- un bouton **Exporter en PDF 16:9** (impression navigateur, une planche par page).

### Dans PowerPoint

Ouvrir `dist/Cartographie-flux-de-donnees.pptx`. Tout est modifiable dans
PowerPoint : textes, couleurs, positions, courbes des connecteurs. Le format est
16:9 (33,87 × 19,05 cm) ; police Calibri.

## Régénérer le PowerPoint après modification du HTML

Le HTML est la source unique de vérité. `pptx/build.py` ouvre `index.html` dans
Chromium (Playwright), relève la géométrie réelle de chaque planche via
`pptx/extract.js`, puis la rejoue en formes natives PowerPoint : rectangles
arrondis, ovales, courbes de Bézier en géométrie personnalisée, zones de texte.

```bash
pip install python-pptx playwright && playwright install chromium
python3 pptx/build.py
```

Options : `--html`, `--out`, `--chromium` (chemin d'un binaire Chromium existant).

## Modifier la cartographie

- **Ajouter ou retirer un flux** : la liste `FLOWS`, dans le script de `index.html`,
  alimente à la fois les connecteurs de la planche 1 et le tableau de la planche 2.
  Les extrémités sont désignées par l'`id` de l'élément (`daf-budget`, `app-excel`…).
  Les connecteurs sont tracés à partir de la position réelle des éléments : aucune
  coordonnée n'est à saisir.
- **Changer les couleurs** : variables `--dt`, `--daf`, `--qe`, `--tv`, `--core`
  en tête de la feuille de styles.
- **Renommer une direction ou un processus** : directement dans le HTML de la
  planche 1 ; le tableau de la planche 2 se met à jour depuis `FLOWS`.

## Écarts assumés avec le document source

- Ordre des lignes réagencé dans les cartes DAF et Qualité & Environnement pour
  supprimer les croisements de connecteurs — les flux sont inchangés.
- Corrections rédactionnelles : « levées de réserve », « gestion des contractuels »,
  « gestion des tiers intervenants ».
- La planche 2 signale que la direction Travaux n'a aucun flux vers le socle
  commun et que Power BI n'est raccordé à aucun flux : ce sont des constats de
  lecture du document source, pas des données ajoutées.
