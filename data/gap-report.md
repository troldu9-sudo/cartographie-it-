# Gap-report — cartographie IT chantier

Généré par `data/extract_bdd.py` le 2026-08-25, depuis `BDD_Cartographie_Outils_IT.xlsx`.
Ce rapport ne signale pas des défauts de génération : il liste ce que la
base d'entretiens ne permet pas de trancher, et ce qui a été ajouté sur
décision métier hors de la base.

> Les numéros de ligne sont ceux de la feuille Excel (l'en-tête est en
> ligne 3, la première ligne de données en ligne 4), et non la colonne `N°`.

## 1. Colonne « socle commun »

La BDD **ne porte pas** de champ booléen `socle_commun`. La seule colonne
qui qualifie la nature d'un outil est **`Nature` (colonne F)** :

| Valeur | Lignes |
|---|---:|
| `Application métiers` | 48 |
| `Socle de données` | 24 |
| `*(vide)*` | 23 |
| `Application métiers/Socle de donnée` | 4 |

Elle est exploitable mais **contradictoire d'un entretien à l'autre**
(section 8). Règle d'arbitrage retenue : *majorité des entretiens ;
égalité → socle ; les 4 outils du hub forcés socle*.

## 2. Outils sans service

Aucune ligne de la table d'entretiens n'est privée de service : les
11 libellés bruts sont tous rattachés à l'un des 9 services.

En revanche, l'onglet `📊 Tableau de Bord` liste **4 outils hors de la
table d'entretiens**, sans service, sans usage et sans flux :

| Ligne | Libellé brut | Note | Retenu sous | Services |
|---:|---|---|---|---|
| 57 | `CEMEX` | — | CEMEX | **aucun** |
| 58 | `GMAO(gestion du stock => tunnel)` | gestion du stock => tunnel | GMAO | Travaux, Travaux tunnel |
| 59 | `Achat +(achat pablo => opérationnel)` | achat pablo => opérationnel | HA+ | Travaux |
| 60 | `Neoacces` | créer les badges => gère les accès | Visioaccès | DAF |

## 3. Flux sans source ou cible

Sur 99 lignes d'entretien, la colonne « Alimente quel outil » donne :

- **39 cellules vides** — aucun flux déclaré ;
- **27** — cellule « Aucun » ;
- **7** — cible générique, aucun outil nommé ;
- **1** — cible non nommée ;
- **1** — désigne une personne, pas un outil ;
- **1** — sens inversé et repris par la correction « processus paye » ;
- **1** — cible non nommée (tableaux internes) ;
- **1** — décrit une fréquence, pas une cible ;
- **1** — formulation ambiguë, sens du flux indéterminé ;
- **1** — cellule « Non » ;
- **1** — cibles inconnues de la BDD (pixys tool, mobidic) ;

Détail des cibles écartées qui ne sont ni `Aucun` ni vides :

| Ligne | Outil source | Cible déclarée | Motif |
|---:|---|---|---|
| 7 | CWT | `Relier à un autre logica bouyguesiel de compt` | cible non nommée |
| 31 | E-Project | `Martin` | désigne une personne, pas un outil |
| 36 | Excel | `Tous` | cible générique, aucun outil nommé |
| 40 | Excel | `Tout` | cible générique, aucun outil nommé |
| 41 | DocuSign | `Tout` | cible générique, aucun outil nommé |
| 43 | E-Paraph | `Tout` | cible générique, aucun outil nommé |
| 45 | Excel | `Tout` | cible générique, aucun outil nommé |
| 46 | QuickConnect | `Tout` | cible générique, aucun outil nommé |
| 49 | Excel | `Tout` | cible générique, aucun outil nommé |
| 52 | PUMA | `BIP(chef de chantier qui deverse dans PUMA)` | sens inversé et repris par la correction « processus paye » |
| 82 | Live Objects (Orange) | `Export manuel des pulsations pour les réinjecter dasn des tab internes` | cible non nommée (tableaux internes) |
| 83 | Sixsense Monitoring | `Ponctuellement en fonction des demandes` | décrit une fréquence, pas une cible |
| 84 | Waste Marketplace (Wakio) | `Avant POWER BI,` | formulation ambiguë, sens du flux indéterminé |
| 94 | Pixis | `pixys tool(rapport de lever d'anneau, position machine) , mobidic` | cibles inconnues de la BDD (pixys tool, mobidic) |

**12 flux retenus depuis la BDD**, 8 ajoutés par les corrections métier, soit 20 au total.

La BDD **ne qualifie jamais l'automatisation d'un flux**. Règle retenue :
un flux est automatisé quand son outil de destination est lui-même
alimenté automatiquement.

## 4. Sens de flux corrigé

Ligne 52 (`N° 49`) — la cellule `BIP(chef de chantier qui deverse dans PUMA)` porte
PUMA comme source alors que le texte décrit **BIP → PUMA**. La correction
métier « processus paye » tranche dans l'autre sens encore : *pointage
PUMA → BIP (intérimaires) / BYCN (compagnons Bouygues)*. **C'est la
correction qui a été appliquée**, la ligne 52 est écartée.

## 5. Processus sans début ni fin

La BDD ne porte **aucune** notion d'étape ordonnée : la colonne `FLUX`
nomme des thèmes, pas des chaînes. Conséquences sur le livrable 2 :

- **34 processus** issus de la BDD n'ont ni début ni fin → bloc vide
  dimensionné et éditable, marqué `TODO` ;
- l'**ordre des étapes** de ces processus est celui des lignes
  d'entretien, il n'est attesté par rien ;
- seuls les processus issus des corrections (*Paye*, *Chaîne
  documentaire des plans*) portent un ordre et des bornes explicites.

## 6. Colonne FLUX vide

**26 lignes** n'ont aucune valeur de FLUX :

| Service | Lignes |
|---|---:|
| TUNNEL | 8 |
| Environnement | 7 |
| TOPO | 7 |
| Direction | 3 |
| Ingé travaux | 1 |

Leurs outils sont regroupés sous un processus « Processus non renseigné »
marqué `TODO`, dont les étapes portent l'usage déclaré en BDD.

## 7. Doublons de nommage

Aucun doublon strict. **30 libellés réécrits**, dont cinq qui créaient
de vrais écarts, pas seulement de la casse :

| Libellé BDD | Retenu | Effet |
|---|---|---|
| `Baseware` | **Basware** | orthographe de l'éditeur |
| `HARMONIE` / `Harmonie` (colonne cible) | **Harmony** | 6 flux pointaient vers un outil inexistant |
| `Cabine de pilotage(Power BI)` | **Power BI** | rapport fusionné dans l'outil |
| `Achat +` | **HA+** | correction métier |
| `Neoacces` | **Visioaccès** | correction métier |

Le reste n'est que de la casse et des espaces (`POWER BI`, `IDCAPTURE`,
`Trimbleconnect`, `Sharepoint`, `Docusign`, `Autocad`, `Puma`, `Pablo`…).

## 8. Contradictions sur la colonne `Nature`

**6 outils** sont qualifiés différemment selon l'entretien :

| Outil | Votes | Retenu |
|---|---|---|
| Basware | application ×3 · socle ×2 | **application** |
| E-Checking | application ×1 · socle ×1 | **socle** |
| Excel | application ×7 · socle ×3 | **application** |
| IDCapture | application ×2 · socle ×1 | **application** |
| PABLO | application ×2 · socle ×3 | **socle** |
| QuickConnect | application ×3 · socle ×4 | **socle** |

## 9. Contradictions sur `Interne / Externe`

Après arbitrage : **17 internes · 30 externes · 14 `TODO`**.
Les `TODO` forment une troisième part dans le graphique de la synthèse ;
aucune valeur n'est devinée.

| Outil | Votes | Retenu |
|---|---|---|
| Covadis | *aucun* | **TODO** |
| Cyclone 3DR | *aucun* | **TODO** |
| E-Checking | externe ×1 · interne ×1 | **TODO** |
| E-Project | *aucun* | **TODO** |
| Excel | externe ×8 · interne ×1 | **externe** |
| IDCapture | externe ×3 · interne ×1 | **externe** |
| La Scene | *aucun* | **TODO** |
| MS Project | externe ×1 · interne ×1 | **TODO** |
| Miro | *aucun* | **TODO** |
| PABLO | externe ×3 · interne ×1 | **externe** |
| Pixis | *aucun* | **TODO** |
| Power BI | externe ×2 · interne ×2 | **TODO** |
| QuickConnect | externe ×1 · interne ×4 | **interne** |

## 10. Colonnes non exploitées

- **`Interfaces / Connexions` (O)** est polluée : des niveaux de maturité
  (`3 – Mature`, `4 – Optimisé`) et 10 valeurs `Fluide` y ont été saisis à
  la place de la colonne P. Inexploitable.
- **`Niveau Maturité` (P)** et **`Fréquence d'utilisation` (L)** sont
  écartées sur consigne : ni extraites, ni affichées nulle part.
- **`Alimenté par qui` (M)** nomme des personnes et des équipes, pas des
  outils : elle ne peut pas servir à tracer des flux entrants.
- **`Colonne32` (R)** est renseignée sur 1 ligne sur 99.

## 11. Éléments ajoutés hors BDD

Chacun porte un champ `origine` dans `carto.json` et un badge dans le
référentiel. `brief` = imposé par le cahier des charges ; `correction` =
section CORRECTIONS MÉTIER ; `bdd-annexe` = onglet Tableau de Bord.

| Outil | Origine | Services | TODO |
|---|---|---|---|
| OneDrive | `brief` | — | absent de la BDD — ajouté au titre du hub de données |
| Probity | `brief` | — | absent de la BDD — ajouté au titre du hub de données |
| Chorus | `correction` | Contrat | — |
| Sogelink/DICT | `correction` | Qualité & Environnement | — |
| RNDTS | `correction` | Qualité & Environnement | valider avec Camille |
| Remind | `correction` | Méthode & BIM | — |
| BYCN | `correction` | DAF | entité ou applicatif ? cible du pointage à préciser |
| GMAO | `bdd-annexe` | Travaux, Travaux tunnel | — |
| Visioaccès | `bdd-annexe` | DAF | — |
| HA+ | `bdd-annexe` | Travaux | service déduit de la note « achat pablo ⇒ opérationnel », à confirmer |
| CEMEX | `bdd-annexe` | — | aucun service ni usage renseigné dans la BDD |

Flux ajoutés ou requalifiés par les corrections : **10**.

| De | Vers | Mode | Objet |
|---|---|---|---|
| Basware | Harmony | auto | Retour de facturation |
| Harmony | Basware | auto | Engagement des dépenses |
| PABLO | GMAO | manuel | Commande vers gestion de stock |
| GMAO | PABLO | manuel | Réapprovisionnement |
| PUMA | BIP/BAPS | manuel | Pointage des intérimaires |
| PUMA | BYCN | manuel | Pointage des compagnons Bouygues |
| AutoCAD | SharePoint | manuel | Dépôt des plans produits |
| SharePoint | Mezzoteam | manuel | Diffusion au client |
| Mezzoteam | Remind | manuel | Archivage 10 ans |
| Chorus | Basware | manuel | Facturation du maître d'ouvrage |

## 12. Reste à faire (`TODO` visibles dans les livrables)

- **33 `TODO` d'outils** (fonction, interne/externe, alimentation, service) ;
- **34 processus sans début ni fin** ;
- le service **Sécurité** est volontairement vide : 3 blocs processus
  vierges, éditables, aucun outil prélevé aux autres services ;
- **RNDTS** : traçabilité des déblais — *valider avec Camille* ;
- **BYCN** : entité ou applicatif ? cible du pointage à préciser ;
- **CEMEX** : aucun service ni usage renseigné.

