# Templates des livrables ECOSLI

Inventaire des templates utilisés par les générateurs (`B_ATM_*`, `C_EMMST_*`, `D_Providor_*`).
Les fichiers de référence sont stockés dans le dossier `Templates/` du repo (non distribués avec le `.xlam` — les chemins de production pointent vers SharePoint).

Chemins déclarés dans `Z_Variables.bas` (constantes `TEMPLATE_*_PATH`).

---

## ATM V1 — un seul fichier multi-feuilles par projet

Un seul classeur de sortie qui contient un Sommaire + N feuilles ATM (une par DTR, dupliquées depuis les templates `ATM_HARD` / `ATM_SOFT`).

| Langue | Fichier template | Constante | Feuilles utilisées par le code |
|---|---|---|---|
| FR | `ATM_FR-Template VBN v2.xlsx` | `TEMPLATE_V1_FR_PATH` | `Couverture`, `Sommaire`, `ATM_HARD`, `ATM_SOFT` |
| EN | `MTA_EN-Template VBN v2.xlsx` | `TEMPLATE_V1_EN_PATH` | `CoverPage`, `Index`, `MTA_HARD`, `MTA_SOFT` |

Feuilles annexes (référencées par formules ou listes déroulantes, non manipulées par VBA) : `Liste Choix` / `Choices list`, `Acteurs` / `Actors`, `ATMC_Design`, `ATMC_Diagnostic`, `ATMC_Illustrations`, `ATMC_Refit`, `ATMC_Remove`, `ATMC_Security_measures`, `ATMC_Test_control`, `MTTR_Correctif`, `versionning` / `Versionning`.

Côté code : `B_ATM_V1.bas → ATM_V1()`, `CopierTemplateV1`, `RemplirIndexV1`, `GenererFicheATM_V1`, `FinaliserFichierV1`.

---

## ATM V2 (ATMC 2.0) — un fichier par DTR

Un fichier Excel généré par DTR (template recopié pour chaque DTR), puis remplissage dynamique de la feuille `DESCRIPTION_TACHE` / `TASK_DESCRIPTION` par insertion de lignes dans les chapitres A → E.

| Langue | Fichier template | Constante | Feuille principale | Feuille couverture |
|---|---|---|---|---|
| FR | `ATMC2.0_FR-Template.xlsx` | `TEMPLATE_V2_FR_PATH` | `DESCRIPTION_TACHE` | `Page de garde ATM` |
| EN | `ATMC2.0_EN-Template.xlsx` | `TEMPLATE_V2_EN_PATH` | `TASK_DESCRIPTION` | `CoverPage ATM` |

Feuilles annexes (FR) : `Traçabilité`, `Règle`, `Recapitulatif`, `Illustration`, `DESCRIPTION_TACHE (exemple)`, `Illustration (exemple)`, `Liste de choix`, plus une cinquantaine de feuilles `_Cat*`, `Liste_*`, `Indice_*`, `Date_MON_*` héritées du référentiel Alstom (calculs LCC / paramètres usine).
Feuilles annexes (EN) : `History`, `Link to other sheet`, `Summary`, `Picture`, `TASK_DESCRIPTION (example)`, `Picture 1 (example)`, `Liste_Compé`, `Liste_Conso`, `Liste_Figur`, `Liste_Outil`, `Liste_Recha`, `Liste_RelatTtask`, `Liste_TâcheIndui`, `Liste_Visse`, `ListePicto`, `Verbe_action`.

### Structure de `DESCRIPTION_TACHE` / `TASK_DESCRIPTION`

Colonne `B` = marqueur de chapitre / rang de tâche. Le code repère 4 chapitres par `Find` sur colonne B (cf. `B_ATM_V2.InitialiserChapitres`) :

| Marqueur | Contenu | Cumul de temps écrit en `L<ligneChap>` |
|---|---|---|
| `A` | Tâches rang 1 | Oui |
| `B` | Tâches rang 2 | Oui |
| `C` | Tâches rang 3 (Dépose/Remove) puis rang 4 (Pose/Refit) avec sous-chapitres B1, B2…, ligne d'en-tête « SOUS-CHAPITRE » / « SUB-CHAPTER » | Oui |
| `D` | Tâches rang 5 | Oui |
| `E` | Ajouté en fin de fiche par `AjouterChapitreE` — « Note : À définir par intégrateur » / « Note : To be defined by integrator » | Non |

Colonnes remplies par tâche : `B` rang court, `C` catégorie/premier mot, `D` description, `G` opérateur (Chap C), `I` pièce de rechange (Chap C), `L` temps estimé (`0.000`).

Couverture (`RemplirCouverture`) : `F26` = Projet, `F28` = DTR, `F30` = Descriptif FR/EN.

### Bibliothèque illustrée

Avant génération, le code teste pour chaque DTR la présence d'un fichier déjà illustré dans `BIBLIO_PATH` (`Bibliothèques ATM2.0 illustrées` côté SharePoint). Si présent → copie directe sans génération. Sinon → génération depuis template + copie locale en cache (`templateLocalV2` dans `%TEMP%`).

Naming des sorties :
- DTR illustré présent en bibliothèque : `<DTR>_<langue>_*.xlsx` (recopie directe)
- DTR connu en BDD, non illustré : `(non illustre)<DTR>_<langue>_<10chars Designation>.xlsx`
- DTR connu en BDD mais pas dans la langue : `(vide)<DTR>_<langue>_<10chars Designation>.xlsx`
- DTR inconnu en BDD : `(vide)<DTR>_<langue>_<10chars Designation>.xlsx`

---

## EMMST / PHLST

Un seul classeur de sortie par projet, peuplé en feuille `EMMST_VBN_FR` / `PHLST` puis dupliqué par DTR (logique similaire à ATM V1).

| Langue | Fichier template | Feuilles utilisées par le code |
|---|---|---|
| FR | `Template EMMST-FR_V02.xlsx` | `Couverture`, `Sommaire`, `EMMST_VBN_FR` |
| EN | `Template PHLST-ENG_V02.xlsx` | `CoverPage`, `Index`, `PHLST` |

Chemins en dur dans `C_EMMST_Main.CopierTemplateEMMST` (pas de constante dédiée) — sous `…\005-03 Templates\02_Phase Projet\02. ASL\Fiches EMMST\`.

Feuilles annexes : `Liste choix` / `Choices list`, `acteurs` / `Acteurs` / `actors`, `REQ_EMST_SLI`, `versionning` / `Versionning`.

Naming sortie : `<Projet_Abrev>_EMMST_<Revision>.xlsx` (FR) ou `<Projet_Abrev>_PHLST_<Revision>.xlsx` (EN).

---

## Providor (LCC)

Template binaire `.xlsb` (47 feuilles) avec moteur LCC complet et tables référentielles. Le code copie le template, ouvre, remplit 7 feuilles d'entrée puis active la feuille de calcul finale.

| Fichier | Constante |
|---|---|
| `Providor v5.021.xlsb` | `PROVIDOR_TEMPLATE_PATH` |

Feuilles manipulées par `D_Providor_Main.bas` + `D_Providor_TNR.bas` :

| Feuille | Rôle | Module |
|---|---|---|
| `General Parameters` | Paramètres projet (durée, nb trains, devise…) | `D_Providor_Main` |
| `Breakdown structure` | PBS / arborescence produit | `D_Providor_BreakDownStructure` |
| `Material list` | Liste matière (Spares / Repairs) | `D_Providor_Main` |
| `Reliability & Corr. maintenance` | MTBF + maintenance corrective | `D_Providor_Main` |
| `Preventive maintenance` | Tâches préventives | `D_Providor_Maintenance` |
| `Skill List` | Niveaux de qualification | `D_Providor_Main` |
| `Tool List` | Outillage | `D_Providor_Main` |
| `LCC Calculation` | Activation finale + lecture `P18` pour TNR | `D_Providor_Main` / `D_Providor_TNR` |

TNR (`D_Providor_TNR`) compte les lignes des 7 feuilles d'entrée pour comparer avant/après refactor.

Naming sortie : selon le module `D_Providor_Finalisation`.

---

## Convention de copie locale (cache `%TEMP%`)

Pattern introduit pour ATM V2 (commit récent) — à généraliser au besoin : copier le template depuis SharePoint vers `%TEMP%\ECOSLI_template_*` **une seule fois** en début de procédure, puis utiliser cette copie locale comme source pour les N copies par DTR. Disque local ~10× plus rapide que OneDrive synchronisé.

Implémenté dans `B_ATM_V2.bas` : variable module `templateLocalV2`.
