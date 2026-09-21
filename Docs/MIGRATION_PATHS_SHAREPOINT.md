# Migration paths SharePoint — inventaire à modifier

Cartographie des chemins hardcodés dans ECOSLI.xlam impactés par le remastering
de l'architecture SharePoint BdD SLI (nouveau dossier cible : `01. BDD SLI\02_BDD_Prod_LGM`).

Document de travail — à supprimer après exécution de la migration et bump du CHANGELOG.

---

## 1. Fichier principal (95% du travail) — `Modules/Z_Variables.bas`

Deux constantes à mettre à jour :

- **Ligne 126** : `CONNEXION_EXCEL_BDD_PROD`
  ```
  Data Source=UserProfile\Alstom\DIS Service VBN Pole ISR - Outils Metier ISR\01. BDD SLI\11_Livraison de l'OUTIL excel BDD\BDD excel original V4.1.xlsb
  ```
  → nouveau segment `\01. BDD SLI\02_BDD_Prod_LGM\BDD excel original V4.1.xlsb`

- **Ligne 127** : `CONNEXION_EXCEL_BDD_TRAVAIL`
  ```
  Data Source=UserProfile\Alstom\DIS Service VBN Pole ISR - Outils Metier ISR\01. BDD SLI\11_BdD de travail LGM\BDD excel original V4.1.xlsb
  ```
  → nouveau segment (à confirmer : même fichier que PROD ? ou un `02_BDD_Travail_LGM` distinct ?)

**Attention accent** : si le dossier racine s'appelle désormais "Outils **Métier**" (avec accent) et non "Metier", il faudra reconstruire la string via `ChrW(&HE9)` pour rester ASCII pur dans le `.bas` (même convention que €, …).

Ces 2 constantes sont consommées via `Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", Environ$("UserProfile"))` dans **18 endroits** — ces 18 endroits n'ont **rien** à modifier, ils héritent automatiquement.

### Consommateurs des constantes (héritent auto)

- `B_ATM_V1.bas:957`
- `B_ATM_V2.bas:106`
- `C_EMMST_Remplissage.bas:417`
- `F_FonctionsGlobales.bas:440`
- `G_PBS_Refresh.bas:227, 279, 355, 662, 702, 1087, 1117, 1188, 1278, 1326`
- `G_SIRUS_Generation.bas:67`
- `G_SIRUS_PBS.bas:402, 404, 775, 1423, 1549`
- `G_SIRUS_Regles.bas:315`
- `X_Test_Appel_BDD.bas:26`

---

## 2. Fichiers du complément ECOSLI sur SharePoint (installation / auto-update)

- **`Modules/A_1_ModEvents.bas:203`** — chemin `VersionComplement.txt` (`11_Livraison de l'OUTIL excel BDD`)
- **`Modules/A_2_ModUpdate.bas:22`** — chemin `ecosli.xlam` source (dans `Install`)
- **`Modules/A_2_ModUpdate.bas:248`** — même chemin `ecosli.xlam` (dans une autre sub, sans doute `Update`)

Ces 3 lignes utilisent l'ancien schéma SharePoint `Svcs_DIS_Fr_VBN - Documents\01_Metiers\01_ISR\007 - Outils metier ISR\...` (différent de la BdD, plus long). À migrer vers la nouvelle archi **si** `ecosli.xlam` et `VersionComplement.txt` déménagent aussi.

---

## 3. Références obsolètes (probablement mortes, à vérifier avant modif)

Ces fichiers pointent vers d'anciennes versions (`V03.xlsm`, `V04.xlsb`) — sans doute des reliquats non appelés :

- **`Modules/V_Nom_ADODB_Type.bas:20`** — sub `Base_V04` : path vers `BDD excel original V04.xlsb`
- **`Modules/O_Projet.bas:26`** — `Workbooks.Open` vers `BDD excel original V03.xlsm`
- **`Modules/X_Test_Appel_BDD.bas:51, 95, 127`** — 3 chemins de test (V03.xlsm et V04.xlsb)
- **`Modules/Q_TestComparaison.bas:10`** — dossier `VBA_Export` sous `01_Outils MACRO EXCEL` (sûrement mort aussi)

**Recommandation** : soit tu me dis de les migrer aussi, soit tu me dis qu'ils sont morts et je te propose de les supprimer une fois qu'on a confirmé qu'ils ne sont jamais appelés.

---

## 4. Templates et hors-BdD (à confirmer)

- **`Modules/Z_Variables.bas:124`** — `PROVIDOR_TEMPLATE_PATH` : `Svcs_DIS_Fr_VBN - Documents\01_Metiers\01_ISR\005 - Qualite\...\Providor v5.021.xlsb`
   - **Pas dans le scope BdD** — c'est le template Providor stocké dans le dossier Qualité SharePoint. À ne PAS modifier sauf si ce dossier bouge aussi.

---

## 5. Fichiers auxiliaires du repo (hors code VBA)

- **`Docs/CHANGELOG.md`** — nouvelle entrée à ajouter à la fin de la session
- **`Docs/ECOSLI_SYNTHESE.md`** et **`README.md`** — si mention des paths (à vérifier, souvent hors-jour)

---

## Résumé chiffré

| Zone | Nb lignes à toucher | Blocage |
|---|:---:|---|
| BdD (Z_Variables) | 2 | attend confirmation PROD/TRAVAIL + accent |
| Auto-update ecosli.xlam | 3 | attend confirmation destination `ecosli.xlam` |
| Refs mortes V03/V04 | 5 | à valider si obsolètes ou à migrer |
| Template Providor | 0 (probable) | confirmer si le dossier Qualité bouge |
| Doc | 1-2 | après code |

---

## Questions ouvertes (à confirmer avant modif)

1. **Nom du dossier racine synchronisé** : "Outils Métier ISR" (avec accent aigu) ou "Outils Metier ISR" (sans accent) côté FS Windows ?
2. **BdD PROD vs TRAVAIL** :
   - a) Les deux constantes pointent désormais sur le même `02_BDD_Prod_LGM\BDD excel original V4.1.xlsb` (auquel cas on peut envisager de fusionner les 2 constantes) ?
   - b) Il existe un `02_BDD_Travail_LGM` (ou équivalent) que le tree fourni n'affiche pas ?
3. **`ecosli.xlam` + `VersionComplement.txt`** : où vont-ils dans la nouvelle archi ?
   - a) Déplacés ailleurs (chemin à préciser) ?
   - b) `11_Livraison...` existe encore ailleurs ?
   - c) À placer à la racine `02_BDD_Prod_LGM` à côté du `.xlsb` ?

---

## Après la migration (checklist test)

1. Ouvrir un classeur projet avec `DATA_ASL`.
2. `Ctrl+W` → UserForm ECOSLI ouvre sans erreur.
3. Bouton "Rafraichir Data PBS" → succès sur toutes les colonnes.
4. ListView Generation → MeQ affiché sur tous les DTR.
5. Génération SIRUS Produit Créer → ADODB OK.
6. Envoi ATMAP (dry-run si possible).
7. Vider `%TEMP%\Providor_Etalon.xlsb` et autres caches locaux avant les 1ers tests.
