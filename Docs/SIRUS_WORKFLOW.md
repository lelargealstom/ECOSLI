# SIRUS — Workflow des demandes d'import ATMAP

Référence des statuts d'une demande d'import (ticket ATMAP_TICKETS, SharePoint) et des transitions autorisées.

Document de spécification — sera implémenté côté SharePoint (validation de transition) et Power Automate (alertes).

---

## États

| Statut | Sens | Terminal ? |
|---|---|:---:|
| **Nouveau** | Ticket créé par le Demandeur depuis SIRUS, en attente de prise en charge. | non |
| **En cours de traitement** | Le Data Owner a pris en charge le ticket et est en train d'intégrer la donnée en BdD. | non |
| **Refusé** | Le Data Owner a rejeté la demande (donnée incomplète, incohérente, hors scope, etc.). Le Demandeur peut corriger pour relancer. | non |
| **Résolu** | La donnée est intégrée en BdD. Fin du cycle de vie du ticket. | **oui** |

## Transitions autorisées

| De \ Vers | Nouveau | En cours | Refusé | Résolu |
|---|:---:|:---:|:---:|:---:|
| **Nouveau** | — | ✅ | ✅ | ✅ |
| **En cours** | ❌ | — | ✅ | ✅ |
| **Refusé** | ❌ | ✅ | — | ✅ |
| **Résolu** | ❌ | ❌ | ❌ | ✅ (no-op) |

Lecture : depuis `Nouveau` on peut passer en `En cours`, `Refusé` ou `Résolu`. Depuis `Résolu`, plus aucune transition possible (état terminal, le ticket reste figé).

### Schéma

```
                    ┌────────────────┐
        ┌──────────►│   En cours     │◄─────────┐
        │           └────────┬───────┘          │
        │                    │                  │
        │                    ▼                  │
   ┌────┴────┐   ┌─────────────────┐            │
   │ Nouveau │──►│     Refusé      │────────────┘
   └────┬────┘   └────────┬────────┘
        │                 │
        │                 ▼
        │           ┌──────────┐
        └──────────►│  Résolu  │ (terminal)
                    └──────────┘
```

## Création du ticket (côté SIRUS)

Lorsque le Demandeur clique sur **Btn_SIRUS_Envoyer** dans le UserForm ECOSLI, `SIRUS_UI_BtnEnvoyerClick` (G_SIRUS_UserForm) crée le ticket avec :

| Champ SP | Valeur initiale | Source |
|---|---|---|
| `Numero_ticket` | `VBN_NN` | `SIRUS_NumeroVBNSuivant()` — auto-incrément SP, max+1 |
| `Numero_Import` | `IMP-YYYY-NNN` | `SIRUS_NumeroTicketSuivant()` — extrait du sujet pré-rempli |
| `Sujet` | `IMP-YYYY-NNN - Import <Op> <Table>` | Pré-rempli, modifiable |
| `Statut` | `Nouveau` | constante (G_SIRUS_ATMAP.SIRUS_CreerTicketEnvoi) |
| `Description` | `<Op> de <N> reference(s) en <Table> pour le projet <P>` | helper `SIRUS_MajDescriptionEnvoi` |
| `Attentes` | `Traitement sous N jour(s)` | helper `SIRUS_JoursTraitementSelonPriorite` |
| `Typedemande` | `Import` | constante |
| `Priorité` | High / Medium / Low | choix Demandeur, défaut Medium |
| `Date_creation` | aujourd'hui ISO | `Format(Now, "yyyy-mm-dd") & "T00:00:00Z"` |
| `Date_estimee` | aujourd'hui + 1/3/5 j | `SIRUS_DateEstimeeISO(priorite)` |
| `DemandeurId` | userID SP | `SIRUS_GetCurrentUserID(digest)` |
| `Assigné_à_Id` | userID SP (auto-assigné Demandeur en attendant prise en charge) | idem |
| `NomoutilId` | 43 (ECOSLI / SIRUS) | constante |
| PJ attachée | classeur Excel `.xlsx` avec toutes les feuilles SIRUS_* | `SIRUS_ConstruirePJ` + `SIRUS_AttacherFichierAuTicket` |

## Règle priorité → délai

Centralisé dans `G_SIRUS_ATMAP.SIRUS_JoursTraitementSelonPriorite(priorite)`. Une seule source de vérité, utilisée par :
- `SIRUS_UI_PrioriteChange` (texte du champ `Attentes` côté UI)
- `SIRUS_DateEstimeeISO` (champ SP `Date_estimee`)

| Priorité | Délai cible | Date_estimee |
|---|:---:|---|
| High | 1 jour | aujourd'hui + 1 |
| Medium (défaut) | 3 jours | aujourd'hui + 3 |
| Low | 5 jours | aujourd'hui + 5 |

## Côté Data Owner (hors SIRUS, dans ATMAP)

Le Data Owner reçoit les tickets dans la vue ATMAP_TICKETS filtrée sur `Typedemande = "Import"`. Pour chaque ticket :

1. Ouvre la PJ Excel attachée → vérifie les valeurs proposées (Création) ou le diff (Modification, à venir Phase 3.2).
2. Passe le statut à **En cours de traitement** (signale au Demandeur que la prise en charge a démarré).
3. Soit :
   - **Intégration OK** → copie-colle dans la BdD Excel partagée, passe le statut à **Résolu**.
   - **Donnée à corriger** → passe le statut à **Refusé** + commente le motif dans `Description` ou champ dédié `Commentaire_DataOwner` (à créer si besoin).

Le Demandeur ne touche jamais directement à la BdD. C'est un point de gouvernance volontaire.

## Power Automate — alertes à mettre en place

Les flows sont à créer côté Power Automate sur la liste SharePoint ATMAP_TICKETS.

### Flow 1 — Alerte Refusé (priorité)

| Élément | Valeur |
|---|---|
| Trigger | `When an item is created or modified` sur ATMAP_TICKETS |
| Condition | `Statut` a changé et nouvelle valeur = `Refusé` |
| Action | `Send email (V2)` à `Demandeur.Email` avec sujet `[SIRUS] Demande X refusée` |
| Corps mail | `Numero_Import`, `Sujet`, motif (= `Description` ou `Commentaire_DataOwner`), lien vers le ticket SP |

### Flow 2 — Alerte Résolu (optionnel, traçabilité)

| Élément | Valeur |
|---|---|
| Trigger | idem |
| Condition | `Statut` a changé et nouvelle valeur = `Résolu` |
| Action | Mail Demandeur `[SIRUS] Demande X intégrée en BdD` |

### Flow 3 — Alerte retard (optionnel)

| Élément | Valeur |
|---|---|
| Trigger | `Recurrence` quotidien |
| Condition | `Date_estimee < utcNow()` ET `Statut` ∉ (`Résolu`, `Refusé`) |
| Action | Mail Data Owner `[SIRUS] Ticket X en retard de N jour(s)` |

## À venir / non implémenté

- Champ `Commentaire_DataOwner` (motif de refus) — à créer côté SP + à intégrer dans le flow 1.
- Réception du Refus côté SIRUS : afficher l'historique des refus pour les tickets du Demandeur dans le UserForm (cf. Phase 6.1 Historique).
- Modification d'une donnée déjà en BdD : ne ré-ouvre pas le ticket original (Résolu = terminal), on crée un **nouveau** ticket avec référence à l'ancien `Numero_Import` (champ `Lien_ticket_origine` à créer si besoin).
