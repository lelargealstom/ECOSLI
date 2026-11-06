# Instructions projet — ECOSLI.xlam

Instructions pour toute session d'assistance IA (Claude Code) sur ce dépôt.

---

## Documentation à tenir à jour à chaque évolution de l'outil

À chaque modification du code ou de la structure du dépôt, mettre à jour **systématiquement** dans le même change set (ou immédiatement après) :

1. **`Docs/CHANGELOG.md`** — entrée datée pour chaque session, avec un bloc par commit de la session. Format existant à respecter (cf. entrées précédentes) :
   - `## AAAA-MM-JJ — Titre synthétique`
   - Phrase de contexte "Session focus : …"
   - Liste des commits de la session.
   - Un sous-bloc `### Commit <hash> — Titre` par commit, détaillant les modifications fichier par fichier.
   - Si pertinent, un bloc final `### À tester en VBE` listant les cas de non-régression / validation.

2. **`Docs/*.md`** concernés — si l'évolution impacte un comportement documenté (règles SIRUS, workflow, roadmap, templates, logique d'extraction, synthèse…), mettre à jour la doc correspondante dans la foulée. Ne pas laisser un `.md` affirmer un fait qui vient d'être invalidé par un commit.

3. **Mémoire persistante Claude** (`~/.claude/projects/<proj>/memory/`) — si l'évolution introduit une convention structurante, une règle métier, ou modifie un comportement durable, mettre à jour la mémoire concernée (`project_*.md`, `feedback_*.md`) pour que les futures sessions soient à jour.

### Quand

- **Toujours après un commit qui modifie le code VBA** (`.bas`, `.frm`, `.cls`) ou le template.
- Toujours après un commit qui modifie la structure du dépôt (ajout/suppression de modules, migration PROD/PREPROD).
- Pas forcément après un commit purement documentaire (le CHANGELOG lui-même, ou un .md de cadrage) — mais signaler l'intention dans le message de commit.

### Comment

- Écrire l'entrée CHANGELOG **dans un commit séparé** avec un message du type `CHANGELOG : session AAAA-MM-JJ <titre>` (cf. commits `eb6081e`, `7a9057a`, `dd2c51f` pour le pattern).
- Si plusieurs commits dans la même session, les regrouper sous une même date dans le CHANGELOG.
- En cas d'oubli, rétablir la cohérence **dès que l'oubli est détecté** avec une entrée rétroactive.

---

## Autres règles projet

Les conventions techniques VBA (Find/Type, ListView, ChrW, ActiveWorkbook, pas de handler défensif…) sont documentées dans la mémoire Claude (`feedback_vba_conventions.md`). Toute règle structurante découverte en session doit y être consolidée.
