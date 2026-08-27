# PlanningPacking.xlsm — version corrigée

Corrections apportées suite à l'analyse du fichier `PlanningPacking.xlsm` (macro `GenererPlanningPacking`).

## Ce qui a été corrigé directement dans `PlanningPacking.xlsm`

Ces corrections touchent uniquement les données/structure du classeur (XML interne), **sans toucher au binaire VBA** (`vbaProject.bin` est resté strictement identique, octet pour octet — vérifié). Zéro risque de corruption des macros, tableaux, cases à cocher ou TCD.

1. **Formules cassées supprimées** — `Plan AM + Day!W2:W89` et `Plan PM!W2:W89` contenaient `=#REF!&#REF!` (176 cellules au total, colonne masquée). Ces cellules ont été vidées.
2. **Nom LAMBDA corrompu supprimé** — `_xlpm.uniqueRand` et les `_xleta.AND/DAY/IF/OR/SMALL/VLOOKUP` (tous en `#NAME?`) étaient un résidu d'une ancienne fonction de tirage aléatoire, remplacée depuis par `ShuffleCollection` en VBA. Retirés du gestionnaire de noms.
3. **Historique allégé** — la plage utilisée allait jusqu'à la ligne 5917 pour seulement 299 lignes de données réelles (formatage résiduel sur ~5600 lignes vides). Tronquée à `A1:D299`.
4. **TCD (feuille "Sheet7") forcé à se rafraîchir à l'ouverture** (`refreshOnLoad="1"`) — son cache était désynchronisé des données actuelles (7397 enregistrements en cache vs. 299 réels), ce qui faisait apparaître une colonne fantôme « SNYO » (ancienne faute de frappe de « SYNO » qui n'existe plus dans les données). Il se mettra à jour automatiquement à l'ouverture dans Excel.

## Ce qui N'A PAS été modifié dans le binaire, et pourquoi

Le code VBA compilé (`vbaProject.bin`) est un format binaire propriétaire (OLE/Compound File). Je n'ai pas d'Excel ni de méthode fiable et vérifiable, dans cet environnement, pour le réécrire sans risquer de corrompre l'ensemble du classeur (au pire, un fichier qu'Excel refuse d'ouvrir). Plutôt que de tenter une modification binaire non vérifiable sur votre fichier de production, les 4 corrections de code sont livrées en clair dans `vba/Module1.bas` et `vba/Module2.bas`, prêtes à coller dans l'éditeur VBA (2 minutes, sans risque) :

### `vba/Module1.bas`
- **Réutilisation de `wsFormAM` / `wsFormPM`** au lieu de refaire `wb.Sheets("Formule AM + Day")` / `wb.Sheets("Formule PM")` en dur : ces deux variables étaient déjà déclarées et assignées avec `On Error Resume Next` (donc tolérantes si la feuille est absente/renommée), mais jamais réutilisées ensuite. Le code rappelait les feuilles par leur nom sans protection, ce qui plantait avec une erreur non gérée (« Subscript out of range ») si l'une de ces feuilles venait à être renommée — alors que `RemplirPlanBanc` a justement un test `If wsForm Is Nothing... Exit Sub` prévu pour ce cas, mais qui n'était jamais atteint.
- **`Randomize` appelé une seule fois** en tête de `GenererPlanningPacking` au lieu d'être appelé à chaque exécution de `ShuffleCollection` (jusqu'à 8 fois par génération, en quelques millisecondes) — évite un ré-amorçage du générateur aléatoire à partir d'une valeur d'horloge quasi identique, qui pouvait corréler les mélanges de listes tirées à la suite.

- **Erreur 1004 `.PaperSize = xlPaperA3` neutralisée** dans `PreparerFeuilleSO_Dual` et `PreparerFeuilleLO_Dual` — cette propriété dépend du pilote d'imprimante installé sur la machine (aucune imprimante par défaut, ou imprimante par défaut ne supportant pas le A3 → erreur d'exécution 1004 qui plantait toute la macro). Encadré d'un `On Error Resume Next` / `On Error GoTo 0` pour que la génération du planning continue même si la mise en page A3 ne peut pas s'appliquer. **Configurez tout de même une imprimante compatible A3 par défaut** (ou "Microsoft Print to PDF") si vous voulez que la mise en page A3 soit effectivement appliquée.

### `vba/Module2.bas`
- **Confirmation ajoutée avant `Clean()`** : la macro effaçait silencieusement `ActiveSheet.Rows("2:...")` sans aucune confirmation ni vérification de l'onglet actif — risque de perte de données si lancée depuis le mauvais onglet. Une boîte de dialogue Oui/Non a été ajoutée.
- **`Option Explicit` ajouté** (déjà présent dans Module1, manquant dans Module2 — bonne pratique pour détecter les fautes de frappe de variables).

### Comment les appliquer
1. Ouvrez `PlanningPacking.xlsm` dans Excel, `Alt+F11` pour l'éditeur VBA.
2. Double-cliquez sur `Module1` dans l'explorateur de projet, sélectionnez tout (`Ctrl+A`), collez le contenu de `vba/Module1.bas`.
3. Idem pour `Module2` avec `vba/Module2.bas`.
4. `Ctrl+S` (garder le format `.xlsm`).

**Attention à l'encodage en ouvrant les `.bas`** : ce sont des fichiers UTF-8 (avec BOM). Le Bloc-notes Windows moderne (Windows 10/11) et VS Code les détectent correctement — les accents (`é`, `à`, `ç`) s'affichent bien. Si un accent s'affiche comme `Ã©` ou `Ã ` à l'ouverture, l'éditeur a mal détecté l'encodage : ne copiez pas ce texte tel quel (vous copieriez la corruption dans le code VBA), rouvrez plutôt le fichier en forçant l'encodage UTF-8 (dans Notepad : Fichier > Ouvrir > choisir "UTF-8" dans le menu déroulant d'encodage en bas de la boîte de dialogue).

## Non traité (nécessite un choix métier, pas juste un bug)

- Le rôle « LIFTTRUCK DRIVER » (3 personnes dans l'onglet Import) ne suit aucune règle dédiée et passe par le chemin normal SO/LO selon sa zone — à confirmer que c'est le comportement voulu.
- Duplication de code entre les 4 blocs SO-AM / SO-PM / LO-AM / LO-PM dans `GenererPlanningPacking` (~500 lignes quasi identiques) — piste de refactorisation en fonction générique paramétrée, non appliquée ici pour ne pas modifier le comportement métier sans validation.
