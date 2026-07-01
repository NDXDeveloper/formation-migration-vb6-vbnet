🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9. Opérateurs, chaînes et structures de contrôle

> **Partie 3 — Migrer le langage (le cœur de la formation)**
> **Pré-requis** : module 7 (types et variables) et module 8 (tableaux et collections).
> **Cible** : .NET Framework 4.7.2 · Visual Studio 2026

Les opérateurs, les fonctions de chaînes et les structures de contrôle forment la **grammaire quotidienne** de votre code : on en trouve dans presque chaque ligne. C'est précisément ce qui rend ce chapitre à la fois rassurant et dangereux.

**Bonne nouvelle** : l'essentiel de cette grammaire migre de façon *mécanique*. Une boucle `For…Next` reste une boucle `For…Next` ; un `Select Case` reste un `Select Case`. L'assistant de mise à niveau traite correctement la plus grande partie.

**Mauvaise nouvelle** : c'est aussi ici que se cachent quelques-uns des **pièges silencieux** les plus sournois de toute la migration — du code qui **compile sans broncher** mais qui **ne s'exécute plus de la même manière**. Aucune erreur de compilation ne vous préviendra. C'est exactement le type de régression que cette formation place en priorité (voir **Annexe B**).

Ce chapitre s'organise autour de **trois fils rouges**.

---

## 🧵 Fil rouge n°1 — L'évaluation **court-circuit** (le piège le plus coûteux)

C'est, de loin, le changement de comportement le plus important du chapitre. Il se manifeste à **deux endroits distincts**.

**Les opérateurs logiques.** En VB6, `And` et `Or` évaluent **toujours les deux opérandes** (ce sont avant tout des opérateurs *bit à bit*, utilisés aussi comme opérateurs logiques). VB.NET **conserve ce comportement à l'identique** pour `And` / `Or` — mais introduit deux nouveaux opérateurs **à court-circuit** : `AndAlso` et `OrElse`, qui n'évaluent le second opérande **que si nécessaire**.

**La fonction `IIf`.** En VB6, `IIf(condition, valeurSiVrai, valeurSiFaux)` est une **fonction** : ses **deux** arguments sont évalués, quel que soit le résultat de la condition. VB.NET propose l'**opérateur** `If(condition, valeurSiVrai, valeurSiFaux)` qui, lui, **court-circuite** (un seul des deux côtés est évalué). Subtilité : `IIf` **existe toujours** en VB.NET et conserve son ancien comportement « tout évaluer ». Migrer `IIf` → `If()` **change donc la sémantique**.

> ⚠️ **Pourquoi c'est dangereux dans les deux sens.**
> Du code reposant sur l'évaluation systématique des deux branches — par exemple une fonction à **effet de bord** glissée dans un `IIf` — **cessera de s'exécuter** si vous basculez naïvement vers `If()`. À l'inverse, ne pas introduire `AndAlso` / `OrElse` là où le code *suppose* un court-circuit (test de `Nothing` puis accès à un membre dans la même condition) laisse subsister un risque d'erreur à l'exécution. Détail complet et stratégie de conversion en **9.1** et **9.3**.

---

## 🧵 Fil rouge n°2 — « Même nom, comportement différent » (les chaînes)

VB.NET hérite de l'espace de noms `Microsoft.VisualBasic`, qui propose les **mêmes fonctions de chaînes** qu'en VB6 (`Left`, `Right`, `Mid`, `Len`…). Très pratique pour migrer vite — mais ces fonctions cohabitent désormais avec les **méthodes de l'objet `String` du .NET**, et les deux familles **ne se comportent pas de la même façon** :

- **Indexation 1-based vs 0-based** : `Mid(s, 1, 3)` part du **1er** caractère ; `s.Substring(0, 3)` part de l'**indice 0**. Mélanger les deux conventions est une cause classique de décalages d'un caractère.
- **L'instruction `Mid`** (à distinguer de la *fonction*) : `Mid(s, 2, 3) = "abc"` modifie une chaîne **en place** en VB6. En .NET, les chaînes sont **immuables** — ce motif demande réflexion (réaffectation, `StringBuilder`…).
- **`Left` / `Right` ambigus** : sur un formulaire, `Left` et `Right` sont aussi des **propriétés de position**. Selon le contexte, le code peut ne pas désigner ce que l'on croit.

> 🔑 **La vraie question n'est pas technique, elle est stratégique** : garder les fonctions `Microsoft.VisualBasic` (rapide, familier, 1-based) ou basculer vers les méthodes `String` (idiomatique, 0-based) ? C'est l'un des arbitrages « béquille vs code propre » récurrents de la formation. À traiter en **9.2**, puis lors de la fiabilisation (**module 17**).

---

## 🧵 Fil rouge n°3 — Ce qui **disparaît** purement et simplement

Certaines constructions héritées des anciens BASIC **n'existent plus** en VB.NET. Pour une fois, c'est plutôt une **bonne nouvelle** : elles ne passent pas inaperçues, elles provoquent une **erreur de compilation** — le compilateur vous **force** à les traiter. Il faut **refactoriser** :

- **`GoSub` / `Return`** : supprimés. Pire, le mot-clé `Return` **existe toujours** mais signifie désormais « sortir de la procédure / renvoyer une valeur » — un **faux-ami** à connaître. → transformer chaque `GoSub` en véritable `Sub` / `Function`.
- **`On…GoTo` / `On…GoSub`** (saut calculé) : supprimés. → réécrire en `Select Case`.
- **Numéros de ligne** (`10 …`, `20 …`) et l'usage de `Erl` : à éliminer. → refactoriser, en particulier la gestion d'erreurs qui s'appuyait dessus (voir **module 11**).

Détaillé en **9.5** et **9.6**.

---

## 📘 Ce que vous allez apprendre

| Section | Sujet | Point de vigilance |
|---------|-------|--------------------|
| **9.1** | Opérateurs : `&` vs `+`, `\` (division entière), `Mod`, `Like`, `Is` / `IsNot` | Concaténation avec `+`, opérateurs logiques |
| **9.2** ⚠️ | Fonctions de chaînes : `Microsoft.VisualBasic` vs `String` | 1-based vs 0-based, instruction `Mid`, immuabilité |
| **9.3** ⚠️ | Conditions : `If…Then`, `Select Case`, `IIf` → opérateur `If()` | Évaluation court-circuit |
| **9.4** | Boucles : `For…Next`, `For Each`, `Do…Loop`, `While…Wend` → `While…End While` | `Wend` supprimé |
| **9.5** ⚠️ | `GoSub` / `Return`, `On…GoTo`, numéros de ligne : **supprimés** | Refactoring obligatoire, faux-ami `Return` |
| **9.6** | Instructions multiples, séparateur « `:` », étiquettes de ligne | Lisibilité, labels conservés |

En clair : **9.1** et **9.3** portent le risque de **court-circuit** ; **9.2** porte le risque d'**indexation et d'immuabilité** des chaînes ; **9.5** et **9.6** concernent les constructions **disparues** qu'il faut réécrire. Les sections **9.4** et la première moitié de **9.1** sont, elles, largement mécaniques.

---

## 🎯 Synthèse — les points de vigilance du chapitre

| Élément | VB6 | VB.NET | Risque à la migration |
|---------|-----|--------|------------------------|
| `And` / `Or` | Évaluent **les 2** opérandes | Idem (`AndAlso` / `OrElse` pour le court-circuit) | Comportement conservé, mais court-circuit « oublié » là où il faudrait |
| `IIf(...)` | Fonction : **2 branches** évaluées | `If(...)` : opérateur, **court-circuit** | Changement **silencieux** de sémantique |
| `+` sur chaînes | Concatène **ou** additionne selon les types | Sous `Option Strict On` : typé / restreint | Conversions ou erreurs inattendues — préférer `&` |
| `Mid`, `Left`, `Right` | 1-based (`Microsoft.VisualBasic`) | 1-based **ou** méthodes `String` 0-based | Décalages d'indice |
| Instruction `Mid` (`=`) | Modifie la chaîne **en place** | Chaînes **immuables** | Motif à repenser (`StringBuilder`) |
| `While…Wend` | Valide | **`Wend` supprimé** → `End While` | Erreur de compilation (traitée par l'assistant) |
| `GoSub` / `Return` | Valides | **Supprimés** ; `Return` = retour de procédure | Refactoring + faux-ami |
| `On…GoTo` / `On…GoSub` | Valides | **Supprimés** | Réécriture en `Select Case` |
| Numéros de ligne / `Erl` | Valides | À éviter / refactoriser | Gestion d'erreurs à reprendre |

---

## 🧭 Comment aborder ce chapitre

- **Activez le filet de sécurité d'abord.** La plupart des pièges de ce chapitre se révèlent **à l'exécution**, pas à la compilation : assurez-vous d'avoir votre **harnais de tests de référence** (*golden master*, voir **module 5.5**) avant de toucher au code.
- **Pensez `Option Strict`.** Plusieurs comportements (notamment `+` sur chaînes) changent selon les options de compilation. La stratégie d'activation **progressive** est traitée au **module 6.4**, puis au **module 17.1**.
- **Gardez les annexes ouvertes.** L'**Annexe A** (correspondance VB6 → VB.NET) recense les équivalents d'opérateurs et de fonctions ; l'**Annexe B** (catalogue des pièges silencieux) détaille notamment `True = -1`, les conversions booléennes et les arrondis.

**Liens utiles dans la formation :**

- Module **2.4** — *ByRef par défaut → ByVal* : un autre changement silencieux d'appel.
- Module **11** — *Gestion des erreurs* : pour la disparition des numéros de ligne et de `Erl`.
- Module **17.4** — *Traquer les pièges silencieux* après migration.
- Annexe **B** — *Catalogue des pièges silencieux* : la référence à garder ouverte.

---

> 🏷️ **Légende** : ⭐ point clé · ⚠️ piège fréquent · 🔗 interop COM
>
> **Section suivante** → [9.1 — Opérateurs (`&` vs `+`, `\`, `Mod`, `Like`, `Is` / `IsNot`)](01-operateurs.md)

⏭️ [Opérateurs (`&` vs `+`, `\` division entière, `Mod`, `Like`, `Is`/`IsNot`)](/09-operateurs-controle/01-operateurs.md)
