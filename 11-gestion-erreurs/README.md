🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11. Gestion des erreurs — de `On Error` aux exceptions ⭐

> **Pas une traduction, un changement de paradigme.** On passe du modèle VB6 « piéger et brancher » (lignes, étiquettes, objet `Err`) au modèle VB.NET des **exceptions structurées** (blocs `Try`/`Catch`/`Finally`, exceptions **typées**). Le vrai danger n'est pas de traduire `On Error GoTo` — c'est `On Error Resume Next`, qui **avale les erreurs en silence**.

**Partie 3 — Migrer le langage (le cœur de la formation)** · .NET Framework 4.7.2
⏱️ *Lecture du chapitre : ~2 h · Prérequis : module 2 (notamment § 2.6) et module 10*

---

## 🧭 Pourquoi ce chapitre est important

La gestion d'erreurs **touche presque toutes les procédures** d'une application VB6 réelle. Or VB6 et VB.NET ne proposent pas deux **syntaxes** pour la même idée : ils proposent deux **modèles mentaux différents**.

- VB6 fonctionne par **« piéger et brancher »** : on arme un piège (`On Error GoTo`), et à la moindre erreur, le flux **saute** vers une étiquette ou **continue** à la ligne suivante (`Resume Next`). L'information sur l'erreur vit dans un objet global, `Err`.
- VB.NET fonctionne par **exceptions structurées** : une erreur devient un **objet** (`System.Exception`) qui **remonte la pile d'appels** jusqu'à un bloc `Catch` capable de le traiter, avec un `Finally` pour le nettoyage garanti.

Migrer, c'est donc **repenser** la logique d'erreur, pas seulement remplacer des mots-clés.

> ⚠️ **Le piège central de ce chapitre est `On Error Resume Next`** (traité en **§ 11.3**). Il demande à VB6 d'**ignorer** les erreurs et de poursuivre. Migré naïvement, il se transforme soit en **perte pure** de la gestion d'erreur, soit en **`Catch` vides** qui **enterrent les bugs**. C'est exactement le genre de *changement silencieux* que cette formation traque en priorité (voir aussi l'**Annexe B.6**).

---

## 🎯 Ce que vous saurez faire à la fin du chapitre

- Comprendre **les deux modèles** (objet `Err` / branchement vs exceptions typées) et ce qui les distingue vraiment ;
- Traduire `On Error GoTo`, `Resume`, `Resume Next` vers la **logique** `Try`/`Catch`/`Finally` (et non mot à mot) ;
- Utiliser les **filtres `When`** et capturer des **types** d'exception précis plutôt que tester un numéro après coup ;
- Identifier et **désamorcer** le cas dangereux de `On Error Resume Next` ;
- Faire correspondre l'objet **`Err`** à **`Exception`**, et `Err.Raise` à **`Throw`** ;
- Mettre en place une **migration progressive**, en faisant cohabiter temporairement `On Error` et `Try` (à l'échelle de la procédure).

---

## 🔍 Le cœur du changement : deux modèles face à face

| Aspect | VB6 | VB.NET (4.7.2) |
|---|---|---|
| Modèle | « **piéger et brancher** » (ligne / étiquette) | **exceptions structurées** (blocs / types) |
| Nature d'une erreur | un **numéro** sur l'objet global `Err` | un **objet** `System.Exception` |
| Armer / capturer | `On Error GoTo label` / `Resume Next` | `Try … Catch … Finally` |
| Sélectivité | tester `Err.Number` **après coup** | **`Catch` par type** + **filtres `When`** |
| Nettoyage garanti | **manuel** | bloc **`Finally`** |
| Propagation | remonte à l'appelant si non gérée | remonte la **pile d'appels** (typée) |
| Lever une erreur | `Err.Raise` | **`Throw`** |
| Portée | par **procédure**, état quasi global | par **bloc**, exception comme objet |

> 💡 La bonne nouvelle : `Try`/`Catch` est **plus puissant et plus lisible** (types, filtres, `Finally`, propagation propre). La migration est l'occasion de **gagner** en robustesse, pas seulement de transposer l'existant.

---

## 🌉 Une migration qui peut être progressive

Détail capital pour la stratégie (et tout le **§ 11.5**) : **VB.NET conserve `On Error`, `Err`, `Resume`** pour la compatibilité. On **n'est donc pas obligé** de tout réécrire d'un coup.

La seule règle structurante : **on ne peut pas mélanger `On Error` et `Try`/`Catch` dans la *même* procédure** (c'est une erreur de compilation). L'**unité de migration** est donc **la procédure** : on peut convertir les procédures **une par une** vers `Try`/`Catch`, en laissant les autres en `On Error` le temps de la transition.

> 🎯 Objectif final : du `Try`/`Catch` idiomatique partout. `On Error` n'est qu'une **passerelle temporaire**, à ne pas pérenniser.

---

## 🗺️ Carte du chapitre

**[11.1 — `On Error GoTo`, `Resume`, `Resume Next` : équivalents et logique de remplacement](01-on-error.md)**
Le point de départ : les mécanismes VB6 (armement, branchement, reprise) et **comment raisonner** leur remplacement — en logique, pas en traduction littérale.

**[11.2 — Convertir vers `Try`/`Catch`/`Finally` (et les filtres `When`)](02-vers-try-catch.md)**
Le modèle cible : structurer un bloc protégé, capturer par **type**, garantir le nettoyage avec `Finally`, et affiner la capture avec les **filtres `When`**.

**[11.3 — Le cas épineux de `On Error Resume Next` : pourquoi c'est dangereux à migrer](03-on-error-resume-next.md)** ⭐ ⚠️
Le point critique du chapitre. Pourquoi « ignorer et continuer » résiste à toute conversion mécanique, et comment le remplacer **sans** masquer de bugs ni perdre de robustesse.

**[11.4 — L'objet `Err` vs `Exception` ; `Err.Raise` → `Throw`](04-err-vs-exception.md)**
La correspondance des informations d'erreur : du `Err.Number`/`Err.Description` aux **propriétés d'une exception**, et de `Err.Raise` à **`Throw`** (avec les types d'exception adaptés).

**[11.5 — Stratégie de migration progressive (cohabitation `On Error` / `Try` temporaire)](05-migration-progressive.md)**
La méthode d'ensemble : migrer **procédure par procédure**, faire cohabiter les deux modèles le temps de la transition, et converger proprement vers `Try`/`Catch`.

---

## 📋 En un coup d'œil

| Élément VB6 | Équivalent / cible VB.NET | À surveiller |
|---|---|---|
| `On Error GoTo label` | `Try … Catch` | Repenser en blocs (→ § 11.1, 11.2) |
| `On Error Resume Next` | *(pas d'équivalent direct)* | ⭐ ⚠️ Ne pas masquer les erreurs (→ § 11.3) |
| `On Error GoTo 0` | Fin de portée d'un `Try` | Mécanique |
| `Resume` / `Resume Next` / `Resume label` | Restructuration du flux | Logique, pas littéral (→ § 11.1) |
| `Err.Number` / `.Description` / `.Source` | Propriétés d'`Exception` | → § 11.4 |
| `Err.Raise` | `Throw` (type adapté) | → § 11.4 |
| Nettoyage manuel | `Finally` (et `Using`) | → § 11.2 et § 12.3 |

---

## 🔗 Liens avec le reste de la formation

- **[§ 2.6 — Gestion d'erreurs : `On Error` vs exceptions structurées](../02-differences-fondamentales/06-erreurs-apercu.md)** : l'aperçu qui pose le décor ; ce chapitre en est le traitement détaillé.
- **[Module 10 — Procédures, fonctions et paramètres](../10-procedures-fonctions/README.md)** : la gestion d'erreur s'inscrit dans les procédures ; les deux chapitres se complètent.
- **[§ 12.3 — `IDisposable` / `Using` / `Finalize`](../12-poo/03-idisposable-using.md)** : `Finally` et `Using` sont la réponse moderne au **nettoyage déterministe** perdu avec `Class_Terminate`.
- **[Annexe B.6 — `On Error Resume Next` masquant des bugs](../annexes/pieges-silencieux/README.md)** : l'entrée du catalogue des pièges silencieux directement liée au **§ 11.3**.
- **[§ 19.x — Migrer avec l'aide de l'IA](../19-migration-ia/README.md)** : l'IA peut proposer une restructuration `Try`/`Catch`, mais la **logique de reprise** (`Resume`) reste à valider.

---

## ✅ À retenir avant d'attaquer la suite

1. **C'est un changement de modèle**, pas de syntaxe : « piéger et brancher » → **exceptions structurées**. La traduction se fait en **logique**.
2. **Le `Try`/`Catch` est un gain** : types, filtres `When`, `Finally`, propagation propre. Profitez-en pour **fiabiliser**.
3. **Le danger n°1 est `On Error Resume Next`** : « ignorer et continuer » n'a pas d'équivalent — à traiter avec soin (§ 11.3), sans `Catch` vides.
4. **La migration peut être progressive** : `On Error` survit, l'unité de migration est **la procédure** (on ne mélange pas les deux modèles dans une même procédure).
5. **Validez par le comportement** : une erreur autrefois ignorée qui « remonte » soudain — ou l'inverse — se détecte par les **tests de non-régression**.

---

*Première section à lire : **[§ 11.1 — `On Error GoTo`, `Resume`, `Resume Next` : équivalents et logique de remplacement](01-on-error.md)**.*

**Cible** : .NET Framework 4.7.2 · **Indicateurs** : ⭐ point clé · ⚠️ piège · 🛠️ outils

⏭️ [`On Error GoTo`, `Resume`, `Resume Next` : équivalents et logique de remplacement](/11-gestion-erreurs/01-on-error.md)
