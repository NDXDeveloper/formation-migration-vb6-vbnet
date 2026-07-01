🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.3 — Le cas épineux de `On Error Resume Next` : pourquoi c'est dangereux à migrer ⭐ ⚠️

> **Le point le plus dangereux du chapitre.** « Ignorer et continuer » n'a **aucun équivalent** en `Try`/`Catch` — pas même un gros `Try` à `Catch` vide, qui, lui, **s'arrête à la première erreur**. Pire : `On Error Resume Next` **compile encore** en VB.NET. Le migrer, c'est **repenser** la logique d'erreur, pas la transposer.

📍 *Module 11 · § 11.3 · [↑ Introduction du chapitre](README.md) · [← § 11.2](02-vers-try-catch.md)*

---

## ⚡ En bref

- `On Error Resume Next` (OERN) installe un **mode** : à la moindre erreur, VB6 **ignore** et **continue à l'instruction suivante**, en gardant tout l'état local.
- Il **n'a pas d'équivalent** en `Try`/`Catch` : un `Try` **abandonne** le bloc à la première erreur, là où OERN **poursuit après chaque** erreur. Le seul équivalent *littéral* serait un `Try`/`Catch` **par instruction** — absurde.
- ⚠️ **Un gros `Try` à `Catch` vide n'est PAS équivalent** : il s'arrête à la première erreur **et** masque les bugs — **doublement faux**.
- **OERN compile en VB.NET.** Le danger n'est donc pas qu'il bloque la migration, mais qu'il **survive en silence**, en continuant d'**enterrer les erreurs**. Le migrer demande une **décision par erreur** : gérer, ignorer (rarement), ou laisser propager.

---

## 1. Ce que fait vraiment `On Error Resume Next`

OERN demande : « à la moindre erreur, **n'interromps rien**, passe simplement à l'instruction suivante. » L'objet `Err` reste positionné, d'où l'usage **idiomatique** : tenter une opération, puis **vérifier** `Err.Number`.

```vb
' VB6 — usage idiomatique : tenter, puis VÉRIFIER
On Error Resume Next
valeur = CInt(texte)               ' peut échouer (erreur 13 : incompatibilité de type)
If Err.Number <> 0 Then
    valeur = 0
    Err.Clear                      ' sinon Err reste "sale" pour la suite
End If
On Error GoTo 0                    ' on rétablit le comportement normal
```

La **portée** d'OERN court jusqu'à `On Error GoTo 0` (désarmement), un autre `On Error GoTo …`, ou la **fin de la procédure**. C'est cette portée qui distingue un usage **localisé** d'une **suppression massive** (§ 3).

---

## 2. Pourquoi il n'a aucun équivalent direct

C'est le cœur du problème. `Try`/`Catch` et OERN ont des comportements **opposés** sur le point essentiel : *que se passe-t-il après une erreur ?*

- **`Try`/`Catch`** : l'erreur **interrompt** le bloc et transfère au `Catch`. Les instructions **suivantes du `Try` ne s'exécutent pas**.
- **OERN** : l'erreur est **ignorée**, et l'instruction **suivante s'exécute quand même**.

```vb
' VB6 — chaque ligne qui échoue est IGNORÉE ; les suivantes s'exécutent
On Error Resume Next
a = Operation1()      ' si échec → ignoré
b = Operation2()      ' s'exécute de toute façon
c = Operation3()      ' s'exécute de toute façon
```

```vb
' VB.NET — ❌ un gros Try + Catch vide n'est PAS équivalent
Try
    a = Operation1()
    b = Operation2()  ' si Operation1 a échoué, on N'ARRIVE JAMAIS ici
    c = Operation3()  ' idem
Catch ex As Exception
    ' rien
End Try
```

Si `Operation1` échoue, OERN exécute tout de même `b` et `c` ; le gros `Try` **les saute**. Les deux comportements **diffèrent**, et le `Catch` vide **enterre** en prime l'erreur. **Doublement faux.**

Le seul équivalent **littéral** d'OERN serait d'isoler **chaque** instruction :

```vb
' VB.NET — seul équivalent LITTÉRAL : un Try/Catch par instruction (À NE PAS FAIRE)
Try : a = Operation1() : Catch : End Try
Try : b = Operation2() : Catch : End Try
Try : c = Operation3() : Catch : End Try
```

Illisible, et presque jamais l'intention réelle. **Conclusion** : on **ne peut pas** préserver la sémantique d'OERN avec un `Try`/`Catch` ; il faut donc **comprendre l'intention** et **reconstruire**.

---

## 3. Les deux visages d'OERN (le diagnostic clé)

Avant de migrer, **diagnostiquez** chaque OERN. Il en existe deux usages très différents.

### 3.1 Vérification en ligne, **localisée** (légitime)

OERN **immédiatement** suivi d'un test de `Err.Number`, autour d'**une** (ou de très peu d') opération. L'intention est claire : « tenter ceci, et gérer si ça rate ». Ce cas se migre **proprement** (§ 5, § 6).

### 3.2 Suppression **en bloc** (le vrai problème)

OERN en **tête de procédure**, **aucun** test de `Err.Number`, sur une **large** portée :

```vb
' VB6 — OERN avale TOUTES les erreurs de la procédure
Sub Traiter()
    On Error Resume Next
    ' ... 50 lignes ...
    ' AUCUN test de Err.Number
End Sub
```

Ici, **toute** erreur — attendue ou non — est silencieusement avalée. L'intention est **inconnue** : negligence, suppression d'une erreur précise mal comprise, ou simple méconnaissance des conséquences.

**Pourquoi c'est nocif** : une erreur à l'instruction *N* est ignorée, mais *N+1*, *N+2*… s'exécutent sur un **état invalide ou partiel**. Une variable qui aurait dû être renseignée ne l'est pas ; un calcul ultérieur utilise une valeur par défaut ou périmée ; le résultat est **faux** — **sans plantage, sans trace, sans signe**. C'est le *changement silencieux* ultime que cette formation traque (voir **[Annexe B.6](../annexes/pieges-silencieux/README.md)**).

---

## 4. Le double tranchant à la migration

Comme OERN **compile en VB.NET**, on est face à trois mauvaises tentations et une bonne voie :

| Tentation | Conséquence |
|---|---|
| **Laisser `On Error Resume Next`** tel quel (ça compile !) | La dette est **conservée** : les bugs restent enterrés |
| **Supprimer** purement OERN | Les erreurs autrefois avalées **remontent** : l'appli **plante** là où elle « marchait » |
| **Envelopper** dans un gros `Catch` vide | Comportement **modifié** (arrêt à la 1ʳᵉ erreur) **et** masquage **préservé** |
| ✅ **Décider erreur par erreur** | La seule voie correcte (§ 5) |

> ⚠️ Aucune transformation **automatique** n'est correcte. Supprimer et envelopper sont **tous deux** des pièges : l'un change le comportement vers le plantage, l'autre fige un masquage de bugs. Il faut **trancher**, opération par opération.

---

## 5. La méthode de migration

### Étape 1 — Localiser chaque OERN et sa **portée**

Repérez tous les `On Error Resume Next` et délimitez où chacun **se termine** (`On Error GoTo 0`, autre `On Error`, ou fin de procédure). La portée révèle l'ampleur de la suppression.

### Étape 2 — **Classer** : localisé ou en bloc ?

- **Localisé** (OERN + test `Err.Number` immédiat sur une opération) → § 5.3.
- **En bloc** (aucun test, large portée) → § 5.4.

### Étape 3 — Cas **localisé** → `Try`/`Catch` resserré (ou mieux, § 6)

```vb
' VB6
On Error Resume Next
valeur = CInt(texte)
If Err.Number <> 0 Then valeur = 0
On Error GoTo 0
```

```vb
' VB.NET — Try/Catch resserré, à intention identique
Try
    valeur = CInt(texte)
Catch ex As Exception
    valeur = 0
End Try
```

### Étape 4 — Cas **en bloc** → analyser et décider, instruction par instruction

Pour **chaque** opération de la portée, posez la question : *que doit-il arriver si elle échoue ?*

| Décision | Quand | Mise en œuvre |
|---|---|---|
| **Gérer** | L'erreur est attendue et **a un sens** | `Try`/`Catch` **resserré** avec un vrai traitement |
| **Ignorer** | Échec réellement **sans conséquence** (rare) | `Try`/`Catch` **étroit**, `Catch` **commenté** expliquant *pourquoi* |
| **Propager** | L'erreur signale un **vrai problème** | **Ne pas** attraper — laisser l'exception remonter (souvent le bon choix qu'OERN masquait) |

Le résultat : la suppression massive disparaît, remplacée par des `Try`/`Catch` **ciblés** là où c'est justifié, et par de la **propagation** ailleurs.

### Étape 5 — **Valider** au golden master

Parce que supprimer **comme** envelopper modifient des comportements subtils, comparez l'avant/après avec un **[harnais de tests de référence](../05-preparer-code-vb6/05-golden-master.md)** et des **[tests de non-régression](../17-valider-refactoriser/03-non-regression.md)**. Chaque divergence (une erreur autrefois avalée qui « remonte » soudain) doit être **investiguée**, pas masquée à nouveau (**[§ 17.4](../17-valider-refactoriser/04-traquer-pieges.md)**).

---

## 6. Le bon réflexe moderne : **éviter** l'exception

Beaucoup d'OERN localisés ne servaient qu'à **tenter une conversion** et gérer l'échec. En .NET, de nombreuses API offrent une variante **non levante** (`TryParse`, `TryGetValue`…) qui renvoie un **booléen de succès** au lieu de lever une exception :

```vb
' VB.NET — API non-levante : aucune exception pour un cas ATTENDU
If Not Integer.TryParse(texte, valeur) Then
    valeur = 0
End If
```

C'est **préférable** au `Try`/`Catch` pour un échec **courant et attendu** : plus lisible, et sans le **coût** d'une exception (lever/attraper est une opération onéreuse, à réserver aux cas réellement **exceptionnels**). Quand une telle API existe, **privilégiez-la**.

---

## 7. Ce que font (et ne font pas) les outils

- **Les assistants** laissent souvent `On Error Resume Next` **intact** (VB.NET le supporte) ou produisent un enveloppement générique. Dans les deux cas, ce **n'est pas** une migration terminée : c'est une **dette** à traiter à la main.
- **Aucun outil ne comprend l'intention** derrière un OERN : était-ce une vérification voulue, ou un masquage paresseux ? Seul un humain tranche.
- **L'IA** peut **aider** à repérer les OERN, proposer des `Try`/`Catch` ciblés ou des `TryParse`, et signaler les erreurs supprimées — mais la **décision** (gérer / ignorer / propager) reste **humaine** (**[§ 19.3](../19-migration-ia/03-detecter-pieges.md)**).

---

## 📋 Tableau de synthèse

| Aspect | Réalité |
|---|---|
| Sémantique d'OERN | Ignorer l'erreur, **continuer à l'instruction suivante**, garder l'état |
| Équivalent `Try`/`Catch` | **Aucun** (le `Try` s'arrête à la 1ʳᵉ erreur) |
| Gros `Try` + `Catch` vide | ❌ Comportement **différent** **et** masquage des bugs |
| Compile en VB.NET ? | **Oui** — d'où le piège : il survit en silence |
| Visage « localisé » | OERN + test `Err.Number` → `Try`/`Catch` resserré, ou `TryParse` |
| Visage « en bloc » | Aucun test, large portée → **analyser et décider** par opération |
| Décisions possibles | **Gérer** / **Ignorer** (rare, commenté) / **Propager** (souvent le bon choix) |
| Validation | **Golden master** + non-régression, obligatoire |

---

## ✅ À retenir

1. **OERN n'a pas d'équivalent** : un `Try` **abandonne** à la première erreur, OERN **poursuit**. Le seul équivalent littéral (un `Try` par instruction) est absurde.
2. **Un gros `Try` à `Catch` vide est doublement faux** : il change le comportement (arrêt anticipé) **et** masque les bugs.
3. **Le vrai piège : OERN compile en VB.NET.** Le laisser, c'est **figer** la dette ; le supprimer, c'est risquer des **plantages**. Aucune voie automatique n'est correcte.
4. **Diagnostiquez avant de migrer** : usage **localisé** (→ `Try`/`Catch` resserré ou `TryParse`) vs **suppression en bloc** (→ décision **par opération** : gérer / ignorer / propager).
5. **Préférez éviter l'exception** (`TryParse` & co) pour les échecs attendus, et **validez tout au golden master** : une erreur qui « remonte » soudain est à **comprendre**, pas à re-masquer.

---

*Section suivante : **[§ 11.4 — L'objet `Err` vs `Exception` ; `Err.Raise` → `Throw`](04-err-vs-exception.md)***

**Cible** : .NET Framework 4.7.2 · **Indicateurs** : ⭐ point clé · ⚠️ piège · 🛠️ outils

⏭️ [L'objet `Err` vs `Exception` ; `Err.Raise` → `Throw`](/11-gestion-erreurs/04-err-vs-exception.md)
