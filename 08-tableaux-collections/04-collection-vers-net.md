🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.4 — L'objet `Collection` VB6 → collections .NET (`List(Of T)`, `Dictionary(Of K,V)`)

> **Chapitre 8 — Tableaux et collections** · Partie 3 — Migrer le langage
>
> L'objet `Collection` de VB6 — non typé, **1-based**, avec des clés `String` optionnelles — se
> migre vers les **génériques** de .NET : **`List(Of T)`** quand il n'y avait pas de clé,
> **`Dictionary(Of K,V)`** quand il y en avait. À la clé, plusieurs changements **silencieux** : le
> passage **1-based → 0-based**, la **casse des clés**, et la disparition de l'**ordre garanti** côté
> dictionnaire.

---

## 🧭 VB6 : l'objet `Collection`

```vb
' VB6 — l'objet Collection
Dim c As New Collection
c.Add "Alice"                 ' sans clé
c.Add "Bob", "b"              ' avec clé "b"
Debug.Print c(1)              ' "Alice"  ⚠️ 1-based !
Debug.Print c("b")            ' "Bob"    (accès par clé)
Debug.Print c.Count           ' 2
c.Remove 1                    ' retire par index (ou par clé)

' Tester l'existence d'une clé : par GESTION D'ERREUR (aucun autre moyen)
Dim v As Variant
On Error Resume Next
v = c("x")
If Err.Number <> 0 Then       ' clé absente
On Error GoTo 0
```

Caractéristiques à retenir :

- **1-based** : `c(1)` est le **premier** élément.
- **Clés** `String`, **optionnelles**, **insensibles à la casse**, sans doublon (sinon erreur).
- **Éléments non typés** (`Variant`) → hétérogènes.
- **Ordonnée** : l'index reflète l'ordre d'insertion (avec positionnement `Before`/`After`).
- **Aucun test d'existence** propre : il fallait **tenter l'accès** et **intercepter l'erreur**.

---

## 🔀 Le choix de migration : `List` ou `Dictionary` ?

| Usage VB6 de `Collection` | Cible .NET recommandée |
|---------------------------|------------------------|
| Liste **ordonnée sans clé** | **`List(Of T)`** (typée, 0-based) |
| Accès **par clé** (`String`) | **`Dictionary(Of String, V)`** |
| **Clé ET ordre** préservé | `OrderedDictionary` (non générique) **ou** `List` + `Dictionary` |
| **Fidélité maximale** (1-based, non typé) | `Microsoft.VisualBasic.Collection` *(béquille)* |

> 💡 **Option de fidélité.** L'objet `Collection` **existe encore** en VB.NET
> (`Microsoft.VisualBasic.Collection`) : il reste **1-based** et non typé, donc le code 1-based
> **compile et tourne sans changement**. C'est le pendant de la stratégie « garder le 1-based » des
> tableaux (8.1) — pratique au départ, mais à **refactoriser** ensuite vers des génériques (dépendance
> de compatibilité, aucune sûreté de type).

---

## 📋 `List(Of T)` — sans clé, typée, 0-based

```vb
' Sans clé → List(Of T)
Dim noms As New List(Of String)
noms.Add("Alice")
noms.Add("Bob")
Console.WriteLine(noms(0))     ' "Alice"  ⚠️ 0-based !
Console.WriteLine(noms.Count)  ' 2
noms.Insert(0, "Zoé")          ' équivaut au Before/After de VB6
noms.RemoveAt(0)               ' retirer par INDICE
```

> 🪤 **Deux pièges :**
> - **1-based → 0-based** : `c(1)` (VB6) devient `noms(0)` (.NET). Tout accès indexé **se décale**.
> - **`Remove` retire par VALEUR, pas par indice** : `noms.Remove("Bob")` retire l'élément **égal à
>   "Bob"** ; pour retirer **l'élément n°i**, c'est **`RemoveAt(i)`**. Migrer `c.Remove(i)` en
>   `Remove(i)` est un bug classique → utiliser `RemoveAt`.

---

## 🔑 `Dictionary(Of K, V)` — accès par clé

```vb
' Accès par clé → Dictionary(Of K, V)
Dim annuaire As New Dictionary(Of String, String)
annuaire.Add("b", "Bob")
annuaire("c") = "Carol"        ' ajoute OU met à jour (l'indexeur ne lève pas si nouvelle clé)

' Existence : API claire (fini la gestion d'erreur)
If annuaire.ContainsKey("b") Then ...
Dim nom As String = Nothing
If annuaire.TryGetValue("x", nom) Then ...   ' False si absent, SANS exception
```

Trois différences de comportement à connaître :

- **Test d'existence** : `ContainsKey`/`TryGetValue` **remplacent l'idiome `On Error`** de VB6. Plus
  propre, et lié à la migration de la gestion d'erreurs (**chapitre 11**).
- **Clé absente / en double** : l'**indexeur** sur une clé absente lève **`KeyNotFoundException`** ;
  `Add` d'une clé **déjà présente** lève **`ArgumentException`**. (D'où l'usage de `TryGetValue` et de
  l'indexeur `dict(clé) = valeur` pour « ajouter ou mettre à jour ».)
- **Sensibilité à la casse** ⚠️ : un `Dictionary(Of String, …)` est **ordinal (sensible à la casse)**
  par défaut, alors que les clés de la `Collection` VB6 étaient **insensibles**. Pour retrouver le
  comportement VB6 :

  ```vb
  Dim annuaire2 As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
  ```

> 🪤 **L'ordre n'est pas garanti.** La `Collection` VB6 était **ordonnée** ; un `Dictionary` **ne
> garantit pas** l'ordre d'insertion. Si l'on a besoin **à la fois** de l'accès par clé **et** de
> l'ordre, sur **.NET Framework 4.7.2** il n'existe pas de dictionnaire ordonné *générique* intégré :
> utiliser `OrderedDictionary` (espace `System.Collections.Specialized`, **non générique** → clés et
> valeurs en `Object`), ou maintenir une **`List(Of T)` en parallèle** d'un `Dictionary`, ou une
> petite **classe dédiée**.

---

## 🔁 Équivalence des opérations

| `Collection` VB6 | `List(Of T)` | `Dictionary(Of K,V)` |
|------------------|--------------|----------------------|
| `c.Add item` | `list.Add(item)` | `dict.Add(key, value)` |
| `c(i)` *(1-based)* | `list(i)` *(0-based)* | `dict(key)` |
| `c.Remove(i / clé)` | `list.RemoveAt(i)` / `list.Remove(item)` | `dict.Remove(key)` |
| `c.Count` | `list.Count` | `dict.Count` |
| existence *(via `On Error`)* | `list.Contains(item)` | `dict.ContainsKey(key)` / `TryGetValue` |
| `.Add …, Before/After` | `list.Insert(i, item)` | *(pas de position)* |
| `For Each x In c` | `For Each x In list` | `For Each kv In dict` *(`KeyValuePair`)* |

> 💡 L'énumération d'un `Dictionary` parcourt des **`KeyValuePair(Of K,V)`** (`kv.Key`, `kv.Value`),
> ou ses **`.Keys`** / **`.Values`**. Détails et pièges de parcours en **8.5**.

---

## ⚠️ Pièges silencieux — récapitulatif

- **`Collection` (1-based) → `List` (0-based)** : tout accès indexé **se décale** d'un cran.
- **`List.Remove` retire par VALEUR** ; pour l'indice, **`RemoveAt`**.
- **Clés `Dictionary` sensibles à la casse** par défaut (VB6 = insensible) → `StringComparer.OrdinalIgnoreCase`.
- **`Dictionary` non ordonné** : si l'ordre comptait, ne pas s'y fier (4.7.2 : `OrderedDictionary`
  non générique, ou `List`+`Dictionary`).
- **Clé absente/en double** : `KeyNotFoundException` / `ArgumentException` → préférer `TryGetValue`
  et l'indexeur.
- **`Microsoft.VisualBasic.Collection`** reste **1-based** et non typée : béquille à retirer.

---

## 🧭 Stratégie de migration

1. **Déterminer l'usage** de chaque `Collection` : **avec ou sans clé** ? **ordre** significatif ?
2. **Sans clé → `List(Of T)`** ; **avec clé → `Dictionary(Of String, V)`** ; **clé + ordre →**
   `OrderedDictionary` ou `List`+`Dictionary`.
3. **Décaler les indices** 1→0 (ou conserver la `Collection` VB le temps de stabiliser).
4. **Remplacer l'idiome `On Error`** de test d'existence par `ContainsKey`/`TryGetValue`.
5. **Aligner la casse des clés** (`StringComparer.OrdinalIgnoreCase`) pour préserver le comportement
   VB6.
6. **Vérifier `Remove`** : index → `RemoveAt`, valeur/clé → `Remove`.
7. **Filet de sécurité** : le ***golden master*** (chapitre 5.5) révèle un décalage d'indice, une
   clé introuvable par différence de casse, ou un ordre perdu.

---

## ✅ À retenir

- **`Collection` → `List(Of T)`** (sans clé) **ou `Dictionary(Of K,V)`** (avec clé) ; la `Collection`
  VB **subsiste** (1-based) pour une migration fidèle, à refactoriser ensuite.
- **1-based → 0-based** : le décalage d'indice est le piège principal.
- **`Dictionary`** : `ContainsKey`/`TryGetValue` remplacent l'`On Error` ; clés **sensibles à la
  casse** par défaut (→ `OrdinalIgnoreCase`) ; **ordre non garanti**.
- **`List.RemoveAt`** (indice) **≠ `List.Remove`** (valeur).

---

## 🔗 Pour aller plus loin

- **8.1 — Tableaux 0-based** : la même bascule d'indexation (et la stratégie « fidélité vs idiomatique »).
- **8.2 — `ReDim`** : `List(Of T)` comme alternative aux tableaux qui grandissent.
- **8.3 — Tableaux de `Variant`** : du non typé vers les génériques.
- **8.5 — `For Each` et énumération** : parcourir listes et dictionnaires (et le piège de la
  modification pendant le parcours).
- **Chapitre 11 — Gestion des erreurs** : remplacer l'`On Error` du test d'existence.
- **Chapitre 17.5 — Refactoring idiomatique** : génériques, `LINQ`, suppression des béquilles.
- **Annexe D** — correspondance des types de données.

➡️ Section suivante : **[8.5 — `For Each` et énumération](05-foreach.md)**  
⬆️ Retour au **[sommaire du chapitre 8](README.md)**

⏭️ [`For Each` et énumération (différences de comportement)](/08-tableaux-collections/05-foreach.md)
