🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.1 — Tableaux toujours 0-based : `Option Base`, bornes non nulles, `LBound`/`UBound` ⭐ ⚠️

> **Chapitre 8 — Tableaux et collections** · Partie 3 — Migrer le langage
>
> VB6 laissait choisir où commençaient les tableaux : `Option Base 0` ou `1`, voire des bornes
> arbitraires (`Dim a(5 To 15)`). **.NET ne laisse aucun choix : tout tableau commence à 0.** Le
> piège est qu'un décalage d'indice **ne plante pas** — il lit la mauvaise case, ajoute un élément
> fantôme ou en saute un. C'est l'entrée **B.4** du catalogue des pièges silencieux.

---

## 🧭 VB6 : des bornes libres

En VB6, la **borne basse** d'un tableau est paramétrable :

```vb
' VB6 — bornes libres
Option Base 1               ' au niveau module : borne basse par défaut = 1
Dim a(10)                   ' indices 1..10  → 10 éléments
Dim b(1 To 10)              ' indices 1..10  → 10 éléments
Dim c(5 To 15)              ' indices 5..15  → 11 éléments
Dim m(1 To 3, 1 To 4)       ' tableau 2D, 1-based sur les deux dimensions

For i = LBound(a) To UBound(a)   ' parcours indépendant des bornes
    ' ...
Next
```

- **`Option Base 0|1`** fixe, au niveau **module**, la borne basse par défaut des tableaux déclarés
  **sans** bornes explicites (par défaut, `0`).
- **`Dim a(N)`** déclare des indices **borne_basse..N**.
- **`LBound(a)`** renvoie la borne basse (0, 1, 5…), **`UBound(a)`** la borne haute.
- **Nombre d'éléments** = `UBound − LBound + 1`.

---

## 🎯 VB.NET : toujours 0-based

```vb
' VB.NET — TOUJOURS 0-based
Dim a(10)        ' indices 0..10  → 11 éléments
Dim b(0 To 10)   ' autorisé (borne basse = 0), strictement identique à Dim a(10)
Dim c(1 To 10)   ' ❌ erreur : « les limites inférieures ne peuvent être que '0' »
' Option Base 1  ' ❌ instruction supprimée

Console.WriteLine(LBound(a))   ' 0   (toujours)
Console.WriteLine(UBound(a))   ' 10
Console.WriteLine(a.Length)    ' 11
```

- **La borne basse est toujours 0.** Toute borne non nulle (`1 To 10`, `5 To 15`) est une **erreur
  de compilation**.
- **`Dim a(N)`** déclare des indices **0..N**, soit **N+1 éléments**. Le nombre entre parenthèses est
  la **borne haute**, **pas** le nombre d'éléments.
- **`Option Base` est supprimé.**
- **`LBound(a)` renvoie toujours 0** (et `UBound`/`LBound` existent encore, dans
  `Microsoft.VisualBasic`).

> 💡 La forme `Dim a(0 To 10)` est **acceptée** (utile pour la lisibilité), mais uniquement avec une
> borne basse **égale à 0**.

---

## 🧨 Le piège central : l'élément fantôme et le décalage d'un cran

Le code écrit sous **`Option Base 1`** (ou avec `1 To n`) change de **nombre d'éléments** une fois  
migré. L'indice **0**, qui n'existait pas, **apparaît**.

```vb
' VB6 sous Option Base 1
Dim notes(10)            ' 10 éléments (1..10)
For i = 1 To 10
    total = total + notes(i)
Next

' Migration NAÏVE en VB.NET
Dim notes(10)            ' 11 éléments (0..10) — l'indice 0 existe désormais !
For i = 1 To 10          ' saute notes(0) (qui vaut 0) → cette boucle-ci reste juste…
    total = total + notes(i)
Next

' …mais ailleurs, tout parcours 0-based bascule :
For i = 0 To UBound(notes)   ' inclut notes(0) fantôme → total faussé
    ' ...
Next
Console.WriteLine(notes.Length)  ' 11, et non 10 → comptage erroné
```

> ⚠️ **Trois conséquences typiques :**
> - **Comptage faux** : `Length` (ou `UBound + 1`) renvoie **N+1** au lieu de N.
> - **Élément fantôme** : un parcours `0 To UBound` traite l'indice 0 (valeur par défaut **0**,
>   `Nothing` ou `""`) → somme faussée, ou **`NullReferenceException`** sur un tableau d'objets.
> - **`Dim a(1 To 10)` ne compile pas** → réécriture obligatoire.

---

## 🛠️ Deux stratégies de migration

Face à un tableau 1-based, deux approches — au **compromis opposé**.

```vb
' (A) RE-BASER À 0 (idiomatique) : décaler toute l'indexation de −1
Dim notes(9)             ' 10 éléments (0..9)
For i = 0 To 9
    ' ...
Next

' (B) CONSERVER la logique 1-based, gaspiller l'indice 0 (fidélité maximale)
Dim notes(10)            ' 11 éléments ; on n'utilise QUE 1..10
For i = 1 To 10
    ' ... le code d'origine est préservé tel quel
Next
```

| | **(A) Re-baser à 0** | **(B) Garder 1-based, gaspiller l'indice 0** |
|---|---|---|
| **Principe** | décaler chaque indice de −1 | déclarer `(N)` et ignorer l'indice 0 |
| **Fidélité** | sémantique différente (indices changés) | comportement **identique** |
| **Effort** | touche **chaque** accès | minimal |
| **Risque** | élevé (indices calculés, partagés entre modules) | faible |
| **Coût** | aucun gaspillage | une case perdue par tableau |

> 🎯 **Recommandation.** Pour la **fidélité** de la migration, **(B)** est généralement le **choix
> sûr** au départ : on préserve la logique 1-based à l'identique. C'est d'ailleurs ce que font
> souvent les outils automatiques. Le passage en **(A)**, idiomatique, est un **refactoring
> ultérieur** à mener **délibérément**, sous protection de tests (voir **chapitre 17.5**).

---

## 🔁 `LBound`/`UBound` et les méthodes idiomatiques

Bonne nouvelle : le **parcours indépendant des bornes survit** sans modification (puisque `LBound`  
renvoie 0) :

```vb
For i = LBound(a) To UBound(a)   ' 0..N en VB.NET — fonctionne tel quel
    ' ...
Next
```

> 🪤 En revanche, une boucle qui **codait en dur le `1`** (`For i = 1 To UBound(a)`) en supposant un
> tableau 1-based **saute désormais l'indice 0** — où .NET range pourtant le premier élément.

Les équivalents idiomatiques .NET, plus clairs, reposent sur les membres du type `Array` :

| Besoin | VB6 | VB.NET idiomatique |
|--------|-----|--------------------|
| Borne basse | `LBound(a)` | `a.GetLowerBound(0)` (= 0) |
| Borne haute | `UBound(a)` | `a.GetUpperBound(0)` **ou** `a.Length - 1` |
| Nombre d'éléments | `UBound − LBound + 1` | **`a.Length`** |
| Borne haute d'une dimension | `UBound(a, d)` | `a.GetUpperBound(d)` |
| Taille d'une dimension | `UBound(a, d) − LBound(a, d) + 1` | `a.GetLength(d)` |
| Nombre de dimensions | *(aucun)* | `a.Rank` |

---

## 🧩 Tableaux multidimensionnels

Mêmes règles, sur chaque dimension :

```vb
' VB6 : Dim m(1 To 3, 1 To 4)
' VB.NET — re-basé (A)            ' …ou conservé 1-based (B)
Dim m(2, 3)                       ' Dim m(3, 4)  (lignes/colonnes 0 ignorées)
```

`LBound(m, 2)`/`UBound(m, 2)` ciblent la 2ᵉ dimension (la borne basse reste 0) ; en idiomatique,
`m.GetLength(1)` et `m.GetUpperBound(1)`.

---

## 🧱 Cas particuliers

- **Tableau non initialisé** : en VB6, `UBound` d'un tableau dynamique non encore dimensionné
  **lève une erreur**. En VB.NET, un tableau non initialisé vaut **`Nothing`** → `a.Length` (ou
  `UBound(a)`) lève une **`NullReferenceException`**. **Tester `Is Nothing`** avant d'accéder (cf.
  7.7).
- **Tableau vide** (0 élément) : sa borne haute est **−1** (et `Length` = 0). D'où une propriété
  agréable : `For i = 0 To UBound(a)` sur un tableau vide donne `0 To -1` → **zéro itération** (la
  boucle ne s'exécute pas). C'est un argument de plus pour **re-baser les boucles** sur
  `0 To UBound`.
- **Tableaux dans une `Type`/`Structure` ou en interop** : les bornes et le 0-based comptent pour la
  **disposition mémoire**. Un champ `Dim buf(1 To 256) As Byte` se re-base, et côté interop natif on
  fixe la taille avec `<MarshalAs(UnmanagedType.ByValArray, SizeConst:=256)>` (voir **7.4**,
  **12.7** et **16.5**).

---

## ⚠️ Pièges silencieux — récapitulatif

- **`Dim a(N)` = N+1 éléments** (0..N) : le nombre entre parenthèses est la **borne haute**.
- **`Option Base 1` / `1 To n`** : l'indice **0 apparaît** → comptage faux, élément fantôme, ou
  `NullReferenceException` sur des objets ; `1 To n` **ne compile pas**.
- **Boucle codant `1` en dur** : saute le premier élément (rangé en 0 par .NET).
- **`LBound` vaut toujours 0** : les boucles `LBound To UBound` survivent ; les hypothèses 1-based
  non.
- **Tableau `Nothing`** : `Length`/`UBound` lève une `NullReferenceException`.

---

## 🧭 Stratégie de migration

1. **Recenser les tableaux 1-based** : `Option Base 1`, déclarations `(1 To n)`, et toute boucle
   `For i = 1 To …`.
2. **Choisir la stratégie** par tableau : **(B) fidélité** (garder 1-based, gaspiller l'indice 0)
   par défaut ; **(A) re-basage** idiomatique en refactoring ultérieur.
3. **Re-baser les boucles** sur `LBound`/`0 To UBound` plutôt que sur des bornes codées en dur —
   cela gère aussi les tableaux vides.
4. **Protéger les accès** contre les tableaux `Nothing` (cf. 7.7).
5. **Filet de sécurité** : le ***golden master*** (chapitre 5.5) révèle immédiatement un élément
   sauté, un total faussé par l'indice 0, ou un `Length` erroné.

> 🛠️ Les outils de migration (assistant VS, VBUC, VB Migration Partner — module 4.3) appliquent
> généralement la stratégie **(B)** et qualifient les bornes ; **valider** néanmoins les indices
> **calculés** ou **partagés** entre procédures.

---

## ✅ À retenir

- **En .NET, tout tableau est 0-based** : pas d'`Option Base`, bornes non nulles **interdites**.
- **`Dim a(N)` déclare N+1 éléments** (indices 0..N) — `N` est la **borne haute**.
- Le code **1-based** se migre par **fidélité** (garder 1..N, indice 0 inutilisé) ou par
  **re-basage** (idiomatique, à faire sous tests).
- **`LBound` = 0 toujours** : préférer `LBound`/`0 To UBound` aux bornes en dur ; côté idiomatique,
  `a.Length`, `a.GetUpperBound(d)`, `a.GetLength(d)`, `a.Rank`.
- **Attention** aux tableaux `Nothing` (NRE) et au **comptage** (`Length` = N+1).

---

## 🔗 Pour aller plus loin

- **8.2 — `ReDim`, `ReDim Preserve`** : redimensionnement (toujours 0-based) et bascule vers
  `List(Of T)`.
- **8.4 — `Collection` → collections .NET** : attention, l'objet `Collection` reste **1-based**.
- **7.7 — `Nothing`** : tableaux non initialisés et accès protégés.
- **Chapitre 12.7** (`Type` → `Structure`) et **module 16.5** (*marshaling* de tableaux à taille
  fixe).
- **Chapitre 17.5 — Refactoring idiomatique** : re-basage et passage aux génériques/`LINQ`.
- **Annexe B.4** (Tableaux 0-based et `Option Base`) ; **Annexe D** (correspondance des types).

➡️ Section suivante : **[8.2 — `ReDim`, `ReDim Preserve`, tableaux dynamiques](02-redim.md)**  
⬆️ Retour au **[sommaire du chapitre 8](README.md)**

⏭️ [`ReDim`, `ReDim Preserve`, tableaux dynamiques](/08-tableaux-collections/02-redim.md)
