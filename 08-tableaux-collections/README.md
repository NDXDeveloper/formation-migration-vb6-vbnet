🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8. Tableaux et collections ⚠️

> **Partie 3 — Migrer le langage (le cœur de la formation)**
>
> Le chapitre 7 a posé le **socle des types**. Celui-ci s'attaque à la façon dont VB6 **regroupe**
> les données : les **tableaux** et l'objet **`Collection`**. Deux mondes qui changent de règles —
> à commencer par la plus structurante de toutes : en .NET, **les tableaux commencent toujours
> à 0**.

---

## 🧭 Pourquoi ce chapitre mérite de la vigilance

Parce qu'il touche à deux choses omniprésentes et faciles à **décaler silencieusement** :

1. **L'indexation.** VB6 autorisait des bornes basses arbitraires (`Dim a(1 To 10)`) et un
   `Option Base` réglable. .NET impose **0** partout. Un simple décalage d'indice ne plante pas : il
   **lit la mauvaise case** ou **saute un élément**.
2. **Le typage des regroupements.** L'objet `Collection` de VB6 et les tableaux de `Variant`
   n'étaient **pas typés**. .NET pousse vers des **génériques typés** (`List(Of T)`,
   `Dictionary(Of K,V)`) — un gain de sûreté, mais une **réécriture** à mener avec méthode.

> ⚠️ Comme partout dans cette partie, **le danger n'est pas le code qui ne compile pas — c'est le
> code qui compile mais ne parcourt plus les mêmes éléments** (voir **Annexe B.4 — Tableaux
> 0-based et `Option Base`**).

---

## 🧨 Le piège fondateur : `Dim a(10)` n'a pas le même nombre d'éléments

S'il ne fallait retenir qu'une chose : **en .NET, un tableau est toujours 0-based**, et `Dim a(N)`  
déclare des indices **0 à N**, soit **N+1 éléments**.

| Déclaration | En **VB6** | En **VB.NET** |
|-------------|------------|---------------|
| `Dim a(10)` *(Option Base 0, défaut)* | 11 éléments (0..10) | **11 éléments (0..10)** |
| `Dim a(10)` *(Option Base 1)* | **10 éléments (1..10)** | **11 éléments (0..10)** ⚠️ |
| `Dim a(1 To 10)` | 10 éléments (1..10) | **erreur de compilation** |

```vb
' VB.NET : la borne basse ne peut être que 0
Dim a(10)        ' indices 0..10  → 11 éléments (toujours)
Dim b(1 To 10)   ' ❌ erreur : « les limites inférieures ne peuvent être que '0' »
```

> 🔎 Conséquence : le code écrit sous **`Option Base 1`** (ou avec `1 To n`) **décale d'un cran** une
> fois migré. C'est le sujet — central — de la section **8.1** (signalée ⭐ ⚠️), avec `LBound`/
> `UBound` et la disparition d'`Option Base`.

---

## 🗺️ Les changements, en bref

- **Bornes** : plus d'`Option Base`, plus de `1 To n` — **tout est 0-based** (8.1).
- **Tableaux dynamiques** : `ReDim`/`ReDim Preserve` **existent encore**, mais redimensionner en
  boucle est coûteux → souvent un signal pour passer à **`List(Of T)`** (8.2).
- **Tableaux de `Variant`** → tableaux d'**`Object`** (avec *boxing*, cf. 7.1) ou, mieux, **typés**
  (8.3).
- **`Collection` (1-based !)** → **`List(Of T)`** / **`Dictionary(Of K,V)`** (0-based, typés) (8.4).
- **Énumération** : `For Each` subsiste, mais **modifier une collection pendant le parcours** lève
  désormais une exception (8.5).

---

## 📚 Ce que couvre ce chapitre

- **[8.1 — Tableaux toujours 0-based](01-tableaux-0-based.md).** ⭐ ⚠️ `Option Base` supprimé, bornes
  non nulles interdites, `LBound`/`UBound`, et le décalage d'indice à traquer.
- **[8.2 — `ReDim`, `ReDim Preserve`, tableaux dynamiques](02-redim.md).** Ce qui subsiste, le coût
  caché du redimensionnement, et quand basculer vers `List(Of T)`.
- **[8.3 — Tableaux de `Variant` / d'objets](03-tableaux-variant.md).** Vers `Object()` ou des
  tableaux **typés** ; rappels de *boxing* (cf. 7.1).
- **[8.4 — L'objet `Collection` → collections .NET](04-collection-vers-net.md).** `Collection`
  (1-based) vers **`List(Of T)`** et **`Dictionary(Of K,V)`** (génériques, 0-based).
- **[8.5 — `For Each` et énumération](05-foreach.md).** Différences de comportement, dont la
  modification pendant le parcours.

---

## 🧾 Aide-mémoire express

| Sujet | VB6 | VB.NET |
|-------|-----|--------|
| Borne basse d'un tableau | variable (`Option Base`, `1 To n`) | **toujours 0** |
| `Dim a(10)` | 11 él. (base 0) ou 10 él. (base 1) | **11 éléments (0..10)**, toujours |
| `Option Base` | disponible | **supprimé** |
| Redimensionner | `ReDim` / `ReDim Preserve` | idem (mais préférer **`List(Of T)`**) |
| Tableau de `Variant` | `As Variant` | `Object()` (ou **typé**) |
| `Collection` | objet **1-based** | **`List(Of T)`** / **`Dictionary`** (0-based) |
| Parcourir | `For Each` | `For Each` (**ne pas modifier** pendant) |

---

## 🎯 Objectifs du chapitre

À l'issue de ce chapitre, tu sauras :

- **convertir** toute déclaration de tableau vers le modèle **0-based**, en repérant les décalages
  hérités d'`Option Base` ou de `1 To n` ;
- **décider** entre tableau et **collection générique** (`List(Of T)`, `Dictionary(Of K,V)`) selon
  l'usage (taille fixe vs croissante, accès par clé) ;
- **migrer l'objet `Collection`** en tenant compte du passage **1-based → 0-based** ;
- **parcourir** des données sans tomber dans le piège de la **modification pendant l'énumération**.

---

## 🔗 À garder ouvert en parallèle

- **Annexe B.4 — Tableaux 0-based et `Option Base`** : le piège d'indexation, symptôme/cause/correction.
- **Chapitre 7 — Types** : le *boxing* (7.1) et l'« objet nul » (7.7) reviennent pour les tableaux
  d'`Object` et les collections.
- **Chapitre 17.5 — Refactoring idiomatique** : remplacer durablement tableaux dynamiques et
  `Collection` par des génériques et `LINQ`.

> 💡 Et toujours le **golden master** (chapitre 5.5) : un décalage d'indice ou un élément sauté se
> voit immédiatement dans la comparaison des sorties.

---

## ➡️ Pour commencer

On attaque par le changement le plus structurant — et le plus piégeux — du chapitre : les tableaux
**toujours 0-based**, la disparition d'`Option Base` et le sort de `LBound`/`UBound`.

➡️ **[8.1 — Tableaux toujours 0-based : `Option Base`, bornes non nulles, `LBound`/`UBound`](01-tableaux-0-based.md)**  
⬆️ Retour au **[sommaire général](../SOMMAIRE.md)**

⏭️ [Tableaux **toujours 0-based** : `Option Base`, bornes non nulles, `LBound`/`UBound`](/08-tableaux-collections/01-tableaux-0-based.md)
