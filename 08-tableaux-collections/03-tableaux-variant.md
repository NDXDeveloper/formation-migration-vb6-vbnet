🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.3 — Tableaux de `Variant` / d'objets

> **Chapitre 8 — Tableaux et collections** · Partie 3 — Migrer le langage
>
> Un tableau de `Variant` migre vers un tableau d'**`Object`** — avec tout ce que 7.1 a montré :
> *boxing* des types valeur, casts à la relecture, perte de sûreté. Cette section prolonge ce thème
> dans le contexte des tableaux, distingue **deux notions VB6** souvent confondues, et porte un
> message simple : **dès que les données sont homogènes, typer le tableau**.

---

## 🧭 VB6 : deux notions distinctes

Attention à ne pas confondre **un tableau de `Variant`** et **un `Variant` contenant un tableau** :

```vb
' VB6 — DEUX notions différentes

' (1) Tableau de Variant : un TABLEAU, chaque case est un Variant
Dim a(2) As Variant
a(0) = 42 : a(1) = "texte" : a(2) = #1/1/2026#

' (2) Variant CONTENANT un tableau : un SEUL Variant dont la valeur est un tableau
Dim v As Variant
v = Array(10, 20, 30)        ' la fonction Array() renvoie un Variant-tableau
Debug.Print IsArray(v)       ' True
Debug.Print v(1)             ' 20
```

Le tableau de `Variant` est aussi le **type par défaut** : `Dim a(10)` sans `As` est un tableau de
`Variant`. On l'utilisait pour des données **hétérogènes**, des retours multiples, des conteneurs
génériques, et l'**interop COM** (`SAFEARRAY` de `VARIANT`).

---

## 🎯 VB.NET : tableau d'`Object`

`Variant` devient `Object` (cf. 7.1) : un tableau d'`Object`, chaque case une **référence**.

```vb
' VB.NET — tableau d'Object
Dim a(2) As Object
a(0) = 42                    ' ⚠️ boxing : l'Integer est emballé sur le tas
a(1) = "texte"
a(2) = #1/1/2026#

' Relecture : il faut CASTER (Option Strict On)
Dim n As Integer = DirectCast(a(0), Integer)
```

> ⚠️ **Deux coûts hérités de 7.1 :**
> - **Boxing** : chaque type valeur rangé dans une case `Object` est copié sur le tas → pression sur
>   le *garbage collector*, parcours numériques lents.
> - **Sûreté perdue** : pour relire une case, il faut **caster** (et, si le contenu est mélangé,
>   tester le type avec `TypeOf`/`GetType`).

---

## 🎯 Le bon réflexe : **typer** le tableau

C'est le message principal de cette section. Selon la nature **réelle** des données :

```vb
' Données HOMOGÈNES → tableau typé (sûr, AUCUN boxing)
Dim nombres() As Integer = {42, 17, 8}

' …ou une collection générique
Dim noms As New List(Of String) From {"Alice", "Bob"}

' HÉTÉROGÈNE mais STRUCTURÉ → classe/structure, ou tuple nommé
Dim ligne = (Id:=42, Nom:="texte", LeJour:=#1/1/2026#)   ' tuple nommé
```

> 🎯 Ne conserver **`Object()`** que là où les données sont **vraiment hétérogènes** ou requises
> pour l'**interop**/la **liaison tardive**. Pour des données mélangées mais **structurées**, une
> **classe**, une **structure** ou un **tuple** exprime bien mieux l'intention qu'un fourre-tout
> `Object`.

---

## 🧩 « `Variant` contenant un tableau » → `Object` détenant un tableau

La seconde notion VB6 se traduit par un `Object` qui **référence** un tableau. Pour l'indexer sous
`Option Strict On`, il faut **caster** vers le type de tableau réel :

```vb
' VB.NET — un Object détenant un tableau
Dim o As Object = {10, 20, 30}       ' o référence un tableau d'Integer (une référence, pas de boxing)
If TypeOf o Is Array Then
    Dim arr() As Integer = DirectCast(o, Integer())   ' caster pour indexer
    Debug.Print(arr(1))               ' 20
End If
```

Et les outils VB6 associés évoluent :

| VB6 | VB.NET |
|-----|--------|
| `Array(1, 2, 3)` (renvoie un Variant-tableau) | initialiseur **`{1, 2, 3}`** (ou `New Object() {1, 2, 3}`) |
| `IsArray(v)` | `TypeOf o Is Array` (ou `IsArray`, encore présent) |
| accès `v(i)` sur un Variant-tableau | caster d'abord : `DirectCast(o, T())(i)` |

---

## 🪜 Multidimensionnel **vs** dentelé : une clarification utile

VB6 simulait parfois des « tableaux de tableaux » via des `Variant` contenant eux-mêmes des  
tableaux. .NET offre **deux structures distinctes**, à ne pas confondre :

```vb
' Multidimensionnel (RECTANGULAIRE) — un seul objet, dimensions fixes
Dim grille(2, 3) As Integer          ' accès : grille(i, j)

' Dentelé (tableau de tableaux) — chaque sous-tableau de taille LIBRE
Dim lignes()() As Integer            ' accès : lignes(i)(j)
lignes = New Integer(1)() {}
lignes(0) = New Integer() {1, 2, 3}
lignes(1) = New Integer() {4, 5}
```

> 💡 Un `Variant` VB6 dont chaque case contenait un tableau correspond naturellement à un **tableau
> dentelé** `()()` en .NET. Attention à la syntaxe : `a(i, j)` (rectangulaire) **≠** `a(i)(j)`
> (dentelé).

---

## 🔗 Cas connexes (renvois)

- **`ParamArray … As Variant`** (arguments variables) → **`ParamArray … As Object()`** : détaillé au
  **chapitre 10.4** (avec le passage `ByRef` VB6 → tableau `ByVal` en .NET).
- **`SAFEARRAY` de `VARIANT`** franchissant une frontière **COM** → tableaux d'`Object` *marshalés* :
  voir **module 16** (interop COM et appels d'API).

---

## ⚠️ Pièges silencieux — récapitulatif

- **`Object()` = boxing + casts** : un tableau d'`Object` de nombres est lent et non typé → préférer
  un **tableau typé**.
- **Deux notions à distinguer** : *tableau de `Variant`* (`Object()`) **vs** *`Variant` contenant un
  tableau* (`Object` détenant un tableau).
- **`Array()` disparaît** au profit des **initialiseurs `{…}`** ; `IsArray` → `TypeOf … Is Array`.
- **Rectangulaire `a(i, j)` ≠ dentelé `a(i)(j)`** : ne pas confondre les deux structures.
- Tout reste **0-based** (cf. 8.1) et l'**affectation aliase** (cf. 8.2).

---

## 🧭 Stratégie de migration

1. **Classer chaque tableau de `Variant`** : homogène ou hétérogène ?
2. **Homogène → typer** (`Integer()`, `String()`…) ou **`List(Of T)`** : on gagne la sûreté et on
   supprime le *boxing*.
3. **Hétérogène mais structuré → classe/structure/tuple** plutôt qu'`Object()`.
4. **Vraiment générique/interop → `Object()`**, avec casts explicites (`DirectCast`/`CType`) et tests
   de type (`TypeOf`).
5. **Distinguer rectangulaire et dentelé** lors de la traduction des « tableaux de tableaux ».
6. **Filet de sécurité** : le ***golden master*** (chapitre 5.5) révèle un cast manquant, une case
   `Object` mal typée ou une confusion d'accès `(i, j)` / `(i)(j)`.

---

## ✅ À retenir

- **Tableau de `Variant` → tableau d'`Object`** : *boxing* et casts (cf. 7.1) ; le réflexe est de
  **typer** dès que c'est possible.
- **Deux notions VB6** distinctes (tableau de `Variant` vs `Variant`-tableau) → `Object()` **vs**
  `Object` détenant un tableau.
- **`Array()` → initialiseurs `{…}`** ; **`IsArray` → `TypeOf … Is Array`**.
- .NET distingue **rectangulaire** (`a(i, j)`) et **dentelé** (`a(i)(j)`) — un VB6 « Variant de
  tableaux » devient souvent un **tableau dentelé**.

---

## 🔗 Pour aller plus loin

- **7.1 — `Variant` → `Object`** : *boxing*, `CType`/`DirectCast`, `Option Strict`.
- **8.1 — Tableaux 0-based** et **8.2 — `ReDim`** : bornes, indexation et alias.
- **8.4 — `Collection` → collections .NET** : `List(Of T)`, `Dictionary(Of K,V)`.
- **Chapitre 10.4 — `ParamArray`** ; **module 16** — interop COM (`SAFEARRAY`/`VARIANT`).
- **Chapitre 17.5 — Refactoring idiomatique** : du fourre-tout `Object` aux génériques et `LINQ`.
- **Annexe D** — correspondance des types de données.

➡️ Section suivante : **[8.4 — L'objet `Collection` → collections .NET](04-collection-vers-net.md)**  
⬆️ Retour au **[sommaire du chapitre 8](README.md)**

⏭️ [L'objet `Collection` VB6 → collections .NET (`List(Of T)`, `Dictionary(Of K,V)`)](/08-tableaux-collections/04-collection-vers-net.md)
