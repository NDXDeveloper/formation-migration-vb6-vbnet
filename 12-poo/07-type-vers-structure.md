🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.7 — `Type…End Type` → `Structure` (sémantique de valeur, attention au comportement) ⚠️

> **Chapitre 12 — Programmation orientée objet** · Section 12.7
> Le conteneur se traduit en une ligne. Le **comportement de valeur**, lui, réserve des surprises — et impose un **vrai choix** : `Structure` ou `Class` ?

---

## 🧭 Périmètre

Les **types utilisateur** (`Type…End Type`) de VB6 — de simples agrégats de données — deviennent des **`Structure`** en VB.NET. Les deux sont des **types valeur** : c'est rassurant… et trompeur. Car .NET applique la sémantique de valeur de façon **plus visible** (et plus piégeuse) qu'en VB6, notamment au **passage de paramètres** et dans les **collections**.

---

## 1. La traduction de base

```vb
' VB6
Public Type Point
    X As Long
    Y As Long
End Type
```

```vb
' VB.NET
Public Structure Point
    Public X As Integer        ' Long → Integer (7.2) ; champ explicitement Public
    Public Y As Integer
End Structure
```

Avec une chaîne de longueur fixe et un type imbriqué :

```vb
' VB6
Public Type Personne
    Nom As String * 30
    Age As Integer
    Position As Point
End Type
```

```vb
' VB.NET
Public Structure Personne
    <VBFixedString(30)> Public Nom As String   ' longueur fixe (voir 7.4)
    Public Age As Short                        ' Integer 16 bits → Short (7.2)
    Public Position As Point
End Structure
```

Quelques points sur le conteneur :

- **Déclarez les champs explicitement `Public`.** Détail amusant : dans une `Structure`, un champ `Dim` est **`Public`** par défaut (à l'inverse d'une **classe**, où `Dim` est `Private`, voir 12.6). Restez explicite.
- **Plus de souplesse de déclaration.** En VB6, un `Public Type` ne pouvait vivre que dans un **module standard** (`.bas`) ; une classe ne pouvait déclarer que des `Type` **privés**. En VB.NET, une `Structure` se déclare au niveau d'un espace de noms ou **imbriquée** dans une classe/un module.
- `<VBFixedString>` provient de l'espace `Microsoft.VisualBasic`, **importé par défaut**.

---

## 2. Sémantique de valeur : identique en théorie, **piégeuse** en pratique ⚠️

VB6 et VB.NET **copient tous deux** à l'affectation : `b = a` duplique les champs, et modifier `b` n'affecte pas `a`. Jusque-là, **cohérent**. Les écarts surgissent ailleurs.

### 2.1 Passage de paramètre : la copie qui fait disparaître la modification

En **VB6**, un type utilisateur est **toujours passé `ByRef`** : une procédure qui le modifie **répercute** le changement chez l'appelant. En **VB.NET**, le défaut est **`ByVal`** (rappel du grand changement, **10.2**) : pour une `Structure`, c'est une **copie** qui est passée — la modification **ne remonte pas**.

```vb
' VB6 — la UDT est TOUJOURS ByRef : la modification remonte
Sub Deplacer(p As Point)            ' implicitement ByRef
    p.X = p.X + 1
End Sub
' Après l'appel, le Point de l'appelant EST modifié.
```

```vb
' VB.NET — traduction NAÏVE : ByVal par défaut → COPIE → modification PERDUE
Sub Deplacer(p As Point)            ' implicitement ByVal !
    p.X = p.X + 1                   ' modifie une copie locale
End Sub
' Après l'appel, le Point de l'appelant N'EST PAS modifié.  ⚠️

' Correction : passer ByRef explicitement si la mutation doit remonter
Sub Deplacer(ByRef p As Point)
    p.X = p.X + 1
End Sub
```

> ⚠️ C'est l'un des **changements silencieux** les plus traîtres : ça compile, ça tourne, mais le comportement diffère. À croiser avec **10.2** (ByRef → ByVal) et **10.4** (`ParamArray`).

### 2.2 Type valeur dans une collection : la modification fantôme

Indexer une `List(Of T)` de structures **renvoie une copie**. Modifier un champ sur cette copie ne change **rien** dans la liste — et, contrairement à C#, **VB.NET laisse compiler** ce code, rendant le bug **silencieux**.

```vb
' VB.NET — piège du type valeur dans une List
Dim liste As New List(Of Point)()
liste.Add(New Point With {.X = 0, .Y = 0})

liste(0).X = 5                      ' ⚠️ modifie une COPIE temporaire : sans effet
Console.WriteLine(liste(0).X)       ' affiche 0, pas 5

' Correction : remplacer l'élément entier
Dim p = liste(0)
p.X = 5
liste(0) = p
Console.WriteLine(liste(0).X)       ' affiche 5
```

> ⚠️ **Les tableaux se comportent différemment.** Indexer un **tableau** renvoie l'élément **lui-même** (modifiable en place) :
>
> ```vb
> Dim tab(0) As Point
> tab(0).X = 5            ' ✅ modifie bien l'élément du tableau
> Console.WriteLine(tab(0).X)   ' affiche 5
> ```
>
> Cette **asymétrie** (`List` copie, `Array` non) n'existait pas dans l'univers VB6 et **surprend** systématiquement.

---

## 3. Le vrai choix de migration : `Structure` ou `Class` ?

C'est **la** décision de cette section. Tout `Type` VB6 ne devient pas forcément une `Structure` : selon l'usage, une **`Class`** (sémantique de **référence**) évite les surprises ci-dessus.

| Indice | Plutôt **`Structure`** | Plutôt **`Class`** |
|--------|-----------------------|--------------------|
| Taille | petite (quelques champs) | grande |
| Mutation | quasi **immuable** | **souvent modifiée** (via procédures, collections) |
| Sémantique voulue | **valeur** (copie) | **référence** (partage) |
| Passée à des procédures qui la modifient | attention au `ByRef` | naturel |
| Stockée dans des `List`/collections **et** modifiée | à éviter | recommandé |
| Besoin de `Nothing` (null réel) | non (`Nothing` = valeur zéro) | oui |

> ✅ **Règle empirique .NET** : privilégier des **structures immuables** (de petites valeurs comme `Point`), et choisir une **classe** dès qu'il y a **mutation**, **partage** ou **stockage modifiable** en collection. Une `Structure` **mutable** est un **anti-patron** reconnu — et la plupart des UDT VB6 sont des sacs de données **mutables**. Ce choix relève de la **migration** (pas seulement du refactoring) car il conditionne le **comportement**.

---

## 4. Ce que .NET ajoute… et ses limites

Une `Structure` VB.NET est plus riche qu'une UDT, mais avec des **restrictions propres aux types valeur** :

- ✅ elle peut avoir des **méthodes**, **propriétés**, **constructeurs paramétrés**, et **implémenter des interfaces** ;
- ⚠️ sur **.NET Framework 4.7.2**, on **ne peut pas** définir de **constructeur sans paramètre** : les champs sont **toujours** initialisés à zéro/vide. `Dim p As Point` donne `X=0, Y=0` (sans `New`) — cohérent avec la UDT VB6 ;
- ⚠️ une `Structure` **ne peut ni hériter ni être héritée** ;
- ⚠️ `Nothing` affecté à une `Structure` donne sa **valeur par défaut** (tout à zéro), **pas** une référence nulle ;
- ⚠️ l'opérateur **`=`** n'est **pas** défini par défaut sur une `Structure` (il faut surcharger `Operator =`) ; un `Equals` par défaut existe mais est **lent** (comparaison champ à champ par réflexion). En VB6, on ne comparait pas non plus une UDT avec `=` ;
- ⚠️ traitée comme `Object` (collection non générique, etc.), une `Structure` est **boxée** (copiée sur le tas) — coût et perte d'identité (voir **7.1**).

---

## 5. Les champs particuliers

- **Chaînes de longueur fixe** (`Nom As String * 30`) : attribut **`<VBFixedString(30)>`** sur un champ `String`, ou **refactoring** en `String` ordinaire. L'attribut compte surtout pour la **disposition binaire** (voir §6 et **7.4**).
- **Types imbriqués** (`Position As Point`) : une `Structure` imbriquée — sans difficulté.
- **Champs tableau** : un tableau dans une `Structure` est une **référence** (le tableau vit sur le tas), ce qui **mélange** sémantiques valeur et référence au sein d'une même structure — à manier avec prudence.

---

## 6. Disposition binaire : fichiers `Put`/`Get` et API Windows

En VB6, on écrivait souvent une UDT **directement** dans un fichier à accès direct (`Put`/`Get`), ou on la passait à une **API Win32** — ce qui suppose une **disposition mémoire précise**.

> ⚠️ Si vous devez **relire des fichiers binaires hérités** (ou interopérer avec une API), reproduisez la disposition avec **`<StructLayout(LayoutKind.Sequential)>`**, des **types de taille correcte**, et `<VBFixedString>` pour les chaînes fixes (voir **16.5**, *marshaling*). Sinon, le format binaire **ne correspondra plus**.

Pour de **nouveaux** fichiers, .NET ne reproduit pas `Put`/`Get` : on emploie plutôt `BinaryReader`/`BinaryWriter` ou la sérialisation. La correspondance directe « UDT ↔ fichier » de VB6 **disparaît**.

---

## 🔁 Récapitulatif VB6 → VB.NET

| VB6 | VB.NET |
|-----|--------|
| `Type…End Type` | `Structure…End Structure` |
| champ d'UDT | champ **`Public`** explicite (`Dim` = Public dans une `Structure`) |
| UDT passée (toujours **ByRef**) | `Structure` passée **ByVal** par défaut → **copie** ⚠️ |
| `Nom As String * 30` | `<VBFixedString(30)> Public Nom As String` (7.4) |
| UDT écrite via `Put`/`Get` | `<StructLayout(Sequential)>` + marshaling (16.5), ou `BinaryReader/Writer` |
| *(données mutables par nature)* | choisir **`Structure`** (valeur) **ou** **`Class`** (référence) selon l'usage |

---

## ✅ Points clés

- `Type` → `Structure` : conteneur **facile**, **comportement** délicat.
- La sémantique de **valeur** est cohérente à l'affectation, mais : passage **`ByVal` = copie** (la modification ne remonte plus, ⚠️ 10.2), et **type valeur dans une `List`** = modification **fantôme** (alors que les **tableaux** modifient en place).
- **Décision centrale** : `Structure` pour de petites valeurs **immuables** ; **`Class`** dès qu'il y a **mutation**, **partage** ou **stockage modifiable**. Une structure **mutable** est un **anti-patron**.
- Limites des types valeur : pas de constructeur sans paramètre (sur 4.7.2), pas d'héritage, `Nothing` = valeur zéro, `=` à surcharger, **boxing** en `Object`.
- Pour la **compatibilité binaire** (fichiers `Put`/`Get`, API), reproduire la disposition avec **`StructLayout`** + `VBFixedString` (16.5).

---

## 🔗 Renvois

- **7.2** — entiers redimensionnés. · **7.4** — chaînes de longueur fixe (`VBFixedString`). · **7.1** — boxing (`Variant`→`Object`).
- **10.2** — `ByRef`→`ByVal`. · **10.4** — `ParamArray`.
- **12.6** — accès par défaut (`Dim` dans classe vs structure).
- **16.5** — *marshaling*, `StructLayout`, ANSI/Unicode.
- **Suite → [12.8 — Événements : `Event` / `RaiseEvent` / `WithEvents` / `Handles` ; `AddHandler` / `RemoveHandler`](08-evenements.md)**

⏭️ [Événements : `Event`/`RaiseEvent`/`WithEvents`/`Handles` ; `AddHandler`/`RemoveHandler`](/12-poo/08-evenements.md)
