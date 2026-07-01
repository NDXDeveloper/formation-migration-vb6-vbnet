🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.4 Boucles

> **Module 9 — Opérateurs, chaînes et structures de contrôle** · [↩ Retour au chapitre](README.md)

Bonne nouvelle : les boucles font partie de ce qui migre le plus **proprement**. Presque toutes les formes passent **à l'identique** de VB6 à VB.NET. Une seule construction disparaît (`Wend`), et quelques **points d'attention** comportementaux — la **portée du compteur**, le **typage** et le **comportement** de `For Each` — méritent qu'on s'y arrête. VB.NET ajoute par ailleurs l'instruction `Continue`.

---

## Vue d'ensemble

| Boucle VB6 | VB.NET | Statut |
|------------|--------|--------|
| `For…Next` | `For…Next` | Conservée |
| `For Each…Next` | `For Each…Next` | Conservée — typage de l'élément à préciser |
| `Do While/Until…Loop` | identique | Conservée |
| `Do…Loop While/Until` | identique | Conservée |
| `Do…Loop` (infinie) | identique | Conservée |
| `While…Wend` | `While…End While` | **`Wend` supprimé** |
| *(aucune)* | `Continue For` / `Continue Do` / `Continue While` | ✨ Nouveau |

---

## `While…Wend` → `While…End While` (la seule suppression)

Le mot-clé `Wend` **n'existe plus** en VB.NET : il faut fermer la boucle par `End While`. C'est une **erreur de compilation** si `Wend` subsiste — l'assistant de mise à niveau opère la substitution automatiquement.

```vb
' VB6
While i < 10
    i = i + 1
Wend

' VB.NET — Wend remplacé par End While
While i < 10
    i += 1
End While
```

Le comportement est **identique** (test en haut, on boucle tant que la condition est vraie). VB.NET ajoute `Continue While` pour passer à l'itération suivante.

---

## `For…Next`

Syntaxe **conservée** : `For compteur = début To fin [Step pas] … Next`. La fermeture groupée de boucles imbriquées (`Next i, j`) reste valide, même si un `Next` par boucle est plus lisible.

### ⚠️ Point d'attention n°1 — la **portée du compteur**

C'est la principale subtilité de cette section. En VB6, le compteur est déclaré **hors** de la boucle : il **survit** après et conserve sa valeur de sortie. Un idiome très courant en dépend (« boucler jusqu'à trouver, puis lire le compteur ») :

```vb
' VB6 — i existe AVANT et APRÈS la boucle
Dim i As Integer
For i = 1 To 10
    If Trouve(i) Then Exit For
Next i
MsgBox "Position : " & i        ' i = indice trouvé (ou 11 si la boucle est allée au bout)
```

VB.NET permet de déclarer le compteur **en ligne** — mais il est alors **limité au bloc** de la boucle et **inaccessible** après :

```vb
' VB.NET — déclaration en ligne : i n'existe que DANS la boucle
For i As Integer = 1 To 10
    If Trouve(i) Then Exit For
Next
Console.WriteLine(i)            ' ERREUR : i est hors de portée
```

Si le code lit le compteur **après** la boucle, déclarez-le **à l'extérieur**, comme en VB6 :

```vb
' VB.NET — compteur déclaré à l'extérieur : valeur lisible après la boucle
Dim i As Integer
For i = 1 To 10
    If Trouve(i) Then Exit For
Next
Console.WriteLine($"Position : {i}")   ' i accessible (valeur de sortie, 11 si parcours complet)
```

### Point d'attention n°2 — type du compteur et débordement

Avec le **redimensionnement des entiers** (module **7.2**), un compteur déclaré au plus près de l'ancien type peut **déborder** : `Integer` VB6 (16 bits) correspond à `Short` (max 32 767). Une boucle `For i As Short = 1 To 40000` lève une `OverflowException` — comme un `Integer` VB6 levait déjà l'erreur 6. Choisissez `Integer` (.NET, 32 bits) pour les grandes bornes.

### Ce qui ne change **pas** (et qu'on croit souvent à tort)

Les **bornes et le pas sont évalués une seule fois**, à l'entrée de la boucle — en VB6 **comme** en VB.NET. Modifier la variable de borne à l'intérieur ne change donc **pas** le nombre d'itérations :

```vb
Dim n As Integer = 5
For i As Integer = 1 To n
    n = 100        ' n'a AUCUN effet : la boucle s'arrête bien à 5
Next
```

Enfin, les **compteurs en virgule flottante** (`For x = 0 To 1 Step 0.1`) accumulent des erreurs d'arrondi : fragiles dans les deux langages — préférez un compteur entier.

### `Exit For` et le nouveau `Continue For`

```vb
For i As Integer = 1 To 100
    If i Mod 2 = 0 Then Continue For   ' ✨ passe à l'itération suivante (sans GoTo)
    If i > 50 Then Exit For            ' sort de la boucle (VB6 et VB.NET)
    Traiter(i)
Next
```

---

## `Do…Loop`

**Toutes** les variantes sont conservées, avec le même comportement qu'en VB6 : test **en haut** (peut ne jamais s'exécuter) ou **en bas** (s'exécute au moins une fois), condition `While` (tant que vrai) ou `Until` (jusqu'à ce que vrai).

```vb
' Test en haut : peut ne jamais s'exécuter
Do While Not lecteur.EOF
    Lire()
Loop

' Test en bas : s'exécute AU MOINS une fois
Do
    Tenter()
Loop Until reussi
```

`Exit Do` est conservé ; `Continue Do` est nouveau. Avec `Exit Do`, la boucle `Do…Loop` est aussi la **cible naturelle** pour réécrire les anciennes boucles construites au `GoTo` (module **9.5**).

---

## `For Each…Next`

Syntaxe **conservée** : `For Each élément In collection … Next`. Mais trois différences de fond méritent l'attention.

### Typer l'élément (sous `Option Strict On`)

En VB6, `For Each` sur une `Collection` produit un `Variant`/`Object`. En VB.NET, sous `Option Strict On`, on **type** l'élément — ou il vient typé d'une collection générique (`List(Of T)` → `T`) :

```vb
' VB6 — élément non typé
Dim c As Variant
For Each c In maCollection
    ' ...
Next

' VB.NET — élément typé
For Each ligne As String In lignes
    Console.WriteLine(ligne.Trim())
Next
```

Sur une ancienne `Collection` migrée, l'élément reste souvent `Object` et demande un transtypage. Voir modules **7.1** (`Variant` → `Object`) et **8.4** (`Collection` → collections .NET).

### Modifier la collection **pendant** l'énumération

Différence de comportement nette : en VB.NET, modifier une collection **en cours de `For Each`** lève une exception. VB6 était laxiste (comportement indéfini) ; .NET est **strict**.

```vb
' VB.NET
For Each item In liste
    liste.Remove(item)      ' InvalidOperationException : « Collection was modified »
Next
' Solutions : itérer sur une copie, ou boucler à l'envers avec un index (For … Step -1)
```

### Ordre d'énumération

L'**ordre** de parcours et le comportement fin de `For Each` peuvent changer — en particulier si une `Collection` VB6 est migrée vers un `Dictionary`, dont l'ordre d'énumération **n'est pas garanti**. Ce sujet est traité en détail au module **8.5**.

---

## Nouveautés VB.NET : `Continue` et déclaration en ligne

VB.NET ajoute `Continue For` / `Continue Do` / `Continue While`, qui remplacent l'idiome VB6 du **`GoTo` vers une étiquette placée juste avant `Next`/`Loop`** pour « passer à l'itération suivante ». C'est une cible de refactoring naturelle une fois les `GoTo` traités (modules **9.5** et **17.5**). La déclaration **en ligne** du compteur (vue plus haut) est l'autre nouveauté courante — utile, à condition de surveiller la portée.

---

## 🎯 Synthèse — points d'attention de la section

| Élément | Point d'attention | Réflexe |
|---------|-------------------|---------|
| `While…Wend` | `Wend` supprimé | `End While` (substitution automatique) |
| Compteur `For` (portée) | En ligne = limité au bloc | Déclarer à l'extérieur si lu **après** la boucle |
| Compteur `For` (type) | Débordement (`Short` = 16 bits) | `Integer` pour les grandes bornes (7.2) |
| Bornes / pas | Évalués **une seule fois** (inchangé) | *(rien — comportement identique)* |
| `For Each` (typage) | Élément non typé sous `Option Strict` | Typer : `For Each x As T` (7.1, 8.4) |
| `For Each` (mutation) | Modifier la collection en cours → exception | Itérer sur une copie / boucle inverse par index |
| `For Each` (ordre) | Ordre non garanti (`Dictionary`) | Voir 8.5 |
| `GoTo` « continuer » | Idiome hérité | `Continue For/Do/While` (9.5, 17.5) |

---

## 🔗 Pour aller plus loin

- Module **7.2** — Entiers redimensionnés (débordement de compteur `For`)
- Module **7.1** — `Variant` → `Object` (typage de l'élément de `For Each`)
- Module **8.4** / **8.5** — `Collection` → collections .NET ; comportement de `For Each` et ordre d'énumération
- Module **9.5** — `GoTo` / `On…GoTo` supprimés → `Continue`, `Select Case`, refactoring
- Module **17.5** — Refactoring idiomatique (boucles, LINQ)
- Annexe **A** — Tableau de correspondance VB6 → VB.NET

---

> **Section précédente** → [9.3 — `If…Then`, `Select Case`, `IIf` → `If()`](03-conditions.md)
> **Section suivante** → [9.5 — `GoSub`/`Return`, `On…GoTo`, numéros de ligne : supprimés](05-gosub-ongoto-supprimes.md) ⚠️

⏭️ [`GoSub`/`Return`, `On…GoTo`/`On…GoSub`, numéros de ligne : **supprimés** → refactoring](/09-operateurs-controle/05-gosub-ongoto-supprimes.md)
