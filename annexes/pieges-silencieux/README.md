🔝 Retour au [Sommaire](/SOMMAIRE.md)

# B. Catalogue des pièges silencieux — le « top des bugs de migration » ⭐ ⚠️

**L'annexe la plus importante de cette formation.**
Pour chaque piège : **symptôme**, **cause**, **correction**.

> ⚠️ **Pourquoi cette annexe existe.** Le vrai danger d'une migration VB6 → VB.NET n'est pas le
> code qui **refuse de compiler** (l'erreur saute aux yeux, on la corrige). C'est le code qui
> **compile parfaitement mais ne se comporte plus pareil**. Ces régressions sont **silencieuses** :
> aucun message, aucune alerte — juste un résultat faux, un fichier corrompu, une ressource qui
> fuit, ou un calcul décalé de quelques centimes. Ce sont elles qui coûtent le plus cher, car on
> les découvre **en production**. D'où la priorité absolue : les connaître, les traquer
> (module 17.4) et les verrouiller avec un **golden master** (modules 5.5 et 17.3).

> 🧪 **Comment lire chaque fiche.** À chaque piège :
> - 🔴 **Symptôme** — ce que l'on observe (et qui ne ressemble pas forcément à un bug de migration).
> - 🔎 **Cause** — le mécanisme exact, côté langage / runtime / modèle objet.
> - ✅ **Correction** — quoi faire, avec un avant/après. *(Code d'illustration, pas d'exercice.)*

---

## 🏆 Le palmarès en un coup d'œil

| # | Piège | Détecté à la compilation ? | Fréquence | Gravité |
|---|---|---|---|---|
| **B.1** | ByRef → ByVal (et l'inverse) | ❌ Non | 🔁 Très élevée | 🔴 Critique |
| **B.2** | `Integer`/`Long`/`Currency` redimensionnés | ⚠️ Parfois (débordement) | 🔁 Élevée | 🔴 Critique |
| **B.3** | Finalisation : `Class_Terminate` perdu | ❌ Non | 🔁 Élevée | 🔴 Critique |
| **B.4** | Tableaux 0-based / `Option Base` | ⚠️ Parfois (hors bornes) | 🔁 Moyenne | 🟠 Élevée |
| **B.5** | `Variant` : `Empty`/`Null`/`Missing` perdus | ⚠️ Parfois | 🔁 Moyenne | 🟠 Élevée |
| **B.6** | `On Error Resume Next` masquant des bugs | ❌ Non | 🔁 Élevée | 🔴 Critique |
| **B.7** | Propriétés par défaut et `Set` disparu | ⚠️ Parfois | 🔁 Élevée | 🟠 Élevée |
| **B.8** | `True = -1` et conversions booléennes | ❌ Non | 🔁 Moyenne | 🟠 Élevée |
| **B.9** | Dates (`Double` → `DateTime`) et arrondis monétaires | ⚠️ Parfois (`Option Strict`) | 🔁 Moyenne | 🔴 Critique |
| **B.10** | Conversions sensibles à la culture | ❌ Non | 🔁 Moyenne | 🟠 Élevée |

> 💡 La colonne « Détecté à la compilation ? » est la plus parlante : la plupart répondent **Non**.
> C'est toute la définition d'un **piège silencieux** — et la raison pour laquelle on ne peut pas
> se reposer sur le compilateur seul. Activer **`Option Strict On`** (module 17.1) déplace
> *plusieurs* de ces ❌ vers ⚠️ : c'est l'un des meilleurs filets de sécurité.

---

## B.1 — ByRef devenu ByVal (et l'inverse)

> **Carte d'identité** — Compilation : ❌ silencieux · Fréquence : 🔁 très élevée · Gravité : 🔴 critique
> · **C'est LE piège n°1.**

### 🔴 Symptôme
Une procédure censée **modifier son argument** ne le modifie plus. Après l'appel, la variable de  
l'appelant a gardé son ancienne valeur. Aucun avertissement : les calculs en aval sont simplement  
faux. À l'inverse, du code converti « par sécurité » en tout-`ByRef` crée des **effets de bord  
inattendus** (une routine modifie une variable que l'appelant ne s'attendait pas à voir changer).

### 🔎 Cause
Le **mode de passage par défaut** s'inverse entre les deux langages :

| | Sans mot-clé `ByRef`/`ByVal` |
|---|---|
| **VB6** | passe **ByRef** (par référence) |
| **VB.NET** | passe **ByVal** (par valeur) |

Donc `Sub Foo(x As Integer)` signifie **ByRef en VB6** et **ByVal en VB.NET**. Une conversion qui  
recopie la signature sans ajouter le mot-clé **inverse la sémantique** — sans rien signaler.

### ✅ Correction
- **Toujours expliciter** `ByVal` ou `ByRef` sur chaque paramètre.
- **Idéalement, le faire dans le source VB6 *avant* de migrer** (module 5.2) : le code reste valide
  en VB6 et l'intention devient non ambiguë pour l'outil de conversion comme pour l'IA.
- Après migration, **auditer chaque procédure** : modifie-t-elle réellement son argument ? Oui → `ByRef`.
  Non → `ByVal` (plus clair, plus sûr).
- L'assistant de mise à niveau ajoute généralement `ByRef` partout pour *préserver* le comportement
  VB6 : vérifier ensuite que `ByVal` n'était pas l'intention réelle.

```vb
' VB6 — "valeur" est passée ByRef (par défaut)
Sub Incremente(valeur As Long)
    valeur = valeur + 1
End Sub

Dim n As Long : n = 5
Incremente n          ' après l'appel, n vaut 6
```

```vbnet
' VB.NET — MÊME signature, mais ByVal implicite → n NE change PAS
Sub Incremente(valeur As Integer)     ' implicitement ByVal !
    valeur = valeur + 1
End Sub

Dim n As Integer = 5
Incremente(n)         ' n vaut TOUJOURS 5 — régression silencieuse
```

```vbnet
' Correction : rendre le mode explicite
Sub Incremente(ByRef valeur As Integer)
    valeur = valeur + 1
End Sub
```

> 🔎 **Subtilité supplémentaire** : encadrer un argument de parenthèses **force le passage par
> valeur**, même si la procédure déclare `ByRef`. `Foo(x)` (parenthèses superflues sur un appel
> sans `Call`) et `Foo x` peuvent donc différer. À surveiller lors d'un remplacement mécanique.

---

## B.2 — `Integer` / `Long` / `Currency` redimensionnés et débordements

> **Carte d'identité** — Compilation : ⚠️ parfois (`OverflowException`) · Fréquence : 🔁 élevée · Gravité : 🔴 critique

### 🔴 Symptôme
Le code compile, puis : `OverflowException` à l'exécution là où VB6 ne bronchait pas, **ou**  
résultats faux **en interop / API / fichiers binaires** parce que la **taille des champs** ne  
correspond plus. Une structure passée à une API Win32 envoie des champs de 32 bits là où l'API en  
attend 16 → valeurs erronées ou corruption mémoire.

### 🔎 Cause
Les mots-clés sont conservés mais **les tailles changent** :

| VB6 | Taille VB6 | Cible **par taille** | Taille VB.NET |
|---|---|---|---|
| `Integer` | **16 bits** | `Short` | 16 bits |
| `Long` | **32 bits** | `Integer` | 32 bits |
| `Currency` | 64 bits (×10000) | `Decimal` | 128 bits |

Le piège : `As Integer` (VB6, 16 bits) **reste** `As Integer` en VB.NET… mais y fait désormais
**32 bits**. De même `Long` (32 bits) devient `Long` (64 bits). La bonne correspondance est par  
**taille**, pas par **nom**. Par ailleurs, VB.NET active par défaut le **contrôle de débordement
des entiers**, ce qui fait surgir des exceptions auparavant absentes.

### ✅ Correction
- Mapper **par taille** : `Integer`(VB6) → **`Short`** ; `Long`(VB6) → **`Integer`** ;
  `Currency` → **`Decimal`** (jamais `Double`, voir B.9).
- Pour les **API Win32 et structures binaires**, être explicite sur la taille et utiliser
  `<StructLayout>` (module 16.5).
- **Garder activé** le contrôle de débordement pendant la migration (projet → ne *pas* cocher
  « Supprimer les contrôles de dépassement de capacité ») : il révèle les vrais problèmes.
- Voir aussi l'**Annexe D** (tailles, plages, conversions).

```vb
' VB6 — COORD (API console Win32) : 4 octets, 2 champs SHORT de 16 bits
Private Type COORD
    x As Integer   ' SHORT Win32 = 16 bits
    y As Integer   ' SHORT Win32 = 16 bits
End Type
```

```vbnet
' VB.NET — conversion NAÏVE : Integer y fait 32 bits → 8 octets, désaligné avec l'API
<StructLayout(LayoutKind.Sequential)>
Structure COORD
    Public x As Integer   ' 32 bits !
    Public y As Integer
End Structure
```

```vbnet
' Correction : Short préserve les 16 bits → structure de 4 octets (correcte)
<StructLayout(LayoutKind.Sequential)>
Structure COORD
    Public x As Short
    Public y As Short
End Structure

' Et pour les montants :
' VB6  : Dim prix As Currency
Dim prix As Decimal      ' surtout PAS Double
```

---

## B.3 — Finalisation : `Class_Terminate` qui ne se déclenche plus

> **Carte d'identité** — Compilation : ❌ silencieux · Fréquence : 🔁 élevée · Gravité : 🔴 critique
> · **Le piège n°1 du *modèle objet*** (module 2.2).

### 🔴 Symptôme
Des ressources libérées « au bon moment » en VB6 **traînent** : fichiers restés verrouillés,  
connexions qui épuisent le pool, *handles* GDI qui fuient, nettoyage/journalisation de
`Class_Terminate` qui s'exécute **trop tard** ou jamais avant la fin du processus. Des effets de
bord d'**ordre de destruction** disparaissent.

### 🔎 Cause
Deux modèles de durée de vie radicalement différents :

- **VB6** — **comptage de références** (COM) : quand la **dernière** référence disparaît (variable
  hors de portée ou `Set x = Nothing`), le compteur tombe à 0 et `Class_Terminate` se déclenche
  **immédiatement et de façon déterministe**.
- **VB.NET** — **ramasse-miettes (GC)** : l'objet est récupéré **quand le GC le décide**
  (non déterministe). Le destructeur `Finalize` peut s'exécuter bien plus tard, sur un autre thread,
  voire pas avant la fermeture. **Il n'y a plus de destruction déterministe.**

### ✅ Correction
- Traduire le nettoyage de `Class_Terminate` via le motif **`IDisposable`** (méthode `Dispose`) —
  **et non** via `Finalize`/destructeur (qui ne garantit pas le *quand*).
- Au point d'appel, encadrer par **`Using`** pour garantir un `Dispose` **déterministe**.
- Pour une classe encapsulant des ressources non managées, implémenter le **motif Dispose complet**
  (`Dispose(disposing)` + finaliseur optionnel + `GC.SuppressFinalize`).
- **Ne pas compter sur `Set x = Nothing`** : cela ne fait que retirer la référence ; le GC décide
  toujours du moment (voir aussi B.7 et l'Annexe A, §11).

```vb
' VB6 — libération immédiate quand l'objet meurt
Private Sub Class_Terminate()
    cn.Close
    Set cn = Nothing
End Sub
```

```vbnet
' VB.NET — équivalent DÉTERMINISTE via IDisposable
Public Class Depot
    Implements IDisposable

    Private cn As SqlConnection

    Public Sub Dispose() Implements IDisposable.Dispose
        cn?.Close()
        cn?.Dispose()
    End Sub
End Class

' Au point d'appel : libération garantie en fin de bloc
Using d As New Depot()
    ' ... travail ...
End Using   ' Dispose() appelé ICI, de façon déterministe
```

---

## B.4 — Tableaux 0-based et `Option Base`

> **Carte d'identité** — Compilation : ⚠️ parfois (`IndexOutOfRangeException`) · Fréquence : 🔁 moyenne · Gravité : 🟠 élevée

### 🔴 Symptôme
Erreurs de décalage d'un cran (*off-by-one*). Une boucle `For i = 1 To UBound(a)` **saute l'élément
0** ou en manque un. Un tableau dimensionné `Dim a(10)` n'a plus le nombre d'éléments attendu.
`IndexOutOfRangeException`, ou données silencieusement décalées.

### 🔎 Cause
- VB6 autorisait `Option Base 0` **ou** `Option Base 1`, ainsi que les bornes explicites
  `Dim a(1 To 10)`. Beaucoup de bases de code VB6 sont **1-based**.
- VB.NET : les tableaux sont **toujours 0-based**. `Option Base` **n'existe plus** et
  `Dim a(1 To 10)` est **interdit**.
- Conséquence subtile : `Dim a(10)` reste « indice max = 10 » dans les deux langages (donc **11
  éléments**), ce qui correspond à VB6 *base 0*. Mais le code écrit pour `Option Base 1` (premier
  élément à l'indice 1) lit désormais un élément 0 inattendu — d'où les décalages.

### ✅ Correction
- **Normaliser tous les tableaux en 0-based.**
- Réécrire les boucles : `For i = 0 To UBound(a)` (ou `For i = 0 To a.Length - 1`).
- Remplacer `Dim a(1 To 10)` par `Dim a(9)` (et décaler les indices) **ou** passer à une
  collection (`List(Of T)`).
- **Auditer en priorité** les modules contenant `Option Base 1` : ce sont les plus à risque.
- Éviter l'émulation 1-based de l'espace `Microsoft.VisualBasic.Compatibility` (la béquille à
  retirer, module 4.2).

```vb
' VB6 avec Option Base 1 — 12 éléments, indices 1..12
Option Base 1
Dim mois(12) As String
For i = 1 To 12
    mois(i) = ...
Next
```

```vbnet
' VB.NET — Dim mois(12) crée 13 éléments (indices 0..12) !
Dim mois(12) As String
For i = 1 To 12          ' l'indice 0 est ignoré → gaspillage et risque de décalage
    mois(i) = ...
Next
```

```vbnet
' Correction : 0-based propre
Dim mois(11) As String  ' 12 éléments, indices 0..11
For i = 0 To 11
    mois(i) = ...
Next

' ou, idiomatique :
Dim mois As New List(Of String)
```

---

## B.5 — `Variant` : `Empty` / `Null` / `Missing` perdus

> **Carte d'identité** — Compilation : ⚠️ parfois · Fréquence : 🔁 moyenne · Gravité : 🟠 élevée

### 🔴 Symptôme
Les tests `If IsEmpty(v)`, `If IsNull(v)`, `If IsMissing(arg)` ne se comportent plus correctement
(ou renvoient toujours faux). La logique qui distinguait « jamais initialisé » (`Empty`), « valeur
nulle voulue / venue de la base » (`Null`) et « argument non fourni » (`Missing`) s'effondre :  
détection d'« initialisé ? » défaillante, exceptions sur des nulls de base de données.

### 🔎 Cause
- Le `Variant` VB6 portait **trois états spéciaux** : `Empty` (non initialisé), `Null` (absence de
  valeur intentionnelle), `Missing` (argument optionnel omis).
- VB.NET remplace `Variant` par **`Object`**. Un `Object` non initialisé vaut simplement
  **`Nothing`**. Il n'y a **plus** d'`Empty` ni de `Null`-Variant. `Missing` n'a plus de sens car
  les paramètres `Optional` exigent une **valeur par défaut**.
- Le « null de base de données » se représente désormais par **`DBNull.Value`** (singleton
  distinct), à ne pas confondre avec `Nothing`.

### ✅ Correction
- `IsEmpty(v)` → tester **`v Is Nothing`** (ou comparer à la valeur par défaut du type).
- `IsNull(v)` (base de données) → **`IsDBNull(v)`** ou `v Is DBNull.Value`.
- `IsMissing(arg)` → donner une **valeur par défaut explicite** au paramètre et tester contre elle
  (ex. `Optional id As Integer = -1` puis `If id = -1`), ou utiliser des **surcharges**, ou
  **`Nullable(Of T)`**.
- Distinguer **délibérément** les trois concepts, autrefois confondus sous `Variant`.

```vb
' VB6
Function Lire(Optional cle As Variant) As Variant
    If IsMissing(cle) Then cle = "defaut"
    If IsNull(rs!champ) Then
        Lire = ""
    Else
        Lire = rs!champ
    End If
End Function
```

```vbnet
' VB.NET — chaque état Variant traité explicitement
Function Lire(Optional cle As String = Nothing) As String
    If cle Is Nothing Then cle = "defaut"        ' remplace IsMissing
    If IsDBNull(rs!champ) Then                    ' remplace IsNull (base de données)
        Return ""
    Else
        Return rs!champ.ToString()
    End If
End Function
```

---

## B.6 — `On Error Resume Next` masquant des bugs

> **Carte d'identité** — Compilation : ❌ silencieux · Fréquence : 🔁 élevée · Gravité : 🔴 critique

### 🔴 Symptôme
Des erreurs qui *devraient* remonter sont **avalées** : le programme continue dans un état  
incohérent et produit des résultats faux ou des données corrompues au lieu d'échouer franchement.  
Pire après migration : `On Error Resume Next` laissé en place masque des **exceptions .NET**
(référence nulle, débordement, format) qui n'existaient pas en VB6.

### 🔎 Cause
- `On Error Resume Next` demande d'**ignorer l'erreur et de continuer** à l'instruction suivante.
  Souvent employé pour sauter une erreur *attendue et bénigne* — mais il masque aussi les
  **inattendues**.
- En .NET, ce code **fonctionne encore** (modèle `On Error` hérité) et continue d'avaler les
  erreurs. Or .NET lève **plus d'exceptions, et différentes**, que VB6 : des blocs `Resume Next`
  jadis « sûrs » masquent désormais de vrais problèmes (un `Nothing` là où VB6 avait `Empty`, une
  `FormatException` de culture, une `OverflowException`).

### ✅ Correction
- **Ne pas traduire en bloc** `On Error Resume Next`. Traiter **chaque occurrence** :
  - si elle protégeait **une** opération connue et bénigne → encadrer **cette seule** opération
    d'un `Try...Catch` **ciblé** (capturant uniquement l'exception attendue) ;
  - si c'était de la suppression d'erreur paresseuse → **la supprimer** et gérer correctement.
- Ne **jamais** laisser un `On Error Resume Next` large couvrir de gros blocs dans le code migré.
- Préférer `Catch TypePrécis` plutôt que tout capturer.
- Endroit idéal pour **ajouter de la journalisation** pendant la validation (module 17.4) afin de
  découvrir ce qui était silencieusement avalé.

```vb
' VB6 — masque TOUT entre les deux instructions
On Error Resume Next
valeur = CInt(saisie)        ' saisie invalide ? valeur inchangée, on continue
resultat = 100 / diviseur    ' division par zéro ? ignorée !
On Error GoTo 0
```

```vbnet
' VB.NET — ciblé : on ne masque que ce qui est attendu
Dim valeur As Integer
If Not Integer.TryParse(saisie, valeur) Then
    valeur = 0               ' décision EXPLICITE
End If

Try
    resultat = 100 / diviseur
Catch ex As DivideByZeroException
    ' traitement explicite (et journalisé)
End Try
```

---

## B.7 — Propriétés par défaut et `Set` disparu

> **Carte d'identité** — Compilation : ⚠️ parfois · Fréquence : 🔁 élevée · Gravité : 🟠 élevée

### 🔴 Symptôme
Des affectations/comparaisons qui s'appuyaient sur une **propriété par défaut** pointent désormais  
vers **l'objet** au lieu de sa valeur (ou ne compilent plus). `If txtNom = "x"` ne compare plus le  
texte ; `x = Text1` récupère la *référence* du contrôle, pas son texte. La disparition de `Set`  
brouille la distinction « affectation d'objet » vs « affectation de valeur ».

### 🔎 Cause
- VB6 acceptait les **propriétés par défaut sans argument** : `Label1 = "Salut"` signifiait
  `Label1.Caption = "Salut"` ; `x = Text1` signifiait `x = Text1.Text` ; `rs("col")` renvoyait la
  **`Value`** du champ.
- VB.NET **supprime les propriétés par défaut sans paramètre**. Seules subsistent les propriétés
  **paramétrées** (indexeurs, ex. `Collection(i)`). Donc `Text1` seul désigne **l'objet contrôle**,
  plus son texte.
- Le mot-clé **`Set` disparaît** : l'affectation d'objet s'écrit simplement `=`. En VB6, `Set x = o`
  (référence) et `x = o` (valeur de la propriété par défaut) étaient **deux instructions
  différentes** ; supprimer `Set` mécaniquement peut donc changer le sens.

### ✅ Correction
- Rendre **chaque accès explicite** : `.Text`, `.Caption`, `.Value`… **ne jamais** se reposer sur
  une propriété par défaut.
- Supprimer `Set`, mais **vérifier chaque cas** : là où VB6 lisait une valeur par défaut **sans**
  `Set`, ajouter la propriété explicite (`.Value`) en VB.NET ; là où il y avait `Set`, c'est une
  simple affectation de référence.
- Pour l'accès aux données, remplacer les lectures `rs("col")` par `rs("col").Value` (ou migrer
  vers ADO.NET, module 15).
- L'assistant insère souvent `.Text`/`.Value` : **vérifier qu'il a choisi le bon**.

```vb
' VB6
Dim s As String
s = txtNom                 ' lit txtNom.Text (propriété par défaut)
lblTitre = "Bonjour"       ' écrit lblTitre.Caption
Set ctl = Me.Controls(0)   ' référence d'objet
```

```vbnet
' VB.NET — tout explicite
Dim s As String = txtNom.Text
lblTitre.Text = "Bonjour"            ' Caption → Text en WinForms (Annexe C)
Dim ctl As Control = Me.Controls(0)  ' plus de Set
```

---

## B.8 — `True = -1` et conversions booléennes

> **Carte d'identité** — Compilation : ❌ silencieux · Fréquence : 🔁 moyenne · Gravité : 🟠 élevée

### 🔴 Symptôme
Du code arithmétique ou bit-à-bit fondé sur des booléens donne des **nombres différents**. Les  
motifs du type `total = total - (condition)` (compter les « vrais ») ou `drapeaux And masqueBool`  
produisent de mauvaises valeurs. En **interop / marshaling**, un booléen converti vers un champ  
numérique envoie **1** là où VB6 envoyait **-1** (ou l'inverse), cassant systèmes externes, formats  
de fichiers ou API. Idem au stockage en base.

### 🔎 Cause
- En VB6, `True` vaut **-1** (tous les bits à 1, `0xFFFF`) et `False` vaut 0. D'où l'usage des
  opérateurs bit-à-bit `And`/`Or` comme opérateurs logiques.
- En VB.NET, la valeur CLR sous-jacente de `True` est **1** (`System.Boolean`, 1 octet).
  **Mais**, par compatibilité, **`CInt(True)` renvoie toujours -1**. D'où une **incohérence** :

| Conversion en VB.NET | Résultat pour `True` |
|---|---|
| `CInt(True)` / `CType(True, Integer)` | **-1** (compatibilité VB) |
| `Convert.ToInt32(True)` | **1** |
| Marshaling / `BitConverter` (octet réel) | **1** |

Mélanger les fonctions de conversion VB (qui donnent -1) et les conversions du framework ou  
l'interop (qui donnent 1) produit des résultats incohérents.

### ✅ Correction
- **Ne pas s'appuyer sur une valeur numérique du booléen.** Utiliser la **logique booléenne**
  (`If`, `AndAlso`, `OrElse`) plutôt que de l'arithmétique sur des booléens.
- Quand une représentation numérique est *réellement* nécessaire, **choisir et fixer** la règle
  (0/1 ou -1) et utiliser **la même voie de conversion partout** : `If(b, 1, 0)` pour du 0/1
  portable ; `CInt(b)` pour la compatibilité VB6 historique (-1).
- En **interop / formats externes**, contrôler la représentation explicitement (ex.
  `<MarshalAs(UnmanagedType.Bool)>` vs `VariantBool`). Le `VARIANT_BOOL` COM vaut -1 pour vrai.
- En **base de données**, mapper booléen ↔ colonne explicitement (souvent 1/0) sans supposer -1.

```vb
' VB6 — compte les "vrais" en s'appuyant sur True = -1
Dim nb As Long
nb = nb - (estValide)        ' soustrait -1 → ajoute 1 quand vrai
```

```vbnet
' VB.NET — fragile : CInt(True) = -1, mais d'autres voies donnent 1
' À RÉÉCRIRE en logique booléenne explicite :
Dim nb As Integer
If estValide Then nb += 1

' Si une valeur numérique est nécessaire, la rendre explicite et portable :
Dim drapeau As Integer = If(estValide, 1, 0)   ' et NON CInt(estValide), qui vaut -1
```

---

## B.9 — Dates (`Double` → `DateTime`) et arrondis monétaires

> **Carte d'identité** — Compilation : ⚠️ parfois (`Option Strict`) · Fréquence : 🔁 moyenne · Gravité : 🔴 critique

Deux pièges jumeaux, tous deux liés à un changement de **représentation interne**.

### 🔴 Symptôme
- **Dates** : l'arithmétique qui traitait les dates comme des `Double` (`dteFin = dteDebut + 30`
  pour ajouter 30 jours, comparaison de dates comme des nombres, stockage d'une date dans un
  `Double`) ne compile plus proprement ou donne des résultats faux. Les conversions implicites
  `Date`↔`Double` sont refusées sous `Option Strict`.
- **Monétaire** : des totaux diffèrent de quelques centimes ; l'arrondi des cas « .5 » part dans
  l'« autre sens » ; des erreurs d'accumulation apparaissent quand `Currency` a été remplacé par
  `Double` au lieu de `Decimal`.

### 🔎 Cause
- **Dates** : VB6 stockait `Date` en interne comme un **`Double`** (partie entière = jours depuis le
  30/12/1899, partie fractionnaire = heure). On pouvait donc additionner un nombre, soustraire deux
  dates pour obtenir des jours, et mélanger `Date` et `Double`. VB.NET `DateTime` est une
  **structure distincte** (ticks depuis l'an 1), **pas un `Double`** : on utilise `.AddDays(n)`, et
  une différence donne un **`TimeSpan`**, pas un nombre.
- **Monétaire** : VB6 `Currency` est un entier 64 bits **mis à l'échelle (×10000)**, exact sur 4
  décimales. Le migrer en **`Double`** (virgule flottante binaire) introduit des erreurs de
  représentation (0,1 n'est pas exact) et un arrondi différent. La bonne cible est **`Decimal`**
  (base 10, exacte pour les montants).

### ✅ Correction
**Dates :**
- Utiliser les méthodes de `DateTime` : `.AddDays`, `.AddMonths`, `.AddHours` ; différences via
  `TimeSpan` (`(d2 - d1).TotalDays`).
- Supprimer toute dépendance à la représentation « date = Double » ; ne pas stocker de dates dans
  des `Double`.
- **Activer `Option Strict`** (module 17.1) pour faire ressortir ces conversions implicites.
- Pour des données héritées stockées en double OLE Automation, convertir **à la frontière** avec
  `Date.FromOADate` / `ToOADate`.

**Monétaire :**
- Mapper `Currency` → **`Decimal`**, **jamais** `Double`.
- Rendre l'**arrondi explicite** : `Math.Round(valeur, 2, MidpointRounding.ToEven)` **ou**
  `MidpointRounding.AwayFromZero`, selon la **règle métier** — et la documenter (module 5.4).
- **Re-vérifier** les cas de test monétaires contre le **golden master** (module 17.3).

```vb
' VB6 — Date est un Double sous le capot
Dim echeance As Date
echeance = Date + 30          ' ajoute 30 jours
Dim jours As Long
jours = dteFin - dteDebut     ' différence en jours
```

```vbnet
' VB.NET — DateTime n'est pas un nombre
Dim echeance As Date = Today.AddDays(30)
Dim jours As Integer = CInt((dteFin - dteDebut).TotalDays)   ' via TimeSpan

' À la frontière avec des données héritées (Double OLE) :
Dim d As Date = Date.FromOADate(valeurDouble)
```

```vbnet
' Monétaire : Decimal, surtout pas Double
' VB6 : Dim montant As Currency
Dim montant As Decimal
Dim arrondi As Decimal = Math.Round(montant, 2, MidpointRounding.AwayFromZero)
```

---

## B.10 — Conversions sensibles à la culture (`CDbl` / `CDate` / `Format`)

> **Carte d'identité** — Compilation : ❌ silencieux · Fréquence : 🔁 moyenne · Gravité : 🟠 élevée
> · **Le piège « ça marche sur ma machine ».**

### 🔴 Symptôme
L'application fonctionne chez le développeur mais **produit des nombres ou des dates faux** sur des  
machines aux **paramètres régionaux** différents (ou l'inverse). `CDbl("1.5")` et `CDbl("1,5")`  
s'interprètent différemment selon le **séparateur décimal**. Les dates se lisent avec mois/jour  
inversés. La lecture/écriture de fichiers, de CSV ou l'interop corrompt des valeurs. Symptôme  
sournois : n'échoue que pour **certains** utilisateurs/régions, difficile à reproduire.

### 🔎 Cause
- En VB.NET, les fonctions sensibles à la culture — `CDbl`, `CSng`, `CDate`, `CInt` (depuis une
  chaîne), `Format`, `.ToString`, `Double.Parse`, `DateTime.Parse` — utilisent par défaut la
  **`CurrentCulture`** du thread. Le **séparateur décimal**, le **séparateur de milliers** et le
  **format de date** dépendent donc des réglages régionaux.
- Conséquence : analyser une saisie utilisateur, un fichier de données ou un format fixe avec des
  fonctions sensibles à la culture rend le comportement **dépendant de l'environnement** (FR : « , »
  décimal et `jj/MM/aaaa` ; US : « . » et `MM/jj/aaaa`).

### ✅ Correction
**Séparer les conversions de « données » des conversions d'« affichage ».**
- Pour les **données techniques / persistantes** (fichiers, protocoles, interop — tout ce qui doit
  rester **stable**) : conversions **invariantes de culture** via
  `CultureInfo.InvariantCulture`. Ex. `Double.Parse(s, CultureInfo.InvariantCulture)`,
  `valeur.ToString(CultureInfo.InvariantCulture)`,
  `DateTime.ParseExact(s, "yyyy-MM-dd", CultureInfo.InvariantCulture)`. Les fonctions VB
  **`Val`** et **`Str`** sont **déjà invariantes** : de bons substituts pour lire/écrire des nombres
  dans un format stable.
- Pour l'**affichage utilisateur** : utiliser délibérément la `CurrentCulture` (comportement par
  défaut), ce qui est précisément voulu.
- **Réserver** `CDbl`/`CDate`/`Format` à l'IHM ; ne pas les employer pour analyser des données à
  format fixe.
- **Passer explicitement** l'`IFormatProvider`/`CultureInfo` plutôt que de dépendre de la culture
  ambiante.
- **Tester sous au moins deux cultures** (ex. `en-US` et `fr-FR`) pour détecter ces régressions
  (module 17.3).

```vb
' VB6 / conversion naïve en VB.NET — dépend des paramètres régionaux
Dim prix As Double
prix = CDbl(ligneFichier)         ' "3.14" KO en culture FR (qui attend "3,14")
Dim s As String
s = Format(prix, "0.00")          ' "3,14" en FR vs "3.14" en US — incohérent en fichier
```

```vbnet
Imports System.Globalization

' VB.NET — DONNÉES : invariant de culture
Dim prix As Double = Double.Parse(ligneFichier, CultureInfo.InvariantCulture)
' ou l'équivalent VB historique, déjà invariant :
Dim prix2 As Double = Val(ligneFichier)

' Écriture de données stables :
Dim aEcrire As String = prix.ToString(CultureInfo.InvariantCulture)

' AFFICHAGE à l'utilisateur — culture courante (voulue) :
Dim affichage As String = prix.ToString("N2")
```

---

## 🎯 Comment chasser ces pièges (méthode)

Aucun de ces pièges ne se découvre « par hasard » de façon fiable. La formation propose un **filet
à plusieurs mailles** :

1. **En amont (modules 5.1–5.2)** — réduire la surface : `Option Explicit` partout, typage et
   `ByVal`/`ByRef` explicites **dans le source VB6** avant migration.
2. **Golden master (modules 5.5 et 17.3)** — capturer le comportement de référence de l'appli VB6,
   puis comparer **automatiquement** après migration. C'est le détecteur principal des écarts
   silencieux (ByRef, arrondis, dates, culture).
3. **`Option Strict On` progressif (module 17.1)** — transforme plusieurs pièges ❌ *silencieux* en
   ⚠️ *erreurs de compilation* (conversions Date↔Double, *late binding* résiduel, pertes de
   précision).
4. **Traque ciblée (module 17.4)** — revue dédiée : ByRef, entiers redimensionnés, finalisation,
   bornes de tableaux, `Variant`, `On Error Resume Next`, propriétés par défaut, `True = -1`, dates,
   culture.
5. **Tests multi-culture** — exécuter la suite sous au moins deux cultures (B.10).
6. **L'IA comme détecteur (module 19.3)** — utile pour **repérer** les motifs à risque… mais à
   **valider systématiquement** (l'IA invente parfois de faux équivalents, module 19.5).

---

## 🔗 Annexes et modules liés

- **Annexe A** — Tableau de correspondance VB6 → VB.NET (la forme « directe » vs « idiomatique »).
- **Annexe C** — Correspondance des contrôles (dont `Caption` → `Text`, couleurs `OLE_COLOR`).
- **Annexe D** — Tailles, plages et règles de conversion des types (détail de B.2).
- **Annexe F** — Modèles de tests *golden master* (le détecteur du §🎯).
- **Module 2** — Différences fondamentales (2.2 finalisation, 2.4 ByRef/ByVal…).
- **Module 16** — Interop COM & API (marshaling, `StructLayout`, FSO → `System.IO`).
- **Module 17** — Valider, fiabiliser, refactoriser (`Option Strict`, non-régression, traque).

---

**Juin 2026**  
**Cible** : .NET Framework 4.7.2 (le « pont » depuis VB6)  
**Licence** : Creative Commons BY-NC-SA 4.0

⏭️ [Correspondance des contrôles VB6 → Windows Forms](/annexes/correspondance-controles/README.md)
