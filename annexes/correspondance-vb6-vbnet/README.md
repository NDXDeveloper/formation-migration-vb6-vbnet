🔝 Retour au [Sommaire](/SOMMAIRE.md)

# A. Tableau de correspondance VB6 → VB.NET ⭐

**Aide-mémoire de conversion : mots-clés, fonctions intrinsèques, instructions, opérateurs.**
La référence à garder ouverte pendant toute la migration.

> 🎯 **Cible** : VB.NET sur **.NET Framework 4.7.2**. Beaucoup de fonctions VB6 *existent
> toujours* via l'espace de noms `Microsoft.VisualBasic` (livré avec le runtime). Elles
> **compilent et fonctionnent** — c'est volontaire, pour faciliter la transition. Mais
> « ça fonctionne » ne veut pas dire « c'est idiomatique » ni « ça se comporte à l'identique ».
> Ce tableau donne donc **deux colonnes de cible** : l'équivalent **direct** (souvent
> `Microsoft.VisualBasic`, pour migrer vite) et l'équivalent **idiomatique .NET** (vers lequel
> tendre, voir module 17 « refactoring »).

---

## 🧭 Comment lire ce tableau

| Colonne | Sens |
|---|---|
| **VB6** | La forme d'origine dans le code source à migrer. |
| **Équivalent direct** | Ce qui compile sans (ou presque) changer la sémantique. Idéal pour une **première passe** rapide. |
| **Idiomatique .NET** | La forme cible recommandée à terme (lisibilité, performance, `Option Strict On`). |
| **⚠️ Piège / Note** | Le **changement silencieux** ou le point d'attention. C'est la colonne la plus importante. |

> ⚠️ **Rappel central de la formation** : le danger n'est pas le code qui ne compile pas,
> c'est le code qui **compile mais ne se comporte plus pareil**. Chaque ⚠️ ci-dessous renvoie,
> en pratique, à un piège détaillé dans l'**Annexe B** (catalogue des pièges silencieux).

---

## 1. Déclarations et mots-clés

| VB6 | Équivalent direct | Idiomatique .NET | ⚠️ Piège / Note |
|---|---|---|---|
| `Dim x As Integer` | `Dim x As Short` | `Dim x As Short` | **`Integer` VB6 = 16 bits → `Short`**. En VB.NET, `Integer` fait **32 bits**. Voir §2 et Annexe D. |
| `Dim a, b As Integer` | `Dim a, b As Short` | déclarer séparément | ⚠️ **Trap classique** : en VB6, `a` est un **`Variant`**, seul `b` est typé. En VB.NET, **les deux** prennent le type. Le comportement change silencieusement. |
| `Set obj = New Classe` | `obj = New Classe` | `obj = New Classe` | ⚠️ Le mot-clé **`Set` disparaît**. Une recherche/remplacement aveugle de `Set ` est risquée (ne pas confondre avec `Property Set`). |
| `Let x = 5` | `x = 5` | `x = 5` | Le mot-clé **`Let` disparaît** (l'affectation par valeur est implicite). |
| `Global x` | `Public x` | `Public x` (dans un `Module`) | `Global` supprimé → `Public`. |
| `Public` / `Private` / `Friend` | identiques | identiques | `Friend` conservé. Ajout de `Protected`, `Protected Friend`. |
| `Static x As Long` | `Static x As Integer` | `Static x As Integer` | Conservé. Attention au type (`Long` VB6 → `Integer` .NET). |
| `Const PI = 3.14` | identique | `Const PI As Double = 3.14` | Conservé. Préférer le **typage explicite**. |
| `Type T ... End Type` | `Structure T ... End Structure` | `Structure` (+ `<StructLayout>` si interop) | ⚠️ `Type` → **`Structure`**. Pour le marshaling Win32, voir module 16.5 (disposition mémoire). |
| `Enum E ... End Enum` | identique | identique | Conservé. |
| `ReDim a(10)` | `ReDim a(10)` | `ReDim a(10)` | ⚠️ Conservé **mais 0-based** : `a(10)` a **11 éléments** (indices 0→10). Voir §4 et Annexe B.4. |
| `ReDim Preserve a(n)` | identique | `Array.Resize(a, n + 1)` | Conservé. `Array.Resize` plus idiomatique. |
| `Dim s As String * 10` | `<VBFixedString(10)> Dim s As String` | `String` simple | ⚠️ **Chaîne de longueur fixe** : plus de syntaxe `* n`. L'attribut `VBFixedString` n'agit qu'au marshaling, pas en mémoire. |
| `Option Explicit` | identique | identique (activé par défaut) | Conservé. **Toujours le garder.** |
| `Option Base 1` | *(supprimé)* | — | ⚠️ **`Option Base` n'existe plus** : les tableaux sont **toujours 0-based**. Annexe B.4. |
| `Option Compare Text/Binary` | identique | identique | Conservé (affecte `=`, `Like`, tris de chaînes). |
| `Declare Function ... Lib "..."` | `Declare` (conservé) | `<DllImport("...")>` (P/Invoke) | ⚠️ `Declare` fonctionne encore, mais P/Invoke est plus sûr (ANSI/Unicode explicite). Module 16.4. |
| `Implements Interface` | identique | identique | Conservé. |
| *(pas d'équivalent)* | — | `Namespace`, `Imports`, `Inherits`, `Overrides`, `Overloads`, `Shadows` | **Nouveautés** .NET (héritage, espaces de noms, surcharge). |

---

## 2. Types de données (résumé)

> 📌 Vue express ci-dessous. Tailles, plages et règles de conversion **complètes** : **Annexe D**.

| VB6 | Taille VB6 | Équivalent .NET (4.7.2) | Type CLR | ⚠️ Piège / Note |
|---|---|---|---|---|
| `Integer` | 16 bits | `Short` | `System.Int16` | ⚠️ **Le renommage le plus dangereux.** |
| `Long` | 32 bits | `Integer` | `System.Int32` | ⚠️ `Long` VB6 ≠ `Long` VB.NET (qui fait **64 bits**). |
| *(nouveau)* | — | `Long` | `System.Int64` | 64 bits en VB.NET. |
| `Currency` | 64 bits (×10000) | `Decimal` | `System.Decimal` | ⚠️ Pas d'équivalent exact. `Decimal` est le plus proche (montants). Arrondis : Annexe B.9. |
| `Single` / `Double` | 32 / 64 bits | identiques | `Single` / `Double` | Conservés. |
| `Byte` | 8 bits | identique | `System.Byte` | Conservé. |
| `Boolean` | 16 bits | `Boolean` | `System.Boolean` (1 octet) | ⚠️ `True` valait **-1** en VB6. Conversions et opérations bit-à-bit à revoir. Annexe B.8. |
| `String` | — | `String` | `System.String` | Conservé, mais **toujours Unicode** ; les API ANSI exigent du marshaling. |
| `Date` | 64 bits (`Double`) | `Date` (alias `DateTime`) | `System.DateTime` | ⚠️ Représentation interne **différente** (plus un `Double`). Arithmétique de dates à revoir. Annexe B.9. |
| `Variant` | 16 octets | `Object` | `System.Object` | ⚠️ `Empty` / `Null` / `Missing` **disparaissent**. Annexe B.5. |
| `Object` | — | `Object` | `System.Object` | Conservé. |

---

## 3. Opérateurs

| VB6 | VB.NET | ⚠️ Piège / Note |
|---|---|---|
| `&` (concaténation) | `&` | Conservé. **À privilégier pour concaténer.** |
| `+` (sur chaînes) | `+` | ⚠️ **Ambigu** : avec `Option Strict On`, `String + Number` provoque une **erreur** ; sans, le résultat dépend des opérandes. **Toujours utiliser `&`** pour concaténer. |
| `\` (division entière) | `\` | Conservé. |
| `Mod` | `Mod` | Conservé. |
| `^` (puissance) | `^` | Conservé. |
| `=` `<>` `<` `>` `<=` `>=` | identiques | Conservés. |
| `Is` (comparaison d'objets) | `Is`, `IsNot` | Conservé ; ajout de `IsNot` (plus lisible que `Not ... Is`). |
| `Like` | `Like` | Conservé (sensible à `Option Compare`). |
| `And` `Or` `Not` `Xor` | identiques | ⚠️ Sur des `Boolean` : opérateurs **logiques** ; sur des entiers : **bit-à-bit**. En VB6, `And`/`Or` n'étaient **jamais court-circuités**. |
| *(pas d'équivalent)* | `AndAlso`, `OrElse` | **Nouveau** : opérateurs logiques **court-circuités**. ⚠️ Ne pas remplacer mécaniquement `And`→`AndAlso` si le second membre a des **effets de bord** (appel de fonction, etc.). |
| `Eqv` | *(supprimé)* | ⚠️ **Retiré.** Remplacer `a Eqv b` par `a = b` (booléens) ou `Not (a Xor b)`. |
| `Imp` | *(supprimé)* | ⚠️ **Retiré.** Remplacer `a Imp b` par `(Not a) Or b`. |
| *(pas d'équivalent)* | `+=` `-=` `*=` `/=` `\=` `^=` `&=` `<<=` `>>=` | **Nouveaux** opérateurs d'affectation composée. |
| *(pas d'équivalent)* | `<<` `>>` | **Nouveaux** opérateurs de décalage de bits. |

---

## 4. Structures de contrôle

| VB6 | VB.NET | ⚠️ Piège / Note |
|---|---|---|
| `If ... Then ... Else ... End If` | identique | Conservé. |
| `Select Case ...` | identique | Conservé. |
| `For i = a To b [Step s] ... Next` | identique | Conservé. ⚠️ Attention aux **bornes de tableau** (0-based) dans `For i = 0 To UBound(a)`. |
| `For Each x In coll ... Next` | identique | Conservé. |
| `Do ... Loop` (`While`/`Until`) | identique | Conservé. |
| `While ... Wend` | `While ... End While` | ⚠️ **`Wend` supprimé** → `End While`. |
| `With obj ... End With` | identique | Conservé. |
| `GoTo label` | identique | Conservé (mais déconseillé). |
| `GoSub label ... Return` | *(supprimé)* | ⚠️ **`GoSub`/`Return` retiré.** Refactoriser en **procédure** (`Sub`/`Function`). |
| `On expr GoTo l1, l2` | *(supprimé)* | ⚠️ **`On...GoTo`/`On...GoSub` retirés.** Réécrire en `Select Case`. |
| `Exit For / Do / Sub / Function` | identiques | Conservés. Ajout de `Exit Property`, `Exit Try`. |
| *(pas d'équivalent)* | `Continue For / Do / While` | **Nouveau** : passer à l'itération suivante. |

---

## 5. Procédures, paramètres et appels

| VB6 | VB.NET | ⚠️ Piège / Note |
|---|---|---|
| `Sub Foo(x As Long)` | `Sub Foo(ByVal x As Integer)` | ⚠️⚠️ **LE piège n°1.** Sans mot-clé, VB6 passe **`ByRef`**, VB.NET passe **`ByVal`**. Une conversion qui omet le mot-clé **change la sémantique**. **Toujours expliciter `ByVal`/`ByRef`.** Annexe B.1. |
| `Function Bar() As Long` | `Function Bar() As Integer` | Conservé (attention au type de retour). |
| `Bar = résultat` (retour) | `Bar = résultat` **ou** `Return résultat` | Conservé ; **`Return valeur`** est nouveau et recommandé. |
| `Call Foo(a, b)` | `Foo(a, b)` | `Call` devient **optionnel** (souvent supprimé). |
| `Foo a, b` (sans parenthèses) | `Foo(a, b)` | ⚠️ Les **parenthèses deviennent obligatoires** lors d'un appel. |
| `Optional x As Variant` + `IsMissing(x)` | `Optional x As Object = Nothing` | ⚠️ Un paramètre `Optional` **doit avoir une valeur par défaut**. `IsMissing` **ne fonctionne pas** pour les types valeur. Annexe B.5. |
| `ParamArray a()` | `ParamArray a()` | Conservé (mais **`ByVal`** et tableau typé). |
| *(pas d'équivalent)* | `Overloads`, paramètres par défaut typés | **Nouveau** : surcharge de méthodes. |

> ⚠️ **Astuce de préparation (module 5.2)** : rendre les `ByVal`/`ByRef` **explicites côté VB6
> *avant* de migrer**. Le code reste valide en VB6 et la sémantique devient non ambiguë pour
> l'outil de conversion comme pour l'IA.

---

## 6. Propriétés

| VB6 | VB.NET | ⚠️ Piège / Note |
|---|---|---|
| `Property Get Nom() As T` | `Property Nom() As T` (bloc `Get`) | Fusionné dans une **propriété unique**. |
| `Property Let Nom(v As T)` | bloc `Set(value As T)` | ⚠️ **`Let` et `Set` fusionnent** dans le bloc `Set` unique. |
| `Property Set Nom(o As Obj)` | bloc `Set(value As Obj)` | ⚠️ Idem : plus de distinction valeur/objet. |
| Propriété par défaut (`TextBox1 = "x"` ⇒ `.Text`) | *(supprimée sauf indexée)* | ⚠️ **Les propriétés par défaut sans paramètre disparaissent.** Il faut écrire `.Text` explicitement. Seules les propriétés **paramétrées** (indexeurs) restent. Annexe B.7. |
| `obj!Champ` (opérateur bang) | `obj.Item("Champ")` | ⚠️ Le `!` dépend des propriétés par défaut indexées ; à expliciter. |

---

## 7. Gestion d'erreurs

> 📖 Vue d'ensemble au **module 2.6** ; mise en œuvre détaillée au module dédié aux exceptions.

| VB6 | Équivalent direct | Idiomatique .NET | ⚠️ Piège / Note |
|---|---|---|---|
| `On Error GoTo label` | conservé | `Try ... Catch ... Finally` | Le modèle `On Error` **fonctionne encore**, mais `Try/Catch` est la cible. Ne pas **mélanger les deux** dans une même procédure. |
| `On Error Resume Next` | conservé | à **éliminer** | ⚠️ **Masque les erreurs** : première cause de bugs silencieux après migration. Annexe B.6. |
| `On Error GoTo 0` | conservé | — | Désactive le gestionnaire courant. |
| `Resume` / `Resume Next` | conservés (modèle `On Error`) | — | Pas d'équivalent direct en `Try/Catch` ; repenser la logique. |
| `Err.Number` / `Err.Description` | `Err.Number` / `Err.Description` | `ex.Message`, `ex.HResult` | L'objet `Err` subsiste ; préférer l'objet `Exception` (`Catch ex As Exception`). |
| `Err.Raise n, , "msg"` | `Err.Raise(...)` | `Throw New Exception("msg")` | ⚠️ Préférer **`Throw`** + type d'exception précis. |
| `Error nn` (instruction) | *(éviter)* | `Throw` | Forme héritée à remplacer. |

---

## 8. Fonctions intrinsèques

> Sauf mention contraire, les fonctions « directes » proviennent de l'espace `Microsoft.VisualBasic`
> (livré avec le runtime) et **fonctionnent toujours**. La colonne idiomatique tend vers le BCL .NET.

### 8.1 Chaînes de caractères

| VB6 | Direct | Idiomatique .NET | ⚠️ Piège / Note |
|---|---|---|---|
| `Len(s)` | `Len(s)` | `s.Length` | — |
| `Left(s, n)` | `Left(s, n)` | `s.Substring(0, n)` | ⚠️ `Substring` lève une exception si `n > Length` (pas `Left`). |
| `Right(s, n)` | `Right(s, n)` | `s.Substring(s.Length - n)` | Idem prudence sur les bornes. |
| `Mid(s, start, len)` | `Mid(s, start, len)` | `s.Substring(start - 1, len)` | ⚠️ **`Mid` est 1-based**, `Substring` est **0-based** (`start - 1`). |
| `Mid(s, p, n) = "x"` (instruction) | `Mid(s, p, n) = "x"` | manipulation via `StringBuilder` | Affectation par `Mid` : pas d'équivalent BCL direct. |
| `InStr([d,] s, t)` | `InStr(s, t)` | `s.IndexOf(t)` | ⚠️ **`InStr` renvoie 0** si absent (1-based) ; **`IndexOf` renvoie -1** (0-based). **Très fréquent comme bug.** |
| `InStrRev(s, t)` | `InStrRev(s, t)` | `s.LastIndexOf(t)` | Mêmes décalages 0/1-based. |
| `Replace(s, a, b)` | `Replace(s, a, b)` | `s.Replace(a, b)` | — |
| `UCase(s)` / `LCase(s)` | idem | `s.ToUpper()` / `s.ToLower()` | ⚠️ Sensibilité à la **culture** : préférer `ToUpperInvariant()` pour des comparaisons techniques. |
| `Trim(s)` / `LTrim` / `RTrim` | idem | `s.Trim()` / `s.TrimStart()` / `s.TrimEnd()` | ⚠️ `Trim` VB6 retire les **espaces** ; `.Trim()` retire **tous les blancs** (tab, etc.). |
| `Space(n)` | `Space(n)` | `New String(" "c, n)` | — |
| `String(n, c)` | `String(n, c)` | `New String(c, n)` | — |
| `StrComp(a, b)` | `StrComp(a, b)` | `String.Compare(a, b)` | Tenir compte de `Option Compare` / `StringComparison`. |
| `StrReverse(s)` | `StrReverse(s)` | `New String(s.Reverse().ToArray())` | — |
| `Asc(c)` / `Chr(n)` | `Asc` / `Chr` | `AscW` / `ChrW` | ⚠️ `Asc`/`Chr` sont **ANSI** (dépendent de la page de code) ; `AscW`/`ChrW` sont **Unicode**. |
| `Format(v, "...")` | `Format(v, "...")` | `v.ToString("...")` / `String.Format(...)` | ⚠️ **Sensible à la culture** (séparateur décimal, format de date). Annexe B.10. |

### 8.2 Conversion de types

| VB6 | Direct | Idiomatique .NET | ⚠️ Piège / Note |
|---|---|---|---|
| `CInt(x)` | `CInt(x)` | `CInt(x)` / `Convert.ToInt32(x)` | ⚠️ **`CInt` convertit désormais vers 32 bits**. Arrondi « **au pair le plus proche** » (banquier) en VB6 **et** VB.NET. |
| `CLng(x)` | `CLng(x)` | `CLng(x)` | ⚠️ Convertit vers **64 bits** en VB.NET. |
| `CCur(x)` | *(supprimé)* | `CDec(x)` | ⚠️ **`CCur` retiré** → utiliser **`CDec`**. |
| `CDbl` / `CSng` / `CStr` / `CBool` / `CDate` / `CByte` | identiques | identiques | Conservés. ⚠️ `CDbl`/`CDate` sont **sensibles à la culture**. Annexe B.10. |
| *(nouveaux)* | — | `CShort`, `CDec`, `CObj`, `CChar`, `CUInt`, `CULng`, `CUShort`, `CSByte` | **Nouveaux** convertisseurs. |
| `Val(s)` | `Val(s)` | `Val(s)` / `Double.Parse(s, Invariant)` | ✅ `Val` est **toujours invariant de culture** (point décimal) : utile pour lire des données « techniques ». |
| `Str(n)` | `Str(n)` | `n.ToString(Invariant)` | ✅ `Str` toujours invariant (point + espace de signe). |
| `Hex(n)` / `Oct(n)` | identiques | `Convert.ToString(n, 16)` | Conservés. |
| *(transtypage objet)* | — | `CType(o, T)`, `DirectCast(o, T)`, `TryCast(o, T)` | **Nouveaux.** `DirectCast` (cast strict), `TryCast` (renvoie `Nothing` si échec). |

### 8.3 Date et heure

| VB6 | Direct | Idiomatique .NET | ⚠️ Piège / Note |
|---|---|---|---|
| `Now` | `Now` | `DateTime.Now` | — |
| `Date` (date du jour) | `Today` | `DateTime.Today` | ⚠️ La **fonction** `Date` devient `Today` (`Date` est désormais le **type**). |
| `Time` | `TimeOfDay` | `DateTime.Now.TimeOfDay` | — |
| `Timer` | `Timer` | `Stopwatch` (mesures précises) | Conservé (secondes depuis minuit). |
| `DateAdd("d", n, d)` | `DateAdd(...)` | `d.AddDays(n)` | Conservé ; méthodes `Add*` plus lisibles. |
| `DateDiff("d", a, b)` | `DateDiff(...)` | `(b - a).TotalDays` (`TimeSpan`) | Conservé. |
| `DatePart("yyyy", d)` | `DatePart(...)` | `d.Year` (etc.) | Conservé. |
| `DateSerial(y, m, d)` | `DateSerial(...)` | `New DateTime(y, m, d)` | Conservé. |
| `Year/Month/Day/Hour/Minute/Second(d)` | identiques | `d.Year` / `d.Month` / … | Conservés ; propriétés plus idiomatiques. |
| `Weekday(d)` | `Weekday(d)` | `d.DayOfWeek` | ⚠️ Numérotation différente (`DayOfWeek` est une énumération 0-based dimanche). |
| `CDate(s)` / `Format(d, "...")` | identiques | `DateTime.Parse` / `d.ToString("...")` | ⚠️ **Culture !** Format de date et séparateurs. Annexe B.10. |

### 8.4 Mathématiques

| VB6 | Direct | Idiomatique .NET | ⚠️ Piège / Note |
|---|---|---|---|
| `Abs(x)` | `Abs(x)` | `Math.Abs(x)` | — |
| `Sgn(x)` | `Sgn(x)` | `Math.Sign(x)` | — |
| `Sqr(x)` | `Sqr(x)` | `Math.Sqrt(x)` | ⚠️ **Renommage** : `Sqr` → `Sqrt`. |
| `Atn(x)` | `Atn(x)` | `Math.Atan(x)` | ⚠️ **Renommage** : `Atn` → `Atan`. |
| `Exp` / `Log` / `Sin` / `Cos` / `Tan` | identiques | `Math.Exp` / `Math.Log` / … | Conservés ; versions `Math.*` recommandées. |
| `Int(x)` | `Int(x)` | `Math.Floor(x)` | ⚠️ `Int` arrondit vers **−∞**. |
| `Fix(x)` | `Fix(x)` | `Math.Truncate(x)` | ⚠️ `Fix` tronque vers **zéro** (≠ `Int` pour les négatifs). |
| `Rnd()` / `Randomize` | identiques | classe `Random` | Conservés ; `Random` recommandé (plusieurs flux, thread-safety à gérer). |

### 8.5 Tableaux

| VB6 | Direct | Idiomatique .NET | ⚠️ Piège / Note |
|---|---|---|---|
| `UBound(a)` | `UBound(a)` | `a.GetUpperBound(0)` / `a.Length - 1` | — |
| `LBound(a)` | `LBound(a)` | `0` | ⚠️ **Toujours 0** (plus d'`Option Base`). Annexe B.4. |
| `Array(1, 2, 3)` | `Array(...)` | `{1, 2, 3}` (littéral) | Littéraux de tableau recommandés. |
| `Split(s, ",")` | `Split(s, ",")` | `s.Split(","c)` | ⚠️ Surcharges et options (vides, limites) différentes. |
| `Join(a, ",")` | `Join(a, ",")` | `String.Join(",", a)` | — |
| `Filter(a, "x")` | `Filter(a, "x")` | `a.Where(Function(e) e.Contains("x"))` (LINQ) | LINQ recommandé (module 17.5). |
| `Erase a` | `Erase a` | `a = Nothing` / `Array.Clear(a, 0, a.Length)` | ⚠️ `Erase` libère ; sémantique à clarifier selon l'intention. |

### 8.6 Test de type et inspection

| VB6 | Direct | Idiomatique .NET | ⚠️ Piège / Note |
|---|---|---|---|
| `IsNumeric(x)` / `IsDate(x)` / `IsArray(x)` | identiques | (conservés) / `Double.TryParse`, etc. | Conservés. |
| `IsEmpty(v)` | *(sans objet)* | `x Is Nothing` / valeur par défaut | ⚠️ **Plus de `Empty`** (Variant disparu). Annexe B.5. |
| `IsNull(v)` | *(BD)* | `IsDBNull(x)` / `x Is Nothing` | ⚠️ Distinguer `Nothing` (référence) et `DBNull.Value` (base de données). |
| `IsMissing(arg)` | *(sans objet)* | comparer à la valeur par défaut | ⚠️ Ne s'applique plus aux paramètres `Optional` typés. Annexe B.5. |
| `IsObject(x)` | — | `TypeOf x Is Object` | — |
| `VarType(v)` | `VarType(v)` | `x.GetType()` | Variant disparu : raisonner par type réel. |
| `TypeName(x)` | `TypeName(x)` | `x.GetType().Name` | Conservé. |
| *(pas d'équivalent)* | — | `TypeOf x Is T`, `x.GetType()`, `GetType(T)` | **Nouveaux** (réflexion, tests de type). |
| *(pas d'équivalent)* | — | `IsNothing(x)` | **Nouveau** (équivaut à `x Is Nothing`). |

### 8.7 Interaction utilisateur

| VB6 | Direct | Idiomatique .NET | ⚠️ Piège / Note |
|---|---|---|---|
| `MsgBox "x", vbYesNo` | `MsgBox(...)` | `MessageBox.Show(...)` (WinForms) | `MsgBox` conservé ; `MessageBox.Show` plus riche. |
| `InputBox("?")` | `InputBox("?")` | (conservé) | Pas d'équivalent BCL direct → `InputBox` reste la voie simple. |
| `Beep` | `Beep` | `Console.Beep()` | — |

### 8.8 Fichiers — E/S héritées

> ⚠️ Toutes ces formes **fonctionnent encore** via `Microsoft.VisualBasic`, mais **`System.IO`**
> est la cible idiomatique (gestion des ressources via `Using`, encodages explicites). Module 16.6.

| VB6 | Direct | Idiomatique .NET | ⚠️ Piège / Note |
|---|---|---|---|
| `Open f For Input As #1` | `FileOpen(...)` | `StreamReader` / `StreamWriter` | ⚠️ Numéros de fichiers (`#1`) → objets. Encapsuler avec `Using`. |
| `Print #1, x` / `Write #1, x` | `Print(...)` / `Write(...)` | `writer.WriteLine(x)` | ⚠️ `Write #` met des guillemets/virgules (format spécifique). |
| `Input #1, x` / `Line Input #1, s` | `Input(...)` / `LineInput(...)` | `reader.ReadLine()` | — |
| `Get #1, , rec` / `Put #1, , rec` | `FileGet` / `FilePut` | `BinaryReader` / `BinaryWriter` | ⚠️ Fichiers binaires/aléatoires : **disposition mémoire** des structures critique (Annexe B.2, module 16.5). |
| `EOF(1)` / `LOF(1)` / `FreeFile` / `Seek` | identiques | propriétés de flux | Conservés. |
| `Close #1` | `FileClose(1)` | `.Dispose()` (via `Using`) | — |
| `Dir(...)`, `Kill`, `MkDir`, `RmDir`, `Name`, `FileCopy` | identiques | `System.IO.File` / `Directory` | `File.Delete`, `Directory.CreateDirectory`, `File.Move`, `File.Copy`. |
| `FileLen`, `FileDateTime`, `GetAttr`, `SetAttr` | identiques | `FileInfo` (propriétés) | — |
| `CurDir`, `ChDir`, `ChDrive` | identiques | `Directory.GetCurrentDirectory` / `SetCurrentDirectory` | ⚠️ Le « lecteur courant » n'a plus vraiment de sens. |
| `FileSystemObject` (Scripting) | RCW d'interop | `System.IO` | ⚠️ Remplacer le **FSO COM** par `System.IO`. Module 16. |

---

## 9. Objets globaux

| VB6 | Direct | Idiomatique .NET | ⚠️ Piège / Note |
|---|---|---|---|
| `App.Path` | — | `Application.StartupPath` / `My.Application.Info.DirectoryPath` | ⚠️ Plusieurs cibles selon le besoin (dossier de l'exe, de l'assembly…). |
| `App.EXEName` / `App.Title` | — | `Application.ProductName` / nom d'assembly | — |
| `App.PrevInstance` | — | `Mutex` / `Process.GetProcessesByName(...)` | ⚠️ Plus d'équivalent direct : gérer l'instance unique soi-même. |
| `App.Major/Minor/Revision` | — | `My.Application.Info.Version` | — |
| `Screen.Width` / `Screen.Height` | — | `Screen.PrimaryScreen.Bounds` | — |
| `Screen.MousePointer` | — | `Cursor` / `Cursor.Current` | — |
| `Printer` / `Printers` | — | `PrintDocument`, `PrinterSettings` (`System.Drawing.Printing`) | ⚠️ **Refonte complète** du modèle d'impression. Voir module Forms / impression. |
| `Clipboard.SetText` / `.GetText` | — | `Clipboard.SetText` (WinForms) / `My.Computer.Clipboard` | API conservée mais espace de noms différent. |
| `Forms` (collection) | — | `Application.OpenForms` / `My.Application.OpenForms` | — |
| `Err` (objet) | `Err` | `Exception` (objet `ex`) | Conservé sous le modèle `On Error`. |
| `Debug.Print x` | — | `Debug.WriteLine(x)` (`System.Diagnostics`) | ⚠️ **Renommage** : `Print` → `WriteLine`. |
| `Debug.Assert c` | — | `Debug.Assert(c)` | Conservé. |
| *(pas d'équivalent)* | — | `My.Computer`, `My.Application`, `My.Settings`, `My.Resources` | **Nouvel** espace `My` (raccourcis pratiques). |

---

## 10. Constantes intrinsèques

| VB6 | Direct | Idiomatique .NET | ⚠️ Piège / Note |
|---|---|---|---|
| `vbCrLf` | `vbCrLf` | `ControlChars.CrLf` / `Environment.NewLine` | ⚠️ `Environment.NewLine` dépend de la plateforme (toujours CrLf sous Windows). |
| `vbCr` / `vbLf` / `vbTab` / `vbBack` / `vbFormFeed` | identiques | `ControlChars.Cr` / `.Lf` / `.Tab` … | Conservés. |
| `vbNullChar` | `vbNullChar` | `ControlChars.NullChar` | — |
| `vbNullString` | `vbNullString` | `""` / `String.Empty` | ⚠️ **Subtil** : `vbNullString` est un **pointeur nul**, `""` une chaîne vide. La différence compte lors d'appels d'**API** (P/Invoke). |
| `vbObjectError` | `vbObjectError` | (base pour codes d'erreur) | Conservé. |
| `vbYesNo`, `vbOKCancel`, `vbCritical`, `vbExclamation`, `vbInformation` | identiques | `MsgBoxStyle` / `MessageBoxButtons` + `MessageBoxIcon` | Conservés (combinables). |
| `vbYes`, `vbNo`, `vbOK`, `vbCancel` (résultats) | identiques | `MsgBoxResult` / `DialogResult` | — |
| `vbModal` / `vbModeless` | — | `.ShowDialog()` / `.Show()` | ⚠️ Le caractère modal se choisit par **la méthode appelée**. |
| `vbBinaryCompare` / `vbTextCompare` | identiques | `StringComparison.*` | — |
| `vbRed`, `vbBlue`, `vbGreen`, … | — | `Color.Red`, `Color.Blue`, … (`System.Drawing`) | ⚠️ Couleurs **OLE_COLOR** (entiers BGR) → structure `Color`. `ColorTranslator` pour convertir. Annexe C. |
| `True` / `False` | identiques | identiques | ⚠️ `True` valait **-1** en VB6 ; conversions arithmétiques et bit-à-bit à revérifier. Annexe B.8. |
| `Nothing` | `Nothing` | `Nothing` | Conservé. ⚠️ Remplace aussi `Null`/`Empty` du `Variant`, **mais** la valeur « base de données » est `DBNull.Value`. |

---

## 11. Instructions diverses et formes supprimées

| VB6 | VB.NET | ⚠️ Piège / Note |
|---|---|---|
| `Set obj = Nothing` | `obj = Nothing` | ⚠️ En VB6, libérait l'objet **immédiatement** (`Class_Terminate`). En .NET, le **GC** décide *quand* : prévoir `Dispose()` / `Using` pour les ressources. **Piège n°1**, module 2.2 et Annexe B.3. |
| `End` (arrêt brutal du programme) | `Application.Exit()` / `Environment.Exit(0)` | ⚠️ `End` était **abrupt** (pas de finalisation). À remplacer par une sortie propre. |
| `Stop` | `Stop` | Conservé (point d'arrêt). |
| `DoEvents` | `Application.DoEvents()` | Conservé, **toujours déconseillé** (préférer `Async/Await`, `Task`, `BackgroundWorker`). |
| `Load Form1` | `Dim f As New Form1` | ⚠️ Les formulaires ne se « chargent » plus ainsi. Modules 13-14. |
| `Unload Me` | `Me.Close()` | ⚠️ Et libérer via `Dispose()` si nécessaire. |
| `Me.Show` / `Me.Hide` | `Me.Show()` / `Me.Hide()` | Conservés (parenthèses obligatoires). |
| `Print` (sur formulaire/Printer) | dessin via `Graphics` / `PrintDocument` | ⚠️ Le dessin direct sur formulaire disparaît : modèle `Paint`/`Graphics`. |
| `Line` / `Circle` / `PSet` (méthodes graphiques) | `Graphics.DrawLine` / `DrawEllipse` / … | ⚠️ Méthodes graphiques VB6 retirées → **GDI+** (`System.Drawing`). |
| `Rem` / `'` (commentaires) | `'` (et `Rem`) | Conservés. |
| `Call` (mot-clé) | optionnel | Souvent supprimé. |

---

## 🔗 Pour aller plus loin

Ce tableau est une **carte de surface**. Les points marqués ⚠️ sont approfondis là où ils  
comptent vraiment :

- **Annexe B — Catalogue des pièges silencieux** ⚠️ : pour chaque piège (ByRef, entiers
  redimensionnés, finalisation, 0-based, `Variant`, `On Error Resume Next`, propriétés par
  défaut, `True = -1`, dates/arrondis, culture) → **symptôme, cause, correction**.
- **Annexe C — Correspondance des contrôles VB6 → Windows Forms** : contrôle par contrôle,
  propriétés/événements, *control arrays*, couleurs `OLE_COLOR`.
- **Annexe D — Correspondance des types de données** : tailles, plages, et règles de conversion
  (`CType`, `DirectCast`, `CInt`/`CLng`/`CDec`).
- **Module 16 — Interop COM & API** : `Declare` → `P/Invoke`, marshaling ANSI/Unicode,
  `Type` → `Structure` avec `StructLayout`, FSO → `System.IO`.
- **Module 17 — Valider et refactoriser** : activer `Option Strict` progressivement et passer
  des formes « directes » aux formes **idiomatiques** (LINQ, génériques, espace `My`).

---

**Juin 2026**  
**Cible** : .NET Framework 4.7.2 (le « pont » depuis VB6)  
**Licence** : Creative Commons BY-NC-SA 4.0

⏭️ [**Catalogue des pièges silencieux** — le « top des bugs de migration »](/annexes/pieges-silencieux/README.md)
