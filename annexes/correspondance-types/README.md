🔝 Retour au [Sommaire](/SOMMAIRE.md)

# D. Correspondance des types de données

`Variant`, `Integer`/`Long`, `Currency`, `Date`, `String * n`, `Boolean`… avec **tailles**,
**plages** et **règles de conversion** (`CType`, `DirectCast`, fonctions `CInt`/`CLng`/`CDec`).

> 🎯 **À quoi sert cette annexe.** L'Annexe A donne la vue *express* des types ; ici, on a les
> **chiffres exacts** (tailles, plages) et la **mécanique des conversions**. Pour le détail des
> *régressions* liées aux types (débordements, `Variant` perdu, dates/arrondis), voir le catalogue
> **Annexe B** (pièges B.2, B.5, B.8, B.9, B.10), auquel cette annexe renvoie.

> ⚠️ **La règle d'or, à retenir avant tout** : on migre les types **par taille, pas par nom**.
> `Integer` (VB6, 16 bits) ne devient **pas** `Integer` (VB.NET, 32 bits) mais **`Short`** ; `Long`
> (VB6, 32 bits) devient **`Integer`**. Conserver le nom **double silencieusement** la taille — avec
> toutes les conséquences en interop, fichiers binaires et `StructLayout` (Annexe B.2, module 16.5).

---

## 1. Tableau maître des correspondances

| Type VB6 | Type VB.NET | Type CLR | Taille | Plage / précision |
|---|---|---|---|---|
| `Byte` | `Byte` | `System.Byte` | 8 bits | 0 à 255 |
| `Integer` | **`Short`** | `System.Int16` | 16 bits | −32 768 à 32 767 |
| `Long` | **`Integer`** | `System.Int32` | 32 bits | −2 147 483 648 à 2 147 483 647 |
| *(aucun)* | `Long` | `System.Int64` | 64 bits | ≈ ±9,2 × 10¹⁸ |
| `Single` | `Single` | `System.Single` | 32 bits | ≈ ±3,4 × 10³⁸ (≈ 7 chiffres signif.) |
| `Double` | `Double` | `System.Double` | 64 bits | ≈ ±1,8 × 10³⁰⁸ (≈ 15-16 chiffres signif.) |
| `Currency` | **`Decimal`** | `System.Decimal` | 64 → **128 bits** | VB6 : ±922 337 203 685 477,5807 (**4 décimales fixes**) |
| *(Variant Decimal)* | `Decimal` | `System.Decimal` | 128 bits | ≈ ±7,9 × 10²⁸ (28-29 chiffres signif.) |
| `Boolean` | `Boolean` | `System.Boolean` | 16 → **1 bit/octet** | `True` / `False` (⚠️ `True` = **-1** en VB6) |
| `Date` | `Date` (alias `DateTime`) | `System.DateTime` | 64 bits | 01/01/0001 à 31/12/9999 (⚠️ plus stocké comme `Double`) |
| `String` (longueur variable) | `String` | `System.String` | variable | Unicode (jusqu'à ≈ 2³¹ caractères) |
| `String * n` (longueur fixe) | `String` + `<VBFixedString(n)>` | `System.String` | — | ⚠️ **Pas de type natif** longueur fixe. Voir §4. |
| `Variant` | **`Object`** | `System.Object` | 16 octets | n'importe quel type (⚠️ `Empty`/`Null`/`Missing` perdus) |
| `Object` (COM générique) | `Object` | `System.Object` | réf. | ⚠️ Sens différent (voir §4). |
| `Collection` | `Collection` | `Microsoft.VisualBasic.Collection` | réf. | Conservée ; préférer `List(Of T)` / `Dictionary`. |
| *(aucun)* | `Char` | `System.Char` | 16 bits | U+0000 à U+FFFF (un caractère Unicode) |
| *(aucun)* | `SByte` | `System.SByte` | 8 bits | −128 à 127 |
| *(aucun)* | `UShort` | `System.UInt16` | 16 bits | 0 à 65 535 |
| *(aucun)* | `UInteger` | `System.UInt32` | 32 bits | 0 à 4 294 967 295 |
| *(aucun)* | `ULong` | `System.UInt64` | 64 bits | 0 à ≈ 1,8 × 10¹⁹ |

> 💡 Les types **non signés** (`UShort`/`UInteger`/`ULong`) et **`SByte`** sont **nouveaux** en
> VB.NET (VB6 n'avait que `Byte` comme non signé). Ils ne sont **pas conformes CLS** : à réserver à
> l'interop ou à des besoins précis, pas comme remplacement systématique.

---

## 2. Caractères de déclaration de type (`%`, `&`, `!`, `#`, `@`, `$`)

> ⚠️ **Piège discret mais réel** : VB.NET conserve ces caractères, **mais leur taille suit les
> tailles VB.NET**. Le **même** caractère désigne donc un type **plus large** qu'en VB6.

| Caractère | Signification VB6 | Signification VB.NET | ⚠️ |
|---|---|---|---|
| `%` | `Integer` (16 bits) | `Integer` (**32 bits**) | Taille **doublée** silencieusement |
| `&` | `Long` (32 bits) | `Long` (**64 bits**) | Taille **doublée** silencieusement |
| `@` | `Currency` (64 bits) | `Decimal` (128 bits) | Type **et** taille changent |
| `!` | `Single` | `Single` | Inchangé |
| `#` | `Double` | `Double` | Inchangé |
| `$` | `String` | `String` | Inchangé |

Exemple : `Dim n&` déclare un entier **32 bits** en VB6 mais **64 bits** en VB.NET.

> ✅ **Recommandation** : supprimer ces caractères au profit du **typage explicite** (`As Short`,
> `As Integer`…), idéalement **dès la préparation du source VB6** (module 5.2). On lève l'ambiguïté
> *avant* la conversion.

---

## 3. Fiches par type clé

### `Variant` → `Object`
- **Taille** : 16 octets (VB6) → référence (.NET).
- **Règle** : `Variant` devient `Object` ; le *late binding* qui en découle exige `Option Strict Off`
  (à éliminer progressivement, modules 16.3 et 17.1).
- ⚠️ **Les trois états spéciaux disparaissent** : `Empty` → `Nothing` ; `Null` (base de données) →
  `DBNull.Value` ; `Missing` (paramètre omis) → valeur par défaut explicite / `Nullable(Of T)`.
  Détail et corrections : **Annexe B.5**.

### `Integer` / `Long`
- **Tailles** : `Integer` 16 bits → `Short` ; `Long` 32 bits → `Integer` ; le nouveau `Long` fait
  64 bits.
- **Règle** : mapper **par taille**. En interop / structures binaires, c'est **critique** (un champ
  de 16 bits ne doit pas devenir 32 bits). VB.NET active par défaut le **contrôle de débordement**
  (`OverflowException`). Détail : **Annexe B.2**.

### `Currency` → `Decimal`
- **VB6** : entier 64 bits **mis à l'échelle ×10000**, soit **4 décimales fixes**, exact (pas
  d'erreur binaire).
- **VB.NET** : pas de type `Currency`. Cible = **`Decimal`** (base 10, exact, idéal pour les
  montants). ⚠️ **Jamais `Double`** (virgule flottante binaire → erreurs d'accumulation et arrondis
  faux).
- **Conversion** : `CCur` **supprimé** → utiliser **`CDec`**.
- Arrondis monétaires : **Annexe B.9**.

### `Date` → `DateTime`
- **VB6** : 8 octets, stocké comme **`Double`** (partie entière = jours depuis le 30/12/1899, partie
  fractionnaire = heure). On pouvait additionner un nombre, soustraire deux dates → jours.
- **VB.NET** : **`DateTime`** est une **structure distincte** (ticks de 100 ns depuis l'an 1), **pas
  un `Double`**. Arithmétique via `.AddDays(...)` ; différence = **`TimeSpan`**.
- **Frontière avec des données héritées** stockées en double OLE : `Date.FromOADate(d)` /
  `date.ToOADate()`. Détail : **Annexe B.9**.

### `String * n` (longueur fixe)
- **VB6** : `Dim s As String * 10` → toujours 10 caractères (complétée par des espaces).
- **VB.NET** : **pas de type natif** de longueur fixe. Options :
  - `String` simple (recommandé dans la plupart des cas) ;
  - attribut `<VBFixedString(10)> Public s As String` (`Microsoft.VisualBasic`) — n'agit qu'au
    **marshaling** (interop / fichiers à structure fixe), **pas** en mémoire.
- ⚠️ Le comportement de remplissage/troncature automatique disparaît : à gérer explicitement si du
  code en dépendait.

### `Boolean`
- **VB6** : 16 bits, `True` = **-1** (tous les bits à 1), `False` = 0.
- **VB.NET** : `System.Boolean` (1 octet), `True` = 1 en interne — **mais `CInt(True)` renvoie
  toujours -1** (compatibilité), alors que `Convert.ToInt32(True)` et le marshaling donnent 1.
  Incohérence à connaître. Détail : **Annexe B.8**.

### `Object` (un faux ami)
- **VB6** : `Object` = référence d'objet **COM générique** (IDispatch), à liaison tardive.
- **VB.NET** : `Object` = **`System.Object`**, racine de **tous** les types (valeur **et**
  référence). Un `As Object` VB6 reste `As Object`, mais le *late binding* nécessite
  `Option Strict Off` (module 16.3).

### `Char` (nouveau)
- VB6 n'avait pas de type `Char` (on utilisait un `String` de longueur 1, ou un code via
  `Asc`/`Chr`). VB.NET a **`Char`** (`System.Char`, 16 bits, Unicode). Littéral : `"A"c`.
  Conversions : `CChar`, `AscW`/`ChrW` (Unicode), `Asc`/`Chr` (ANSI). Voir Annexe A, §8.1.

---

## 4. Règles de conversion

### 4.1 Trois familles de conversion (ne pas les confondre)

| Famille | Exemples | Usage | ⚠️ |
|---|---|---|---|
| **Opérateurs de cast/conversion** | `CType`, `DirectCast`, `TryCast` | Transtypage et conversions générales | Sémantiques **différentes** (voir 4.2). |
| **Fonctions de conversion VB** | `CInt`, `CLng`, `CShort`, `CByte`, `CDec`, `CDbl`, `CSng`, `CStr`, `CBool`, `CDate`, `CChar`, `CObj` | Conversion de **données** avec sémantique VB | Arrondi **banquier** ; certaines **sensibles à la culture** (4.5). |
| **Framework (BCL)** | `Convert.ToInt32`, `Integer.Parse`/`TryParse`, `.ToString()` | Conversions .NET explicites | `Convert.ToInt32(True)` = **1** (≠ `CInt`). |

### 4.2 `CType` vs `DirectCast` vs `TryCast`

| | `CType(expr, T)` | `DirectCast(expr, T)` | `TryCast(expr, T)` |
|---|---|---|---|
| **Effectue une conversion ?** | **Oui** (élargissement, restriction, opérateurs définis) | **Non** : simple cast | **Non** : simple cast |
| **Exige que le type réel soit déjà `T` ?** | Non | **Oui** (sinon `InvalidCastException`) | **Oui** (sinon **`Nothing`**) |
| **Types concernés** | valeur **et** référence | valeur (unboxing) **et** référence | **référence uniquement** |
| **Échec** | exception si impossible | `InvalidCastException` | renvoie `Nothing` |
| **Quand l'utiliser** | quand il faut **convertir** (ex. `String` → `Integer`) | quand on **connaît** le type exact (downcast, unboxing) — plus rapide | downcast **incertain** sans exception |

```vbnet
Dim o As Object = 123                         ' Integer "boxé"
Dim a As Integer = DirectCast(o, Integer)     ' OK : type exact (unboxing)

Dim s As Object = "123"
' Dim n As Integer = DirectCast(s, Integer)   ' EXCEPTION : String n'EST PAS Integer
Dim n As Integer = CType(s, Integer)          ' OK : CType convertit String → Integer

Dim b As Button = TryCast(unControle, Button) ' Nothing si ce n'est pas un Button
If b IsNot Nothing Then b.Text = "OK"
```

### 4.3 Élargissement (*widening*) vs restriction (*narrowing*)

- **Élargissement** : toujours sûr, implicite et autorisé même sous `Option Strict On`.
  `Byte → Short → Integer → Long → Decimal` ; `Integer → Single → Double` ; tout type → `Object`.
- **Restriction** : peut **perdre des données** ou **lever une exception**. Sous **`Option Strict
  On`**, elle exige une **conversion explicite** (`CType`/`CInt`/…). C'est précisément ce qui rend
  `Option Strict On` si utile en fin de migration (module 17.1) : il **force** à expliciter les
  conversions risquées.

> ⚠️ Démarrer la migration en `Option Strict Off` (toléré pour le *late binding* hérité), puis
> l'activer **progressivement** pour faire ressortir les conversions implicites dangereuses
> (module 6.4 puis 17.1).

### 4.4 Arrondi : `CInt`/`CLng` (banquier) vs `Int`/`Fix`

⚠️ Les fonctions `CInt`/`CLng` utilisent l'**arrondi au pair le plus proche** (*banker's rounding*),
**en VB6 comme en VB.NET** — un comportement souvent surprenant mais **cohérent** entre les deux :

```vbnet
CInt(0.5)   ' = 0     (vers le pair : 0)
CInt(1.5)   ' = 2     (vers le pair : 2)
CInt(2.5)   ' = 2     (vers le pair : 2)
CInt(3.5)   ' = 4     (vers le pair : 4)
```

À distinguer des fonctions de **troncature** (voir aussi Annexe A, §8.4) :

| Fonction | Comportement | `(-2,5)` → |
|---|---|---|
| `CInt(x)` | arrondi **banquier** | `-2` |
| `Int(x)` | plancher (vers **−∞**) | `-3` |
| `Fix(x)` | troncature (vers **zéro**) | `-2` |

> Pour un arrondi **commercial** explicite (demi loin de zéro), utiliser
> `Math.Round(x, n, MidpointRounding.AwayFromZero)` et **documenter la règle** (module 5.4).

### 4.5 Conversions sensibles à la culture

- **Sensibles à la culture** (séparateur décimal, format de date) : `CDbl`, `CSng`, `CDate`, `CStr`
  (d'un nombre), `Format`, `.ToString` sans `IFormatProvider`. → comportement **dépendant des
  paramètres régionaux**.
- **Invariantes** (sûres pour des **données** persistantes) : les fonctions VB **`Val`** (lecture)
  et **`Str`** (écriture), ou les méthodes BCL avec `CultureInfo.InvariantCulture`.

> ⚠️ Séparer **données** (invariant) et **affichage** (culture courante). C'est la cause du bug
> « ça marche sur ma machine ». Détail et exemples : **Annexe B.10**.

### 4.6 `Nothing` et valeurs par défaut

⚠️ En VB.NET, `Nothing` affecté à un **type valeur** ne signifie **pas « null »** mais la **valeur
par défaut** du type :

```vbnet
Dim i As Integer = Nothing   ' i = 0   (et non "null")
Dim b As Boolean = Nothing   ' b = False
Dim d As Date = Nothing      ' d = 01/01/0001 (DateTime.MinValue)
```

Pour un type **référence**, `Nothing` est bien la **référence nulle**. Pour représenter  
explicitement « **valeur ou absente** » sur un type valeur, utiliser **`Nullable(Of T)`** (ex.
`Integer?`), avec `.HasValue` et `.Value`.

### 4.7 Boxing / unboxing

Affecter un **type valeur** à une variable `Object` le **« boxe »** (copie sur le tas) ; le  
récupérer le **« déboxe »**. Le déboxage par `DirectCast` exige le **type exact** :

```vbnet
Dim o As Object = 42          ' boxing d'un Integer
Dim x As Integer = DirectCast(o, Integer)   ' OK
' Dim y As Long = DirectCast(o, Long)       ' EXCEPTION : le type boxé est Integer, pas Long
Dim y As Long = CLng(o)                      ' OK : CLng convertit
```

---

## 5. 🔗 Annexes et modules liés

- **Annexe A** — Tableau de correspondance (vue express des types, fonctions de conversion §8.2).
- **Annexe B** — Pièges silencieux : **B.2** (entiers redimensionnés), **B.5** (`Variant`),
  **B.8** (`True = -1`), **B.9** (dates / `Currency` → `Decimal`), **B.10** (culture).
- **Module 6.4** — Options de compilation (`Option Strict`/`Explicit`/`Infer`).
- **Module 16.5** — Marshaling et `StructLayout` (où les **tailles** de type sont critiques).
- **Module 17.1** — Activer `Option Strict` progressivement (force l'explicitation des conversions).

---

**Juin 2026**  
**Cible** : .NET Framework 4.7.2 (le « pont » depuis VB6)  
**Licence** : Creative Commons BY-NC-SA 4.0

⏭️ [Checklist de migration (avant / pendant / après)](/annexes/checklist-migration/README.md)
