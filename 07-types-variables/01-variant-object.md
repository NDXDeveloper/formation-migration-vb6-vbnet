🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.1 — `Variant` → `Object` : conversions, boxing, `CType`/`DirectCast` ⚠️

> **Chapitre 7 — Types, variables et déclarations** · Partie 3 — Migrer le langage
>
> Le `Variant` est le symbole de la souplesse de VB6 : une variable qui peut **tout** contenir.
> Son successeur naturel en .NET est `Object`. La ressemblance est réelle… mais elle masque
> trois différences de fond — le **boxing**, le **typage exact à la relecture**, et la **liaison
> tardive** — qui sont autant de sources de régressions silencieuses.

---

## 🧭 Rappel : qu'est-ce qu'un `Variant` en VB6 ?

En VB6, le `Variant` est le **type universel**. Il peut contenir un nombre, une chaîne, une date,  
un objet, un tableau, et même des **états particuliers** (`Empty`, `Null`, `Missing`). C'est aussi  
le **type par défaut** : sans `Option Explicit`, toute variable non déclarée est un `Variant`, et
`Dim x` sans `As` en crée un.

```vb
' VB6
Dim v As Variant
v = 42              ' un entier
v = "Bonjour"       ' puis une chaîne
v = #6/29/2026#     ' puis une date (littéral VB au format US M/J/AAAA)
Set v = New Client  ' puis un objet — Set requis pour affecter une référence
```

Sous le capot, un `Variant` transporte **en lui-même** une étiquette de type (interrogeable via
`VarType(v)` ou `TypeName(v)`). On le retrouve partout dans le code VB6 : paramètres optionnels,
retours de valeurs hétérogènes, **liaison tardive**, et interopérabilité COM (le type `VARIANT`).

---

## 🎯 La cible : `Object`, la racine de tout en .NET

En .NET, **`System.Object` est la racine de la hiérarchie de types** : *tout* en dérive. Une  
variable `Object` peut donc référencer n'importe quoi. C'est la traduction que retiennent les  
outils de migration : **`Variant` → `Object`**.

```vb
' VB.NET
Dim o As Object
o = 42
o = "Bonjour"
o = New Client()
```

> ⚠️ **Mais `Object` n'est pas un `Variant`.** Trois écarts changent le comportement :
> 1. les **types valeur** (`Integer`, `Double`, `Date`, structures…) subissent un **boxing** ;
> 2. pour **relire** la valeur, il faut le **bon type**, et `CType` ≠ `DirectCast` ;
> 3. les **états** du `Variant` (`Empty`/`Null`/`Missing`) n'existent plus tels quels
>    (→ traité en **7.7**).

---

## 📦 Le *boxing* : le mécanisme invisible

C'est **le** concept nouveau de cette section. En .NET, ranger un **type valeur** dans une variable
`Object` (ou une interface) déclenche un **boxing** (« mise en boîte ») : la valeur est **copiée
sur le tas**, emballée dans un objet, et `o` en garde une référence. La relire impose un
**unboxing** (déballage), via un cast.

```vb
' Type VALEUR → boxing (allocation sur le tas)
Dim n As Integer = 42
Dim o As Object = n          ' boxing : 42 est emballé dans un objet du tas

' Relecture → unboxing : le type doit correspondre EXACTEMENT
Dim n2 As Integer = DirectCast(o, Integer)   ' ✅ OK : o contient bien un Integer
Dim l  As Long    = DirectCast(o, Long)      ' ❌ InvalidCastException ! (Integer ≠ Long)
Dim l2 As Long    = CType(o, Long)           ' ✅ OK : CType convertit (déballe puis élargit)
```

À l'inverse, un **type référence** (une instance de classe, une `String`) n'est **pas** *boxé* :
il est déjà une référence.

```vb
Dim client As New Client()
Dim o As Object = client                 ' PAS de boxing : 'client' est déjà une référence
Dim c2 As Client = DirectCast(o, Client) ' simple cast de référence
```

> 🧠 **Deux conséquences pratiques :**
> - **Performance.** En VB6, le `Variant` « emballait » déjà tout, sans qu'on y pense. En .NET,
>   laisser des valeurs en `Object` provoque un **boxing/unboxing permanent** : allocations sur le
>   tas et pression sur le *garbage collector*. Dans une boucle serrée, c'est nettement plus lent
>   qu'un code **typé**.
> - **Règle du type exact.** L'unboxing **n'élargit pas**. Un `Integer` *boxé* ne se déballe pas
>   directement en `Long` avec `DirectCast` — il faut un `Integer`, ou passer par `CType`.

---

## 🔍 Relire le contenu : `DirectCast`, `CType`, `TryCast`

Quand une valeur est dans un `Object`, trois opérateurs permettent d'en récupérer le type réel.
**Les confondre est une cause classique de bugs de migration** (et les outils automatiques
choisissent souvent `CType` par sécurité, ce qui masque parfois un vrai problème).

| | `DirectCast` | `CType` | `TryCast` |
|---|---|---|---|
| **Rôle** | cast **strict** (aucune conversion) | **conversion** (élargit, rétrécit, *parse*, opérateurs définis) | cast strict qui **échoue en douceur** |
| **Exige le type exact ?** | oui | non (tente la conversion) | oui |
| **En cas d'échec** | lève `InvalidCastException` | lève si la conversion est impossible | renvoie `Nothing` |
| **Types valeur** | oui (unbox exact) | oui (unbox + conversion) | ❌ (réf. et `Nullable` seulement) |
| **Performance** | la plus rapide | plus lourde (machinerie de conversion) | rapide |
| **Quand l'utiliser** | tu **connais** le type réel | tu dois **transformer** la valeur | « **peut-être** ce type ? » (réf.) |

```vb
Dim o As Object = "123"

Dim a As Integer = DirectCast(o, Integer) ' ❌ o est une String, pas un Integer boxé → exception
Dim b As Integer = CType(o, Integer)      ' ✅ 123 : CType convertit la chaîne en nombre
Dim s As String  = DirectCast(o, String)  ' ✅ o EST une String : cast direct, rapide

' TryCast : pour les types RÉFÉRENCE, sans lever d'exception
Dim c As Client = TryCast(o, Client)      ' renvoie Nothing si o n'est pas un Client
If c IsNot Nothing Then c.Facturer()
```

> ⚠️ **`CType` fait parfois trop.** Parce qu'il *convertit*, il peut réussir là où le code
> dissimule en réalité une erreur de type. `CType("12,5", Double)` dépend en plus des **paramètres
> régionaux** (séparateur décimal) — un piège sensible à la culture (voir **Annexe B.10**).
> Règle de pouce : **`DirectCast` quand tu connais le type** (plus rapide, et il *échoue fort* en
> cas d'erreur de migration), **`CType` seulement quand une vraie conversion est nécessaire**.

---

## 🏷️ Interroger le type : `VarType` → `TypeOf…Is` / `GetType`

Beaucoup de code VB6 **branche sur le type** d'un `Variant` :

```vb
' VB6
Select Case VarType(v)
    Case vbInteger, vbLong : Traiter_Nombre v
    Case vbString          : Traiter_Texte v
    Case vbObject          : Traiter_Objet v
End Select
```

En VB.NET, on dispose d'équivalents plus sûrs et plus lisibles :

```vb
' VB.NET — idiomatique : TypeOf … Is
If TypeOf v Is Integer Then
    Traiter_Nombre(v)
ElseIf TypeOf v Is String Then
    Traiter_Texte(v)
End If

' Comparaison de type exacte
If v IsNot Nothing AndAlso v.GetType() Is GetType(Integer) Then ...

' Compatibilité : VarType() et TypeName() existent ENCORE (espace Microsoft.VisualBasic)
Select Case VarType(v) : '... : End Select
```

> 💡 `TypeOf v Is Integer` fonctionne aussi sur une valeur *boxée* (il teste le type emballé).
> Pense à **garder le `Nothing`** : `TypeOf Nothing Is String` vaut `False`, donc protège les
> accès (`v IsNot Nothing AndAlso …`).

---

## 🔗 Liaison tardive : `Object` + `Option Strict Off`

En VB6, un `Variant` (ou un `Object`) permet la **liaison tardive** : on appelle une méthode dont  
le type n'est connu qu'à l'exécution.

```vb
' VB6 — liaison tardive
Dim x As Object
Set x = CreateObject("Excel.Application")
x.Visible = True            ' résolu à l'exécution
```

En VB.NET, ce code migre vers `Object`, mais son sort dépend d'une option de compilation :

- **`Option Strict Off`** (point de départ conseillé pour du *legacy*) : la liaison tardive sur
  `Object` est **tolérée**, comme en VB6.
- **`Option Strict On`** (objectif de fin de migration) : appeler un membre sur `Object` devient
  une **erreur de compilation** — il faut d'abord **caster vers le vrai type**.

```vb
' VB.NET, Option Strict On : il faut typer avant d'appeler
Dim app As Excel.Application = DirectCast(x, Excel.Application)
app.Visible = True
```

> 🎯 La trajectoire est claire : on **démarre en `Option Strict Off`** pour ne rien casser, puis on
> **réintroduit les types** et on l'active **progressivement** (voir **module 17.1**). La liaison
> tardive et l'interop COM sont traitées au **module 16.3**.

---

## ⚠️ Les pièges silencieux de cette traduction

- **Unboxing au mauvais type → `InvalidCastException`.** Un `Short` *boxé* ne se déballe pas en
  `Integer` avec `DirectCast`. L'erreur n'apparaît qu'à **l'exécution**, sur un chemin de code
  parfois rare. (Le décalage de tailles est le sujet de **7.2**.)
- **Tout laisser en `Object`** annule le bénéfice de la migration : on perd la vérification de
  type à la compilation **et** on paie le coût du *boxing*. `Object` doit rester l'exception, pas
  la règle.
- **Arithmétique sur `Object`.** Avec `Option Strict Off`, `o + 1` fonctionne par **liaison
  tardive** et suit des règles de coercition VB parfois surprenantes ; des différences subtiles
  existent (voir **chapitre 9 — Opérateurs**).
- **États perdus.** `Empty`, `Null`, `Missing` ne se ramènent pas mécaniquement à `Nothing` (et la
  « valeur nulle » d'une base de données devient `DBNull`). C'est l'objet de **7.7**.

---

## 🧭 Stratégie de migration recommandée

1. **Laisser l'outil traduire** `Variant` → `Object` : c'est la bonne cible mécanique.
2. **Re-typer dès que possible.** Pour chaque variable `Object`, se demander : *son type est-il en
   réalité connu ?* Si oui, remplacer `Object` par le **type réel** (`Integer`, `String`,
   `Client`…). C'est le gain principal de la migration.
3. **Ne garder `Object`** que là où il est vraiment justifié : polymorphisme réel, **interop COM**,
   liaison tardive irréductible.
4. **Choisir le bon cast :** `DirectCast` quand le type est connu, `CType` pour une **vraie**
   conversion, `TryCast` pour un « peut-être ce type » sur une référence.
5. **Activer `Option Strict` progressivement** pour faire émerger, à la compilation, les
   conversions implicites restantes (module 17).

> 💡 Et toujours le filet de sécurité : les **tests de non-régression** (*golden master*,
> chapitre 5.5) révèlent, chiffres à l'appui, qu'une conversion `Object` a changé un résultat.

---

## ✅ À retenir

- **`Variant` → `Object`**, mais `Object` **n'est pas** un `Variant` : *boxing*, typage exact et
  liaison tardive changent la donne.
- **Boxing** : les **types valeur** rangés dans un `Object` sont copiés sur le tas ; l'**unboxing
  n'élargit pas** (il faut le type exact, ou `CType`).
- **`DirectCast`** = cast strict (rapide, type exact, échoue fort) ; **`CType`** = conversion
  (souple, mais peut masquer un bug et dépend parfois de la culture) ; **`TryCast`** = cast strict
  sans exception, réservé aux **références**.
- **Interroger le type** : `VarType` → `TypeOf…Is` / `GetType` (en gardant `TypeName`/`VarType`
  pour la compatibilité).
- **Objectif** : faire **disparaître** un maximum de `Object` au profit de **types explicites**.

---

## 🔗 Pour aller plus loin

- **7.2 — Entiers redimensionnés** : pourquoi un `Short` *boxé* ne se déballe pas en `Integer`.
- **7.7 — `Nothing`, `Empty`, `Null`, `Missing`** : le sort des états du `Variant`.
- **Chapitre 9 — Opérateurs** : l'arithmétique et la comparaison sur `Object` (liaison tardive).
- **Module 16.3 — *Early/late binding*** et **Module 17.1 — `Option Strict`** : éliminer la liaison
  tardive résiduelle.
- **Annexe B** (pièges silencieux) et **Annexe D** (correspondance des types).

➡️ Section suivante : **[7.2 — Entiers redimensionnés](02-entiers-redimensionnes.md)**  
⬆️ Retour au **[sommaire du chapitre 7](README.md)**

⏭️ [**Entiers redimensionnés** : `Integer`→`Short`, `Long`→`Integer`, et `Currency`→`Decimal`](/07-types-variables/02-entiers-redimensionnes.md)
