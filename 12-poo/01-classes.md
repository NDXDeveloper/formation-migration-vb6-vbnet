🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.1 — Classes (`.cls`) → classes `.vb` : champs, propriétés, méthodes

> **Chapitre 12 — Programmation orientée objet** · Section 12.1
> Objectif : traduire **la structure** d'une classe VB6 vers VB.NET — le conteneur, les **champs**, les **propriétés** et les **méthodes**.

---

## 🧭 Périmètre de cette section

On part ici du **squelette** d'une classe : comment un fichier `.cls` devient un fichier `.vb`, et comment se traduisent ses trois familles de membres — **champs**, **propriétés**, **méthodes**.

On laisse **délibérément de côté** ce qui touche à la **durée de vie** des objets :

- la **création** (`Class_Initialize` → constructeur) et la **destruction** (`Class_Terminate`) sont traitées en **12.2** ;
- le nettoyage déterministe (`IDisposable` / `Using`) en **12.3** ;
- les interfaces et l'héritage en **12.5**, la portée fine (`Friend`) en **12.6**, les `Structure` en **12.7**, les événements en **12.8**.

> 💡 Bonne nouvelle : la traduction *structurelle* d'une classe est **la partie la plus mécanique** de tout le chapitre. Le risque, lui, est concentré sur la durée de vie (sections suivantes) et sur quelques **faux-amis** signalés au fil de l'eau.

---

## 1. Du fichier `.cls` au fichier `.vb` : le conteneur

En **VB6**, une classe **est** un fichier `.cls`, et **un seul** fichier `.cls` contient **une seule** classe. Le corps de la classe est **implicite** : on n'écrit jamais `Class … End Class`. En revanche, le fichier commence par un **en-tête caché** (visible si on ouvre le `.cls` dans un éditeur de texte) :

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
' … membres de la classe …
```

En **VB.NET**, tout change de forme :

- la classe est **explicite** : `Public Class CClient … End Class` ;
- **plusieurs** classes peuvent cohabiter dans un même fichier `.vb` ;
- le **nom du fichier est libre** (il n'a plus à coïncider avec le nom de la classe) ;
- l'en-tête d'attributs disparaît ; ses informations sont reportées dans des **mots-clés** ou, au besoin, des **attributs .NET**.

Correspondance des principaux attributs d'en-tête :

| Attribut VB6 | Rôle en VB6 | Traduction en VB.NET |
|--------------|-------------|----------------------|
| `VB_Name` | nom de la classe | nom après `Class` |
| `VB_Creatable` | instanciable depuis l'extérieur (COM) | géré par les **modificateurs d'accès** (et attributs COM si exposition COM) |
| `VB_Exposed` | visible hors du projet (COM) | idem |
| `VB_GlobalNameSpace = True` | membres accessibles **globalement** | ⚠️ pas d'équivalent direct → **refactoring** (importer, qualifier, ou rendre `Shared`) |
| `VB_PredeclaredId = True` | **instance globale** automatique portant le nom de la classe | ⚠️ pas d'équivalent pour une classe ordinaire → repenser en membre **`Shared`** ou en *singleton* |

> ⚠️ **Deux pièges d'en-tête.** `VB_PredeclaredId = True` (instance globale gratuite, comme pour les formulaires) et `VB_GlobalNameSpace = True` (membres visibles partout) reposent sur des mécanismes **absents** de .NET. Le code qui en dépend **compilera après refactoring seulement** — repérez-les **avant** de migrer la classe.

Enfin, l'instruction `Option Explicit` de VB6 devient `Option Explicit On` en VB.NET. **Profitez-en pour ajouter `Option Strict On`** dès que possible (le typage strict est votre meilleur garde-fou ; voir **6.4** pour le démarrage et **17.1** pour l'activation progressive).

```vb
' VB.NET — en-tête de fichier recommandé
Option Explicit On
Option Strict On

Public Class CClient
    ' … champs, propriétés, méthodes …
End Class
```

---

## 2. Les champs (variables membres)

La déclaration des champs est **presque identique** d'un langage à l'autre.

```vb
' VB6
Private mId As Long
Private mNom As String
Public  Etiquette As String      ' champ public (à éviter, voir plus bas)
```

```vb
' VB.NET
Private _id As Integer           ' Long (VB6) = 32 bits → Integer (.NET) !
Private _nom As String
Public  Etiquette As String
```

Trois points méritent l'attention :

**a) Les types changent parfois de taille.** Un `Long` VB6 (32 bits) devient un `Integer` .NET ; un `Integer` VB6 (16 bits) devient un `Short` ; `Currency` devient `Decimal`. C'est l'un des pièges majeurs de la migration — il est traité en détail en **7.2** (entiers redimensionnés). Gardez-le en tête dès l'écriture des champs.

**b) Les champs peuvent enfin être initialisés à la déclaration.** ⭐ En VB6, toute initialisation devait passer par `Class_Initialize`. En VB.NET, on écrit directement :

```vb
' VB.NET — initialiseurs de champ (impossible en VB6)
Private _compteur As Integer = 0
Private _lignes   As New List(Of String)()
```

> ⚠️ **Attention au faux-ami `As New`.** En VB6, `Private mColl As New Collection` est une variable **auto-instanciée** : elle se **recrée toute seule** dès qu'on y accède après l'avoir mise à `Nothing`. En VB.NET, `Private _coll As New List(Of String)()` crée l'objet **une seule fois** (à la construction) et ne le ressuscite **jamais**. Même syntaxe, **comportement différent** — détaillé en **7.5**.

**c) Préférez `Private` + propriété aux champs `Public`.** Un champ `Public` reste valide en .NET, mais l'usage idiomatique est de **l'encapsuler** dans une propriété (voir section suivante). La conversion est quasi gratuite avec les propriétés auto-implémentées ; ce n'est pas obligatoire pour migrer, mais c'est un gain de propreté immédiat. Par convention, le préfixe hongrois `m` des champs VB6 cède souvent la place à un simple `_` (`mNom` → `_nom`).

---

## 3. Les propriétés

C'est ici que se concentrent les différences de **syntaxe** les plus visibles.

### 3.1 Le schéma VB6 : `Get` / `Let` / `Set`

En VB6, une propriété s'écrit avec des blocs **séparés**, adossés à un champ privé :

```vb
' VB6
Private mNom As String

Public Property Get Nom() As String
    Nom = mNom
End Property

Public Property Let Nom(ByVal Value As String)
    mNom = Value
End Property
```

VB6 distingue deux accesseurs d'écriture :

- **`Property Let`** pour affecter une **valeur** (types simples) ;
- **`Property Set`** pour affecter un **objet** (avec le mot-clé `Set` à l'appel).

### 3.2 Le schéma VB.NET : une propriété **unifiée**

VB.NET réunit lecture et écriture dans un **bloc unique**, et **fusionne `Let` et `Set`** en un seul accesseur `Set` :

```vb
' VB.NET
Private _nom As String

Public Property Nom() As String
    Get
        Return _nom
    End Get
    Set(value As String)
        _nom = value
    End Set
End Property
```

> ⚠️ **Faux-ami `Set`.** Le mot `Set` existe dans les deux langages mais ne désigne **plus la même chose** : en VB6, `Property Let` (valeur) **et** `Property Set` (objet) sont **fusionnés** en un unique accesseur `Set` côté .NET — et, par ailleurs, le `Set` d'**affectation** d'objet (`Set obj = …`) a, lui, **purement disparu**. Le détail de ce double faux-ami est en **10.6**.

### 3.3 Lecture seule, écriture seule

En VB6, on exprimait une propriété **en lecture seule** en ne fournissant **que** `Property Get` (pas de `Let`/`Set`), et **en écriture seule** en ne fournissant que `Let`/`Set`. VB.NET rend cela **explicite** avec des modificateurs :

| VB6 | VB.NET |
|-----|--------|
| `Get` seul | `ReadOnly Property` (bloc `Get` uniquement) |
| `Let` / `Set` seul | `WriteOnly Property` (bloc `Set` uniquement) |
| `Get` + `Let` (valeur) | `Property` (lecture-écriture) |
| `Get` + `Set` (objet) | `Property` (lecture-écriture) — `Let`/`Set` fusionnés |

```vb
' VB.NET — propriété en lecture seule
Public ReadOnly Property Credit() As Decimal
    Get
        Return _credit
    End Get
End Property
```

### 3.4 Propriétés **auto-implémentées** (la version concise) ⭐

Pour une propriété triviale (un simple champ exposé), VB.NET génère le champ de stockage **à votre place**. Plus besoin d'écrire le champ ni les accesseurs :

```vb
' VB.NET — propriétés auto-implémentées
Public Property Id  As Integer
Public Property Nom As String
```

Le compilateur crée un **champ privé caché** nommé `_Id`, `_Nom`, etc. (préfixe `_` + nom de la propriété). Ce champ reste **accessible et modifiable depuis l'intérieur** de la classe — utile pour une propriété en lecture seule que des méthodes internes doivent mettre à jour :

```vb
' VB.NET — propriété auto-implémentée en lecture seule
Public ReadOnly Property Credit As Decimal   ' champ généré : _Credit

Public Sub Crediter(montant As Decimal)
    _Credit += montant                        ' on écrit dans le champ généré
End Sub
```

On peut aussi leur donner une **valeur initiale** : `Public Property Actif As Boolean = True`.

### 3.5 Propriétés paramétrées et propriété par défaut

VB6 autorisait des **propriétés paramétrées** (`Property Get Item(ByVal i As Integer)`) et une **propriété par défaut** (via l'attribut caché `VB_UserMemId = 0`, ce qui permettait `objet(i)` sans nommer la propriété). En VB.NET :

- les propriétés paramétrées existent toujours ;
- la propriété par défaut exige le mot-clé **`Default`** *et* qu'elle prenne **au moins un paramètre** (les propriétés par défaut **sans** paramètre, fréquentes en VB6, n'ont **pas** d'équivalent direct).

C'est un point sensible traité en **2.5** (propriétés par défaut) ; à ce stade, retenez seulement que toute propriété par défaut VB6 **sans argument** demandera un **appel explicite** après migration.

---

## 4. Les méthodes

Les méthodes sont les membres qui se traduisent **le plus naturellement** : `Sub` et `Function` existent à l'identique.

```vb
' VB6
Public Sub Crediter(ByVal montant As Currency)
    mCredit = mCredit + montant
End Sub

Public Function EstSolvable() As Boolean
    EstSolvable = (mCredit >= 0)
End Function
```

```vb
' VB.NET
Public Sub Crediter(montant As Decimal)
    _credit += montant                 ' '+=' (affectation composée) est une nouveauté .NET
End Sub

Public Function EstSolvable() As Boolean
    Return _credit >= 0D               ' 'Return' idiomatique
End Function
```

Quatre différences à connaître — toutes approfondies au **chapitre 10** :

- **Valeur de retour.** `NomFonction = valeur` (VB6) reste **toléré**, mais l'écriture idiomatique est `Return valeur` (**10.5**).
- **⚠️ `ByRef` → `ByVal` par défaut.** En VB6, un paramètre **sans mot-clé** est passé **`ByRef`** ; en VB.NET, il est passé **`ByVal`**. Une méthode qui « renvoyait » des valeurs en modifiant ses paramètres **cessera silencieusement de le faire**. C'est **le** piège d'appel n°1 — auditez chaque signature (**10.2**).
- **Surcharge.** VB.NET autorise plusieurs méthodes de même nom à signatures différentes (`Overloads`) — une **nouveauté** sans équivalent VB6 (**10.5**).
- **Affectation composée.** Les opérateurs `+=`, `-=`, `&=`… n'existent pas en VB6 ; ils sont disponibles en VB.NET (gain de lisibilité, comportement identique à la forme longue).

Les **modificateurs d'accès** des méthodes (`Public` / `Private` / `Friend`) se transposent presque tels quels — avec une nuance sur `Friend` (portée **assembly** en .NET) détaillée en **12.6**.

---

## 5. Exemple complet : une classe migrée

### 5.1 La classe d'origine (VB6)

```vb
' Fichier : CClient.cls  (en-tête VB_Name = "CClient" omis ici)
Option Explicit

Private mId     As Long
Private mNom    As String
Private mCredit As Currency

Public Property Get Id() As Long
    Id = mId
End Property
Public Property Let Id(ByVal Value As Long)
    mId = Value
End Property

Public Property Get Nom() As String
    Nom = mNom
End Property
Public Property Let Nom(ByVal Value As String)
    mNom = Value
End Property

Public Property Get Credit() As Currency   ' lecture seule (pas de Let)
    Credit = mCredit
End Property

Public Sub Crediter(ByVal montant As Currency)
    mCredit = mCredit + montant
End Sub

Public Function EstSolvable() As Boolean
    EstSolvable = (mCredit >= 0)
End Function
```

### 5.2 Traduction **fidèle à la structure** (VB.NET explicite)

Version proche de ce que produit l'assistant de mise à niveau : chaque membre est transposé un pour un. Notez les **types redimensionnés** (`Long` → `Integer`, `Currency` → `Decimal`).

```vb
' Fichier : CClient.vb  (nom de fichier libre)
Option Explicit On
Option Strict On

Public Class CClient

    Private _id     As Integer
    Private _nom    As String
    Private _credit As Decimal

    Public Property Id() As Integer
        Get
            Return _id
        End Get
        Set(value As Integer)
            _id = value
        End Set
    End Property

    Public Property Nom() As String
        Get
            Return _nom
        End Get
        Set(value As String)
            _nom = value
        End Set
    End Property

    Public ReadOnly Property Credit() As Decimal
        Get
            Return _credit
        End Get
    End Property

    Public Sub Crediter(montant As Decimal)
        _credit += montant
    End Sub

    Public Function EstSolvable() As Boolean
        Return _credit >= 0D
    End Function

End Class
```

### 5.3 Traduction **idiomatique** (propriétés auto-implémentées)

Une fois la non-régression assurée, on peut condenser les propriétés triviales. Le code se réduit nettement, à comportement **identique** :

```vb
' Fichier : CClient.vb
Option Explicit On
Option Strict On

Public Class CClient

    Public Property Id  As Integer
    Public Property Nom As String

    Public ReadOnly Property Credit As Decimal   ' champ généré : _Credit

    Public Sub Crediter(montant As Decimal)
        _Credit += montant
    End Sub

    Public Function EstSolvable() As Boolean
        Return Credit >= 0D
    End Function

End Class
```

> 💡 Les versions **5.2** et **5.3** sont **équivalentes**. Pendant la migration, privilégiez la **5.2** (traçable, calquée sur l'original) ; gardez la **5.3** pour la phase de **refactoring idiomatique** (**17.5**), une fois les tests de non-régression au vert.

---

## 6. Ce que cette section ne traite pas (encore)

Pour éviter toute confusion, voici ce qui **n'apparaît volontairement pas** dans la classe ci-dessus, et où le trouver :

- **Construction / destruction** : `Class_Initialize` → `Sub New`, et le **problème** de `Class_Terminate` → **12.2**.
- **Nettoyage déterministe** des ressources : `IDisposable` / `Using` → **12.3**.
- **Interfaces et héritage** (`Implements`, `Inherits`) → **12.5**.
- **Portée `Friend`** (= assembly) → **12.6**.
- **`Type` → `Structure`** (et sa sémantique de valeur) → **12.7**.
- **Événements** (`Event` / `RaiseEvent` / `WithEvents`) → **12.8**.

---

## ✅ Points clés

- Une classe VB6 = **un** fichier `.cls` à en-tête caché ; une classe VB.NET = un bloc **`Class … End Class`** explicite, **plusieurs par fichier**, **nom de fichier libre**.
- Surveillez `VB_PredeclaredId` et `VB_GlobalNameSpace` : ils **n'ont pas d'équivalent** et imposent un **refactoring**.
- Les **champs** se traduisent presque à l'identique, mais : **types redimensionnés** (7.2), **initialiseurs** désormais permis, et **`As New`** au comportement différent (7.5).
- Les **propriétés** se **fusionnent** (`Let` + `Set` → `Set`), gagnent des modificateurs **`ReadOnly`/`WriteOnly`** explicites, et peuvent devenir **auto-implémentées**.
- Les **méthodes** changent peu… sauf le piège **`ByRef` → `ByVal` par défaut** (10.2), à auditer systématiquement.

---

## 🔗 Renvois

- **7.2** — Entiers redimensionnés (`Long`→`Integer`, `Currency`→`Decimal`) et **7.5** — `As New`.
- **2.5** / **10.6** — Propriétés par défaut, et le double faux-ami `Set`.
- **10.2** — `ByRef` → `ByVal`, et **10.5** — surcharge et valeur de retour.
- **6.4** / **17.1** — Activation d'`Option Strict`. · **17.5** — Refactoring idiomatique.
- **Suite directe → [12.2 — `Class_Initialize` → constructeur ; `Class_Terminate` → le problème de la finalisation](02-initialize-terminate.md)** ⭐ ⚠️

⏭️ [`Class_Initialize` → constructeur ; `Class_Terminate` → **le problème de la finalisation**](/12-poo/02-initialize-terminate.md)
