🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.5 — Interfaces : `Implements` (différences) et **héritage** (nouveauté .NET)

> **Chapitre 12 — Programmation orientée objet** · Section 12.5
> Deux sujets : faire **évoluer `Implements`** (les interfaces existaient déjà en VB6), et **aborder l'héritage** — une capacité **nouvelle** qu'il faut savoir… **ne pas** précipiter.

---

## 🧭 Périmètre

VB6 connaissait **un seul** mécanisme de polymorphisme : l'**implémentation d'interface** (`Implements`). Il **ignorait totalement** l'héritage d'implémentation. Cette section traite donc :

- **Partie A** — les **interfaces** : `Implements` existe dans les deux langages, mais sa **forme** et sa **souplesse** changent ;
- **Partie B** — l'**héritage** : une **nouveauté .NET** — et la question centrale, *faut-il l'utiliser pendant la migration ?*

---

## Partie A — Les interfaces : `Implements`

### A.1 En VB6 : une **classe** tient lieu d'interface

VB6 n'a **pas** de mot-clé `Interface`. Pour définir un contrat, on crée une **classe** (`.cls`) dont les membres publics (au corps vide) servent de signature. On l'implémente ensuite avec `Implements`, et les membres implémentés sont **`Private`** et **nommés `Interface_Membre`** :

```vb
' VB6 — l'"interface" est une classe : IFigure.cls
Public Function Aire() As Double
End Function
```

```vb
' VB6 — Cercle.cls
Implements IFigure

Private mRayon As Double

Private Function IFigure_Aire() As Double          ' Private + préfixe imposé
    IFigure_Aire = 3.14159 * mRayon * mRayon
End Function
```

```vb
' VB6 — appel : uniquement via une référence d'interface
Dim f As IFigure
Set f = New Cercle
MsgBox f.Aire
```

Détail important : ces membres `IFigure_Aire` étant **`Private`**, ils ne font **pas** partie de l'API publique de `Cercle` — ils ne sont accessibles **que** via une référence `IFigure`.

### A.2 En VB.NET : un mot-clé `Interface` dédié + une **clause `Implements`**

VB.NET sépare clairement le **contrat** (l'interface) de son **implémentation** :

```vb
' VB.NET — une vraie interface
Public Interface IFigure
    Function Aire() As Double
    ReadOnly Property Nom() As String      ' une interface peut déclarer
                                           ' propriétés et événements, pas que des méthodes
End Interface
```

L'implémentation utilise une **clause `Implements`** sur chaque membre — et le nom du membre est **libre** :

```vb
' VB.NET — Cercle
Public Class Cercle
    Implements IFigure

    Private _rayon As Double

    Public Function Aire() As Double Implements IFigure.Aire
        Return Math.PI * _rayon * _rayon
    End Function

    Public ReadOnly Property Nom() As String Implements IFigure.Nom
        Get
            Return "Cercle"
        End Get
    End Property
End Class
```

```vb
' VB.NET — appel
Dim f As IFigure = New Cercle()
Console.WriteLine(f.Aire())
```

Ce que VB.NET apporte par rapport à VB6 :

- le **nom** du membre implémentant est **libre** (plus de préfixe `Interface_` imposé) ;
- un **même membre** peut implémenter **plusieurs** membres d'interfaces : `… Implements IA.Faire, IB.Exécuter` ;
- une classe peut implémenter **plusieurs interfaces** (comme en VB6, via plusieurs `Implements`) ;
- les **interfaces peuvent hériter** d'autres interfaces (`Interface IB : Inherits IA`).

> 💡 Comme en VB6, il faut implémenter **tous** les membres de l'interface (sinon la classe doit être `MustInherit`).

### A.3 La traduction, et un piège de **visibilité**

| VB6 | VB.NET |
|-----|--------|
| « interface » = une **classe** au corps vide | mot-clé **`Interface … End Interface`** |
| `Implements IFigure` | `Implements IFigure` (inchangé) |
| membre `Private IFigure_Aire()` | membre (au nom libre) + clause **`Implements IFigure.Aire`** |
| accès **uniquement via l'interface** | dépend de la **visibilité** du membre (voir ci-dessous) |

> ⚠️ **Changement de surface d'API.** En VB6, les membres implémentés sont `Private` : invisibles sur la classe. En VB.NET, si vous les déclarez **`Public`** (le style courant), ils deviennent accessibles **à la fois** via l'interface **et** directement sur la classe — la surface publique **s'élargit**. Pour **retrouver le comportement VB6** (accessible *uniquement* via l'interface), déclarez le membre **`Private`** (implémentation explicite d'interface) :
>
> ```vb
> Private Function Aire() As Double Implements IFigure.Aire   ' accessible SEULEMENT via IFigure
> ```

L'assistant de mise à niveau gère l'essentiel de cette transposition ; vérifiez surtout les cas où une **même classe** servait à la fois d'**interface** et de type **instanciable** — en .NET, on **sépare** l'`Interface` (contrat) de la `Class` (implémentation).

---

## Partie B — L'héritage : une **nouveauté .NET**

### B.1 VB6 ne connaît **pas** l'héritage d'implémentation

En VB6, on **ne peut pas** hériter du code d'une classe. Pour réutiliser du code, on emploie la **contenance** (délégation) : une classe **possède** une instance d'une autre et **lui délègue** des appels.

```vb
' Réutilisation par CONTENANCE — le motif VB6 (à conserver tel quel en migration)
Public Class Voiture
    Private _moteur As Moteur          ' Voiture "A UN" Moteur (et non "EST UN")
    Public Sub Demarrer()
        _moteur.Demarrer()             ' délégation
    End Sub
End Class
```

### B.2 VB.NET : l'héritage **simple** complet

VB.NET introduit l'**héritage simple** (une seule classe de base) et tout son vocabulaire :

| Mot-clé | Rôle |
|---------|------|
| `Inherits` | déclare la classe de base |
| `MyBase` | accède au membre / constructeur de base (`MyBase.New(...)`) |
| `Me` / `MyClass` | l'instance courante / le membre tel que défini **dans cette classe** |
| `Overridable` | rend un membre **redéfinissable** (⚠️ **pas** par défaut !) |
| `Overrides` | **redéfinit** un membre de base |
| `MustInherit` / `MustOverride` | classe **abstraite** / membre **abstrait** |
| `NotInheritable` / `NotOverridable` | classe **scellée** / membre **non redéfinissable** |
| `Shadows` | **masque** un membre de base par son nom (à manier avec prudence) |

```vb
' VB.NET — syntaxe d'héritage (illustration)
Public Class Animal
    Public Overridable Function Cri() As String    ' opt-in EXPLICITE
        Return "..."
    End Function
End Class

Public Class Chien
    Inherits Animal
    Public Overrides Function Cri() As String
        Return "Ouaf"
    End Function
End Class
```

> ⚠️ **Deux pièges majeurs de l'héritage.**
> - **Rien n'est redéfinissable par défaut.** Contrairement à d'autres langages, une méthode VB.NET n'est **pas** redéfinissable tant qu'on n'écrit pas `Overridable`. Oublier ce mot-clé empêche toute redéfinition.
> - **`Shadows` vs `Overrides`.** Redéclarer un membre **sans** `Overrides` le **masque** (`Shadows`) au lieu de le redéfinir : un appel via une référence de la classe de base utilisera alors la version **de base** — un **changement de comportement silencieux**.
>
> Enfin, le **chaînage des constructeurs** s'applique : le constructeur de base s'exécute en premier ; si la base n'a **pas** de constructeur sans paramètre, l'appel `MyBase.New(args)` est **obligatoire** (voir **12.2**).

### B.3 Que faire en migration ? → **ne pas introduire d'héritage**

C'est le message essentiel de cette partie. Puisque VB6 n'avait pas d'héritage, **il n'y a rien à migrer** côté `Inherits`. La tentation est de **redessiner** les hiérarchies pour « profiter » de l'héritage — **résistez-y**.

> ⚠️ **Migration ≠ réécriture.** Conservez la **contenance** VB6 **telle quelle** (une `Voiture` qui *possède* un `Moteur` reste ainsi). Introduire de l'héritage **modifie l'architecture** et **multiplie les risques** de régression. La modélisation idiomatique (héritage, classes abstraites, interfaces enrichies) est un **refactoring à part entière**, à mener **après** la migration et la non-régression (**17.5**).

### B.4 Là où l'héritage **apparaît malgré tout**

Même sans en introduire dans votre code métier, vous **rencontrerez** l'héritage car le **framework** s'en sert :

- **Windows Forms** : chaque formulaire migré **`Inherits System.Windows.Forms.Form`**. C'est le premier endroit où le mot `Inherits` surgit concrètement (voir **chapitre 13**).
- **Contrôles personnalisés**, **exceptions personnalisées** (`Inherits Exception`, voir **11.4**), et divers points d'extension du framework reposent sur l'héritage.

Dans ces cas, l'héritage n'est **pas un choix de conception** de votre part : c'est le **modèle imposé par .NET**, et il est légitime de le suivre.

---

## 🧭 Interface ou héritage ? (repère de conception)

Pour mémoire, une fois venu le temps du refactoring :

- une **interface** définit un **contrat** sans partage d'implémentation — idéale pour le polymorphisme ;
- l'**héritage** exprime une relation **« est un »** avec **code partagé** — puissant mais **rigidifiant** ;
- la **contenance/délégation** (« a un ») reste souvent **préférable** à l'héritage.

En migration, la règle est simple : **les interfaces restent des interfaces, la contenance reste de la contenance.** On ne convertit pas l'un en l'autre.

---

## 🔁 Récapitulatif VB6 → VB.NET

| VB6 | VB.NET |
|-----|--------|
| « interface » = classe au corps vide | mot-clé **`Interface`** dédié |
| `Private IFace_Membre()` | membre (nom libre) + **`Implements IFace.Membre`** |
| membre implémenté **invisible** sur la classe | `Public` → visible ; `Private` → **uniquement via l'interface** |
| réutilisation par **contenance** uniquement | contenance **+** héritage (`Inherits`) — héritage **nouveau** |
| *(pas d'héritage)* | `Overridable`/`Overrides`, `MustInherit`, `NotInheritable`, `Shadows` |
| *(rien d'imposé)* | formulaires/contrôles/exceptions **héritent** de classes du framework |

---

## ✅ Points clés

- Les **interfaces** existaient en VB6 (sous forme de classes) ; VB.NET offre un mot-clé **`Interface`** dédié et une **clause `Implements`** plus souple (nom de membre libre, implémentations multiples).
- ⚠️ La **visibilité** des membres implémentés change la **surface d'API** : `Public` les expose sur la classe ; `Private` reproduit le comportement VB6 (accès **via l'interface** seulement).
- L'**héritage** est **nouveau** en .NET — donc **rien à migrer**, et **rien à introduire** pendant la migration (**17.5** pour plus tard).
- ⚠️ En VB.NET, **rien n'est redéfinissable par défaut** (`Overridable` explicite), et `Shadows` (masquage) ≠ `Overrides` (redéfinition).
- L'héritage **surgit quand même** via le framework : **Windows Forms**, contrôles, exceptions.

---

## 🔗 Renvois

- **12.1** — structure de classe. · **12.2** — chaînage des constructeurs (`MyBase.New`).
- **11.4** — exceptions personnalisées (`Inherits Exception`).
- **13** — formulaires Windows Forms (`Inherits …Form`).
- **17.5** — refactoring idiomatique (où l'on s'autorise enfin l'héritage).
- **Suite → [12.6 — `Public` / `Private` / `Friend` : `Friend` = portée assembly en .NET](06-portee.md)** ⚠️

⏭️ [`Public`/`Private`/`Friend` : `Friend` = portée **assembly** en .NET](/12-poo/06-portee.md)
