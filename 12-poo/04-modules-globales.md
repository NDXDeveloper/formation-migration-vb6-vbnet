🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.4 — Modules standards (`.bas`) → `Module` ; variables globales et leur portée

> **Chapitre 12 — Programmation orientée objet** · Section 12.4
> Le pendant **procédural et global** de la 12.1 : le conteneur `.bas` devient un `Module`, et l'on clarifie la **portée** des variables globales.

---

## 🧭 Périmètre

La section 12.1 traitait les **modules de classe** (`.cls`). Ici, on s'occupe des **modules standards** (`.bas`) — ceux qui hébergent le code **procédural** et l'**état global** d'une application VB6. Bonne nouvelle : le **conteneur** se traduit très directement. Le vrai sujet, c'est la **portée** — savoir ce qui reste accessible « de partout », et ce qui change subtilement.

---

## 1. Du `.bas` au `Module` : un conteneur **statique**

En **VB6**, un fichier `.bas` est un **module standard** : un conteneur **sans instance**, toujours disponible. Ses membres `Public` sont **globaux** — accessibles **partout dans le projet, sans qualification**. Ses membres `Private` sont limités au module.

En **VB.NET**, l'équivalent direct est le bloc **`Module … End Module`** :

- ses membres sont **implicitement `Shared`** (on n'écrit jamais `Shared`) ;
- il est **non instanciable** (`New Module1` est une erreur) — cohérent avec le `.bas` ;
- ses membres `Public` sont, comme en VB6, accessibles **sans qualification** dans tout le projet.

```vb
' VB6 — Module1.bas
Public g_Compteur As Long
Public Const MAX As Integer = 100

Public Sub Initialiser()
    g_Compteur = 0
End Sub

Public Function Doubler(ByVal n As Long) As Long
    Doubler = n * 2
End Function
```

```vb
' VB.NET — Module1.vb
Module Module1

    Public g_Compteur As Integer        ' Long (VB6) → Integer (.NET) — voir 7.2
    Public Const MAX As Integer = 100

    Public Sub Initialiser()
        g_Compteur = 0
    End Sub

    Public Function Doubler(n As Integer) As Integer
        Return n * 2
    End Function

End Module
```

Un `Module` peut contenir **tout ce que contenait un `.bas`** : procédures, variables, constantes, **énumérations**, types utilisateur et déclarations d'API. Leur contenu suit toutefois ses **propres règles** de migration, traitées ailleurs :

- `Type…End Type` → **`Structure`** (voir **12.7**) ;
- `Declare` (API Win32) → **`P/Invoke`** (voir **16.4**) ;
- énumérations et constantes (voir **7.6**).

---

## 2. `Module` ou classe à membres `Shared` ? (le bon choix pour migrer)

VB.NET offre une alternative : regrouper des fonctions dans une **classe** à membres `Shared`. La différence est **l'accès** :

| Aspect | `Module` | `Class` + membres `Shared` |
|--------|----------|----------------------------|
| Accès aux membres publics | **sans qualification** (projet) | **qualifié** : `Classe.Membre` |
| Instanciation | impossible | possible (mais inutile pour du `Shared`) |
| Idéal pour… | **migrer un `.bas` fidèlement** | regrouper des *helpers* en .NET idiomatique |

> ✅ Pour une migration **fidèle**, choisissez le **`Module`** : il **préserve l'accès non qualifié** dont dépend le code VB6 (`Doubler(x)` plutôt que `Module1.Doubler(x)`). Le passage à une classe à membres `Shared`, ou à un *singleton*, relève d'un **refactoring** ultérieur (**17.5**) — pas de la migration.

---

## 3. Les portées, de VB6 à VB.NET

Voici la correspondance des niveaux de portée. C'est le **cœur** de cette section.

| Portée | Déclaration VB6 | Équivalent VB.NET |
|--------|-----------------|-------------------|
| Locale (procédure) | `Dim` dans `Sub`/`Function` | `Dim` dans la méthode |
| Locale **persistante** | `Static x` | `Static x` (**préservé**) |
| Module, **privée** | `Private` (ou `Dim` au niveau module) | `Private` |
| Module standard, **publique = globale** | `Public` (ou l'ancien `Global`) dans un `.bas` | **`Public` dans un `Module`** (non qualifié, projet) |
| Assembly | *(pas d'équivalent direct)* | **`Friend`** (voir **12.6**) |

Deux précisions :

- **`Global` (VB6) = `Public`.** L'ancien mot-clé `Global` des modules standards se traduit simplement par `Public`.
- **`Static` est conservé**, avec la même sémantique : dans un `Module` (ou une méthode `Shared`), il existe **une seule copie** globale ; dans une **méthode d'instance**, il y a **une copie par objet** — exactement comme en VB6.

> ⚠️ **Attention au `Public` dans une *classe*.** Dans un module **standard** (`.bas`), `Public` au niveau module signifie **global**. Dans un module de **classe** (`.cls`), `Public` au niveau module désigne un **membre d'instance**, **pas** une variable globale. Ne confondez pas les deux en migrant : seul le `Public` d'un `Module` produit une variable réellement globale.

---

## 4. Variables globales : ce qui change vraiment

### 4.1 L'accès non qualifié **fonctionne toujours** ✅

C'est le point rassurant : les membres `Public` d'un `Module` restent appelables **sans préfixe** dans tout le projet. Le code VB6 qui lit `g_Compteur` ou appelle `Doubler(x)` continue de le faire **tel quel**.

### 4.2 Les **collisions de noms** ⚠️

Si **deux modules** définissent un membre public de **même nom**, l'appel non qualifié devient **ambigu** : VB.NET exige alors la **qualification** (`Module1.Doubler`). Ce cas surgit souvent quand on **consolide** d'anciens modules.

Surveillez aussi les collisions avec le **framework** ou avec le **runtime VB** : une variable globale nommée `Environment`, ou une fonction maison `Left`, peut entrer en conflit avec `System.Environment` ou avec la fonction `Left` de `Microsoft.VisualBasic`. La solution est de **qualifier** ou de **renommer**.

### 4.3 Les fonctions « globales » de VB6 viennent désormais de `Microsoft.VisualBasic`

Beaucoup de fonctions que VB6 exposait « globalement » (`MsgBox`, `Left`, `Mid`, `Format`, `IsNumeric`…) ne sont pas des mots-clés du langage : elles proviennent du **runtime VB**. En VB.NET, elles sont fournies par l'espace de noms **`Microsoft.VisualBasic`**, **importé par défaut** dans les projets VB. Elles restent donc disponibles **sans qualification** — mais leur comportement mérite parfois un second regard (voir **9.2** pour les fonctions de chaînes).

### 4.4 Initialisation et **ordre** ⚠️

Les membres d'un `Module` peuvent recevoir une **valeur initiale** (`Public g_Compteur As Integer = 0`). Mais l'initialisation d'un `Module` est portée par son **constructeur statique**, déclenché **avant le premier accès** au module — et il n'existe **aucun ordre garanti** d'initialisation **entre** modules. Ne faites donc **pas** dépendre l'initialisation d'un module de celle d'un autre. Pour une séquence d'amorçage **déterministe**, centralisez-la dans **`Sub Main`** (ou une initialisation paresseuse explicite).

### 4.5 `Option Explicit` / `Option Strict` : un garde-fou pour les globales

En VB6 **sans** `Option Explicit`, une variable mal orthographiée crée silencieusement une **nouvelle variable** (un `Variant`) — source classique de faux « globaux » fantômes. En VB.NET, `Option Explicit On` (par défaut) **interdit** cela, et `Option Strict On` renforce le typage. Activez-les (voir **5.1** et **6.4**) : ils éliminent toute une catégorie de bugs sur l'état global.

### 4.6 Le point d'entrée `Sub Main`

En VB6, `Sub Main` réside dans un `.bas`. En VB.NET, il peut rester dans un **`Module`**, désigné comme **objet de démarrage** du projet. La transposition est directe.

---

## 5. Un mot sur l'état global (sans le sur-jouer)

Les programmes VB6 s'appuient souvent **beaucoup** sur des variables globales. Elles se migrent **directement** en membres `Public` de `Module` — c'est le bon réflexe **pendant** la migration. Gardez toutefois deux choses en tête :

- l'**état global mutable** est un **candidat naturel au refactoring** — mais **plus tard** (**17.5**), une fois la non-régression assurée. Migration ≠ redécoupage de l'architecture ;
- si l'application devient **multithread**, l'état global partagé demande une **attention particulière** (synchronisation) : un point absent de l'univers VB6 monothread, à ne pas négliger lors d'évolutions futures.

---

## 🔁 Récapitulatif VB6 → VB.NET

| VB6 | VB.NET |
|-----|--------|
| Fichier `.bas` (module standard) | bloc `Module … End Module` |
| Membre `Public` / `Global` d'un `.bas` | `Public` dans un `Module` (global, non qualifié) |
| Membre `Private` d'un module | `Private` |
| `Static x` | `Static x` (préservé) |
| `Sub Main` dans un `.bas` | `Sub Main` dans un `Module` (objet de démarrage) |
| Fonctions runtime « globales » (`MsgBox`, `Left`…) | espace `Microsoft.VisualBasic`, **importé par défaut** |
| `Type` / `Declare` dans un module | `Structure` (12.7) / `P/Invoke` (16.4) |

---

## ✅ Points clés

- Un `.bas` devient un **`Module`** : conteneur **statique, non instanciable**, dont les membres `Public` restent **accessibles sans qualification** dans tout le projet.
- Pour migrer **fidèlement**, préférez `Module` à une classe `Shared` : il **préserve l'accès non qualifié**.
- Le `Public` d'un **`Module`** est **global** ; le `Public` d'une **classe** est un **membre d'instance** — ne pas confondre.
- Surveillez les **collisions de noms** (entre modules, ou avec le framework / le runtime VB) : elles imposent **qualification** ou **renommage**.
- **Aucun ordre garanti** d'initialisation entre modules → centralisez l'amorçage dans **`Sub Main`**.
- L'**état global** se migre tel quel, mais reste un **objet de refactoring** ultérieur (17.5), pas de la phase de migration.

---

## 🔗 Renvois

- **12.1** — modules de **classe** (`.cls`). · **12.6** — `Friend` = portée **assembly**.
- **7.2** — types redimensionnés (`Long`→`Integer`). · **7.6** — constantes et énumérations.
- **9.2** — fonctions de chaînes (`Microsoft.VisualBasic` vs `String`).
- **12.7** — `Type` → `Structure`. · **16.4** — `Declare` → `P/Invoke`.
- **5.1** / **6.4** — `Option Explicit` / `Option Strict`. · **17.5** — refactoring idiomatique.
- **Suite → [12.5 — Interfaces : `Implements` (différences) et héritage (nouveauté .NET)](05-interfaces-heritage.md)**

⏭️ [Interfaces : `Implements` (différences) et **héritage** (nouveauté .NET)](/12-poo/05-interfaces-heritage.md)
