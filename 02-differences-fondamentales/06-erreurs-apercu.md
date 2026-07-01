🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2.6 Gestion d'erreurs : `On Error` vs exceptions structurées

> **Chapitre 2 — Différences fondamentales VB6 ↔ VB.NET**
> Un changement de **paradigme** : on passe d'un modèle de **codes d'erreur et de sauts** à un modèle d'**objets exception**. Vue d'ensemble ici ; traitement détaillé au **chapitre 11**.

---

## Un aperçu, pas le mode d'emploi

Cette section est un **panorama**. Son rôle est de poser les **deux modèles** face à face et de faire  
comprendre **en quoi** ils diffèrent dans leur philosophie — pas de dérouler la mécanique de  
conversion. Tout le « comment migrer » (équivalents de `Resume`, conversion vers `Try`/`Catch`, le  
cas épineux de `On Error Resume Next`, l'objet `Err`, la stratégie progressive) est traité **en  
profondeur au chapitre 11**, vers lequel cette section renvoie systématiquement.

Pourquoi en parler dès le chapitre 2 ? Parce que ce changement **découle directement du changement  
de runtime** (2.1) : les **exceptions structurées** sont un service du **CLR**. Comprendre le modèle  
d'exceptions, c'est aussi mieux comprendre des mécanismes déjà rencontrés — à commencer par le bloc
`Using` de la section 2.2, qui n'est, on le verra, **que du sucre syntaxique au-dessus d'un
`Try…Finally`**.

> ⚠️ **La grille du chapitre, appliquée ici.** Ce sujet est plutôt rassurant côté détection :
> contrairement à `Set` ou aux entiers, **`On Error` n'a *pas* disparu de VB.NET** (point clé
> ci-dessous). Le code à base de `On Error` **compile** encore. Le risque silencieux est donc
> **concentré** sur un seul cas — `On Error Resume Next` — délibérément renvoyé à la section
> **11.3**, qui lui est entièrement consacrée.

---

## 🔵 Côté VB6 : `On Error`, un modèle de codes et de sauts

En VB6, la gestion d'erreurs repose sur l'instruction **`On Error`** et sur un objet global,
**`Err`**. Le principe : quand une erreur survient, on **saute** vers un endroit désigné, ou on  
**l'ignore**. Trois formes principales :

```vb
' Forme 1 — On Error GoTo : sauter vers un gestionnaire étiqueté
Sub Traiter()
    On Error GoTo Gestionnaire
    OuvrirFichier "data.txt"          ' si ça échoue → saut vers Gestionnaire
    Exit Sub
Gestionnaire:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description
End Sub
```

```vb
' Forme 2 — On Error Resume Next : IGNORER l'erreur et continuer (le cas dangereux)
Sub Traiter()
    On Error Resume Next
    OuvrirFichier "data.txt"          ' si ça échoue… on passe à la ligne suivante
    Traiter Données                   ' …sans savoir que la précédente a échoué
End Sub
```

```vb
' Forme 3 — On Error GoTo 0 : désactiver la gestion d'erreurs en cours
On Error GoTo 0
```

À quoi s'ajoutent l'instruction **`Resume`** (reprendre l'exécution **à la ligne fautive**),
**`Resume Next`** (reprendre **à la ligne suivante**) ou **`Resume étiquette`**, et l'objet **`Err`**
porteur de l'information (`Err.Number`, `Err.Description`, `Err.Source`), que l'on déclenche soi-même  
avec **`Err.Raise`**.

Trois traits **caractérisent** ce modèle — et préparent le contraste :

- **L'erreur est un *numéro*.** On discrimine en testant `Err.Number` (souvent dans un gros `Select Case`).
- **L'information est *globale* et *éphémère*.** Un seul objet `Err`, écrasé à chaque nouvelle erreur.
- **Le contrôle se fait par *sauts*.** `GoTo`, `Resume`, `Resume Next` : le flux se déroute, puis éventuellement reprend.

---

## 🟢 Côté VB.NET : les exceptions structurées

En VB.NET, le modèle change de nature : une erreur n'est plus un **numéro**, c'est un **objet  
exception** — une instance d'une classe dérivant de **`System.Exception`**. On l'**émet** avec
`Throw`, on l'**intercepte** avec `Try`/`Catch`, et on garantit le nettoyage avec `Finally` :

```vb
' VB.NET — Try / Catch / Finally
Sub Traiter()
    Try
        OuvrirFichier("data.txt")
    Catch ex As IOException                 ' on capture par TYPE d'erreur
        Console.WriteLine($"Erreur d'E/S : {ex.Message}")
    Catch ex As Exception                   ' filet plus général
        Console.WriteLine($"Erreur : {ex.Message}")
    Finally
        ' nettoyage GARANTI — exécuté qu'il y ait eu erreur ou non
    End Try
End Sub
```

```vb
' Déclencher une erreur : Err.Raise  →  Throw
Throw New ArgumentException("Argument invalide")
```

Les mêmes trois traits, **renversés** :

- **L'erreur est un *type*.** On discrimine en **capturant par classe** (`Catch ex As IOException`), pas en testant un numéro. La hiérarchie d'exceptions permet d'attraper précisément ou largement.
- **L'information est *riche* et *portée par l'objet*.** L'exception transporte un message, le **type**, la **pile d'appels** (`StackTrace`), et même une **exception interne** (`InnerException`) — bien plus qu'un `Err.Number`.
- **La propagation est *automatique*.** Une exception non interceptée **remonte la pile d'appels** toute seule, jusqu'au premier `Catch` capable de la traiter. *(VB6 propageait aussi les erreurs non gérées vers le gestionnaire appelant — mais sans la richesse ni le typage du modèle .NET.)*

VB.NET ajoute encore les **filtres d'exception** (`Catch … When …`), qui permettent de conditionner  
une capture — détaillés en **section 11.2**.

> 💡 **Le lien avec 2.2.** Le `Finally` garantit qu'un bloc s'exécute **quoi qu'il arrive**. C'est
> exactement le mécanisme sur lequel repose **`Using`** : `Using obj … End Using` équivaut à un
> `Try … Finally obj.Dispose() End Try`. La parade au piège n°1 (libération déterministe) est donc
> **bâtie sur le modèle d'exceptions** présenté ici.

---

## 🆚 Les deux modèles, côte à côte

| Aspect | **VB6 — `On Error`** | **VB.NET — exceptions** |
|---|---|---|
| **Nature de l'erreur** | Un **numéro** (`Err.Number`) | Un **objet typé** (classe d'exception) |
| **Discrimination** | `If/Select Case` sur `Err.Number` | `Catch ex As TypeException` (par type) |
| **Information portée** | Objet `Err` **global**, écrasé | Objet **riche** : message, type, pile, `InnerException` |
| **Déclenchement** | `Err.Raise` | `Throw` |
| **Interception** | `On Error GoTo étiquette` | `Try … Catch` |
| **Nettoyage garanti** | À gérer **manuellement** | **`Finally`** (intégré) |
| **Reprise après erreur** | **`Resume` / `Resume Next`** | **Pas d'équivalent direct** (→ 11.1) |
| **Contrôle du flux** | Par **sauts** (`GoTo`) | Par **propagation** structurée |

---

## 🔑 Le point clé pour la migration : `On Error` **survit** en VB.NET

C'est la particularité qui change toute la stratégie. **VB.NET a conservé `On Error GoTo` et
`Resume`** (par compatibilité — c'est l'un des rares langages .NET à le faire). Conséquence directe :

> ⭐ **On n'est pas obligé de tout convertir d'un coup.** Le code VB6 à base de `On Error` **compile
> tel quel** en VB.NET. On peut donc le **laisser en place** dans un premier
> temps et migrer la gestion d'erreurs **progressivement**, module par module, vers `Try`/`Catch`.
> C'est précisément la **stratégie de cohabitation** détaillée en **section 11.5**.

Concrètement, les deux styles peuvent **coexister dans un même module** — chacun **confiné à sa  
procédure** — le temps de migrer au rythme voulu :

```vb
' Même module, en cours de migration : les deux styles cohabitent

' ── Procédure héritée : encore en On Error (migrée « telle quelle »)
Sub ChargerAncien()
    On Error GoTo Gestion
    ' ... code hérité, pas encore restructuré ...
    Exit Sub
Gestion:
    Journaliser("Erreur " & Err.Number & " : " & Err.Description)
End Sub

' ── Procédure neuve (ou déjà refondue) : en Try/Catch
Sub ChargerNouveau()
    Try
        ' ... code neuf ...
    Catch ex As Exception
        Journaliser(ex.Message)
    End Try
End Sub
```

Cette tolérance vient toutefois avec **deux limites** à connaître :

- **On ne peut pas mélanger les deux modèles dans une *même* procédure.** `On Error` **et**
  `Try`/`Catch` dans la même méthode → **erreur de compilation** (donc **visible**). La cohabitation
  se fait **entre** procédures, pas **à l'intérieur** d'une procédure.
- **`Resume` n'a pas d'équivalent structuré.** Une fois dans un `Catch`, on **ne peut pas** revenir
  exécuter la ligne fautive comme le faisait `Resume`. Tout code qui s'appuyait sur `Resume` /
  `Resume Next` pour **réessayer** ou **poursuivre** doit être **restructuré** (boucles de
  *retry*, conditions…). Mécanique détaillée en **section 11.1**.

---

## ⚠️ Où se cache le danger silencieux ? (renvoi 11.3)

Si ce sujet n'est globalement **pas** un nid à pièges silencieux — le code compile, et `On Error`  
est préservé —, **un cas** concentre tout le risque : **`On Error Resume Next`**.

En ignorant **silencieusement** les erreurs pour « passer à la suite », cette construction pouvait  
déjà **masquer des bugs en VB6**. À la migration, le danger se **double** : non seulement le  
comportement « ignorer et continuer » est délicat à reproduire fidèlement, mais le traduire
**mécaniquement** peut soit **étouffer** des exceptions qui auraient dû remonter, soit **changer** le
point où le programme reprend. C'est un piège à part entière, marqué **⭐ ⚠️**, et c'est pourquoi il  
fait l'objet d'une section dédiée :

> ⚠️ **Le cas `On Error Resume Next` est traité en profondeur en [section 11.3](../11-gestion-erreurs/03-on-error-resume-next.md).**
> N'essayez pas de le convertir « à l'instinct » : c'est l'un des points où une traduction naïve
> introduit des régressions difficiles à diagnostiquer.

---

## 💡 Pourquoi le modèle .NET est (vraiment) meilleur

Au-delà de l'obligation de migrer, le passage aux exceptions est un **gain réel** :

- **On capture ce qui nous concerne.** `Catch ex As IOException` n'attrape que les erreurs d'E/S ; le reste **continue de remonter**. Fini le `Select Case Err.Number` fourre-tout.
- **On en sait beaucoup plus.** Type, message, **pile d'appels**, exception interne : le diagnostic est sans commune mesure avec un simple numéro.
- **Le nettoyage est fiable.** `Finally` (et donc `Using`) garantit la libération des ressources **même en cas d'erreur** — ce qui, on l'a vu en 2.2, n'est pas un détail.
- **On ne peut plus « oublier » une erreur par inadvertance.** Une exception non gérée **interrompt** le programme de façon visible, au lieu de se faire avaler discrètement comme avec `On Error Resume Next`.

> 💡 **L'état d'esprit de migration.** Reproduire fidèlement `On Error` d'abord (la cohabitation),
> puis **réécrire vers `Try`/`Catch`** là où c'est utile — pas l'inverse. On **préserve le
> comportement** avant de **moderniser** (le même principe que partout dans cette formation).

---

## ↔️ Le détail, au chapitre 11 (renvois)

Cette section s'arrête au **panorama**. Tout le traitement opérationnel se trouve ici :

| Sujet | Section |
|---|---|
| `On Error GoTo`, `Resume`, `Resume Next` : équivalents et logique de remplacement | **11.1** |
| Convertir vers `Try`/`Catch`/`Finally` (et les filtres `When`) | **11.2** |
| **`On Error Resume Next`** : pourquoi c'est dangereux à migrer ⭐ ⚠️ | **11.3** |
| L'objet `Err` vs `Exception` ; `Err.Raise` → `Throw` | **11.4** |
| Stratégie de migration **progressive** (cohabitation `On Error` / `Try`) | **11.5** |

---

## ✅ À retenir

- La gestion d'erreurs change de **paradigme** : VB6 raisonne en **codes d'erreur et sauts** (`On Error`, `Err`, `Resume`) ; VB.NET en **objets exception** (`Try`/`Catch`/`Finally`, `Throw`).
- Ce changement **découle du runtime** (2.1) : les exceptions sont un service du **CLR**. Et **`Using` (2.2) est bâti sur `Try…Finally`**.
- **Différences de fond** : l'erreur devient un **type** (capture par classe) au lieu d'un **numéro** ; l'information est **riche** (message, pile, `InnerException`) au lieu d'un `Err` global ; la propagation est **automatique** ; le nettoyage est **garanti** par `Finally`.
- **Point clé** : **`On Error` survit en VB.NET** → le code compile, et la migration peut être **progressive** (cohabitation, → 11.5). Limites : **pas de mélange** des deux modèles dans une même procédure (erreur visible), et **`Resume` sans équivalent** structuré (→ 11.1).
- Le **seul vrai danger silencieux** est **`On Error Resume Next`** → traité à part en **section 11.3** ; ne pas le convertir à l'instinct.
- Le modèle .NET est un **gain réel** (capture ciblée, diagnostic riche, nettoyage fiable, erreurs non « avalées »). On **préserve d'abord**, on **modernise ensuite**.

---

> 🧭 **Section suivante** → [2.7 Tableau de synthèse des incompatibilités majeures](07-synthese.md)  
> 🔝 **Retour** → [Sommaire du chapitre 2](README.md)

⏭️ [Tableau de synthèse des incompatibilités majeures](/02-differences-fondamentales/07-synthese.md)
