🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.4 — L'objet `Err` vs `Exception` ; `Err.Raise` → `Throw`

> **Deux façons de décrire une erreur.** VB6 la résume par un **numéro** posé sur un objet **global unique** (`Err`), qu'il faut lire **immédiatement**. VB.NET la décrit par un **objet typé** (`Exception`) qui **se propage** et reste **stable**. On discrimine désormais par **type**, plus par numéro — et `Err.Raise` devient `Throw` d'un objet **bien choisi**.

📍 *Module 11 · § 11.4 · [↑ Introduction du chapitre](README.md) · [← § 11.3](03-on-error-resume-next.md)*

---

## ⚡ En bref

- **`Err`** (VB6) : objet **global** et **unique**, écrasé à chaque erreur. On le lit par `Err.Number` (un **code**), tout de suite, avant qu'il ne change.
- **`Exception`** (VB.NET) : un **objet à part entière**, dont la **classe** porte le sens. Il est passé au `Catch`, **propagé** dans la pile, et **stable**.
- **Discriminer** : par **type** d'exception (et filtres `When`), plus par `Err.Number`.
- **`Err.Raise N`** → **`Throw`** d'un **objet du bon type** (de préférence un type cadre précis, ou une **exception personnalisée** — qui remplace `vbObjectError + N`).

---

## 1. Deux façons de décrire une erreur

| Aspect | `Err` (VB6) | `Exception` (VB.NET) |
|---|---|---|
| Nature | objet **global** et **unique** | **objet** (instance) par erreur |
| Identité de l'erreur | un **numéro** (`Err.Number`) | un **type** (la classe de l'exception) |
| Durée de vie | écrasé par l'erreur suivante → **lire tout de suite** | **stable**, transmis au `Catch`, **propagé** |
| Pile d'appels | **aucune** | **`StackTrace`** (où l'erreur s'est produite) |
| Cause sous-jacente | aucune | **`InnerException`** (chaînage) |
| Discrimination | `Select Case Err.Number` | plusieurs **`Catch` typés** + `When` |

C'est le **basculement central** : on passe d'un modèle **par numéro et global** à un modèle **par objet typé et local**.

---

## 2. L'objet `Err` (VB6)

`Err` porte les informations de la **dernière** erreur :

- `Err.Number` (code), `Err.Description` (texte), `Err.Source` (origine) ;
- `Err.HelpFile` / `Err.HelpContext` (aide), `Err.LastDllError` (code Win32 d'un appel de DLL) ;
- méthodes : `Err.Raise …` (lever) et `Err.Clear` (réinitialiser).

```vb
' VB6 — lire Err IMMÉDIATEMENT après l'erreur
On Error Resume Next
Ouvrir fichier
If Err.Number <> 0 Then
    MsgBox "Erreur " & Err.Number & " : " & Err.Description
    Err.Clear
End If
```

Les **numéros standards** sont connus (6 dépassement, 9 indice hors limites, 11 division par zéro, 13 incompatibilité de type, 53 fichier introuvable, 70 permission refusée, 91 variable objet non définie…). Les **erreurs personnalisées** s'écrivent `Err.Raise vbObjectError + N`, la constante `vbObjectError` évitant les collisions avec les codes système.

> ⚠️ **Fragilité du global** : `Err` étant unique, une opération ultérieure peut le **modifier ou l'effacer**. Il faut **capturer `Number`/`Description` sur-le-champ**. Cette fragilité **disparaît** avec les exceptions (§ 3).

---

## 3. L'objet `Exception` (VB.NET)

En VB.NET, une erreur est un **objet typé**. C'est d'abord la **classe** qui renseigne sur sa nature (`FileNotFoundException`, `OverflowException`, `InvalidCastException`, `ArgumentNullException`…). L'objet expose ensuite des **propriétés** :

| Propriété | Rôle |
|---|---|
| `Message` | description textuelle (≈ `Err.Description`) |
| `Source` | application/objet à l'origine (≈ `Err.Source`) |
| `StackTrace` | **pile d'appels** — **nouveau**, sans équivalent VB6 |
| `InnerException` | **cause enveloppée** — **nouveau** |
| `HResult` | code entier brut (le plus proche de `Err.Number`, rarement utilisé directement) |
| `Data` | dictionnaire de contexte clé/valeur — **nouveau** |
| `HelpLink` | lien d'aide (≈ `HelpFile`/`HelpContext`) |
| `GetType()` | **type** réel de l'exception (sert à discriminer) |

```vb
' VB.NET — un objet riche et stable, transmis au Catch
Try
    Ouvrir(fichier)
Catch ex As Exception
    Journaliser("Type    : " & ex.GetType().Name)
    Journaliser("Message : " & ex.Message)
    Journaliser("Pile    : " & ex.StackTrace)        ' aucun équivalent VB6
    If ex.InnerException IsNot Nothing Then
        Journaliser("Cause   : " & ex.InnerException.Message)
    End If
End Try
```

> 💡 La **pile d'appels** (`StackTrace`) et la **cause** (`InnerException`) sont des **gains majeurs** pour le diagnostic : VB6 n'offrait rien d'équivalent.

---

## 4. Tableau de correspondance `Err` → `Exception`

| `Err` (VB6) | `Exception` (VB.NET) | Note |
|---|---|---|
| `Err.Number` | le **type** de l'exception (principal) ; `ex.HResult` pour le code brut | Discriminer par **type** |
| `Err.Description` | `ex.Message` | |
| `Err.Source` | `ex.Source` | |
| `Err.HelpFile` / `Err.HelpContext` | `ex.HelpLink` | |
| `Err.LastDllError` | `Marshal.GetLastWin32Error()` | contexte P/Invoke → **[§ 16.4](../16-interop-com-api/04-declare-vers-pinvoke.md)** |
| `Err.Clear` | *(inutile)* | l'exception n'est pas un état global |
| — | `ex.StackTrace` | **nouveau** |
| — | `ex.InnerException` | **nouveau** |
| — | `ex.Data` | **nouveau** |

---

## 5. `Err.Raise` → `Throw`

VB6 **levait un numéro** (avec description optionnelle) ; VB.NET **lève un objet**, dont on choisit le **type** le plus parlant.

```vb
' VB6
Err.Raise Number:=vbObjectError + 1000, _
          Source:="Banque.Compte", _
          Description:="Solde insuffisant"
```

```vb
' VB.NET — lever un OBJET, du bon type
Throw New InvalidOperationException("Solde insuffisant")
```

**Choisir le bon type** est essentiel :

- utilisez un **type cadre précis** quand il correspond : `ArgumentException`, `ArgumentNullException`, `ArgumentOutOfRangeException`, `InvalidOperationException`, `NotSupportedException`, `FormatException`… ;
- **ne levez pas `Exception` nu** (trop générique) ni `NullReferenceException`/`IndexOutOfRangeException` (réservées au moteur) — préférez `ArgumentNullException`/`ArgumentOutOfRangeException` ;
- pour une erreur **métier** (ce que VB6 faisait avec `vbObjectError + N`), définissez une **exception personnalisée** (§ 6).

**Correspondance des erreurs standard courantes** (utile quand on *relève* l'équivalent d'une erreur VB6) :

| Erreur VB6 | N° | Type .NET courant |
|---|---|---|
| Overflow | 6 | `OverflowException` |
| Subscript out of range | 9 | `IndexOutOfRangeException` / `ArgumentOutOfRangeException` |
| Division by zero | 11 | `DivideByZeroException` |
| Type mismatch | 13 | `InvalidCastException` / `FormatException` |
| File not found | 53 | `FileNotFoundException` |
| Permission denied | 70 | `UnauthorizedAccessException` |
| Object variable not set | 91 | `NullReferenceException` *(à capturer ; à ne pas lever)* |
| Invalid procedure call/argument | 5 | `ArgumentException` / `ArgumentOutOfRangeException` |

> 🔎 `Err.Raise` **existe encore** en VB.NET (compatibilité), mais le code idiomatique utilise **`Throw`** d'un objet typé.

---

## 6. Les erreurs personnalisées : `vbObjectError` → classe d'exception

Là où VB6 distinguait ses erreurs métier par un **numéro** (`vbObjectError + N`), VB.NET définit une **classe** dédiée — discriminée ensuite par **type**, bien plus clairement.

```vb
' VB.NET — exception personnalisée (remplace vbObjectError + N)
Public Class SoldeInsuffisantException
    Inherits Exception

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(message As String)
        MyBase.New(message)
    End Sub

    Public Sub New(message As String, inner As Exception)
        MyBase.New(message, inner)
    End Sub
End Class
```

```vb
' Lever puis attraper PAR TYPE — sans tester de numéro
Throw New SoldeInsuffisantException("Solde insuffisant")
' ...
Catch ex As SoldeInsuffisantException
    ' traitement spécifique
End Try
```

Bonnes pratiques :

- fournir les **trois constructeurs standards** (sans paramètre, `message`, `message + inner`) ;
- **hériter directement de `Exception`** — **pas** d'`ApplicationException`, dont l'usage est **déconseillé** (recommandation obsolète) ;
- définir la classe est un sujet de **POO** : voir le **[§ 12.5 (héritage)](../12-poo/05-interfaces-heritage.md)**.

> 💡 Besoin d'un **code** en plus du type (codes d'un domaine, d'un protocole…) ? Ajoutez une **propriété** à votre exception (ex. `CodeMetier`), ou rangez le détail dans `ex.Data`. On garde alors le **type** pour le sens général et la **propriété** pour la précision.

---

## 7. Relancer et enrichir (rappel)

Pour relancer depuis un `Catch`, les règles du **[§ 11.2](02-vers-try-catch.md)** s'appliquent :

- **`Throw` nu** relance en **préservant la pile** d'origine ;
- **`Throw ex` réinitialise la pile** (à éviter) ;
- **envelopper** ajoute du contexte sans perdre la cause :

```vb
Catch ex As SqlException
    Throw New DataException("Échec de la mise à jour du compte", ex)   ' ex → InnerException
End Try
```

---

## 📋 Tableau de synthèse

| Aspect | VB6 (`Err`) | VB.NET (`Exception`) | Nature |
|---|---|---|---|
| Modèle | numéro, objet **global** | **objet typé**, par instance | Structurel |
| Identité | `Err.Number` | **type** de l'exception | Discriminer par type |
| Description | `Err.Description` | `ex.Message` | Mécanique |
| Origine | `Err.Source` | `ex.Source` | Mécanique |
| Pile / cause | — | `StackTrace` / `InnerException` | ✅ Gain majeur |
| Lever | `Err.Raise N` | `Throw New Type(message)` | Choisir le **bon type** |
| Erreur perso | `vbObjectError + N` | **classe** héritant d'`Exception` | Type, pas numéro |
| Durée de vie | écrasé → lire vite | **stable**, propagé | ✅ Fin de la fragilité |

---

## ✅ À retenir

1. **Numéro → type** : VB6 identifie l'erreur par `Err.Number` ; VB.NET par la **classe** de l'exception. Discriminez par **type** (et `When`), plus par numéro.
2. **`Err` est global et fragile** (à lire immédiatement) ; **`Exception` est un objet stable**, transmis au `Catch` et propagé — avec, en prime, **`StackTrace`** et **`InnerException`**.
3. **`Err.Raise` → `Throw`** d'un objet **bien choisi** : un type cadre précis quand il existe ; **jamais `Exception` nu**, ni les types réservés au moteur.
4. **Erreurs métier** : remplacez `vbObjectError + N` par une **exception personnalisée** (héritant d'`Exception`, **pas** d'`ApplicationException`), avec ses trois constructeurs standards.
5. **Relancez proprement** : `Throw` nu pour la pile, enveloppement (`Throw New …(…, ex)`) pour le contexte.

---

*Section suivante : **[§ 11.5 — Stratégie de migration progressive (cohabitation `On Error` / `Try` temporaire)](05-migration-progressive.md)***

**Cible** : .NET Framework 4.7.2 · **Indicateurs** : ⭐ point clé · ⚠️ piège · 🛠️ outils · 🔗 interop

⏭️ [Stratégie de migration **progressive** (cohabitation `On Error` / `Try` temporaire)](/11-gestion-erreurs/05-migration-progressive.md)
