🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.3 — Finalisation déterministe perdue → `IDisposable` / `Using` / `Finalize` (le patron Dispose) ⭐

> **Chapitre 12 — Programmation orientée objet** · Section 12.3
> La **réponse** au problème posé en [12.2](02-initialize-terminate.md) : rétablir un nettoyage **déterministe**, mais **explicite**.

---

## 🧭 D'où l'on part

La section 12.2 a établi le diagnostic : en .NET, la finalisation est **non déterministe**, et traduire `Class_Terminate` en `Finalize` est **le piège n°1**. La cure tient en une phrase : **rendre la libération déterministe et explicite**, à la demande — et ne garder `Finalize` que comme **dernier recours**.

> 📌 Rappel du tri de la 12.2 : un `Class_Terminate` qui ne relâchait que des objets se **supprime**. Cette section ne concerne que l'autre cas — celui où la classe détient une **ressource rare** (fichier, connexion, *handle*, socket, objet COM…).

---

## 1. Les trois briques, et leur rôle

| Brique | Rôle | Qui la déclenche |
|--------|------|------------------|
| **`IDisposable`** | le **contrat** : « cet objet sait se nettoyer » (méthode `Dispose`) | — |
| **`Using`** | l'**invocation déterministe** : appelle `Dispose` à coup sûr | **vous** (l'appelant) |
| **`Finalize`** | le **filet de sécurité** : nettoyage de dernier recours | le **GC**, si `Dispose` n'a pas été appelé |

> 🔑 Modèle mental : **`Dispose` = le nettoyage déterministe que *vous* déclenchez ; `Finalize` = le filet que le *GC* déclenche** (tardivement, et seulement si nécessaire). On vise toujours le premier ; le second n'est là que pour le cas où l'on aurait oublié.

---

## 2. Le bloc `Using` : le nettoyage déterministe au quotidien

C'est la **brique que vous écrirez le plus souvent**. `Using` garantit l'appel à `Dispose` à la **sortie du bloc** — y compris si une **exception** est levée, ou si l'on fait `Return` au milieu.

```vb
' VB.NET
Using j As New CJournal()
    j.Ecrire("Bonjour")
End Using            ' ← j.Dispose() est appelé ICI, quoi qu'il arrive
```

C'est exactement ce qui **remplace**, côté appelant, le `Set obj = Nothing` de VB6 :

```
VB6  :  Set j = Nothing        → Class_Terminate (nettoyage) ici
.NET :  End Using              → Dispose (nettoyage) ici
```

`Using` n'est qu'un **raccourci** pour un `Try…Finally` :

```vb
' Équivalent de Using j As New CJournal() … End Using
Dim j As New CJournal()
Try
    j.Ecrire("Bonjour")
Finally
    If j IsNot Nothing Then j.Dispose()
End Try
```

Variantes utiles :

```vb
' Plusieurs ressources dans un même bloc (libérées en ordre inverse)
Using cnx As New SqlConnection(chaine), cmd As New SqlCommand(sql, cnx)
    cnx.Open()
    ' …
End Using

' Sur une ressource déjà obtenue
Using lecteur = cmd.ExecuteReader()
    ' …
End Using
```

> 💡 `Using` gère gracieusement une ressource à `Nothing` (il vérifie avant d'appeler `Dispose`). Et la plupart des classes .NET « à ressource » (`StreamWriter`, `SqlConnection`, `Font`, `Bitmap`…) implémentent **déjà** `IDisposable` : vous pouvez les envelopper dans un `Using` **immédiatement**, sans rien écrire de plus.

---

## 3. Implémenter `IDisposable` — le **cas simple** (le plus fréquent en migration)

Votre classe **enveloppe** une ou plusieurs ressources **managées** déjà `IDisposable` (un `StreamWriter`, une `SqlConnection`…). C'est le cas courant, car en .NET les ressources sont quasiment toujours déjà encapsulées dans de telles classes.

Règle d'or de ce cas : **vous n'avez PAS besoin de finaliseur.** Les objets managés que vous détenez ont **déjà** le leur comme filet de sécurité. Vous vous contentez d'implémenter `IDisposable` et d'y libérer vos membres.

Reprenons le `CJournal` de la 12.2 — voici sa version **correcte** (à comparer à la version naïve à base de `Finalize`) :

```vb
' VB.NET — la BONNE version (cas simple : ressource managée enveloppée)
Imports System.IO

Public Class CJournal
    Implements IDisposable

    Private _writer As StreamWriter

    Public Sub New()
        _writer = New StreamWriter("journal.log", append:=True)
    End Sub

    Public Sub Ecrire(texte As String)
        _writer.WriteLine(texte)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        _writer?.Dispose()       ' libère le writer managé (ferme le fichier)
        _writer = Nothing        ' rend Dispose idempotent (appelable plusieurs fois)
    End Sub

End Class
```

```vb
' Appel — le fichier est fermé de façon déterministe
Using j As New CJournal()
    j.Ecrire("Bonjour")
End Using
```

> ✅ `Dispose` doit être **idempotent** : pouvoir être appelé plusieurs fois sans erreur. Ici, l'opérateur `?.` court-circuite si `_writer` est déjà `Nothing` — un second appel ne fait rien.

---

## 4. Le **patron Dispose complet** — quand on détient une ressource **non managée**

Si votre classe détient **directement** une ressource non managée — un *handle* Win32 brut obtenu par `P/Invoke`, un `IntPtr`, de la mémoire non managée — alors le GC **ne sait pas** la libérer tout seul. Il faut un **finaliseur** comme filet de sécurité, **en plus** de `Dispose`. C'est le **patron canonique** :

```vb
' VB.NET — patron Dispose COMPLET (ressource non managée + ressources managées)
Public Class RessourceBrute
    Implements IDisposable

    Private _disposed As Boolean = False
    Private _handle As IntPtr              ' ressource NON managée
    Private _managee As StreamWriter       ' ressource managée (exemple)

    Public Sub New()
        ' _handle = OuvrirHandle()  (P/Invoke) ; _managee = New StreamWriter(...)
    End Sub

    ' (1) Méthode publique — appelée par l'utilisateur / Using
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)            ' déjà nettoyé : inutile de finaliser
    End Sub

    ' (2) Filet de sécurité — appelé par le GC si Dispose a été oublié
    Protected Overrides Sub Finalize()
        Try
            Dispose(False)
        Finally
            MyBase.Finalize()
        End Try
    End Sub

    ' (3) Le nettoyage réel — un SEUL endroit
    Protected Overridable Sub Dispose(disposing As Boolean)
        If _disposed Then Return

        If disposing Then
            ' Ressources MANAGÉES : à libérer SEULEMENT si on vient de Dispose
            _managee?.Dispose()
        End If

        ' Ressources NON managées : à libérer DANS TOUS LES CAS
        If _handle <> IntPtr.Zero Then
            ' FermerHandle(_handle)   (P/Invoke)
            _handle = IntPtr.Zero
        End If

        _disposed = True
    End Sub

End Class
```

Pourquoi cette mécanique :

- **`Dispose()` publique** appelle `Dispose(True)` (nettoyage complet) puis **`GC.SuppressFinalize(Me)`** : on signale au GC que le finaliseur **n'a plus besoin de s'exécuter** — on évite ainsi le coûteux cycle de finalisation (rappel : un objet finalisable survit à **deux** GC).
- **`Finalize`** appelle `Dispose(False)`.
- **`Dispose(disposing)`** centralise tout :
  - `disposing = True` (venu de `Dispose`) → on libère **managé ET non managé** ;
  - `disposing = False` (venu du **finaliseur**) → on libère **uniquement le non managé**. **Pourquoi ?** Parce qu'à l'heure du finaliseur, les objets managés que l'on référence peuvent **déjà avoir été finalisés** (aucun ordre garanti) : il serait dangereux d'y toucher.
- **`_disposed`** garantit l'**idempotence**.

> 💡 Quand vous tapez `Implements IDisposable`, **Visual Studio propose de générer** ce squelette complet. Pratique — mais ne le gardez que si vous détenez réellement du **non managé** ; sinon, restez sur le cas simple de la section 3.

---

## 5. La recommandation moderne : **`SafeHandle`** (éviter d'écrire un finaliseur)

Écrire un finaliseur à la main est **délicat** et **rarement nécessaire**. La pratique recommandée consiste à **envelopper** tout *handle* non managé dans un **`SafeHandle`** (depuis `Microsoft.Win32.SafeHandles`). Un `SafeHandle` possède **son propre finaliseur** (à finalisation critique) : votre classe le traite alors comme un **membre managé `IDisposable`**, et **retombe sur le cas simple** de la section 3.

```vb
' VB.NET — avec SafeHandle : plus besoin d'écrire Finalize() soi-même
Imports Microsoft.Win32.SafeHandles

Public Class FichierBasNiveau
    Implements IDisposable

    Private _handle As SafeFileHandle      ' encapsule le handle non managé

    Public Sub Dispose() Implements IDisposable.Dispose
        _handle?.Dispose()                 ' le SafeHandle gère sa finalisation
    End Sub

End Class
```

> ✅ Hiérarchie de décision : **(1)** ne détenir que des membres managés `IDisposable` (y compris un `SafeHandle` pour tout *handle*) → **patron simple, sans finaliseur** ; **(2)** seulement si l'on détient un non managé **non encapsulable** → **patron complet** avec finaliseur. En pratique, le cas (1) couvre la grande majorité des migrations.

---

## 6. Propager la libération : la notion de **possession** (*ownership*)

Toutes les ressources ne vivent pas le temps d'un `Using`. Si un objet **possède** un membre `IDisposable` qui **vit aussi longtemps que lui**, alors **cet objet doit lui aussi être `IDisposable`** et libérer ce membre dans son propre `Dispose`. La libération **se propage en cascade** le long de la chaîne de possession.

```vb
' VB.NET — CRapport possède un CJournal pour toute sa durée de vie
Public Class CRapport
    Implements IDisposable

    Private _journal As CJournal

    Public Sub New()
        _journal = New CJournal()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        _journal?.Dispose()       ' on propage la libération au membre possédé
    End Sub

End Class
```

```vb
' Appel — Dispose se propage : r.Dispose() → _journal.Dispose()
Using r As New CRapport()
    ' …
End Using
```

On a donc **deux formes d'usage** :

- **ressource de courte durée** (locale à une méthode) → **`Using`** ;
- **ressource qui vit avec l'objet** (un champ) → **l'objet propriétaire implémente `IDisposable`** et libère ce champ dans son `Dispose`.

---

## 7. Quand **ne pas** implémenter `IDisposable`

Si votre classe ne contient que des **données managées ordinaires** (chaînes, nombres, listes, autres objets **non** `IDisposable`, aucune ressource non managée), elle **n'a pas besoin** d'`IDisposable`. Le **GC** récupère la mémoire ; ajouter `Dispose` « par principe » ne fait **qu'alourdir** l'API sans bénéfice. C'est le prolongement direct du **Cas 1** du tri de la 12.2.

---

## 8. Erreurs fréquentes (à relire avant de valider)

- ❌ **Oublier `Using`** (ou l'appel à `Dispose`) → la ressource fuit : on retombe **exactement** sur le problème de la 12.2.
- ❌ **Mettre le nettoyage uniquement dans `Finalize`** → non déterministe : le piège n°1, à nouveau.
- ❌ **Écrire un finaliseur dont on n'a pas besoin** → pénalité de performance (deux cycles de GC) ; ne le faites **que** pour du non managé non encapsulé.
- ❌ **Toucher un objet managé dans la branche `disposing = False`** → il peut être **déjà finalisé**.
- ❌ **Oublier `GC.SuppressFinalize(Me)`** dans le patron complet → finalisation inutile et coûteuse.
- ❌ **Rendre `Dispose` non idempotent** → erreurs au double appel (utilisez le drapeau `_disposed`).
- ❌ **Lever une exception depuis `Dispose`** → à éviter (peut masquer l'exception d'origine, notamment dans un `Finally`).

---

## 🔁 Du `Class_Terminate` VB6 au bon patron .NET

| Situation VB6 (`Class_Terminate`) | Cible .NET |
|-----------------------------------|------------|
| Ne relâche que des objets (`Set x = Nothing`) | **Rien** — supprimer (GC) |
| Enveloppe une ressource **managée** (fichier, connexion…) | **`IDisposable`** simple, **sans** finaliseur (§3) + `Using` côté appelant |
| Détient une ressource **non managée** encapsulable | Envelopper dans un **`SafeHandle`** → patron simple (§5) |
| Détient une ressource **non managée** brute | **Patron Dispose complet** avec `Finalize` (§4) |
| Ressource qui **vit avec l'objet** | L'objet **propriétaire** implémente `IDisposable` et propage (§6) |

---

## ✅ Points clés

- La libération déterministe en .NET est **explicite** : **`IDisposable` + `Using`**. `Finalize` n'est qu'un **filet de sécurité**, souvent **inutile**.
- **`Using`** garantit `Dispose` à la sortie du bloc (même sur exception) ; c'est le **remplaçant** du `Set obj = Nothing → Class_Terminate` de VB6.
- **Cas simple** (ressources managées enveloppées) : implémenter `Dispose`, libérer les membres, **pas de finaliseur**. C'est le cas **le plus fréquent**.
- **Patron complet** (non managé brut) : `Dispose()` → `Dispose(True)` + `GC.SuppressFinalize` ; `Finalize` → `Dispose(False)` ; nettoyage centralisé dans `Dispose(disposing)`, drapeau `_disposed`.
- **`SafeHandle`** évite d'écrire un finaliseur dans la quasi-totalité des cas.
- La possession se **propage** : un objet qui détient un `IDisposable` durable est lui-même `IDisposable`.

---

## 🔗 Renvois

- **[12.2](02-initialize-terminate.md)** — le **problème** que cette section résout (à lire d'abord).
- **2.2** — *La finalisation déterministe perdue* (cadrage). · **Annexe B.3** — symptôme/cause/correction.
- **16** — Interop COM : libérer un objet COM (`Marshal.ReleaseComObject`) dans un `Dispose`.
- **17.6** — Revue de performance et de mémoire (`IDisposable`, `Using`, fuites de *handles*).
- **Suite → [12.4 — Modules standards (`.bas`) → `Module` ; variables globales et leur portée](04-modules-globales.md)**

⏭️ [Modules standards (`.bas`) → `Module` ; variables globales et leur portée](/12-poo/04-modules-globales.md)
