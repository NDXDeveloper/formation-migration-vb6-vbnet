🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.2 — `Class_Initialize` → constructeur ; `Class_Terminate` → **le problème de la finalisation** ⭐ ⚠️

> **Chapitre 12 — Programmation orientée objet** · Section 12.2
> **La section la plus importante du chapitre — et l'un des points les plus risqués de toute la formation.**

---

## 🧭 Pourquoi cette section est cruciale

La traduction d'une classe (section 12.1) était mécanique. **La durée de vie des objets, non.** C'est ici que se manifeste le changement de *runtime* annoncé depuis le début : le passage du **compteur de références** de VB6 au **garbage collector** (GC) de .NET.

Cette section se découpe en deux moitiés très inégales :

- **La construction** (`Class_Initialize` → `Sub New`) : facile, et même **enrichie** par .NET.
- **La destruction** (`Class_Terminate` → ?) : **le piège n°1 de la migration**. On y **diagnostique** le problème.

> 📌 Répartition avec la section suivante : **12.2 pose le problème**, **[12.3](03-idisposable-using.md) apporte la solution** (le patron `IDisposable` / `Using` / `Finalize`). Ne mettez pas en place le patron complet ici — comprenez d'abord *pourquoi* il est nécessaire.

---

## Partie A — La construction : `Class_Initialize` → `Sub New`

### A.1 Le mécanisme

En **VB6**, `Class_Initialize` s'exécute **automatiquement** à la création de l'objet. Deux limites importantes :

- il **ne prend aucun paramètre** ;
- il y en a **un seul** par classe.

Conséquence : VB6 ne permet **pas** de construire un objet *avec des arguments*. D'où un idiome très répandu — **« créer puis initialiser »** — où l'on appelle une méthode `Init` juste après la création.

```vb
' VB6
Private Sub Class_Initialize()
    mCredit = 0                 ' valeurs par défaut
End Sub

Public Sub Init(ByVal id As Long, ByVal nom As String)
    mId = id
    mNom = nom
End Sub
```

```vb
' VB6 — appel
Dim c As CClient
Set c = New CClient
c.Init 42, "Durand"             ' initialisation manuelle, en deux temps
```

### A.2 La traduction de base

En **VB.NET**, `Class_Initialize` devient le **constructeur** `Sub New()`. Le corps se transpose tel quel.

```vb
' VB.NET
Public Sub New()
    ' équivalent de Class_Initialize
    ' souvent vide, grâce aux initialiseurs de champ (voir 12.1)
End Sub
```

### A.3 Ce que .NET apporte — et le piège associé

VB.NET autorise des **constructeurs paramétrés** et **surchargés**. L'idiome `Init` disparaît : on **replie** l'initialisation dans le constructeur.

```vb
' VB.NET
Public Sub New()
End Sub

Public Sub New(id As Integer, nom As String)
    _id = id
    _nom = nom
End Sub
```

```vb
' VB.NET — appel : construction + initialisation en une seule fois
Dim c As New CClient(42, "Durand")
```

> ⚠️ **Piège du constructeur par défaut.** Dès que vous déclarez **un** constructeur paramétré, VB.NET **cesse de générer** le constructeur sans paramètre implicite. Si le code VB6 créait l'objet **sans argument** (`New CClient`, toujours permis), il **cessera de compiler**. Conservez donc un `Sub New()` explicite — c'est exactement ce qu'on a fait ci-dessus.

Trois points complémentaires :

- **Ordre d'initialisation** : à la construction, le constructeur de la **classe de base** s'exécute d'abord (appel implicite `MyBase.New()`), puis les **initialiseurs de champ** de la classe, puis le **corps** du constructeur.
- **`As New` au niveau d'un champ** : en VB6, l'auto-instanciation **différait** `Class_Initialize` jusqu'au premier accès ; en VB.NET, l'objet est créé **à la construction**. Décalage de **timing** à connaître (voir **7.5** et **12.1**).
- **Exceptions dans le constructeur** : si `Sub New` lève une exception, **l'objet n'est pas créé** et la référence **n'est jamais affectée**. C'est plus net qu'en VB6, mais différent — pensez-y si `Class_Initialize` s'appuyait sur `On Error`.

---

## Partie B — La destruction : **le cœur du problème**

### B.1 VB6 : une destruction **déterministe** (compteur de références)

Les objets VB6 sont des objets **COM** gérés par **comptage de références**. Dès que la **dernière** référence à un objet disparaît — variable hors de portée, mise à `Nothing`, ou réaffectée — le compteur tombe à **0** et l'objet est détruit **immédiatement**. `Class_Terminate` s'exécute **à cet instant précis**, de façon **prévisible**.

C'est ce qui rend possible le motif « **j'acquiers une ressource à la création, je la libère à la destruction** », avec une **garantie de moment** :

```vb
' VB6 — CJournal : ouvre un fichier, le referme à la destruction
Private mNumFichier As Integer

Private Sub Class_Initialize()
    mNumFichier = FreeFile
    Open "journal.log" For Append As #mNumFichier
End Sub

Private Sub Class_Terminate()
    Close #mNumFichier          ' libéré DÈS que l'objet n'est plus référencé
End Sub

Public Sub Ecrire(ByVal texte As String)
    Print #mNumFichier, texte
End Sub
```

```vb
' VB6 — appel
Dim j As CJournal
Set j = New CJournal
j.Ecrire "Bonjour"
Set j = Nothing                 ' ← le fichier est fermé ICI, immédiatement
```

> 💡 **Le seul cas où VB6 lui-même n'est pas déterministe** : les **références circulaires**. Si deux objets se référencent mutuellement, leur compteur ne tombe jamais à 0 et `Class_Terminate` **ne se déclenche jamais** — une fuite classique en VB6.

### B.2 .NET : une finalisation **non déterministe** (garbage collector)

.NET n'utilise **pas** de comptage de références, mais un **ramasse-miettes à traçage**. Un objet est récupéré quand le GC s'exécute **et** constate qu'il est **inaccessible** — à un moment **imprévisible** : quelques millisecondes plus tard, plusieurs minutes plus tard, ou seulement à la fin du processus.

Conséquence directe, **contre-intuitive pour qui vient de VB6** :

```vb
' VB.NET
obj = Nothing        ' ne détruit RIEN, n'exécute AUCUN nettoyage.
                     ' Cela retire seulement une référence.
```

Le plus proche équivalent de `Class_Terminate` est le **finaliseur** :

```vb
Protected Overrides Sub Finalize()
    ' …
End Sub
```

Mais le finaliseur cumule les **inconvénients** :

- il s'exécute sur un **thread de finalisation séparé**, à un moment **choisi par le GC** ;
- un objet **finalisable** survit à **au moins deux cycles** de GC (coût mémoire et performance) ;
- **aucun ordre** de finalisation n'est garanti entre objets ;
- il peut **ne pas s'exécuter** dans certains scénarios d'arrêt — on ne peut donc **pas s'y fier** pour une libération critique ;
- mal écrit, il peut **ressusciter** l'objet (le rendre à nouveau accessible) — un bug subtil.

### B.3 Le piège : `Class_Terminate` → `Finalize`

La traduction **naïve** (souvent produite par l'assistant de mise à niveau) consiste à transformer `Class_Terminate` en `Finalize`. **Le code compile. Il s'exécute même.** Mais le nettoyage **change de moment** : il devient indéterministe.

```vb
' VB.NET — traduction NAÏVE et DANGEREUSE
Public Class CJournal
    Private _writer As StreamWriter

    Public Sub New()
        _writer = New StreamWriter("journal.log", append:=True)
    End Sub

    Protected Overrides Sub Finalize()   ' ⚠️ ne se déclenche PAS quand vous croyez
        _writer.Close()
        MyBase.Finalize()
    End Sub

    Public Sub Ecrire(texte As String)
        _writer.WriteLine(texte)
    End Sub
End Class
```

```vb
' VB.NET — appel
Dim j As New CJournal()
j.Ecrire("Bonjour")
j = Nothing          ' ← ne ferme RIEN ; le fichier reste ouvert jusqu'à
                     '    un passage du GC, à un instant indéterminé
```

Le décalage de timing, vu côte à côte :

```
VB6  :  … Set obj = Nothing ──► Class_Terminate s'exécute ICI (immédiat, garanti)
.NET :  … obj = Nothing ─────► (rien) … … … ► GC ► Finalize (plus tard, indéterminé)
```

C'est l'illustration parfaite du fil rouge de la formation : **ça compile, ça tourne, mais ça ne se comporte plus pareil.** Les **symptômes** typiques en production :

- des **fichiers verrouillés** : impossible de les rouvrir ou de les supprimer, car le *handle* traîne ;
- un **pool de connexions épuisé** : les connexions ne sont pas rendues à temps, l'application finit par ne plus pouvoir se connecter ;
- des **fuites de *handles* / objets GDI+** : on atteint les **limites de l'OS** ;
- des **verrous (mutex)** tenus trop longtemps.

> ⚠️ **Ne tentez pas de « forcer » le déterminisme avec `GC.Collect()`.** Appeler le GC à la main pour simuler `Class_Terminate` est un **anti-patron** : c'est coûteux, cela ne donne **aucune** garantie de timing fiable, et cela masque le vrai problème. La bonne réponse est `IDisposable` / `Using`.

### B.4 Que faire du corps de `Class_Terminate` ? (le tri à opérer)

Tout `Class_Terminate` ne pose pas problème. **Avant de migrer**, classez son contenu :

**Cas 1 — Il ne fait que relâcher des références d'objets** (`Set mObj = Nothing`, `Set mAutre = Nothing`…).
→ **Supprimez-le purement et simplement.** En .NET, c'est le **GC** qui récupère la mémoire ; remettre les champs à `Nothing` à la destruction est **inutile**. Une grande partie des `Class_Terminate` VB6 tombent dans ce cas et deviennent des **no-ops**.

**Cas 2 — Il libère une ressource rare ou non managée** : fichier, connexion à une base, *handle* Win32, objet GDI+, socket, mutex, **objet COM** (RCW)…
→ **C'est là qu'il faut un nettoyage déterministe.** Il ne se traduit **pas** par `Finalize`, mais par le patron **`IDisposable` / `Using`** → **section [12.3](03-idisposable-using.md)**. (Le cas particulier de la libération d'objets COM via `Marshal.ReleaseComObject` relève de l'interop, **chapitre 16**.)

> ✅ **Un gain au passage** : contrairement à VB6, le GC à traçage de .NET **récupère correctement les références circulaires**. Le piège de fuite par cycle, classique en VB6 (voir B.1), **disparaît**. .NET perd le *timing* déterministe mais gagne la **collecte des cycles**.

### B.5 La bonne direction (aperçu de la 12.3)

La libération déterministe, en .NET, est **explicite** : on implémente **`IDisposable`** (méthode `Dispose`), et on l'appelle de façon déterministe — idéalement via un bloc **`Using`**, qui **garantit** l'appel à `Dispose` à la sortie du bloc, **même en cas d'exception**.

```vb
' VB.NET — la bonne approche, dans l'esprit (patron complet en 12.3)
Public Class CJournal
    Implements IDisposable
    ' … Dispose() ferme _writer de façon déterministe …
End Class
```

```vb
' VB.NET — appel : le fichier est fermé à la sortie du bloc, c'est garanti
Using j As New CJournal()
    j.Ecrire("Bonjour")
End Using            ' ← Dispose() est appelé ICI, déterministe
```

> 🔑 **Le changement de mentalité.** En VB6, la libération était **automatique** (portée par l'objet). En .NET, la libération déterministe est **à la charge de l'appelant** : c'est lui qui doit écrire `Using` (ou appeler `Dispose`). Le rôle de `Finalize` n'est plus que celui d'un **filet de sécurité** — et, bien souvent, il n'est même pas nécessaire. Tout cela est détaillé en **12.3**.

---

## 🔁 Tableau de correspondance

| VB6 | Traduction **naïve** (à éviter) | **Bonne** cible |
|-----|-------------------------------|-----------------|
| `Class_Initialize` (sans paramètre) | `Sub New()` | `Sub New()`, et constructeurs **paramétrés** au besoin ✓ |
| Idiome `Init(...)` après création | méthode `Init` conservée | replié dans un **constructeur paramétré** |
| `Class_Terminate` — relâche des objets (`Set x = Nothing`) | `Finalize` | **rien** : à **supprimer** (le GC suffit) |
| `Class_Terminate` — libère une ressource rare/non managée | `Finalize` ⚠️ | **`IDisposable` + `Using`** → [12.3](03-idisposable-using.md) |
| (référence circulaire qui fuyait en VB6) | — | **résolue** par le GC à traçage |

---

## ✅ Points clés

- `Class_Initialize` → **`Sub New`** : transposition simple, **enrichie** par les constructeurs **paramétrés/surchargés** (fin de l'idiome `Init`).
- ⚠️ Déclarer un constructeur paramétré **supprime** le constructeur sans paramètre implicite : **gardez un `Sub New()`** si le code créait l'objet sans argument.
- En VB6, `Class_Terminate` est **déterministe** (compteur de références → 0) ; en .NET, la finalisation est **non déterministe** (le GC décide *quand*).
- **`obj = Nothing` ne libère rien** et n'exécute aucun nettoyage en .NET.
- Traduire `Class_Terminate` en **`Finalize`** est **le piège n°1** : ça compile, ça tourne, mais fichiers verrouillés, connexions épuisées et *handles* fuités s'ensuivent.
- **Triez** le corps de `Class_Terminate` : **supprimez-le** s'il ne relâche que des objets ; passez à **`IDisposable`/`Using`** (12.3) s'il libère une ressource rare.

---

## 🔗 Renvois

- **2.2** — *La finalisation déterministe perdue* : le même piège, en version « cadrage ».
- **7.5** — `As New` (instanciation implicite) et son décalage de timing. · **12.1** — structure de classe.
- **16** — Interop COM : libérer un objet COM (`Marshal.ReleaseComObject`).
- **Annexe B.3** — *`Class_Terminate` qui ne se déclenche plus* : symptôme, cause, correction.
- **Suite directe → [12.3 — Finalisation déterministe perdue → `IDisposable` / `Using` / `Finalize` (le patron Dispose)](03-idisposable-using.md)** ⭐

⏭️ [Finalisation déterministe perdue → `IDisposable` / `Using` / `Finalize` (le pattern Dispose)](/12-poo/03-idisposable-using.md)
