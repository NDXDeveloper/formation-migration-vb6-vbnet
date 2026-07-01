🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.6 — L'objet `App`, le registre, et `FileSystemObject` → `System.IO` / espace `My`

**Les services « runtime » de VB6 — informations sur l'application, accès au registre, manipulation de fichiers — n'ont pas besoin d'un pont d'interopérabilité : on les *remplace* par des équivalents .NET natifs, plus simples et plus sûrs.**

> 🔗 **Le contrepoint du chapitre.** Les cinq sections précédentes ont construit des ponts vers le
> monde COM/natif (RCW, CCW, P/Invoke, marshaling). Celle-ci tient le discours inverse : pour toute
> une catégorie de services VB6, le bon geste n'est **pas** de faire le pont, mais de **réécrire** avec
> ce que le framework offre déjà — **sans COM ni handle à gérer**. C'est aussi la meilleure façon de
> *réduire* la surface d'interop (philosophie « pont, pas destination », et objectif du module 17.2).

---

## Le principe : remplacer, pas faire le pont

Trois cas typiques où une migration naïve serait tentée de conserver du COM ou du P/Invoke — à tort :

- L'objet global **`App`** de VB6 → propriétés de l'espace **`My`** et classes du framework.
- L'accès au **registre** (fonctions VB6 ou API Win32) → **`Microsoft.Win32.Registry`** / `My.Computer.Registry` (et souvent : plus de registre du tout).
- Le **`FileSystemObject`** (un composant COM) et les E/S intrinsèques VB6 → **`System.IO`** / `My.Computer.FileSystem`.

> 💡 **L'espace `My`, trait d'union convivial.** VB.NET fournit `My` comme raccourci « à la VB6 » vers
> les fonctionnalités courantes : `My.Application`, `My.Computer.FileSystem`, `My.Computer.Registry`,
> `My.Settings`… Souvent plus proche de l'esprit VB6 que les API .NET brutes, il facilite la
> transition tout en restant 100 % managé. En dessous, c'est `System.IO`, `Environment`, etc. — que
> l'on peut toujours utiliser directement.

---

## L'objet `App` → `My.Application.Info` & compagnie

La quasi-totalité des membres de `App` trouve son équivalent dans **`My.Application.Info`** (qui lit  
les attributs d'assembly définis dans les propriétés du projet) :

| VB6 `App` | Équivalent .NET |
|-----------|-----------------|
| `App.Path` | `Application.StartupPath` (WinForms) / `My.Application.Info.DirectoryPath` |
| `App.EXEName` | `My.Application.Info.AssemblyName` |
| `App.Title` | `My.Application.Info.Title` |
| `App.Major` / `Minor` / `Revision` | `My.Application.Info.Version` (`.Major` / `.Minor` / `.Build` / `.Revision`) |
| `App.ProductName` | `My.Application.Info.ProductName` |
| `App.CompanyName` | `My.Application.Info.CompanyName` |
| `App.FileDescription` | `My.Application.Info.Description` |
| `App.LegalCopyright` | `My.Application.Info.Copyright` |
| `App.PrevInstance` | `Mutex` nommé / application mono-instance (WinForms) |
| `App.LogEvent` | `My.Application.Log` / `System.Diagnostics.EventLog` |
| `App.hInstance` | `Marshal.GetHINSTANCE(...)` (rarement nécessaire) |

> ⚠️ **`App.Path` n'est PAS `Environment.CurrentDirectory`.** `App.Path` renvoyait le **dossier de
> l'exécutable** ; `Environment.CurrentDirectory` renvoie le **répertoire de travail**, qui peut
> **changer** en cours d'exécution (après une boîte de dialogue de fichiers, par exemple). Les
> confondre est un bug silencieux classique : utilisez **`Application.StartupPath`** ou
> **`My.Application.Info.DirectoryPath`**.

> ⚠️ **Numéro de version : 3 parties (VB6) → 4 parties (.NET).** VB6 exposait
> `Major`/`Minor`/`Revision` ; un `System.Version` .NET est `Major.Minor.Build.Revision`. La
> correspondance du troisième nombre **diffère** — vérifiez le mappage si votre code compare des
> versions.

```vbnet
' App.Path → le dossier de l'exécutable (et SURTOUT pas le répertoire de travail)
Dim dossier As String = My.Application.Info.DirectoryPath
```

### `App.PrevInstance` → un `Mutex` nommé

`App.PrevInstance` (déjà peu fiable en VB6) se remplace par un **`Mutex` nommé**, robuste :

```vbnet
Dim creeNouveau As Boolean
Using verrou As New Threading.Mutex(True, "MaSociete.MonApp.Unique", creeNouveau)
    If Not creeNouveau Then
        MessageBox.Show("L'application est déjà en cours d'exécution.")
        Return
    End If
    ' ... démarrage normal ...
End Using
```

> 💡 En Windows Forms, l'option **« application à instance unique »** (`My.Application` mono-instance +
> événement `StartupNextInstance`) offre une alternative intégrée et idiomatique.

---

## Le `FileSystemObject` (et les E/S intrinsèques VB6) → `System.IO` / `My.Computer.FileSystem`

Le `FileSystemObject` est un **composant COM** : on pourrait l'atteindre par RCW (module 16.1)… mais  
ce serait une **erreur**. Il en va de même pour les **instructions de fichier intrinsèques** de VB6
(`Open`, `Print #`, `Get`, `Put`, `Dir`, `Kill`…), qui subsistent dans `Microsoft.VisualBasic` (un
portage littéral compile) mais ne sont **plus idiomatiques**. La cible est **`System.IO`**, doublée du  
raccourci convivial **`My.Computer.FileSystem`**.

| VB6 (FSO **ou** intrinsèque) | `System.IO` | `My.Computer.FileSystem` |
|------------------------------|-------------|--------------------------|
| `FileExists` / `Dir()` | `File.Exists` | `FileExists` |
| `FolderExists` | `Directory.Exists` | `DirectoryExists` |
| `OpenTextFile` (lecture) / `Open For Input` | `File.OpenText` / `File.ReadAllText` | `ReadAllText` |
| `CreateTextFile` / `Open For Output` | `File.CreateText` / `File.WriteAllText` | `WriteAllText` |
| `Open For Append` | `File.AppendText` | `WriteAllText(…, append:=True)` |
| `CopyFile` / `FileCopy` | `File.Copy` | `CopyFile` |
| `MoveFile` / `Name … As` | `File.Move` | `MoveFile` / `RenameFile` |
| `DeleteFile` / `Kill` | `File.Delete` | `DeleteFile` |
| `CreateFolder` / `MkDir` | `Directory.CreateDirectory` | `CreateDirectory` |
| `BuildPath` | `Path.Combine` | `CombinePath` |
| `GetBaseName` | `Path.GetFileNameWithoutExtension` | — |
| `GetExtensionName` | `Path.GetExtension` | — |
| `GetParentFolderName` | `Path.GetDirectoryName` | `GetParentPath` |
| `FileLen` | `New FileInfo(p).Length` | `GetFileInfo(p).Length` |
| `GetAttr` / `SetAttr` | `File.GetAttributes` / `SetAttributes` | — |
| `Dir` (énumération) | `Directory.EnumerateFiles` / `GetFiles` | `GetFiles` |

### Lire un fichier texte — trois idiomes

```vb
' VB6 — FileSystemObject
Dim fso As Object, ts As Object, contenu As String
Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.OpenTextFile("C:\app\notes.txt", 1)   ' 1 = ForReading
contenu = ts.ReadAll()
ts.Close
```

```vbnet
' VB.NET — System.IO (le plus direct)
Dim contenu As String = IO.File.ReadAllText("C:\app\notes.txt")

' ... ou en flux, avec libération déterministe (voir ci-dessous)
Using lecteur As New IO.StreamReader("C:\app\notes.txt")
    Dim contenu2 As String = lecteur.ReadToEnd()
End Using

' ... ou façon conviviale VB
Dim contenu3 As String = My.Computer.FileSystem.ReadAllText("C:\app\notes.txt")
```

### ⚠️ `Open` / `Close` → `Using` (la finalisation déterministe, encore)

En VB6, on faisait `Open … : … : Close`. En .NET, les flux, lecteurs et écrivains
(`FileStream`, `StreamReader`, `StreamWriter`) sont **`IDisposable`** : il faut les **fermer
explicitement**, et le bon outil est **`Using`** (le patron du module 12.3). **Ne comptez pas sur le  
ramasse-miettes** pour fermer un fichier — c'est le piège de la finalisation non déterministe (modules
2.2 / 12.3) appliqué aux E/S : un fichier laissé ouvert peut rester verrouillé bien après que vous en
ayez « fini ».

### Le point dur : fichiers binaires / à accès direct

Les fichiers VB6 ouverts `For Binary` / `For Random` avec `Get`/`Put` sur des `Type` contenant des
**chaînes de longueur fixe** sont le cas le plus délicat à migrer. Côté .NET, on s'appuie sur
`FileStream` + `BinaryReader`/`BinaryWriter`, en **gérant soi-même la disposition des enregistrements**
— ce qui mobilise les notions des modules 7.4 (chaînes de longueur fixe), 12.7 (`Type` → `Structure`)
et 16.5 (marshaling). `FreeFile`/`EOF`/`LOF`/`Seek` se traduisent par la position et la longueur du  
flux (`stream.Position`, `stream.Length`, `reader.EndOfStream`).

---

## Le registre → `Microsoft.Win32.Registry` / `My.Computer.Registry`

VB6 abordait le registre de deux manières.

### Les fonctions intrinsèques VB6 (`GetSetting` / `SaveSetting`…)

`SaveSetting`, `GetSetting`, `GetAllSettings`, `DeleteSetting` **existent toujours** dans
`Microsoft.VisualBasic` : un portage littéral compile.

> ⚠️ **Mais leur usage est déconseillé**, et — surtout — **ne présumez pas que l'emplacement de
> stockage est strictement identique** à celui de VB6 (`HKEY_CURRENT_USER\Software\VB and VBA Program
> Settings\…`). Si un autre outil lit ces clés, **vérifiez** le comportement avant de vous y fier.

### L'accès direct au registre (API ou besoin réel) → la classe `Registry`

Les déclarations d'API VB6 (`RegOpenKeyEx`, `RegQueryValueEx`…) se remplacent par la classe managée
**`Microsoft.Win32.Registry`**. Atout important : **`RegistryKey` est `IDisposable`** → **`Using`**
(toujours le réflexe du module 12.3) :

```vbnet
Imports Microsoft.Win32

Using cle As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\MaSociete\MonApp")
    If cle IsNot Nothing Then
        Dim serveur As String = CStr(cle.GetValue("Serveur", "localhost"))
    End If
End Using
```

> ⚠️ **Vues 32/64 bits (WOW6432Node).** Un processus 32 bits voit `HKLM\Software\WOW6432Node`. Pour
> cibler une vue précise indépendamment de la « bitness » du processus, utilisez
> `RegistryKey.OpenBaseKey(hive, RegistryView.Registry32)` (ou `Registry64`). C'est la déclinaison,
> côté registre, du thème de la « bitness » récurrent dans ce chapitre.

### La vraie question : faut-il *encore* utiliser le registre ?

Souvent, **non**. Pour des **préférences applicatives**, le successeur naturel de
`GetSetting`/`SaveSetting` est **`My.Settings`** (concepteur *Settings*), qui persiste dans un fichier
`user.config` **sans toucher au registre** :

```vbnet
' Préférences utilisateur, sans registre
Dim serveur As String = My.Settings.Serveur
My.Settings.Serveur = "db.example.com"
My.Settings.Save()
```

> 🎯 Réservez l'accès direct au registre aux **vrais** besoins système (interopérer avec une clé
> imposée par Windows ou un tiers). Pour la configuration de votre application, `My.Settings` ou
> `app.config` (`ConfigurationManager`) sont plus propres, plus portables, et déployables sans droits
> particuliers.

---

## ⚠️ Pièges à connaître

- **`App.Path` ≠ `Environment.CurrentDirectory`** : utilisez `Application.StartupPath` /
  `My.Application.Info.DirectoryPath`, jamais le répertoire de travail. ⚠️
- **Version 3 parties → 4 parties** : la correspondance du troisième nombre change (VB6 `Revision` vs
  .NET `Build`/`Revision`). ⚠️
- **Fichiers et flux sont `IDisposable`** : remplacez `Open … Close` par **`Using`** ; ne laissez pas
  le GC fermer vos fichiers (finalisation non déterministe, module 12.3). ⚠️
- **Fichiers binaires/à accès direct** : la migration des enregistrements à champs de longueur fixe est
  le point dur (modules 7.4, 12.7, 16.5).
- **`GetSetting`/`SaveSetting`** : subsistent mais déconseillés ; **emplacement de stockage à
  vérifier**. ⚠️
- **Registre 32/64 bits** : attention à `WOW6432Node` ; ciblez la vue voulue avec `OpenBaseKey` +
  `RegistryView`. ⚠️
- **Ne faites pas de pont là où une réécriture suffit** : le FSO via RCW, ou le registre via P/Invoke,
  sont des erreurs — préférez le managé natif.

---

## En résumé

- Les services « runtime » de VB6 se **remplacent** par du .NET natif, **sans interop** : c'est plus
  simple, plus sûr, et cela réduit la surface COM/native à maintenir.
- **`App`** → **`My.Application.Info`** (+ `Application.StartupPath`, `Mutex` pour l'instance unique).
  Attention au piège `App.Path` ≠ répertoire de travail et au format de version.
- **`FileSystemObject`** et les E/S intrinsèques VB6 → **`System.IO`** / **`My.Computer.FileSystem`**,
  avec le réflexe **`Using`** sur tous les flux (`IDisposable`). Les fichiers binaires/à accès direct
  restent le cas délicat.
- **Registre** → **`Microsoft.Win32.Registry`** (`RegistryKey` jetable → `Using`) / `My.Computer.Registry` ;
  mais la meilleure réponse pour la configuration est souvent **`My.Settings`** / `app.config`, **sans
  registre**.
- L'espace **`My`** est le trait d'union convivial qui rend ces remplacements proches de l'esprit VB6.

---

## Le chapitre 16 en perspective

Vous disposez désormais des deux faces de l'interopérabilité : **faire le pont** quand c'est utile —  
RCW pour consommer du COM (16.1), CCW pour exposer du .NET et migrer **par étapes** (16.2), le bon
**liage** (16.3), le **P/Invoke** pour l'API Windows (16.4), le **marshaling** des chaînes et
structures (16.5) — et **ne pas le faire** quand une réécriture native suffit (16.6). Le fil rouge :  
l'interop est un **outil de transition**, à employer à bon escient et à **réduire** au fil de la  
migration.

> ➡️ **Direction la partie 6 — *Finaliser, valider et moderniser*** : activer `Option Strict`
> progressivement, traquer les pièges silencieux et mener les tests de non-régression (module 17),
> avant le déploiement et la bascule (module 18).

---

🏷️ **Indicateurs** : 🔗 Interop · lié aux modules 12.3 (`IDisposable`/`Using`), 7.4 et 12.7 (fichiers à enregistrements), 17.2 (réduction des dépendances)
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026 · `System.IO` · espace `My` · `Microsoft.Win32`

⏭️ [Valider, fiabiliser et refactoriser après migration](/17-valider-refactoriser/README.md)
