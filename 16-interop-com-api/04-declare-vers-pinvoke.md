🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.4 — Déclarations d'API Win32 : `Declare` → P/Invoke (`DllImport`) ⚠️

**Traduire les appels directs à l'API Windows : l'instruction `Declare` de VB6 cède la place au P/Invoke (`DllImport`) de .NET — une opération truffée de pièges silencieux, à commencer par la taille des types.**

> 🔗 Nous quittons les composants COM (RCW/CCW) pour l'**autre frontière** : les appels à des
> fonctions **natives** exportées par des DLL Windows (`kernel32`, `user32`, `gdi32`…) ou par des DLL
> tierces. En VB6, chaque `Declare Function … Lib …` était un tel appel. Ce module traite la
> **déclaration** et la **correspondance des types** ; le *marshaling* fin des chaînes et des
> structures fait l'objet du module **16.5**.

---

## Avant tout : faut-il *vraiment* appeler l'API ?

Beaucoup de `Declare` VB6 existaient parce que VB6 **n'offrait pas** la fonctionnalité nativement.  
Or, .NET en **enveloppe une grande partie**. Le premier réflexe n'est donc pas de traduire l'appel,  
mais de chercher l'**équivalent managé** — plus sûr, sans marshaling à gérer, et indépendant de la
« bitness ».

| `Declare` VB6 (API) | Équivalent .NET managé |
|---------------------|------------------------|
| `GetTickCount` | `Environment.TickCount` |
| `Sleep` | `System.Threading.Thread.Sleep` |
| `GetComputerName` | `Environment.MachineName` |
| `GetWindowsDirectory` / `GetSystemDirectory` | `Environment.GetFolderPath(...)` / `Environment.SystemDirectory` |
| `GetTempPath` / `GetTempFileName` | `IO.Path.GetTempPath` / `IO.Path.GetTempFileName` |
| `CopyFile` / `DeleteFile` / `MoveFile` | `IO.File.Copy` / `Delete` / `Move` |
| `RegOpenKeyEx`, `RegQueryValueEx`… | `Microsoft.Win32.Registry` / `My.Computer.Registry` (module 16.6) |

> 🎯 **Règle, cohérente avec tout le chapitre : on ne fait du P/Invoke que pour ce que .NET ne sait pas
> faire nativement.** Restent légitimes, par exemple : la manipulation fine de fenêtres
> (`FindWindow`/`SendMessage`), certaines API système sans équivalent managé, ou des **DLL natives
> tierces** (un SDK fourni en C). Pour ces cas, voici comment déclarer proprement.

---

## Ce que faisait VB6 : l'instruction `Declare`

Anatomie d'une déclaration VB6 typique :

```vb
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" ( _
    ByVal hWnd As Long, ByVal lpText As String, _
    ByVal lpCaption As String, ByVal wType As Long) As Long
```

- `Lib "user32"` : la **DLL** exportatrice.
- `Alias "MessageBoxA"` : le **nom réel** de l'export (souvent suffixé **`A`** — la variante **ANSI**,
  car VB6 passait des chaînes ANSI).
- La **liste des paramètres** (avec `ByVal`/`ByRef`) et le **type de retour**.

Deux caractéristiques de VB6 sont au cœur des pièges qui suivent : les chaînes étaient **ANSI** par  
défaut, et les **paramètres étaient `ByRef` par défaut** (module 2.4).

---

## En VB.NET : deux chemins

### Le chemin fidèle — `Declare` (conservé)

L'instruction `Declare` **existe toujours** en VB.NET, avec un modificateur `Ansi` / `Unicode` /
`Auto` :

```vbnet
Private Declare Unicode Function MessageBox Lib "user32.dll" Alias "MessageBoxW" ( _
    ByVal hWnd As IntPtr, ByVal lpText As String, _
    ByVal lpCaption As String, ByVal wType As UInteger) As Integer
```

C'est le **moindre changement** depuis VB6 — pratique pour un portage 1:1. Mais `Declare` est
**limité** : toujours convention `StdCall`, pas d'option pour capturer le code d'erreur Win32, etc.

### Le chemin idiomatique — `<DllImport>` (P/Invoke)

L'attribut **`<DllImport>`** (espace `System.Runtime.InteropServices`) est la forme recommandée :  
plus explicite et **bien plus contrôlable** (jeu de caractères, code d'erreur, convention d'appel,  
attributs de marshaling). La méthode doit être **`Shared`** (ou placée dans un `Module`) et son corps  
est **vide** (l'attribut fournit l'implémentation) :

```vbnet
Imports System.Runtime.InteropServices

<DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
Private Shared Function MessageBox(hWnd As IntPtr, lpText As String,
                                   lpCaption As String, wType As UInteger) As Integer
End Function
```

> 🎯 **Recommandation** : gardez `Declare` pour des portages triviaux en `StdCall` si cela réduit la
> charge de migration ; passez à **`<DllImport>`** dès que vous avez besoin de contrôle (jeu de
> caractères, `SetLastError`, convention d'appel) et pour tout code neuf. Les deux sont acceptables.

---

## ⚠️ Le piège n°1 : la correspondance des types

C'est **la** raison du panneau ⚠️ sur cette section. Le piège des entiers redimensionnés (module 7.2)
**ressurgit à la frontière native, où il est bien plus dangereux** : un paramètre de mauvaise taille
ne lève pas une erreur propre — il **corrompt la pile d'appel** ou passe des octets erronés.

| Type Win32 | `Declare` VB6 | **Type .NET correct** |
|------------|---------------|------------------------|
| `LONG`, `INT`, `DWORD` | `As Long` | **`Integer`** (32 bits) — *pas* `Long` ! |
| `SHORT`, `WORD` | `As Integer` | **`Short`** (16 bits) |
| `BYTE` | `As Byte` | `Byte` |
| `BOOL` | `As Long` | `Integer`, ou `Boolean` avec `<MarshalAs(UnmanagedType.Bool)>` |
| `HANDLE`, `HWND`, `HDC`, pointeur | `As Long` | **`IntPtr`** (voir ci-dessous) ⚠️ |
| `LPSTR` / `LPWSTR` (chaîne) | `As String` | `String` (selon `CharSet` — module 16.5) |
| `__int64`, `LARGE_INTEGER` | (malaisé en VB6) | `Long` (64 bits) |

> ⚠️ **Rappel décisif : `Long` en VB6 = 32 bits → `Integer` en .NET.** Traduire `As Long` par `Long`
> (qui fait 64 bits en .NET) est l'erreur la plus fréquente — et silencieuse.

### Les handles : **`IntPtr`**, jamais `Integer` ⚠️

VB6 rangeait les **handles** de fenêtres, de contextes de périphérique, etc., dans des `Long`
(32 bits). Mais sur un **Windows 64 bits**, un handle est un pointeur de **64 bits**. Si vous le
mappez sur `Integer`, votre code **fonctionne en 32 bits puis casse en 64 bits**.

Le type .NET correct pour **tout handle ou pointeur** est **`IntPtr`** : il s'adapte automatiquement à  
la « bitness » du processus (32 bits sur x86, 64 bits sur x64). C'est l'un des points les plus souvent  
oubliés d'un portage — et il rejoint directement la problématique de la « bitness » vue tout au long  
du chapitre.

---

## Les champs de `<DllImport>` qui comptent

L'attribut accepte plusieurs paramètres nommés ; quatre méritent votre attention :

- **`EntryPoint`** — le **nom réel** de l'export (l'équivalent de l'`Alias` VB6). Inutile si votre
  méthode porte déjà le nom de l'export.
- **`CharSet`** — `Ansi` / `Unicode` / `Auto`. Il pilote à la fois le marshaling des **chaînes** *et*
  la résolution du **suffixe `A`/`W`** du nom de fonction (`MessageBox` → `MessageBoxA` **ou**
  `MessageBoxW`). VB6 était **ANSI** par défaut ; les API Windows modernes sont nativement **Unicode
  (`W`)**. Choisir `Unicode` est généralement préférable aujourd'hui, mais **change le comportement**
  par rapport au défaut ANSI de VB6 — un risque de changement silencieux **détaillé au module 16.5**.
- **`SetLastError`** — mettez-le à **`True`** pour pouvoir lire ensuite le code d'erreur via
  **`Marshal.GetLastWin32Error()`**. C'est l'équivalent de l'**`Err.LastDllError`** de VB6. Sans ce
  drapeau, le code d'erreur n'est pas fiable.
- **`CallingConvention`** — **`StdCall`** par défaut (`__stdcall`, la convention de la quasi-totalité
  des API Win32) ou **`Cdecl`** (certaines DLL en C / *runtime* C). VB6 supposait toujours `StdCall`.
  Si vous appelez une **DLL native tierce en `__cdecl`** avec la mauvaise convention, **la pile est
  corrompue**. ⚠️ À vérifier impérativement pour les DLL non Windows.

---

## ⚠️ Le piège n°2 : le défaut `ByRef` → `ByVal` à la frontière native

Le retournement du défaut de passage (module 2.4) est **anodin entre deux fonctions .NET** ; à la  
frontière native, il est **catastrophique**. En VB6, un paramètre **non marqué** était `ByRef` : un
**pointeur** était passé, et l'API y **écrivait** un résultat. En VB.NET, le défaut est `ByVal` : la  
**valeur** est passée. Si l'API attendait un pointeur, **vous corrompez la mémoire** au lieu d'obtenir
une erreur propre.

```vb
' VB6 : lpdwProcessId est ByRef (défaut) → un pointeur est passé, l'API y écrit le PID
Private Declare Function GetWindowThreadProcessId Lib "user32" ( _
    ByVal hWnd As Long, lpdwProcessId As Long) As Long
```

```vbnet
' VB.NET : il FAUT marquer ByRef explicitement, sinon le défaut ByVal casse l'appel
<DllImport("user32.dll", SetLastError:=True)>
Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr,
                                                 ByRef lpdwProcessId As Integer) As Integer
End Function
```

> ⚠️ **Auditez chaque paramètre de chaque `Declare`** : tout ce qui s'appuyait sur le défaut `ByRef`
> de VB6 doit redevenir **explicitement `ByRef`** en .NET. C'est l'un des bugs silencieux les plus
> redoutables de cette frontière (et il figure au catalogue de l'annexe B).

---

## Exemple complet : `MessageBox` (VB6 → `Declare` → `<DllImport>`)

```vb
' (1) VB6 — variante ANSI, hWnd et wType en Long, paramètres ByVal explicites
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" ( _
    ByVal hWnd As Long, ByVal lpText As String, _
    ByVal lpCaption As String, ByVal wType As Long) As Long
```

```vbnet
' (2) VB.NET — P/Invoke : hWnd → IntPtr, Long → Integer/UInteger, Unicode (→ MessageBoxW)
Imports System.Runtime.InteropServices

<DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
Private Shared Function MessageBox(hWnd As IntPtr, lpText As String,
                                   lpCaption As String, wType As UInteger) As Integer
End Function
```

Notez la transformation : `hWnd` devient un **`IntPtr`** (c'est un handle !), les `Long` deviennent
`Integer`/`UInteger`, et `CharSet:=Unicode` fait résoudre le nom vers **`MessageBoxW`** — plus besoin
d'`Alias` pour ajouter le suffixe.

> ❌ **Cela dit, pour `MessageBox` précisément, n'utilisez pas P/Invoke.** L'équivalent managé est
> `MessageBox.Show(…)` (Windows Forms). Cet exemple n'illustre que la **mécanique** de traduction —
> exactement comme le FSO au module 16.1 : la bonne réponse reste souvent le framework natif.

---

## Chaînes et structures : le marshaling proprement dit

Deux sujets décisifs ont été seulement **effleurés** ici et sont traités en profondeur au module
**16.5** :

- **Les chaînes** : le choix `CharSet` (ANSI/Unicode) et son piège silencieux ; les **tampons de
  sortie** que l'API remplit, où l'on utilise un **`System.Text.StringBuilder`** plutôt qu'un `String`.
- **Les structures** : `Type … End Type` (VB6) → `Structure … End Structure`, avec
  `<StructLayout(LayoutKind.Sequential)>`, l'option `Pack`, et les `<MarshalAs>` pour les chaînes de
  longueur fixe et les tableaux.

---

## ⚠️ Récapitulatif des pièges

- **Types redimensionnés** : `As Long` (VB6, 32 bits) → **`Integer`** (et non `Long`) ; `As Integer`
  → **`Short`**. Une erreur silencieuse qui corrompt l'appel. ⚠️
- **Handles → `IntPtr`**, jamais `Integer` : sous peine de casse en **64 bits**. ⚠️
- **Convention d'appel** : `StdCall` par défaut ; passez à **`Cdecl`** pour les DLL natives qui
  l'exigent, sinon la **pile est corrompue**. ⚠️
- **Jeu de caractères** : VB6 = ANSI ; les API modernes = Unicode. Mauvais `CharSet` → **chaînes
  corrompues** ; changer de jeu modifie le comportement (détails au 16.5). ⚠️
- **Défaut `ByRef` → `ByVal`** : remarquez explicitement `ByRef` partout où VB6 s'appuyait sur son
  défaut, sinon **corruption mémoire**. ⚠️
- **`SetLastError:=True`** est nécessaire pour lire le code d'erreur via `Marshal.GetLastWin32Error()`
  (l'équivalent de `Err.LastDllError`).
- **Préférez le managé** : ne faites du P/Invoke que pour ce que .NET n'enveloppe pas déjà.

---

## En résumé

- Le **P/Invoke** (`<DllImport>`) est l'équivalent .NET de l'instruction **`Declare`** de VB6 pour
  appeler des fonctions **natives**. `Declare` subsiste (modificateurs `Ansi`/`Unicode`/`Auto`) comme
  chemin fidèle ; **`<DllImport>`** est la forme idiomatique et contrôlable, à privilégier.
- **Premier réflexe** : chercher l'**équivalent managé** ; ne P/Invoker que l'irréductible.
- **Le piège majeur est la taille des types** : `Long` VB6 = 32 bits → **`Integer`** ; et **tout
  handle/pointeur → `IntPtr`** (sinon casse en 64 bits).
- Les champs `<DllImport>` qui comptent : **`EntryPoint`** (= `Alias`), **`CharSet`** (ANSI/Unicode,
  + résolution `A`/`W`), **`SetLastError`** (→ `Marshal.GetLastWin32Error`), **`CallingConvention`**
  (`StdCall`/`Cdecl`).
- Le **défaut de passage s'inverse** (`ByRef` → `ByVal`) : à auditer paramètre par paramètre, car
  l'erreur **corrompt la mémoire** silencieusement.
- Le **marshaling** des chaînes et des structures — le cœur technique — est traité au module 16.5.

> ➡️ **Plongeons maintenant dans le marshaling : ANSI vs Unicode, tampons de chaînes, et la traduction
> des structures (`Type` → `Structure` avec `StructLayout`).** (Module 16.5)

---

🏷️ **Indicateurs** : 🔗 Interop · ⚠️ Pièges majeurs (taille des types, handles, `ByRef`/`ByVal`, convention d'appel) · lié au module 7.2
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026 · `System.Runtime.InteropServices`

⏭️ [Marshaling, **ANSI/Unicode**, types et structures (`Type` → `Structure` avec `StructLayout`)](/16-interop-com-api/05-marshaling.md)
