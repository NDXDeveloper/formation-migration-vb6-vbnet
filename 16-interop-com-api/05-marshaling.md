🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.5 — Marshaling, ANSI/Unicode, types et structures (`Type` → `Structure` avec `StructLayout`) ⚠️

**Le cœur technique de l'interopérabilité : comment les données franchissent réellement la frontière entre le monde managé (.NET, Unicode, ramassé par le GC) et le monde natif/COM — et pourquoi les chaînes et les structures concentrent l'essentiel des bugs silencieux.**

> 🔗 Le module 16.4 a posé la **déclaration** des appels d'API ; il a renvoyé ici tout ce qui touche à
> la **conversion des données**. C'est la section la plus dense du chapitre, et la plus piégeuse :
> chaîne ANSI vs Unicode, tampon de sortie, disposition mémoire d'une structure… autant de détails où
> une erreur ne lève pas une exception propre, mais **corrompt silencieusement** les données.

---

## Qu'est-ce que le marshaling ?

Le **marshaling**, c'est la **conversion automatique** des données entre les deux mondes lors d'un  
appel d'interop. Le CLR comporte un *marshaler* qui, à chaque appel, **emballe** vos arguments managés  
en représentation native, **appelle** la fonction COM/Win32, puis **déballe** les résultats. Il fait  
beaucoup de travail pour vous — mais **pas tout**, et ses choix par défaut ne coïncident pas toujours  
avec ceux que VB6 supposait.

La première question à se poser pour chaque type : **a-t-il besoin d'être converti, ou non ?**

---

## Blittable vs non-blittable — le concept fondateur

- **Type *blittable*** : sa représentation binaire est **identique** en mémoire managée et non managée.
  Aucune conversion : le marshaler passe un pointeur direct (l'objet est juste « épinglé » le temps de
  l'appel). Rapide et prévisible.
- **Type *non-blittable*** : il **doit être converti** (copié et transformé). C'est là que vivent les
  pièges.

| Catégorie | Exemples |
|-----------|----------|
| **Blittable** ✅ | `Byte`, `Short`, `Integer`, `Long`, `Single`, `Double`, `IntPtr` ; tableaux 1D et structures **composés uniquement** de types blittables |
| **Non-blittable** ⚠️ | `Boolean` (taille variable !), `Char` / `String` (encodage !), tableaux et structures contenant du non-blittable, `Decimal`, `DateTime` |

> 💡 **À retenir** : tant que vous restez en types blittables (entiers, flottants, `IntPtr`), le
> marshaling est trivial. Dès que des **chaînes**, des **booléens** ou des **structures complexes**
> entrent en jeu, il faut être explicite — c'est l'objet de tout ce qui suit.

---

## La question **ANSI / Unicode**

C'est **le** point central du marshaling des chaînes, et l'une des premières causes de migrations  
ratées.

- **VB6** passait les chaînes aux API en **ANSI** (le marshaler convertissait, et appelait les
  variantes de fonction suffixées **`A`**). Son `Declare` était ANSI par défaut.
- **.NET** : `String` et `Char` sont **Unicode (UTF-16)**.

Le réglage **`CharSet`** (sur `<DllImport>` ou `<StructLayout>`) décide de l'encodage **et** de la  
résolution du suffixe `A`/`W` (rappel du module 16.4) :

| `CharSet` | Encodage des chaînes | Fonction résolue |
|-----------|----------------------|------------------|
| `Ansi` | ANSI (1 octet/caractère, perte possible) | `…A` |
| `Unicode` | UTF-16 (2 octets/caractère) | `…W` |
| `Auto` | Défaut de la plateforme (Unicode sur Windows moderne) | selon plateforme |

> ⚠️ **Le piège du changement silencieux.** VB6 était ANSI ; les API Windows modernes sont nativement
> **Unicode**. Choisir `Unicode` est généralement préférable, mais **modifie le comportement** par
> rapport au défaut VB6 — et surtout, **un caractère ne pèse plus le même nombre d'octets** (1 en
> ANSI, 2 en Unicode). C'est dévastateur pour les **tampons de longueur fixe** : une taille calculée
> « en caractères » et une autre « en octets » ne coïncident plus. Si des données contiennent des
> caractères non-ASCII, un mauvais `CharSet` produit des chaînes **tronquées ou illisibles**.

> 🧩 **Distinction importante : le COM, lui, est toujours Unicode.** Une chaîne COM est un **BSTR**
> (UTF-16). La question ANSI/Unicode est donc essentiellement une affaire d'**API Win32 (P/Invoke)**,
> pas d'interop COM (RCW/CCW). Pour le COM, le marshaler gère les BSTR en Unicode sans que vous ayez à
> arbitrer le `CharSet`.

---

## Marshaler les chaînes

### En entrée (vous passez une chaîne à l'API)

Un `String` suffit ; `CharSet` pilote l'encodage. Le marshaler alloue un tampon natif, copie la  
chaîne, puis le libère après l'appel. Rien de spécial à faire — hormis choisir le bon `CharSet`.

### En sortie : **`StringBuilder`** — le grand classique ⚠️

Quand l'API **écrit** dans un tampon que vous fournissez (`GetWindowText`, `GetWindowsDirectory`…),  
un `String` est **inutilisable** : les chaînes .NET sont **immuables**. Il faut un
**`System.Text.StringBuilder`** dont vous **fixez la capacité** ; le marshaler passe son tampon
interne, et vous lisez le résultat avec `.ToString()` :

```vbnet
Imports System.Runtime.InteropServices
Imports System.Text

<DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
Private Shared Function GetWindowText(hWnd As IntPtr, lpString As StringBuilder,
                                      nMaxCount As Integer) As Integer
End Function

' Utilisation
Dim tampon As New StringBuilder(256)               ' capacité = taille max du tampon
GetWindowText(hWnd, tampon, tampon.Capacity)
Dim titre As String = tampon.ToString()
```

> ⚠️ **Capacité insuffisante = troncature ou débordement.** Dimensionnez le `StringBuilder` à la
> taille maximale attendue, et passez bien `tampon.Capacity` (et non une constante désynchronisée) à
> l'API.

### Contrôle fin avec `<MarshalAs>`

Pour forcer un type de chaîne précis, indépendamment du `CharSet`, on annote le paramètre ou le champ :

- `<MarshalAs(UnmanagedType.LPStr)>` — pointeur vers une chaîne **ANSI**.
- `<MarshalAs(UnmanagedType.LPWStr)>` — pointeur vers une chaîne **Unicode**.
- `<MarshalAs(UnmanagedType.LPTStr)>` — selon le `CharSet`.
- `<MarshalAs(UnmanagedType.BStr)>` — **BSTR** (cas COM).

---

## Le cas des booléens (1, 2 ou 4 octets — et `True = -1`)

`Boolean` est non-blittable, et **sa taille native varie** :

| Contexte | Type natif | Taille | Vrai |
|----------|-----------|--------|------|
| Win32 `BOOL` | `int` | **4 octets** | ≠ 0 |
| COM `VARIANT_BOOL` | `short` | **2 octets** | **-1** (0xFFFF) |
| .NET `Boolean` | — | 1 octet (managé) | — |

Bonne nouvelle : **pour un `BOOL` Win32, `As Boolean` fonctionne** avec le marshaling **par défaut**  
de P/Invoke (qui traite `Boolean` comme un `BOOL` 4 octets, soit `UnmanagedType.Bool`). En revanche,
**en interop COM**, le défaut de `Boolean` est `VARIANT_BOOL` (2 octets, **-1**). Pour lever toute
ambiguïté, soyez explicite : `<MarshalAs(UnmanagedType.Bool)>` (Win32) ou
`<MarshalAs(UnmanagedType.VariantBool)>` (COM).

> 🔗 Ce **`True = -1`** n'est pas un détail : c'est le piège booléen du module 7.6 (et de l'annexe
> B.8), vu ici sous l'angle du marshaling.

---

## `Type … End Type` → `Structure` avec `StructLayout`

Le sujet structurel majeur. Un **type défini par l'utilisateur** VB6 (`Type`) devient une
**`Structure`** VB.NET — mais pour la passer à du code natif, sa **disposition mémoire** doit
correspondre **exactement** à la définition C attendue.

### `LayoutKind.Sequential` (+ le piège des types redimensionnés)

`<StructLayout(LayoutKind.Sequential)>` dispose les champs **dans l'ordre de déclaration**, comme une
`struct` C. C'est ce que l'on veut presque toujours pour le P/Invoke. (Les structures VB.NET sont déjà
`Sequential` par défaut, mais être explicite est une bonne pratique — et nécessaire dès qu'on précise
`CharSet` ou `Pack`.)

```vb
' VB6
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
```

```vbnet
' VB.NET — attention : Long (VB6, 32 bits) → Integer (cf. modules 7.2 et 16.4) !
<StructLayout(LayoutKind.Sequential)>
Public Structure RECT
    Public Left As Integer
    Public Top As Integer
    Public Right As Integer
    Public Bottom As Integer
End Structure
```

> ⚠️ Le piège du **redimensionnement** s'applique **aussi à l'intérieur des structures** : chaque
> `As Long` y devient `Integer`. Une seule erreur décale tous les champs suivants.

### `Pack` — l'alignement des octets

Le **`Pack`** contrôle l'alignement des champs. Le défaut .NET convient à la plupart des structures
« simples » (champs de même taille), mais un struct C compilé avec `#pragma pack(1)` (sans
remplissage) exige `Pack:=1`. **Un `Pack` qui ne correspond pas à la définition native fait atterrir  
les champs à de mauvais décalages** → données illisibles. Reproduisez fidèlement l'alignement du C.

### Chaînes fixes et tableaux fixes **dans** une structure

```vb
' VB6
Private Type DEVICE_INFO
    Id As Long
    Name As String * 32        ' chaîne de longueur fixe
    Flags(0 To 3) As Byte      ' tableau de taille fixe
End Type
```

```vbnet
' VB.NET
<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
Public Structure DEVICE_INFO
    Public Id As Integer

    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)>   ' String * 32 → chaîne inline
    Public Name As String

    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=4)>   ' tableau inline de 4 octets
    Public Flags As Byte()
End Structure
```

- **`String * n`** (VB6) → **`<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=n)>`** sur un `String`
  (le sujet général des chaînes de longueur fixe est traité au module 7.4).
- **Tableau de taille fixe** → **`<MarshalAs(UnmanagedType.ByValArray, SizeConst:=n)>`**.
- ⚠️ **La taille est en *caractères/éléments*, pas en octets** : avec `CharSet.Unicode`, un
  `ByValTStr` de 32 occupe **64 octets**. Ce nombre doit coller à la définition C, jeu de caractères
  inclus.

### Passer une structure (souvent `ByRef`)

Beaucoup d'API attendent un **pointeur** vers la structure → on passe **`ByRef`** (rappel du module
16.4) :

```vbnet
<DllImport("user32.dll", SetLastError:=True)>
Private Shared Function GetClientRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
End Function

' Utilisation
Dim zone As RECT
If GetClientRect(hWnd, zone) Then
    Dim largeur As Integer = zone.Right - zone.Left
End If
```

### `LayoutKind.Explicit` / `<FieldOffset>` — les unions

Pour des champs **superposés** (une *union* C) ou un contrôle d'offset au plus près, on utilise
`<StructLayout(LayoutKind.Explicit)>` avec un `<FieldOffset(n)>` sur **chaque** champ. VB6 ne proposait
pas d'unions simples, mais certaines structures système l'exigent. À réserver aux cas qui le  
nécessitent vraiment.

---

## La boîte à outils `Marshal` (pour le marshaling manuel)

Quand le marshaling automatique ne suffit pas (tableaux de taille variable, pointeurs vers pointeurs,  
gestion fine de tampons natifs), la classe **`Marshal`** offre l'outillage bas niveau :

- `Marshal.SizeOf(...)` — taille native d'un type/instance.
- `Marshal.PtrToStructure(...)` / `Marshal.StructureToPtr(...)` — conversion structure ↔ pointeur.
- `Marshal.AllocHGlobal(...)` / `Marshal.FreeHGlobal(...)` — allouer/libérer un tampon natif (⚠️ **à
  vous** de libérer).
- `Marshal.Copy(...)` — copier entre tableau managé et mémoire native.
- `Marshal.StringToHGlobalAnsi/Uni(...)` et `Marshal.PtrToStringAnsi/Uni(...)` — chaînes ↔ mémoire
  native.

> 💡 Ces API sont puissantes mais **manuelles** : vous reprenez la gestion mémoire que le COM gérait
> par comptage de références. À n'employer que lorsque l'automatique est insuffisant.

---

## Les *callbacks* : `Delegate` → pointeur de fonction ⚠️

Certaines API Win32 **rappellent** votre code : elles attendent un **pointeur de fonction**
(`EnumWindows`, `SetWindowsHookEx`, `SetTimer`…). En VB6, on le fournissait péniblement via
**`AddressOf`** (limité aux **modules standard**). En .NET, on déclare un **`Delegate`** dont la
signature correspond au callback attendu, et le marshaler le convertit en pointeur de fonction natif
(via un *thunk*) :

```vbnet
' Le TYPE du callback (signature imposée par l'API)
Private Delegate Function EnumWindowsProc(hWnd As IntPtr, lParam As IntPtr) As Boolean

<DllImport("user32.dll")>
Private Shared Function EnumWindows(callback As EnumWindowsProc, lParam As IntPtr) As Boolean
End Function
```

> ⚠️ **Le piège mortel : le delegate collecté par le GC.** Le marshaler ne crée **aucune** référence
> managée durable vers votre delegate. Si vous le passez « en vol » et que l'API **conserve** le
> pointeur **au-delà** de l'appel (*hook*, timer), le GC peut **collecter** le delegate → le prochain
> rappel déclenche une **violation d'accès** (assistant de débogage `CallbackOnCollectedDelegate`).
> La panne est **aléatoire** — elle dépend du moment du GC. **Parade** : conservez le delegate dans un
> **champ** vivant aussi longtemps que le natif peut l'appeler (ou `GC.KeepAlive`). Pour un
> `EnumWindows` **synchrone**, le CLR le protège le temps de l'appel ; mais gardez le réflexe dès
> qu'une API **stocke** le callback pour plus tard.

---

## `Currency` et `Date` : rappels de marshaling

- **`Currency`** (VB6, entier mis à l'échelle) ↔ **`CY`** (COM) ↔ **`Decimal`** (.NET), au besoin avec
  `<MarshalAs(UnmanagedType.Currency)>`. C'est le bon type pour la **précision monétaire** (modules 7.2
  et annexe B.9).
- **`Date`** (VB6, un `Double`) ↔ **`DATE`** (COM) ↔ **`DateTime`** (.NET). En **interop COM**, la
  conversion `DateTime` ↔ `DATE` est **automatique** ; au niveau **API**, un `DATE` est un `Double` que
  l'on convertit avec `DateTime.ToOADate` / `DateTime.FromOADate` (modules 7.3 et B.9).

---

## ⚠️ Récapitulatif des pièges

- **ANSI vs Unicode** : VB6 = ANSI, API modernes = Unicode. Mauvais `CharSet` → **chaînes corrompues** ;
  et surtout, **un caractère change de taille en octets** → tampons fixes faussés. ⚠️
- **Tampons de sortie** : jamais un `String` (immuable) — un **`StringBuilder`** à **capacité fixée**,
  sinon **troncature/débordement**. ⚠️
- **Disposition des structures** : `Pack`/alignement **doivent** correspondre à la définition C, sinon
  champs à de mauvais décalages → **données illisibles**. ⚠️
- **Types redimensionnés dans les structures** : `As Long` → **`Integer`** ; une erreur décale toute
  la suite. ⚠️
- **Booléens** : taille 1/2/4 octets selon le contexte, **`True = -1`** pour `VARIANT_BOOL` (COM) ;
  défauts différents en P/Invoke (4 octets) et en COM (2 octets). ⚠️
- **Tailles fixes** : `ByValTStr`/`ByValArray` exigent un `SizeConst` exact, **en caractères/éléments**
  (× 2 octets en Unicode).
- **Marshaling manuel** : ce que vous allouez avec `AllocHGlobal`, **vous** devez le libérer ; n'utilisez
  pas un pointeur natif au-delà de l'appel.
- **COM = Unicode** : la question ANSI/Unicode ne concerne (presque) que les API Win32, pas le COM.

---

## En résumé

- Le **marshaling** convertit les données entre managé et natif/COM. Les types **blittables** (entiers,
  flottants, `IntPtr`) passent sans conversion ; les **non-blittables** (chaînes, booléens, structures
  complexes) exigent d'être explicite.
- La question **ANSI/Unicode** est centrale pour les **API Win32** : `CharSet` pilote l'encodage **et**
  la résolution `A`/`W`. VB6 était ANSI, les API modernes Unicode — d'où un risque de **changement
  silencieux** et de tampons mal dimensionnés. Le **COM**, lui, est toujours **Unicode (BSTR)**.
- Les **chaînes en sortie** se marshalent avec un **`StringBuilder`** à capacité fixée, pas un `String`.
- Les **structures** : `Type` → `Structure` avec **`<StructLayout(LayoutKind.Sequential)>`**, en
  veillant à `Pack`/alignement, au **redimensionnement des types** (`Long`→`Integer`), et aux champs de
  taille fixe (`ByValTStr` pour `String * n`, `ByValArray` pour les tableaux). On les passe souvent
  **`ByRef`**.
- Les **booléens** ont une taille native variable et le piège `True = -1` (COM) — à relier au module
  7.6 / B.8.
- La classe **`Marshal`** fournit l'outillage **manuel** pour les cas que l'automatique ne couvre pas.

> ➡️ **Dernière étape du chapitre : les services « runtime » de VB6 — l'objet `App`, le registre, le
> `FileSystemObject` — que l'on remplace, le plus souvent *sans* interop, par `System.IO` et l'espace
> `My`.** (Module 16.6)

---

🏷️ **Indicateurs** : 🔗 Interop · ⚠️ Pièges majeurs (ANSI/Unicode, tampons, disposition des structures) · lié aux modules 7.2, 7.3, 7.4, 7.6 et aux annexes B.8/B.9
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026 · `System.Runtime.InteropServices`

⏭️ [L'objet `App`, le registre, et `FileSystemObject` → `System.IO` / espace `My`](/16-interop-com-api/06-app-registre-fso.md)
