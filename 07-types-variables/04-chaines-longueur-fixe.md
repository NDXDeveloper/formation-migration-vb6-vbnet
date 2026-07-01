🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.4 — Chaînes de longueur fixe (`Dim s As String * 10`) : `VBFixedString` ou refactoring ⚠️

> **Chapitre 7 — Types, variables et déclarations** · Partie 3 — Migrer le langage
>
> VB6 connaît un type particulier : la **chaîne de longueur fixe**, `Dim s As String * 10`. Elle
> occupe **toujours** 10 caractères, **complète par des espaces** ce qui est trop court et
> **tronque** ce qui est trop long. .NET **n'a aucun équivalent natif** : sa `String` est
> **immuable** et **de longueur variable**. La migration passe donc par un attribut de
> compatibilité (`VBFixedString`) ou par un **refactoring** — et, entre les deux, par des
> changements de comportement **silencieux**.

---

## 🧭 Rappel VB6 : ce que fait `String * n`

```vb
' VB6
Dim s As String * 10
s = "abc"                  ' s vaut en réalité "abc       " (complété à 10 par des espaces)
s = "Bonjour tout le monde"' s vaut "Bonjour to" (TRONQUÉ à 10)
```

Deux comportements définissent ce type — et ce sont eux qui piègent :

1. **Remplissage** : une valeur **plus courte** est complétée par des **espaces** à droite,
   jusqu'à la longueur déclarée.
2. **Troncature** : une valeur **plus longue** est **coupée** à la longueur déclarée.

La taille en mémoire est **fixe et connue**. C'est précisément pour cela que VB6 l'utilisait :

- **enregistrements de fichiers à accès direct** (`Type` avec des champs `String * n`,
  `Open … For Random`, `Get #` / `Put #`) — la disposition binaire est figée ; *l'usage n°1* ;
- **tampons d'API Win32** (passer un buffer pré-dimensionné que l'API remplit) ;
- **protocoles binaires / formats à largeur fixe** (interop, échanges) ;
- **alignement** ou remplissage de colonnes à l'affichage.

---

## 🎯 .NET : pas de chaîne de longueur fixe

En .NET, **`String` est immuable et de longueur variable** : il n'existe **pas** de type
« chaîne de N caractères ». Selon la **raison d'être** du `String * n`, on dispose de mécanismes
différents :

- **`<VBFixedString(n)>`** — attribut de **compatibilité** pour les **entrées/sorties de fichier**
  `Microsoft.VisualBasic` ;
- **`<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=n)>`** — pour la **disposition mémoire native**
  (structures passées à Win32/COM) ;
- **`StringBuilder(capacité)`** — pour les **tampons** remplis par une API ;
- **`PadRight`/`PadLeft`, `Substring`, `TrimEnd`** — pour gérer largeur et remplissage
  **explicitement**.

---

## 💥 Le changement de comportement (le vrai piège)

Quand un `String * n` devient une `String` ordinaire, **le remplissage et la troncature  
disparaissent**. C'est silencieux, et cela change la logique.

### L'égalité bascule (remplissage perdu)

```vb
' VB6 : s est un String * 10
s = "abc"                  ' s = "abc       "
If s = "abc" Then ...      ' FAUX  (s ≠ "abc" à cause du padding)
If Trim$(s) = "abc" Then   ' VRAI  (l'idiome habituel)

' VB.NET : s est un String normal
s = "abc"                  ' s = "abc" (longueur 3)
If s = "abc" Then ...      ' VRAI désormais
```

> ⚠️ Le **résultat d'une comparaison s'inverse**. De même, **`Len(s)` change** (10 → 3),
> et tout code qui s'appuyait sur la présence des **espaces de fin** (export à largeur fixe,
> concaténation, écriture en base) se comporte différemment.

### La troncature ne se fait plus toute seule

```vb
' VB6 : String * 10 → troncature automatique
s = "Bonjour tout le monde"   ' s = "Bonjour to"

' VB.NET : String normal → AUCUNE troncature
s = "Bonjour tout le monde"   ' s = la chaîne entière
' → débordement si la cible (colonne CHAR(10), protocole…) attend 10 caractères
```

> ⚠️ Le code qui **comptait sur** la troncature (pour respecter une largeur max) doit la
> **réintroduire explicitement** :
> ```vb
> s = valeur.Substring(0, Math.Min(valeur.Length, 10))
> ```

### La disposition des fichiers binaires se casse

Un fichier à accès direct écrit par VB6 a des **champs de largeur fixe**. Migrer le `Type` →
`Structure` avec des `String` ordinaires **modifie la taille des enregistrements** → les fichiers
de données existants deviennent **illisibles**. Il faut soit **préserver** le format (attribut
`VBFixedString` + E/S `Microsoft.VisualBasic`), soit **ré-architecturer** la persistance.

---

## 🩹 Option A — `<VBFixedString(n)>` : la béquille (et ses limites)

L'attribut s'applique à un champ `String`, surtout **dans une structure utilisée avec les fonctions  
de fichier `Microsoft.VisualBasic`** (`FileOpen`/`FilePut`/`FileGet`), pour **préserver la  
disposition** des enregistrements à accès direct.

```vb
' VB6
Private Type Client
    Nom  As String * 30
    Code As String * 5
End Type

' VB.NET — préserver le format binaire des fichiers à accès direct
Structure Client
    <VBFixedString(30)> Public Nom  As String
    <VBFixedString(5)>  Public Code As String
End Structure
```

> 🚨 **Le piège de l'attribut :** `VBFixedString` **n'impose PAS la longueur dans le code
> ordinaire**. Le champ reste une `String` **normale** : `Nom = "abc"` donne `"abc"` (longueur 3),
> **pas** une chaîne complétée à 30. Le remplissage/la troncature à 30 ou 5 **n'ont lieu que** via
> les **fonctions de fichier** `Microsoft.VisualBasic`. L'attribut est donc une **information de
> taille pour l'E/S**, pas un véritable type de longueur fixe.

**Quand l'utiliser :** des **fichiers à accès direct existants** qu'il faut continuer à lire/écrire
au **même format**, pendant et juste après la migration.
**Réserve :** c'est une **dépendance à `Microsoft.VisualBasic.Compatibility`/aux fonctions
héritées**, que l'on cherchera à **retirer** ensuite (**module 17.2**).

---

## 🔧 Option B — `<MarshalAs(...)>` : pour l'interop **native**

Si la chaîne fixe servait à une **structure passée à du code natif** (Win32/COM) — et non à de  
l'E/S de fichier VB6 — c'est la **disposition mémoire non managée** qu'il faut contrôler :

```vb
' Structure passée à du code natif (Win32/COM)
<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
Structure DonneesNatives
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=30)> Public Nom As String
End Structure
```

C'est le bon outil quand le `String * n` était là pour des raisons **d'API/COM** (détails au
**module 16.5 — *marshaling***, ANSI/Unicode, `StructLayout`).

---

## ✨ Option C — Refactoriser vers une `String` (le chemin propre)

Remplacer `String * n` par une `String` ordinaire et rendre largeur et remplissage **explicites**  
là où ils comptent :

```vb
s = s.PadRight(10)          ' compléter par des espaces (ou PadRight(10, "0"c))
s = s.PadLeft(10)           ' aligner à droite
s = s.Substring(0, Math.Min(s.Length, 10))   ' imposer une longueur max
nom = s.TrimEnd()           ' retirer les espaces de fin que VB6 ajoutait silencieusement
```

Pour les **tampons** remplis par une API, l'idiome .NET est **`StringBuilder`** :

```vb
' VB6 : tampon de longueur fixe rempli par l'API
Dim buf As String * 260
GetWindowsDirectory buf, Len(buf)
chemin = Left$(buf, InStr(buf, vbNullChar) - 1)

' VB.NET : StringBuilder avec capacité (le marshaleur gère la terminaison)
Dim buf As New StringBuilder(260)
GetWindowsDirectory(buf, buf.Capacity)
chemin = buf.ToString()
```

**Avantages :** idiomatique, intention claire, **aucune dépendance** de compatibilité.  
**Coût :** il faut **retrouver chaque endroit** qui dépendait de la largeur fixe.

> 🔗 `LSet`/`RSet` (justifier à gauche/droite dans une largeur, avec remplissage d'espaces) existent
> encore dans `Microsoft.VisualBasic`, mais opèrent sur des `String` ordinaires ; la variante VB6 de
> **`LSet` qui copiait des types définis** (UDT) **n'existe plus**. À refactoriser en
> `PadRight`/`PadLeft`/`Substring` (voir **chapitre 9.2**).

---

## 🧭 Grille de décision

| Pourquoi le `String * n` était-il là ? | Approche recommandée |
|----------------------------------------|----------------------|
| Enregistrement de **fichier à accès direct** (`Type` + `Get`/`Put`) | `<VBFixedString(n)>` (+ E/S `Microsoft.VisualBasic`) pour **préserver**, ou **ré-architecturer** la persistance |
| Champ d'une **structure passée à Win32/COM** | `<MarshalAs(ByValTStr, SizeConst:=n)>` (disposition native) |
| **Tampon** rempli par une API | `StringBuilder(capacité)` |
| Imposer une **largeur maximale** (colonne BD, protocole) | `String` + troncature **explicite** (`Substring`) |
| **Alignement** / remplissage d'affichage | `String` + `PadRight` / `PadLeft` |

---

## 🧭 Stratégie de migration

1. **Recenser tous les `String * n`** (déclarations isolées **et** champs de `Type`).
2. Pour chacun, **identifier la raison d'être** (fichier, interop native, largeur max, affichage,
   tampon) — c'est elle qui dicte l'approche via la grille ci-dessus.
3. **Préserver** d'abord ce qui doit l'être (formats de fichiers existants) avec `VBFixedString`,
   quitte à **refactoriser ensuite** vers une `String` propre.
4. **Réintroduire explicitement** remplissage et troncature partout où le code en **dépendait**
   (égalités, largeurs, exports), sans oublier de **retirer les espaces de fin** (`TrimEnd`).
5. **Filet de sécurité** : le ***golden master*** (chapitre 5.5) compare entrées/sorties à largeur
   fixe et révèle un padding ou une troncature manquants.

---

## ✅ À retenir

- **Aucun type de chaîne de longueur fixe en .NET** : `String` est **immuable** et **variable**.
- Les deux comportements de VB6 — **remplissage par espaces** (trop court) et **troncature** (trop
  long) — **disparaissent**, ce qui peut **inverser une égalité**, changer `Len`, ou provoquer un
  **débordement** de largeur.
- **`<VBFixedString(n)>`** = béquille pour les **E/S de fichier** `Microsoft.VisualBasic` ; il
  **n'impose pas** la longueur dans le code ordinaire et reste une **dépendance à retirer**.
- **`<MarshalAs(ByValTStr, …)>`** pour l'**interop native** ; **`StringBuilder`** pour les
  **tampons** d'API.
- Le **refactoring** vers `String` + `PadRight`/`Substring`/`TrimEnd` est le chemin **propre**, au
  prix de retrouver chaque dépendance à la largeur.

---

## 🔗 Pour aller plus loin

- **7.7 — `Nothing`, `Empty`, `Null`** : valeur par défaut et absence de valeur pour les chaînes.
- **Chapitre 9.2** — fonctions de chaînes (`Left`/`Right`/`Mid`, `LSet`/`RSet`) :
  `Microsoft.VisualBasic` vs `String`.
- **Chapitre 12.7** — `Type` → `Structure` : sémantique de valeur et disposition mémoire.
- **Module 16.5** — *marshaling*, ANSI/Unicode, `StructLayout` (interop native) ; **16.6** —
  `FileSystemObject` → `System.IO`.
- **Module 17.2** — supprimer les dépendances à `Microsoft.VisualBasic.Compatibility`.
- **Annexe D** (correspondance des types de données).

➡️ Section suivante : **[7.5 — Déclarations](05-declarations.md)**  
⬆️ Retour au **[sommaire du chapitre 7](README.md)**

⏭️ [Déclarations : `Dim` multiple, `DefInt`/`DefStr`…, portée, `As New` (instanciation implicite)](/07-types-variables/05-declarations.md)
