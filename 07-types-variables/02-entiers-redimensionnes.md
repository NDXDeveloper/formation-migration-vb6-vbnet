🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.2 — Entiers redimensionnés : `Integer`→`Short`, `Long`→`Integer`, `Currency`→`Decimal` ⭐ ⚠️

> **Chapitre 7 — Types, variables et déclarations** · Partie 3 — Migrer le langage
>
> Si une seule section de toute la formation devait être lue les yeux grands ouverts, c'est
> celle-ci. Le piège est d'une simplicité redoutable : **les mots-clés numériques n'ont pas changé
> de nom, mais ils ont changé de taille.** Un `Dim x As Integer` recopié tel quel **compile**,
> **tourne**… et ne désigne plus le même type. C'est l'entrée **B.2** du catalogue des pièges
> silencieux — et la première cause de régressions sournoises.

---

## 🧨 Le fait central : même mot-clé, taille différente

| Mot-clé | En **VB6** | En **VB.NET** |
|---------|------------|---------------|
| `Byte` | 8 bits non signé (0 → 255) | 8 bits non signé (0 → 255) — *inchangé* |
| `Integer` | **16 bits** (-32 768 → 32 767) | **32 bits** (≈ ±2,1 milliards) |
| `Long` | **32 bits** (≈ ±2,1 milliards) | **64 bits** (≈ ±9,2 × 10¹⁸) |
| `Currency` | **64 bits**, 4 décimales fixes | *(n'existe plus)* |
| *(aucun)* | — | `Decimal` : 128 bits, ~28-29 chiffres |
| *(aucun)* | — | `Long` 64 bits, `Short` 16 bits, types **non signés** |

Le glissement saute aux yeux quand on l'aligne : **`Integer` et `Long` ont chacun doublé de  
taille**, et le mot `Long` désigne maintenant un entier **64 bits** qui n'existait pas en VB6.

> ⚠️ **Pourquoi c'est silencieux ?** Parce qu'un type **plus large contient** l'ancien. Une valeur
> qui tenait sur 16 bits tient évidemment sur 32. Dans 95 % du code, « ça marche » sans rien dire.
> Le bug n'apparaît qu'aux **frontières** : débordements, interop, fichiers binaires, structures.

---

## 🎯 La bonne équivalence : raisonner par la **taille**, pas par le **nom**

La règle d'or de la migration numérique : **on conserve la taille, donc on change souvent le  
mot-clé.**

| Type **VB6** | Taille | Équivalent **VB.NET** (même taille) |
|--------------|--------|-------------------------------------|
| `Byte` | 8 bits | `Byte` |
| `Integer` | 16 bits | **`Short`** |
| `Long` | 32 bits | **`Integer`** |
| `Currency` | 64 bits (4 déc.) | **`Decimal`** |

```vb
' VB6                          ' VB.NET (équivalence par la taille)
Dim compteur As Integer    →   Dim compteur As Short      ' 16 bits → 16 bits
Dim total    As Long       →   Dim total    As Integer    ' 32 bits → 32 bits
Dim montant  As Currency   →   Dim montant  As Decimal    ' money → Decimal
```

> 💡 **Nuance importante.** Conserver strictement la taille est **obligatoire** pour l'**interop**,
> les **fichiers binaires** et les **structures** (voir plus bas). Mais pour une variable purement
> « métier » (un compteur, un identifiant interne), élargir n'est pas un problème — on peut alors
> **garder `Integer`** (32 bits) sans risque, voire en profiter. **Le danger, ce n'est pas la
> largeur — c'est le décalage silencieux là où la taille compte.**

---

## 💥 Où ça passe… puis où ça casse

### 1. Le **débordement** change de seuil

En VB6, dépasser la capacité d'un `Integer` (16 bits) lève une **erreur « Overflow »** (erreur 6).  
En VB.NET, le **contrôle de débordement est actif par défaut** lui aussi (et lève une
`System.OverflowException`) — **mais le seuil a changé**, puisque le type est désormais 32 bits.

```vb
' VB6 : Integer = 16 bits
Dim i As Integer
i = 30000
i = i + 10000      ' ❌ 40000 > 32767 → erreur « Overflow »

' VB.NET, traduit naïvement : Integer = 32 bits
Dim i As Integer
i = 30000
i = i + 10000      ' ✅ 40000 : aucun débordement, résultat silencieusement différent
```

> ⚠️ **Conséquence :** un calcul qui *échouait* (ou bouclait, ou s'enroulait) à 32 767 ne le fait
> plus. Tout code qui **comptait sur** la borne 16 bits — compteur cyclique, détection volontaire
> d'overflow, masque — se comporte différemment. (Le contrôle de débordement est détaillé plus
> bas.)

### 2. Interop, API Win32, fichiers binaires et **structures** — la zone rouge 🔴

C'est ici que le piège devient **un vrai bug**, pas une simple différence de seuil. Dès que la
**taille en octets** est significative, le décalage corrompt les données.

**Structures / `Type` (disposition mémoire).** Un enregistrement change de taille si on garde les
mots-clés :

```vb
' VB6 — 8 octets
Private Type Enregistrement
    Id   As Integer   ' 2 octets
    Code As Integer   ' 2 octets
    Prix As Long      ' 4 octets
End Type

' VB.NET, traduit NAÏVEMENT — 16 octets ! (le double)
Structure Enregistrement
    Dim Id   As Integer   ' 4 octets (32 bits)
    Dim Code As Integer   ' 4 octets
    Dim Prix As Long      ' 8 octets (64 bits) !!
End Structure

' VB.NET, migration CORRECTE — 8 octets, disposition préservée
<StructLayout(LayoutKind.Sequential)>
Structure Enregistrement
    Dim Id   As Short     ' 2 octets
    Dim Code As Short     ' 2 octets
    Dim Prix As Integer   ' 4 octets
End Structure
```

**API Win32 (`Declare` → `P/Invoke`).** Un paramètre `WORD` (16 bits) déclaré `As Integer` en VB6  
**doit** devenir `Short` en .NET, pas `Integer`. Sinon, la pile d'arguments est mal alignée et
l'appel lit ou écrit de la mémoire erronée.

> 🪤 **Faux-ami à double détente :** en terminologie **Win32**, le type `LONG` fait **32 bits** — il
> correspond donc à `Integer` en VB.NET, **surtout pas** à `Long` (64 bits) ! C'est exactement le
> genre de confusion que la migration doit traquer (détails au **module 16.4** : `Declare` →
> `P/Invoke`, et **16.5** : *marshaling*, `StructLayout`).

**Fichiers binaires à format fixe** (`Open … For Random`, `Get #`, `Put #`). La **taille de chaque
enregistrement** dépend de la taille des champs. Un champ `Integer` qui passe de 2 à 4 octets
**décale tout** : les fichiers de données existants deviennent **illisibles**. Là encore, il faut
`Short` pour préserver la disposition.

---

## 💰 `Currency` → `Decimal` : la bonne cible (et le piège du `Double`)

En VB6, `Currency` est un type **monétaire à virgule fixe** : 64 bits, **toujours 4 décimales**,  
plage ≈ **±922 000 milliards** (exactement ±922 337 203 685 477,5807 — un `Int64` divisé par
10 000). En .NET, son successeur est **`Decimal`** : 128 bits, ~28-29 chiffres significatifs,
**arithmétique décimale exacte**, précision **variable**.

> 🚫 **Ne jamais migrer `Currency` vers `Double`.** `Double` est un flottant **binaire** : il ne
> peut pas représenter exactement `0,10` et accumule des erreurs d'arrondi sur l'argent — une faute
> classique et coûteuse. **`Currency` → `Decimal`, point final.** (Voir **Annexe B.9** : dates et
> arrondis monétaires.)

```vb
' VB6                              ' VB.NET
Dim prix As Currency           →   Dim prix As Decimal     ' ✅
                                   Dim prix As Double       ' ❌ erreurs d'arrondi monétaires
```

**Différences à connaître malgré la bonne cible :**

- **Précision/échelle.** `Currency` est **figé à 4 décimales** ; `Decimal` en garde **plus**. Un
  enchaînement `a * b / c` peut donc produire des **chiffres de fin différents** de l'ancien
  calcul à virgule fixe. Tout code qui supposait « exactement 4 décimales » doit être **revérifié**
  (arrondir explicitement si nécessaire).
- **Quand ce n'était pas de l'argent.** En VB6, `Currency` servait parfois de **conteneur 64 bits**
  (pour de grands entiers, des compteurs de précision, certains appels d'API). Dans ce cas, la
  bonne cible peut être **`Long`** (entier 64 bits, désormais natif) plutôt que `Decimal`. À juger
  au cas par cas, selon l'**usage réel**.

---

## 🔣 Les suffixes de type : un glissement, lui aussi silencieux

Les **caractères de déclaration de type** existent toujours, mais leur signification a bougé :

| Suffixe | En **VB6** | En **VB.NET** |
|---------|------------|---------------|
| `%` | `Integer` (16 bits) | `Integer` (**32 bits**) |
| `&` | `Long` (32 bits) | `Long` (**64 bits**) |
| `@` | **`Currency`** | **`Decimal`** |
| `!` | `Single` | `Single` |
| `#` | `Double` | `Double` |
| `$` | `String` | `String` |

```vb
Dim n%      ' VB6 : entier 16 bits   →   VB.NET : entier 32 bits
Dim taille& ' VB6 : entier 32 bits   →   VB.NET : entier 64 bits
Dim prix@   ' VB6 : Currency         →   VB.NET : Decimal
```

> 💡 Conseil de migration : **supprimer les suffixes** au profit de déclarations explicites
> (`As Short`, `As Integer`, `As Decimal`). C'est plus lisible et cela force à **choisir
> consciemment** la taille — exactement ce que ce piège réclame.

---

## 🔄 Les fonctions de conversion : `CInt`, `CLng`, `CCur`…

Mêmes noms, mais elles **convertissent vers les nouvelles tailles** :

| Fonction | En **VB6** | En **VB.NET** |
|----------|------------|---------------|
| `CInt(x)` | → `Integer` (16 bits) | → `Integer` (**32 bits**) |
| `CLng(x)` | → `Long` (32 bits) | → `Long` (**64 bits**) |
| `CCur(x)` | → `Currency` | **supprimée** |

- Pour une conversion **16 bits** (l'ancien `CInt` de VB6), utiliser le **nouveau `CShort()`**.
- L'équivalent **32 bits** de l'ancien `CLng` est désormais **`CInt()`**.
- `CCur` n'existe plus : la convertir en **`CDec()`** (Decimal).
- Nouveautés .NET : `CShort`, `CByte`, et les non signés `CUShort`/`CUInt`/`CULng`, `CSByte`.

```vb
' VB6                       ' VB.NET (intention 16 bits préservée)
n = CInt(valeur)        →   n = CShort(valeur)      ' si l'on veut vraiment 16 bits
m = CCur(montant)       →   m = CDec(montant)       ' CCur supprimée
```

---

## 🛡️ Le contrôle de débordement en .NET

Par défaut, **VB.NET vérifie les débordements arithmétiques** et lève `OverflowException` — comme  
VB6 levait l'erreur 6. Deux points de vigilance :

1. **Le seuil a changé** avec la taille (32 bits au lieu de 16 pour `Integer`) : un calcul qui
   débordait en VB6 peut désormais **réussir silencieusement**.
2. Il existe une option de compilation **« Supprimer les contrôles de dépassement d'entier »**
   (*Remove integer overflow checks*). Si elle est activée pour la performance, un dépassement ne
   lève **plus rien** et « enroule » la valeur. À utiliser en connaissance de cause.

---

## ⚠️ Pièges silencieux — récapitulatif

- **Mot-clé conservé, taille doublée** : `Integer` (16→32), `Long` (32→64). Compile et tourne, mais
  fausse l'**interop**, les **structures** et les **fichiers binaires**.
- **Disposition mémoire** : une structure ou un enregistrement binaire **change de taille** si l'on
  garde les mots-clés → données corrompues / fichiers illisibles. Utiliser `Short` + `StructLayout`.
- **Win32 `LONG` = 32 bits** → `Integer` en .NET, jamais `Long` (64 bits).
- **`Currency` → `Double`** : faute d'arrondi monétaire. Cible = **`Decimal`** (ou `Long` si c'était
  un grand entier).
- **Suffixes `%` / `&` / `@`** : sens silencieusement modifié.
- **`CInt` / `CLng`** ciblent désormais 32/64 bits ; **`CCur` supprimée** → `CDec`.
- **Débordement** : seuil déplacé ; un overflow VB6 peut devenir un calcul silencieux en .NET.

---

## 🧭 Stratégie de migration

1. **Auditer chaque déclaration numérique** (`Integer`, `Long`, `Currency`, suffixes `%`/`&`/`@`,
   `CInt`/`CLng`/`CCur`). C'est un **passage systématique**, pas un survol.
2. **Mapper par la taille** quand la taille compte (interop COM, `Declare`/`P/Invoke`, `Structure`
   avec `StructLayout`, accès binaire/`Random`) : `Integer`→`Short`, `Long`→`Integer`.
3. **Pour le métier pur**, élargir est acceptable (garder `Integer` 32 bits), mais **par choix
   conscient**.
4. **Argent → `Decimal`** ; **grand entier 64 bits → `Long`** ; **jamais `Double` pour de la
   monnaie**.
5. **Garder le contrôle de débordement actif** pendant la validation, pour faire **émerger** les
   dépassements aux nouveaux seuils.
6. **S'appuyer sur le *golden master*** (chapitre 5.5) : c'est lui qui révèle, chiffres à l'appui,
   qu'un type redimensionné a changé un résultat ou un arrondi.

> 🛠️ Les outils commerciaux (VBUC, VB Migration Partner — module 4.3) gèrent une grande part de ce
> remappage automatiquement, **mais pas tout** : la décision « élargir ou préserver » et les cas
> d'interop restent à **valider à la main**.

---

## ✅ À retenir

- **`Integer`→`Short`, `Long`→`Integer`, `Currency`→`Decimal`** : on raisonne en **taille**, pas en
  **nom**.
- Le danger n'est pas la largeur, c'est le **décalage silencieux** là où les **octets comptent** :
  interop, structures, fichiers binaires.
- **`Currency` → `Decimal`** (exact, décimal), **jamais `Double`**.
- Les **suffixes** (`%`/`&`/`@`) et les **fonctions `CInt`/`CLng`/`CCur`** ont aussi changé de
  cible : `CCur` est supprimée (→ `CDec`), et le nouveau **`CShort`** rend l'ancien comportement
  16 bits.
- Le **contrôle de débordement** existe toujours, mais à un **seuil déplacé**.

---

## 🔗 Pour aller plus loin

- **7.1 — `Variant` → `Object`** : pourquoi un `Short` *boxé* ne se déballe pas en `Integer`.
- **7.3 — Le type `Date`** : un autre faux-ami de taille/sémantique (`Double` → `DateTime`).
- **Module 16.4 / 16.5** — `Declare` → `P/Invoke` et *marshaling* : où la taille exacte est vitale.
- **Chapitre 12.7** — `Type` → `Structure` : sémantique de valeur et disposition mémoire.
- **Annexe B.2** (entiers redimensionnés) et **B.9** (arrondis monétaires) ; **Annexe D**
  (correspondance des types, plages et conversions).

➡️ Section suivante : **[7.3 — Le type `Date`](03-dates.md)**  
⬆️ Retour au **[sommaire du chapitre 7](README.md)**

⏭️ [`Date` (un `Double` en VB6 → `DateTime` en .NET) : conversions et pièges](/07-types-variables/03-dates.md)
