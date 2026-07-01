🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2.3 Le typage : `Variant`, entiers redimensionnés, `Currency` ⚠️

> **Chapitre 2 — Différences fondamentales VB6 ↔ VB.NET**
> Trois types qui changent de taille, de nom ou de nature — dont un piège qui se cache derrière un mot-clé **identique**.

---

## Pourquoi les types changent (et pourquoi c'est piégeux)

La section 2.1 l'a annoncé : en passant au CLR, on adopte un **système de types unifié** (le CTS),  
où tout dérive d'`Object` et où l'on distingue types **valeur** et types **référence**. Les types  
historiques de VB6 (`Variant`, `Currency`, et même les entiers) **n'y trouvent pas leur place  
telle quelle** : certains disparaissent, d'autres **changent de taille**.

Le danger n'est pas uniforme. Certains changements sont **visibles** — un mot-clé qui n'existe plus  
provoque une erreur de compilation, le compilateur vous force la main. Mais l'un d'eux est
**parfaitement silencieux** : il se cache derrière un mot-clé qui existe **dans les deux langages**,
avec la **même orthographe** et une **taille différente**. C'est le piège des **entiers  
redimensionnés**, et c'est le cœur de cette section.

> ⚠️ **La grille du chapitre, appliquée ici.** Posez-vous à chaque type : *« le mot-clé existe-t-il
> encore en VB.NET ? »*. S'il **disparaît** (`Variant`, `Currency`), le changement est **visible**
> (erreur de compilation). S'il **survit avec un autre sens** (`Integer`, `Long`), le changement est
> **silencieux** — et c'est là qu'on se fait piéger.

---

## 1️⃣ `Variant` → `Object` : la fin du type « fourre-tout »

### Ce qui change

En VB6, le **`Variant`** est le type universel : il peut contenir **n'importe quoi** (un nombre, une  
chaîne, un objet, une date…), et c'est même le type **par défaut** d'une variable non typée. En  
VB.NET, `Variant` **n'existe plus** : son rôle de « type pouvant tout contenir » revient à
**`Object`**, la racine de **tous** les types du CTS.

```vb
' VB6
Dim valeur As Variant       ' peut tout contenir
valeur = 42
valeur = "bonjour"
valeur = #1/15/2026#        ' littéral date VB : format US M/j/aaaa
```

```vb
' VB.NET — Variant n'existe plus → Object
Dim valeur As Object        ' peut tout contenir (racine du CTS)
valeur = 42
valeur = "bonjour"
valeur = #1/15/2026#
```

### Le côté visible… et le côté silencieux

- **Visible** : le mot-clé `Variant` est **invalide** en VB.NET. Le code ne compile pas tant qu'on ne l'a pas remplacé — l'erreur vous guide. Bonne nouvelle.
- **Silencieux** : ce qui change **sous** le mot-clé. Trois points de vigilance :
  - **Le boxing.** Stocker un type **valeur** (un `Integer`, par exemple) dans un `Object` l'« emballe » sur le tas (*boxing*) ; le relire le « déballe ». C'est transparent à l'écriture, mais cela a un **coût** et change la nature de la donnée manipulée.
  - **Les états perdus du `Variant`.** Le `Variant` VB6 connaissait des états spéciaux — **`Empty`**, **`Null`**, **`Missing`** — qu'`Object` **ne reproduit pas** à l'identique. Ce point, source de bugs à part entière, est traité en **section 7.7** et fiché en **annexe B.5**.
  - **Les coercitions implicites.** Le `Variant` convertissait automatiquement entre types selon le contexte. En VB.NET, ces conversions deviennent **explicites** sous `Option Strict On` ; sous `Option Strict Off`, le *late binding* en imite une partie, mais **pas à l'identique** (→ sections 16.3 et 17.1).

> 💡 **L'objectif à terme.** `Object` doit rester l'**exception**, pas la règle. Là où le `Variant`
> VB6 servait par **commodité** (ou par défaut), la migration est l'occasion de rétablir un
> **typage explicite** — bénéfice direct d'`Option Strict On` (chapitre 17.1).

---

## 2️⃣ Entiers redimensionnés : **le piège silencieux de cette section** ⭐

Voici le changement le plus dangereux du lot, précisément parce qu'il **ne déclenche aucune erreur**.

### Le fait central : mêmes noms, tailles différentes

Les mots-clés `Integer` et `Long` existent **dans les deux langages**. Mais ils **ne désignent pas  
la même taille** :

| Mot-clé | Taille en **VB6** | Taille en **VB.NET** |
|---|---|---|
| `Integer` | **16 bits** (−32 768 à 32 767) | **32 bits** (≈ ±2,1 milliards) |
| `Long` | **32 bits** (≈ ±2,1 milliards) | **64 bits** (≈ ±9,2 × 10¹⁸) |

Autrement dit, chaque type entier a « **glissé d'un cran** » vers le haut. Le modèle mental à  
graver :

```
   Largeur     VB6          VB.NET
   ───────     ───          ──────
   16 bits     Integer  ──► Short
   32 bits     Long     ──► Integer
   64 bits      (—)          Long
```

Pour **préserver la taille** d'origine, la correspondance de migration est donc un **décalage** :

| Type VB6 | Taille | Cible .NET **de même taille** |
|---|---|---|
| `Integer` (16 bits) | 16 bits | **`Short`** |
| `Long` (32 bits) | 32 bits | **`Integer`** |

### Pourquoi c'est silencieux

Parce que `Dim x As Integer` est du **VB6 valide** *et* du **VB.NET valide**. Le code **compile des  
deux côtés** — mais `x` fait **16 bits** dans un monde et **32 bits** dans l'autre.

```vb
' Ce code est valide en VB6 ET en VB.NET — mais "compteur" n'a pas la même taille
Dim compteur As Integer     ' VB6 : 16 bits  |  VB.NET : 32 bits
Dim total As Long           ' VB6 : 32 bits  |  VB.NET : 64 bits
```

Le compilateur ne dira **rien**. Le piège ne se révèle que **par le comportement**, dans quatre  
situations précises.

### Où le piège mord vraiment

**a) Les appels d'API Windows (`Declare`).** C'est le cas le plus fréquent et le plus traître. En
VB6, les fonctions de l'API Win32 (32 bits) se déclaraient avec **`Long`**. Or un `Long` VB6 (32  
bits) devient, à taille égale, un **`Integer`** VB.NET — **surtout pas** un `Long` VB.NET (qui ferait
64 bits et corromprait l'appel) :

```vb
' VB6 : paramètre 32 bits de l'API déclaré en Long
Declare Function GetTickCount Lib "kernel32" () As Long
```

```vb
' VB.NET : le 32 bits de l'API devient Integer — PAS Long !
Declare Function GetTickCount Lib "kernel32" () As Integer
' (Long ici = 64 bits = mauvaise taille → bug de marshaling)
```

Traduire mécaniquement `Long` → `Long` produit un code qui **compile** mais **marshale faux**. Sujet  
détaillé en **sections 16.4 et 16.5**.

**b) Les structures binaires et formats de fichiers.** Un enregistrement écrit en VB6 avec un
`Integer` occupait **2 octets** ; le « même » `Integer` en VB.NET en occupe **4**. Tout format de
fichier binaire, toute structure passée à une API, tout protocole à taille fixe **se décale** si  
l'on ne préserve pas les largeurs (→ `Short`).

**c) Les débordements (overflow).** Un calcul qui débordait — ou *ne* débordait *pas* — selon la
taille change de comportement :

```vb
' VB6 : Integer 16 bits → 40000 DÉBORDE (erreur d'exécution "Overflow")
' VB.NET : Integer 32 bits → 40000 passe sans broncher
Dim n As Integer
n = 40000
```

Un code qui s'appuyait (volontairement ou non) sur les bornes 16 bits **ne se comporte plus pareil**.

**d) Les opérations bit à bit et l'arithmétique sensible à la largeur.** Masques, décalages,
combinaisons de drapeaux : tout ce qui raisonne sur un **nombre de bits précis** doit être réexaminé.

> ⚠️ **L'embûche de l'assistant.** Les outils de migration mappent généralement `Integer` → `Short`
> et `Long` → `Integer` **pour préserver les tailles** — ce qui est le choix **sûr**. Conséquence :
> votre code converti peut se retrouver **truffé de `Short`**. C'est **correct** du point de vue du
> comportement, mais peu idiomatique. Le bon arbitrage (garder `Short` là où la taille **compte** —
> interop, binaire — et revenir à `Integer` dans la logique métier après vérification d'absence de
> débordement) relève du refactoring (chapitre 17). **Ne « nettoyez » jamais les `Short` à la légère
> sans vérifier le contexte.**

👉 **À creuser** : sections **7.2** (entiers redimensionnés en pratique) et **annexe D**
(correspondance des types, tailles et plages).

---

## 3️⃣ `Currency` → `Decimal` : le type monétaire change

### Ce qui change

Le **`Currency`** de VB6 était un type **64 bits à virgule fixe** (toujours **4 décimales**), conçu  
pour les **calculs monétaires** sans les erreurs d'arrondi du virgule flottante. En VB.NET, `Currency`
**n'existe plus** ; son remplaçant est **`Decimal`** (128 bits, à virgule flottante décimale, haute
précision).

```vb
' VB6
Dim prix As Currency        ' 64 bits, 4 décimales fixes
prix = 19.99
```

```vb
' VB.NET — Currency n'existe plus → Decimal
Dim prix As Decimal         ' 128 bits, précision décimale élevée
prix = 19.99D               ' le suffixe D force un littéral Decimal
```

### `Decimal` n'est pas `Currency` à l'identique

Le remplacement est **meilleur** (plus de précision, plus de plage), mais **pas équivalent au bit  
près** :

| Aspect | `Currency` (VB6) | `Decimal` (VB.NET) |
|---|---|---|
| **Taille** | 64 bits | **128 bits** |
| **Décimales** | **Fixe : 4** | **Variable : 0 à 28** |
| **Plage** | ± 922 337 203 685 477,5807 | ± ~7,9 × 10²⁸ |
| **Usage** | Monétaire | Monétaire / haute précision |

### Le côté visible… et le côté silencieux

- **Visible** : `Currency` est **invalide** en VB.NET → erreur de compilation, donc remplacement forcé. Le compilateur vous guide.
- **Silencieux** : le **comportement d'arrondi**. `Currency` arrondissait **systématiquement à 4 décimales** ; `Decimal` **conserve davantage de précision**. Un calcul qui donnait un résultat « propre » à 4 décimales en VB6 peut produire un résultat **plus long** en `Decimal` — ce qui décale ensuite les **comparaisons**, les **totaux** et les **affichages**. En migration financière, ces écarts au centime (ou en deçà) sont scrutés de près. *(→ fiché en annexe B.9, arrondis monétaires.)*

> ⚠️ **L'erreur à ne JAMAIS commettre.** Remplacer `Currency` par **`Single`** ou **`Double`**. Ce
> sont des types à **virgule flottante binaire**, **inadaptés à la monnaie** : `0.1 + 0.2`
> n'y vaut pas exactement `0.3`. La cible monétaire correcte est **`Decimal`**, et **uniquement**
> `Decimal`.

> 💡 **Et le `Decimal` de VB6 ?** En VB6, `Decimal` n'existait **pas** comme type déclarable à part
> entière : il n'était accessible **qu'à travers un `Variant`** (via `CDec`). En VB.NET, `Decimal`
> devient un **type de premier rang**, que l'on peut déclarer directement — une **amélioration**.

👉 **À creuser** : **annexe B.9** (arrondis monétaires) et **annexe D** (types et conversions).

---

## ↔️ Et les autres types ? (renvois)

Pour rester focalisée, cette section traite les **trois** types de son titre. Les autres écarts de  
typage, tout aussi importants, sont couverts ailleurs — autant savoir où :

| Type / sujet | Ce qui change | Où c'est traité |
|---|---|---|
| **`Date`** | `Double` (VB6) → `DateTime` (.NET) | Section **7.3**, annexe **B.9** |
| **Chaînes de longueur fixe** (`String * n`) | `VBFixedString` ou refactoring | Section **7.4** |
| **`Boolean`** (`True = -1`) | Conversions booléennes ↔ entier | Sections **7.6**, annexe **B.8** |
| **`Nothing` / `Empty` / `Null` / `Missing`** | États du `Variant` perdus | Section **7.7**, annexe **B.5** |
| **`Byte`** | 8 bits non signé — **inchangé** | Annexe **D** |

---

## 🗺️ Synthèse : visible ou silencieux ?

| Type | Devient | Détection | Nature du risque |
|---|---|---|---|
| **`Variant`** | `Object` | **Visible** (mot-clé invalide) | Boxing, états perdus (→ 7.7), coercitions |
| **`Integer`** (16 b) | `Short` (pour garder 16 b) | **🔴 Silencieux** (même mot-clé !) | API, binaire, débordements, bits |
| **`Long`** (32 b) | `Integer` (pour garder 32 b) | **🔴 Silencieux** (même mot-clé !) | API, binaire, débordements |
| **`Currency`** | `Decimal` | **Visible** (mot-clé invalide) | Arrondis monétaires (→ B.9) |

La leçon de ce tableau : les types qui **disparaissent** (`Variant`, `Currency`) sont **les moins  
risqués**, parce que le compilateur les signale. Le **vrai danger**, ce sont les entiers, dont le  
mot-clé **survit en changeant de taille** — silencieux par construction. C'est exactement le profil  
de risque que le chapitre 2 s'attache à rendre **visible**.

---

## ✅ À retenir

- Le passage au **CTS** (→ 2.1) redéfinit les types : certains **disparaissent**, d'autres **changent de taille**.
- **`Variant` → `Object`** : changement **visible** (mot-clé invalide), mais attention au **boxing**, aux **états perdus** (`Empty`/`Null`/`Missing` → 7.7) et aux **coercitions** désormais explicites.
- **Entiers redimensionnés — le piège silencieux de la section** : `Integer` passe de **16 à 32 bits**, `Long` de **32 à 64 bits**. Pour **préserver la taille** : VB6 `Integer` → **`Short`**, VB6 `Long` → **`Integer`**. Le code **compile** dans les deux langages — le danger est dans l'**API**, le **binaire**, les **débordements** et les **opérations bit à bit**.
- L'assistant préserve les tailles (`Integer`→`Short`, `Long`→`Integer`) : code **correct** mais **truffé de `Short`** ; ne les « nettoyez » pas sans **vérifier le contexte**.
- **`Currency` → `Decimal`** (jamais `Single`/`Double` !) : remplacement **visible**, mais **arrondis monétaires** à surveiller (→ B.9).
- **Où chercher le reste** : `Date` (7.3), chaînes fixes (7.4), `Boolean`/`True = -1` (7.6/B.8), états du `Variant` (7.7), et l'**annexe D** pour toutes les tailles et plages.

---

> 🧭 **Section suivante** → [2.4 ByRef par défaut → ByVal par défaut](04-byref-byval.md) ⭐ ⚠️  
> 🔝 **Retour** → [Sommaire du chapitre 2](README.md)

⏭️ [**ByRef par défaut → ByVal par défaut** (changements silencieux d'appel)](/02-differences-fondamentales/04-byref-byval.md)
