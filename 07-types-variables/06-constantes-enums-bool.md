🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.6 — Constantes, énumérations, et le booléen (`True` = -1 en VB6 ?) ⚠️

> **Chapitre 7 — Types, variables et déclarations** · Partie 3 — Migrer le langage
>
> Trois sujets réunis ici. Les **constantes** se migrent presque sans heurt. Les **énumérations**
> cachent un piège de **taille** (déjà rencontré en 7.2) et un changement de **qualification**.
> Quant au **booléen**, il pose une question célèbre — *« `True` vaut-il vraiment -1 ? »* — dont la
> réponse, en .NET, est **« ça dépend du chemin de conversion »**.

---

## 🔒 Constantes (`Const`)

La déclaration `Const` existe en VB.NET avec la même logique. Quelques points d'attention :

```vb
' VB6
Const MAX As Integer = 100
Public Const APP_NAME = "MonAppli"

' VB.NET — constante TYPÉE (recommandé sous Option Strict)
Const MAX As Integer = 100
Public Const APP_NAME As String = "MonAppli"
```

- **Typage.** Sous `Option Strict On`, donner un **type explicite** aux constantes (ne pas laisser
  un `Const` non typé devenir `Object`).
- **Constante mais calculée à l'exécution → `ReadOnly`, pas `Const`.** Un `Const` doit être une
  **expression constante à la compilation**. Pour une valeur figée mais évaluée au démarrage, .NET
  offre **`ReadOnly`** (initialisé une seule fois, à la déclaration ou dans le constructeur) — un
  outil que VB6 n'avait pas proprement :

  ```vb
  Public ReadOnly Demarrage As DateTime = DateTime.Now   ' figé, mais calculé à l'exécution
  ```

- **Constantes intrinsèques.** Beaucoup subsistent dans `Microsoft.VisualBasic` (`vbCrLf`, `vbTab`,
  `vbNullChar`…), donc le code compile. Les équivalents idiomatiques existent toutefois :

  | VB6 | Idiomatique .NET |
  |-----|------------------|
  | `vbCrLf` / `vbNewLine` | `Environment.NewLine` (ou `ControlChars.CrLf`) |
  | `""` | `String.Empty` (ou `""`) |
  | `vbNullString` | **`Nothing`** (pointeur de chaîne nul — **≠ `""`**, surtout en interop) |
  | `vbRed`, `vbYesNo`… | `Color.Red`, l'énum `MsgBoxStyle`/`MessageBoxButtons`… *(UI, voir 13.9)* |

> ⚠️ **`vbNullString` n'est pas `""`.** C'est une **chaîne nulle** (pointeur null), utilisée
> notamment pour des appels d'API qui distinguent « chaîne vide » et « pas de chaîne ». Sa
> traduction est **`Nothing`**, pas la chaîne vide.

---

## 🔢 Énumérations (`Enum`)

La syntaxe est quasi identique, mais trois choses changent.

### Le type de base : un piège de taille (rappel de 7.2)

En VB6, **une énumération est toujours un `Long` (32 bits)**. En VB.NET, la base par défaut est
**`Integer` (32 bits)** — donc **la même taille**, par chance :

```vb
' VB6 — toujours Long (32 bits) sous le capot
Enum Couleur
    Rouge      ' 0
    Vert       ' 1
    Bleu       ' 2
End Enum

' VB.NET — base Integer (32 bits) par défaut : ✅ même taille que le Long VB6
Enum Couleur
    Rouge
    Vert
    Bleu
End Enum
```

> 🪤 **Le piège :** vouloir « garder le `Long` » de VB6 en écrivant `Enum Couleur As Long` produit
> en VB.NET une énumération **64 bits** (puisque `Long` y fait 64 bits — voir **7.2**) ! Pour
> l'**interop** et la **persistance** (valeurs stockées, structures binaires), il faut **laisser la
> base par défaut `Integer`** (32 bits), pas `As Long`.

### L'accès devient **qualifié**

```vb
' VB6 : un membre d'enum Public était accessible SANS qualification
c = Vert

' VB.NET : accès QUALIFIÉ requis
Dim c As Couleur = Couleur.Vert
```

> 🪤 Du code VB6 qui utilisait des membres d'énumération **non qualifiés** (`Vert` au lieu de
> `Couleur.Vert`) ne compile plus tel quel. Les outils de migration **qualifient** automatiquement
> ces accès.

### Validation et champs de bits

Une énumération .NET **n'est pas exhaustivement contrôlée** : elle peut contenir une valeur **sans  
membre correspondant**. Pour valider, utiliser **`[Enum].IsDefined`**. Et pour les **champs de  
bits**, ajouter l'attribut **`<Flags>`** (que VB6 n'avait pas) :

```vb
Dim x As Couleur = CType(99, Couleur)                 ' compile et tourne (99 n'est pas un membre)
If [Enum].IsDefined(GetType(Couleur), x) Then ...     ' valider explicitement

<Flags>
Enum Permissions
    Lecture   = 1
    Ecriture  = 2
    Execution = 4
End Enum
' <Flags> améliore ToString() et documente l'usage en combinaisons binaires
```

> 💡 Les énumérations .NET gagnent aussi des méthodes utiles : `ToString()` (renvoie le **nom** du
> membre), `[Enum].Parse`, `[Enum].GetNames`. (`Enum` étant un mot réservé, on l'entoure de
> crochets : `[Enum].…`)

---

## ⚖️ Le booléen : `True` vaut-il -1 ?

### En VB6 : oui, `True` = -1

Le `Boolean` VB6 occupe **2 octets**. `True` vaut **-1** (tous les bits à 1, `&HFFFF`) et `False`  
vaut `0`. Les conversions sont symétriques :

```vb
' VB6
n = CInt(True)    ' -1
n = CInt(False)   ' 0
b = CBool(5)      ' True  (toute valeur non nulle → True)
b = CBool(0)      ' False
```

Cette valeur **-1** était parfois **exploitée** : stockage de booléens en `-1`/`0` dans une base ou  
un fichier, arithmétique (`True * 5` = -5), opérations binaires, ou tests `If (a > b) = -1`.

### En .NET : `CInt(True)` reste -1… mais pas partout

Le `Boolean` .NET occupe **1 octet**. Surtout, **la valeur numérique dépend du chemin de  
conversion** :

```vb
' VB.NET
Dim n As Integer
n = CInt(True)              ' -1  ✅ compatibilité VB conservée
n = Convert.ToInt32(True)   ' 1   ⚠️ chemin BCL .NET → 1, PAS -1 !
Dim b As Boolean = CBool(5) ' True (inchangé)
```

> 🚨 **C'est LE piège de cette section.** Les **fonctions de conversion VB** (`CInt`, `CType` vers
> un entier) renvoient **-1** pour `True` (compatibilité VB6). Mais le **reste du monde .NET** traite
> `True` comme **1** : `Convert.ToInt32(True)` = 1, `True.GetHashCode()` = 1, la **représentation
> binaire** sous-jacente vaut typiquement 1. **Mélanger les deux mécanismes** (ou lire la
> représentation brute) donne **-1 d'un côté, 1 de l'autre**.

### Interop : -1 (COM) vs 1 (Win32)

La même ambiguïté existe côté natif, et le ***marshaling*** doit être **explicite** :

```vb
' COM / Automation : VARIANT_BOOL — TRUE = -1 (comme VB6)
<MarshalAs(UnmanagedType.VariantBool)> Public Actif As Boolean
' Win32 BOOL (4 octets) — TRUE = 1
<MarshalAs(UnmanagedType.Bool)> Public Visible As Boolean
```

> ⚠️ Par défaut, un `Boolean` *marshalé* en `P/Invoke` est traité comme un **Win32 `BOOL` 4 octets
> (TRUE = 1)**. Si la cible attend un **`VARIANT_BOOL` (-1)** ou un **`bool` 1 octet**, il faut le
> préciser via `MarshalAs`. Détails au **module 16.5**.

### La règle d'or

> 🎯 **Ne jamais s'appuyer sur la valeur numérique d'un booléen.** L'utiliser **comme un booléen**.
> Là où `-1`/`0` compte **réellement** — données persistées, interop — rendre la conversion
> **explicite et maîtrisée** : `CInt` pour le -1 « façon VB », un mappage `-1`/`0` assumé, ou le bon
> `MarshalAs`. Et pour relire des données existantes en `-1`/`0`, `CBool(valeur)` reste fiable (tout
> non-nul → `True`).

---

## ⚠️ Pièges silencieux — récapitulatif

- **`vbNullString` ≠ `""`** : c'est `Nothing` (pointeur de chaîne nul), important en interop.
- **`Const` calculé à l'exécution** : utiliser **`ReadOnly`**, pas `Const`.
- **Énumération `As Long`** en VB.NET = **64 bits** (≠ le `Long` 32 bits de VB6) → garder la base
  **`Integer`** par défaut pour l'interop/la persistance.
- **Membres d'énumération non qualifiés** : ne compilent plus → `Type.Membre`.
- **Énumération non validée** : peut contenir une valeur hors membres → `[Enum].IsDefined`.
- **`True` = -1** : vrai via **`CInt`** (compat VB), mais **1** via `Convert.ToInt32`, la
  représentation binaire et l'interop Win32. **Ne pas dépendre de la valeur numérique.**

---

## 🧭 Stratégie de migration

1. **Typer** explicitement les constantes ; basculer en **`ReadOnly`** celles qui sont calculées à
   l'exécution.
2. **Vérifier la base des énumérations** : conserver `Integer` (32 bits) par défaut ; n'utiliser une
   autre base que par choix conscient (jamais `As Long` « par réflexe VB6 »).
3. **Qualifier** les accès aux membres d'énumération ; ajouter **`<Flags>`** aux énumérations de
   bits ; valider avec **`[Enum].IsDefined`** là où des valeurs externes entrent.
4. **Auditer toute dépendance à `True = -1`** : stockage `-1`/`0`, arithmétique/bits sur booléens,
   tests `= -1`. Rendre les conversions **explicites** ; choisir le bon `MarshalAs` en interop.
5. **Filet de sécurité** : le ***golden master*** (chapitre 5.5) révèle un booléen persisté à `1` au
   lieu de `-1`, ou une énumération dont la taille a changé.

---

## ✅ À retenir

- **Constantes** : quasi inchangées ; penser **`ReadOnly`** pour le « constant mais à l'exécution »
  et `Nothing` pour **`vbNullString`**.
- **Énumérations** : base **`Integer`** par défaut (≈ le `Long` 32 bits de VB6) — **`As Long` = 64
  bits, piège** ; accès **qualifié** ; valeurs **non validées** automatiquement ; `<Flags>` pour les
  bits.
- **Booléen** : `True` vaut **-1 via `CInt`** (compat VB6) mais **1** via le BCL/l'interop/la
  représentation binaire. **Règle d'or : ne pas se fier à la valeur numérique d'un booléen.**

---

## 🔗 Pour aller plus loin

- **7.2 — Entiers redimensionnés** : pourquoi `As Long` fait 64 bits, et la persistance des valeurs.
- **7.7 — `Nothing`, `Empty`, `Null`** : `vbNullString` → `Nothing`, et l'absence de valeur.
- **Chapitre 9** — opérateurs et conversions (comparaisons renvoyant des booléens).
- **Module 16.5** — *marshaling* : `VariantBool` vs `Bool`, ANSI/Unicode, `StructLayout`.
- **Chapitre 13.9** — `MsgBox`/constantes d'interface (`vbYesNo`, `vbInformation`…).
- **Annexe B.8** (`True = -1` et conversions booléennes) ; **Annexe D** (correspondance des types).

➡️ Section suivante : **[7.7 — `Nothing`, `Empty`, `Null`, `Missing`](07-nothing-empty-null.md)**  
⬆️ Retour au **[sommaire du chapitre 7](README.md)**

⏭️ [`Nothing`, `Empty`, `Null`, `Missing` : la fin du `Variant` et de ses états](/07-types-variables/07-nothing-empty-null.md)
