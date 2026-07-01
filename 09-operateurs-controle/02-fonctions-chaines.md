🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.2 Fonctions de chaînes : `Microsoft.VisualBasic` vs `String`

> **Module 9 — Opérateurs, chaînes et structures de contrôle** · [↩ Retour au chapitre](README.md)
> **Indicateur de cette section** : ⚠️ changements silencieux fréquents

Le code VB6 manipule les chaînes avec des fonctions comme `Left`, `Right`, `Mid`, `Len`, `InStr`. **Toutes existent encore** en VB.NET. On pourrait donc croire qu'il n'y a rien à faire — et c'est précisément le piège de cette section.

En VB.NET, ces fonctions **cohabitent** désormais avec les **méthodes de l'objet `String` du .NET**, et les **deux familles ne se comportent pas pareil**. Deux différences de comportement traversent toute la section :

1. **L'indexation** : les fonctions VB sont **1-based** (le 1ᵉʳ caractère est en position 1) ; les méthodes `String` sont **0-based**.
2. **Le dépassement de bornes** : les fonctions VB sont **tolérantes** (elles renvoient une chaîne vide ou tronquée) ; les méthodes `String` sont **strictes** (elles **lèvent une exception**).

---

## Deux familles, une seule à choisir par appel

| Famille | Exemples | Indexation | Hors limites |
|---------|----------|------------|--------------|
| `Microsoft.VisualBasic` (classe `Strings`) | `Mid`, `Left`, `Right`, `Len`, `InStr`, `Replace`… | **1-based** | Tolérant (chaîne vide / tronquée) |
| `System.String` (méthodes) | `.Substring`, `.IndexOf`, `.Length`, `.Replace`… | **0-based** | Strict (**exceptions**) |

> ⚠️ **Ne confondez pas deux espaces de noms.** Les fonctions `Left`/`Right`/`Mid`/`InStr`… appartiennent à `Microsoft.VisualBasic`, qui est un composant **pleinement supporté et légitime** de VB.NET : les conserver **n'est pas une dette technique**. À ne pas confondre avec l'espace `Microsoft.VisualBasic.Compatibility`, qui est, lui, une **béquille de l'assistant de migration à supprimer** (modules **4.2** et **17.2**). Garder `Mid` est idiomatique ; garder `Compatibility.VB6.*`, non.

---

## L'indexation : 1-based vs 0-based

C'est le décalage à garder en tête en permanence :

```
Chaîne  :   A   B   C   D   E   F
position :   1   2   3   4   5   6     <- Mid, InStr, Len  (Microsoft.VisualBasic)
indice   :   0   1   2   3   4   5     <- Substring, IndexOf, Length  (System.String)
```

Toute traduction d'une fonction VB vers une méthode `String` qui manipule une **position** implique de **soustraire 1**. C'est la première cause de bugs « d'un caractère » de la migration.

---

## `Left` / `Right` (fonctions)

`Left(s, n)` renvoie les `n` premiers caractères, `Right(s, n)` les `n` derniers. Les deux familles savent le faire :

```vb
Left("Bonjour", 3)                          ' "Bon"   (Microsoft.VisualBasic)
"Bonjour".Substring(0, 3)                   ' "Bon"

Right("Bonjour", 3)                         ' "our"
"Bonjour".Substring("Bonjour".Length - 3)   ' "our"   (indice à calculer à la main)
```

Deux pièges :

**Bornes — tolérant vs strict.**

```vb
Dim s As String = "AB"
Left(s, 10)            ' "AB"   -> tronque silencieusement, AUCUNE erreur
s.Substring(0, 10)     ' ArgumentOutOfRangeException !
```

Migrer `Left(s, 10)` en `s.Substring(0, 10)` **sans garde** transforme un comportement tolérant en **plantage** dès que la chaîne est plus courte que prévu. L'équivalent sûr :

```vb
s.Substring(0, Math.Min(10, s.Length))   ' "AB" — équivalent robuste de Left(s, 10)
```

**Collision de noms avec les propriétés `Left`/`Right`.** Dans une classe de formulaire ou de contrôle, `Left` et `Right` sont **aussi des propriétés de position**. Le compilateur peut interpréter `Left(...)` autrement que la fonction de chaîne. Levez l'ambiguïté en qualifiant explicitement :

```vb
Dim p As String = Strings.Left(nom, 3)    ' explicite : la fonction Microsoft.VisualBasic
```

*(Note : les variantes `Left$`/`Right$`/`Mid$` de VB6 — renvoyant une `String` plutôt qu'un `Variant` — existent encore, mais la distinction `Variant`/`Null` disparaît avec le `Variant`. Un `Nothing` passé à `Left` est traité comme `""`, là où `.Substring` sur `Nothing` lèverait une `NullReferenceException`.)*

---

## `Mid` (fonction)

`Mid(s, début[, longueur])` extrait une sous-chaîne. Son argument **`début` est 1-based** :

```vb
Dim s As String = "ABCDEF"

Mid(s, 3, 2)         ' "CD"   (à partir de la position 3 — Microsoft.VisualBasic)
s.Substring(2, 2)    ' "CD"   (indice 2 = position 3) -> le début est décalé de 1
```

Même piège de bornes que `Left`/`Right` : `Mid` est tolérant, `Substring` est strict.

```vb
Mid(s, 8)            ' ""                              (début au-delà de la fin : toléré)
s.Substring(7)       ' ArgumentOutOfRangeException !   (indice 7 hors limites : strict)
```

---

## ⚠️ L'**instruction** `Mid` (le faux-jumeau à repérer)

Voici le point le plus délicat de la section. En VB6, `Mid` n'est pas qu'une fonction : c'est aussi une **instruction d'affectation** qui **remplace des caractères *en place*** dans une variable chaîne. On la distingue de la fonction au seul fait qu'elle se trouve **à gauche du `=`** :

```vb
Dim s As String
s = "ABCDEF"
Mid(s, 2, 3) = "xyz"      ' s devient "AxyzEF"
                          '  -> "A" + "xyz" + "EF" : la LONGUEUR ne change pas
```

Le remplacement est **borné par `longueur`** (et par la taille de `s`) : un texte de remplacement trop long est **tronqué**, la longueur de `s` reste constante.

```vb
s = "ABCDEF"
Mid(s, 2, 3) = "xyzABC"   ' s devient "AxyzEF" : seuls 3 caractères sont écrits
```

**Pourquoi c'est un piège :**

- **Facile à manquer.** L'instruction `Mid` ressemble *trait pour trait* à un appel de la fonction `Mid` ; seule sa position à gauche de `=` les distingue. Un relecteur — comme un outil de migration ou une IA — peut la confondre. À repérer systématiquement.
- **Immuabilité .NET.** L'instruction `Mid` **existe encore** en VB.NET et fonctionne, mais les chaînes .NET sont **immuables** : la modification « en place » est une **illusion**. En coulisses, une **nouvelle chaîne est allouée** à chaque appel. Dans une boucle, c'est un piège de performance.
- **Réflexe idiomatique** pour des modifications répétées : `StringBuilder` (mutable, indices **0-based**), puis `.ToString()` à la fin.

```vb
Dim sb As New StringBuilder("ABCDEF")
sb(1) = "x"c                ' indices 0-based : position 2
sb(2) = "y"c
sb(3) = "z"c
Dim r As String = sb.ToString()   ' "AxyzEF"
```

> 🔗 L'instruction `Mid` accompagne souvent les **chaînes de longueur fixe** (`Dim s As String * n`) : on réservait une chaîne fixe et on en écrasait des tranches. Migrer l'un sans l'autre crée des bugs. Voir module **7.4**.

---

## `LSet` / `RSet`

En VB6, ces instructions **justifient** une chaîne dans une largeur donnée (typiquement une chaîne de longueur fixe), en complétant avec des espaces :

- `RSet` : justifie **à droite** (espaces à gauche) ;
- `LSet` : justifie **à gauche** (espaces à droite).

Le réflexe de migration est de passer aux méthodes `String` `PadLeft` / `PadRight` :

```vb
' VB6 — sur une chaîne de longueur fixe 10
Dim champ As String * 10
RSet champ = "42"          ' "        42"   (justifié à droite)
LSet champ = "42"          ' "42        "   (justifié à gauche)

' VB.NET — équivalents idiomatiques
champ = "42".PadLeft(10)   ' "        42"   (≈ RSet)
champ = "42".PadRight(10)  ' "42        "   (≈ LSet)
```

> ⚠️ **Une nuance de troncature.** `PadLeft`/`PadRight` **ne tronquent jamais** : si la source dépasse la largeur, elle est renvoyée intacte. `LSet`/`RSet`, dans une chaîne de **longueur fixe**, **tronquent** pour tenir dans la taille. Si votre source peut dépasser la largeur, combinez avec `Substring` pour reproduire la troncature d'origine.

> 🔥 **Le piège caché de `LSet`.** En VB6, `LSet` a un **second usage**, obscur et dangereux : **copier une variable d'un type défini par l'utilisateur (`Type`) vers un autre** par recopie mémoire brute. Cet usage n'a **aucun équivalent sûr** en .NET (il dépendait de l'alignement mémoire). Il doit être **réécrit explicitement**, champ par champ — ou via le *marshaling* si une réinterprétation binaire est réellement nécessaire (module **16.5**).

---

## Tableau de correspondance : fonction VB → méthode `String`

| Fonction VB (`Microsoft.VisualBasic`) | Méthode `String` équivalente | Différences clés |
|---------------------------------------|------------------------------|------------------|
| `Mid(s, d, l)` | `s.Substring(d - 1, l)` | début **1-based → 0-based** |
| `Left(s, n)` | `s.Substring(0, n)` | `Substring` **lève** si `n > Len(s)` |
| `Right(s, n)` | `s.Substring(s.Length - n)` | idem + calcul d'indice manuel |
| `Len(s)` | `s.Length` | fonction → **propriété** ; `Len` sur un `Type`/byte diffère (voir 7.4) |
| `InStr(s, t)` | `s.IndexOf(t)` | **1-based → 0-based** *et* **0 (absent) → -1** |
| `InStrRev(s, t)` | `s.LastIndexOf(t)` | idem |
| `Replace(s, a, b)` | `s.Replace(a, b)` | proche (sans arguments de position) |
| `UCase(s)` / `LCase(s)` | `s.ToUpper()` / `s.ToLower()` | ⚠️ sensibilité à la **culture** (annexe B.10) |
| `Trim` / `LTrim` / `RTrim` | `.Trim()` / `.TrimStart()` / `.TrimEnd()` | proche |
| `Split` / `Join` | `s.Split(...)` / `String.Join(...)` | proche ; tableau résultat **0-based** |

> ⚠️ **Le double piège d'`InStr` → `IndexOf`.** `InStr` est **1-based** et renvoie **0** quand le motif est absent ; `IndexOf` est **0-based** et renvoie **-1** quand il est absent. Le test de présence change donc :
>
> ```vb
> If InStr(s, "@") > 0 Then ...        ' VB6 : présence
> If s.IndexOf("@"c) >= 0 Then ...     ' .NET : présence  (>= 0, et SURTOUT pas > 0 !)
> ```
>
> Traduire `InStr(...) > 0` en `IndexOf(...) > 0` **manque toute occurrence en première position** (indice 0). Il faut `>= 0`.

---

## 🔑 Garder `Microsoft.VisualBasic` ou passer aux méthodes `String` ?

C'est l'arbitrage stratégique de la section. Les deux options sont valables ; elles ne s'opposent pas sur le « propre » mais sur le **moment** et le **risque**.

| Critère | Garder `Microsoft.VisualBasic` | Passer aux méthodes `String` |
|---------|-------------------------------|------------------------------|
| Effort de migration | **Faible** (code quasi inchangé) | Élevé (réécriture + indices) |
| Indexation | 1-based (cohérent avec l'existant) | 0-based (à recalculer) |
| Dépassement de bornes | Tolérant (chaîne vide) | Strict (**exceptions**) |
| `Nothing` | Toléré (≈ `""`) | `NullReferenceException` |
| Idiomatique .NET / portable C# | Non | **Oui** |
| Risque de régression | Faible | Plus élevé |

**Recommandation, fidèle à la philosophie « pont » de la formation :**

1. **Pendant la migration**, **conservez** les fonctions `Microsoft.VisualBasic`. Elles préservent le comportement (indices, bornes tolérantes) et **réduisent le nombre de changements simultanés** — exactement l'objectif de la cible 4.7.2.
2. **Plus tard, pendant la fiabilisation/refactoring** (module **17.5**), migrez **sélectivement** vers les méthodes `String`, fonction par fonction, **sous la protection des tests de non-régression** (*golden master*, module **5.5**) qui capturent le passage 1-based → 0-based et le changement de bornes.
3. **Ne faites jamais les deux en même temps** (migrer *et* réécrire l'API de chaînes) : vous perdriez la capacité de distinguer une régression de migration d'un bug de refactoring.

---

## 🎯 Synthèse — points de vigilance de la section

| Élément | Risque | Réflexe |
|---------|--------|---------|
| `Mid`/`Left`/`Right` → `Substring` | Décalage 1-based → 0-based | Soustraire 1 au début |
| `Substring`, `IndexOf`… | **Exceptions** là où VB renvoyait une chaîne vide | Garder VB.* ou ajouter `Math.Min`/gardes |
| **Instruction** `Mid` (à gauche de `=`) | Confondue avec la fonction ; réalloue (immuabilité) | Repérer ; `StringBuilder` si répété |
| `LSet`/`RSet` | Justification → `PadLeft`/`PadRight` (sans troncature) | Pad* + `Substring` si dépassement |
| `LSet` (copie de `Type`) | Aucun équivalent sûr | Réécrire champ par champ / marshaling |
| `InStr > 0` → `IndexOf > 0` | Rate une occurrence en position 0 | Utiliser `>= 0` |
| `Left`/`Right` dans un formulaire | Collision avec les propriétés de position | Qualifier : `Strings.Left(...)` |

---

## 🔗 Pour aller plus loin

- Module **7.4** — Chaînes de longueur fixe (`String * n`) : souvent liées à l'instruction `Mid` et à `LSet`/`RSet`
- Module **4.2** / **17.2** — L'espace `Microsoft.VisualBasic.Compatibility` (à ne pas confondre avec les fonctions de chaînes)
- Module **16.5** — Marshaling et structures (pour la réécriture de l'usage `Type` de `LSet`)
- Module **17.5** — Refactoring idiomatique (migration sélective vers les méthodes `String`)
- Module **5.5** — Harnais de tests de référence (*golden master*) pour sécuriser le passage 1-based → 0-based
- Annexe **A** — Tableau de correspondance VB6 → VB.NET (fonctions intrinsèques)
- Annexe **B.10** — Conversions sensibles à la culture (`UCase`/`LCase`, `Format`)

---

> **Section précédente** → [9.1 — Opérateurs](01-operateurs.md)
> **Section suivante** → [9.3 — `If…Then`, `Select Case`, `IIf` → opérateur `If()`](03-conditions.md) ⚠️

⏭️ [`If…Then`, `Select Case`, `IIf` → opérateur `If()` ternaire (et l'évaluation court-circuit)](/09-operateurs-controle/03-conditions.md)
