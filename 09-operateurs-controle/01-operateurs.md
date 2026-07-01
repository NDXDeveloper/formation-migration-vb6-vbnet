🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.1 Opérateurs

> **Module 9 — Opérateurs, chaînes et structures de contrôle** · [↩ Retour au chapitre](README.md)
> **Indicateurs de cette section** : ⚠️ plusieurs changements silencieux

À première vue, les opérateurs sont la partie la plus anodine de la migration : `+`, `-`, `*`, `=`, `<`, `>`… s'écrivent **exactement pareil** en VB6 et en VB.NET. C'est justement le danger. Sous cette ressemblance se cachent **trois catégories de surprises** :

1. des opérateurs qui **changent de sémantique** sans changer de syntaxe (`+` sur des chaînes, `Mod` sur des flottants) ;
2. des opérateurs **supprimés** (`Eqv`, `Imp`) qu'il faut réécrire ;
3. des opérateurs **nouveaux** (`AndAlso`, `OrElse`, `IsNot`, affectation composée) qui ne sont pas obligatoires, mais qui changent la façon idiomatique d'écrire — et qu'on ne peut **pas** appliquer mécaniquement (voir le piège `And` → `AndAlso`).

---

## Vue d'ensemble

| Catégorie | Opérateur(s) | Statut en VB.NET |
|-----------|--------------|------------------|
| Arithmétique | `+` `-` `*` `/` `^` | Conservés |
| Division entière | `\` | Conservé — ⚠️ tailles d'entiers, division par zéro |
| Reste | `Mod` | Conservé — ⚠️ modulo en virgule flottante possible |
| Concaténation | `&` | Conservé — **à privilégier** |
| Concaténation **ou** addition | `+` | Conservé — ⚠️ ambigu, **à auditer** |
| Logique / bit à bit | `And` `Or` `Not` `Xor` | Conservés — évaluent **les deux** opérandes |
| Logique **court-circuit** | `AndAlso` `OrElse` | ✨ Nouveaux |
| Logique (obsolètes) | `Eqv` `Imp` | ❌ **Supprimés** |
| Comparaison | `=` `<>` `<` `>` `<=` `>=` | Conservés — sur chaînes : ⚠️ `Option Compare` |
| Identité de référence | `Is` | Conservé |
| Identité (négation) | `IsNot` | ✨ Nouveau |
| Test de type | `TypeOf…Is` | Conservé |
| Correspondance de motif | `Like` | Conservé — ⚠️ `Option Compare` |
| Affectation composée | `+=` `-=` `*=` `/=` `\=` `^=` `&=` | ✨ Nouveaux |
| Décalage de bits | `<<` `>>` (et `<<=` `>>=`) | ✨ Nouveaux |

Le reste de cette section détaille les lignes marquées ⚠️ et les nouveautés utiles.

---

## `&` vs `+` : la concaténation (le piège du quotidien)

C'est l'opérateur le plus fréquent du code VB et la **première source de bugs silencieux** de ce chapitre.

En VB6, le sens de `+` **dépend du type des opérandes** : avec deux chaînes il concatène, mais dès qu'un nombre entre en jeu il peut additionner (après conversion implicite) — ou lever une erreur de type.

```vb
' VB6 — la sémantique de "+" dépend des types
Dim r As Variant
r = "10" + 5       ' r = 15      (numérique : addition)
r = "10" + "5"     ' r = "105"   (chaîne : concaténation)
r = "10" & 5       ' r = "105"   (concaténation : TOUJOURS, quels que soient les types)
r = "abc" + 5      ' ERREUR 13   (Type mismatch)
```

L'opérateur `&`, lui, **convertit toujours en chaîne et concatène**, sans ambiguïté.

En VB.NET sous `Option Strict On`, l'ambiguïté disparaît — mais d'une manière qui peut **casser la compilation** d'un code qui utilisait `+` pour concaténer :

```vb
' VB.NET avec Option Strict On
Dim r As String
r = "10" & 5       ' "105"  — OK, & convertit et concatène
r = "10" + "5"     ' "105"  — OK, deux chaînes
r = "10" + 5       ' ERREUR DE COMPILATION : pas de conversion implicite String <-> Integer
```

> ⚠️ **Règle de migration.** L'assistant de mise à niveau **ne convertit pas** `+` en `&` : il ne peut pas deviner si vous vouliez additionner ou concaténer. Tout `+` utilisé pour de la concaténation en VB6 est donc un **point à auditer**. La consigne est simple : **utiliser `&` pour toute concaténation**, et réserver `+` à l'arithmétique. Sous `Option Strict Off`, un `+` ambigu *compile* mais reproduit (et parfois aggrave) le comportement dépendant des types — d'où l'intérêt d'activer `Option Strict` (module **17.1**).

À ne pas confondre avec les **autres usages** du caractère `&` : suffixe de type (`100&` = littéral `Long`) et littéraux `&H` (hexadécimal) / `&O` (octal). Ils n'ont rien à voir avec l'opérateur de concaténation. La taille des entiers ayant changé (module **7.2**), surveillez plutôt ces suffixes lors de la migration des types.

---

## `\` : division entière

L'opérateur `\` réalise une **division entière** (le résultat est tronqué vers zéro). Si les opérandes sont flottants, ils sont d'abord **arrondis** à un entier — ce comportement est **identique** en VB6 et en VB.NET.

```vb
7 \ 2        ' 3
7.8 \ 2      ' 4   (7.8 est arrondi à 8, puis 8 \ 2)
```

Deux points de vigilance, hérités d'ailleurs :

- **Tailles d'entiers.** Le type du résultat suit les opérandes, et les entiers ont été **redimensionnés** (`Integer` VB6 = 16 bits → `Short` ; `Long` VB6 = 32 bits → `Integer`). Un calcul qui débordait silencieusement en VB6 peut lever une `OverflowException` en .NET, et inversement. Voir module **7.2**.
- **Division par zéro** : voir l'encadré ci-dessous.

> ⚠️ **Division par zéro : un comportement qui change.**
> En VB6, toute division par zéro déclenche l'**erreur 11** à l'exécution. En VB.NET :
> - `\` (entier) et la division `Decimal` lèvent une **`DivideByZeroException`** ;
> - mais `/` (flottant `Double`) **ne lève rien** : `1.0 / 0.0` renvoie `Double.PositiveInfinity`, et `0.0 / 0.0` renvoie `NaN` (norme IEEE 754).
>
> Conséquence : du code VB6 qui **comptait sur l'erreur 11** pour intercepter une division flottante par zéro (via `On Error`) ne la verra **plus jamais** se déclencher après migration — le calcul produira silencieusement `Infinity` ou `NaN`. À traquer explicitement (module **17.4**).

---

## `Mod` : le reste

`Mod` renvoie le **reste** d'une division. Le **signe du résultat suit le dividende** (l'opérande de gauche), à l'identique dans les deux langages :

```vb
-7 Mod 3     ' -1   (signe du dividende)
 7 Mod -3    '  1
```

La différence est ailleurs, et elle est **silencieuse** : en VB6, `Mod` **arrondit d'abord ses opérandes en entiers**. En VB.NET, `Mod` accepte les **types flottants** et calcule un **modulo en virgule flottante** — sans arrondir.

```vb
' VB6
5.3 Mod 2     ' 1     (5.3 arrondi à 5, puis 5 Mod 2)

' VB.NET
5.3 Mod 2     ' 1.3   (modulo flottant : 5.3 - 2*2 = 1.3, AUCUN arrondi)
```

> ⚠️ **Piège silencieux.** Si du code VB6 appliquait `Mod` à des `Variant`/`Single`/`Double` en s'appuyant sur l'arrondi implicite (calculs de cycles, de positions, de pagination…), le résultat **change** après migration. La correction est généralement d'**arrondir explicitement** les opérandes (`CInt`, `Math.Round`…) avant le `Mod`, pour retrouver la sémantique d'origine — ou d'assumer le modulo flottant en connaissance de cause.

---

## Opérateurs logiques et bit à bit : `And`, `Or`, `Not`, `Xor` — et le court-circuit

En VB6, `And` `Or` `Not` `Xor` `Eqv` `Imp` sont des opérateurs **bit à bit** qui servent **aussi** d'opérateurs logiques (parce qu'un booléen VB6 est un entier : `True` = `-1`, tous les bits à 1). Ils **évaluent toujours les deux opérandes**.

En VB.NET :

- `And` `Or` `Not` `Xor` sont **conservés** — toujours bit à bit sur les entiers, logiques sur les `Boolean`, et ils continuent d'**évaluer les deux opérandes** ;
- `Eqv` et `Imp` sont **supprimés** ;
- deux opérateurs **à court-circuit** apparaissent : `AndAlso` et `OrElse`, qui n'évaluent le second opérande **que si nécessaire**.

**Remplacer les opérateurs supprimés :**

| VB6 | Équivalent VB.NET |
|-----|-------------------|
| `a Eqv b` | `a = b` (booléens) ou `Not (a Xor b)` |
| `a Imp b` | `(Not a) Or b` |

**Le court-circuit, en pratique.** Comme VB6 n'a pas de court-circuit, les gardes d'objet y sont typiquement **imbriquées** :

```vb
' VB6 — pas de court-circuit : on imbrique les tests
If Not (obj Is Nothing) Then
    If obj.EstValide Then
        ' ...
    End If
End If
```

En VB.NET, la même garde s'écrit en **une seule condition** grâce à `AndAlso` (et `IsNot`, voir plus bas) :

```vb
' VB.NET — court-circuit explicite
If obj IsNot Nothing AndAlso obj.EstValide Then
    ' ...
End If
```

> ⚠️ **Ne remplacez JAMAIS `And` par `AndAlso` mécaniquement.** Le `And` de VB6 est peut-être utilisé **bit à bit**, pas logiquement :
>
> ```vb
> Dim octetBas As Integer = valeur And &HFF   ' masque binaire : And est CORRECT
> ```
>
> Ici, `AndAlso` serait une **erreur** (il exige des opérandes `Boolean` et fait un ET logique court-circuité). La règle : n'introduire `AndAlso`/`OrElse` que dans les **conditions logiques** (`If`, `While`, `Do…Loop While`), après avoir vérifié qu'il s'agit bien d'un test booléen — et jamais sur un usage de masquage de bits. L'assistant de migration, prudemment, conserve `And`/`Or` partout ; c'est à vous d'introduire le court-circuit **là où il a un sens**.

---

## `Like` : correspondance de motifs

L'opérateur `Like` est **conservé** en VB.NET, avec la **même syntaxe de motifs** qu'en VB6 :

| Motif | Signifie |
|-------|----------|
| `?` | un caractère quelconque |
| `*` | zéro caractère ou plus |
| `#` | un chiffre (`0`–`9`) |
| `[liste]` | un caractère de la liste (ex. `[A-Z]`) |
| `[!liste]` | un caractère **absent** de la liste |

```vb
"Fichier12.txt" Like "Fichier##.*"    ' True   (## = deux chiffres)
```

> ⚠️ **Dépendance à `Option Compare`.** La sensibilité à la casse de `Like` (et l'interprétation des plages comme `[A-Z]`) dépend de `Option Compare` : `Binary` = sensible à la casse, `Text` = insensible.
>
> ```vb
> "abc" Like "[A-Z]*"   ' Binary -> False  |  Text -> True
> ```
>
> En VB6, `Option Compare` se déclare **par module** (défaut `Binary`). En VB.NET, il se règle **au niveau du projet** ou par fichier, et un réglage non reporté change **silencieusement** le résultat. Vérifiez que la valeur d'origine est bien préservée (module **6.4**). Ce même piège vaut pour les **comparaisons de chaînes** ci-dessous.

Certaines équipes remplacent `Like` par des **expressions régulières** (`System.Text.RegularExpressions`) pour une logique de motifs plus puissante. Ce n'est **pas** requis par la migration, et attention : la syntaxe des motifs `Like` **n'est pas** celle des regex (`*` et `?` n'y ont pas le même sens). À réserver à du refactoring volontaire (module **17.5**).

---

## `Is` / `IsNot` (et `TypeOf…Is`)

`Is` compare deux **références d'objet** : il est vrai si les deux désignent **la même instance** (ou pour tester `obj Is Nothing`). Il est **conservé** tel quel.

VB.NET ajoute `IsNot`, qui remplace avantageusement le `Not (… Is …)` de VB6 :

```vb
' VB6
If Not (ref1 Is ref2) Then ...

' VB.NET — plus lisible
If ref1 IsNot ref2 Then ...
If obj Is Nothing Then ...          ' inchangé
```

`TypeOf obj Is NomDeType` (test de type) fonctionne dans les deux langages ; VB.NET accepte aussi `TypeOf obj IsNot NomDeType`.

> ⚠️ **`Is` reste une comparaison de *référence*.** Pour les **types valeur** / `Structure`, et compte tenu du nouveau sens de `Nothing` (valeur par défaut de **tout** type : `null` pour une référence, mais zéro/valeur par défaut pour un type valeur), un `If x Is Nothing` peut ne pas signifier ce que vous croyez. Ce point est traité en détail au module **7.7** (`Nothing`, `Empty`, `Null`, `Missing`).

---

## Comparaisons de chaînes et `Option Compare`

Les opérateurs `=` `<>` `<` `>` `<=` `>=` sont conservés. Sur les **nombres**, rien ne change. Sur les **chaînes**, comme pour `Like`, le résultat dépend de `Option Compare` :

```vb
"Été" = "ÉTÉ"     ' Binary -> False (sensible à la casse)
                  ' Text   -> True  (insensible à la casse)
```

> ⚠️ **Piège silencieux et coûteux.** Si l'application VB6 utilisait `Option Compare Text` (comparaisons insensibles à la casse, très courant pour des recherches ou des tris), et que le projet .NET retombe sur le défaut `Binary`, **toutes** les égalités et comparaisons de chaînes basculent en sensible à la casse — sans le moindre avertissement. Recherches qui ne trouvent plus, doublons qui réapparaissent, tris réordonnés. À vérifier dès le départ (module **6.4**) et à couvrir par le *golden master* (module **5.5**).

---

## Nouveautés VB.NET utiles (non obligatoires)

Ces opérateurs n'existent pas en VB6 ; le code hérité ne les utilise donc pas. Ils relèvent du **refactoring idiomatique** (module **17.5**), pas de la migration stricte.

**Affectation composée** — évite de répéter la variable :

```vb
total += montant        ' équivaut à : total = total + montant
message &= ligne        ' équivaut à : message = message & ligne
compteur -= 1
```

**Décalage de bits** `<<` et `>>` (et `<<=` / `>>=`) — là où VB6 imposait des multiplications/divisions ou des appels d'API.

---

## 🎯 Synthèse — points de vigilance de la section

| Opérateur | Risque | Réflexe de migration |
|-----------|--------|----------------------|
| `+` (sur chaînes) | Concaténation ambiguë / erreur de compilation | Remplacer par `&` ; auditer chaque `+` |
| `\` et `/` par zéro | `/` flottant ne lève **plus** d'erreur (`Infinity`/`NaN`) | Tester explicitement le diviseur |
| `Mod` (flottant) | Plus d'arrondi implicite → résultat différent | Arrondir les opérandes si besoin |
| `And` / `Or` | Pas de court-circuit (conservé) | Introduire `AndAlso`/`OrElse` **uniquement** en contexte logique |
| `Eqv` / `Imp` | Supprimés | `=` / `Not (a Xor b)` · `(Not a) Or b` |
| `Like`, comparaisons de chaînes | Casse selon `Option Compare` | Reporter fidèlement `Option Compare` |
| `Is` | Comparaison de référence ; `Nothing` redéfini | Voir module 7.7 ; moderniser en `IsNot` |

---

## 🔗 Pour aller plus loin

- Module **7.2** — Entiers redimensionnés (débordements de `\`, `Mod`)
- Module **7.7** — `Nothing`, `Empty`, `Null`, `Missing` (sémantique de `Is`)
- Module **6.4** / **17.1** — Options de compilation (`Option Strict`, `Option Compare`)
- Module **9.3** — Conditions : `IIf` → opérateur `If()` (la suite de l'histoire du court-circuit)
- Module **17.4** — Traquer les pièges silencieux après migration
- Annexe **A** — Tableau de correspondance VB6 → VB.NET (référence des opérateurs)
- Annexe **B** — Catalogue des pièges silencieux (`True = -1`, conversions booléennes, arrondis)

---

> **Section précédente** → [9. Introduction du chapitre](README.md)
> **Section suivante** → [9.2 — Fonctions de chaînes (`Microsoft.VisualBasic` vs `String`)](02-fonctions-chaines.md) ⚠️

⏭️ [Fonctions de chaînes (`Left`/`Right`/`Mid`, **instruction** `Mid`, `LSet`/`RSet`) : `Microsoft.VisualBasic` vs `String`](/09-operateurs-controle/02-fonctions-chaines.md)
