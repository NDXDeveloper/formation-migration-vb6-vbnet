🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.1 — `Sub` / `Function` ; appels avec / sans parenthèses, `Call`

> **La brique de base de tout le chapitre.** La *déclaration* d'une procédure change à peine entre VB6 et VB.NET — mais la *syntaxe d'appel*, elle, change pour de bon, et VB6 cachait un piège dans les parenthèses.

📍 *Module 10 · § 10.1 · [↑ Introduction du chapitre](README.md)*

---

## ⚡ En bref

Trois choses à retenir de cette section :

1. **La déclaration `Sub`/`Function` est l'une des constructions les plus stables** de la migration : la structure `Sub … End Sub` / `Function … As T … End Function` est quasi identique.
2. **En VB.NET, les parenthèses sont *systématiques*** à l'appel — pour un `Sub` comme pour une `Function`, avec ou sans arguments. VB6, lui, suivait des règles à géométrie variable (parfois sans parenthèses, parfois avec `Call`).
3. **`Call` devient obsolète** (mais reste toléré), et surtout : sur un seul argument, **le sens des parenthèses bascule** entre VB6 et VB.NET — une source de régression traitée en profondeur au **[§ 10.2](02-byref-byval-audit.md)**.

---

## 1. Déclarer un `Sub` et une `Function`

Bonne nouvelle pour commencer : c'est l'un des points les plus **stables** de toute la migration. La forme générale est identique.

```vb
' VB6
Sub Enregistrer(ByVal nom As String)
    ' ... corps de la procédure ...
End Sub

Function Somme(a As Integer, b As Integer) As Integer
    Somme = a + b          ' affectation au nom de la fonction
End Function
```

```vb
' VB.NET
Sub Enregistrer(ByVal nom As String)
    ' ... corps de la procédure ...
End Sub

Function Somme(a As Integer, b As Integer) As Integer
    Return a + b           ' Return (recommandé) — ou Somme = a + b
End Function
```

Ce qui **ne change pas** :

- les mots-clés `Sub` / `Function`, les blocs `End Sub` / `End Function` ;
- la déclaration du **type de retour** (`… As Integer`) après la liste de paramètres ;
- la possibilité d'affecter la valeur de retour au **nom de la fonction** (`Somme = …`), conservée en VB.NET.

Ce qui **mérite votre attention** (et qui est traité ailleurs) :

- ⚠️ le **mode de passage par défaut** des paramètres bascule de `ByRef` à `ByVal` → **[§ 10.2](02-byref-byval-audit.md)** ;
- le type `Integer` ne désigne **pas la même taille** des deux côtés → **[module 7](../07-types-variables/README.md)** ;
- `Return` vs affectation au nom de la fonction : ce ne sont **pas** exactement la même chose → **[§ 10.5](05-surcharge-retour.md)** ;
- les modificateurs d'accès (`Public`/`Private`/`Friend`) et la **surcharge** → **[module 12](../12-poo/README.md)** et **[§ 10.5](05-surcharge-retour.md)**.

> 💡 Le sujet de **cette** section, c'est uniquement la **syntaxe d'appel** : avec ou sans parenthèses, avec ou sans `Call`.

---

## 2. Appeler une procédure : la règle change vraiment

C'est ici que VB6 et VB.NET divergent nettement.

### 2.1 Côté VB6 — une syntaxe à géométrie variable

En VB6, la forme d'un appel dépend de **trois facteurs** : `Sub` ou `Function`, valeur de retour utilisée ou non, et présence du mot-clé `Call`.

**Appel comme *instruction*** (on ignore un éventuel retour) :

| Forme écrite | Validité VB6 |
|---|---|
| `Proc arg1, arg2` *(sans parenthèses)* | ✅ Forme canonique d'un `Sub` |
| `Call Proc(arg1, arg2)` | ✅ Avec `Call`, les parenthèses sont **obligatoires** |
| `Proc(arg1, arg2)` *(parenthèses, sans `Call`, plusieurs arguments)* | ❌ **Erreur de syntaxe** |
| `Proc(arg)` *(parenthèses, sans `Call`, **un seul** argument)* | ⚠️ Légal, mais `(arg)` est une **expression groupée** → passage **ByVal** forcé (voir § 4) |

**Appel comme *expression*** (on utilise la valeur de retour d'une `Function`) :

```vb
Dim total As Integer
total = Somme(10, 20)      ' parenthèses requises ; passage normal
```

Autrement dit, en VB6, **un même `Sub` peut s'appeler de plusieurs façons** selon le contexte — ce qui explique la diversité du code hérité que vous allez rencontrer.

### 2.2 Côté VB.NET — parenthèses systématiques

VB.NET **unifie** la règle : **les parenthèses entourent toujours la liste d'arguments**, quels que soient le type de procédure, le nombre d'arguments ou l'usage du retour.

```vb
Enregistrer("Dupont")          ' Sub : parenthèses requises
Rafraichir()                   ' Sub sans argument : parenthèses requises aussi
total = Somme(10, 20)          ' Function, retour utilisé
Somme(10, 20)                  ' Function, retour ignoré : autorisé en VB.NET
```

Points à noter :

- même un appel **sans argument** porte des parenthèses : `Rafraichir()`. L'éditeur de Visual Studio les **ajoute automatiquement** si vous les oubliez.
- une **`Function` peut être appelée comme une instruction**, sa valeur de retour étant simplement **abandonnée** (`Somme(10, 20)` sur une ligne) — ce que VB6 n'autorisait qu'avec une forme sans parenthèses ou avec `Call`.

> ✅ **Règle de survie VB.NET : toujours des parenthèses.** Si vous migrez à la main un appel VB6 sans parenthèses (`Proc arg1, arg2`), ajoutez-les (`Proc(arg1, arg2)`).

---

## 3. Le mot-clé `Call`

En VB6, `Call` servait à donner aux appels de `Sub` une allure plus « classique » et imposait l'usage des parenthèses :

```vb
Call Enregistrer(nom)          ' VB6 : Call exige les parenthèses
```

En **VB.NET**, `Call` **existe toujours et reste valide**, mais il est **redondant** : puisque les parenthèses sont désormais obligatoires partout, `Call` n'apporte plus rien.

```vb
Call Enregistrer(nom)          ' VB.NET : fonctionne, mais inutile
Enregistrer(nom)               ' VB.NET : forme idiomatique recommandée
```

> 🛠️ **Recommandation** : supprimez les `Call` lors du refactoring idiomatique (le code migré automatiquement peut en conserver). Ce n'est **pas** une correction *de comportement* — `Call x()` et `x()` sont strictement équivalents en VB.NET — mais une simplification de lisibilité.

---

## 4. ⚠️ Le piège des parenthèses sur un seul argument

Voici le point le plus subtil de cette section, et le lien direct avec le titre « appels **avec/sans parenthèses** ». Sur un **seul argument**, **le sens d'une paire de parenthèses n'est pas le même** en VB6 et en VB.NET.

En **VB6**, dans un appel-instruction, `Proc(x)` ne veut **pas** dire « appeler `Proc` avec l'argument `x` » : la paire de parenthèses est interprétée comme une **expression entre parenthèses**. Or une expression entre parenthèses est **évaluée** puis passée — ce qui revient à un passage **ByVal**, même si le paramètre est déclaré `ByRef`.

En **VB.NET**, `Proc(x)` est tout simplement la **syntaxe d'appel normale** : `x` est passé selon le mode déclaré (`ByVal` ou `ByRef`). Pour **forcer** un passage ByVal en VB.NET, il faut une **paire de parenthèses supplémentaire** : `Proc((x))`.

| Écriture *(appel-instruction)* | Interprétation **VB6** | Interprétation **VB.NET** |
|---|---|---|
| `Proc x` | Appel normal — respecte `ByRef`/`ByVal` déclaré | *(non valide : parenthèses requises)* |
| `Proc(x)` | `(x)` = **expression groupée** → **ByVal forcé** | **Liste d'arguments** → respecte le mode déclaré |
| `Proc((x))` | Double expression groupée → **ByVal forcé** | Parenthèses **supplémentaires** → **ByVal forcé** |

Conséquence pour la migration : un appel VB6 écrit `Proc (x)` (souvent avec une espace, parfois par mégarde) **forçait un ByVal** ; transcrit naïvement en `Proc(x)` côté VB.NET, il **rétablit le mode déclaré** — donc, potentiellement, un `ByRef` qui n'était pas voulu (ou l'inverse).

> ⚠️ **Ce piège est indissociable du basculement `ByRef` → `ByVal`.** Il est analysé, avec sa méthode d'audit, au **[§ 10.2 — ByRef par défaut → ByVal par défaut](02-byref-byval-audit.md)**. Retenez ici simplement que **les parenthèses ne sont pas « neutres »** : elles peuvent changer le mode de passage.

---

## 5. En pratique : migrer les appels

### Ce que font les outils

- L'**[*Upgrade Wizard*](../04-outils-migration/01-upgrade-wizard.md)** de Visual Studio **ajoute les parenthèses** manquantes et adapte la forme des appels ; il s'efforce aussi de **préserver le mode de passage** d'origine (voir § 10.2). C'est l'aspect le plus fiable de la conversion automatique.
- Les **[outils commerciaux](../04-outils-migration/03-outils-commerciaux.md)** (VBUC, VB Migration Partner) font de même, souvent avec une meilleure prise en charge des cas limites.

### En migration manuelle

Une poignée de gestes suffisent pour cette section :

| Code VB6 | Code VB.NET |
|---|---|
| `Proc arg1, arg2` | `Proc(arg1, arg2)` |
| `Call Proc(arg1, arg2)` | `Proc(arg1, arg2)` *(retirer `Call`)* |
| `Func arg` *(retour ignoré)* | `Func(arg)` |
| `x = Func(arg)` | `x = Func(arg)` *(inchangé)* |
| `Proc (x)` *(parenthèses « groupantes »)* | ⚠️ **à examiner** : le ByVal était-il voulu ? → § 10.2 |

> 💡 La transformation purement syntaxique (ajout de parenthèses, retrait de `Call`) est **mécanique et sûre**. Le **seul** point qui demande du jugement est le cas des parenthèses « groupantes » sur un argument — précisément parce qu'il touche au **mode de passage**, pas à la syntaxe d'appel.

---

## 📋 Tableau de synthèse

| Aspect | VB6 | VB.NET (4.7.2) | Nature du changement |
|---|---|---|---|
| Déclaration `Sub`/`Function` | `Sub … End Sub`, `Function … As T` | Identique | ✅ Stable |
| Parenthèses à l'appel | Variables (parfois absentes) | **Toujours présentes** | Mécanique |
| Appel sans argument | `Proc` | `Proc()` | Mécanique |
| `Function` appelée en instruction | Forme sans parenthèses ou `Call` | `Func(args)` *(retour abandonné)* | Mécanique |
| Mot-clé `Call` | Utile (impose les parenthèses) | **Obsolète** (toléré) | Cosmétique |
| `Proc(x)` sur **un** argument (instruction) | **Expression groupée** → ByVal forcé | **Liste d'arguments** → mode déclaré | ⚠️ Sémantique (→ § 10.2) |

---

## ✅ À retenir

1. **La déclaration ne bouge presque pas** : `Sub`/`Function` est l'un des terrains les plus sûrs de la migration.
2. **À l'appel, mettez toujours des parenthèses** en VB.NET — c'est la règle unique qui remplace les conventions à géométrie variable de VB6.
3. **Retirez les `Call`** pour un code idiomatique : ils ne changent rien au comportement.
4. **Méfiez-vous des parenthèses sur un seul argument** : en VB6 elles pouvaient *grouper* (et forcer un ByVal) ; ce n'est pas un détail de syntaxe mais une question de **mode de passage**, à traiter avec le § 10.2.

---

*Section suivante : **[§ 10.2 — ByRef par défaut → ByVal par défaut : auditer chaque appel](02-byref-byval-audit.md)** ⭐ ⚠️*

**Cible** : .NET Framework 4.7.2 · **Indicateurs** : ⭐ point clé · ⚠️ piège · 🛠️ outils

⏭️ [**ByRef par défaut → ByVal par défaut** : auditer chaque appel](/10-procedures-fonctions/02-byref-byval-audit.md)
