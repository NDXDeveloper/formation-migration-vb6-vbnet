🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10. Procédures, fonctions et paramètres ⭐ ⚠️

> **Le squelette de toute application VB6 — et l'endroit où les *changements silencieux* se concentrent.**
> Même syntaxe, sens différent : un appel qui compile sans broncher peut ne plus rien faire de ce qu'il faisait.

**Partie 3 — Migrer le langage (le cœur de la formation)** · .NET Framework 4.7.2
⏱️ *Lecture du chapitre : ~2 h · Prérequis : modules 2, 7, 8 et 9*

---

## 🧭 Pourquoi ce chapitre est critique

Procédures (`Sub`), fonctions (`Function`) et propriétés forment le **squelette** de n'importe quelle application VB6 : tout le reste s'y appelle, s'y enchaîne et s'y passe des données. C'est aussi, et ce n'est pas un hasard, **le chapitre où se concentrent le plus de pièges silencieux** de toute la migration.

Le problème n'est pas la syntaxe — elle se ressemble énormément entre VB6 et VB.NET. Le problème est que **le sens de cette syntaxe a changé**. Plusieurs constructions se *compilent sans erreur* après migration, mais ne se *comportent plus de la même manière* :

- un paramètre qui était modifiable par la procédure ne l'est plus ;
- un paramètre optionnel « non fourni » ne se détecte plus comme avant ;
- un tableau passé à une fonction n'est plus relié à celui de l'appelant ;
- une affectation de valeur de retour ne fait plus exactement la même chose.

> ⚠️ **C'est exactement le scénario que cette formation traque en priorité** (voir l'**Introduction** et l'**Annexe B**) : *le code qui compile mais ne se comporte plus pareil*. Ici, le compilateur ne vous protège pas — il valide une intention qui n'est plus la vôtre.

---

## 🎯 Ce que vous saurez faire à la fin du chapitre

- Convertir les appels VB6 (avec ou sans parenthèses, avec `Call`) vers la syntaxe VB.NET ;
- **Auditer chaque paramètre** pour décider, en connaissance de cause, entre `ByRef` et `ByVal` — au lieu de subir le changement de défaut ;
- Migrer les paramètres `Optional` et remplacer `IsMissing` par une stratégie robuste (valeur sentinelle ou surcharge) ;
- Adapter `ParamArray` au nouveau modèle (typé, passé `ByVal`) ;
- Tirer parti de la **surcharge** (nouveauté .NET) et choisir entre `Return` et `NomFonction = valeur` en mesurant la différence de comportement ;
- Reconstruire les propriétés `Property Get`/`Let`/`Set` en un accesseur `Get`/`Set` unique, en repérant le **faux-ami** que constitue le `Set`.

---

## 🔍 Le cœur du problème : la signature change de *sens*

En VB6, une signature de procédure repose sur une série de **valeurs par défaut implicites** très différentes de celles de VB.NET. Migrer « à l'identique » la syntaxe revient donc, sans le savoir, à changer le contrat de la procédure.

| Élément de la signature | Défaut **VB6** | Défaut **VB.NET** | Conséquence si on ne fait rien |
|---|---|---|---|
| Mode de passage d'un paramètre | **ByRef** | **ByVal** | La procédure ne modifie plus l'appelant |
| `Optional` sans valeur | autorisé (+ `IsMissing`) | **interdit** (valeur obligatoire) | Erreur de compilation → refactoring forcé |
| `ParamArray` | `ByRef`, tableau de `Variant` | `ByVal`, tableau **typé** | Sémantique de modification perdue |
| Renvoi de valeur | `NomFonction = valeur` | idem **ou** `Return` | `Return` **sort immédiatement** |

La bonne nouvelle : la plupart de ces changements deviennent **inoffensifs dès qu'ils sont explicités**. Tout le travail de ce chapitre consiste à **rendre explicite ce que VB6 laissait implicite**.

---

## ⚠️ Le piège-vedette : `ByRef` par défaut → `ByVal` par défaut

C'est **le** changement à retenir de ce chapitre, et l'un des plus dangereux de toute la migration. En VB6, un paramètre sans modificateur est passé **par référence** : la procédure peut modifier la variable de l'appelant. En VB.NET, ce même paramètre, toujours sans modificateur, est passé **par valeur** : la procédure travaille sur une copie, et les modifications sont perdues.

```vb
' VB6 — paramètre ByRef PAR DÉFAUT
Sub AppliquerRemise(prix As Double)
    prix = prix * 0.9
End Sub

Dim total As Double
total = 100
AppliquerRemise total       ' total = 90   ✅  l'appelant a bien été modifié
```

```vb
' VB.NET — le MÊME code, sans ByVal/ByRef explicite → ByVal IMPLICITE
Sub AppliquerRemise(prix As Double)
    prix = prix * 0.9
End Sub

Dim total As Double = 100
AppliquerRemise(total)      ' total = 100  ⚠️  la modification est PERDUE
```

Aucune erreur de compilation, aucun avertissement : juste un calcul qui, soudain, ne « remonte » plus son résultat. Sur des centaines d'appels, repérer ce genre de régression à l'œil nu est illusoire.

> 🛠️ **Bonne nouvelle outillée** : l'*Upgrade Wizard* de Visual Studio insère en général un modificateur **explicite** (`ByVal`/`ByRef`) sur chaque paramètre, préservant ainsi le comportement VB6 d'origine. Le danger se niche surtout dans le **code migré à la main**, le **code neuf**, ou le « nettoyage » bien intentionné qui retire ces modificateurs. **La règle de survie : un paramètre = un mode de passage explicite.** Le détail (et le cas subtil des parenthèses qui forcent le passage `ByVal` en VB6) est traité en **§ 10.2**.

---

## 🗺️ Carte du chapitre

Le chapitre suit la signature d'une procédure, du plus mécanique au plus subtil.

**[10.1 — `Sub`/`Function` ; appels avec/sans parenthèses, `Call`](01-sub-function-appels.md)**
Le socle. VB6 tolérait l'appel d'un `Sub` sans parenthèses (`MaSub arg`) et le mot-clé `Call` ; VB.NET impose les parenthèses et rend `Call` obsolète. Conversion essentiellement **mécanique**, mais c'est le point de départ de tout le reste.

**[10.2 — ByRef par défaut → ByVal par défaut : auditer chaque appel](02-byref-byval-audit.md)** ⭐ ⚠️
Le cœur du chapitre. Comment retrouver, qualifier et corriger chaque paramètre dont le mode de passage change de sens — et pourquoi c'est un audit *systématique*, pas une relecture au hasard.

**[10.3 — `Optional` et valeurs par défaut obligatoires ; fin de `IsMissing`](03-optional-ismissing.md)** ⚠️
En VB.NET, tout paramètre `Optional` exige une valeur par défaut, et `IsMissing` disparaît. Comment détecter « l'argument non fourni » autrement (valeur sentinelle, surcharge), proprement.

**[10.4 — `ParamArray` (passé ByRef en VB6 → tableau ByVal en .NET)](04-paramarray.md)** ⚠️
Le nombre variable d'arguments survit, mais change de nature : tableau **typé** et passé **par valeur**. Conséquences sur le typage et sur la possibilité de modifier les éléments.

**[10.5 — Surcharge et valeurs de retour (`Return` vs `NomFonction = valeur`)](05-surcharge-retour.md)**
Deux différences à exploiter : la **surcharge** (inexistante en VB6, idéale pour remplacer les anciens patterns `Optional`/`IsMissing`) et le choix entre `Return` (qui *sort immédiatement*) et l'affectation au nom de la fonction (qui *poursuit l'exécution*).

**[10.6 — `Property Get`/`Let`/`Set` → propriétés](06-property-get-let-set.md)** ⚠️
Les trois procédures de propriété VB6 fusionnent en un bloc `Get`/`Set` unique. Surtout : `Let` (valeur) **et** `Set` (objet) se rejoignent dans le **même** accesseur `Set` — un **faux-ami** du `Set` VB6, qui disparaît par ailleurs comme mot-clé d'affectation d'objet.

---

## 📋 En un coup d'œil

| Sujet | VB6 | VB.NET (4.7.2) | Point de vigilance |
|---|---|---|---|
| Appel de `Sub` | `MaSub arg` (sans parenthèses) | `MaSub(arg)` (parenthèses requises) | Mécanique |
| `Call` | Optionnel | Obsolète (toléré) | Cosmétique |
| Mode de passage par défaut | **ByRef** | **ByVal** | ⭐ ⚠️ Régression silencieuse |
| `Optional` | Sans valeur possible + `IsMissing` | Valeur par défaut **obligatoire**, plus de `IsMissing` | ⚠️ Refactoring (erreur de compilation) |
| `ParamArray` | `ByRef`, tableau de `Variant` | `ByVal`, tableau **typé** | ⚠️ Typage + sémantique |
| Surcharge | Inexistante | Disponible | ✅ Opportunité |
| Valeur de retour | `NomFonction = valeur` | `Return valeur` (ou `NomFonction = valeur`) | ⚠️ `Return` sort immédiatement |
| Propriétés | `Get` / `Let` / `Set` (3 procédures) | Un bloc `Get`/`Set` ; `Let` + `Set` → un seul `Set` | ⚠️ Faux-ami du `Set` |

---

## 🔗 Liens avec le reste de la formation

- **[Module 2 — Différences fondamentales](../02-differences-fondamentales/README.md)** : le changement ByRef→ByVal y est présenté comme l'un des pièges majeurs (§ 2.4). Ce chapitre-ci en donne le traitement opérationnel.
- **[Module 7 — Types, variables et déclarations](../07-types-variables/README.md)** : indispensable pour les paramètres, car un `Integer` ou un `Currency` ne désigne pas la même chose des deux côtés (un piège **distinct** de celui du mode de passage).
- **[Module 12 — Programmation orientée objet](../12-poo/README.md)** : les propriétés (§ 10.6) y sont reprises dans le contexte plus large des classes, de l'encapsulation et des accesseurs.
- **[Annexe A — Correspondance VB6 → VB.NET](../annexes/correspondance-vb6-vbnet/README.md)** : l'aide-mémoire à garder ouvert pour les mots-clés d'appel et de déclaration.
- **[Annexe B — Catalogue des pièges silencieux](../annexes/pieges-silencieux/README.md)** : entrées **B.1** (ByRef/ByVal) et **B.7** (propriétés par défaut et `Set`) — à lire en parallèle de ce chapitre.

---

## ✅ À retenir avant d'attaquer la suite

1. **Même syntaxe ≠ même comportement.** Ici, le compilateur valide du code dont le *sens* a changé : la vigilance ne peut pas être déléguée à la compilation seule.
2. **Le mode de passage est la priorité absolue.** `ByRef` par défaut (VB6) devient `ByVal` par défaut (VB.NET) : tant qu'un paramètre n'est pas explicitement qualifié, considérez-le comme suspect.
3. **Expliciter = se protéger.** La quasi-totalité des pièges de ce chapitre se neutralisent en rendant explicite ce que VB6 laissait implicite (mode de passage, valeur par défaut, intention de retour).
4. **Certains changements sont « bruyants », d'autres « silencieux ».** `Optional` sans valeur *ne compile pas* (vous serez prévenu) ; ByRef→ByVal *compile très bien* (personne ne vous prévient). Ce sont les seconds qui coûtent cher.

---

*Chapitre suivant à lire : **[§ 10.1 — `Sub`/`Function` ; appels avec/sans parenthèses, `Call`](01-sub-function-appels.md)**.*

**Cible** : .NET Framework 4.7.2 · **Indicateurs** : ⭐ point clé · ⚠️ piège · 🛠️ outils

⏭️ [`Sub`/`Function` ; appels avec/sans parenthèses, `Call`](/10-procedures-fonctions/01-sub-function-appels.md)
