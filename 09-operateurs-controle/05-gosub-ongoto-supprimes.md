🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.5 `GoSub`/`Return`, `On…GoTo`/`On…GoSub`, numéros de ligne : **supprimés** → refactoring

> **Module 9 — Opérateurs, chaînes et structures de contrôle** · [↩ Retour au chapitre](README.md)
> **Indicateur de cette section** : ⚠️ le refactoring recèle des pièges (faux-ami `Return`, état partagé)

Ces constructions sont des héritages des **vieux BASIC** (GW-BASIC, QuickBASIC), déjà considérés comme obsolètes en VB6. En VB.NET, elles sont **supprimées** : leur présence déclenche une **erreur de compilation**.

C'est, à un égard, une **bonne nouvelle** : pas de changement de comportement silencieux au niveau de la construction — le compilateur vous **force** à les traiter. Le ⚠️ de cette section ne porte donc pas sur les constructions elles-mêmes, mais sur le **refactoring** : il recèle deux vrais pièges — le **faux-ami `Return`** et la **préservation de l'état partagé** — qui peuvent, eux, changer le comportement en silence.

---

## Vue d'ensemble

| Construction VB6 | Statut en VB.NET | Cible de refactoring |
|------------------|------------------|----------------------|
| `GoSub` / `Return` (retour) | ❌ Supprimé — `Return` **réaffecté** | `Sub`/`Function` (`ByRef`) ou lambda multiligne |
| `On N GoTo …` | ❌ Supprimé | `Select Case` (+ `Case Else`) |
| `On N GoSub …` | ❌ Supprimé | `Select Case` appelant des méthodes |
| Numéros de ligne (`10`, `20`…) | Sans rôle | À supprimer |
| `Erl` (n° de ligne en erreur) | Pas d'équivalent | `Try`/`Catch` + `ex.StackTrace` (module **11**) |

> ℹ️ **À ne pas confondre** : le `GoTo label` **simple subsiste** en VB.NET (toujours déconseillé). Ce sont `GoSub`, le saut calculé `On…GoTo` et `On…GoSub` qui **disparaissent**.

---

## `GoSub` / `Return` (le cas le plus délicat)

### Le comportement VB6

`GoSub label` se comporte comme une **« sous-routine » locale** : on saute à une étiquette **dans la même procédure**, on exécute jusqu'à `Return`, puis on **revient à la ligne suivant le `GoSub`**. La « sous-routine » **partage toutes les variables locales** de la procédure (pas de paramètres, pas d'isolation).

```vb
' VB6 — GoSub partage les variables locales de la procédure
Sub Traiter()
    Dim x As Integer
    x = LireValeur()
    GoSub Normaliser
    Enregistrer x
    Exit Sub
Normaliser:
    If x < 0 Then x = 0
    If x > 100 Then x = 100
    Return              ' revient APRÈS le GoSub
End Sub
```

### ⚠️ Le faux-ami `Return`

Voici le piège central de la section. En VB.NET, le mot-clé `Return` **existe toujours** — mais signifie tout **autre chose** :

```vb
' ⚠️ Le mot-clé Return change de sens
'   VB6 (dans un GoSub) : Return = « revenir après le GoSub »
'   VB.NET              : Return = « SORTIR de la procédure » (= Exit Sub / renvoyer une valeur)
Function Calculer() As Integer
    ' ...
    Return resultat     ' .NET : QUITTE la fonction en renvoyant resultat
End Function
```

Conséquence : une transposition mécanique qui **conserve un `Return`** là où il signifiait « revenir » va, en VB.NET, **quitter la procédure** — un changement de flux de contrôle **catastrophique et difficile à repérer**. Heureusement, `GoSub` étant une **erreur de compilation**, vous serez forcé de toucher à ces blocs ; profitez-en pour **vérifier chaque `Return`**.

### Le refactoring : extraire en méthode

On transforme la « sous-routine » `GoSub` en **vraie `Sub`/`Function`**. Comme le bloc partageait les variables locales, l'état doit être **transmis explicitement** (paramètres `ByRef` pour préserver les modifications) :

```vb
' VB.NET — extraction en méthode ; l'état partagé passe en paramètre ByRef
Sub Traiter()
    Dim x As Integer = LireValeur()
    Normaliser(x)
    Enregistrer(x)
End Sub

Private Sub Normaliser(ByRef valeur As Integer)
    If valeur < 0 Then valeur = 0
    If valeur > 100 Then valeur = 100
End Sub
```

> ⚠️ **Piège de l'état partagé.** Le `GoSub` modifiait des variables **locales partagées**. Si vous extrayez le bloc dans une méthode en passant les variables **`ByVal`** (ou sans les passer du tout), les modifications **ne remontent plus** — régression silencieuse. Passez **`ByRef`** ce que le bloc modifiait (rappel : `ByVal` est le défaut en VB.NET, voir modules **2.4** et **10.2**).

**Variante** — si l'on veut coller au plus près du partage de variables locales (sans paramètres), une **lambda multiligne** qui **capture** les locales fait l'affaire :

```vb
' VB.NET — lambda multiligne capturant les variables locales (modifie x directement)
Sub Traiter()
    Dim x As Integer = LireValeur()
    Dim Normaliser As Action =
        Sub()
            If x < 0 Then x = 0
            If x > 100 Then x = 100
        End Sub
    Normaliser()        ' x est modifié par capture
    Enregistrer(x)
End Sub
```

En général, la **vraie méthode avec paramètres explicites** est préférable : elle rend visibles les entrées/sorties que `GoSub` laissait implicites.

---

## `On…GoTo` / `On…GoSub` (saut calculé)

### Le comportement VB6

`On expression GoTo étiq1, étiq2, étiq3, …` est un **saut indexé** : l'expression est évaluée en un entier `N` ; `N`=1 saute à `étiq1`, `N`=2 à `étiq2`, etc. **Si `N`=0 ou dépasse le nombre d'étiquettes, aucun saut n'a lieu** : l'exécution **continue** après l'instruction.

```vb
' VB6 — saut calculé selon la valeur de choix
On choix GoTo Ajouter, Modifier, Supprimer
' choix = 0 ou hors plage -> AUCUN saut, on tombe ici
MsgBox "Choix invalide"
Exit Sub
Ajouter:
    ' ...
    Exit Sub
Modifier:
    ' ...
    Exit Sub
Supprimer:
    ' ...
End Sub
```

### Le refactoring : `Select Case`

```vb
' VB.NET — Select Case appelant les méthodes extraites
Select Case choix
    Case 1 : Ajouter()
    Case 2 : Modifier()
    Case 3 : Supprimer()
    Case Else                          ' ⚠️ reproduit le « 0 ou hors plage » de On…GoTo
        MessageBox.Show("Choix invalide")
End Select
```

`On…GoSub` se traite de la même façon : les cibles deviennent des **méthodes**, appelées depuis les `Case` (c'est le refactoring `GoSub` ci-dessus, appliqué à chaque branche).

> ⚠️ **Ne perdez pas le cas « hors plage ».** Le comportement de **non-saut** quand `N`=0 ou `N` dépasse le nombre d'étiquettes est facile à oublier. Mappez-le explicitement sur `Case Else` (même si c'était « ne rien faire »), sous peine de changer le comportement.

---

## Numéros de ligne (et `Erl`)

VB6 autorise des **numéros de ligne** en tête d'instruction (`10 …`, `20 …`), héritage des BASIC anciens. Ils servaient de cibles de saut — rôle qui **disparaît** avec `GoSub`/`On…GoTo` — et, surtout, de support à **`Erl`** : la fonction qui renvoie le **numéro de la ligne** où s'est produite la dernière erreur. Beaucoup d'applications VB6 numérotaient chaque ligne pour journaliser « erreur à la ligne 1240 ».

En VB.NET, ce mécanisme **n'a pas d'équivalent** : on passe aux **exceptions** et aux **traces de pile**.

```vb
' VB6 — numéros de ligne + Erl pour journaliser la ligne fautive
Sub Traiter()
    On Error GoTo Gestion
10  Ouvrir()
20  Lire()
30  Fermer()
    Exit Sub
Gestion:
    Journaliser "Erreur ligne " & Erl & " : " & Err.Description
End Sub
```

```vb
' VB.NET — plus de numéros de ligne ; diagnostics par exception et trace de pile
Sub Traiter()
    Try
        Ouvrir()
        Lire()
        Fermer()
    Catch ex As Exception
        Journaliser($"Erreur : {ex.Message}{Environment.NewLine}{ex.StackTrace}")
    End Try
End Sub
```

Avec les **symboles de débogage (PDB)** inclus dans la build, la trace de pile fournit méthode **et** ligne — ce que `Erl` donnait, en mieux. Le passage complet `On Error` → `Try`/`Catch` (et l'objet `Err`) est traité au **module 11**.

---

## L'assistant de migration face à ces constructions

L'assistant de mise à niveau ne traite ces constructions que **partiellement** et de façon **inégale** : il peut retirer les numéros de ligne, mais `GoSub`/`Return` et les sauts calculés demandent presque toujours un **refactoring manuel**. Puisque ce sont des **erreurs de compilation**, vous ne pourrez de toute façon pas les ignorer — l'occasion idéale de les remplacer par des **structures propres** (méthodes, `Select Case`, `Try`/`Catch`) plutôt que par une transposition littérale. Couvrez le résultat par les tests de non-régression (module **5.5**), car le refactoring touche au flux de contrôle.

---

## 🎯 Synthèse — points de vigilance de la section

| Élément | Risque | Réflexe |
|---------|--------|---------|
| Mot-clé `Return` | **Faux-ami** : « revenir » (VB6) → « sortir » (.NET) | Vérifier **chaque** `Return` au refactoring de `GoSub` |
| État partagé du `GoSub` | Locales partagées perdues à l'extraction | Passer en **`ByRef`** (ou capturer via lambda) |
| `On N GoTo` hors plage | `N`=0 / hors plage = aucun saut (fall-through) | Reproduire avec **`Case Else`** |
| Numéros de ligne | Sans rôle ; cassent à la compilation | Supprimer |
| `Erl` | Pas d'équivalent | `Try`/`Catch` + `ex.StackTrace` (module 11) |
| `GoTo label` simple | *Conservé* — souvent cru supprimé | Pas obligatoire de retirer ; structurer si possible |

---

## 🔗 Pour aller plus loin

- Module **11** — Gestion des erreurs : `On Error` → `Try`/`Catch`, objet `Err`, diagnostics (remplace `Erl`)
- Module **2.4** / **10.2** — `ByRef` par défaut → `ByVal` (crucial pour préserver l'état partagé du `GoSub`)
- Module **9.4** — `Continue For/Do/While` remplace le `GoTo` placé près d'un `Next`/`Loop`
- Module **12** — Programmation orientée objet (extraire proprement en méthodes)
- Module **17.5** — Refactoring idiomatique
- Module **5.5** — Harnais de tests de référence (*golden master*) pour sécuriser le refactoring
- Annexe **A** — Tableau de correspondance VB6 → VB.NET

---

> **Section précédente** → [9.4 — Boucles](04-boucles.md)
> **Section suivante** → [9.6 — Instructions multiples, séparateur « `:` », étiquettes de ligne](06-instructions-etiquettes.md)

⏭️ [Instructions multiples, séparateur « `:` », étiquettes de ligne](/09-operateurs-controle/06-instructions-etiquettes.md)
