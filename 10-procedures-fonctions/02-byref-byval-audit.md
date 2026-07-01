🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.2 — ByRef par défaut → ByVal par défaut : auditer chaque appel ⭐ ⚠️

> **Le piège n°1 du chapitre, et l'un des plus dangereux de toute la migration.** La même signature de procédure veut dire l'**inverse** en VB6 et en VB.NET — et le compilateur ne dit rien. La parade n'est pas une relecture au hasard, mais un **audit systématique**.

📍 *Module 10 · § 10.2 · [↑ Introduction du chapitre](README.md) · [← § 10.1](01-sub-function-appels.md)*

---

## ⚡ En bref

- En **VB6**, un paramètre sans modificateur est passé **`ByRef`** (par référence). En **VB.NET**, le même paramètre, toujours sans modificateur, est passé **`ByVal`** (par valeur). **Le défaut s'inverse.**
- Conséquence : une procédure qui « renvoyait » un résultat via un paramètre **cesse silencieusement de le faire**. Le code **compile sans erreur ni avertissement** ; `Option Strict On` n'y change **rien** (il concerne le typage, pas le mode de passage).
- La seule défense fiable : **rendre explicite le mode de passage de *chaque* paramètre**, puis **auditer *chaque* site d'appel**, et **valider par des tests de non-régression** (golden master).

> ⚠️ Cette section mérite votre temps : c'est le passage où une migration « qui compile » peut produire des résultats faux de façon **indétectable à l'œil**.

---

## 1. Le basculement, en une phrase

> En VB6, le défaut est **`ByRef`**. En VB.NET, le défaut est **`ByVal`**.

La **même ligne** de déclaration change donc de sens :

```vb
Sub Calculer(montant As Double)     ' ← exactement la même écriture des deux côtés
```

| Interprétation | VB6 | VB.NET |
|---|---|---|
| `Sub Calculer(montant As Double)` | `montant` est **`ByRef`** | `montant` est **`ByVal`** |
| La procédure peut modifier la variable de l'appelant ? | **Oui** | **Non** |

Petit rappel des deux modes :

- **`ByVal`** : la procédure reçoit une **copie**. Ce qu'elle écrit dans le paramètre reste **local** ; l'appelant ne voit rien.
- **`ByRef`** : la procédure reçoit un **lien vers la variable** de l'appelant. Ce qu'elle écrit **remonte** jusqu'à l'appelant.

C'est exactement ce comportement « qui remonte » qui disparaît silencieusement quand un `ByRef` implicite devient un `ByVal` implicite.

> 🔗 Le même piège est présenté côté « différences fondamentales » au **[§ 2.4](../02-differences-fondamentales/04-byref-byval.md)** et catalogué en **[Annexe B.1](../annexes/pieges-silencieux/README.md)**. La présente section en donne le **traitement opérationnel**.

---

## 2. Pourquoi c'est dangereux : un exemple qui casse en silence

L'idiome le plus touché est **omniprésent** dans le code VB6 : faute de tuples ou de retours multiples, on renvoie un **statut** par la valeur de la fonction, et les **résultats** par des paramètres `ByRef` (implicites).

```vb
' === VB6 ===  paramètre 'resultat' ByRef PAR DÉFAUT
Function TryLireMontant(texte As String, resultat As Double) As Boolean
    If IsNumeric(texte) Then
        resultat = CDbl(texte)          ' écrit dans la variable de l'appelant
        TryLireMontant = True
    Else
        TryLireMontant = False
    End If
End Function

Dim m As Double
If TryLireMontant("42.5", m) Then
    ' m vaut 42.5   ✅
End If
```

Transcrit **« à l'identique »** en VB.NET, sans modificateur, le paramètre devient **`ByVal`** :

```vb
' === VB.NET (NAÏF, CASSÉ) ===  'resultat' ByVal IMPLICITE
Function TryLireMontant(texte As String, resultat As Double) As Boolean
    If IsNumeric(texte) Then
        resultat = CDbl(texte)          ' ⚠️ écrit dans une COPIE locale
        Return True
    Else
        Return False
    End If
End Function

Dim m As Double
If TryLireMontant("42.5", m) Then
    ' m vaut TOUJOURS 0   ⚠️  le résultat ne remonte plus
End If
```

La fonction renvoie bien `True` ; la condition est donc franchie comme avant ; mais `m` reste à `0`. **Aucune erreur, aucun avertissement.** Le programme « marche » — il calcule simplement faux.

La correction consiste à **expliciter** le mode voulu :

```vb
' === VB.NET (CORRECT) ===  mode de passage RENDU EXPLICITE
Function TryLireMontant(ByVal texte As String, ByRef resultat As Double) As Boolean
    If IsNumeric(texte) Then
        resultat = CDbl(texte)
        Return True
    Else
        Return False
    End If
End Function
```

> ⚠️ **À retenir** : le compilateur ne vous protège pas ici. Ni une compilation propre, ni `Option Strict On` ne signalent ce basculement. Seuls un **audit délibéré** et des **tests de comportement** le détectent.

---

## 3. Le cas des objets : `ByRef` n'est *pas* « pouvoir modifier l'objet »

C'est la confusion la plus fréquente, et elle change la façon d'auditer. Pour un **type référence** (un objet), `ByVal` passe une **copie de la référence**, pas une copie de l'objet. Donc :

- avec **`ByVal`**, la procédure peut toujours **modifier les membres** de l'objet existant (les deux variables pointent le même objet) — ces modifications **sont visibles** par l'appelant ;
- ce que `ByVal` **interdit**, c'est de **réaffecter** le paramètre à un **autre** objet (ou à `Nothing`) et que l'appelant le voie ; cela exige **`ByRef`**.

```vb
Class Compte
    Public Solde As Decimal
End Class

' ByVal (défaut VB.NET) : COPIE DE LA RÉFÉRENCE
Sub Mouvementer(ByVal c As Compte, ByVal montant As Decimal)
    c.Solde += montant          ' ✅ modifie l'objet pointé → VISIBLE par l'appelant
    c = New Compte()            ' ⚠️ réaffecte la copie locale → INVISIBLE par l'appelant
End Sub

' ByRef : la VARIABLE de l'appelant elle-même
Sub Remplacer(ByRef c As Compte)
    c = New Compte()            ' ✅ réaffecte la variable de l'appelant → VISIBLE
End Sub
```

| La procédure, sur un paramètre objet… | Mode nécessaire |
|---|---|
| lit l'objet, ou **modifie ses membres** (`c.Solde = …`) | **`ByVal`** suffit |
| **réaffecte** le paramètre (`c = New …` / `c = Nothing`) et l'appelant doit le voir | **`ByRef`** ⚠️ |

> 💡 **Conséquence pour l'audit** : un paramètre objet ne « bascule » de façon visible que s'il est **réaffecté** dans le corps de la procédure. S'il n'est que **muté** (membres modifiés), le passage en `ByVal` ne change rien. Cela **réduit** le nombre de paramètres objets réellement à risque.

---

## 4. Le cas des expressions et des parenthèses

Tous les sites d'appel ne sont pas concernés de la même manière.

**Arguments non modifiables.** Si l'appelant passe une **constante**, un **littéral** ou une **expression** (et non une variable affectable), aucune donnée ne peut « remonter », ni en VB6 ni en VB.NET. Ces appels-là sont **insensibles** au basculement et peuvent être écartés de l'audit. (VB.NET accepte d'ailleurs de passer un littéral à un paramètre `ByRef` : il utilise une variable temporaire, et l'écriture est simplement abandonnée.)

**Parenthèses « groupantes ».** Rappel du **[§ 10.1](01-sub-function-appels.md)** : en VB6, sur un **seul** argument en position d'instruction, `Proc (x)` n'est pas un appel « avec argument » mais une **expression entre parenthèses**, ce qui **force un `ByVal`** — même si le paramètre est `ByRef`. En VB.NET, c'est l'inverse : `Proc(x)` est la syntaxe d'appel normale, et il faut une paire **supplémentaire**, `Proc((x))`, pour forcer le `ByVal`.

| Au site d'appel | VB6 | VB.NET |
|---|---|---|
| `Proc(x)` *(un argument, instruction)* | `(x)` groupé → **ByVal forcé** | liste d'arguments → **mode déclaré respecté** |
| `Proc((x))` | → **ByVal forcé** | parenthèses en trop → **ByVal forcé** |

> ⚠️ Lors de l'audit des **appels**, repérez ces parenthèses : un appel VB6 qui *semblait* `ByRef` pouvait en réalité être `ByVal` à cause d'un `(x)` groupant. Transcrit en `Proc(x)`, il **rétablit** le mode déclaré — et peut donc **changer de comportement** dans un sens comme dans l'autre.

---

## 5. La méthode d'audit : auditer *chaque* appel

L'expression du titre n'est pas une figure de style. Voici une démarche **systématique** et reproductible.

### Étape 0 — Tout expliciter

Avant de raisonner, **aucun paramètre ne doit rester sans `ByVal`/`ByRef`**. Tant que des modificateurs sont implicites, on ne peut rien affirmer. Les outils (§ 6) le font automatiquement ; en code écrit ou migré à la main, **faites-le délibérément**, paramètre par paramètre.

### Étape 1 — Classer chaque *paramètre* par intention

Pour chaque paramètre, posez **une** question : *la procédure doit-elle écrire dans la variable de l'appelant ?*

| Intention du paramètre | Type | Mode |
|---|---|---|
| **Entrée pure** (lecture seule) | valeur **ou** objet | **`ByVal`** |
| **Entrée/sortie** ou **sortie** (on écrit une valeur qui doit remonter) | valeur : `Integer`, `Double`, `String`, `Boolean`, `Date`, `Structure`… | **`ByRef`** ⚠️ |
| Modification des **membres** d'un objet existant | objet (référence) | **`ByVal`** suffit |
| **Réaffectation** du paramètre objet (`= New …` / `= Nothing`) visible par l'appelant | objet (référence) | **`ByRef`** ⚠️ |

Les lignes ⚠️ sont les **seules** réellement dangereuses : ce sont elles qui *doivent* rester `ByRef` sous peine de régression silencieuse.

### Étape 2 — Auditer chaque *site d'appel*

Pour chaque paramètre classé `ByRef`, parcourez **tous ses appels** et vérifiez :

1. l'appelant passe-t-il une **variable affectable** (variable, élément de tableau, champ) ? Si oui, **des données remontent** — confirmez que l'appelant en dépend réellement.
2. y a-t-il des **parenthèses groupantes** (`Proc (x)`) qui masquaient un `ByVal` côté VB6 ? (voir § 4)
3. l'argument est-il une **expression/un littéral** ? Alors ce site est **hors risque**.

### Étape 3 — Valider par le comportement

L'audit visuel ne suffit jamais sur un code réel. Comparez l'avant/après avec un **[harnais de tests de référence (golden master)](../05-preparer-code-vb6/05-golden-master.md)** et des **[tests de non-régression](../17-valider-refactoriser/03-non-regression.md)**. Le basculement `ByRef`/`ByVal` est précisément le genre de régression que ces tests rattrapent là où l'œil échoue (voir aussi **[§ 17.4 — traquer les pièges silencieux](../17-valider-refactoriser/04-traquer-pieges.md)**).

### Étape 4 — S'appuyer sur l'outillage

Le compilateur ne signale pas le basculement, mais d'autres aides existent : recherche systématique des signatures, et **[détection assistée par IA](../19-migration-ia/03-detecter-pieges.md)** des paramètres écrits puis attendus par l'appelant. **Aucun outil ne décide de l'*intention*** à votre place — d'où la nécessité de l'étape 1.

> 🧭 **Cas particuliers à connaître** : (a) passer une **propriété** `ByRef` en VB.NET fonctionne par **copie-entrée / copie-sortie** (le `Get` est lu, la valeur passée, puis le `Set` rappelé au retour) ; (b) une `Structure` est un **type valeur** : pour qu'une nouvelle structure remonte, il faut **`ByRef`**, comme pour un `Integer`.

---

## 6. Ce que font (et ne font pas) les outils

- L'**[*Upgrade Wizard*](../04-outils-migration/01-upgrade-wizard.md)** de Visual Studio insère un modificateur **explicite** sur chaque paramètre afin de **préserver le comportement VB6 d'origine**. C'est, sur ce point précis, l'aspect le plus **fiable** de la conversion automatique.
- **Zones de danger**, à l'inverse :
  - le **code migré à la main** (où l'on oublie le modificateur → `ByVal` par défaut) ;
  - le **code neuf** écrit avec un réflexe VB6 (« le défaut, c'est ByRef ») ;
  - le **« nettoyage »** bien intentionné qui retire les `ByRef` explicites posés par l'outil — et réintroduit le bug.
- **Aucun outil** ne tranche la question d'intention de l'étape 1 : faut-il que ce paramètre remonte une valeur ? Cela reste un **jugement humain**, à valider par les tests.

> 🛠️ **Règle de survie** : *un paramètre = un mode de passage explicite.* Ne laissez jamais le défaut décider à votre place.

---

## 📋 Tableau de synthèse

| Aspect | VB6 | VB.NET (4.7.2) | Risque |
|---|---|---|---|
| Mode de passage **par défaut** | `ByRef` | `ByVal` | ⭐ ⚠️ Régression **silencieuse** |
| Détection par le compilateur | — | **Aucune** (même avec `Option Strict On`) | ⚠️ |
| Paramètre **valeur** écrit pour remonter | `ByRef` implicite | doit être `ByRef` **explicite** | ⚠️ Élevé |
| Paramètre **objet** dont on modifie les membres | indifférent | `ByVal` **suffit** | ✅ Faible |
| Paramètre **objet réaffecté** (`= New…`) | `ByRef` implicite | doit être `ByRef` **explicite** | ⚠️ |
| `Proc(x)` sur un argument (instruction) | `(x)` groupé → ByVal | liste d'arguments → mode déclaré | ⚠️ Sens inversé |
| Posture recommandée | — | **Tout expliciter + auditer chaque appel + tester** | — |

---

## ✅ À retenir

1. **Le défaut s'inverse** : `ByRef` (VB6) → `ByVal` (VB.NET). La même signature change de sens, et **le compilateur se tait**.
2. **L'idiome « statut + résultats via ByRef »** est le plus exposé : c'est lui qui casse en silence en premier.
3. **Pour les objets**, `ByVal` permet toujours de **muter les membres** ; seul un paramètre **réaffecté** exige `ByRef`. Cela cible l'audit.
4. **Les parenthèses ne sont pas neutres** : un `(x)` groupant masquait un `ByVal` en VB6 — à repérer au site d'appel.
5. **La méthode** : tout expliciter → classer chaque paramètre → auditer chaque appel → valider par le **golden master**. *Aucun outil ne décide de l'intention à votre place.*

---

*Section suivante : **[§ 10.3 — `Optional` et valeurs par défaut obligatoires ; fin de `IsMissing`](03-optional-ismissing.md)** ⚠️*

**Cible** : .NET Framework 4.7.2 · **Indicateurs** : ⭐ point clé · ⚠️ piège · 🛠️ outils · 🔗 interop

⏭️ [`Optional` et **valeurs par défaut obligatoires** ; fin de `IsMissing`](/10-procedures-fonctions/03-optional-ismissing.md)
