🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.3 `If…Then`, `Select Case`, `IIf` → opérateur `If()` (et l'évaluation court-circuit)

> **Module 9 — Opérateurs, chaînes et structures de contrôle** · [↩ Retour au chapitre](README.md)
> **Indicateur de cette section** : ⚠️ un changement de sémantique facile à manquer

Les structures conditionnelles migrent, dans l'ensemble, **proprement** : un `If…Then…Else` reste un `If…Then…Else`, un `Select Case` reste un `Select Case`. L'exception — et c'est tout l'objet du ⚠️ — est le passage de la **fonction `IIf`** à l'**opérateur `If()`**. Cela ressemble à un simple renommage. **Ce n'en est pas un** : c'est un **changement de comportement** (l'apparition du court-circuit).

Cette section **referme l'histoire du court-circuit** ouverte en **9.1** (avec `And`/`Or` vs `AndAlso`/`OrElse`).

---

## `If…Then…Else`

Syntaxe **conservée** : `If…Then…Else`, `ElseIf`, `End If`, et la forme sur une seule ligne. Deux points d'attention.

**La condition doit être un `Boolean`.** En VB6, un booléen *est* un entier (`True` = `-1`, `False` = `0`), et `If` accepte donc une expression **numérique** : non nul = vrai, zéro = faux. Sous `Option Strict On`, VB.NET exige un vrai `Boolean` :

```vb
' VB6 — condition numérique tolérée
If compteur Then ...          ' vrai si compteur <> 0

' VB.NET avec Option Strict On
If compteur Then ...          ' ERREUR : Integer n'est pas convertible en Boolean
If compteur <> 0 Then ...     ' correction explicite
```

Le même réflexe vaut pour `If unObjet Then` ou `If uneChaine Then`, idiomes hérités à rendre explicites (`If unObjet IsNot Nothing Then`, `If uneChaine <> "" Then`). Sous `Option Strict Off`, ces formes *compilent* encore — raison de plus pour activer `Option Strict` progressivement (module **17.1**).

> 🔗 Le piège `True = -1` et les comparaisons du type `If x = True` sont traités au module **7.6** et à l'**annexe B.8**. Ne comparez jamais une condition à `True`/`1` : testez la condition elle-même.

*(La forme sur une seule ligne `If cond Then a = 1 : b = 2` conserve son comportement : tout ce qui suit `Then` sur la ligne appartient à la branche conditionnelle. Lisible, mais à éviter en code neuf.)*

---

## `Select Case`

`Select Case` migre **directement** et reste tout aussi expressif qu'en VB6 :

```vb
Select Case note
    Case Is >= 16 : mention = "Très bien"
    Case 14 To 15 : mention = "Bien"
    Case 10, 11   : mention = "Passable"
    Case Else     : mention = "À revoir"
End Select
```

L'expression de test est évaluée **une seule fois**, les `Case` sont examinés **dans l'ordre**, le **premier qui correspond gagne**, et il **n'y a pas de fall-through** (pas de cascade entre `Case`) — exactement comme en VB6. Aucune surprise de ce côté.

> ⚠️ **Le seul vrai piège : la casse des chaînes.** Comme pour `Like` et les comparaisons (section 9.1/9.2), la correspondance d'un `Case` sur une **chaîne** dépend d'`Option Compare` :
>
> ```vb
> Select Case role
>     Case "admin"     ' Binary -> "Admin" ne correspond PAS
>         ...          ' Text   -> "Admin" correspond
> End Select
> ```
>
> Un `Option Compare` non reporté fait basculer **silencieusement** tous ces tests en sensible à la casse. À vérifier (module **6.4**).

> 🔗 `Select Case` est aussi l'**outil de remplacement** des instructions `On…GoTo` / `On…GoSub` (saut calculé), **supprimées** en VB.NET. Voir module **9.5**.

---

## ⚠️ `IIf` → opérateur `If()` : un changement de **sémantique**, pas un renommage

Voici le cœur de la section.

### `IIf` est une **fonction** — elle évalue **les deux** branches

En VB6, `IIf(condition, valeurSiVrai, valeurSiFaux)` est une **fonction**. Comme toute fonction, **ses arguments sont évalués avant l'appel** — donc **les deux** branches sont **toujours** calculées, quelle que soit la condition. D'où deux classes de bugs bien connues :

```vb
' (1) Division : la branche "vraie" est calculée MÊME quand diviseur = 0
resultat = IIf(diviseur <> 0, numerateur / diviseur, 0)
'                             └────────────────────┘  -> erreur de division par zéro

' (2) Effets de bord : les DEUX fonctions sont appelées
x = IIf(condition, Enregistrer(), Annuler())   ' Enregistrer() ET Annuler() s'exécutent
```

De plus, `IIf` renvoie un `Object` (anciennement `Variant`), ce qui ajoute des questions de typage.

### L'opérateur `If()` **court-circuite**

VB.NET introduit l'opérateur conditionnel **ternaire** `If(condition, valeurSiVrai, valeurSiFaux)` — même mot-clé `If`, mais à **trois arguments**. Contrairement à `IIf`, il **court-circuite** : **seule la branche retenue** est évaluée.

```vb
' Si diviseur = 0, la division n'est PAS évaluée -> renvoie 0 sans erreur
resultat = If(diviseur <> 0, numerateur / diviseur, 0)

' Une seule des deux fonctions s'exécute
x = If(condition, Enregistrer(), Annuler())
```

Il **infère** aussi un type commun aux deux branches (typage plus sûr que l'`Object` d'`IIf`), et permet enfin des **expressions gardées contre `Nothing`** :

```vb
' Possible en VB.NET grâce au court-circuit
nom = If(client IsNot Nothing, client.Nom, "(inconnu)")
'   client.Nom n'est lu que si client n'est pas Nothing
'   En VB6, IIf aurait lu client.Nom même sur Nothing -> erreur
```

Sa **forme à deux arguments** est un bonus : l'opérateur de **coalescence de null** (équivalent du `??` de C#) :

```vb
affichage = If(valeur, "(vide)")    ' valeur si non Nothing, sinon "(vide)"
```

### Le piège : `IIf` **existe toujours** et garde son ancien comportement

C'est ce qui rend la migration délicate. `IIf` **n'a pas disparu** : il reste disponible dans `Microsoft.VisualBasic` et **continue d'évaluer les deux branches**. Deux scénarios :

- **On laisse `IIf` tel quel** → comportement **préservé** (les deux branches évaluées). C'est sûr du point de vue de la migration… mais cela **transporte les bugs latents** (division, effets de bord) tels quels.
- **On remplace `IIf` par `If()`** → le comportement **change** (court-circuit). C'est presque toujours **souhaitable** (cela corrige la classe de bugs « division par zéro »), mais c'est une **modification observable** : si du code s'appuyait sur l'exécution des **deux** effets de bord, l'un d'eux **cessera de s'exécuter**.

> ⚠️ **Règle de migration.** Migrer `IIf` → `If()` est **recommandé**, mais ce **n'est pas un rechercher-remplacer**. Pour chaque `IIf`, vérifiez :
> 1. Les branches ont-elles des **effets de bord** qui doivent **tous deux** se produire ? Si oui, `If()` en supprimerait un.
> 2. L'évaluation des deux branches **masquait-elle** ou **compensait-elle** quelque chose ?
>
> Traitez chaque conversion comme un **changement de comportement à valider** contre le *golden master* (module **5.5**), pas comme un simple renommage. Et ne confondez pas les deux usages du mot-clé : `If(...)` (parenthèses + virgules) = **opérateur** ; `If … Then` = **instruction**.

---

## Les autres fonctions « qui évaluent tout » : `Switch`, `Choose`

`IIf` n'est pas seule dans son genre. VB6 propose aussi `Switch` et `Choose`, **toutes deux conservées** en VB.NET (`Microsoft.VisualBasic`) et **toutes deux sans court-circuit** : elles évaluent **l'intégralité** de leurs arguments.

```vb
' Switch : valeur associée à la 1re expression vraie -- mais TOUTES les paires sont évaluées
r = Switch(x < 0, "négatif", x = 0, "nul", x > 0, "positif")

' Choose : valeur à la position 'index' -- mais TOUTES les valeurs sont évaluées
c = Choose(index, OptionA(), OptionB(), OptionC())   ' les trois fonctions s'exécutent
```

Le même réflexe qu'avec `IIf` s'applique : attention aux **effets de bord** et aux calculs coûteux. Pour un choix multi-branches **à court-circuit**, préférez un `Select Case` ou des `If()` imbriqués.

---

## 🔁 Le fil du court-circuit — vue d'ensemble (sections 9.1 + 9.3)

Le modèle mental est simple : **VB6 n'a de court-circuit nulle part**. VB.NET l'introduit à **deux niveaux**, et dans les deux cas la **forme héritée est conservée** (donc l'assistant préserve le comportement), tandis que la **forme à court-circuit doit être introduite volontairement**.

| Besoin | Forme « évalue tout » (héritée, conservée) | Forme « court-circuit » (VB.NET) |
|--------|--------------------------------------------|----------------------------------|
| ET / OU logique | `And` / `Or` | `AndAlso` / `OrElse` |
| Expression conditionnelle | `IIf(c, a, b)` | `If(c, a, b)` |
| Choix multi-branches | `Switch` / `Choose` | `Select Case` / `If()` imbriqués |

Les deux niveaux se combinent dans une même expression :

```vb
' VB.NET — court-circuit aux deux niveaux
montant = If(commande IsNot Nothing AndAlso commande.EstValide,
             commande.Total,
             0D)
' AndAlso : commande.EstValide n'est testé que si commande <> Nothing
' If()    : commande.Total n'est lu que si la condition est vraie
```

**Pourquoi l'introduire (et avec quelle prudence).** Passer à `AndAlso`/`OrElse` et à `If()` rend possibles les **gardes contre `Nothing`** propres et évite les évaluations inutiles — c'est généralement un gain. Mais partout où la seconde branche (ou le second opérande) a un **effet de bord**, c'est un **changement de comportement** : on l'introduit donc **délibérément**, en **contexte logique uniquement**, et **sous protection des tests** — jamais par remplacement aveugle.

---

## 🎯 Synthèse — points de vigilance de la section

| Élément | Risque | Réflexe |
|---------|--------|---------|
| `If n Then` (n entier) | Erreur sous `Option Strict On` | `If n <> 0 Then` (idem objets/chaînes) |
| `If x = True` | `True` = `-1` en VB6 | Tester la condition, pas l'égalité à `True` (B.8) |
| `Select Case` sur chaîne | Casse selon `Option Compare` | Reporter `Option Compare` (6.4) |
| `IIf` laissé tel quel | Évalue **les deux** branches (division, effets de bord) | Auditer ; migrer vers `If()` |
| `IIf` → `If()` | **Changement de sémantique** (court-circuit) | Vérifier les effets de bord ; valider (golden master) |
| `Switch` / `Choose` | Évaluent **tous** les arguments | Même prudence qu'`IIf` |
| `If(...)` vs `If … Then` | Même mot-clé, sens différent | Lecture attentive (opérateur vs instruction) |

---

## 🔗 Pour aller plus loin

- Module **9.1** — Opérateurs logiques `And`/`Or` vs `AndAlso`/`OrElse` (l'autre moitié du court-circuit)
- Module **7.6** / Annexe **B.8** — Le booléen VB6 (`True` = `-1`) et les conversions
- Module **6.4** / **17.1** — `Option Compare` (casse des `Case`) et `Option Strict` (conditions booléennes)
- Module **9.5** — `On…GoTo` / `On…GoSub` supprimés → réécriture en `Select Case`
- Module **5.5** — Harnais de tests de référence (*golden master*) pour valider le passage `IIf` → `If()`
- Module **17.4** — Traquer les pièges silencieux après migration
- Annexe **A** — Tableau de correspondance VB6 → VB.NET

---

> **Section précédente** → [9.2 — Fonctions de chaînes](02-fonctions-chaines.md)
> **Section suivante** → [9.4 — Boucles (`For…Next`, `For Each`, `Do…Loop`, `While…Wend`)](04-boucles.md)

⏭️ [Boucles (`For…Next`, `For Each`, `Do…Loop`, `While…Wend` → `While…End While`)](/09-operateurs-controle/04-boucles.md)
