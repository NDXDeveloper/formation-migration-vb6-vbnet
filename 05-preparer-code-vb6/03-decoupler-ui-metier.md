🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 🧩 5.3 — Découpler l'UI de la logique métier (faciliter la migration progressive)

**Module 5 — Préparer le code VB6 avant la migration** · Section 5.3
Cible : .NET Framework 4.7.2

> 🧩 **En une phrase** : tant que vos règles de calcul vivent à l'intérieur des
> `Button_Click`, vous ne pourrez migrer ni les tester sans traîner avec elles tout le
> formulaire — c'est-à-dire le morceau le plus dur de la migration. Les **séparer en VB6**,
> c'est se rendre capable de migrer **par tranches**.

---

## 🧭 De quoi parle cette section

Les [sections 5.1](01-nettoyer.md) et [5.2](02-reduire-pieges-amont.md) ont nettoyé et sécurisé le  
code *ligne à ligne* (ménage, typage, `ByVal`). La 5.3 change d'échelle : elle touche à  
l'**organisation** du code. L'objectif n'est pas de réécrire l'application, mais de **déplacer la  
logique métier hors des formulaires**, là où elle pourra être migrée et testée indépendamment de  
l'interface.

C'est une préparation **structurelle**, et elle est au service d'un objectif très concret nommé dans  
le titre : rendre possible une **migration progressive** (par opposition au *big-bang* — voir
[§3.4](../03-evaluer-strategie/04-strategies.md)).

> ⭐ **L'idée maîtresse — l'asymétrie de risque.** Dans une migration VB6 → VB.NET, les deux moitiés
> du travail n'ont pas du tout le même profil :
>
> | | La **logique métier** | L'**interface** (formulaires) |
> |---|---|---|
> | Nature du travail | surtout **mécanique** (types, `ByRef`, `On Error` → `Try`) | le **vrai chantier** : Twips→pixels, *control arrays* disparus, GDI+, MDI, cycle de vie des formulaires… |
> | Testabilité | **élevée** : entrées → sorties comparables au *golden master* | **faible** : difficile à automatiser, validation surtout visuelle |
> | Renvoi | [Partie 3](../07-types-variables/README.md) | [Module 13](../13-formulaires-winforms/README.md) ⚠️ |
>
> Tant que les deux sont **enchevêtrés**, vous héritez du pire des deux mondes : impossible de
> migrer la partie facile et testable sans affronter en même temps la partie dure et peu testable.
> Le découplage **dissocie** ces deux risques pour pouvoir les traiter **séparément**.

---

## 🔎 Le symptôme : la logique dans les gestionnaires d'événements

Le code VB6 hérité loge très souvent ses règles métier **directement dans le code des formulaires**,  
au creux des gestionnaires d'événements (`Click`, `Change`, `LostFocus`…). Validation, calcul, accès  
base de données, tout y est mêlé à la manipulation des contrôles :

```vb
' Form1.frm — TOUT est dans le gestionnaire de bouton
Private Sub cmdValider_Click()
    ' --- lecture de l'UI ---
    Dim prix As Currency, quantite As Long
    prix = CCur(txtPrix.Text)
    quantite = CLng(txtQuantite.Text)

    ' --- validation (logique métier) ---
    If quantite <= 0 Then
        MsgBox "Quantité invalide"
        Exit Sub
    End If

    ' --- calcul (logique métier) ---
    Dim remise As Currency
    If quantite >= 10 Then remise = 0.05 Else remise = 0
    Dim total As Currency
    total = prix * quantite * (1 - remise)

    ' --- accès données (logique métier) ---
    cnx.Execute "INSERT INTO Commandes ... VALUES (" & total & ")"

    ' --- retour vers l'UI ---
    lblTotal.Caption = Format(total, "Currency")
End Sub
```

Le problème saute aux yeux une fois nommé : **la règle « remise de 5 % à partir de 10 unités » est  
prisonnière du formulaire**. On ne peut ni la migrer seule, ni la tester sans cliquer sur un bouton,  
ni la réutiliser ailleurs.

---

## 🛠️ Le geste : extraire la logique vers des modules / classes

La préparation consiste à **déplacer la logique métier** des `.frm` vers des unités séparées — un  
module standard (`.bas`) ou, mieux, une **classe** (`.cls`) — en laissant au formulaire un rôle
**mince** : lire les contrôles, appeler la logique, afficher le résultat.

```vb
' CalculCommande.cls — la logique métier, SANS aucune référence à l'UI
Public Function CalculerTotal(ByVal prix As Currency, _
                              ByVal quantite As Long) As Currency
    Dim remise As Currency
    If quantite >= 10 Then remise = 0.05 Else remise = 0
    CalculerTotal = prix * quantite * (1 - remise)
End Function

Public Function QuantiteEstValide(ByVal quantite As Long) As Boolean
    QuantiteEstValide = (quantite > 0)
End Function
```

```vb
' Form1.frm — devenu MINCE : il orchestre, il ne décide plus
Private Sub cmdValider_Click()
    Dim calc As New CalculCommande
    Dim prix As Currency, quantite As Long
    prix = CCur(txtPrix.Text)
    quantite = CLng(txtQuantite.Text)

    If Not calc.QuantiteEstValide(quantite) Then
        MsgBox "Quantité invalide"
        Exit Sub
    End If

    Dim total As Currency
    total = calc.CalculerTotal(prix, quantite)
    EnregistrerCommande total                 ' couche données, elle aussi extraite
    lblTotal.Caption = Format(total, "Currency")
End Sub
```

La règle métier est désormais dans `CalculCommande`, **indépendante de tout formulaire**. Notez au  
passage que les paramètres sont annotés `ByVal` et typés précisément : c'est la
[section 5.2](02-reduire-pieges-amont.md) qui paye ici — une frontière propre est aussi une
frontière **explicite**.

### La règle de partage : qu'est-ce qui va où ?

| Reste dans le **formulaire** (`.frm`) | Part vers **module / classe** (`.bas` / `.cls`) |
|---|---|
| Lire et écrire les contrôles (`.Text`, `.Caption`…) | Règles de calcul et de validation |
| Réagir aux événements UI (clic, saisie, focus) | Accès aux données (requêtes, transactions) |
| Afficher messages, dialogues, mises en forme | Transformations, formats métier, algorithmes |
| Enchaîner l'appel à la logique | Tout ce qui ne « connaît » aucun contrôle |

> 💡 **Le test décisif** : une unité de logique bien découplée **ne mentionne aucun contrôle**. Si
> elle référence `txtPrix`, `lblTotal` ou `Me.`, c'est qu'elle est encore soudée à l'UI. Le but est
> qu'elle puisse être appelée **sans qu'aucun formulaire ne soit chargé**.

---

## 🎯 Ce que le découplage débloque pour la migration

### 1. Migrer la partie facile *en premier* et la valider

Une fois la logique sortie des formulaires, elle devient **migrable et testable seule**. On peut la  
porter en VB.NET, puis vérifier `CalculerTotal(100, 12)` → résultat attendu, **sans toucher aux  
formulaires ni aux contrôles**. C'est exactement la matière que le *golden master*
([§5.5](05-golden-master.md)) sait arbitrer : entrées connues, sorties comparables. La partie
difficile — l'UI — est **isolée** et traitée à son rythme (module 13).

### 2. Migrer par tranches (la migration progressive)

Le découplage transforme un bloc monolithique en **unités à frontières nettes**. On peut alors  
migrer **module par module**, valider chaque tranche, et avancer sans devoir tout basculer d'un coup.  
C'est la condition pratique d'une stratégie **incrémentale**
([§3.4](../03-evaluer-strategie/04-strategies.md)) plutôt que d'un *big-bang* risqué.

### 3. Rendre la cohabitation VB6 ↔ .NET réaliste

C'est le bénéfice le plus puissant. Une logique propre, regroupée dans des classes **sans  
dépendance à l'UI**, offre des **« coutures »** où brancher l'interopérabilité COM. Pendant la  
transition, formulaires VB6 et logique .NET (ou l'inverse) peuvent **coexister** de part et d'autre  
de ces frontières :

- migrer d'abord la logique en .NET et l'exposer au VB6 restant via **CCW**
  ([§16.2](../16-interop-com-api/02-ccw.md)) ;
- ou migrer d'abord des formulaires en .NET qui consomment la logique VB6 encore en place via
  **RCW** ([§16.1](../16-interop-com-api/01-rcw.md)).

> 🔗 Dans les deux cas, **la qualité de la couture dépend du découplage fait ici**. Une logique
> empêtrée dans les formulaires n'offre aucune frontière exploitable. Le « comment » de cette
> cohabitation appartient au [module 16](../16-interop-com-api/README.md) et à la stratégie
> [§3.5](../03-evaluer-strategie/05-cohabitation-com.md) ; ce qu'on prépare **ici**, ce sont les
> frontières qui la rendront possible.

### 4. Profiter d'une frontière naturelle pour tester

Une logique appelable sans formulaire est, par construction, une logique qu'on peut **piloter par du  
code**. C'est ce qui permettra, plus tard, d'écrire de vrais tests de non-régression
([§17.3](../17-valider-refactoriser/03-non-regression.md)) sur la partie métier — chose impossible
tant qu'il faut un clic humain pour déclencher chaque calcul.

---

## 🧭 Jusqu'où aller — la lucidité d'abord

Cette section peut donner envie de tout réarchitecturer proprement (MVP, MVVM, couches en bonne et  
due forme). **Résistez.** L'objectif est de **faciliter la migration**, pas de livrer une  
architecture idéale dans un langage qu'on s'apprête à quitter.

| ✅ Découplage *utile* (à faire) | ❌ Sur-ingénierie (à éviter ici) |
|---|---|
| Sortir les règles de calcul et de validation des `.frm` | Refondre toute l'application en architecture en couches |
| Regrouper la logique en `.bas` / `.cls` appelables sans UI | Introduire injection de dépendances, abstractions élaborées |
| Couper les calculs métier de la manipulation des contrôles | Viser une séparation « parfaite » sur 100 % du code |
| Créer des frontières là où une **couture COM** sera utile | Réécrire ce qu'on va de toute façon transformer |

> 🎯 **Le bon dosage.** Concentrez l'effort là où il **rapporte** : le cœur métier réutilisé, les
> calculs sensibles ([§5.4](04-documenter-comportements.md)), et les frontières qui serviront à la
> cohabitation. Un écran purement informatif, sans logique, n'a rien à gagner à être disséqué — il
> sera migré tel quel au module 13.

> ⚠️ **Réalité du *legacy* VB6.** L'IDE VB6 n'aide pas à ce genre de refactoring (pas de « extraire
> une méthode » automatique). Le déplacement est **manuel**, donc faillible : avancez par **petites
> passes**, recompilez, **comparez au comportement de référence** après chaque extraction. Déplacer
> du code, c'est risquer de le casser — d'où l'importance du filet.

---

## ⚠️ Les pièges de cette section

- **Déplacer la logique *et* la transformer en même temps.** L'extraction doit être à **iso-
  comportement** : on *bouge* le code, on ne le *réécrit* pas. Mélanger les deux rend toute
  régression impossible à attribuer. Si une amélioration s'impose, faites-la **après**, dans une
  passe distincte.
- **Laisser des références à l'UI fuiter dans la logique.** Une classe « métier » qui lit encore
  `txtPrix.Text` n'est pas découplée. Faites passer les valeurs **en paramètres** (typés, `ByVal` —
  cf. 5.2), ne laissez pas la logique aller chercher les contrôles.
- **Vouloir tout découpler.** Le découplage total n'est pas l'objectif ; la *facilitation de la
  migration* l'est. Au-delà, le coût dépasse le bénéfice.
- **Oublier le code derrière les événements moins visibles.** La logique ne se cache pas que dans
  les `Click` : `Form_Load`, `Change`, `Timer`, `LostFocus` en contiennent souvent aussi. L'ordre
  d'enchaînement de ces événements est lui-même un comportement à préserver — à consigner en
  [§5.4](04-documenter-comportements.md).
- **Découpler sans filet.** Sans recompilation fréquente ni *golden master*
  ([§5.5](05-golden-master.md)), une extraction ratée se découvre trop tard.

---

## 🔗 Pour aller plus loin

- **Avant** : [5.1 — Nettoyer](01-nettoyer.md) · [5.2 — Réduire les pièges en amont](02-reduire-pieges-amont.md)
  (les frontières propres sont aussi des frontières **typées et explicites**).
- **Après** : [5.4 — Documenter les comportements critiques](04-documenter-comportements.md) (dont
  l'**ordre des événements**, fragile au découplage).
- **Filet de sécurité** : [5.5 — Harnais de tests de référence (*golden master*)](05-golden-master.md).
- **Stratégie de migration** : [§3.4 — Big-bang vs incrémentale](../03-evaluer-strategie/04-strategies.md)
  · [§3.5 — Cohabitation COM](../03-evaluer-strategie/05-cohabitation-com.md).
- **La cohabitation, en pratique** : [Module 16 — Interop COM](../16-interop-com-api/README.md)
  ([RCW §16.1](../16-interop-com-api/01-rcw.md), [CCW §16.2](../16-interop-com-api/02-ccw.md)).
- **La partie difficile — l'UI** : [Module 13 — Formulaires → Windows Forms](../13-formulaires-winforms/README.md)
  (et les *control arrays* disparus, [§13.6](../13-formulaires-winforms/06-control-arrays.md)).
- **Tester la logique après migration** : [§17.3 — Non-régression](../17-valider-refactoriser/03-non-regression.md).

---

> ✅ **À retenir** : la logique métier enfermée dans les formulaires est l'**ennemie de la migration
> progressive**. En la déplaçant **en VB6** vers des modules et classes appelables **sans aucun
> formulaire**, vous séparez la partie *facile et testable* (le métier) de la partie *dure et
> risquée* (l'UI), vous ouvrez la voie à une migration **par tranches**, et vous créez les coutures
> qui rendront la **cohabitation COM** réaliste. Le tout à **iso-comportement**, par petites passes,
> sous la surveillance du *golden master* — et sans céder à la tentation de tout réarchitecturer.

---

**Section précédente** : [5.2 — Réduire les pièges en amont](02-reduire-pieges-amont.md)  
**Section suivante** : [5.4 — Documenter les comportements critiques →](04-documenter-comportements.md)

⏭️ [Documenter les comportements critiques (arrondis, dates, ordre d'événements)](/05-preparer-code-vb6/04-documenter-comportements.md)
