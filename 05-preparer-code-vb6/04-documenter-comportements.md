🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 📝 5.4 — Documenter les comportements critiques (arrondis, dates, ordre d'événements)

**Module 5 — Préparer le code VB6 avant la migration** · Section 5.4
Cible : .NET Framework 4.7.2

> 📝 **En une phrase** : votre application VB6 qui tourne **est** la spécification — mais sa façon
> exacte d'arrondir un montant, de calculer sur une date ou d'enchaîner ses événements n'est écrite
> **nulle part**. Cette section consiste à **mettre ce comportement noir sur blanc** pendant que
> l'oracle vivant est encore là pour répondre.

---

## 🧭 De quoi parle cette section

Les sections précédentes ont préparé le code lui-même : [ménage (5.1)](01-nettoyer.md),
[typage et `ByVal` (5.2)](02-reduire-pieges-amont.md), [découplage (5.3)](03-decoupler-ui-metier.md).
La 5.4 ne touche pas au code : elle **capture du savoir**. Précisément, elle transforme en référence
**écrite** les comportements que .NET risque de reproduire *différemment, et en silence*.

Le SOMMAIRE en nomme trois, et ce sont eux qui structurent cette section :

1. **Les arrondis** (et la monnaie) ;
2. **Les dates** ;
3. **L'ordre des événements**.

Pourquoi ces trois-là ? Parce qu'ils partagent une caractéristique redoutable : **le code compile à  
l'identique des deux côtés, mais le résultat dérive.** Ce sont des pièges *silencieux* typiques
(catalogue complet en [Annexe B](../annexes/pieges-silencieux/README.md)) — et, contrairement à un
`ByRef` qu'on peut *corriger*, ceux-ci doivent d'abord être **documentés** pour qu'on sache, après
migration, ce que « correct » voulait dire.

> ⭐ **La distinction clé avec la [section 5.5](05-golden-master.md).** Ne confondez pas *documenter*
> et *capturer un golden master* — ils sont complémentaires, pas redondants :
>
> | | **5.4 — Documenter** (cette section) | **5.5 — Golden master** |
> |---|---|---|
> | Forme | **prose / tableaux** lisibles par un humain | **captures exécutables** comparées automatiquement |
> | Répond à | « **quel** est le comportement attendu, et **pourquoi** ? » | « le comportement a-t-il changé ? » |
> | Public | l'équipe, le validateur ([§17.4](../17-valider-refactoriser/04-traquer-pieges.md)), l'IA ([module 19](../19-migration-ia/README.md)) | la machine, en non-régression |
> | Force | explique l'**intention**, couvre l'invisible (ordre, timing) | détecte la **dérive** sur des sorties mesurables |
>
> Le golden master vous dit *qu'*une valeur a changé ; la documentation vous dit *quelle* valeur
> était la bonne et *pourquoi*. Les deux se nourrissent l'un l'autre — on y revient plus bas.

---

## 🔥 Pourquoi documenter *maintenant* (et pas après)

### 1. L'oracle est encore vivant — interrogez-le

Tant que l'application VB6 tourne, vous pouvez l'**observer**, la **sonder**, l'**expérimenter** :  
quel résultat produit-elle sur un montant à `,xx5` ? sur un calcul d'échéance à cheval sur un  
changement d'heure ? dans quel ordre ses événements se déclenchent-ils ? Après la bascule, cet oracle
**disparaît** : il ne reste que le code migré, et plus aucune façon de savoir ce que l'original
*faisait vraiment*.

### 2. Ces comportements sont **tacites**

Ils ne vivent ni dans une spec ni dans des commentaires : ils vivent dans le **runtime VB6** et,  
parfois, dans la tête des développeurs d'origine — souvent partis. Sans capture écrite, ce savoir  
s'évapore. Documenter, c'est **extérioriser** ce qui n'existait que dans l'exécution.

### 3. Le découplage (5.3) peut *lui-même* perturber l'ordre des événements

Déplacer de la logique hors des formulaires peut, par effet de bord, **modifier l'enchaînement** des
événements. D'où une règle d'ordre : **documentez l'ordre des événements *avant* de refactorer**,
pour disposer de la référence contre laquelle vérifier que le découplage n'a rien déplacé.

> 🧭 **Sans référence écrite, on ne sait pas distinguer un *bug* d'un *« différent mais acceptable »*.**
> C'est le vrai enjeu : après migration, face à un montant qui change d'un centime ou à un événement
> qui ne se déclenche plus, la seule question qui compte est « était-ce censé faire ça ? ». La 5.4
> est ce qui permet d'y répondre.

---

## 1️⃣ Les arrondis (et la monnaie)

### Le risque

L'arrondi est un nid à surprises silencieuses, pour deux raisons distinctes :

- **La règle d'arrondi sur les valeurs « pile au milieu ».** VB6 applique sur ses conversions
  (`CInt`, `CLng`) l'**arrondi au pair** (dit *« bancaire »* : `0,5` est arrondi vers le pair le plus
  proche), ce qui surprend qui attend un arrondi arithmétique. `CInt(2.5)` donne **2**, pas 3 ;
  `CInt(3.5)` donne **4**. .NET a ses propres règles par défaut, *souvent* alignées sur ce point —
  mais « souvent » n'est pas « toujours », et les fonctions employées (`Format`, arrondi maison,
  `Math.Round` avec un mode explicite) peuvent diverger sur les cas limites.
- **`Currency` → `Decimal`.** `Currency` est un entier 64 bits à virgule fixe (**exactement 4
  décimales**) ; `Decimal` a une précision et un suivi d'échelle **différents**. Sur une **somme de
  nombreux montants**, un calcul de TVA ou une remise en chaîne, le dernier centime peut bouger. Et
  si la monnaie a été logée dans un `Double` par commodité (une faute que la
  [5.2](02-reduire-pieges-amont.md) cherchait à corriger), l'arrondi devient carrément instable.

```vb
' VB6 — l'arrondi "bancaire", la surprise classique à consigner
Debug.Print CInt(2.5)    ' -> 2  (pas 3 !)
Debug.Print CInt(3.5)    ' -> 4
' À documenter : "ici l'app s'appuie sur l'arrondi au pair" — ou au contraire
' "l'app attend un arrondi arithmétique et passe par TelleFonction()"
```

### Ce qu'on documente ici

Pas la *correction* (elle relève de [§7.2](../07-types-variables/02-entiers-redimensionnes.md) et de  
l'[Annexe B.9](../annexes/pieges-silencieux/README.md)), mais le **comportement de référence** :

- **quelle règle d'arrondi** l'application produit réellement, vérifiée sur des **valeurs
  représentatives** — surtout les cas « pile au milieu » (`,5`, `,xx5`) et les négatifs ;
- les **points de calcul monétaire sensibles** : totaux, sous-totaux, TVA, remises en cascade,
  conversions de devises — avec, pour chacun, **entrées connues → résultat attendu** ;
- le **nombre de décimales** significatif et le moment où l'arrondi est censé intervenir (à chaque
  étape ? seulement à l'affichage ?).

> 🏷️ Ces couples *entrée → sortie* attendus ne sont pas perdus : ils deviennent directement des **cas
> du golden master** ([§5.5](05-golden-master.md)). Documenter la monnaie, c'est déjà écrire la
> moitié de son test.

---

## 2️⃣ Les dates

### Le risque

C'est l'un des changements de modèle les plus profonds. En VB6, **une `Date` *est* un `Double`** (le  
SOMMAIRE le rappelle en [§7.3](../07-types-variables/03-dates.md)) : un nombre de jours depuis une
**époque** de référence, la partie fractionnaire codant l'heure. Conséquence directe : **l'arithmétique
numérique sur les dates « marche »**.

```vb
' VB6 — la date est un Double : on calcule dessus comme sur un nombre
Dim echeance As Date
echeance = Now + 30              ' "+ 30" = + 30 jours
Dim delai As Long
delai = dateFacture - datePaiement   ' une soustraction = un nombre de jours
If demiJournee Then rdv = rdv + 0.5  ' "+ 0,5" = + 12 heures
```

En .NET, `DateTime` est un **type distinct** : `Now + 30` ne veut plus rien dire de la même façon
(il faut `.AddDays(30)`, `.Subtract(...)`, des `TimeSpan`). Tout code qui traite les dates **comme des
nombres** se comporte différemment — ou ne compile plus. S'y ajoutent :

- l'**époque** : du code qui lit la valeur numérique d'une date (`CDbl(uneDate)`, stockage en base
  d'un sérial) dépend du point zéro VB6 ;
- le **formatage et l'analyse** (`Format`, `CDate`) sont **sensibles à la culture** (séparateurs,
  ordre jour/mois) — c'est le piège [B.10](../annexes/pieges-silencieux/README.md) — et **les jetons
  de format diffèrent** entre VB6 et .NET ;
- les **dates « vides »**, l'année sur deux chiffres, les bornes min/max.

### Ce qu'on documente ici

- les endroits où l'application fait de l'**arithmétique de date** et **ce que chaque opération
  signifie** (« `+ 30` = 30 jours », « la soustraction rend un nombre de jours ») ;
- toute dépendance à la **valeur numérique** d'une date (et donc à l'**époque**) ;
- les **formats** que l'application **produit** et ceux qu'elle **analyse**, avec les **paramètres
  régionaux** supposés (lié à [§5.2](02-reduire-pieges-amont.md) pour les types, et à
  [B.10](../annexes/pieges-silencieux/README.md) pour la culture) ;
- la gestion des **dates vides / nulles** et des cas limites.

> Le traitement complet (conversions `Double` ↔ `DateTime`, jetons de format, époque) est en
> [§7.3](../07-types-variables/03-dates.md). Ici, on **capture l'intention**, pas la solution.

---

## 3️⃣ L'ordre des événements

### Le risque

VB6 et Windows Forms déclenchent leurs événements dans un **ordre différent**, et certains événements  
n'ont **pas d'équivalent direct**. Le cycle de vie d'un formulaire change
([§13.2](../13-formulaires-winforms/02-cycle-de-vie.md) : `Load`/`Unload`/`Activate` →
`Load`/`FormClosing`/`Shown`), tout comme la gestion du focus
(`GotFocus`/`LostFocus` → `Enter`/`Leave`/`Validating`/`Validated`) et l'enchaînement des `Change`.

Le piège n'est pas l'événement isolé, mais la **dépendance d'ordre** : du code qui suppose qu'un
événement se déclenche **avant** un autre.

```vb
' VB6 — dépendance d'ordre fragile : Load initialise une valeur que Change exploite
Private Sub Form_Load()
    txtTaux.Text = "5"        ' ceci déclenche-t-il txtTaux_Change ? et quand ?
End Sub

Private Sub txtTaux_Change()
    Recalculer                ' s'appuie sur un état supposé déjà initialisé
End Sub
```

Si, après migration, l'ordre diffère — `Change` qui se déclenche plus tôt, plus tard, ou pas du tout  
au chargement — le résultat dérive **sans aucune erreur**.

### Pourquoi c'est *ici* que la documentation est la plus précieuse

L'ordre des événements est **le comportement le plus difficile à capturer par un golden master** :  
c'est une affaire de **séquencement**, souvent **invisible dans les sorties** finales. Deux exécutions  
peuvent produire le même résultat affiché tout en ayant suivi des chemins différents — jusqu'au jour  
où la différence d'ordre se traduit en bug. **D'où la primauté de la documentation écrite sur ce  
point précis** : c'est parfois le *seul* moyen de fixer la référence.

### Ce qu'on documente ici

- les **séquences d'événements observées** aux moments sensibles : ouverture de formulaire,
  changement de focus, validation de saisie, fermeture ;
- les **dépendances d'ordre dont le code dépend** : « tel calcul suppose que tel champ a déjà été
  initialisé par tel événement » ;
- les interactions avec **`DoEvents`** (réentrance) et les **`Timer`**, dont l'ordre relatif est
  particulièrement piégeux ([§13.9](../13-formulaires-winforms/09-msgbox-doevents.md)) ;
- les enchaînements que le **découplage (5.3)** a pu déplacer — d'où l'intérêt d'avoir capturé la
  référence **avant** de refactorer.

> Le modèle d'événements lui-même (`Event`/`RaiseEvent`/`WithEvents`/`Handles`) est traité en
> [§12.8](../12-poo/08-evenements.md) et le cycle de vie des formulaires en
> [§13.2](../13-formulaires-winforms/02-cycle-de-vie.md). En 5.4, on note **ce qui se passe et dans
> quel ordre**, pas comment le réimplémenter.

---

## 🛠️ Comment documenter — formes utiles

La documentation n'a de valeur que si elle est **retrouvable** au bon moment. Trois supports  
complémentaires :

| Support | Usage | Atout |
|---|---|---|
| **Document de référence des comportements** (un tableau central) | recense arrondis, dates, séquences d'événements critiques | vue d'ensemble, transférable à l'équipe ([§18.4](../18-deploiement-bascule/04-documentation-transfert.md)) |
| **Commentaires balisés *au site*** (p. ex. `' ⚑ COMPORTEMENT : arrondi bancaire attendu`) | accrochent la note **là où le piège vit**, dans le code | impossible à rater pendant la migration de ce bout de code |
| **Tableaux *entrée → sortie attendue*** | figent le résultat de référence sur des cas représentatifs | **deviennent les cas du golden master** ([§5.5](05-golden-master.md)) |

> 🏷️ Réutilisez la **convention de balises** amorcée en [5.2](02-reduire-pieges-amont.md)
> (`' ⚑ TAILLE` pour les entiers redimensionnés). En ajoutant `' ⚑ COMPORTEMENT`,
> `' ⚑ DATE`, `' ⚑ ORDRE`, vous construisez une **checklist grepable** que la validation
> ([§17.4](../17-valider-refactoriser/04-traquer-pieges.md)) n'aura plus qu'à parcourir.

> 🤖 **Bénéfice pour la suite assistée par IA.** Une intention écrite explicitement (« cet arrondi
> doit rester bancaire », « cette date supporte une arithmétique en jours ») est exactement ce qu'on
> fournit à l'IA pour qu'elle **convertisse juste** et **n'invente pas** un faux équivalent
> ([§19.3](../19-migration-ia/03-detecter-pieges.md), [§19.5](../19-migration-ia/05-limites-pieges.md)).
> Un comportement non documenté est un comportement que l'IA — comme l'humain — devra *deviner*.

---

## 🔗 Documentation (5.4) et golden master (5.5) : un duo

Les deux sections forment un cycle vertueux ; mieux vaut les voir ensemble :

- la **documentation** fournit la **spec** : *quels* cas comptent, *quel* est le bon résultat,
  *pourquoi* ;
- ces cas alimentent le **golden master**, qui **automatise** la comparaison après chaque tranche de
  migration ;
- en retour, quand le golden master signale un écart, la **documentation explique** s'il s'agit d'une
  régression ou d'une différence acceptable ;
- et là où le golden master est **aveugle** (ordre, timing, réentrance), la **documentation reste le
  seul garde-fou**.

> ⭐ Autrement dit : la 5.4 répond à *« quoi et pourquoi »*, la 5.5 répond à *« est-ce que ça a
> bougé »*. Ni l'une ni l'autre seule ne suffit sur les comportements critiques.

---

## 🧭 Jusqu'où aller — la lucidité, encore

Documenter a un **coût** et la documentation **vieillit**. L'objectif n'est pas de décrire toute  
l'application, mais de capturer ce qui est **fragile *et* porteur** — ce dont une dérive aurait des  
conséquences réelles.

| ✅ À documenter en priorité | ❌ Effort à ne pas gaspiller |
|---|---|
| Calculs **monétaires** et règles d'arrondi | Logique triviale, sans cas limite |
| **Arithmétique de dates** et formats/régional | Écrans purement informatifs, sans comportement subtil |
| **Séquences d'événements** dont du code dépend | Ordre d'événements sans aucune dépendance |
| Cas **limites** (`,5`, dates vides, focus) | Re-décrire ce que le code dit déjà clairement |

> 🎯 **Le bon dosage** : suivez les **points sensibles** déjà repérés en
> [5.2](02-reduire-pieges-amont.md) (les balises `' ⚑ TAILLE`, la monnaie) et les frontières créées
> en [5.3](03-decoupler-ui-metier.md). C'est là que se concentrent les comportements à fixer. Le
> reste se vérifiera par le golden master sans nécessiter de prose dédiée.

---

## ⚠️ Les pièges de cette section

- **Documenter ce qu'on *croit*, pas ce que l'application *fait*.** La référence doit venir de
  l'**observation** de l'oracle vivant, pas d'un souvenir ou d'une intention supposée. Sur les
  arrondis et l'ordre des événements en particulier, l'intuition trompe souvent — **vérifiez en
  exécutant**.
- **Reporter l'ordre des événements à *après* le découplage.** Le refactoring de la
  [5.3](03-decoupler-ui-metier.md) peut **déplacer** les enchaînements. La référence d'ordre se
  capture **avant**, sinon on documente déjà un comportement altéré.
- **Compter sur le seul golden master pour l'ordre et le timing.** Ces comportements sont souvent
  **invisibles dans les sorties** : sans documentation écrite, leur dérive passe inaperçue jusqu'en
  production.
- **Oublier les paramètres régionaux.** Un format de date ou un séparateur décimal supposé est un
  comportement à part entière ([B.10](../annexes/pieges-silencieux/README.md)) — souvent tacite, donc
  facile à omettre.
- **Tout documenter.** La documentation exhaustive ne sera ni tenue à jour, ni lue. Visez le
  fragile et le porteur ; laissez le golden master couvrir le reste.

---

## 🔗 Pour aller plus loin

- **Avant** : [5.2 — Réduire les pièges en amont](02-reduire-pieges-amont.md) (les balises
  `' ⚑ TAILLE`, la monnaie) · [5.3 — Découpler l'UI](03-decoupler-ui-metier.md) (qui peut déplacer
  l'ordre des événements).
- **Après / duo** : [5.5 — Harnais de tests de référence (*golden master*)](05-golden-master.md) —
  l'exécutable qui consomme les cas documentés ici.
- **Traitement technique des sujets** :
  [§7.2 — Entiers redimensionnés & `Currency`](../07-types-variables/02-entiers-redimensionnes.md),
  [§7.3 — Dates (`Double` → `DateTime`)](../07-types-variables/03-dates.md),
  [§12.8 — Événements](../12-poo/08-evenements.md),
  [§13.2 — Cycle de vie des formulaires](../13-formulaires-winforms/02-cycle-de-vie.md),
  [§13.9 — `DoEvents`](../13-formulaires-winforms/09-msgbox-doevents.md).
- **Catalogue des pièges** : [Annexe B](../annexes/pieges-silencieux/README.md) — **B.9** (dates &
  arrondis monétaires), **B.10** (conversions sensibles à la culture).
- **Côté IA** : [§19.3 — Détecter les pièges](../19-migration-ia/03-detecter-pieges.md),
  [§19.5 — Limites & faux équivalents](../19-migration-ia/05-limites-pieges.md).
- **Côté validation** : [§17.4 — Traquer les pièges silencieux](../17-valider-refactoriser/04-traquer-pieges.md).

---

> ✅ **À retenir** : l'application VB6 connaît, dans son exécution, des comportements **tacites** —
> *comment* elle arrondit, *comment* elle calcule sur les dates, *dans quel ordre* ses événements se
> déclenchent — que .NET reproduira peut-être **autrement, et en silence**. Pendant que l'oracle est
> encore vivant, **transformez ce savoir en référence écrite** : règle d'arrondi et cas monétaires,
> sémantique d'arithmétique des dates et hypothèses régionales, séquences d'événements dont du code
> dépend. Cette documentation répond à *« quoi et pourquoi »* là où le golden master ne répond qu'à
> *« est-ce que ça a bougé »* — et elle reste le **seul garde-fou** sur l'ordre et le timing, que la
> machine ne sait pas voir.

---

**Section précédente** : [5.3 — Découpler l'UI de la logique métier](03-decoupler-ui-metier.md)  
**Section suivante** : [5.5 — Mettre en place un harnais de tests de référence (*golden master*) →](05-golden-master.md)

⏭️ [Mettre en place un **harnais de tests de référence** (*golden master*)](/05-preparer-code-vb6/05-golden-master.md)
