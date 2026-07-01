🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 🎯 5.2 — Réduire les pièges *en amont* (typage explicite, `ByVal` explicite côté VB6) ⭐

**Module 5 — Préparer le code VB6 avant la migration** · Section 5.2
Cible : .NET Framework 4.7.2

> 🛡️ **En une phrase** : tout ce que vous rendez **explicite en VB6** — le type d'une variable, le
> mode de passage d'un paramètre — **survit** à la migration et y verrouille le comportement voulu.
> C'est la section où l'on désamorce, à la source, les deux changements silencieux les plus
> dangereux.

---

## 🧭 De quoi parle cette section

La [section 5.1](01-nettoyer.md) a fait le **ménage** : `Option Explicit` partout, code mort  
supprimé, `Variant` *manifestement* inutiles retirés. La 5.2 va plus loin et de manière
**systématique** : passer chaque variable et chaque paramètre au crible pour figer, côté VB6, ce que
.NET changerait sinon en silence.

Deux piliers, ce sont exactement ceux du titre :

1. **Le typage explicite** — donner à *chaque* variable un type précis et correct, et repérer les
   cas où la **taille** comptera ;
2. **Le `ByVal`/`ByRef` explicite** — annoter *chaque* paramètre, parce que le mode de passage par
   défaut s'**inverse** entre VB6 et VB.NET.

> ⭐ **L'idée qui rend cette section décisive** : en VB6, ajouter ces annotations est presque
> toujours un **non-événement** — le comportement de l'application ne change pas. Mais une fois la
> migration faite, ce que vous aviez écrit explicitement **reste vrai**, alors qu'un paramètre ou un
> type laissé implicite aurait été *réinterprété* par le nouveau langage. On épingle le contrat
> pendant qu'on dispose encore de l'oracle vivant qu'est l'application VB6.

> ⚠️ **Filet de sécurité** : annoter `ByVal` ou resserrer un type **peut** modifier le comportement
> si le code reposait sur l'ancien (un paramètre réellement réécrit, par exemple). Le risque est
> ciblé, pas nul. Travaillez par **petites passes**, recompilez et **comparez au *golden master***
> ([§5.5](05-golden-master.md)) après chaque lot. C'est le réflexe de tout ce module.

---

## 1️⃣ Pilier n°1 — Le `ByVal`/`ByRef` explicite (le piège n°1)

### Le mécanisme : un défaut qui s'inverse

C'est le piège le plus cité de toute la formation, et pour une raison simple :

| | Mode de passage **par défaut** (paramètre non annoté) |
|---|---|
| **VB6** | **`ByRef`** (par référence) |
| **VB.NET** | **`ByVal`** (par valeur) |

Un paramètre **sans mot-clé** veut donc dire **l'exact opposé** dans les deux langages. Le code  
compile des deux côtés ; rien ne s'allume ; mais le comportement bascule.

```vb
' VB6 — paramètre non annoté = ByRef (par défaut)
Sub Incrementer(valeur As Long)
    valeur = valeur + 1        ' écrit DANS la variable de l'appelant
End Sub

Sub Appelant()
    Dim compteur As Long
    compteur = 10
    Incrementer compteur
    ' En VB6 : compteur vaut 11 (passage par référence)
    ' Si ce paramètre devient ByVal à la migration : compteur reste 10 — sans la moindre erreur
End Sub
```

### La préparation : annoter *chaque* paramètre, en connaissance de cause

Le geste consiste à parcourir toutes les signatures (`Sub`, `Function`, `Property`) et à écrire
**explicitement** `ByVal` ou `ByRef` devant chaque paramètre. Mais ce n'est pas mécanique : pour
chacun, il faut **trancher une question de contrat**.

| Le code appelé… | Intention réelle | À écrire en VB6 |
|---|---|---|
| modifie l'argument **et** l'appelant en dépend | passage par référence | **`ByRef`** (explicite) |
| lit seulement l'argument, ne réécrit rien que l'appelant attende | passage par valeur | **`ByVal`** |
| réassigne un objet (`Set param = New …`) que l'appelant doit récupérer | par référence | **`ByRef`** |
| appelle des méthodes / lit des propriétés d'un objet | la valeur suffit | **`ByVal`** (l'objet reste modifiable) |

> 💡 **Pourquoi ajouter `ByRef` ne « coûte » rien en VB6** : `ByRef` *étant déjà* le défaut, l'écrire
> explicitement ne change **rien** au comportement VB6. En revanche, écrire `ByVal` n'est sûr **que
> là où le code ne dépend pas** de la réécriture — et se forcer à le vérifier, paramètre par
> paramètre, **est tout l'intérêt** de l'exercice. Là où le doute subsiste, on garde `ByRef`
> explicite : c'est le choix conservateur, qui préserve le comportement actuel.

### Pourquoi le faire *maintenant*, et pas s'en remettre à l'outil

Un bon convertisseur (l'assistant de Visual Studio,
[§4.1](../04-outils-migration/01-upgrade-wizard.md)) ajoute en général `ByRef` explicite pour
préserver la sémantique. **Mais ne misez pas dessus** : une migration manuelle, un outil moins  
fidèle, ou un copier-coller assisté par IA peuvent laisser le paramètre glisser vers le `ByVal` par  
défaut de .NET. En annotant **vous-même en VB6**, le résultat devient **déterministe** quel que soit  
l'outil employé ensuite.

> 🔎 **Clarification utile — `ByVal` sur un objet ne fait pas de copie de l'objet.** En VB6 comme en
> .NET, `ByVal` sur une variable objet passe **la référence par valeur** : on peut toujours
> **modifier** l'objet (appeler ses méthodes, changer ses propriétés) ; ce qu'on ne peut pas, c'est
> **réassigner** la variable de l'appelant à un *autre* objet. Donc ne basculez en `ByVal` que les
> objets que le code appelé **ne réassigne pas**. Ce comportement est identique des deux côtés : pas
> de surprise à la migration sur ce point précis.

### ⚠️ Deux limites de périmètre à connaître

- **Les subtilités du *site d'appel*** (et non de la déclaration) relèvent du
  [module 10.2](../10-procedures-fonctions/02-byref-byval-audit.md). Exemple classique : en VB6, des
  **parenthèses autour d'un argument** forcent un passage par valeur — `Incrementer (compteur)`
  passe une copie même si le paramètre est `ByRef`. Ici, en 5.2, on fixe les **déclarations** ;
  l'audit fin des appels viendra plus tard.
- **`ParamArray`** mérite une vigilance à part : en VB6 il est passé **par référence**, mais devient
  un **tableau par valeur** en .NET. Repérez-le dès maintenant et notez-le pour
  [§10.4](../10-procedures-fonctions/04-paramarray.md).

> Détail complet du retournement et de ses conséquences :
> [§2.4 — ByRef → ByVal](../02-differences-fondamentales/04-byref-byval.md) et
> [Annexe B.1](../annexes/pieges-silencieux/README.md).

---

## 2️⃣ Pilier n°2 — Le typage explicite (et la question de la taille)

### Au-delà du `Variant` : pourquoi typer *précisément*

La 5.1 a retiré les `Variant` *évidemment* inutiles. La 5.2 généralise : **chaque** variable reçoit  
un type explicite **et** le bon. Deux bénéfices distincts :

- moins de `Variant` résiduels = moins de surface où les pièges `Empty`/`Null`/`Missing` peuvent
  surgir à la conversion (`Variant` → `Object`,
  [§7.1](../07-types-variables/01-variant-object.md)) ;
- un type explicite **enlève toute devinette** au convertisseur : il sait exactement quoi produire,
  de façon déterministe.

Mais il y a un piège propre aux entiers, qui rend le typage en 5.2 *stratégique* et pas seulement  
cosmétique.

### Le redimensionnement silencieux des entiers

Les types entiers **changent de taille** entre VB6 et VB.NET — et le code compile sans broncher :

| Type | En **VB6** | En **VB.NET** |
|---|---|---|
| `Integer` | **16 bits** (−32 768 … 32 767) | **32 bits** |
| `Long` | **32 bits** | **64 bits** |
| `Currency` | 64 bits à virgule fixe (4 décimales) | **pas d'équivalent** → `Decimal` |

Concrètement : une variable déclarée `As Integer` en VB6 (16 bits) devient, après migration, un  
entier **32 bits**. La plupart du temps c'est **bénin** (la plage s'élargit). Mais cela **mord** dans  
des cas précis, qu'il faut savoir repérer.

> 🧭 **Honnêteté technique — on ne « pré-redimensionne » pas en VB6.** VB6 ne possède pas les types
> aux tailles .NET (`Short` pour 16 bits, `Integer` pour 32 bits). « Typer explicitement en VB6 » ne
> veut donc **pas** dire forcer des tailles .NET dans VB6 — ce serait une confusion. Cela veut dire
> trois choses : (a) donner à chaque variable un type **correct pour VB6** ; (b) s'assurer qu'il
> **colle à l'usage réel** ; (c) **signaler** les endroits où la taille comptera. La *décision* de
> mapping (par ex. mapper le `Integer` VB6 vers `Short` pour conserver 16 bits) appartient au
> [module 7.2](../07-types-variables/02-entiers-redimensionnes.md) — mais c'est **ici**, pendant
> qu'on comprend le code, qu'on dresse la liste des candidats.

### Les cas où la taille compte (à signaler dès maintenant)

Le redimensionnement est sans conséquence… sauf là où le **nombre de bits** fait partie du contrat :

- **Interop COM et appels d'API Win32** : une structure ou un champ attendu sur 16 bits ne tolère
  pas un 32 bits (voir [§16.4–16.5](../16-interop-com-api/04-declare-vers-pinvoke.md)) ;
- **Enregistrements binaires de taille fixe**, sérialisation, `Get`/`Put` sur fichiers : la taille
  conditionne la disposition des octets ;
- **Opérations sur les bits** (`And`/`Or`/`Xor`, masques, décalages) : le résultat dépend de la
  largeur du type ;
- **Débordement volontairement exploité** (rare, mais on en voit) : un calcul qui « bouclait » à
  32 767 ne bouclera plus au même seuil.

```vb
' À signaler : un masque 16 bits passé à une API attendant un entier court
Dim drapeaux As Integer        ' 16 bits en VB6
drapeaux = lpAPI And &H7FFF&    ' la largeur 16 bits fait partie du contrat ici
' ⚑ candidat à un traitement explicite au module 7 (mapper vers Short ?)
```

> 🏷️ **Pratique recommandée** : marquez ces points d'un commentaire repérable (par ex. `' ⚑ TAILLE`)
> ou consignez-les dans le document de comportements critiques ([§5.4](04-documenter-comportements.md)).
> Ils deviendront la *checklist* de la validation ([§17.4](../17-valider-refactoriser/04-traquer-pieges.md)).

### Le cas `Currency` — la monnaie, piège classique

`Currency` (entier 64 bits à virgule fixe, 4 décimales) **n'a pas d'équivalent exact** en .NET et se
traduit en `Decimal`, dont la précision et le comportement d'arrondi diffèrent. Pour de la monnaie,  
c'est un piège silencieux répertorié ([Annexe B.9](../annexes/pieges-silencieux/README.md)).

La préparation en 5.2 : **identifier** non seulement les variables déjà typées `Currency`, mais  
aussi celles qui *devraient* l'être — typiquement des montants logés dans un `Double` ou un
`Variant` par commodité. Les ramener sur le bon type **en VB6**, c'est éviter qu'un `Double` flottant
ne s'installe durablement là où il fallait une valeur décimale exacte.

> Détail des conversions de types et de leurs règles :
> [§2.3](../02-differences-fondamentales/03-typage.md),
> [§7.2](../07-types-variables/02-entiers-redimensionnes.md) et
> [Annexe D](../annexes/correspondance-types/README.md).

---

## 🔁 Méthode de passage

Les deux piliers se traitent bien **module par module**, dans cet ordre pragmatique :

1. **Le typage d'abord, sur un module.** On resserre les types, on flague les cas sensibles à la
   taille et les montants. Recompiler, relancer, comparer au *golden master*.
2. **Le `ByVal`/`ByRef` ensuite, sur le même module.** On annote chaque signature en tranchant le
   contrat de chaque paramètre, en gardant `ByRef` explicite là où le doute persiste. Recompiler,
   relancer, comparer.
3. **Module suivant.** On répète, du plus simple (modules utilitaires) au plus sensible (cœur
   métier).

À chaque lot : **petites passes → recompilation → comparaison au comportement de référence.** Le
*golden master* est ce qui transforme « ça compile encore » en « ça se comporte encore pareil ».

---

## ⚠️ Les pièges de cette section

- **Mettre `ByVal` partout sans réfléchir.** Basculer tous les paramètres en `ByVal` « pour coller à
  .NET » casse, **dès VB6**, ceux qui réécrivaient réellement leur argument. Le geste est *par
  paramètre*, avec jugement — sinon on remplace un piège futur par un bug immédiat.
- **Croire qu'on « règle » le redimensionnement ici.** En 5.2 on **prépare et on signale** ; la
  *résolution* (choix du type cible, `Short`/`Integer`/`Decimal`) est l'affaire du
  [module 7](../07-types-variables/README.md). Ne tranchez pas le mapping ici.
- **Confondre déclaration et appel.** Les parenthèses qui forcent `ByVal`, le mot-clé `Call`, etc.,
  sont des subtilités de *site d'appel* — audit en [§10.2](../10-procedures-fonctions/02-byref-byval-audit.md),
  pas ici.
- **Travailler sans filet.** Resserrer un type ou annoter `ByVal` *peut* changer un comportement.
  Sans recompilation fréquente ni *golden master*, la régression se découvre trop tard et sans cause
  identifiable.
- **Vouloir tout perfectionner.** L'objectif reste un code **prêt à migrer**, pas un code idéal. On
  va de toute façon le transformer ; visez le bon niveau pour **vos** risques et **votre** budget.

---

## 🔗 Pour aller plus loin

- **Avant** : [5.1 — Nettoyer](01-nettoyer.md) (le ménage de base, dont les `Variant` manifestement
  inutiles).
- **Après** : [5.3 — Découpler l'UI de la logique métier](03-decoupler-ui-metier.md).
- **Filet de sécurité** : [5.5 — Harnais de tests de référence (*golden master*)](05-golden-master.md).
- **Le « pourquoi » des deux pièges** :
  [§2.4 — ByRef → ByVal](../02-differences-fondamentales/04-byref-byval.md),
  [§2.3 — Typage](../02-differences-fondamentales/03-typage.md).
- **Traitement détaillé (cœur de la migration)** :
  [§7.1 — Variant → Object](../07-types-variables/01-variant-object.md),
  [§7.2 — Entiers redimensionnés](../07-types-variables/02-entiers-redimensionnes.md),
  [§10.2 — Audit ByRef/ByVal](../10-procedures-fonctions/02-byref-byval-audit.md),
  [§10.4 — ParamArray](../10-procedures-fonctions/04-paramarray.md).
- **Aide-mémoire** : [Annexe B](../annexes/pieges-silencieux/README.md) (B.1 ByRef, B.2 entiers,
  B.9 dates/monnaie) et [Annexe D — Correspondance des types](../annexes/correspondance-types/README.md).
- **Côté .NET** : [§6.4 — Options de compilation](../06-preparer-environnement/04-options-compilation.md).

---

> ✅ **À retenir** : rendez **explicite en VB6** ce que .NET réinterpréterait en silence. Annotez
> chaque paramètre (`ByVal`/`ByRef`) en tranchant son contrat — `ByRef` explicite ne coûte rien,
> `ByVal` se vérifie au cas par cas. Typez chaque variable précisément, et **signalez** les endroits
> où la taille des entiers ou la précision de `Currency` fera la différence. Ces annotations, posées
> dans l'environnement stable de VB6, **traversent** la migration et y verrouillent le comportement
> voulu — c'est ce qui rend cette section la plus rentable du module.

---

**Section précédente** : [5.1 — Nettoyer le code VB6](01-nettoyer.md)  
**Section suivante** : [5.3 — Découpler l'UI de la logique métier →](03-decoupler-ui-metier.md)

⏭️ [Découpler l'UI de la logique métier (faciliter la migration progressive)](/05-preparer-code-vb6/03-decoupler-ui-metier.md)
