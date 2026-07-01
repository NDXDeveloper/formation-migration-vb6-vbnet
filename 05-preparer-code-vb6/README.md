🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 🏗️ Module 5 — Préparer le code VB6 avant la migration

**Partie 2 — Préparer le terrain (avant de migrer)** · Cible : .NET Framework 4.7.2

> 🩺 **L'image qui résume ce module** : avant une opération, on stabilise le patient, on
> documente son état et on réduit les complications prévisibles. La migration, c'est pareil.
> Le meilleur moment pour neutraliser la plupart des pièges, c'est *tant que le code VB6
> compile encore, tourne encore, et que son comportement de référence est sous la main*.

---

## 🧭 L'idée maîtresse

Une grande partie du travail de migration ne se fait **pas** en VB.NET : elle se fait en amont,  
dans le code VB6 lui-même.

C'est contre-intuitif, mais décisif. Tant qu'on est encore en VB6, l'application tourne — on  
dispose donc d'un *oracle* vivant, c'est-à-dire d'une référence fiable de ce qu'est le comportement
« correct ». Chaque correction y est petite, isolée et testable : on recompile, on relance, on
vérifie. Et chaque construction piégeuse qu'on désamorce avant la bascule réduit d'autant le nombre  
de changements simultanés au moment de migrer — exactement la logique du « pont » 4.7.2 (voir
[module 1.2](../01-cadrage-migration/02-pourquoi-472.md)).

Autrement dit : **chaque piège qu'on désamorce en VB6 est un piège qui ne pourra pas régresser  
silencieusement en VB.NET.**

> ⚠️ Le danger n°1 d'une migration n'est pas le code qui ne compile pas — c'est le code qui
> compile mais **ne se comporte plus pareil** (ByRef devenu ByVal, entier redimensionné,
> finalisation perdue…). Ce module attaque ce danger *avant* qu'il n'apparaisse. Le catalogue
> complet de ces pièges est en **[Annexe B](../annexes/pieges-silencieux/README.md)**.

---

## 🎯 Objectifs du module

À la fin de ce module, vous saurez :

- assainir un projet VB6 pour qu'il soit **prêt à être migré** (et pas seulement « qui marche ») ;
- **déplacer le risque en amont** : rendre explicites, côté VB6, les comportements que .NET
  changerait sinon en silence (typage, `ByVal`) ;
- **découpler** la logique métier de l'interface, condition d'une migration progressive ;
- **documenter** les comportements fragiles (arrondis, dates, ordre des événements) ;
- mettre en place un **harnais de tests de référence** (*golden master*) qui servira d'arbitre tout
  au long de la migration et lors de la validation finale (module 17) ;
- **geler le périmètre** pour ne pas migrer une cible mouvante.

---

## Pourquoi préparer *avant* de basculer

### 1. On a encore un oracle

L'application VB6 en production **est** la spécification. Elle définit, au comportement près, ce que
« correct » veut dire. Une fois la migration commencée, cet oracle devient plus difficile à
interroger. D'où deux réflexes complémentaires : capturer ce comportement de référence (le *golden  
master*, §5.5) et réduire l'écart qu'il faudra franchir.

### 2. La boucle de rétroaction est courte

Corriger un `Variant` superflu, expliciter un `ByVal`, sortir une règle de calcul d'un gestionnaire  
de bouton : en VB6, chacune de ces opérations se vérifie immédiatement contre une base connue et  
stable. En VB.NET, la même correction se mélangerait à des centaines d'autres changements —  
impossible de savoir lequel a cassé quoi.

### 3. On migre moins de complexité

Plus il reste de constructions piégeuses au moment du basculement (`Variant` implicites, sémantique
`ByRef` par défaut, propriétés par défaut, tableaux de contrôles soudés à l'UI), plus le
convertisseur — automatique ou humain — a d'occasions de se tromper en silence. Nettoyer en amont,  
c'est offrir à l'outil (et à l'IA, voir module 19) un terrain plus régulier.

---

## Le principe directeur : déplacer le risque en amont

| Piège typique à la migration | Préparation possible **en VB6** | Section |
|------------------------------|----------------------------------|---------|
| `ByRef` par défaut devient `ByVal` par défaut | Rendre tous les passages explicites (`ByVal`/`ByRef`) | §5.2 · [2.4](../02-differences-fondamentales/04-byref-byval.md) |
| `Variant` aux états flous (`Empty`/`Null`/`Missing`) | Typer explicitement, retirer le `Variant` superflu | §5.1 · §5.2 |
| Logique métier piégée dans le code des formulaires | Découpler UI ↔ métier | §5.3 |
| Arrondis monétaires, dates, ordre d'événements qui dérivent | Documenter le comportement attendu | §5.4 |
| « Ça compile, donc c'est bon » | Comparer à un *golden master* | §5.5 |
| Régressions noyées dans de nouvelles fonctionnalités | Geler le périmètre | §5.6 |

> ⭐ Retenez la formule : **« shift left »**. Un problème traité tôt, dans un environnement stable,
> coûte beaucoup moins cher qu'un bug silencieux découvert en recette — voire en production —
> après la bascule.

---

## 🗺️ Feuille de route du module

Les six sections suivantes vont du plus mécanique (l'hygiène du code) au plus stratégique
(le gel du périmètre) :

- **[5.1 — Nettoyer](01-nettoyer.md)** : supprimer le code mort, imposer `Option Explicit` partout,
  éliminer les `Variant` qui ne servent à rien. Le ménage de base.
- **[5.2 — Réduire les pièges en amont](02-reduire-pieges-amont.md)** ⭐ : typage explicite et
  `ByVal` explicite **côté VB6**, pour neutraliser à la source les changements silencieux les plus
  fréquents.
- **[5.3 — Découpler l'UI de la logique métier](03-decoupler-ui-metier.md)** : séparer le « quoi »
  (les règles) du « comment on l'affiche », afin de migrer par tranches et de tester la logique
  indépendamment des formulaires.
- **[5.4 — Documenter les comportements critiques](04-documenter-comportements.md)** : consigner
  noir sur blanc ce qui se brise facilement — arrondis, formats de dates, ordre d'enchaînement des
  événements.
- **[5.5 — Mettre en place un harnais de tests de référence](05-golden-master.md)** ⭐ : capturer le
  comportement de l'application VB6 (*golden master*) pour le comparer automatiquement après
  migration. Le filet de sécurité central, réutilisé au module 17 et détaillé en
  [Annexe F](../annexes/golden-master/README.md).
- **[5.6 — Geler le périmètre](06-geler-perimetre.md)** : arrêter les évolutions fonctionnelles
  pendant l'opération, pour ne pas viser une cible mouvante.

---

## ⚠️ Le piège de l'impatience

La tentation est forte de sauter cette partie et de « lancer l'assistant tout de suite ». C'est  
généralement un mauvais calcul. Migrer un code non nettoyé revient à migrer aussi ses `Variant`  
inutiles, son code mort et ses ambiguïtés : du travail en plus, et des pièges en plus. Sans oracle  
ni *golden master*, les régressions se découvrent trop tard, sans qu'on sache d'où elles viennent.  
Et sans gel du périmètre, on vise une cible mouvante, où chaque nouvelle fonctionnalité brouille la  
comparaison « avant / après ».

> 🧭 **Repère honnête** : préparer n'est pas *tout* perfectionner. L'objectif est de rendre la
> migration plus sûre, pas de réécrire idéalement un code qu'on va de toute façon transformer.
> Visez le bon niveau de préparation pour **votre** budget et **vos** risques — ni bâclé, ni
> interminable.

---

## Où ce module s'inscrit

- **Avant** : la [Partie 1](../01-cadrage-migration/README.md) vous a permis de comprendre et
  cadrer (pourquoi migrer, ce qui change, quelle stratégie). En particulier, l'inventaire et la
  cartographie des dépendances ([module 3](../03-evaluer-strategie/README.md)) vous disent *quoi*
  préparer.
- **Ici** : on prépare la matière première — le code VB6 source.
- **En parallèle** : le [module 6](../06-preparer-environnement/README.md) prépare l'autre côté du
  pont — l'environnement .NET (solution, ciblage 4.7.2, options de compilation, gestion de
  versions).
- **Ensuite** : la [Partie 3](../07-types-variables/README.md) attaque le cœur — la migration du
  langage lui-même.

---

> ✅ **À retenir** : on ne migre bien que ce qu'on a d'abord stabilisé, allégé, documenté et mis
> sous filet. Le temps investi dans ce module se rembourse intégralement lors de la validation
> (module 17), quand il faudra prouver que **rien n'a changé** dans le comportement.

---

**Module précédent** : [4. Outils de migration](../04-outils-migration/README.md)  
**Module suivant** : [6. Préparer l'environnement .NET](../06-preparer-environnement/README.md)  
**Section suivante** : [5.1 — Nettoyer le code VB6 →](01-nettoyer.md)

⏭️ [Nettoyer (code mort, `Option Explicit` partout, supprimer le `Variant` superflu)](/05-preparer-code-vb6/01-nettoyer.md)
