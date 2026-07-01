🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 5.6 Geler le périmètre (arrêter les évolutions fonctionnelles pendant la migration)

> *Partie 2 — Préparer le terrain · Module 5 — Préparer le code VB6 avant la migration*

---

## Le problème que résout cette section

Migrer, c'est transformer une application **sans changer ce qu'elle fait**. C'est une **transformation à comportement préservé** : même entrées, mêmes sorties, simplement un autre langage et un autre runtime dessous.

Or une telle transformation suppose une chose évidente mais souvent négligée : **que le comportement à préserver reste stable pendant qu'on le transforme.** Si l'application continue d'évoluer fonctionnellement en même temps qu'on la migre, on poursuit une **cible mouvante** : on ne peut plus comparer « avant » et « après », parce que le « avant » lui-même n'arrête pas de bouger.

> ⚠️ **Migrer et faire évoluer en même temps, c'est changer deux variables à la fois.** Quand un écart apparaît, impossible de savoir s'il vient de la migration ou de la nouvelle fonctionnalité. On perd l'attribution de cause à effet — le fondement même du débogage.

**Geler le périmètre**, c'est décider explicitement de **suspendre les évolutions fonctionnelles** le temps de la migration, pour la ramener à ce qu'elle doit être : un changement d'**une seule** variable.

---

## Qu'est-ce que « geler le périmètre » ?

C'est un engagement, pris **avant** de commencer, à ne migrer que l'**existant tel quel** :

- on **arrête** d'ajouter des fonctionnalités ;
- on **arrête** de modifier les comportements métier ;
- on **reporte** les améliorations, refontes et corrections non bloquantes ;
- on **migre** le comportement actuel à l'identique, puis on rouvre les évolutions **après** validation.

C'est le pendant organisationnel du principe vu au § 5.5 : on **caractérise** le comportement existant (bugs compris) au lieu de chercher à l'améliorer en chemin. Le gel est ce qui rend ce principe **tenable dans la durée**.

---

## Pourquoi geler : trois raisons qui se renforcent

**1. Isoler une seule variable.** La migration change déjà énormément de choses simultanément : le langage, le runtime (VBVM → CLR), le modèle objet, le typage. C'est déjà beaucoup de surface au risque. Y superposer des changements fonctionnels multiplie les sources d'erreur possibles et rend chaque régression bien plus difficile à diagnostiquer.

**2. Garder une référence valide (le lien direct avec le § 5.5).** Le *golden master* repose sur une référence figée, capturée sur l'application VB6 intacte. **Si le comportement fonctionnel change pendant la migration, cette référence devient caduque** : on compare le code migré à un étalon qui ne décrit plus ce que l'application est censée faire. Sans gel, pas de golden master fiable ; sans golden master, pas de filet de sécurité.

**3. Pouvoir déboguer.** Quand le comportement de départ est constant, toute différence observée après migration **est** un effet de la migration. On sait où chercher. C'est cette certitude qui transforme la chasse aux régressions en travail méthodique plutôt qu'en enquête sans fin.

---

## Ce qui est gelé — et ce qui ne l'est pas

Un gel **total** est rarement réaliste en entreprise. La vraie compétence consiste à distinguer ce qui se reporte de ce qui ne peut pas attendre.

| Catégorie de changement | Statut pendant la migration |
|-------------------------|------------------------------|
| Nouvelles fonctionnalités | 🔒 **Gelé** |
| Améliorations « tant qu'on y est » | 🔒 **Gelé** |
| Refonte de l'UI / de l'ergonomie | 🔒 **Gelé** |
| Refactoring qui **change le comportement** | 🔒 **Gelé** |
| Correction de bugs historiques **non bloquants** | 🔒 **Gelé** *(reporté, voir backlog)* |
| Optimisations de performance non critiques | 🔒 **Gelé** |
| — | — |
| Correctif d'un bug **bloquant** en production (P1) | ⚠️ **Exception encadrée** |
| Patch de **sécurité** critique | ⚠️ **Exception encadrée** |
| Changement **légal / réglementaire** non négociable | ⚠️ **Exception encadrée** |

> 💡 **Le piège du « tant qu'on y est ».** C'est la façon la plus courante dont un gel s'effrite. « Pendant que je migre ce module, autant corriger ce petit truc / améliorer cet écran / nettoyer cette règle… » Chaque micro-amélioration est une occasion de nouveau bug **et** une divergence avec la référence. La règle est nette : **on migre, on ne redessine pas.** Les bonnes idées vont au backlog (voir plus bas), pas dans le commit de migration.

---

## Le coût caché : la double maintenance

Pourquoi le gel doit-il être **court** autant que strict ? À cause d'un coût qui grandit avec le temps : la **double maintenance**.

Dès qu'un changement *doit* être appliqué pendant la migration (un correctif légal, un bug P1), la question devient : **où** ? Et la réponse dépend de l'avancement :

- Si le module concerné **n'est pas encore migré**, on corrige le **VB6**.
- S'il **est déjà migré**, on corrige le **VB.NET**.
- Si l'on est en migration **incrémentale avec cohabitation** (VB6 et .NET tournent côte à côte via interop COM — voir **§ 3.5** et **§ 18.2**), la logique peut être **scindée entre les deux mondes**, et le correctif doit alors être porté **des deux côtés**. C'est coûteux et propice aux erreurs.

> ⚠️ **C'est « l'impôt des deux bases de code ».** Chaque jour passé avec un pied dans chaque monde augmente la probabilité qu'un changement urgent doive être fait deux fois — et qu'il soit fait correctement d'un côté, mal de l'autre. La parade n'est pas d'interdire les exceptions (impossible), mais de **raccourcir la fenêtre** pendant laquelle elles coûtent double : un gel court réduit mécaniquement la facture.

---

## Le gel dépend de la stratégie de migration

La **forme** du gel découle directement de la stratégie choisie au **module 3** (*big-bang* vs incrémentale) :

- **Big-bang** : un **gel global** unique, suivi d'une bascule. La fenêtre de gel est plus courte mais plus dense, et tout l'effort de migration tombe d'un coup.
- **Incrémentale** : un gel **par module** au moment où on le migre, pendant que la cohabitation COM laisse le reste de l'application continuer à fonctionner. La durée totale est plus longue, mais le risque est étalé et la double maintenance est circonscrite aux modules en cours.

> 💡 D'où l'importance de l'**estimation de l'effort** (§ 3.3) : **un gel sans date de fin sera violé.** On ne « met pas en pause les évolutions indéfiniment » ; on annonce une fenêtre **bornée et crédible**. C'est la seule sorte de gel qui tienne réellement.

---

## Gérer les exceptions : un processus, pas du cas par cas

Les exceptions existeront (l'État ne suspend pas un changement de taux légal parce que vous migrez). L'enjeu n'est pas de les éliminer, mais de les rendre **contrôlées et tracées** plutôt qu'improvisées. Un processus simple en cinq temps :

1. **Qualifier.** Le changement relève-t-il vraiment d'une catégorie autorisée (légal, sécurité, bug bloquant P1) ? Sinon → **backlog**, sans exception.
2. **Décider.** L'exception est validée par un **responsable identifié** (chef de projet, *product owner*), jamais par une décision individuelle au fil de l'eau.
3. **Appliquer à la source de vérité.** Selon la stratégie : sur le VB6 s'il reste maître, ou des **deux côtés** en cas de cohabitation.
4. **Mettre à jour le golden master.** Le comportement de référence a **légitimement** changé : on « bénit » le nouvel écart et on met à jour la référence en conscience (mécanique décrite au **§ 5.5**). Sans cette mise à jour, le filet déclencherait un faux positif sur un changement voulu.
5. **Documenter.** Quoi, pourquoi, où appliqué — idéalement dans le message de commit qui modifie la référence.

> ⚠️ La règle d'or du § 5.5 s'applique ici aussi : **on ne met jamais à jour la référence « pour faire passer le test ».** On la met à jour **après** avoir compris et validé l'écart. Une exception non documentée est une exception qui, dans six mois, ressemblera à une régression inexpliquée.

---

## Que faire des demandes qui arrivent pendant le gel ?

Geler les évolutions ne veut **pas** dire ignorer les besoins métier. On ne **stoppe** pas le recueil des demandes — on les **met de côté** dans un backlog dédié. Ce backlog a deux vertus :

- Il **rassure** les parties prenantes : leurs demandes ne sont pas perdues, elles sont **planifiées pour après**.
- Il **devient la feuille de route** post-migration : la première vague de travail une fois le gel levé.

> 💡 **Certaines évolutions seront même *meilleures* à faire après.** Bon nombre d'« améliorations » sont plus simples et plus propres à implémenter en **VB.NET idiomatique** (LINQ, génériques, espace `My`, *data binding* moderne — voir **§ 17.5**) que dans le VB6 d'origine. Pour celles-là, le gel n'est pas du temps perdu : c'est « la même fonctionnalité, mieux faite, un peu plus tard ». Cet argument aide beaucoup à faire accepter le gel.

---

## Faire accepter le gel : un enjeu de communication

Le gel est autant un **sujet de gouvernance** qu'un sujet technique. Un long gel des fonctionnalités est difficile à « vendre » : il faut un accord **explicite et préalable** du métier, pas un gel décrété en silence par l'équipe technique. Quelques principes :

- **Cadrer le « pourquoi ».** Le gel n'est pas un caprice d'ingénieurs : c'est ce qui rend la migration **sûre et rapide** et qui protège l'investissement déjà consenti.
- **Borner la durée.** S'appuyer sur l'estimation (§ 3.3) pour annoncer une fenêtre crédible, avec un terme visible.
- **Montrer la contrepartie.** Le backlog gelé + la promesse d'évolutions plus rapides ensuite donnent quelque chose à attendre.
- **Obtenir un sponsor.** Un gel tenu repose sur un décideur qui arbitre les exceptions et tient la ligne face à la pression.

Le gel agit alors comme un **moteur de vitesse** : parce que le métier attend, l'équipe est incitée à **finir la migration vite**, sans la sur-ingénierie ni le *gold plating* qui la feraient déraper. À l'inverse, une migration qui s'éternise rend le gel **intenable** : il finit par être violé, la référence s'effondre, et tout l'édifice de sécurité avec elle.

---

## Lever le gel

On rouvre les évolutions **une fois la migration validée**, pas avant :

- *golden master* au vert et tests de **non-régression** passés (**§ 17.3**) ;
- application **en production** ou pilote validé, plan de *rollback* en place (**§ 18.3**).

À cet instant, le backlog gelé **devient la feuille de route** — et, bonne nouvelle, beaucoup de ces évolutions se construisent désormais plus vite, sur une base VB.NET moderne et maintenable.

---

## Les limites, en toute lucidité

- Un gel **parfait** (zéro changement) est rarement atteignable dans une vraie organisation. L'objectif réaliste est de **minimiser et contrôler** le changement, pas de l'annuler.
- Plus la migration est **courte**, plus le gel est **facile** à tenir et moins la double maintenance coûte. Discipline de gel et **rapidité de migration se renforcent mutuellement** : c'est un cercle vertueux qu'il faut entretenir activement.
- Le gel ne supprime pas le besoin de juger : il fournit un **cadre** pour décider, vite et au bon niveau, ce qui se reporte et ce qui passe en exception.

---

## À retenir

> - **Geler le périmètre**, c'est suspendre les évolutions fonctionnelles pendant la migration pour ne changer qu'**une seule variable** à la fois.
> - C'est la condition de validité du **golden master** (§ 5.5) : sans comportement stable, pas de référence comparable.
> - On distingue ce qui se **reporte** (nouvelles fonctionnalités, améliorations, bugs non bloquants) de ce qui passe en **exception encadrée** (légal, sécurité, bugs P1).
> - Toute exception suit un **processus** : qualifier → décider → appliquer à la source de vérité → **mettre à jour la référence** → documenter.
> - Le gel doit être **borné, communiqué et sponsorisé** : un gel sans date de fin sera violé.
> - Le **coût de la double maintenance** croît avec la durée → un gel **court** est aussi important qu'un gel **strict**.
> - On lève le gel **après validation** ; le backlog gelé devient alors la feuille de route, souvent réalisable plus vite en VB.NET moderne.

---

## Liens utiles

- **§ 5.5** — Harnais de tests de référence (*golden master*) *(ce que le gel protège)*
- **§ 3.3** — Estimer l'effort *(pour borner la durée du gel)*
- **§ 3.4** — Stratégies : *big-bang* vs incrémentale *(détermine la forme du gel)*
- **§ 3.5** / **§ 18.2** — Cohabitation VB6 / .NET par interop COM *(d'où vient la double maintenance)*
- **§ 17.3** — Tests de non-régression *(le feu vert pour lever le gel)*
- **§ 17.5** — Refactoring idiomatique *(pourquoi certaines évolutions sont meilleures après)*
- **§ 18.3** — Stratégie de bascule : pilote, *parallel run*, *rollback* *(quand le gel se lève réellement)*

⏭️ [Préparer l'environnement .NET](/06-preparer-environnement/README.md)
