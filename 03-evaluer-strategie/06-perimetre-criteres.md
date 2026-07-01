🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.6 Définir le périmètre, les critères de réussite et les indicateurs

> **Module 3 — Évaluer l'existant et choisir une stratégie**
> L'inventaire (3.1) a recensé, la cartographie (3.2) a décidé du sort de chaque dépendance,
> l'estimation (3.3) a chiffré, et le choix de stratégie (3.4-3.5) a fixé la trajectoire. Il
> reste à **cadrer** : que met-on dans la migration, à quoi reconnaît-on qu'elle a réussi, et
> comment sait-on, à tout moment, où l'on en est ? Cette section **clôt** le travail
> d'évaluation et le transforme en un projet pilotable.

---

## 🧭 Pourquoi cadrer avant de migrer

Tout le travail des sections précédentes converge ici. Sans cadrage explicite, une migration n'a  
ni **limites** (on ne sait pas où elle s'arrête), ni **but mesurable** (on ne sait pas quand elle  
est réussie), ni **tableau de bord** (on ne sait pas où l'on en est). Trois manques qui ouvrent la  
porte aux trois dérives les plus classiques d'un projet de migration :

- **le périmètre qui enfle** — on en profite pour « améliorer au passage », et la migration ne
  finit jamais ;
- **la réussite indéfinissable** — « c'est migré » devient une affaire d'opinion, jamais de
  preuve ;
- **l'avancement opaque** — on ne sait pas si l'on est à 30 % ou à 70 %, et l'effet tunnel
  (surtout en big-bang, 3.4) s'installe.

Cadrer, c'est se prémunir contre ces trois dérives **avant** d'écrire la première ligne de code  
migré.

> ⚠️ **Le piège fondateur : confondre migration et amélioration.** La tentation est immense — tant
> qu'on est dans le code, autant corriger ce bug, ajouter cette fonctionnalité, refactoriser cette
> horreur. C'est presque toujours une erreur. Migrer **et** améliorer en même temps, c'est mélanger
> deux sources de changement, et donc rendre impossible de distinguer une régression de migration
> d'un effet de la nouvelle fonctionnalité. Le périmètre clair est ce qui maintient les deux
> séparés.

---

## 🎯 Définir le périmètre

Le périmètre, c'est la **frontière** de la migration : ce qui en fait partie, et — tout aussi  
important — ce qui n'en fait pas partie. Le définir, c'est répondre à plusieurs questions.

### Que migre-t-on, et que ne migre-t-on pas ?

L'inventaire (3.1) et la cartographie (3.2) ont fourni la liste complète des constituants. Le  
cadrage tranche, pour chacun, son inclusion :

- **Le code mort** repéré au balayage (3.1) entre-t-il dans le périmètre ? En principe **non** : on
  ne migre pas ce qui ne sert plus. Il relève du **nettoyage en amont** (module 5.1), pas de la
  migration.
- **Les composants à éliminer** (décision de la cartographie 3.2 : `FileSystemObject` → `System.IO`,
  etc.) sont **dans** le périmètre, mais comme remplacement par une API native.
- **Les composants à réutiliser durablement en COM** (décision « réutiliser » de 3.2) sont, par
  définition, **hors** du périmètre de migration : ils restent tels quels derrière la frontière
  d'interop (3.5).
- **Les fonctionnalités obsolètes** que personne n'utilise plus : faut-il les migrer ou les
  **abandonner** ? Le cadrage est l'occasion de cette décision.

### Le principe directeur : iso-fonctionnel

Le cadrage le plus sûr pose un principe simple : la migration est **iso-fonctionnelle**. L'objectif  
n'est pas une application *meilleure*, mais une application qui fait **exactement la même chose**,  
sur une nouvelle plateforme. Le critère de réussite découle directement de ce principe (voir plus  
bas) : *se comporter comme avant*.

> 💡 **« Iso-fonctionnel » n'interdit pas tout changement — il les sépare dans le temps.** Les
> améliorations légitimes (corriger un vrai bug, moderniser une UI vieillissante, exploiter LINQ ou
> les génériques) ne sont pas proscrites : elles sont **reportées après** la migration, une fois le
> nouveau socle validé. C'est exactement l'esprit du **refactoring idiomatique** post-migration
> (module 17.5) et de la modernisation traitée au module 20. D'abord migrer à l'identique et
> prouver l'équivalence ; **ensuite** améliorer.

### Geler le périmètre

Définir le périmètre ne suffit pas s'il **bouge** pendant la migration. D'où la décision, étroitement  
liée, de **geler les évolutions fonctionnelles** le temps de la migration (module 5.6) : tant qu'on  
migre, on n'ajoute pas de fonctionnalités à la version VB6. Viser une cible mobile est l'un des  
moyens les plus sûrs de ne jamais l'atteindre.

> 💡 Le gel est plus facile à tenir en **big-bang** (campagne courte) qu'en **incrémental**
> (calendrier long), où l'activité doit souvent continuer à évoluer (3.4). Quand un gel total est
> impossible, la stratégie incrémentale et sa cohabitation (3.5) offrent une alternative : on
> **fige par morceau**, le temps de migrer chaque morceau, plutôt que l'application entière d'un
> bloc.

---

## ✅ Définir les critères de réussite

Un critère de réussite répond à une question redoutablement simple : **à quoi reconnaîtra-t-on que  
la migration est terminée et réussie ?** Tant que cette réponse n'est pas écrite **avant** de  
commencer, « c'est fini » restera une affaire de ressenti.

### Le critère central : la non-régression comportementale

Pour une migration iso-fonctionnelle, le critère de réussite **n°1** est que l'application migrée
**se comporte comme l'originale**. Ce n'est pas un vœu : c'est une propriété **vérifiable**, à
condition de s'en être donné les moyens **en amont**.

C'est tout l'enjeu du **harnais de tests de référence** — le *golden master* (module 5.5) : on
**capture** le comportement de l'application VB6 *avant* de migrer (entrées → sorties attendues), et
on **rejoue** ces mêmes cas sur l'application migrée pour vérifier qu'elle produit **les mêmes  
résultats** (tests de non-régression, module 17.3). La réussite n'est alors plus une opinion mais  
une **mesure** : tant de cas rejoués, tant de cas conformes.

> ⚠️ **Sans golden master, le critère de réussite est invérifiable.** C'est précisément contre les
> **pièges silencieux** (le `ByRef` devenu `ByVal`, l'arrondi qui change, la date mal convertie —
> Annexe B) que ce critère protège : ces bugs ne se voient pas à la lecture du code et ne font pas
> planter la compilation. Seule la comparaison méthodique du comportement *avant/après* les
> débusque. Définir « réussir = se comporter comme avant » **impose** de construire ce harnais —
> c'est l'une des raisons pour lesquelles il se met en place **dès la préparation** (module 5).

### Les autres dimensions de la réussite

La non-régression comportementale est le critère central, mais pas le seul. Selon le contexte, la  
réussite peut intégrer d'autres dimensions, à **expliciter dès le cadrage** :

| Dimension | Question | Lien formation |
|-----------|----------|----------------|
| **Comportement** (le cœur) | Produit-elle les mêmes résultats que l'originale ? | Golden master 5.5, 17.3 |
| **Périmètre couvert** | Tout ce qui devait être migré l'a-t-il été ? | Inventaire 3.1, checklist (Annexe E) |
| **Qualité du code** | `Option Strict` activé ? *late binding* résiduel éliminé ? | Module 17.1 |
| **Dépendances de compatibilité** | `Microsoft.VisualBasic.Compatibility` supprimé ? | Module 17.2 |
| **Performance / mémoire** | Pas de fuite de *handles* ? `IDisposable`/`Using` en place ? | Module 17.6 |
| **Déployabilité** | S'installe et tourne dans l'environnement cible ? | Module 18.1 |

> 💡 **Réussite « technique » vs réussite « propre ».** Une application peut « marcher » tout en
> traînant du *late binding*, des béquilles de compatibilité et une finalisation hasardeuse. Le
> cadrage doit décider du **niveau d'exigence** : se contente-t-on que « ça fonctionne », ou vise-t-on
> un code idiomatique et fiabilisé (module 17) ? Les deux sont des choix légitimes — mais ils
> doivent être **explicites**, car ils changent radicalement l'effort (3.3) et le calendrier.

---

## 📊 Définir les indicateurs

Le périmètre dit *où l'on va* ; les critères de réussite disent *à quoi ressemble l'arrivée* ; les
**indicateurs** disent *où l'on en est en chemin*. Ils transforment la migration en un projet  
**piloté** plutôt que subi — particulièrement crucial pour contrer l'effet tunnel du big-bang
(3.4).

### Des indicateurs d'avancement

Ils mesurent la **progression** vers la cible. Quelques exemples, à rapporter aux mesures de  
l'inventaire (3.1) et de l'estimation (3.3) :

- **Fichiers / modules migrés** sur le total inventorié (avec la nuance de 3.3 : un fichier migré
  n'égale pas un autre — pondérer par la complexité donne une image plus juste qu'un simple
  décompte) ;
- **Formulaires migrés** sur le total (la charge d'UI étant souvent la plus lourde) ;
- **Points durs traités** sur les points durs identifiés (3.3) — un indicateur particulièrement
  parlant, puisque c'est là que se concentre le risque ;
- **Dépendances résolues** sur les dépendances cartographiées (3.2), par décision (réutilisées /
  remplacées / éliminées).

> ⚠️ **L'indicateur trompeur : le pourcentage de code converti.** « 80 % du code migré » ne signifie
> **pas** « 80 % du travail fait » — c'est l'illusion dénoncée au module 1.5 et chiffrée en 3.3 : le
> volume restant (la strate rouge) peut concentrer la majorité de l'effort. Un bon indicateur
> d'avancement **pondère** par la difficulté, ou suit séparément la progression sur les **points
> durs**, plutôt que de se fier au volume brut converti.

### Des indicateurs de qualité

Ils mesurent non plus *combien* est migré, mais *dans quel état*. Ils relient le cadrage aux  
critères de réussite « propres » :

- **Taux de non-régression** : proportion des cas du *golden master* (5.5) qui passent sur la
  version migrée — **l'indicateur reine** de la réussite comportementale ;
- **Progression de `Option Strict`** : part du code basculé en typage strict (module 17.1), témoin
  de l'élimination du *late binding* résiduel ;
- **Dépendances de compatibilité restantes** : nombre de références à
  `Microsoft.VisualBasic.Compatibility` encore présentes (module 17.2), qui doit tendre vers zéro.

### Des indicateurs de risque résiduel

Plus subtils, ils mesurent **ce qui reste incertain** — directement issus des zones grises de la  
cartographie (3.2) et de la borne haute de l'estimation (3.3) :

- **zones grises non éclaircies** (DLL sans documentation, composant non identifié) encore ouvertes ;
- **points durs non encore abordés**, donc dont le coût réel reste à confirmer.

> 💡 **Un bon tableau de bord croise les trois familles.** L'avancement seul rassure à tort (« on
> avance ! ») s'il ignore la qualité (« mais ça se comporte-t-il comme avant ? ») et le risque
> résiduel (« et que reste-t-il d'incertain ? »). Les trois ensemble donnent une image **honnête**
> de l'état du projet.

---

## 🔗 Le cadrage, point de jonction du module et de la suite

Cette section **referme** le module 3 — l'évaluation de l'existant et le choix de stratégie sont  
faits — et **ouvre** sur la suite de la formation. Le cadrage produit ici alimente directement :

- la **préparation du code VB6** (module 5) : le périmètre désigne le code mort à nettoyer (5.1) et
  les comportements critiques à documenter (5.4) ; le critère de réussite **impose** le *golden
  master* (5.5) ; le gel du périmètre (5.6) en découle ;
- la **préparation de l'environnement** (module 6) : la stratégie (3.4) et le niveau d'exigence
  qualité (17.1) orientent la configuration de la solution ;
- la **validation finale** (module 17) : les critères de réussite définis ici sont **exactement**
  ce que la non-régression (17.3) et la traque des pièges (17.4) viendront vérifier ;
- la **bascule** (module 18) : la stratégie conditionne le plan de déploiement et de *rollback*
  (18.3) ;
- la **checklist de migration** (Annexe E) : l'opérationnalisation point par point du périmètre et
  des critères posés ici.

---

## ⚠️ Les pièges du cadrage

- **Ne pas cadrer du tout.** Le piège majeur : se lancer sans périmètre ni critères, en pensant que
  « ça se précisera en route ». Ça ne se précise jamais — ça enfle.
- **Mélanger migration et amélioration.** La tentation d'« améliorer au passage » brouille la
  frontière entre régression et nouveauté, et rend la non-régression invérifiable. D'abord migrer à
  l'identique, **ensuite** améliorer (module 17.5, 20).
- **Définir la réussite sans moyen de la vérifier.** « Se comporter comme avant » est un critère
  creux sans *golden master* (5.5) pour le mesurer. Le critère **impose** son outil.
- **Piloter au seul volume de code.** « X % converti » est l'indicateur le plus trompeur (1.5,
  3.3) : il ignore la strate rouge où se concentre l'effort.
- **Oublier les indicateurs de qualité et de risque.** Suivre l'avancement sans la qualité ni le
  risque résiduel donne une fausse assurance — on « avance » vers un résultat dont on ne mesure ni
  l'état ni les incertitudes.
- **Ne pas geler le périmètre.** Une cible qui bouge pendant qu'on la migre (module 5.6) est une
  cible qu'on n'atteint pas.

---

## ✅ En résumé

- Le cadrage **clôt** l'évaluation : il transforme l'inventaire (3.1), la cartographie (3.2),
  l'estimation (3.3) et la stratégie (3.4-3.5) en un projet **pilotable**.
- **Le périmètre** fixe ce qui entre dans la migration et ce qui en sort (code mort, composants
  réutilisés en COM, fonctionnalités obsolètes). Le principe directeur le plus sûr est
  **iso-fonctionnel** : migrer à l'identique, et **reporter après** les améliorations (module
  17.5, 20). Le périmètre doit en outre être **gelé** (module 5.6).
- **Les critères de réussite** répondent à « comment sait-on que c'est réussi ? ». Le critère
  central est la **non-régression comportementale** — *se comporter comme avant* —, vérifiable
  grâce au **golden master** (module 5.5, 17.3). D'autres dimensions (qualité du code, suppression
  des béquilles de compatibilité, performance, déployabilité) sont à expliciter selon le niveau
  d'exigence visé.
- **Les indicateurs** disent où l'on en est en chemin. Trois familles à croiser : **avancement**
  (en se méfiant du « % de code converti », trompeur — 1.5, 3.3), **qualité** (taux de
  non-régression en tête), et **risque résiduel** (zones grises et points durs non traités).
- Le piège fondateur est de **confondre migration et amélioration** ; le piège de pilotage est de
  se fier au **volume brut converti**.
- Ce cadrage alimente directement la **préparation** (modules 5-6), la **validation** (module 17),
  la **bascule** (module 18) et la **checklist** (Annexe E).

---

⬅️ Section 3.5 🔗 — [L'approche par interopérabilité COM](05-cohabitation-com.md)  
➡️ Module 4 🛠️ — [Outils de migration](../04-outils-migration/README.md)

---

**Juin 2026** · .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Outils de migration](/04-outils-migration/README.md)
