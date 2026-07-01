🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17. Valider, fiabiliser et refactoriser après migration ⭐

**Transformer un portage qui *compile* en une application dont on sait qu'elle *se comporte correctement* : prouver l'équivalence, serrer les boulons, traquer les pièges silencieux, puis moderniser — dans cet ordre.**

> ⚠️ **« Ça compile » n'est pas « c'est terminé ».** Rappel du fil rouge de toute la formation (modules
> 1.5 et 2.x) : *le vrai danger n'est pas le code qui refuse de compiler — c'est le code qui compile
> **mais ne se comporte plus pareil**.* Un `ByRef` devenu `ByVal`, un `Integer` redimensionné, un
> `Class_Terminate` qui ne se déclenche plus au bon moment : ces différences **silencieuses** ne se
> voient pas à la compilation. Ce module est précisément l'endroit où on les débusque — et où la
> migration passe du statut « ça marche, je crois » à « c'est correct, je le prouve ».

---

## Où nous en sommes

Le langage est migré (partie 3), l'interface aussi (partie 4), les données et l'interopérabilité
également (partie 5). Le code **compile et s'exécute**. C'est une étape majeure — mais ce n'est **pas**
la fin. Il reste à répondre à trois questions, qui donnent leur nom au module :

- **Valider** — l'application produit-elle exactement les mêmes résultats qu'avant ?
- **Fiabiliser** — peut-on durcir le code pour faire remonter les défauts latents et éliminer les
  béquilles de migration ?
- **Refactoriser** — peut-on, une fois la correction prouvée, passer d'une traduction mécanique à du
  VB.NET **idiomatique et moderne** ?

C'est aussi la **porte d'entrée de la mise en production** : le module 18 (déploiement et bascule) ne  
doit commencer qu'une fois ce module franchi.

---

## Pourquoi l'inspection visuelle ne suffit pas

Les changements que l'on cherche ici sont, par nature, **silencieux** : ils ne lèvent ni erreur de  
compilation, ni exception à l'exécution. Relire le code « à l'œil » ne les révèle pas de façon fiable —  
sinon ils ne seraient pas silencieux. La seule méthode robuste est la **comparaison automatisée** du  
comportement migré avec un **comportement de référence capturé avant la migration** : le *golden  
master* (préparé au module 5.5, exploité ici au 17.3). C'est ce qui justifie la place centrale qu'il  
occupe dans ce module.

---

## Trois verbes, un ordre

L'ordre compte. On ne se jette pas sur le refactoring : on **sécurise** d'abord, on **durcit** ensuite,  
on **traque** les pièges, et l'on **modernise** en dernier — chaque étape étant revalidée contre le  
filet de sécurité.

```
  Filet de sécurité          Durcissement               Traque ciblée          Modernisation
  ─────────────────          ────────────               ─────────────          ─────────────
  Golden master (17.3)  ──►  Option Strict (17.1)  ──►   Pièges          ──►    Refactoring (17.5)
  « où en suis-je ? »        Compatibility (17.2)        silencieux (17.4)      Perf / mémoire (17.6)
                             « on serre les boulons »    « annexe B »           « code idiomatique »
        ▲                                                                              │
        └────────────── à chaque étape, on revalide contre le golden master ───────────┘
```

> ℹ️ **L'ordre des fichiers n'est pas forcément l'ordre d'exécution.** En pratique, on établit
> **d'abord** le filet (17.3) avant de serrer (17.1 / 17.2) et de traquer (17.4) ; la modernisation
> (17.5 / 17.6) vient en **dernier**, une fois la correction prouvée. La numérotation suit l'ordre de
> lecture ; ce schéma donne l'ordre des opérations.

---

## Le principe cardinal : le filet de sécurité **d'abord**

Trois règles non négociables encadrent tout ce module :

1. **Le *golden master* est ce qui rend tout le reste sûr.** Durcir `Option Strict`, retirer les
   dépendances de compatibilité, et surtout refactoriser **modifient le code** — et toute modification
   peut introduire une régression. La comparaison de non-régression est le filet qui les **rattrape**.
   Sans lui, refactoriser, c'est parier.
2. **On ne migre et on ne refactorise jamais en même temps.** Le périmètre fonctionnel a été **gelé**
   au module 5.6 pour cette raison : on ne change qu'**une** chose à la fois (d'abord rendre le
   comportement identique, ensuite seulement l'améliorer), sinon on ne sait plus quelle modification a
   causé quel écart.
3. **Tout écart par rapport au *golden master* doit être compris et justifié** — jamais accepté « parce
   que c'est sûrement bénin ». Un écart est soit un **bug de migration** à corriger, soit une
   **différence attendue et documentée**. Pas de troisième catégorie.

---

## Feuille de route du chapitre (17.1 → 17.6)

- **17.1 — Activer `Option Strict` progressivement (éliminer le *late binding* résiduel). ⭐** Le
  paiement de la promesse des modules 6.4 et 16.3 : passer en mode strict **fichier par fichier**, ce
  qui interdit le liage tardif *et* les conversions implicites, et fait remonter quantité de défauts
  latents.
- **17.2 — Supprimer les dépendances à `Microsoft.VisualBasic.Compatibility`. ⚠️** Retirer la
  « béquille » signalée au module 4.2 : remplacer les appels de l'espace de compatibilité par du code
  .NET idiomatique, pour ne pas figer la migration sur un échafaudage temporaire.
- **17.3 — Tests de non-régression : comparer avec le *golden master*. ⭐** Le **filet de sécurité** :
  rejouer le comportement de référence capturé au module 5.5 et comparer, automatiquement, l'avant et
  l'après. La pièce maîtresse du module.
- **17.4 — Traquer les pièges silencieux (ByRef, arrondis, dates, bornes de tableaux, finalisation). ⭐ ⚠️**
  L'audit systématique du **catalogue de l'annexe B**, appliqué à votre code : la liste de contrôle des
  changements qui compilent mais cassent le comportement.
- **17.5 — Refactoring idiomatique : LINQ, génériques, espace `My`, code VB.NET moderne.** La
  **récompense** : une fois la correction prouvée, passer d'une traduction littérale à du VB.NET
  expressif et maintenable.
- **17.6 — Revue de performance et de mémoire (`IDisposable`, `Using`, fuites de *handles*).** Vérifier
  que la finalisation non déterministe (modules 2.2 / 12.3) ne laisse pas fuir fichiers, connexions ou
  *handles* — un point particulièrement sensible après de l'interop (modules 16.1 / 16.6).

---

## Prérequis et liens utiles

Ce module **récolte** ce qui a été semé plus tôt. Gardez à portée :

- **Module 5.5** — le harnais de tests de référence (*golden master*) : il devait être **capturé avant**
  la migration ; on le **compare** ici.
- **Module 5.6** — le gel du périmètre : la discipline qui rend la validation interprétable.
- **Module 6.4** — les options `Option Strict` / `Explicit` / `Infer` / `Compare` : 17.1 en est
  l'aboutissement.
- **Module 4.2** — l'espace `Microsoft.VisualBasic.Compatibility`, la béquille à retirer (17.2).
- **Module 16.3** — l'élimination du *late binding* résiduel, pilotée globalement par 17.1.
- **Modules 2.2 / 12.3** — la finalisation déterministe perdue et le patron `IDisposable`/`Using` :
  centraux pour 17.4 et 17.6.
- **Module 3.6** — les critères de réussite et indicateurs : ils définissent quand ce module est
  « terminé ».

Et côté annexes : **B** (catalogue des pièges silencieux — la référence de 17.4), **F** (modèles de
*golden master* — l'outillage de 17.3), et **E** (checklist de migration).

---

## Quand peut-on dire que c'est « terminé » ?

Définissez des **critères de sortie** explicites (en cohérence avec le module 3.6) avant d'ouvrir le  
module 18 :

- ✅ Le *golden master* **passe** — ou chaque écart restant est **compris et justifié**.
- ✅ **`Option Strict On`** sur l'ensemble du code — ou le liage tardif résiduel est **isolé et
  documenté** (patron d'isolation du module 16.3).
- ✅ **Aucune** dépendance à `Microsoft.VisualBasic.Compatibility` — ou les rares restantes sont
  **tracées** et planifiées.
- ✅ L'**audit des pièges silencieux** (annexe B) est mené à son terme.
- ✅ Aucune **fuite** de mémoire, de *handles* ou de ressources non managées.

Tant que ces points ne sont pas cochés, l'application **compile et tourne**, mais n'est pas **prête à  
basculer**.

---

## En résumé

- **« Ça compile » n'est pas « c'est terminé ».** Ce module débusque les différences de **comportement**
  silencieuses — le cœur du risque de migration.
- Les changements visés étant **silencieux**, seule la **comparaison automatisée** contre un
  comportement de référence (le *golden master*) les révèle de façon fiable.
- **Trois verbes, un ordre** : **valider** (filet de sécurité, 17.3), **fiabiliser** (durcir avec
  `Option Strict` 17.1 et retirer `Compatibility` 17.2, puis traquer les pièges 17.4), **refactoriser**
  (moderniser 17.5 et réviser perf/mémoire 17.6) — en revalidant à chaque étape.
- **Principe cardinal** : le filet **d'abord** ; on ne migre et on ne refactorise **jamais** en même
  temps ; **tout écart** est compris et justifié.
- Le module a des **critères de sortie** précis : il est la **porte** à franchir avant le déploiement.

> ➡️ **Une fois cette porte franchie, cap sur la mise en production :** packaging, cohabitation
> VB6/.NET pendant la transition, stratégie de bascule et plan de *rollback* (module 18).

---

🏷️ **Indicateurs du module** : ⭐ Point clé (validation) · ⚠️ Pièges (17.2, 17.4) · récolte les modules 4.2, 5.5/5.6, 6.4, 16.3 et l'annexe B
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026  
**Porte de sortie** : critères de réussite remplis (module 3.6) avant le déploiement (module 18)

⏭️ [Activer `Option Strict` **progressivement** (éliminer le *late binding* résiduel)](/17-valider-refactoriser/01-option-strict.md)
