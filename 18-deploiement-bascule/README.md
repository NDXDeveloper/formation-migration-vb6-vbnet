🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18. Déploiement et bascule en production

**Faire passer l'application migrée en production sans mettre l'activité en péril : empaqueter, déployer progressivement, basculer prudemment — et toujours garder une voie de retour.**

> 🎯 **C'est ici qu'une migration réussit ou échoue *aux yeux de l'entreprise*.** Une migration
> techniquement parfaite qui rate sa bascule — basculement *big-bang* sans retour possible, production
> cassée, utilisateurs bloqués — est un **échec**. Ce module applique à l'étape finale la même
> philosophie **incrémentale et prudente** qui a guidé toute la formation : on **dé-risque** la mise en
> production, on **valide en conditions réelles**, et on conserve un **plan de retour arrière**.

---

## Où nous en sommes

Le module 17 était la **porte** : correction prouvée (golden master), code durci (`Option Strict`,
`Compatibility` retiré), pièges silencieux neutralisés, aucune fuite. L'application est **prête**. Ce
module répond à la question suivante : *comment la mettre entre les mains des utilisateurs, en  
production, sans casse ?*

> ⚠️ **Ce module n'est pas l'endroit où l'on corrige encore des bugs de migration** — c'était le rôle du
> module 17, et ses **critères de sortie** doivent être remplis **avant** d'ouvrir celui-ci. Ici, on
> déploie un produit *fini*.

---

## Le principe : dé-risquer la bascule

Quatre idées directrices encadrent tout le module :

1. **Jamais de basculement *big-bang* irréversible.** Comme la migration elle-même (module 3.4), le
   déploiement se fait **par étapes** : on limite le **rayon d'impact** de chaque pas.
2. **Toujours une voie de retour.** Un **plan de *rollback*** testé est **non négociable** : on ne
   bascule pas sans issue de secours (et ce plan couvre aussi les **données**, pas seulement
   l'exécutable).
3. **Valider en conditions réelles.** Le **pilote** et le ***parallel run*** prolongent en production la
   logique du *golden master* (module 17.3) : comparer le nouveau et l'ancien sur des **données et une
   charge réelles** avant de s'engager totalement.
4. **La bascule est un *processus*, pas un *événement*.** Il y a une période de **stabilisation**, un
   **suivi**, et un **transfert de compétences** (18.4) qui rend le résultat **durable**. On ne « bascule
   pas et on s'en va ».

---

## Quatre étapes, une voie de retour

```
  Préparer            Déployer progressivement        Basculer prudemment        Pérenniser
  ────────            ────────────────────────        ───────────────────        ──────────
  Packaging (18.1) ─► Cohabitation VB6/.NET (18.2) ─►  Pilote → parallel run ─►   Doc &
  installeur,         (déploiement par étapes) 🔗       → bascule (18.3)           transfert (18.4)
  prérequis 4.7.2                                            │
                                                             ▼
                                                  Plan de rollback ◄─ à chaque étape, une voie de retour
```

> 🔗 **La cohabitation (18.2) est le versant *déploiement* du travail d'interopérabilité du module 16.**
> Grâce au **CCW** (module 16.2) et à la stratégie de cohabitation COM (module 3.5), on peut **déployer
> des tranches migrées** qui coexistent en production avec le VB6 pas encore migré — sans attendre que
> 100 % du code soit converti pour livrer quoi que ce soit.

---

## Feuille de route du chapitre (18.1 → 18.4)

- **18.1 — Packaging .NET Framework (installeur, prérequis 4.7.2, redistribuables).** Produire
  l'artefact déployable : installeur, présence du **runtime 4.7.2** sur les postes cibles,
  redistribuables, et **enregistrement COM** des composants d'interop (`regasm`/`regsvr32`, **bitness**
  x86 et inscription **32 bits** — les sujets des modules 16.2 / 16.6 arrivent à échéance ici).
- **18.2 — Cohabitation VB6 / .NET pendant la transition (déploiement progressif). 🔗** Faire coexister
  l'ancien et le nouveau en production via COM, pour un déploiement **par étapes** plutôt qu'un saut
  unique — la concrétisation des modules 3.5 et 16.2.
- **18.3 — Stratégie de bascule : pilote, *parallel run*, plan de *rollback*.** Le cœur de la gestion du
  risque : déployer d'abord à un **groupe pilote**, faire tourner ancien et nouveau **en parallèle** pour
  comparer, et préparer un **retour arrière** testé en cas de problème.
- **18.4 — Documentation et transfert de compétences à l'équipe.** Le versant **humain** : l'équipe qui
  maintiendra l'application doit comprendre .NET (pas seulement VB6), les **décisions** de migration, les
  **frontières d'interop** et les pièges — pour que la valeur de la migration **perdure**.

---

## Prérequis et liens utiles

Ce module **met en œuvre** plusieurs fils tissés plus tôt :

- **Module 17 (tout entier)** — la **porte** : ses critères de sortie doivent être remplis avant de
  déployer.
- **Module 3.4** — *big-bang* vs incrémental : la stratégie de **déploiement** reflète celle de la
  **migration**.
- **Module 3.5** — la cohabitation COM : 18.2 en est la réalisation en production.
- **Module 3.6** — critères de réussite et indicateurs : ils définissent une bascule **réussie** et les
  **déclencheurs de *rollback***.
- **Modules 5.5 / 17.3** — le *golden master* : le ***parallel run*** en est l'extension **en
  production**.
- **Module 6.5** — Git, branche de migration, jalons : la gestion des **livraisons**.
- **Module 16** — l'interopérabilité (RCW/CCW, **enregistrement**, **bitness**) : ses implications de
  **packaging** et de cohabitation se concrétisent ici (18.1 / 18.2).

> 🧭 **Au-delà de ce module** : la partie 7 propose deux prolongements — la **migration assistée par IA**
> (module 19, transversal) et, surtout, le **second saut** optionnel de 4.7.2 vers .NET moderne (module
> 20), à n'envisager qu'une fois la présente mise en production **stabilisée** (deux sauts **séparés**).

---

## Quand la bascule est-elle « réussie » ?

Définissez des **critères de réussite** et des **déclencheurs de *rollback*** explicites **avant** de  
basculer (en cohérence avec le module 3.6) — par exemple :

- ✅ Le **pilote** s'est déroulé sans incident bloquant sur une durée représentative.
- ✅ Le ***parallel run*** ne montre **aucun écart** ancien/nouveau non justifié.
- ✅ Les **indicateurs de production** (erreurs, performance, ressources — cf. module 17.6) sont dans les
  seuils attendus pendant la **période de stabilisation**.
- ✅ Le **plan de *rollback*** a été **testé** (et reste actionnable).
- ✅ L'équipe dispose de la **documentation** et des **compétences** pour exploiter le nouveau système
  (18.4).

Tant que ces points ne sont pas réunis, on **ne bascule pas définitivement** — on reste en déploiement  
progressif, filet déployé.

---

## En résumé

- Le déploiement est l'étape où la migration **réussit ou échoue pour l'entreprise** : on la **dé-risque**.
- **Quatre idées directrices** : pas de *big-bang* irréversible, **toujours un *rollback*** (données
  comprises), **valider en réel** (pilote + *parallel run*, prolongement du *golden master*), et traiter
  la bascule comme un **processus** (stabilisation, suivi, transfert).
- **Quatre étapes** : **packaging** (18.1), **cohabitation** progressive VB6/.NET (18.2, le versant
  déploiement du module 16), **bascule prudente** (18.3, pilote → *parallel run* → bascule, avec
  *rollback*), **pérennisation** par la **documentation et le transfert** (18.4).
- Ce module **n'ouvre** qu'une fois les **critères de sortie du module 17** remplis, et il définit ses
  propres **critères de réussite** (module 3.6).
- **Au-delà** : la migration cœur (parties 1-6) est complète ; la partie 7 traite l'**IA** (19) et le
  **second saut** optionnel vers .NET moderne (20), **après** stabilisation.

> ➡️ **Commençons par le commencement concret : produire un artefact déployable — installeur, prérequis
> 4.7.2, redistribuables et enregistrement des composants d'interop.** (Module 18.1)

---

🏷️ **Indicateurs du module** : déploiement et bascule · 🔗 Interop (cohabitation, 18.2) · met en œuvre les modules 3.4/3.5/3.6, 6.5, 16 et 17
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026  
**Porte d'entrée** : critères de sortie du module 17 remplis · **Suite optionnelle** : second saut vers .NET moderne (module 20)

⏭️ [Packaging .NET Framework (installeur, prérequis 4.7.2, redistribuables)](/18-deploiement-bascule/01-packaging.md)
