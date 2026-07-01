🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19. Migrer avec l'aide de l'IA 🤖 ⭐

**L'intelligence artificielle comme accélérateur de migration : puissante sur le travail mécanique, faillible sur les pièges silencieux, jamais une autorité finale.**

> 🧭 *Partie 7 — IA, suite de parcours et ressources*
> Module à aborder une fois les Parties 1 et 2 assimilées.

---

## De quoi parle ce module

En 2026, les assistants de code fondés sur de grands modèles de langage (LLM) sont devenus une aide crédible pour moderniser du *legacy*. Pour une migration VB6 → VB.NET, ils tombent à pic : une grande partie du travail est une **traduction mécanique** — convertir des mots-clés, réécrire des déclarations, transposer des structures — exactement ce que l'IA fait vite et plutôt bien. Mais c'est aussi le terrain le plus **traître** : les pires bugs de cette migration ne font pas planter le compilateur, ils changent **silencieusement** le comportement. Un `ByRef` devenu `ByVal`, un `Integer` redimensionné, un `Class_Terminate` qui ne se déclenche plus au bon moment : autant de régressions que l'IA produit avec un aplomb déconcertant, sous la forme d'un code qui **compile, mais ne se comporte plus pareil**.

Ce module vous apprend à faire de l'IA un **multiplicateur de productivité** sans jamais lui abandonner le volant. La règle de la formation ne bouge pas d'un iota : on **migre un existant qui doit continuer à tourner**. Tout ce que l'IA propose passe donc au crible du *golden master* (module 5, Annexe F) et du **catalogue des pièges silencieux** (Annexe B).

> ⭐ **Le principe directeur de ce module : l'IA propose, vous validez. Toujours.**
> Un assistant n'est ni un oracle ni un collègue senior : c'est un générateur de propositions plausibles. Aucune sortie d'IA n'entre dans la base de code sans avoir été **relue**, **comprise** et **confrontée aux tests de référence**. Vous restez le seul responsable du résultat.

---

## Pourquoi un module dédié

Les outils de migration « classiques » ont déjà été traités au [module 4](../04-outils-migration/README.md) : l'*Upgrade Wizard*, VBUC, VB Migration Partner. L'IA est un outil d'une **autre nature** — non pas un convertisseur déterministe que l'on configure une fois, mais un assistant conversationnel dont la qualité dépend **entièrement de la façon dont on le pilote**.

C'est précisément pour cela que la section 4.6 vous renvoie ici : l'IA mérite son propre mode d'emploi. Il faut savoir **quand** l'utiliser, **comment** la cadrer, et surtout **comment se protéger** de ses erreurs. Mal employée, elle accélère la production de régressions. Bien encadrée, elle absorbe la corvée mécanique et — atout souvent sous-estimé — elle **explique** le code VB6 obscur que plus personne dans l'équipe ne comprend.

---

## Ce que vous allez apprendre

À l'issue de ce module, vous saurez :

- situer **précisément** où l'IA apporte une vraie valeur et où elle reste dangereuse pour cette migration (**19.1**) ;
- rédiger des *prompts* efficaces pour convertir un module ou une classe, et pour **faire expliciter** un comportement VB6 obscur — en levant systématiquement l'ambiguïté **VB6 vs VB.NET** et en rappelant la cible **.NET Framework 4.7.2** (**19.2**) ;
- mettre l'IA au service de la **détection** des pièges (ByRef, `Variant`, finalisation, entiers redimensionnés) plutôt que de leur création (**19.3**) ;
- générer des **tests de non-régression** et de la **documentation XML** pour fiabiliser et documenter le code migré (**19.4**) ;
- reconnaître et neutraliser les **modes de défaillance propres à l'IA** — hallucinations, **faux équivalents** VB6/VB.NET — par une validation systématique (**19.5**).

---

## La place de l'IA dans la méthode de cette formation

L'IA ne remplace pas la démarche : elle s'y **branche**. Tout l'édifice tient toujours — préparation du code (module 5), *golden master*, activation **progressive** d'`Option Strict` ([module 17](../17-valider-refactoriser/README.md)), tests de non-régression — et l'IA vient simplement accélérer certaines étapes **à l'intérieur** de ce cadre.

Là où elle est la plus rentable :

- le **travail mécanique** (les fameux 60-90 % de l'effort de conversion) ;
- l'**explication** d'un fragment VB6 cryptique (effets de bord, propriétés par défaut, ordre d'événements) ;
- la **génération de tests** et de **documentation** à partir du code existant.

Là où le **jugement humain reste indispensable** :

- les **pièges silencieux** — le cœur même de cette formation ;
- l'**UI** et les *control arrays*, que les outils (IA comprise) traitent mal ;
- toute décision d'**architecture** ou de **périmètre**.

> 💡 L'IA est **un outil parmi d'autres**, à côté de l'*Upgrade Wizard* et des outils commerciaux vus au [module 4](../04-outils-migration/README.md). Le bon réflexe n'est pas « IA *ou* outil dédié » mais « le bon instrument pour chaque tâche » — souvent une **combinaison** des deux.

---

## Plan du module

| § | Sujet | Repère |
|---|-------|--------|
| [19.1](01-pourquoi-ia.md) | **Pourquoi l'IA aide pour VB6 → VB.NET (et ses limites)** — le panorama honnête des forces et des angles morts de l'IA sur cette migration précise. | ⭐ |
| [19.2](02-prompting.md) | **Prompting** — convertir un module/une classe, faire **expliquer** un comportement VB6 obscur. *Toujours préciser « VB6 » vs « VB.NET » et la cible .NET Framework 4.7.2.* | ⚠️ |
| [19.3](03-detecter-pieges.md) | **Détecter les pièges avec l'IA** — se servir de l'assistant comme d'un détecteur (ByRef, `Variant`, finalisation, entiers redimensionnés). | |
| [19.4](04-tests-documentation.md) | **Générer tests de non-régression et documentation XML** — transformer le code migré en filet de sécurité testé et documenté. | |
| [19.5](05-limites-pieges.md) | **Pièges de l'IA** — hallucinations, **faux équivalents** VB6/VB.NET, et la discipline de **validation systématique**. | ⚠️ |

---

## Prérequis et liens utiles

Avant d'attaquer ce module, il est recommandé d'avoir parcouru :

- les modules [1](../01-cadrage-migration/README.md) et [2](../02-differences-fondamentales/README.md) (cadrage et différences fondamentales VB6 ↔ VB.NET) — sans cette base, vous ne pourrez pas **juger** les propositions de l'IA ;
- le [module 4](../04-outils-migration/README.md) (outils de migration), dont la section 4.6 renvoie ici ;
- en lecture d'accompagnement permanente : l'[Annexe A](../annexes/correspondance-vb6-vbnet/README.md) (tableau de correspondance) et surtout l'[Annexe B](../annexes/pieges-silencieux/README.md) (catalogue des pièges silencieux), votre grille de contrôle des sorties d'IA.

Pour la suite logique, le [module 17](../17-valider-refactoriser/README.md) (valider, fiabiliser, refactoriser) prolonge directement le travail de validation amorcé ici.

> 📌 Rappel du parcours **« Migration assistée par IA »** 🤖 : modules **1-2, 4, 19** (≈ 3-4 jours).

---

## Les règles d'or (à garder sous les yeux)

> ⚠️ **À ne jamais transgresser**
>
> 1. **Désambiguïser systématiquement.** Indiquez « VB6 » (code source) vs « VB.NET » (cible) et précisez **.NET Framework 4.7.2** dans chaque *prompt*. C'est la première cause de réponses hors-sujet.
> 2. **Rien ne passe sans *golden master*.** Aucune conversion proposée par l'IA n'est validée tant qu'elle n'a pas été **comparée au comportement de référence**.
> 3. **Chaque « équivalent » est une hypothèse.** Traitez toute correspondance suggérée comme une piste à **vérifier**, en priorité sur les pièges connus de l'Annexe B.
> 4. **L'IA explique, vous décidez.** Excellente pour **comprendre** du VB6 obscur ; jamais habilitée à **tamponner** une conversion à votre place.
> 5. **Laisse courte.** Un module ou une classe à la fois — relu, testé, *commité* — plutôt qu'une conversion massive impossible à auditer.

---

## À retenir avant de commencer

Bien pilotée, l'IA rend cette migration **plus rapide** et **moins ingrate** : elle débroussaille le mécanique et éclaire l'obscur. Mal pilotée, elle devient une **machine à régressions silencieuses**, d'autant plus dangereuse qu'elle est **convaincante**. Toute la valeur de ce module tient dans la frontière entre les deux : **accélérer sans jamais déléguer la responsabilité**.

---

> 🧭 **Navigation**  
> ⬅️ Précédent : [18. Déploiement et bascule en production](../18-deploiement-bascule/README.md)  
> ⬆️ [Sommaire de la formation](../SOMMAIRE.md)  
> ➡️ Suivant : [19.1 — Pourquoi l'IA aide pour VB6 → VB.NET (et ses limites)](01-pourquoi-ia.md)

---

**Juin 2026** · Cible : .NET Framework 4.7.2 (le « pont » depuis VB6) · Licence : Creative Commons BY-NC-SA 4.0

⏭️ [Pourquoi l'IA aide pour VB6 → VB.NET (et ses **limites**)](/19-migration-ia/01-pourquoi-ia.md)
