🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.4 — Documentation et transfert de compétences à l'équipe

**Le versant humain de la migration, et celui qui la rend *durable* : transmettre à l'équipe qui maintiendra l'application la compétence .NET, les décisions prises, la carte de l'interop et les pièges — pour ne pas avoir troqué un *legacy* VB6 contre un *legacy* .NET mal maîtrisé.**

> 🤝 La dernière section du module — et du **parcours cœur** (parties 1-6). Toutes les autres sections
> étaient techniques ; celle-ci est **organisationnelle**. Le module 18 a réussi la mise en production ;
> cette section s'assure que sa **valeur perdure** au-delà de la bascule.

---

## Pourquoi : la valeur d'une migration se réalise **après** la bascule

La valeur d'une migration ne se réalise pas le jour de la bascule — elle se réalise sur les **années** où  
l'application migrée est **maintenue et fait évoluer**. Si l'équipe ne sait pas comprendre, corriger et
étendre le nouveau code .NET, cette valeur **s'érode** : les bugs ne sont pas corrigés, le code se fige à
nouveau (un **nouveau** *legacy*), et l'on a simplement échangé un *legacy* VB6 contre un *legacy* .NET
**mal compris**.

La migration représente surtout un énorme **transfert de connaissances** : les idiomes .NET, les
**décisions** prises, les **frontières d'interop**, et les **pièges** débusqués. Si ce savoir vit
uniquement dans la tête de l'équipe de migration (ou, pire, d'un prestataire externe qui s'en va), il est
**perdu**.

> 🎯 **Le risque n°1 de durabilité est le silo de connaissances** (*bus factor*). Cette section consiste
> à **transférer délibérément** ce savoir **et** à le **capturer** dans une documentation utile.

---

## Ce qui doit passer

```
   CE QUI DOIT PASSER                       COMMENT                       RÉSULTAT
   ─────────────────                        ───────                       ────────
   • Compétence .NET / VB.NET               Documentation vivante
   • Décisions de migration (ADR)           (dans le dépôt, versionnée)
   • Carte de l'interop (RCW/CCW, x86)      + Binôme / mentorat       ──►   Équipe
   • Pièges silencieux (annexe B)           + Formation                     autonome,
   • Harnais golden master (17.3)           + Revues de code                capable de
   • Runbooks build / déploiement / rollback + Période de recouvrement      MAINTENIR
   • Feuille de route (cohabitation, mod. 20)                               ET FAIRE ÉVOLUER
```

Le détail de chaque élément :

- **La compétence .NET elle-même.** Si l'équipe de maintenance vient de VB6, elle a besoin d'une réelle
  **fluence .NET** : le modèle CLR/GC (vs VBVM), la discipline `Option Strict`, `IDisposable`/`Using`,
  génériques/LINQ, exceptions structurées, la bibliothèque du framework. Ce n'est pas « lire le nouveau
  code » : c'est un **changement de compétence**. (C'est le versant humain de tout ce qu'enseigne cette
  formation : les mainteneurs ont besoin, au minimum, du socle conceptuel des parties 1-6.)
- **Les décisions de migration et leur *pourquoi*.** Pourquoi **4.7.2** (le « pont », modules 1.2/20) et
  pas encore le .NET moderne ; **où** l'interop a été employée et pourquoi (modules 16.1/16.2), ce qui est
  **temporaire** (l'échafaudage de cohabitation à retirer, 18.2) ; **où** le liage tardif a été
  délibérément **isolé** (16.3/17.1) ; **où** des comportements VB6 ont été **préservés** (équivalence)
  vs **corrigés** sciemment (re-baselinés, 17.3). Un **journal de décisions d'architecture** (ADR) est la
  façon idiomatique de capturer ces « pourquoi ».
- **La carte de l'interop.** Ce qui dépend **encore** de COM, où sont les **coutures**, la contrainte
  **x86** (et pourquoi), les exigences d'**enregistrement** (18.1), la discipline de **libération** des
  objets (16.1). La partie la plus délicate à maintenir et à déboguer **mérite une carte claire**.
- **Les pièges silencieux (annexe B).** Le savoir chèrement acquis sur les différences de comportement
  **silencieuses**. En étendant le code, un mainteneur peut **les réintroduire** (un `ByRef`/`ByVal`, un
  `Currency`→`Double`, une hypothèse de tableau 1-based). Il lui faut le **catalogue** (annexe B) et la
  **vigilance** : la migration les a traqués (17.4), la documentation doit en **prévenir le retour**.
- **Le harnais *golden master* (17.3).** Sans doute l'**artefact le plus précieux** à transmettre :
  l'équipe doit savoir qu'il **existe**, comment l'**exécuter**, comment l'**étendre**, et la discipline
  de **valider chaque changement** contre lui. C'est le **filet de sécurité permanent** de toute
  évolution future — pas seulement de la migration. À transmettre avec son **corpus** et le **savoir-faire**
  (ajouter des scénarios, re-baseliner sciemment).
- **Les runbooks build / déploiement / *rollback*.** Comment produire une *release*, le packaging et
  l'**enregistrement COM** (bitness, 18.1), comment déployer une version, et la **procédure de
  *rollback*** (18.3) — des modes opératoires **actionnables**.
- **La feuille de route.** Le travail **différé** : retrait de l'échafaudage de cohabitation (18.2),
  achèvement d'`Option Strict` résiduel, et le **second saut** optionnel vers .NET moderne (module 20).
  L'équipe doit savoir que la migration n'est pas une impasse mais a un **avenir planifié**.

---

## Comment transférer (au-delà d'« écrire des docs »)

La connaissance se transmet par **plusieurs canaux complémentaires** :

- **Documentation vivante** — l'artefact **durable**, mais *utile* (voir plus bas). À garder **dans le
  dépôt**, **versionnée** (Git, module 6.5), pour qu'elle reste à jour.
- **Binôme / mentorat / observation** — le savoir passe le mieux **de personne à personne** : l'équipe de
  migration fait équipe avec l'équipe de maintenance, qui réalise de **vraies tâches** sous supervision
  avant de voler de ses propres ailes — crucial pour le changement de compétence .NET et les subtilités
  d'interop.
- **Formation** — une montée en compétence .NET/VB.NET **formelle** pour les développeurs VB6 en
  transition (le socle conceptuel — dont cette formation fait partie).
- **Revues de code** — pendant la passation, l'équipe de maintenance commit et l'équipe de migration
  **relit** : transfert des standards et **détection** des pièges réintroduits.
- **Une période de recouvrement** — l'équipe de migration **ne disparaît pas** à la bascule : un
  **chevauchement** permet de répondre aux questions **en contexte**. ⚠️ **Particulièrement critique si la
  migration a été faite par des prestataires externes** — l'échec classique étant : les consultants
  partent, le savoir part avec eux.

---

## Documenter **utile**, pas exhaustif

> ⚖️ **Une documentation de 500 pages que personne ne lit est aussi inutile qu'une absence de
> documentation** — et, pire, elle **pourrit** si on ne la met pas à jour, devenant **trompeuse**.

Le bon équilibre : documenter **ce qui est nécessaire** — l'**aperçu d'architecture**, la **carte de  
l'interop**, les **ADR** (les « pourquoi »), les **runbooks** (build/déploiement/*rollback*), la
**référence des pièges** (pointer vers l'annexe B), et le **guide du harnais *golden master***. Le tout  
**près du code**, **dans le dépôt**, **versionné** — et dont la **mise à jour fait partie du « terminé »**
de chaque évolution future.

> 🧭 **Les annexes de cette formation sont précisément le matériel de référence à transmettre** : **B**
> (pièges), **E** (checklist), **F** (golden master), **G** (glossaire), **H** (ressources).

---

## ⚠️ Pièges à connaître

- **Silo de connaissances / *bus factor*** : le savoir dans quelques têtes (ou un prestataire qui part) —
  le **risque n°1** de durabilité. ⚠️
- **Documentation qui pourrit** : écrite une fois, jamais mise à jour → trompeuse. À **versionner** dans
  le dépôt et à **maintenir**. ⚠️
- **Sur-documentation** : un pavé que personne ne lit. Documenter le **nécessaire**, pas le tout.
- **Pas de période de recouvrement** : l'équipe de migration s'en va à la bascule, les mainteneurs
  **abandonnés**.
- **Mainteneurs non réellement formés** : ils lisent le code mais ne comprennent pas le modèle .NET →
  réintroduisent des idiomes VB6 ou des pièges, ou restent paralysés. Une **vraie** montée en compétence,
  pas une simple remise de documents.
- **Perte de la discipline *golden master*** : le harnais existe, mais l'équipe ne l'exécute pas sur
  chaque changement → les régressions et les pièges silencieux **reviennent**.
- **Considérer la bascule comme la fin** : la valeur dépend de l'**après** ; négliger la passation
  **gâche** tout l'effort.

---

## En résumé

- La valeur d'une migration se réalise **après** la bascule, sur les années de **maintenance** : sans
  transfert, on troque un *legacy* VB6 contre un *legacy* .NET **mal maîtrisé**. Le risque n°1 est le
  **silo de connaissances**.
- **Ce qui doit passer** : la **compétence .NET**, les **décisions** (ADR : les « pourquoi »), la **carte
  de l'interop** (RCW/CCW, x86, enregistrement), les **pièges** (annexe B), le **harnais *golden
  master*** (l'artefact le plus précieux), les **runbooks** (build/déploiement/*rollback*), et la
  **feuille de route** (cohabitation à retirer, module 20).
- **Comment** : **documentation vivante** (dans le dépôt, versionnée), **binôme/mentorat**, **formation**,
  **revues de code**, et une **période de recouvrement** — surtout face à des **prestataires externes**.
- **Documenter utile, pas exhaustif** : le nécessaire, près du code, **maintenu** (les annexes B/E/F/G/H
  en sont le socle).

---

## Le module 18 — et le parcours cœur — en perspective

La mise en production est **complète** et **durable** : l'application est **empaquetée** (18.1),
**déployée progressivement** par cohabitation (18.2), **basculée** sous trois garde-fous (18.3), et son
savoir est **transféré** (18.4).

Et avec ce module s'achève le **parcours cœur** de la formation : on a **cadré** la migration (partie 1),
**préparé le terrain** (partie 2), migré le **langage** (partie 3), l'**interface** (partie 4), les  
**données et l'interopérabilité** (partie 5), puis **validé, fiabilisé et déployé** (partie 6). Le fil
rouge tenu de bout en bout : *le danger n'est pas le code qui ne compile pas, c'est celui qui compile mais  
ne se comporte plus pareil* — et l'on s'est outillé, à chaque étape, pour le maîtriser.

> ➡️ **Reste la partie 7 — *IA, suite de parcours et ressources*** : comment l'**IA** assiste la
> migration (module 19, transversal), et surtout **le second saut** optionnel de 4.7.2 vers **.NET moderne**
> (module 20), à n'envisager **qu'une fois la présente migration stabilisée** — deux sauts **séparés**.

---

🏷️ **Indicateurs du chapitre** : transfert / durabilité · transmet notamment le *golden master* (17.3), la carte d'interop (16) et le catalogue des pièges (annexe B) ; ouvre la partie 7 (modules 19 et 20)
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026 · annexes B / E / F / G / H comme matériel de référence

⏭️ [Migrer avec l'aide de l'IA](/19-migration-ia/README.md)
