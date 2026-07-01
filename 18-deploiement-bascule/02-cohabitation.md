🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.2 — Cohabitation VB6 / .NET pendant la transition (déploiement progressif) 🔗

**Faire coexister l'ancien et le nouveau en production, le temps de la migration : déployer des *tranches* migrées qui dialoguent avec le VB6 restant, plutôt que d'attendre 100 % de conversion pour un saut unique et risqué.**

> 🔗 Le versant **déploiement** de tout le travail d'interopérabilité du module 16. Grâce au **CCW**
> (module 16.2) et à la stratégie de cohabitation COM (module 3.5), on peut livrer en production des
> **morceaux** migrés qui coexistent avec le VB6 pas encore converti. Cette section ne ré-explique pas
> la **mécanique** RCW/CCW (modules 16.1/16.2) ni la **bascule** (module 18.3) : elle traite la
> **topologie** en production et le **séquencement** des déploiements.

---

## L'idée : déployer des tranches, pas un saut unique

Plutôt que « tout migrer, puis basculer d'un coup », on déploie des **tranches migrées  
incrémentalement**, l'ancien VB6 et le nouveau .NET **coexistant en production** et communiquant par  
COM. C'est le **paiement, côté déploiement**, de la migration incrémentale rendue possible par le CCW :  
on livre un module migré pendant que le reste de l'application est encore en VB6, parce que le VB6
**appelle** le nouveau .NET (CCW) et que le .NET appelle le VB6/COM restant (RCW).

---

## Pourquoi déployer progressivement

En termes de déploiement, l'approche par tranches offre :

- **Un rayon d'impact réduit** — chaque livraison touche une **tranche**, pas toute l'application. Un
  problème reste **contenu** et facile à localiser.
- **De la valeur et du retour plus tôt** — les parties migrées atteignent les utilisateurs **plus vite**,
  et l'on apprend de leur comportement réel **tôt**.
- **Un risque étalé dans le temps** — beaucoup de petits pas **réversibles** plutôt qu'un seul **saut
  irréversible**.
- **Une validation incrémentale en production** — chaque tranche peut avoir son **pilote** et son
  ***parallel run*** (module 18.3) sur données réelles avant la suivante.

C'est l'expression, au moment du déploiement, de la philosophie **incrémentale** de toute la formation
(module 3.4).

---

## Comment l'ancien et le nouveau coexistent : les topologies

```
  COHABITATION EN PROCESSUS (CCW/RCW)          COHABITATION HORS PROCESSUS (serveur COM EXE / IPC)
  ─────────────────────────────────           ──────────────────────────────────────────────────
   ┌──────────────────────────────┐             ┌───────────────┐         ┌───────────────┐
   │ Processus hôte VB6 (32 bits)  │             │ Processus VB6 │   COM   │ Processus .NET│
   │  ┌───────┐  CCW   ┌────────┐  │             │   (32 bits)   │ ◄─────► │ (32 ou 64)    │
   │  │  VB6  │ ─────► │  .NET  │  │             └───────────────┘  hors   └───────────────┘
   │  │       │ ◄───── │  (CLR) │  │                                process
   │  └───────┘  RCW   └────────┘  │             Découplé · isolé · franchit la bitness
   │  Rapide · couplé · 32 bits    │             · plus d'overhead
   └──────────────────────────────┘
```

### En processus (via COM : CCW/RCW)

Le processus **hôte VB6** charge les composants .NET **en son sein** via COM : il héberge le CLR et  
appelle les classes .NET comme des objets COM (CCW). À l'inverse, un nouvel exécutable .NET peut charger  
les composants VB6/COM restants en processus (RCW).

- ✅ **Tight et rapide** (pas de franchissement de processus).
- ⚠️ **Couplé** : un défaut non géré dans un monde affecte le **processus partagé** ; et la **bitness
  doit correspondre** → tout le processus reste **32 bits** tant que le VB6 host existe (module 16.1).

C'est la topologie la plus courante d'une migration **classe par classe**.

### Hors processus (serveur COM EXE / IPC)

Les parties VB6 et .NET tournent comme des **processus distincts** communiquant par COM (serveur EXE  
hors processus) ou une autre IPC.

- ✅ **Découplé** : domaines de défaillance **indépendants** ; et — point important — cela **franchit la
  contrainte de bitness** (un processus .NET 64 bits peut dialoguer avec un serveur COM 32 bits hors
  processus via un substitut, module 16.1).
- ⚠️ **Plus d'overhead** (marshaling inter-processus) et plus de complexité.

### Applications côte à côte sur **données partagées**

Parfois, la cohabitation est au niveau des **données** : l'application VB6 et la nouvelle application
.NET sont des **exécutables séparés** qui partagent la **même base / les mêmes fichiers**, chacun
prenant en charge une **partie** du flux de travail. Topologie plus **grossière** (migration  
fonctionnalité par fonctionnalité), où le **magasin de données** est le point d'intégration. Elle exige  
une attention particulière à la **compatibilité de schéma** (voir plus bas).

---

## Séquencer les déploiements

L'**ordre de migration** (choisi à la stratégie, module 3.4) pilote l'**ordre de déploiement**. Quelques  
principes :

- **Tranches à faible risque et faible couplage d'abord** — pour bâtir l'élan, la confiance, et
  **prouver en production** que la plomberie de cohabitation fonctionne ; réserver le plus risqué/couplé
  pour plus tard, équipe et outillage **rodés**.
- **Des unités cohérentes et indépendamment déployables**, chacune dotée d'une **interface stable** (le
  contrat CCW du module 16.2).
- **Compatibilité de versions** — quand une tranche est déployée, l'hôte VB6 et le composant .NET doivent
  être des **versions compatibles**. C'est la discipline des **GUID/interfaces figés** (module 16.2) qui
  rend cela gérable : le **contrat COM reste fixe**, donc l'ancien et le nouveau interopèrent. On déploie
  par **ensembles compatibles**.

---

## Opérer la cohabitation en production

- **Enregistrement par tranche** — chaque composant CCW nouvellement déployé doit être **enregistré**
  (`regasm`, module 18.1) sur la cible ; un composant VB6/COM retiré peut devoir être **désenregistré**.
  L'installeur gère l'enregistrement **tranche par tranche**.
- **Verrouillage de *bitness*** — la cohabitation **en processus** impose le **32 bits** (hôte VB6) :
  cette contrainte **persiste** tant que le VB6 n'a pas disparu (la cohabitation **hors processus** peut
  la relâcher).
- **Diagnostics à la frontière** — déboguer/journaliser un processus **mixte** VB6+.NET est plus
  difficile ; les erreurs traversent le pont COM sous forme de **HRESULT** (module 16.2). Prévoyez une
  **journalisation** qui **couvre les deux mondes**, pour suivre une requête de bout en bout à travers la
  frontière.
- **Performance** — chaque franchissement de frontière a un **coût** (module 16.5) : une cohabitation
  **bavarde** (nombreux appels fins inter-mondes) sera lente. **Concevez des interfaces grossières** à la
  frontière (à relier à la revue de performance du module 17.6).
- **Durée de vie** — la libération des objets COM à travers la frontière (discipline RCW du module 16.1 ;
  durée de vie des objets CCW vus de VB6) doit fonctionner dans le **processus mixte**.

---

## La compatibilité des données partagées

Si l'ancien et le nouveau **partagent** une base ou des fichiers (surtout dans la topologie côte à côte),  
le **schéma doit servir les deux** pendant la transition. Toute évolution (nouvelles colonnes, etc.) doit  
rester **rétrocompatible**, afin que les parties **encore en VB6 continuent de fonctionner**.

> ⚠️ **Évitez les changements de données *cassants*** sur le magasin partagé tant que des consommateurs
> VB6 subsistent. (À relier au module 15 sur l'accès aux données, et à la dimension **données** du *plan
> de rollback* du module 18.3 : un retour arrière doit lui aussi rester cohérent côté données.)

---

## Une cohabitation a une **fin**

La cohabitation est **temporaire** — le thème récurrent « **pont, pas destination** » du module 16. À  
mesure que la migration progresse, l'empreinte VB6 **se réduit** jusqu'à disparaître ; on peut alors
**retirer l'échafaudage d'interop** :

- supprimer les **enregistrements COM** et la plomberie CCW/RCW devenue inutile ;
- **lever la contrainte x86** imposée par la cohabitation en processus ;
- envisager, enfin, le **second saut** optionnel vers .NET moderne (module 20).

> 🎯 **Planifiez la sortie.** Une cohabitation qui s'éternise devient une **architecture permanente** —
> avec sa complexité, son coût, et le blocage du saut vers .NET moderne. Le but reste de **retirer** les
> frontières, pas de les pérenniser.

---

## ⚠️ Pièges à connaître

- **Laisser la cohabitation devenir permanente** : transitoire par nature, à **retirer** selon un plan. ⚠️
- **Verrouillage de *bitness*** : la cohabitation **en processus** épingle tout en **x86** jusqu'au
  retrait du VB6. ⚠️
- **Casser les données partagées** : des changements de schéma qui brisent les consommateurs **encore
  VB6**. ⚠️
- **Frontières bavardes** : des appels fins inter-mondes ruinent la performance → **interfaces
  grossières**.
- **Décalage de versions** : déployer des versions ancien/nouveau incompatibles → s'appuyer sur le
  **contrat COM figé** (16.2) et déployer par **ensembles compatibles**.
- **Dérive d'enregistrement** : une livraison qui oublie d'enregistrer/désenregistrer un composant laisse
  la machine dans un état **incohérent** (« Classe non enregistrée » ou inscriptions périmées).
- **Couplage de crash** (en processus) : un défaut non géré peut abattre le **processus partagé** — le
  **hors processus** isole, au prix de l'overhead.

---

## En résumé

- La cohabitation permet un **déploiement progressif** : on livre des **tranches** migrées qui coexistent
  en production avec le VB6 restant (via CCW/RCW), au lieu d'un **saut unique** — c'est le versant
  déploiement des modules 3.5 et 16.
- **Bénéfices** : rayon d'impact réduit, valeur et retour plus tôt, risque étalé, validation
  incrémentale en production.
- **Trois topologies** : **en processus** (CCW/RCW — rapide, couplé, **32 bits**), **hors processus**
  (COM EXE/IPC — découplé, isolé, **franchit la bitness**, plus d'overhead), et **applications côte à
  côte sur données partagées** (granularité fonctionnelle).
- **Séquencer** : tranches à faible risque d'abord, unités cohérentes à **interface stable**,
  **versions compatibles** grâce au **contrat COM figé** (16.2).
- **Opérer** : enregistrement **par tranche** (18.1), **verrouillage 32 bits** (en processus),
  diagnostics et **journalisation couvrant les deux mondes**, **interfaces grossières** pour la
  performance, et **compatibilité de schéma** sur les données partagées.
- **Planifier la sortie** : la cohabitation est **temporaire** ; son retrait lève la contrainte x86 et
  ouvre le module 20.

> ➡️ **La cohabitation rend le déploiement progressif possible ; reste à orchestrer la bascule
> elle-même : pilote, *parallel run* et plan de *rollback*.** (Module 18.3)

---

🏷️ **Indicateurs** : 🔗 Interop (cohabitation en production) · concrétise les modules 3.5 et 16.1/16.2 ; lié à 15 (données), 16.5/17.6 (performance), 18.1 (enregistrement), 18.3 (bascule), 20 (sortie)
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026 · interop COM (RCW/CCW)

⏭️ [Stratégie de bascule : pilote, *parallel run*, plan de *rollback*](/18-deploiement-bascule/03-bascule-rollback.md)
