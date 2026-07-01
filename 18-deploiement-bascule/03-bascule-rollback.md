🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.3 — Stratégie de bascule : pilote, *parallel run*, plan de *rollback*

**Orchestrer le moment où les utilisateurs passent réellement à la nouvelle application — le point le plus risqué de toute la migration — à l'aide de trois garde-fous en couches : limiter qui est exposé, vérifier que le nouveau égale l'ancien sur des données réelles, et garder une voie de retour testée.**

> 🎯 Le **cœur de la gestion du risque** du déploiement. Le module 18.2 a posé la **topologie** de
> cohabitation et le déploiement par tranches ; cette section traite la **bascule elle-même** — l'événement
> où le nouveau système devient celui que l'on utilise. Trois techniques classiques la **dé-risquent** :
> le **pilote**, le ***parallel run*** et le **plan de *rollback***.

---

## La bascule : le moment de vérité, dé-risqué par trois contrôles

La **bascule** est le point où les utilisateurs commencent à se servir du **nouveau** système à la place  
de l'ancien (pour une **tranche**, ou pour l'application **entière**). C'est le moment le plus risqué.  
Trois garde-fous, **complémentaires** (des **couches**, pas des alternatives) :

- le **pilote** **limite** qui est exposé ;
- le ***parallel run*** **vérifie** que le nouveau égale l'ancien avant de s'engager ;
- le **plan de *rollback*** est le **filet** si tout échoue malgré tout.

> 🔁 Ces techniques s'appliquent à **deux échelles** : la bascule de **chaque tranche** (pendant le
> déploiement incrémental, module 18.2) **et** la **bascule finale** (retrait du dernier VB6). Mêmes
> techniques ; enjeux croissants.

---

## Le pilote : limiter le rayon d'impact

Déployer le nouveau (système ou tranche) auprès d'un **sous-ensemble restreint et représentatif** — un
**groupe pilote** d'utilisateurs, un site, un service —, le reste **restant sur l'ancien**.

- **But** : **contenir** le rayon d'impact. En cas de problème, seul le groupe pilote est touché, et l'on
  **apprend avant** d'exposer tout le monde.
- **Sélection** : représentatif **mais tolérant** — des utilisateurs qui exercent de **vrais** flux mais
  peuvent absorber des accrocs (pas le site le plus critique/à plus fort volume pour un premier pilote ;
  pas non plus un bac à sable qui ne reflète pas la réalité). Des utilisateurs **engagés**, qui
  **remontent** les incidents.
- **Suivi rapproché** : erreurs, performance, ressources (module 17.6), retours utilisateurs. Définir
  une **durée** représentative (couvrant, si pertinent, un **cycle métier** complet — une clôture
  mensuelle, par exemple) et des **critères de succès** **avant** d'élargir.
- **Élargir par paliers** : pilote → groupe plus large → tous. Une **montée progressive** (déploiement
  « en anneaux »/*canary*), pas un « pilote puis tout le monde ».

---

## Le *parallel run* : comparer l'ancien et le nouveau en réel

Faire tourner **simultanément** l'ancien système VB6 **et** le nouveau .NET sur les **mêmes entrées  
réelles**, en production, et **comparer leurs sorties**. L'ancien reste le **système de référence** ; le  
nouveau tourne à côté, et l'on vérifie qu'il produit des résultats **identiques**.

> 🔗 **C'est le prolongement, *en production*, du *golden master*** (module 17.3). Au lieu d'un corpus de
> test capturé, on valide sur des **données et une charge réelles** — ce qui **rattrape** les divergences
> que le corpus avait manquées (les « angles morts » des limites honnêtes du 17.3).

Plusieurs variantes selon le coût et le risque tolérés :

| Variante | Principe | Compromis |
|----------|----------|-----------|
| ***Parallel run* complet** | Les deux systèmes traitent **toutes** les transactions ; sorties comparées | Confiance **maximale**, coût **maximal** (double traitement, réconciliation) |
| **Mode *shadow* (ombre)** | Le nouveau traite les mêmes entrées **sans que ses sorties soient utilisées** — seulement journalisées et **comparées** à l'ancien (live) | Risque opérationnel **faible** (le nouveau ne peut pas casser la prod) ; excellent pour valider |
| ***Parallel run* échantillonné** | Comparer un **échantillon** de transactions plutôt que toutes | Moins coûteux, confiance moindre |

- **La comparaison** suit la **même discipline** que le module 17.3 : normaliser **uniquement** le bruit
  prouvé non pertinent, **jamais** un écart de **valeur** ; **chaque** divergence **investiguée** et
  **classée** (bug → corriger ; différence justifiée → documenter). Pas de troisième catégorie « on
  ignore ».
- **Durée** : assez longtemps pour couvrir des cas représentatifs / un **cycle métier**, jusqu'à **zéro**
  divergence (ou toutes comprises).

> ⚖️ **Coût.** Le *parallel run* (surtout complet) peut être **lourd** (double travail, réconciliation) :
> c'est l'étalon-or des bascules à fort enjeu, mais on **pèse coût vs risque**. Le mode **shadow** offre
> souvent le meilleur rapport validation/risque.

---

## Le plan de *rollback* : la voie de retour

Un plan **testé** pour **revenir à l'ancien système** si le nouveau échoue après la bascule.
**Non négociable.**

> ⚠️ **Un *rollback* n'est PAS « remettre l'ancien exécutable ».** Il doit couvrir les **données et
> l'état** : si le nouveau système a **écrit des données** (nouvelles lignes, nouveau schéma, nouveau
> format) entre la bascule et la décision de retour, remettre les binaires ne suffit pas — il faut
> **traiter ces données** pour que l'ancien système, restauré, reste **cohérent** (et que le travail fait
> dans le nouveau ne soit ni perdu ni ignoré sans décision). À relier à la **compatibilité des données**
> du module 18.2 et à l'accès aux données du module 15.

Les composants d'un plan de *rollback* :

| Composant | Rôle |
|-----------|------|
| **Déclencheurs** | Conditions **prédéfinies** imposant le retour (erreurs critiques, corruption de données, performance sous seuil, métier bloqué) — décidés **à l'avance** (module 3.6), pas dans la panique |
| **Procédure** | Les **étapes** réelles (revenir aux binaires/enregistrements, **restaurer l'état des données**, communiquer), **documentées** et **répétées** |
| **Fenêtre** | **Combien de temps** le retour reste possible — au-delà, on est **engagé** (correction en avant uniquement) |
| **Sauvegardes** | **Points de restauration** (sauvegardes/instantanés de base) **avant** la bascule |

> 🧪 **Testez le *rollback*.** Comme l'installeur sur machine propre (module 18.1), **répétez** le retour
> arrière dans un environnement hors production : un plan de *rollback* **jamais éprouvé** n'est qu'une
> **fausse sécurité**.

---

## Orchestrer la bascule

En réunissant les trois contrôles, le déroulé type (d'une tranche ou de la bascule finale) :

```
  Préparer ──► Pilote ──► Élargir ──► Parallel run ──► BASCULE ──► Stabiliser ──► Décommissionner
  backups,     groupe     par         ancien+nouveau   (nouveau =   suivi          retrait de
  rollback     restreint  paliers     comparés         référence)   (17.6)         l'ancien
  testé,                                                   │
  critères                                                 ▼
  (3.6)        └────────── Fenêtre de rollback : retour possible ──────────┘
                          (déclencheurs prédéfinis · données comprises)
```

1. **Avant** : sauvegardes/instantanés ; **plan de *rollback* testé** ; **critères** de *go/no-go* et
   **déclencheurs** définis (module 3.6).
2. **Pilote** : déployer au groupe pilote ; (éventuellement) *parallel run*/shadow pour ce groupe ;
   suivre les critères.
3. **Élargir** : montée par **paliers** si les critères sont tenus.
4. ***Parallel run*** (enjeux élevés) : les deux systèmes en réel, comparés, jusqu'à divergences résolues.
5. **Bascule** : le nouveau devient le **système de référence** ; l'ancien **reste disponible** dans la
   **fenêtre de *rollback***.
6. **Stabiliser** : surveiller les indicateurs (17.6) ; **prêt à revenir** dans la fenêtre.
7. **Décommissionner** : une fois **stable** au-delà de la fenêtre, retirer l'ancien (et, pour une
   migration incrémentale, l'échafaudage de cohabitation — la **sortie** du module 18.2).

> 📣 **Communiquer** avant, pendant, après : les utilisateurs doivent savoir **ce qui change**, **quoi
> surveiller**, **comment signaler** un incident, et qu'un **repli** existe (à relier au transfert du
> module 18.4). **Choisir le moment** : une fenêtre à faible risque (heures creuses), **pas** une période
> métier critique (clôture annuelle…) — sauf si c'est précisément ce que l'on veut valider.

---

## Bascule par tranches vs bascule finale

- **Migration incrémentale** (module 18.2) : chaque **bascule de tranche** emploie ces techniques **à
  petite échelle** (enjeu faible par tranche, la cohabitation laissant l'ancien et le nouveau coexister).
  La **bascule finale** (retrait du dernier VB6) est le seul grand moment — mais, à ce stade, on a déjà
  **répété** ces techniques de nombreuses fois : même la bascule finale est dé-risquée.
- **Migration *big-bang*** (si choisie, module 3.4) : **une seule** bascule à fort enjeu, sans répétition
  incrémentale — ces techniques y sont **encore plus** critiques (et plus difficiles).

---

## ⚠️ Pièges à connaître

- **Basculer sans *rollback* testé** : un plan jamais répété peut ne pas fonctionner le jour J. ⚠️
- **Un *rollback* qui ignore les données** : remettre l'EXE en laissant des données divergées →
  incohérence / perte. ⚠️
- **Pas de déclencheurs prédéfinis** : décider du retour dans la panique, trop tard, ou jamais (« on
  pousse quand même », coût irrécupérable). ⚠️
- **Pilote trop petit/non représentatif** (fausse confiance) ou **trop large** (rayon d'impact non
  limité).
- **Normalisation du *parallel run* qui masque des bugs** : même piège qu'au 17.3 — ne jamais masquer un
  écart de **valeur**.
- **Basculer à une période métier critique** : amplifie le risque.
- **Décommissionner l'ancien trop tôt** : perdre l'option de *rollback* avant que le nouveau ne soit
  prouvé stable.
- **Absence de communication** : utilisateurs pris au dépourvu, sans canal de remontée ni repli.

---

## En résumé

- La **bascule** est le moment le plus risqué ; on la dé-risque par **trois garde-fous en couches** :
  **pilote**, ***parallel run***, **plan de *rollback***.
- **Pilote** : déployer d'abord à un **sous-ensemble représentatif mais tolérant**, suivre des critères,
  puis **élargir par paliers**.
- ***Parallel run*** : faire tourner ancien et nouveau **en réel** et **comparer** — le prolongement **en
  production** du *golden master* (17.3), avec ses variantes (complet, **shadow**, échantillonné) et la
  **même discipline** de comparaison.
- **Plan de *rollback*** : **non négociable**, couvre les **données et l'état** (pas seulement l'EXE),
  avec **déclencheurs prédéfinis** (3.6), **procédure répétée**, **fenêtre** limitée et **sauvegardes** —
  et **testé** d'avance.
- **Orchestration** : préparer → pilote → élargir → *parallel run* → bascule → stabiliser →
  décommissionner, le **retour** restant possible dans la **fenêtre** ; **communiquer** et **choisir le
  moment**.
- En **incrémental**, chaque bascule de tranche **répète** ces techniques, dé-risquant la **bascule
  finale** ; en **big-bang**, l'unique bascule les rend d'autant plus **critiques**.

> ➡️ **La bascule réussie ne clôt pas tout : pour que la migration *dure*, l'équipe doit comprendre le
> nouveau système. Place à la documentation et au transfert de compétences.** (Module 18.4)

---

🏷️ **Indicateurs** : bascule / gestion du risque · prolonge le *golden master* du module 17.3 en production ; lié à 3.4 (stratégie), 3.6 (critères/déclencheurs), 15 (données), 17.6 (indicateurs), 18.2 (cohabitation/données), 18.4 (communication)
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Documentation et transfert de compétences à l'équipe](/18-deploiement-bascule/04-documentation-transfert.md)
