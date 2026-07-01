🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.4 Stratégies : *big-bang* vs incrémentale ⭐

> **Module 3 — Évaluer l'existant et choisir une stratégie**
> L'inventaire (3.1), la cartographie (3.2) et l'estimation (3.3) ont **éclairé le terrain**.
> Vient maintenant **la décision la plus engageante de toute la migration** : la basculer d'un
> bloc, ou la mener par étapes ? Ce choix structure le coût, la durée, le risque et
> l'organisation de l'équipe.

---

## 🧭 La décision qui structure tout le projet

Il n'y a, au fond, que **deux grandes manières** de migrer une application VB6 :

- **Tout migrer d'un coup**, puis basculer en une fois — l'approche **big-bang** ;
- **Migrer morceau par morceau**, en faisant cohabiter l'ancien et le nouveau pendant la
  transition — l'approche **incrémentale**.

Tout le reste (outils, ordre des modules, plan de tests, plan de bascule) découle de ce choix.  
C'est pourquoi il ne se prend **qu'après** l'inventaire, la cartographie et l'estimation : décidé
à l'aveugle, c'est l'erreur la plus coûteuse du projet.

> ⚠️ **Il n'existe pas de bonne réponse universelle.** Une petite application interne tolère
> très bien un big-bang ; un logiciel critique de production exige presque toujours une bascule
> progressive. La criticité, les délais, la taille de l'équipe et la dette technique pèsent
> autant que des considérations purement techniques. L'erreur n'est pas de choisir telle ou telle
> approche — c'est de la choisir **sans avoir fait le travail des sections précédentes**.

---

## 💥 L'approche *big-bang*

Le principe : on migre **l'intégralité** de l'application (en une campagne dédiée), on valide le  
résultat complet, puis on **remplace** l'ancienne version par la nouvelle en une seule bascule.  
Pendant toute la durée des travaux, l'application VB6 reste en production **inchangée** ; la  
version .NET se construit **à côté**, et ne prend le relais qu'une fois prête.

### Quand elle convient

- **Petites et moyennes applications** : le volume et la durée des travaux restent maîtrisables.
- **Faible couplage aux dépendances problématiques** : peu ou pas de points durs (3.3), pas de
  composant tiers qui impose une cohabitation.
- **Périmètre stable et bien cerné** : on peut **geler les évolutions** (module 5.6) le temps de
  la migration sans pénaliser l'activité.
- **Tolérance à une bascule unique** : l'organisation accepte un basculement « tout ou rien » à
  une date donnée.

### Ses avantages

- **Conceptuellement simple** : une cible, une trajectoire, pas de mécanique de cohabitation à
  construire ni à maintenir.
- **Pas de pont temporaire** : on n'investit pas dans des échafaudages (interop bidirectionnelle,
  doubles déploiements) qui seront jetés ensuite.
- **Cohérence du résultat** : tout le code arrive dans le même état, au même moment, sans phase
  hybride à gérer.
- **Fin nette** : une fois la bascule réussie, le VB6 disparaît — pas de dette de transition qui
  traîne.

### Ses risques

- **L'effet « tunnel »** : pendant toute la durée des travaux, **rien n'est livré**. Plus la
  migration est longue, plus le risque s'accumule sans retour intermédiaire, et plus
  l'estimation (3.3) — déjà incertaine sur la strate rouge — devient difficile à tenir.
- **Une bascule à fort enjeu** : tout se joue en une fois. Si un problème majeur surgit le jour
  J, c'est **toute** l'application qui est concernée. D'où l'importance vitale du plan de
  *rollback* (module 18.3).
- **Le périmètre gelé qui dérive** : sur un projet long, geler les évolutions fonctionnelles
  (module 5.6) devient politiquement et opérationnellement difficile. Si le gel cède, la cible
  bouge pendant qu'on la vise.
- **La validation massive d'un coup** : tester l'application entière en fin de course est plus
  lourd — et plus tardif — que valider des incréments au fil de l'eau.

> 💡 **Le big-bang n'est pas « la mauvaise approche ».** Sur le bon périmètre — une application
> petite à moyenne, sans points durs bloquants, dont on peut geler le périmètre — c'est souvent
> le choix **le plus rationnel** : il évite le coût de la cohabitation. Le danger n'est pas le
> big-bang en soi, mais le big-bang appliqué à un projet **trop gros ou trop critique** pour lui.

---

## 🧩 L'approche incrémentale

Le principe : on découpe l'application en **morceaux** (modules, sous-ensembles fonctionnels) et  
on les migre **un par un**, en maintenant à chaque étape une application qui **fonctionne** —  
faite d'une partie VB6 et d'une partie .NET qui **coexistent**. La bascule n'est plus un événement  
unique mais une **série de petits basculements**.

### Le moteur de l'incrémental : l'interopérabilité COM 🔗

Cette approche n'est possible que parce que **VB6 et .NET peuvent se parler**, dans les deux sens,  
via COM :

- du **.NET appelle du VB6** existant via un **RCW** (*Runtime Callable Wrapper*) ;
- du **VB6 appelle du .NET** déjà migré via un **CCW** (*COM Callable Wrapper*).

C'est ce second mécanisme — exposer du .NET à du VB6 — qui est la **clé technique** de la  
migration incrémentale : il permet à la partie déjà migrée de continuer à être utilisée par la  
partie restée en VB6. La **section 3.5** est entièrement consacrée à cette cohabitation, et le
**module 16** en détaille la mécanique (RCW/CCW, *early/late binding*).

> 💡 Sans interopérabilité, l'incrémental serait impossible : on ne pourrait pas faire tourner une
> moitié d'application. C'est COM qui rend la coexistence viable — et donc l'approche progressive
> envisageable.

### Quand elle convient

- **Applications volumineuses ou critiques** : on réduit le risque en le **fractionnant**, et on
  livre de la valeur **en continu** plutôt qu'au bout du tunnel.
- **Présence de points durs** : un composant délicat (3.3) peut être isolé et traité à part, sans
  bloquer le reste de la migration.
- **Périmètre difficile à geler** : l'activité doit continuer à évoluer pendant la migration —
  l'incrémental compose mieux avec un existant qui bouge.
- **Découplage possible** : l'application se prête à un découpage en morceaux relativement
  indépendants (d'où l'intérêt de **découpler l'UI de la logique métier** en amont, module 5.3).

### Ses avantages

- **Risque fractionné** : chaque incrément est petit, validé et basculé séparément. Un problème
  reste **circonscrit** à un morceau.
- **Valeur livrée en continu** : on ne disparaît pas dans un tunnel ; des bénéfices arrivent au
  fil de l'eau, et l'on **apprend** d'un incrément à l'autre (les premiers calibrent l'estimation
  des suivants).
- **Compatible avec une activité vivante** : pas besoin d'un gel total et prolongé du périmètre.
- **Bascule dédramatisée** : une série de petits basculements réversibles, plutôt qu'un saut
  unique à très fort enjeu.

### Ses risques

- **La complexité de la cohabitation** : faire coexister deux mondes (VB6 + .NET) a un **coût
  propre** — interop à construire et maintenir, déploiement hybride, *marshaling* à la frontière
  (module 16.5). C'est un échafaudage **temporaire mais réel**.
- **Une durée totale potentiellement plus longue** : fractionner et maintenir le pont allonge
  souvent le calendrier global, même si le risque par étape est moindre.
- **Le risque d'enlisement** : une migration incrémentale **abandonnée à mi-parcours** laisse
  l'organisation avec une application **hybride durable** — le pire des deux mondes. La discipline
  de mener la migration **jusqu'au bout** est essentielle.
- **Le surcoût de l'interface VB6/.NET** : tant que la frontière existe, chaque appel qui la
  traverse paie le prix de l'interop (sérialisation, conversions). C'est transitoire, mais à
  surveiller.

> ⚠️ **Le piège de l'incrémental : ne jamais finir.** La coexistence est un **état de
> transition**, pas une destination. Une migration progressive sans engagement clair de la mener
> à son terme risque de se figer en architecture hybride permanente — coûteuse à maintenir et
> jamais aboutie. L'incrémental exige un **cap** et la volonté de le tenir.

---

## ⚖️ Big-bang ou incrémental : le tableau de décision

Aucune des deux approches n'est supérieure dans l'absolu. Le bon choix dépend du contexte —  
précisément celui que les sections 3.1 à 3.3 ont permis de mesurer.

| Critère | Penche vers **big-bang** | Penche vers **incrémental** |
|---------|--------------------------|------------------------------|
| **Taille de l'application** | Petite à moyenne | Grande |
| **Criticité en production** | Faible (outil interne, non vital) | Élevée (cœur de métier, indisponibilité coûteuse) |
| **Points durs (3.3)** | Peu ou pas | Nombreux ou bloquants |
| **Périmètre** | Gelable sans douleur | Doit continuer à évoluer |
| **Taille de l'équipe** | Réduite, mobilisable d'un bloc | Suffisante pour mener migration **et** maintien en parallèle |
| **Tolérance au risque** | Bascule unique acceptable | Risque à fractionner impérativement |
| **Découplage du code** | Faible (monolithe difficile à découper) | Bon (modules relativement indépendants) |
| **Horizon de temps** | Court, délai serré | Long, étalement acceptable |

> 💡 **Lire le tableau comme un faisceau, pas comme une formule.** Un seul critère tranche
> rarement. C'est la **convergence** de plusieurs signaux qui oriente : une grosse application
> critique, truffée de points durs, dont le périmètre ne peut pas geler, appelle clairement
> l'incrémental — quand une petite appli interne sans dépendance problématique appelle tout aussi
> clairement le big-bang.

---

## 🔀 Au-delà du binaire : les approches hybrides

La réalité est souvent moins tranchée que « tout d'un coup » contre « morceau par morceau ». Entre  
les deux extrêmes existent des **stratégies mixtes**, qui combinent les avantages des deux selon  
les zones de l'application :

- **Big-bang par sous-système** : découper l'application en grands sous-ensembles relativement
  autonomes, et mener un big-bang **sur chacun** — incrémental à l'échelle de l'application, mais
  big-bang à l'échelle du sous-système.
- **Cœur d'abord, périphérie ensuite** (ou l'inverse) : migrer en priorité la logique métier
  (souvent du `.bas` plus simple, module 5.3), en gardant l'UI sous interop, puis traiter les
  formulaires ; ou commencer par un module périphérique peu risqué pour **roder** la démarche.
- **Isoler les points durs** : traiter d'abord (ou réserver pour la fin) les zones rouges
  identifiées en 3.3 — composants tiers à remplacer, *control arrays*, reporting — en les sortant
  du flux principal de migration.

> 💡 La frontière entre « incrémental fin » et « big-bang par sous-système » est poreuse — et ce
> n'est pas grave. Ce qui compte n'est pas l'étiquette mais que l'approche soit **adaptée au
> contexte mesuré** et qu'elle soit **menée à son terme**.

---

## ⚠️ Les pièges du choix de stratégie

- **Choisir avant d'avoir mesuré.** Le piège fondamental : trancher la stratégie sans inventaire,
  cartographie ni estimation. La décision n'est défendable qu'**adossée** aux sections 3.1 à 3.3.
- **Croire qu'une approche est universellement meilleure.** Le big-bang n'est pas « risqué » dans
  l'absolu, l'incrémental n'est pas « propre » dans l'absolu. Chacun excelle sur le bon périmètre.
- **Sous-estimer le coût de la cohabitation.** L'incrémental n'est pas gratuit : l'interop est un
  échafaudage réel à construire et maintenir (module 16). À mettre en regard du risque qu'il
  réduit.
- **Sous-estimer le risque du big-bang sur un gros projet.** L'effet tunnel et la bascule unique
  deviennent ingérables au-delà d'une certaine taille ou criticité.
- **Lancer un incrémental sans s'engager à finir.** Le pire scénario : l'hybride permanent. Sans
  cap clair, l'incrémental se fige à mi-chemin.
- **Oublier que la stratégie conditionne le plan de bascule.** Big-bang → *rollback* global
  (module 18.3) ; incrémental → série de bascules réversibles (module 18.2). Le choix d'ici se
  paie là-bas.

---

## ✅ En résumé

- Deux grandes stratégies structurent toute la migration : le **big-bang** (tout migrer, puis
  basculer en une fois) et l'**incrémentale** (migrer morceau par morceau, en faisant coexister
  VB6 et .NET).
- **Aucune n'est supérieure dans l'absolu.** Le bon choix dépend de la **taille**, de la
  **criticité**, des **points durs**, de la possibilité de **geler le périmètre**, de la **taille
  de l'équipe** et de la **tolérance au risque** — tout ce que les sections 3.1 à 3.3 ont mesuré.
- Le **big-bang** est simple et sans pont temporaire, mais expose à l'**effet tunnel** et à une
  **bascule à fort enjeu** : idéal pour les applications **petites à moyennes**, peu critiques,
  sans points durs.
- L'**incrémentale** fractionne le risque et livre de la valeur en continu, au prix de la
  **complexité de la cohabitation** et du risque d'**enlisement** : adaptée aux applications
  **volumineuses ou critiques**. Son moteur technique est l'**interopérabilité COM** (3.5,
  module 16).
- Entre les deux existent des **approches hybrides** (big-bang par sous-système, cœur d'abord,
  isolement des points durs) — souvent les plus réalistes.
- Le piège central est de **choisir sans avoir mesuré** ; le piège de l'incrémental est de **ne
  jamais finir**.

---

⬅️ Section 3.3 — [Estimer l'effort](03-estimer-effort.md)  
➡️ Section 3.5 🔗 — [L'approche par interopérabilité COM : faire cohabiter VB6 et .NET](05-cohabitation-com.md)

---

**Juin 2026** · .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [L'approche par **interopérabilité COM** : faire cohabiter VB6 et .NET pendant la transition](/03-evaluer-strategie/05-cohabitation-com.md)
