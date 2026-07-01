🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3. Évaluer l'existant et choisir une stratégie ⭐

> **Partie 1 — Comprendre et cadrer la migration**
> Ce module fait le pont entre la *compréhension* des différences (module 2) et le
> *choix des outils* (module 4). Avant de migrer, il faut savoir **quoi** migrer,
> **de quoi cela dépend**, **combien cela coûte** et **selon quelle stratégie** l'aborder.

---

## 🧭 Pourquoi ce module est central

On ne migre bien que ce que l'on comprend. Une migration décidée sans inventaire préalable  
est une migration qui découvre ses surprises **en production** — au pire moment, et au prix  
fort. L'objectif de ce module est exactement inverse : transformer une impression vague
(« nous avons une grosse application VB6 ») en une connaissance précise et chiffrée
(« nous avons 142 formulaires `.frm`, 38 classes `.cls`, 11 contrôles OCX dont 3 ne sont plus
maintenus, et 6 modules saturés de `Variant` »).

Cette connaissance n'est pas un luxe documentaire : elle est ce qui rend la migration
**pilotable**. Les pièges silencieux décrits au module 2 — le `ByRef` devenu `ByVal`, les
entiers redimensionnés, la finalisation déterministe perdue — ne se répartissent pas  
uniformément dans le code. Ils se **concentrent** dans des endroits identifiables : modules  
gorgés de `Variant`, API truffées de passages par référence, classes qui s'appuient sur
`Class_Terminate`. L'inventaire et la cartographie des dépendances sont précisément ce qui
fait apparaître ces zones à risque **avant** qu'elles ne se transforment en régressions.

Enfin, le choix de **stratégie** (*big-bang* ou incrémentale) structure tout le projet :  
son coût, sa durée, son profil de risque et l'organisation de l'équipe. C'est la décision la  
plus engageante de toute la migration — et la plus coûteuse à prendre à l'aveugle. Ce module  
ne tranche pas à votre place : il vous donne les **critères** pour décider en connaissance de  
cause.

---

## 🎯 Ce que vous saurez faire à l'issue de ce module

- **Inventorier** méthodiquement les constituants d'un projet VB6 (formulaires, modules,
  classes, références, contrôles OCX, DLL, déclarations d'API) et savoir lire ce que chaque
  type de fichier vous apprend sur l'effort à venir.
- **Cartographier les dépendances** — composants COM/ActiveX, bases de données, bibliothèques
  tierces — pour distinguer ce que vous maîtrisez de ce qui vous échappe (un composant tiers
  non maintenu est un point dur, pas un détail).
- **Estimer l'effort** de façon réaliste à partir de la taille, de la complexité et des points
  durs repérés, plutôt qu'à partir d'une intuition.
- **Choisir entre une bascule globale et une migration incrémentale** en pesant la criticité de
  l'application, les délais, la taille de l'équipe et la tolérance au risque.
- **Comprendre l'interopérabilité COM comme le moteur de la cohabitation** : ce qui permet à du
  VB6 et à du .NET de coexister pendant la transition, et donc rend l'approche incrémentale
  possible.
- **Définir un périmètre, des critères de réussite et des indicateurs** clairs, pour savoir à
  tout moment où vous en êtes et quand la migration est « terminée ».

---

## 📚 Le parcours du module

| § | Sujet | Ce que vous y trouverez |
|---|-------|-------------------------|
| **3.1** | Inventaire | Recenser `.frm`, `.bas`, `.cls`, références, OCX, DLL et API — la photographie de départ |
| **3.2** | Cartographie des dépendances | COM, ActiveX, bases de données, composants tiers : ce dont l'application dépend vraiment |
| **3.3** | Estimer l'effort | Taille, complexité cyclomatique, points durs : du recensement au chiffrage |
| **3.4** | Stratégies *big-bang* vs incrémentale | Les deux grandes approches, leurs avantages, leurs risques |
| **3.5** 🔗 | Cohabitation par interopérabilité COM | Faire tourner VB6 et .NET côte à côte le temps de la transition |
| **3.6** | Périmètre, critères de réussite, indicateurs | Cadrer ce qui entre dans la migration et comment en mesurer le succès |

Le module suit une progression naturelle : on **observe** (3.1–3.2), on **chiffre** (3.3), on
**décide** (3.4–3.5), puis on **cadre** (3.6).

---

## ⚠️ Un principe à garder en tête

> Il n'existe **pas** de bonne stratégie universelle. Une petite application interne tolère un
> *big-bang* ; un logiciel critique de production exige presque toujours une bascule
> progressive avec cohabitation. La criticité, les délais, l'équipe disponible et la dette
> technique pèsent autant que des considérations purement techniques. L'erreur n'est pas de
> choisir telle ou telle approche — c'est de la choisir **sans avoir fait l'inventaire et la
> cartographie** des modules 3.1 et 3.2.

---

## 🔗 Prérequis et liens utiles

- **À lire avant** : module 1 (cadrage) et surtout module 2 (différences fondamentales) — c'est
  là que sont décrits les pièges que l'inventaire de ce module cherche à localiser.
- **Pour aller plus loin sur les outils** : module 4 (assistants, outils commerciaux,
  migration manuelle), qui réalise concrètement le travail que ce module aura permis de cadrer.
- **Sur l'interopérabilité COM en profondeur** : module 16 (RCW/CCW, *early/late binding*),
  pour la mécanique technique de la cohabitation évoquée en 3.5.
- **Sur la bascule en production** : module 18 (cohabitation, *parallel run*, *rollback*).
- **Outil pratique** : **Annexe E — Checklist de migration**, pour ne rien oublier lors de
  l'évaluation de l'existant.

---

⬅️ Module 2 — [Différences fondamentales VB6 ↔ VB.NET](../02-differences-fondamentales/README.md)  
➡️ Section 3.1 — [Inventaire](01-inventaire.md)

---

**Juin 2026** · .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Inventaire (formulaires `.frm`, modules `.bas`, classes `.cls`, références, OCX, DLL, API)](/03-evaluer-strategie/01-inventaire.md)
