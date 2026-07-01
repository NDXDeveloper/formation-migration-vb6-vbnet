🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.5 Grille de décision : outil automatique vs manuel vs hybride

> **Module 4 — Outils de migration**
> Voici l'aboutissement du module. Les sections précédentes ont présenté chaque voie :
> l'assistant disparu (4.1), les outils commerciaux (4.3), le manuel assisté (4.4). Reste à
> **choisir** — ou plutôt, comme l'a annoncé la section 4.4, à **doser**. Cette grille aide à
> placer le curseur, en s'appuyant directement sur le cadrage du module 3.

---

## 🧭 Reformuler la question : un dosage, pas un camp

La section 4.4 l'a posé comme principe : opposer « **outil** » et « **manuel** » est un **faux  
dilemme**. Toute migration massivement outillée **se termine** à la main (le résidu, les points durs,  
l'UI) ; toute migration dite « manuelle » s'appuie sur des **outils** (IDE, IA, parfois un outil  
commercial sur les parties faciles). La vraie décision n'est donc pas binaire.

> 💡 **La bonne question n'est pas « outil *ou* manuel ? » mais « quel *dosage*, et *où* placer la
> frontière ? »** Quelle part automatiser, quelle part traiter à la main, et selon quelle ligne de
> partage. C'est exactement ce que cette grille aide à régler — non pas en tranchant à votre place,
> mais en reliant la décision aux **mesures** déjà faites au module 3 (inventaire, dépendances,
> estimation, stratégie).

Précision de vocabulaire pour 2026 : « **automatique** » désigne ici, en pratique, les **outils  
commerciaux** (section 4.3). L'assistant intégré (4.1) ayant disparu de Visual Studio, il n'est plus  
une option ; quand cette grille dit « outil automatique », elle vise VBUC, VB Migration Partner ou
équivalents.

---

## 🎛️ Les trois voies, en synthèse

Avant de croiser les facteurs, une vue d'ensemble des trois approches et de leur terrain de  
prédilection.

| Voie | Principe | Forces | Limites | Terrain de prédilection |
|------|----------|--------|---------|--------------------------|
| **Automatique** (outils commerciaux, 4.3) | L'outil convertit en masse ; l'humain valide | Rapide sur le **volume** ; couvre le gros œuvre (strates verte/orange) ; métriques et détection de pièges | **Coût** ; courbe d'apprentissage ; sortie **à valider** ; UI à retoucher | Gros volumes ; bases régulières ; cible **C#** (VBUC) |
| **Manuel assisté** (4.4) | L'humain pilote, outillé par IDE/IA | **Jugement** ; code idiomatique ; traite les points durs ; décisions métier | **Lent** sur le volume ; exige une **double expertise** (VB6 **et** .NET) | Petites applications ; **points durs** ; budget contraint avec expertise interne |
| **Hybride** | Outil pour le facile, manuel pour le dur | **Cumule** les forces des deux | Suppose de savoir **où placer la frontière** | **La plupart des projets réels**, surtout volumineux **et** truffés de points durs |

> 💡 L'hybride n'est pas un compromis mou : c'est, le plus souvent, la voie **la plus rationnelle**.
> Le réflexe par défaut sur un projet d'envergure est : *outil sur la strate verte/orange, manuel
> sur la strate rouge* (3.3). Les « voies pures » (tout automatique, tout manuel) restent
> pertinentes surtout aux **extrêmes** (très petit projet, ou code si particulier qu'aucun outil ne
> tient).

---

## 📊 Les facteurs de décision (le module 3 nourrit la grille)

C'est ici que tout le travail d'évaluation **paie**. Chaque facteur ci-dessous a été **mesuré** au  
module 3 ; la grille ne fait que les convertir en orientation.

| Facteur | Mesuré en | Penche vers **automatique** | Penche vers **manuel** | Penche vers **hybride** |
|---------|-----------|------------------------------|-------------------------|--------------------------|
| **Taille** (volume) | 3.1, 3.3 | Grande | Petite | Grande |
| **Densité de points durs** | 3.3 | Faible | Concentrée et limitée | **Élevée** (outil + manuel ciblé) |
| **Stratégie** | 3.4 | *Big-bang* sur gros volume | *Big-bang* sur petit volume | **Incrémentale** |
| **Budget outil** | 3.6 | Disponible | Nul (mais expertise présente) | Disponible |
| **Expertise VB6 *et* .NET** | — | Limitée en interne | **Disponible** | Disponible |
| **Exigence de qualité du code** | 3.6 | Standard | **Maximale** (contrôle fin) | Élevée sur les zones sensibles |
| **Langage cible** | 3.6 | **C#** (atout VBUC) | Indifférent | Indifférent |
| **Horizon de modernisation** (module 20) | 3.6, 20 | Saut prévu → **code natif** (VBUC) | — | Saut prévu → privilégier le natif |

> ⚠️ **Aucun facteur ne tranche seul.** C'est la **convergence** des signaux qui oriente, comme pour
> le choix de stratégie (3.4). Une grosse application critique, à forte densité de points durs, dont
> le périmètre ne peut geler, appelle clairement un **hybride incrémental** (outil sur le volume,
> manuel sur le rouge, cohabitation COM le temps de la transition). Une petite appli interne sans
> dépendance problématique appelle tout aussi clairement le **manuel** (ou un outil gratuit dans la
> limite de volume offerte).

---

## 🧭 Une logique de décision, étape par étape

Plutôt qu'un organigramme rigide, voici une **suite de questions** à se poser dans l'ordre. Chacune  
affine le dosage.

1. **Le volume justifie-t-il l'investissement dans un outil ?**
   Petit projet → le **manuel** est souvent plus économique (apprendre et configurer un outil peut
   coûter plus que migrer à la main). Gros projet → un **outil** devient rentable sur le volume.

2. **Quelle est la densité de points durs ?** (3.3)
   Faible → l'**automatique** couvre l'essentiel. Élevée → de toute façon **hybride** : aucun outil
   ne traite seul un OCX abandonné, un *control array*, un `On Error Resume Next`. La question
   devient *où* placer la frontière entre outil et manuel.

3. **Quelle stratégie a été retenue ?** (3.4)
   *Big-bang* → une passe outil (si volume) puis reprise manuelle. **Incrémentale** → le dosage peut
   **varier par morceau** (3.5), et un outil au cycle *convert-test-fix* (VB Migration Partner, 4.3)
   épouse particulièrement bien cette approche.

4. **De quelles ressources dispose-t-on ?**
   Budget outil **et** expertise → tous les choix sont ouverts. Budget nul **mais** expertise interne
   → **manuel assisté** (IDE + IA). Budget présent **mais** expertise rare → l'**outil** compense en
   partie le manque de spécialistes.

5. **Quelle est la cible et l'horizon ?** (3.6, module 20)
   Cible **C#** → VBUC est un atout. **Second saut** vers .NET moderne envisagé → privilégier le
   **code natif** (sans béquille de 4.2, ni dépendance d'exécution lourde) pour ne pas hypothéquer
   cette suite.

---

## 👥 Quelques profils types

Pour rendre la grille concrète, voici des **situations fréquentes** et le dosage qu'elles appellent  
généralement. (Des repères, pas des verdicts : votre contexte mesuré prime toujours.)

| Profil | Contexte | Dosage recommandé |
|--------|----------|-------------------|
| **Petite appli interne** | Faible volume, peu de points durs, non critique | **Manuel** assisté (ou outil gratuit dans la limite de volume offerte) |
| **Grosse appli critique** | Fort volume, points durs nombreux, périmètre non gelable | **Hybride incrémental** : outil sur le volume, manuel sur le rouge, cohabitation COM (3.5) |
| **Migration vers C#** | Cible stratégique C# | **Outil** (VBUC) sur le gros œuvre + **manuel** sur le résidu |
| **Budget nul, expertise présente** | Pas d'achat possible, spécialistes VB6/.NET en interne | **Manuel** assisté (IDE, `Option Strict`, IA — module 19) |
| **Code très particulier** | *Late binding* massif, patterns exotiques | **Manuel** dominant (les outils patinent), outil ponctuel si utile |
| **Modernisation visée** | Saut 4.7.2 → .NET moderne prévu (module 20) | Outil produisant du **code natif** + manuel, en évitant toute béquille (4.2) |

---

## ⚠️ L'erreur à éviter : choisir « en bloc », une fois pour toutes

Le piège classique de cette décision est de la prendre **globalement et définitivement** : « ce  
sera tout automatique » ou « ce sera tout manuel », décrété au départ et figé.

C'est doublement risqué :

- **Le bon dosage varie selon les *zones*.** La strate verte (3.3) se prête à l'outil, la strate
  rouge au manuel — dans **la même application**. Décider par strate, et non en bloc, est presque
  toujours plus juste. C'est l'essence même de l'**hybride**.
- **Le bon dosage s'**ajuste** en cours de route.** Comme l'estimation (3.3), la décision se
  **réactualise** : un **pilote** sur un sous-ensemble représentatif révèle vite si l'outil tient
  ses promesses sur *votre* code, et permet de **recalibrer** la frontière outil/manuel pour la
  suite.

> 💡 Cette logique de **calibration progressive** est exactement celle de la migration **incrémentale**
> (3.4) : les premiers morceaux **renseignent** le dosage des suivants. Un avantage de plus de
> l'approche par étapes.

---

## 🔬 En pratique : comment trancher concrètement

La grille s'applique d'autant mieux qu'on l'**alimente** par des éléments tangibles plutôt que par  
l'intuition :

- **Partir de l'évaluation du module 3** : l'inventaire (3.1), la cartographie (3.2) et surtout
  l'estimation en strates (3.3) **sont** les données d'entrée de cette grille.
- **Exploiter la phase d'évaluation d'un outil** : VBUC ou VB Migration Partner (4.3) produisent des
  **rapports** (métriques, dépendances, éléments problématiques) qui chiffrent la part « facile » et
  la part « dure » — un excellent moyen objectif de situer la frontière.
- **Faire un pilote** : migrer un **sous-ensemble représentatif** (incluant au moins un point dur)
  avec le dosage envisagé, puis **mesurer** (taux de non-régression contre le *golden master* 5.5,
  effort réel, qualité). Le pilote **valide ou corrige** la décision avant de l'étendre.
- **Décider par strate** : appliquer la grille **zone par zone** (vert → outil, rouge → manuel)
  plutôt qu'à l'application entière.

---

## ✅ En résumé

- La décision n'est pas **« outil ou manuel »** mais **« quel dosage, et où placer la frontière »** :
  toute migration outillée finit à la main, toute migration manuelle s'outille (rappel de 4.4).
- En 2026, « **automatique** » signifie **outils commerciaux** (4.3), l'assistant intégré ayant
  disparu (4.1).
- Trois voies : **automatique** (rapide sur le volume, à valider, coûteuse), **manuel assisté**
  (jugement et qualité, lent, double expertise), **hybride** (cumule les forces — souvent la voie la
  plus rationnelle : *outil sur le vert, manuel sur le rouge*).
- Les **facteurs de décision** viennent tous du module 3 : taille (3.1, 3.3), densité de points durs
  (3.3), stratégie (3.4), budget et expertise, exigence de qualité et **langage cible** (3.6),
  **horizon de modernisation** (module 20). **Aucun ne tranche seul** : c'est leur **convergence**
  qui oriente.
- **Profils types** : petite appli → manuel ; grosse appli critique → **hybride incrémental** ;
  cible C# → outil + manuel ; budget nul avec expertise → manuel assisté ; code très particulier →
  manuel dominant ; modernisation visée → **code natif** + manuel.
- **Erreur à éviter** : choisir **en bloc et définitivement**. Le bon dosage **varie par zone**
  (strate) et **s'ajuste** au fil de l'eau, via un **pilote** et la calibration progressive de
  l'incrémental.
- **En pratique** : alimenter la grille par l'évaluation (3.1-3.3) et la phase d'analyse des outils,
  faire un **pilote** mesuré contre le *golden master* (5.5), et décider **strate par strate**.

---

⬅️ Section 4.4 — [Migration manuelle assistée : quand et pourquoi](04-migration-manuelle.md)  
➡️ Section 4.6 🤖 — [Migration assistée par IA → voir module 19](../19-migration-ia/README.md)

---

**Juin 2026** · .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Préparer le code VB6 avant la migration](/05-preparer-code-vb6/README.md)
