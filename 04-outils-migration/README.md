🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4. Outils de migration 🛠️

> **Partie 1 — Comprendre et cadrer la migration**
> Ce module clôt la partie « cadrage ». Le module 3 a permis de savoir **quoi** migrer et
> **selon quelle stratégie** ; ce module répond à la question suivante : **avec quels outils** ?
> Assistants intégrés, outils commerciaux, migration manuelle, IA — chacun avec ses forces et,
> surtout, ses **limites honnêtes**.

---

## 🧭 Pourquoi ce module — et pourquoi maintenant

L'ordre n'est pas anodin : on choisit ses outils **après** avoir évalué l'existant et arrêté une  
stratégie (module 3), jamais avant. Choisir un outil avant de connaître son projet, c'est laisser  
l'outil dicter la démarche — alors que c'est la démarche qui doit dicter l'outil. L'inventaire
(3.1), la cartographie (3.2), l'estimation (3.3) et le choix big-bang/incrémental (3.4) sont
précisément ce qui permet de décider, en connaissance de cause, quel outillage convient.

Ce module repose sur un **repère honnête**, posé dès le sommaire de la formation et qu'il faut  
graver avant toute chose :

> ⚠️ **Aucun outil ne migre *tout* automatiquement.** Les assistants traitent **60 à 90 %** du
> travail **mécanique** ; le **jugement humain** — et l'IA, bien encadrée — traite le reste. Et ce
> reste n'est pas du détail : c'est précisément là que se logent les **pièges silencieux**
> (le `ByRef` devenu `ByVal`, les entiers redimensionnés, la finalisation perdue — Annexe B) et
> l'**UI** à repenser. Le pourcentage rassurant porte sur le **volume de code**, pas sur l'**effort**
> réel (rappel chiffré du module 3.3 : la fraction résiduelle peut concentrer la majorité du
> travail).

Tout ce module se lit à travers ce prisme : aucun outil n'est une baguette magique, mais aucun  
n'est inutile non plus. La bonne approche est presque toujours **outillée *et* manuelle**, en  
sachant exactement ce que l'on attend de chaque outil — et ce qu'il ne fera pas.

---

## 🎯 Ce que vous saurez faire à l'issue de ce module

- **Comprendre ce que fait — et ne fait pas — l'assistant de mise à niveau** de Visual Studio
  (*Upgrade Wizard*), ses capacités réelles et ses limites structurelles.
- **Reconnaître la béquille `Microsoft.VisualBasic.Compatibility`** : pourquoi les outils s'en
  servent, pourquoi c'est commode à court terme, et pourquoi c'est une **dette à éviter** (ou à
  rembourser, module 17.2).
- **Situer les outils commerciaux** — **VBUC** (Mobilize.Net) et **VB Migration Partner** (Code
  Architects) — et ce qu'ils apportent : conversion vers VB.NET (ou C#), suppression des
  dépendances de compatibilité, gestion des *control arrays*.
- **Savoir quand la migration manuelle s'impose** et pourquoi elle reste irremplaçable sur
  certaines zones.
- **Décider entre automatique, manuel et hybride** à l'aide d'une grille de décision adossée au
  cadrage du module 3.
- **Comprendre la place de l'IA** dans cet outillage, en renvoi au module 19 qui lui est dédié.

---

## 📚 Le parcours du module

| § | Sujet | Ce que vous y trouverez |
|---|-------|-------------------------|
| **4.1** ⚠️ | L'assistant de mise à niveau Visual Studio (*Upgrade Wizard*) | Ses capacités et, surtout, ses **limites** |
| **4.2** ⚠️ | L'espace de noms `Microsoft.VisualBasic.Compatibility` | La **béquille à éviter** — et pourquoi |
| **4.3** | Outils commerciaux (VBUC, VB Migration Partner) | Conversion VB.NET/C#, dépendances de compatibilité, *control arrays* |
| **4.4** | Migration manuelle assistée | Quand et pourquoi elle reste nécessaire |
| **4.5** | Grille de décision | Automatique vs manuel vs hybride |
| **4.6** 🤖 | Migration assistée par IA | Renvoi au **module 19** (dédié) |

La progression va du **plus automatique** (4.1) au **plus manuel** (4.4), avant de fournir une
**grille de décision** (4.5) pour arbitrer, puis d'ouvrir sur l'**IA** (4.6, traitée en profondeur
au module 19).

---

## 🧰 Le paysage des outils en un coup d'œil

Avant d'entrer dans le détail, une vue d'ensemble des grandes familles. Aucune n'est « la  
meilleure » : chacune répond à un besoin différent, et elles se **combinent** plus qu'elles ne  
s'opposent.

| Famille | Exemples | Force principale | Limite principale |
|---------|----------|------------------|-------------------|
| **Assistant intégré** | *Upgrade Wizard* de Visual Studio | Gratuit, disponible immédiatement | S'appuie sur la **béquille** de compatibilité (4.2), limites importantes (4.1) ⚠️ |
| **Outils commerciaux** | VBUC, VB Migration Partner | Conversion plus complète, suppression des dépendances de compatibilité, *control arrays* | Coût, courbe d'apprentissage, validation toujours nécessaire |
| **Migration manuelle** | (le développeur, outillé) | Maîtrise totale, qualité, idiomatique | Lente, exige l'expertise VB6 **et** .NET |
| **IA** 🤖 | Assistants de génération de code | Conversion, **explication**, génération de tests | Hallucinations, **faux équivalents** VB6/VB.NET, validation systématique (module 19) ⚠️ |

> 💡 **Combiner plutôt qu'opposer.** La question n'est presque jamais « quel outil ? » au
> singulier, mais « **quel dosage** ? ». Un projet réel mêle typiquement un outil de conversion
> pour le travail mécanique, du manuel pour les points durs (3.3) et l'UI, et de l'IA bien
> encadrée pour expliquer le code obscur et générer des tests. La grille de décision (4.5) aide à
> régler ce dosage.

---

## ⚠️ Deux principes à garder en tête

> **1. Le travail des outils porte sur le volume, pas sur l'effort.** « L'outil a converti 85 % du
> code » ne veut **pas** dire « 85 % du travail est fait » — c'est l'illusion dénoncée au module
> 1.5 et chiffrée en 3.3. La strate « rouge » (points durs, pièges silencieux, UI) résiste à
> l'automatisation et concentre l'essentiel de l'effort. Mesurer l'avancement au seul pourcentage
> converti est un piège.

> **2. La sortie d'un outil n'est jamais le résultat final.** Quel que soit l'outil — assistant,
> outil commercial ou IA —, son résultat est un **point de départ**, pas une livraison. Il doit
> être **relu**, **validé** et **fiabilisé** : c'est tout l'objet de la **partie 6** (module 17),
> qui active `Option Strict` progressivement, supprime les béquilles de compatibilité (17.2) et
> traque les pièges silencieux (17.4) en comparant au **golden master** (5.5, 17.3). Aucun outil
> ne dispense de cette validation — surtout pas l'IA (module 19.5).

---

## 🔗 Prérequis et liens utiles

- **À lire avant** : tout le **module 3**, dont ce module est le prolongement direct — la grille de
  décision (4.5) s'appuie sur l'inventaire, l'estimation et la stratégie qui y ont été établis.
- **Sur les pièges que les outils ne corrigent pas** : module 2 (différences fondamentales) et
  **Annexe B** (catalogue des pièges silencieux) — la raison d'être de la validation manuelle.
- **Sur la béquille de compatibilité** : la section 4.2 introduit le problème ; le module **17.2**
  explique comment **s'en débarrasser** après coup.
- **Sur la validation après migration** : module **17** dans son ensemble — la suite logique et
  indispensable de toute migration outillée.
- **Sur l'IA** : module **19** (dédié), vers lequel pointe la section 4.6.

---

⬅️ Module 3 — [Évaluer l'existant et choisir une stratégie](../03-evaluer-strategie/README.md)  
➡️ Section 4.1 ⚠️ — [L'assistant de mise à niveau Visual Studio (*Upgrade Wizard*)](01-upgrade-wizard.md)

---

**Juin 2026** · .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [L'assistant de mise à niveau Visual Studio (*Upgrade Wizard*) : capacités et **limites**](/04-outils-migration/01-upgrade-wizard.md)
