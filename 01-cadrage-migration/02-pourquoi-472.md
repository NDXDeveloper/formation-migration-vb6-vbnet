🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 1.2 Pourquoi cibler .NET Framework 4.7.2 : le « pont » naturel ⭐

> **Chapitre 1 — Introduction : pourquoi et comment migrer**
> Deuxième question de cadrage : *vers quoi* migrer, exactement — et pourquoi cette cible-là ?

---

## La cible est, elle aussi, une décision de gestion du risque

On a établi à la section précédente que migrer s'impose. Reste la question qui détermine la moitié  
du succès du projet : **migrer vers quoi ?** Et l'intuition la plus répandue — « tant qu'à migrer,  
autant viser le plus récent, .NET 10 » — est précisément le piège à éviter.

Rappelons le fil rouge de la formation : une migration VB6 → VB.NET, c'est déjà **trois changements  
simultanés** (langage, runtime, modèle objet). Le seul levier dont vous disposez pour garder le  
chantier maîtrisable, c'est de **ne pas en ajouter un quatrième**. Choisir la cible, c'est donc  
choisir **combien de choses changent en même temps**. Et la cible qui en change le **moins**, c'est
**.NET Framework 4.7.2**.

---

## 🌐 D'abord, comprendre qu'il existe deux mondes « .NET »

Le mot « .NET » recouvre aujourd'hui **deux familles distinctes**, et confondre les deux est la  
première source d'erreur de cadrage.

| | **.NET Framework** (4.7.2 / 4.8) | **.NET moderne** (.NET 8, 10…) |
|---|---|---|
| **Origine** | La plateforme « historique », depuis 2002 | Reconstruction issue de .NET Core (depuis 2016) |
| **Système** | **Windows uniquement** | Multiplateforme (Windows, Linux, macOS) |
| **Livraison** | **Composant de Windows** | Installé/embarqué avec l'application |
| **Surface d'API** | Très **proche** de l'univers VB6/Win32 | Modernisée, **épurée**, certaines API retirées |
| **COM / ActiveX** | Support **mature et natif** | Possible (Windows) mais plus limité |
| **Rôle dans la migration** | **Le pont** depuis VB6 | Une **étape ultérieure**, optionnelle |

La distinction qui compte n'est donc **pas** « 4.7.2 ou 4.8 » (deux versions très proches de la même  
famille) : c'est **« .NET Framework classique » contre « .NET moderne »**. Viser d'emblée .NET 10,  
c'est viser une cible **plus éloignée** de VB6 — donc multiplier les écarts à franchir d'un coup.

---

## 🌉 Pourquoi 4.7.2 est le « pont » : quatre raisons concrètes

### 1. La surface d'API la plus proche de VB6

.NET Framework partage l'ADN Windows de VB6 : accès au système, **GDI+** pour le dessin, modèle
d'impression, interactions Win32. Le dessin VB6 (`Line`, `Circle`, `PSet`) se retranscrit  
directement vers l'objet `Graphics`, et `System.Drawing` **fonctionne sans friction** — là où le  
même `System.Drawing` sur .NET moderne a connu des restrictions multiplateformes. Pour une  
application *legacy* riche en graphismes ou en impression, c'est un atout décisif.

### 2. La compatibilité COM / ActiveX / OCX maximale

C'est **le** point qui fait ou défait une migration d'application VB6 d'entreprise. La plupart de  
ces applications dépendent de composants **COM** et de contrôles **ActiveX/OCX** (grilles,  
calendriers, contrôles métier). .NET Framework offre le support d'interopérabilité COM le plus
**mature et le mieux éprouvé** : RCW, *Primary Interop Assemblies*, marshaling — tout l'outillage
existe et fonctionne. C'est ce qui permet la stratégie de **cohabitation** VB6 ↔ .NET pendant la  
transition (chapitres 3 et 16). À cela s'ajoute la **compatibilité de Crystal Reports**, encore  
très présent dans le *legacy* et bien moins bien servi sur .NET moderne.

### 3. La parité du concepteur Windows Forms et des contrôles intrinsèques

Vos formulaires VB6 migrent vers **Windows Forms**, dont le **concepteur visuel** est pleinement  
mature sur .NET Framework. Les contrôles intrinsèques de VB6 ont des équivalents directs
(`CommandButton`→`Button`, `TextBox`, `Label`, `Frame`, `OptionButton`…), et le concepteur se
comporte de façon stable et prévisible. Sur .NET moderne, le concepteur WinForms a connu une  
histoire mouvementée (absent un temps, puis réintroduit) : un terrain moins confortable pour qui  
migre un parc important de formulaires.

### 4. Un runtime supporté, adapté au *legacy* d'entreprise

.NET Framework est un **composant de Windows**, **pris en charge dans le cadre du cycle de vie de
Windows**. C'est exactement le profil recherché pour du *legacy* d'entreprise : une cible stable,  
livrée avec le système, sans dépendance externe à gérer — par opposition au rythme de versions plus  
soutenu du monde .NET moderne (cycles LTS de quelques années).

> 💡 **Et le langage, dans tout ça ?** Bonne nouvelle : **VB.NET est le même langage sur les deux
> cibles**. Migrer vers 4.7.2, c'est donc faire le gros œuvre — changer de **langage** et de
> **runtime** — **une seule fois**. Un éventuel passage ultérieur vers .NET moderne ne rejouera
> *plus* le changement de langage : ce sera une migration de **plateforme**, bien plus circonscrite.
>
> En toute honnêteté : Microsoft **maintient pleinement** VB.NET (corrections, sécurité, accès aux
> nouvelles API du runtime via le CLR) mais a annoncé, dès 2020 et réaffirmé depuis, **ne plus le
> faire évoluer** comme langage — pas de nouvelle syntaxe à venir. Pour une migration dont le but
> est de **préserver un existant**, cette stabilité est un **atout**, pas un handicap : la cible ne
> bougera plus sous vos pieds.

---

## ⚖️ « 4.7.2 ou 4.8 ? » — une question secondaire

Pour être honnête : **4.8** (et 4.8.1) est la **dernière** version de .NET Framework, et tout le  
raisonnement ci-dessus s'applique **à l'identique** à la famille 4.7.x / 4.8. Le choix de **4.7.2**  
dans cette formation est **pragmatique** : c'est une version **omniprésente** sur le parc Windows  
encore supporté, stable et éprouvée, qui constitue une **base commune sûre**. Si votre parc impose  
ou favorise 4.8, l'essentiel de la formation reste valable sans changement — la frontière qui  
compte est *Framework classique* vs *.NET moderne*, pas la version mineure choisie à l'intérieur de  
la famille.

---

## 🔤 Et pourquoi VB.NET plutôt que C# ?

La question revient presque toujours : « tant qu'à migrer, autant passer à C#, non ? » C'est un  
choix **légitime mais orthogonal** : il porte sur le **langage**, pas sur la plateforme (l'axe de  
cette section). Deux repères pour trancher.

- **VB.NET est le chemin de moindre friction.** C'est le **langage jumeau** de VB6 : même famille,
  même paradigme, mots-clés proches. La migration reste un **portage** relativement direct, que les
  assistants et l'IA outillent bien. On change de langage **une seule fois, et le moins possible**.
- **C# est une cible viable, mais ajoute un écart.** Les outils commerciaux (VBUC, VB Migration
  Partner) savent d'ailleurs produire du **C#**. Mais viser C# revient à **changer aussi de
  syntaxe**, en plus du runtime et du modèle objet — un front de plus, à contre-courant du fil
  rouge « minimiser les changements simultanés ».

> 💡 **Le bon ordre, ici aussi.** Migrer d'abord vers **VB.NET** (portage sûr, comportement
> préservé) ; **puis**, si l'équipe le souhaite vraiment, une conversion VB.NET → C# devient une
> opération **distincte**, à froid, sur un code déjà managé et couvert par des tests. Cette
> formation choisit **VB.NET** pour cette raison — et son propos reste, pour l'essentiel, valable
> quelle que soit la décision finale sur le langage.

---

## 🪜 Le principe des « deux sauts séparés »

C'est l'aboutissement logique de tout ce qui précède :

```
   VB6  ──────►  VB.NET / .NET Framework 4.7.2  ──────►  VB.NET / .NET moderne (8/10)
        SAUT 1                                    SAUT 2
   (langage + runtime                       (plateforme seulement,
    + modèle objet)                          optionnel, plus tard)
```

- **Saut 1 — obligatoire :** sortir de VB6 vers VB.NET sur 4.7.2. C'est l'objet de cette formation.
- **Saut 2 — optionnel et ultérieur :** passer de 4.7.2 à .NET moderne, **si et seulement si** le besoin le justifie (multiplateforme, fonctionnalités récentes, fin de vie à anticiper).

| | **Migrer vers 4.7.2 (recommandé)** | **Migrer direct vers .NET 10** |
|---|---|---|
| Changement de **langage** | ✅ une fois | ✅ une fois |
| Changement de **runtime/modèle objet** | ✅ une fois | ✅ une fois |
| Changement de **plateforme/API** | ❌ minimal | ⚠️ majeur, **en plus** |
| Compatibilité **COM/ActiveX/Crystal** | ✅ maximale | ⚠️ dégradée |
| **Risque global** | Maîtrisé | **Cumulé** |

Faire les deux sauts en même temps **multiplie les risques** au lieu de les additionner : on ne sait  
plus, face à une régression, si la cause vient du changement de langage, de runtime ou d'API.  
Séparer les sauts, c'est pouvoir **valider** après chacun. Le **module 20** traite ce second saut et  
fait la passerelle vers la formation dédiée « VB.NET avec .NET 10 LTS ».

---

## 🤔 Quand envisager directement .NET moderne ? (la nuance honnête)

La règle « 4.7.2 d'abord » est la bonne **par défaut**, mais elle n'est pas un dogme. Aller
**directement** vers .NET moderne peut se défendre dans des cas précis :

- application **petite et simple**, sans dépendances COM/ActiveX ni Crystal Reports ;
- équipe **déjà à l'aise** avec .NET moderne, qui ne « subit » pas le changement de plateforme ;
- besoin réel et immédiat de **multiplateforme** ou de fonctionnalités propres à .NET moderne ;
- volonté d'éviter une **double migration** sur un code dont la cible finale est de toute façon .NET moderne.

> ⚠️ Mais attention : ces cas restent l'**exception**. Dès qu'il y a du COM, de l'ActiveX, un parc
> de formulaires conséquent ou une équipe peu familière de .NET, le **pont 4.7.2** reste de loin
> l'option la plus sûre. Dans le doute, on fait le saut 1, on valide, **puis** on décide pour le
> saut 2 (voir module 20.4, « Quand **rester** sur 4.7.2 »).

---

## ✅ À retenir

- Il existe **deux mondes .NET** : le **Framework classique** (4.7.2/4.8, Windows, composant du système) et le **.NET moderne** (multiplateforme, versions LTS). La frontière qui compte pour migrer est **celle-là**, pas « 4.7.2 vs 4.8 ».
- **.NET Framework 4.7.2 est le « pont »** parce qu'il **minimise le nombre de changements simultanés** : API proche de VB6, **COM/ActiveX/OCX** et **Crystal Reports** au mieux, **concepteur Windows Forms** mature, **runtime supporté** via Windows.
- **VB.NET est le même langage** sur les deux cibles : on ne fait le gros œuvre (langage + runtime) **qu'une fois**. Microsoft le **maintient pleinement** mais ne le **fait plus évoluer** — une stabilité bienvenue pour une migration de préservation.
- **VB.NET plutôt que C#** : le langage jumeau de VB6 est le **portage le plus sûr** ; C# est possible (les outils le génèrent) mais ajoute un **changement de syntaxe**. On peut toujours convertir en C# **ensuite**, à froid.
- La bonne méthode, ce sont **deux sauts séparés** : 4.7.2 d'abord (obligatoire), .NET moderne ensuite (optionnel, module 20). Les faire en même temps **cumule les risques** au lieu de permettre de valider après chacun.
- Aller directement vers .NET moderne reste une **exception** réservée aux cas simples et sans dépendances lourdes.

---

> 🧭 **Section suivante** → [1.3 VB6 vs VB.NET : ce qui change vraiment](03-ce-qui-change.md)

⏭️ [VB6 vs VB.NET : ce qui change vraiment (langage, runtime, modèle objet, GC)](/01-cadrage-migration/03-ce-qui-change.md)
