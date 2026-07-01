# 🔄 Formation Migration VB6 → VB.NET (.NET Framework 4.7.2)

![License](https://img.shields.io/badge/License-CC%20BY--NC--SA%204.0-blue.svg)  
![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.7.2-512BD4.svg)  
![Visual Studio](https://img.shields.io/badge/Visual%20Studio-2026-blueviolet.svg)  
![Migration](https://img.shields.io/badge/VB6%20%E2%86%92%20VB.NET-Migration-orange.svg)  
![Modules](https://img.shields.io/badge/Modules-20-green.svg)  
![Annexes](https://img.shields.io/badge/Annexes-8-yellowgreen.svg)  
![Language](https://img.shields.io/badge/Langue-Français-blue.svg)  
![Mise à jour](https://img.shields.io/badge/Mise%20à%20jour-Juin%202026-brightgreen.svg)

**Un guide progressif et honnête pour migrer une application Visual Basic 6 vers VB.NET — en gérant le vrai risque : les changements *silencieux* de comportement.**

<p align="center">
  <img src="https://upload.wikimedia.org/wikipedia/commons/4/40/VB.NET_Logo.svg" alt="VB.NET Logo" width="200"/>
</p>

---

## 📖 Table des matières

- [À propos](#-à-propos)
- [Points forts](#-points-forts)
- [Pour qui ?](#-pour-qui-)
- [Contenu](#-contenu-de-la-formation)
- [Démarrage](#-démarrage-rapide)
- [Structure](#-structure-du-projet)
- [Parcours](#-parcours-dapprentissage-suggéré)
- [Ressources](#-ressources-officielles)
- [FAQ](#-faq)
- [Licence](#-licence)
- [Contact](#-contact)

---

## 📋 À propos

Cette formation accompagne la **migration d'un code Visual Basic 6 vers VB.NET**, en ciblant
**.NET Framework 4.7.2**. Son fil rouge : migrer un existant qui doit **continuer à fonctionner à
l'identique**, pas le réécrire de zéro.

> ⚠️ **Migrer du VB6 vers VB.NET n'est pas une conversion de syntaxe.** C'est un changement
> de **trois choses à la fois** : le **langage**, le **runtime** (moteur VB6 → CLR avec
> *garbage collector*) et le **modèle objet** (instanciation, durée de vie, finalisation).
> Le vrai danger n'est pas le code qui refuse de compiler — c'est le code qui **compile mais ne
> se comporte plus pareil**.

**Pourquoi cibler .NET Framework 4.7.2 (et pas .NET 8/10 directement) ?** Parce que c'est le  
**« pont »** naturel, celui qui **minimise le nombre de changements simultanés** :

- ✅ Surface d'API la plus **proche de VB6** (Windows, COM, GDI+, impression) ;
- ✅ Meilleure **compatibilité COM / ActiveX / OCX** et **Crystal Reports** ;
- ✅ Parité du **concepteur Windows Forms** et des contrôles intrinsèques ;
- ✅ **Runtime supporté** (lié au cycle de vie de Windows), adapté au *legacy* d'entreprise.

Le saut **4.7.2 → .NET moderne** est traité comme une **migration distincte et optionnelle**
(module 20), à faire *ensuite* si nécessaire. Faire les deux en même temps multiplie les risques.

> Cette formation ne cherche pas à vous retenir sur VB6 ni à vous précipiter ailleurs : elle vous
> donne une méthode **outillée et lucide** pour franchir le pas proprement. Elle n'est pas parfaite ;
> l'objectif est de vous faire gagner du temps — et d'éviter les régressions.

**✨ Ce que vous y trouverez :**

- 📚 **20 modules** — du cadrage à la bascule en production, en passant par le langage, l'UI, les données et le COM
- 🎯 **.NET Framework 4.7.2** — la cible « pont » assumée, avec sa justification
- ⚠️ **Un catalogue des pièges** — ByRef/ByVal, entiers redimensionnés, finalisation déterministe perdue, tableaux 0-based…
- 🔗 **L'interop COM comme stratégie** — faire cohabiter VB6 et .NET pour migrer **incrémentalement** (RCW / CCW)
- 🖥️ **Migration de l'UI** — formulaires VB6 → Windows Forms, *control arrays*, graphismes GDI+, impression, ActiveX
- 🗄️ **Migration des données** — ADO/DAO/RDO → ADO.NET, *data binding*, Crystal Reports
- 🛠️ **Outils** — assistant Visual Studio, VBUC (Mobilize.Net), VB Migration Partner (Code Architects), manuel
- 🤖 **Migration assistée par IA** — un module dédié (et ses pièges)
- 📖 **8 annexes** — correspondances VB6↔VB.NET, contrôles, types, checklist, *golden master*
- 🇫🇷 **En français** — parce que c'est plus accessible

**Durée estimée :** ~30-40 h de lecture • 22-28 jours pour le parcours complet en pratiquant  
**Niveau :** vous connaissez le VB6 (le code à migrer) ; des bases en .NET aident mais ne sont pas indispensables

---

## ✨ Points forts

- ✅ **Réaliste** — on migre un existant qui doit continuer à tourner, pas une réécriture idéale
- ⚠️ **Centrée sur les risques** — les changements *silencieux* sont la priorité n°1, pas un détail
- 🔗 **Incrémentale** — l'interop COM permet de migrer module par module sans tout geler
- 🛠️ **Outillée et honnête** — assistants, outils commerciaux et IA, avec leurs **limites** clairement dites
- 🧪 **Sécurisée** — *golden master* et tests de non-régression au cœur de la méthode
- 🧭 **Bien cadrée** — 4.7.2 comme « pont », le saut vers .NET moderne traité comme une étape **séparée**
- 🤖 **IA-aware** — un module complet, avec le risque des « faux équivalents » VB6/VB.NET

---

## 🎯 Pour qui ?

Cette formation s'adresse à différents profils. Choisissez votre parcours :

| 👤 Profil | 📚 Modules recommandés | ⏱️ Durée estimée |
|-----------|------------------------|------------------|
| **Décideur / Cadrage** (évaluer, chiffrer, décider) | 1-4, 20 | 2-3 jours |
| **Migration du langage** (le cœur) ⭐ | 1-2, 5-12 + Annexes A/B/D | 8-10 jours |
| **Migration de l'interface** (Forms) | 1-2, 13-14 + Annexe C | 5-6 jours |
| **Migration données & COM** | 1-2, 15-16 + Annexe A | 5-6 jours |
| **Chef de projet migration** | 1-4, 17-18, 20 | 4-5 jours |
| **Migration assistée par IA** 🤖 | 1-2, 4, 19 | 3-4 jours |
| **Formation complète** | 1-20 + annexes | 22-28 jours |

---

## 📚 Contenu de la formation

### Les 7 parties

| # | Partie | Modules | Niveau | Sujets clés |
|---|--------|---------|--------|-------------|
| **1** | Comprendre et cadrer la migration | 1-4 | 🌱 Débutant | Pourquoi/comment migrer, différences VB6↔VB.NET, évaluation, outils |
| **2** | Préparer le terrain | 5-6 | 🌱 Débutant | Préparer le code VB6, *golden master*, environnement .NET |
| **3** | Migrer le langage ⭐ | 7-12 | 🌿 Intermédiaire | Types, tableaux, contrôle, procédures, erreurs, POO |
| **4** | Migrer l'interface | 13-14 | 🌿 Intermédiaire | Forms → WinForms, *control arrays*, GDI+, impression, ActiveX |
| **5** | Données et interopérabilité | 15-16 | 🌳 Avancé | ADO→ADO.NET, COM (RCW/CCW), `Declare`→P/Invoke |
| **6** | Finaliser, valider, moderniser | 17-18 | 🌳 Avancé | `Option Strict`, non-régression, déploiement, bascule |
| **7** | IA, suite de parcours et ressources | 19-20 | 🌳 Avancé | Migration assistée par IA, cap vers .NET moderne |

### Les modules

1. Introduction : pourquoi et comment migrer
2. **Différences fondamentales VB6 ↔ VB.NET** ⚠️
3. Évaluer l'existant et choisir une stratégie
4. Outils de migration (assistant VS, VBUC, VB Migration Partner) 🛠️
5. Préparer le code VB6 avant la migration
6. Préparer l'environnement .NET (ciblage 4.7.2)
7. Types, variables et déclarations ⚠️
8. Tableaux et collections
9. Opérateurs, chaînes et structures de contrôle
10. **Procédures, fonctions et paramètres (ByRef → ByVal !)** ⚠️
11. Gestion des erreurs (`On Error` → `Try/Catch`)
12. **Programmation orientée objet (la finalisation déterministe)** ⚠️
13. **Des formulaires VB6 à Windows Forms** ⭐
14. Graphismes, impression et contrôles ActiveX
15. Accès aux données (ADO/DAO/RDO → ADO.NET)
16. **Interopérabilité COM et appels d'API** 🔗
17. Valider, fiabiliser et refactoriser après migration
18. Déploiement et bascule en production
19. **Migrer avec l'aide de l'IA** 🤖
20. Et après ? De 4.7.2 vers .NET moderne (optionnel)

### ⚠️ Les pièges majeurs traités (le « top des bugs de migration »)

- 🔁 **ByRef par défaut → ByVal par défaut** — des effets de bord qui disparaissent silencieusement
- 🔢 **Entiers redimensionnés** — `Integer` (16 bits) → `Short`, `Long` (32 bits) → `Integer`, `Currency` → `Decimal`
- ♻️ **Finalisation déterministe perdue** — `Class_Terminate` ne se déclenche plus au bon moment → `IDisposable`/`Using`
- 📐 **Tableaux toujours 0-based** — `Option Base 1` et bornes non nulles n'existent plus
- 🎭 **`Variant` et ses états** — `Empty` / `Null` / `Missing` disparaissent → `Object` + tests explicites
- 🧯 **`On Error Resume Next`** — masque des bugs ; sa migration mécanique est dangereuse
- 🅿️ **Propriétés par défaut et mot-clé `Set`** — supprimés ; accès explicites obligatoires
- 🎛️ ***Control arrays*** — n'existent pas en Windows Forms → solutions de remplacement

### 📎 Les 8 annexes

- **A.** Correspondance syntaxique VB6 → VB.NET (aide-mémoire) ⭐
- **B.** **Catalogue des pièges silencieux** : symptôme, cause, correction ⚠️ ⭐
- **C.** Correspondance des contrôles VB6 → Windows Forms
- **D.** Correspondance des types de données (Variant, entiers, Currency, Date…)
- **E.** Checklist de migration (avant / pendant / après)
- **F.** Modèles de tests de non-régression (*golden master*)
- **G.** Glossaire et acronymes (VB6, VB.NET, COM, CLR…)
- **H.** Ressources et outils

📖 **Sommaire détaillé** → [SOMMAIRE.md](/SOMMAIRE.md)

---

## 🚀 Démarrage rapide

### 1. Installer la cible .NET Framework 4.7.2 et Visual Studio 2026

```powershell
# Installer le Developer Pack .NET Framework 4.7.2 (pour compiler vers cette cible)
# https://dotnet.microsoft.com/download/dotnet-framework/net472

# Installer Visual Studio 2026 Community (gratuit)
# https://visualstudio.microsoft.com/downloads/
# ✅ Charge de travail : ".NET desktop development"
#    (inclut le concepteur Windows Forms et le ciblage .NET Framework)
```

> 💡 Le **code source VB6** reste votre point de départ. L'**EDI VB6 n'étant plus maintenu**, il
> n'est pas requis pour cette formation : on lit/analyse le code VB6 et on le **transpose** vers
> un projet VB.NET ciblant 4.7.2.

### 2. Créer le projet cible VB.NET (.NET Framework 4.7.2)

Dans Visual Studio 2026 : **Nouveau projet → « Application Windows Forms (.NET Framework) » →  
langage Visual Basic → Framework cible : `.NET Framework 4.7.2`**.

Pour les bibliothèques métier réutilisables, créez une **« Bibliothèque de classes (.NET  
Framework) »** en VB, ciblant elle aussi 4.7.2.

### 3. La migration en 4 étapes (le fil rouge de la formation)

```
1) 🔍 INVENTORIER   → formulaires, modules, classes, références, OCX, DLL, API   (modules 3-4)
2) 🧹 PRÉPARER      → nettoyer le VB6 + figer un « golden master » de référence    (modules 5-6)
3) 🔁 MIGRER        → langage, UI, données, COM (manuel / outil / IA)              (modules 7-16, 19)
4) ✅ VALIDER       → non-régression, Option Strict, refactoring, bascule          (modules 17-18)
```

### 4. Cloner cette formation

```bash
git clone https://github.com/NDXDeveloper/formation-migration-vb6-vbnet.git
cd formation-migration-vb6-vbnet
```

---

## 📁 Structure du projet

```
formation-migration-vb6-vbnet/
│
├── 📄 README.md                          # Ce fichier
├── 📄 SOMMAIRE.md                        # Table des matières complète (source de vérité)
├── 📄 LICENSE                            # Licence CC BY-NC-SA 4.0
│
├── 📂 01-cadrage-migration/              # Pourquoi/comment migrer, pourquoi 4.7.2
├── 📂 02-differences-fondamentales/      # ⚠️ VBVM vs CLR, finalisation, ByRef/ByVal
├── 📂 03-evaluer-strategie/              # Inventaire, dépendances, big-bang vs incrémental
├── 📂 04-outils-migration/               # 🛠️ Assistant VS, VBUC, VB Migration Partner
├── 📂 05-preparer-code-vb6/              # Nettoyage, golden master
├── 📂 06-preparer-environnement/         # Ciblage 4.7.2, options de compilation
├── 📂 07-types-variables/                # ⚠️ Variant, entiers redimensionnés, dates
├── 📂 08-tableaux-collections/           # Tableaux 0-based, Collection → .NET
├── 📂 09-operateurs-controle/            # Opérateurs, GoSub supprimé, boucles
├── 📂 10-procedures-fonctions/           # ⚠️ ByRef/ByVal, Optional, Property Get/Let/Set
├── 📂 11-gestion-erreurs/                # On Error → Try/Catch
├── 📂 12-poo/                            # ⚠️ Class_Terminate, IDisposable, Type→Structure
├── 📂 13-formulaires-winforms/           # ⭐ Forms → WinForms, control arrays
├── 📂 14-graphismes-impression-activex/  # GDI+, Printer, ActiveX/OCX
├── 📂 15-acces-donnees/                  # ADO/DAO/RDO → ADO.NET, Crystal Reports
├── 📂 16-interop-com-api/                # 🔗 RCW/CCW, Declare → P/Invoke
├── 📂 17-valider-refactoriser/           # Option Strict, non-régression, refactoring
├── 📂 18-deploiement-bascule/            # Packaging, parallel run, rollback
├── 📂 19-migration-ia/                   # 🤖 Migrer avec l'IA (et ses pièges)
├── 📂 20-apres-net-moderne/              # Cap (optionnel) vers .NET 8/10
│
└── 📂 annexes/
    ├── correspondance-vb6-vbnet/         # A. Aide-mémoire VB6 ↔ VB.NET
    ├── pieges-silencieux/                # B. ⚠️ Le catalogue des bugs de migration
    ├── correspondance-controles/         # C. Contrôles VB6 → Windows Forms
    ├── correspondance-types/             # D. Types de données
    ├── checklist-migration/              # E. Checklist avant/pendant/après
    ├── golden-master/                    # F. Tests de non-régression
    ├── glossaire/                        # G. Glossaire et acronymes
    └── ressources/                       # H. Ressources et outils
```

> Chaque dossier contient un `README.md` (sommaire du module) et un fichier `.md` par section,
> reliés par une navigation 🔝 Sommaire / ⏭️ section suivante.

---

## 🗓️ Parcours d'apprentissage suggéré

```
🌱 DÉBUTANT
│
├─ Partie 1 : Comprendre et cadrer la migration
└─ Partie 2 : Préparer le terrain (code VB6, golden master, environnement)
   │
   ▼
🌿 INTERMÉDIAIRE
│
├─ Partie 3 : Migrer le langage (types, contrôle, POO…)  ← le cœur
└─ Partie 4 : Migrer l'interface (Forms → WinForms)
   │
   ▼
🌳 AVANCÉ
│
├─ Partie 5 : Données et interopérabilité (ADO.NET, COM, P/Invoke)
├─ Partie 6 : Finaliser, valider et basculer en production
└─ Partie 7 : IA et cap (optionnel) vers .NET moderne

🎓 Total : ~30-40 h de lecture (22-28 jours en pratiquant, à 30 min-1 h/jour)
```

**🎯 Parcours Express (3-4 jours)** pour démarrer vite :
- Module 1 — Pourquoi et comment migrer (et pourquoi 4.7.2)
- Module 2 — Les différences fondamentales et les pièges
- Module 4 — Choisir son outil (manuel / VBUC / VB Migration Partner)
- Module 13 — Migrer les formulaires vers Windows Forms
- Annexe B — Le catalogue des pièges silencieux

---

## 🔗 Ressources officielles

### Documentation et support

| Ressource | Lien |
|-----------|------|
| 📖 Documentation .NET Framework | [learn.microsoft.com/dotnet/framework](https://learn.microsoft.com/dotnet/framework/) |
| 📖 Référence du langage VB.NET | [learn.microsoft.com/dotnet/visual-basic](https://learn.microsoft.com/dotnet/visual-basic/) |
| 🛟 **Statut de support de VB6** (officiel) | [Support policy VB6](https://learn.microsoft.com/previous-versions/visualstudio/visual-basic-6/visual-basic-6-support-policy) |
| 🗓️ Annonce cycle de vie VB6 | [Lifecycle announcement](https://learn.microsoft.com/lifecycle/announcements/visual-basic-6-support-announcement) |
| 📥 Télécharger .NET Framework 4.7.2 (Developer Pack) | [dotnet.microsoft.com/download/dotnet-framework/net472](https://dotnet.microsoft.com/download/dotnet-framework/net472) |
| ⚙️ Visual Studio 2026 | [visualstudio.microsoft.com](https://visualstudio.microsoft.com) |
| 💬 Forum Q&A | [learn.microsoft.com/answers](https://learn.microsoft.com/answers) |

### Outils de migration

| Outil | Éditeur | Lien |
|-------|---------|------|
| 🛠️ **VBUC** (Visual Basic Upgrade Companion) | Mobilize.Net (GAPVelocity AI) | [gapvelocity.ai](https://www.gapvelocity.ai) · [page partenaire MS](https://learn.microsoft.com/previous-versions/visualstudio/visual-basic-6/vb6-partners-mobilize-net) |
| 🛠️ **VB Migration Partner** | Code Architects | [vbmigration.com](https://www.vbmigration.com) |
| 🤖 GitHub Copilot (assistance IA) | GitHub | [github.com/features/copilot](https://github.com/features/copilot) |

> ℹ️ VBUC et VB Migration Partner peuvent cibler VB.NET **ou** C#, et différentes versions de .NET.
> Pour cette formation, on les configure pour produire du **VB.NET sur .NET Framework 4.7.2**.
> Mobilize.Net propose (via Microsoft) une **licence gratuite jusqu'à ~10 000 lignes** pour évaluer VBUC.

---

## ❓ FAQ

**Q : Pourquoi cibler .NET Framework 4.7.2 et pas directement .NET 8/10 ?**
> Parce que 4.7.2 est le **« pont »** : la cible la plus proche de VB6 (API Windows, COM, GDI+,
> impression), donc celle qui **minimise les changements simultanés**. Le saut vers .NET moderne est
> une **seconde migration distincte** (module 20), à faire ensuite si nécessaire.

**Q : L'assistant de migration de Visual Studio suffit-il ?**
> Rarement, pour une vraie application. Il produit du code qui s'appuie sur l'espace de noms de
> compatibilité et laisse beaucoup de travail manuel. Le module 4 le compare honnêtement aux outils
> commerciaux (VBUC, VB Migration Partner) et à la migration manuelle assistée par IA.

**Q : Combien de temps prend une migration VB6 → VB.NET ?**
> Cela dépend surtout de la **taille**, de la **complexité** et des **dépendances** (ActiveX, COM,
> Crystal Reports). Le module 3 propose une méthode d'inventaire et d'estimation. Les outils
> accélèrent la partie mécanique ; le **jugement humain** (et les tests) traite le reste.

**Q : Peut-on migrer progressivement, module par module ?**
> Oui — c'est même recommandé. Le module 16 montre comment faire **cohabiter** VB6 et .NET via
> l'interopérabilité COM (RCW pour consommer du COM existant, CCW pour exposer du .NET à VB6).

**Q : Mon application utilise des contrôles ActiveX/OCX. Que faire ?**
> Deux voies (module 14) : **réutiliser** l'ActiveX via interop, ou le **remplacer** par un contrôle
> .NET équivalent. Le choix dépend de la disponibilité, du support et de la dette que vous acceptez.

**Q : Les tableaux de contrôles (*control arrays*) existent-ils en VB.NET ?**
> Non, pas en Windows Forms. Le module 13 détaille les solutions : recréer un tableau à la main,
> partager un gestionnaire d'événements, ou s'appuyer sur la bibliothèque de support d'un outil
> de migration.

**Q : Faut-il tout réécrire ou migrer ?**
> Cette formation **migre** (on préserve les règles métier de l'existant). La réécriture n'est
> pertinente que dans des cas précis (module 3) — elle coûte plus cher et fait perdre le comportement
> éprouvé.

**Q : Crystal Reports, DAO et ADO sont-ils gérés ?**
> Oui : les données (ADO/DAO/RDO → ADO.NET) au module 15, le *data binding* et le **reporting**
> (Crystal Reports et alternatives) inclus. 4.7.2 maximise la compatibilité de ces briques.

**Q : VB6 est-il encore « supporté » ?**
> L'**EDI VB6 n'est plus supporté depuis avril 2008**. En revanche, le **runtime VB6** reste pris en
> charge pour la durée de vie des versions de Windows où il est livré (« It Just Works »). C'est une
> raison de migrer sereinement, pas dans l'urgence — mais une raison réelle (module 1).

**Q : La formation couvre-t-elle la migration avec l'IA ?**
> Oui — **module 19** : prompting pour convertir/expliquer du VB6, détecter les pièges, générer des
> tests… et surtout ses **limites** (faux équivalents VB6/VB.NET, hallucinations de syntaxe), avec
> validation systématique.

**Q : Et après 4.7.2 ?**
> Le module 20 fait la passerelle vers la formation **« VB.NET avec .NET 10 LTS »** : quand sauter,
> ce qui devra encore changer, et quand **rester** sur 4.7.2 (un choix parfois légitime).

---

## 📝 Licence

Ce projet est sous licence **Creative Commons Attribution - Pas d'Utilisation Commerciale - Partage dans les Mêmes Conditions 4.0 International (CC BY-NC-SA 4.0)**.

✅ **Vous pouvez :**
- Partager — copier et redistribuer le matériel
- Adapter — remixer, transformer et créer à partir du matériel

⚠️ **Selon les conditions suivantes :**
- **Attribution** — vous devez créditer l'œuvre originale
- **Pas d'Utilisation Commerciale** — pas d'usage à des fins commerciales
- **Partage dans les Mêmes Conditions** — toute redistribution sous la même licence

📄 Voir le fichier [LICENSE](/LICENSE) pour les détails complets.

**Attribution suggérée :**
```
Formation Migration VB6 → VB.NET (.NET Framework 4.7.2) par Nicolas DEOUX
https://github.com/NDXDeveloper/formation-migration-vb6-vbnet
Licence CC BY-NC-SA 4.0
```

---

## 👨‍💻 Contact

**Nicolas DEOUX**
- 📧 [NDXDev@gmail.com](mailto:NDXDev@gmail.com)
- 💼 [LinkedIn](https://www.linkedin.com/in/nicolas-deoux-ab295980/)
- 🐙 [GitHub](https://github.com/NDXDeveloper)

---

## 🙏 Remerciements

Merci à :
- **Microsoft** et la **.NET Foundation** pour .NET, Visual Studio et le maintien du runtime VB6
- Les éditeurs d'outils de migration (**Mobilize.Net**, **Code Architects**) qui rendent ces projets atteignables
- La **communauté VB6 / VB.NET** qui documente et partage son expérience de migration
- **Anthropic**, **OpenAI** et les créateurs d'outils IA qui accélèrent la modernisation du *legacy*
- **Vous** qui prenez le temps de migrer proprement plutôt que dans la précipitation

---

<div align="center">

## 🔄 Bonne migration de VB6 vers VB.NET ! 🔄

*Cette formation est un travail en cours. Elle n'est pas parfaite, mais j'espère sincèrement qu'elle vous évitera bien des régressions sur votre chemin de migration.*

**[📖 Consulter le sommaire complet →](/SOMMAIRE.md)**

[![Star on GitHub](https://img.shields.io/github/stars/NDXDeveloper/formation-migration-vb6-vbnet?style=social)](https://github.com/NDXDeveloper/formation-migration-vb6-vbnet)
[![Follow](https://img.shields.io/github/followers/NDXDeveloper?style=social)](https://github.com/NDXDeveloper)

**[⬆ Retour en haut](#-formation-migration-vb6--vbnet-net-framework-472)**

*Dernière mise à jour : Juin 2026*

</div>
