🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.1 Visual Studio et le ciblage .NET Framework 4.7.2

> *Partie 2 — Préparer le terrain · Module 6 — Préparer l'environnement .NET*

---

## Objectif de cette section

C'est la **première pierre** de la zone d'atterrissage : installer le bon outil (Visual Studio), configuré comme il faut, et surtout **viser le bon framework cible**. Une erreur fréquente consiste à créer machinalement un projet « .NET » récent et à découvrir, des semaines plus tard, qu'on s'est imposé un second saut (.NET moderne) **en même temps** que la migration du langage. Toute la stratégie du « pont » se joue dans le bon choix de gabarit de projet et de framework cible — c'est précisément ce que cette section verrouille.

---

## Rappel express : pourquoi 4.7.2 ?

Le « pourquoi » de fond a été traité au **§ 1.2** ; on n'y revient ici que sous l'angle **environnement**. Cibler **.NET Framework 4.7.2** plutôt que .NET moderne, c'est minimiser le nombre de changements simultanés :

- **Parité du concepteur Windows Forms** : l'éditeur visuel de formulaires sur .NET Framework est mature et se comporte au plus près de ce qu'on connaît en VB6.
- **Surface d'API proche de VB6** : Windows, GDI+, impression, accès aux composants Windows.
- **Compatibilité COM / ActiveX / OCX** : la meilleure pour réutiliser l'existant pendant la transition (détaillée en § 6.3).

> 💡 Le saut **4.7.2 → .NET 8/10** est une migration **distincte**, traitée séparément au module 20. On ne le cumule pas avec la migration du langage.

À noter honnêtement : **.NET Framework 4.8 / 4.8.1** est la dernière révision de la branche 4.x, et son support est désormais **lié au cycle de vie de Windows** (tant qu'il est installé sur un Windows supporté, il l'est aussi). Cibler **4.7.2** est ici un choix **délibéré** (le « pont »), pas un défaut de mise à jour : il reste parfaitement utilisable, et son *Developer Pack* est toujours disponible au téléchargement.

---

## Choisir et installer Visual Studio

**Visual Studio 2026** est la version courante (successeur de Visual Studio 2022). Trois éditions existent : **Community**, **Professional** et **Enterprise**.

- Pour un *legacy* VB6 d'**entreprise**, on est en général sous **Professional** ou **Enterprise** (l'édition Community est gratuite mais soumise à des restrictions de licence pour les organisations au-delà d'une certaine taille — à vérifier selon votre contexte).
- Visual Studio 2026 s'installe **en parallèle** des versions antérieures et reste **compatible** avec les solutions et projets existants : aucun obstacle à l'adopter progressivement.

### Le workload essentiel

Dans le **programme d'installation de Visual Studio**, un seul *workload* est indispensable pour ce qui nous occupe :

> **« Développement .NET desktop »**

Il apporte tout le nécessaire : **Visual Basic**, **Windows Forms** (et WPF), le **concepteur visuel** de formulaires — pour .NET Framework **comme** pour .NET moderne.

### Le composant qui manque souvent : le pack de ciblage 4.7.2

Le workload installe les versions récentes du framework, mais **pas forcément** le pack de ciblage d'une version précise comme 4.7.2. Il faut souvent l'ajouter explicitement, sous l'onglet **« Composants individuels »** du programme d'installation :

- **Pack de ciblage .NET Framework 4.7.2**
- **SDK .NET Framework 4.7.2**

> ⚠️ **Symptôme classique :** la version **4.7.2 n'apparaît pas** dans la liste « Target framework » des propriétés du projet. C'est presque toujours le **pack de ciblage absent**. On l'ajoute via le programme d'installation (composants individuels) ou en installant le **Developer Pack 4.7.2**, puis on redémarre Visual Studio.

*(Le choix de version de Windows pour la machine de développement n'est pas détaillé ici : une version récente et supportée suffit. La contrainte qui compte vraiment porte sur les machines **cibles** — voir plus bas.)*

---

## Le piège n°1 du ciblage : .NET Framework **vs** .NET « moderne »

Au moment de créer le projet, Visual Studio propose **deux familles** de gabarits aux noms trompeusement proches. **Tout se joue ici.**

| | Gabarits **« (.NET Framework) »** ✅ | Gabarits **.NET « moderne »** ❌ *(pour la migration)* |
|---|---|---|
| Exemple de nom | « Windows Forms App (**.NET Framework**) » | « Windows Forms App » |
| Cible | .NET Framework 4.x (**dont 4.7.2**) | .NET 10, .NET 8… |
| Format de projet | **Classique** (`.vbproj` ancien style) | SDK-style |
| Concepteur WinForms | Parité historique, mature | Correct, mais arrivé plus tard, avec quelques écarts |
| Interop COM/ActiveX | La plus proche de VB6 | Possible, mais davantage de friction |
| Pour le « pont » VB6 → … | **C'est ce qu'on veut** | À réserver au **second saut** (module 20) |

> ⚠️ **À retenir absolument :** pour la migration depuis VB6, choisissez la variante **« (.NET Framework) »** et le langage **Visual Basic**. Prendre par réflexe le gabarit « moderne » revient à s'imposer deux migrations en une.

---

## Cibler concrètement .NET Framework 4.7.2

Une fois le projet **(.NET Framework)** créé, on fixe (ou on vérifie) la version cible :

> **Propriétés du projet → onglet « Application » → liste déroulante « Target framework » → « .NET Framework 4.7.2 »**

Le réglage est inscrit dans le fichier projet. Selon le format, la ligne diffère :

```xml
<!-- Projet classique (.NET Framework, gabarit WinForms) -->
<TargetFrameworkVersion>v4.7.2</TargetFrameworkVersion>
```

```xml
<!-- Forme SDK-style équivalente (moniker) -->
<TargetFramework>net472</TargetFramework>
```

Savoir lire cette ligne est utile : elle permet de **vérifier la cible directement dans le dépôt** (revue de code, gestion de versions — § 6.5), sans ouvrir l'IDE.

### Ne pas confondre les trois « packs »

| Composant | À quoi il sert | Qui en a besoin |
|---|---|---|
| **Pack de ciblage** *(assemblies de référence)* | **Compiler** contre 4.7.2 → remplit la liste « Target framework » | La machine de **développement** |
| **Developer Pack** | Bundle : pack de ciblage **+** SDK **+** runtime de 4.7.2 | La machine de **développement** *(le plus simple à installer)* |
| **Runtime** | **Exécuter** une application .NET Framework 4.7.2 | Les machines **cibles** (déploiement) |

> 💡 En clair : sur le **poste du développeur**, on installe le **Developer Pack** (ou le pack de ciblage + SDK) pour pouvoir *compiler* en 4.7.2 ; sur les **machines des utilisateurs**, c'est le **runtime** qui doit être présent pour *exécuter* l'application.

---

## VB.NET sur 4.7.2 : un scénario de première classe

Bonne nouvelle, et c'est l'une des raisons du choix du « pont » : **VB.NET + Windows Forms sur .NET Framework 4.7.2 est un terrain pleinement supporté et mature.**

- Le **concepteur de formulaires** fonctionne nativement, avec la parité historique attendue.
- Le **format de projet classique** est celui que connaissent les outils et les habitudes d'entreprise.
- Les **références COM/ActiveX** s'y intègrent de la façon la plus directe (§ 6.3).

C'est exactement le confort recherché : on change de langage et de runtime, **mais pas** d'écosystème d'outils ni de modèle d'interface. Un seul saut à la fois.

---

## Vérifier que l'environnement est correct

Votre environnement est prêt lorsque **tous** ces points sont vrais :

- ✅ Un nouveau projet **« Windows Forms App (.NET Framework) »** en **Visual Basic** se crée sans erreur.
- ✅ Dans *Propriétés → Application*, le champ **« Target framework »** propose **et** affiche **« .NET Framework 4.7.2 »**.
- ✅ Le projet **compile et s'exécute** (F5) : une fenêtre vide s'ouvre.
- ✅ Le **concepteur** de formulaires s'ouvre et permet de déposer un contrôle (un bouton, par exemple).
- ✅ Le fichier `.vbproj` contient bien `<TargetFrameworkVersion>v4.7.2</TargetFrameworkVersion>`.

Si l'un de ces points échoue, le suspect n°1 reste le **pack de ciblage 4.7.2 manquant** (voir plus haut).

---

## Côté déploiement : ne pas oublier la machine cible

Cibler 4.7.2 au développement ne sert à rien si les **machines des utilisateurs** ne disposent pas du **runtime** correspondant. Le runtime .NET Framework 4.7.2 est présent en standard sur plusieurs versions récentes de Windows (et installable sur d'autres versions supportées), mais cela **se vérifie et se planifie**.

> 🔗 Le **packaging** (installeur, prérequis 4.7.2, redistribuables) et la stratégie de bascule sont traités en détail au **§ 18.1**. À ce stade, retenez simplement que la cible d'exécution fait partie des contraintes à confirmer **tôt**.

---

## Points de vigilance

> ⚠️ **Le bon gabarit avant tout.** « (.NET Framework) », pas le gabarit .NET moderne. C'est l'erreur qui transforme une migration en deux.

> ⚠️ **Pack de ciblage = condition pour voir 4.7.2.** Si la version n'apparaît pas, installez le pack de ciblage / Developer Pack, puis redémarrez l'IDE.

> 💡 **4.8/4.8.1 existe — on vise 4.7.2 sciemment.** Ce n'est pas un retard de version, c'est la cible du « pont » (§ 1.2). Le second saut éventuel est au module 20.

> 💡 **Nouveau format de solution `.slnx`.** Visual Studio 2026 peut proposer un format de solution XML plus lisible et interopérable avec les versions antérieures. Rien d'obligatoire ; à décider au moment d'organiser et de versionner la solution (§ 6.2 et § 6.5).

---

## À retenir

> - **Étape 1 de l'environnement** : installer Visual Studio (workload **« Développement .NET desktop »**) et **cibler .NET Framework 4.7.2**.
> - Le **piège central** est le **gabarit de projet** : choisir **« (.NET Framework) »** en **Visual Basic**, jamais le gabarit .NET moderne — sinon on cumule deux migrations.
> - Pour voir et choisir 4.7.2, il faut le **pack de ciblage** (ou le **Developer Pack**) installé sur le poste de développement.
> - **Trois packs à ne pas confondre** : *pack de ciblage* et *Developer Pack* pour **compiler** (développeur), *runtime* pour **exécuter** (machines cibles).
> - **VB.NET sur 4.7.2** est un scénario de **première classe** (concepteur WinForms, projet classique, interop COM) : c'est tout l'intérêt du « pont ».

---

## Liens utiles

- **§ 1.2** — Pourquoi cibler .NET Framework 4.7.2 : le « pont » naturel *(le fondement du choix)*
- **§ 6.2** — Structure de solution et de projets *(comment organiser ce qu'on vient de créer)*
- **§ 6.3** — Référencer les composants COM/ActiveX 🔗 *(la suite naturelle de la mise en place)*
- **§ 6.4** — Les options de compilation *(régler `Option Strict`/`Explicit`/`Infer`/`Compare`)*
- **§ 6.5** — Stratégie de gestion de versions *(versionner la solution et lire la cible dans le `.vbproj`)*
- **§ 18.1** — Packaging .NET Framework *(le runtime 4.7.2 sur les machines cibles)*
- **Module 20** — De 4.7.2 vers .NET moderne *(le second saut, optionnel et séparé)*

⏭️ [Structure de solution et de projets (`.frm`/`.bas`/`.cls` → fichiers `.vb`)](/06-preparer-environnement/02-structure-projets.md)
