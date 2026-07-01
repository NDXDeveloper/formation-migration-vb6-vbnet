🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.3 Outils commerciaux : VBUC et VB Migration Partner 🛠️

> **Module 4 — Outils de migration**
> L'assistant intégré n'existe plus (4.1) et s'appuyait sur une béquille à éviter (4.2). Les
> **outils commerciaux** sont aujourd'hui l'option **automatisée sérieuse** : conçus pour la
> migration VB6 → .NET, bien plus capables que l'ancien assistant, et toujours assortis — c'est le
> *repère honnête* de la formation — de **limites** qu'il faut connaître. Deux produits dominent
> ce marché.

---

## 🧭 Pourquoi passer par un outil commercial

L'*Upgrade Wizard* a disparu de Visual Studio et produisait, on l'a vu, du code rabattu sur l'espace  
de compatibilité (`VB6.*`), 32 bits et obsolète. Les outils commerciaux existent précisément pour
**dépasser ces limites** : convertir davantage, produire un code **plus propre** (idéalement sans
béquille), et **outiller** l'ensemble de la démarche (évaluation, métriques, détection de pièges,  
comparaison de comportement).

Ils ne suppriment pas le besoin de jugement humain — aucun outil ne migre *tout* —, mais ils
**déplacent le curseur** : là où l'assistant traitait le gros œuvre mécanique en laissant une dette
de compatibilité, ces outils visent un code **natif**, maintenable, et accompagnent le travail  
résiduel.

> 💡 Rappel de filiation (section 4.1) : le **moteur** de l'ancien *Upgrade Wizard* provenait
> d'**ArtinSoft**, qui a donné **Mobilize.Net**. L'un des deux outils ci-dessous, **VBUC**, en est
> le **descendant commercial** entretenu et étendu. L'autre, **VB Migration Partner**, a été conçu
> **à partir de zéro** par une autre équipe. Deux lignées, deux philosophies — c'est tout l'intérêt
> de les comparer.

---

## 🏢 Les deux acteurs

### VBUC — *Visual Basic Upgrade Companion* (Mobilize.Net)

C'est l'outil issu de la **même technologie** que l'assistant historique : son moteur a été créé à  
l'origine pour Microsoft, pour accompagner les premières versions de .NET. Ses caractéristiques :

- **Cible VB.NET *ou* C#** : VBUC convertit vers l'un ou l'autre langage, avec un même niveau
  d'automatisation et de qualité — la sortie **C#** est l'un de ses arguments distinctifs.
- **Code natif, sans runtime** : sa philosophie centrale est de produire du **code .NET 100 %
  natif**, **sans dépendance** à un runtime tiers ni à l'espace de compatibilité Microsoft. Il
  applique des règles de **refactoring** pour que le résultat ressemble à du code écrit
  directement en .NET.
- **Couverture large** : résolution du *late binding* (typage), remplacement de la gestion
  d'erreurs, correction des **bornes de tableaux** (passage en 0-based), *control arrays*, accès aux
  données ADO/DAO/RDO → **ADO.NET**, projets multiples, et même projets mixtes ASP/VB6 → ASP.NET.
- **Évolution récente** : l'outil s'est doté d'une variante enrichie par l'**IA** (un « migrateur
  IA » combinant génération assistée et correspondances déterministes) et d'un outil d'**évaluation**
  de la base de code. Le produit est aujourd'hui porté par une entité opérant aussi sous la marque
  **GAPVelocity AI** ; on le rencontre donc sous les deux noms.

> 💡 VBUC peut aussi cibler des versions **modernes** de .NET (.NET Core / .NET récents) — mais,
> dans ce cas, la sortie se fait en **C#** (le couple .NET moderne + VB.NET étant plus contraint).
> C'est un point à garder en tête si l'on anticipe le **second saut** du module 20.

### VB Migration Partner (Code Architects)

Conçu **de zéro** et lancé en **2008** par une équipe d'experts VB/.NET menée par **Francesco  
Balena** (auteur de la série *Programming Microsoft Visual Basic* chez Microsoft Press), cet outil  
mise sur la **fidélité** et le **pilotage fin** de la conversion. Ses caractéristiques :

- **Cible VB.NET (et C#)** : historiquement orienté VB.NET, il prend aussi en charge la génération
  C#.
- **Bibliothèque de support** : les applications converties s'appuient sur une **bibliothèque
  d'exécution maison** (`CodeArchitects.VBLibrary.dll`), à **déployer** avec l'exécutable. Elle
  reproduit fidèlement des mécanismes sans équivalent .NET direct (classes `DataEnvironment`,
  méthodes graphiques, liaison de données…). **Attention** : ce n'est **pas** l'espace de
  compatibilité obsolète de Microsoft (section 4.2), mais une bibliothèque **propre et supportée**
  par l'éditeur — un choix d'architecture différent, pas la béquille à éviter (nuance importante,
  voir plus bas).
- **Les *pragmas*** : la signature de l'outil. Ce sont des **remarques spéciales** insérées dans le
  code VB6 d'origine qui **pilotent** le moteur de conversion (par ex. `ArrayBounds` pour les bornes
  de tableaux, `AutoNew` pour l'auto-instanciation `As New`, `SetType`/`ChangeType` pour forcer un
  type, ou encore des pragmas pour exiger une **finalisation déterministe**). Leur portée est
  **hiérarchique** : projet, fichier, méthode, jusqu'à la **variable** individuelle.
- **Cycle *convert-test-fix*** : grâce aux pragmas, on peut **re-migrer** le même code VB6 autant de
  fois que nécessaire en gardant projets VB6 et .NET **synchronisés** — précieux pour les grosses
  applications **maintenues en parallèle** pendant la migration.
- **Outillage de validation** : un outil **Trace Match** trace l'exécution **côté VB6 et côté .NET**
  pour **comparer** les comportements et localiser les divergences ; des **rapports de métriques**
  (indice cyclomatique, profondeur des structures, ratio de commentaires) exportables vers Excel
  aident à l'estimation.

---

## ⚖️ Deux philosophies opposées — le vrai critère de choix

La différence essentielle entre ces deux outils n'est pas une liste de fonctionnalités : c'est une
**conception différente** de ce que doit être un code migré. La comprendre est plus utile que
n'importe quel tableau comparatif de cases à cocher.

| Axe | **VBUC** (Mobilize.Net) | **VB Migration Partner** (Code Architects) |
|-----|-------------------------|---------------------------------------------|
| **Philosophie** | Code **natif**, **sans runtime** : transformer en .NET idiomatique | **Haute fidélité** via **bibliothèque de support** + pragmas |
| **Dépendance d'exécution** | Aucune (objectif : zéro runtime tiers) | `CodeArchitects.VBLibrary.dll` à déployer |
| **Pilotage** | Options + extensibilité, customisations par l'éditeur | **Pragmas** granulaires (jusqu'à la variable) dans le code VB6 |
| **Re-migration** | Conversion (puis on travaille la sortie) | **Convert-test-fix** : re-migration répétée, VB6/.NET en phase |
| **Langage cible** | VB.NET **ou C#** (C# est un point fort) | VB.NET (et C#) |
| **Lignée** | Descendant du moteur de l'*Upgrade Wizard* | Conçu **de zéro** (2008) |

> 💡 **Comment lire cette opposition.** L'approche **« natif, sans runtime »** (VBUC) vise le code le
> plus **propre** et le moins dépendant — cohérent avec l'esprit de la section 4.2 (supprimer les
> béquilles) et plutôt favorable au **second saut** vers .NET moderne (module 20). L'approche
> **« haute fidélité + bibliothèque »** (VB Migration Partner) vise l'**équivalence de comportement**
> la plus exacte et le **pilotage** le plus fin, au prix d'une **dépendance** d'exécution à
> déployer ; son cycle *convert-test-fix* épouse remarquablement la stratégie **incrémentale** et la
> **cohabitation** (3.4, 3.5). Aucune n'est « meilleure » dans l'absolu : le bon choix dépend de
> votre contexte (criticité, taille, exigence de pureté du code, horizon de modernisation).

> ⚠️ **Méfiance envers le marketing.** Chaque éditeur défend sa philosophie avec vigueur — l'un
> raille les approches « à runtime », l'autre l'« épouvantable » assistant d'origine, et tous deux
> avancent des taux de réussite spectaculaires. Ces affirmations sont des **arguments commerciaux**,
> pas des mesures neutres. Évaluez sur **votre** code (les deux proposent des essais), pas sur les
> brochures.

---

## 🎯 Les trois capacités mises en avant

Le sommaire pointe trois capacités précises pour lesquelles ces outils dépassent l'assistant  
historique. Voici comment chacun les traite.

### 1. Conversion vers VB.NET *ou* C#

C'est un acquis majeur par rapport à l'*Upgrade Wizard* (VB.NET uniquement). **VBUC** convertit  
vers **C# comme vers VB.NET** avec la même qualité — un atout si la cible stratégique de  
l'organisation est C#. **VB Migration Partner**, historiquement VB.NET, prend également en charge
**C#**. Dans les deux cas, viser C# implique de gérer des **différences de langage** supplémentaires
(typage strict, opérateurs sur `Object`/`Variant`, déclaration des tableaux, variable de boucle
`foreach`…) que les outils s'efforcent de combler automatiquement.

### 2. Suppression — ou gestion — des dépendances de compatibilité

C'est ici que la **philosophie** de chaque outil se concrétise, et la nuance est cruciale :

- **VBUC** vise la **suppression** pure et simple des dépendances de compatibilité : il génère du
  **code natif** sans s'appuyer sur l'espace `Microsoft.VisualBasic.Compatibility`. C'est
  exactement la cible recommandée en 4.2 — pas de béquille à rembourser ensuite.
- **VB Migration Partner** ne « supprime » pas les dépendances : il **remplace** la béquille
  Microsoft par sa **propre bibliothèque** (`CodeArchitects.VBLibrary.dll`), supportée et bien plus
  complète. Le code ne dépend donc **pas** de l'espace obsolète de 4.2, mais il dépend **d'un
  runtime de l'éditeur** — un arbitrage différent, à intégrer dans la décision (notamment au regard
  du module 20 : toute dépendance d'exécution est un paramètre du second saut).

> 💡 **Ne pas confondre les trois « runtimes ».** (a) L'espace `Microsoft.VisualBasic.Compatibility`
> = la **béquille obsolète à éviter** (4.2). (b) `CodeArchitects.VBLibrary.dll` = la **bibliothèque
> de support supportée** de VB Migration Partner, un choix d'architecture assumé. (c) Le **runtime
> VB.NET** `Microsoft.VisualBasic`, légitime (4.2). Trois choses distinctes, qu'il serait fâcheux de
> mélanger en évaluant un outil.

### 3. Gestion des *control arrays*

Les **tableaux de contrôles** ont **disparu** en .NET (module 13.6) : c'est l'un des points durs  
classiques (3.3). L'assistant historique les gérait mal (ou via la béquille). Les deux outils  
commerciaux font nettement mieux : chacun fournit des **objets de remplacement** qui se comportent  
comme le contrôle VB6, permettant à la fois de **créer des contrôles à l'exécution** et de
**centraliser la gestion d'événements** (les deux usages d'origine d'un *control array*). VB
Migration Partner prend en charge ce mécanisme de façon transparente, y compris la méthode
`Controls.Add` que l'assistant ne savait pas convertir. C'est typiquement le genre de point dur où
un outil commercial fait gagner un temps considérable face à une reprise manuelle.

---

## 🔗 Ce qu'ils apportent au-delà de la simple conversion

Ces outils ne se limitent pas à transcrire du code : ils **outillent la démarche** entière de la  
formation, et c'est souvent là leur vraie valeur.

- **Évaluation et métriques** → ils produisent des rapports de **complexité** (indice cyclomatique),
  de dépendances et de volume qui alimentent directement l'**estimation de l'effort** (module 3.3).
- **Détection de pièges silencieux** → ils repèrent et traitent certains pièges de l'**Annexe B**
  là où l'assistant restait muet : par exemple la **sémantique par valeur** d'un champ passé en
  paramètre (le piège `ByRef`/`ByVal`), l'auto-instanciation `As New`, ou les subtilités de `Null`
  → `DBNull`. Un appui précieux pour la **traque des pièges** (module 17.4).
- **Comparaison de comportement** → l'outil **Trace Match** de VB Migration Partner (exécution
  comparée VB6/.NET) est, dans l'esprit, un **aide au *golden master*** et à la **non-régression**
  (modules 5.5, 17.3).
- **Composants tiers et ActiveX** → ils proposent des **correspondances** vers des équivalents .NET
  natifs ou des versions .NET récentes de contrôles tiers (modules 3.2, 14.6), et permettent de
  conserver des contrôles « en attente » via **interop COM** (module 16) — soutenant ainsi une
  migration **incrémentale**.

---

## ⚠️ Lucidité : ce que les outils commerciaux **ne** font **pas**

Aussi capables soient-ils, ils restent des outils — et le *repère honnête* de la formation  
s'applique :

- **Ils ne migrent pas tout.** On reste dans la logique des 60-90 % de **volume** mécanique ; la
  strate « rouge » (points durs, UI à repenser, pièges résiduels) demande toujours du **jugement
  humain** (et, bien encadrée, l'IA — module 19).
- **Leur sortie doit être validée.** Comme pour tout outil, le code produit est un **point de
  départ** à relire, fiabiliser (module 17) et **tester contre le *golden master*** (5.5, 17.3).
  « Ça compile et ça tourne » ne prouve pas l'équivalence de comportement.
- **L'UI demande souvent des ajustements** (polices, couleurs, disposition, twips → pixels — module
  13.4), même quand la conversion logique est excellente.
- **Ils ont un coût.** La tarification se fait généralement au **volume de code** (lignes
  effectives), par application ; à mettre en regard du temps économisé (et des essais gratuits
  disponibles pour évaluer avant d'acheter).
- **Ils réduisent le risque de pièges silencieux sans l'annuler.** Ils en détectent beaucoup, mais
  la responsabilité finale de la **non-régression** reste au projet.

---

## ✅ En résumé

- Les **outils commerciaux** sont l'option automatisée sérieuse depuis la disparition de l'*Upgrade
  Wizard* (4.1) : plus capables, et visant un code **plus propre** que la béquille de 4.2.
- **VBUC** (Mobilize.Net, aussi sous la marque GAPVelocity AI) descend du moteur de l'assistant
  historique. Sa philosophie : **code natif, sans runtime**, vers **VB.NET *ou* C#** ; il **supprime**
  les dépendances de compatibilité. Variante enrichie par l'**IA**, ciblage possible de .NET
  moderne (en C#).
- **VB Migration Partner** (Code Architects, Francesco Balena, 2008) mise sur la **haute fidélité**
  via une **bibliothèque de support** (`CodeArchitects.VBLibrary.dll`, distincte de la béquille
  obsolète) et sur les ***pragmas*** granulaires pilotant la conversion, avec un cycle
  **convert-test-fix** idéal pour les grosses applications maintenues en parallèle.
- **Le vrai critère de choix** est philosophique : **« natif sans runtime »** (plus propre, mieux
  armé pour le module 20) **vs « haute fidélité avec bibliothèque + pragmas »** (équivalence exacte,
  pilotage fin, affinité avec l'incrémental). Aucune n'est universellement supérieure.
- Sur les **trois capacités** du sommaire : tous deux ciblent **VB.NET ou C#**, **traitent** les
  dépendances de compatibilité (suppression pour VBUC, remplacement par un runtime propre pour
  VBMP), et **gèrent les *control arrays*** via des objets de remplacement (module 13.6).
- Ils **outillent** aussi l'évaluation (métriques → 3.3), la détection de pièges (Annexe B → 17.4)
  et la comparaison de comportement (Trace Match → golden master 5.5/17.3).
- **Lucidité** : ils ne migrent pas tout, leur sortie se **valide** et se **teste** (module 17),
  l'UI demande des retouches, ils ont un **coût**, et la **non-régression** reste la responsabilité
  du projet.

---

⬅️ Section 4.2 ⚠️ — [L'espace de noms `Microsoft.VisualBasic.Compatibility`](02-compatibility-namespace.md)  
➡️ Section 4.4 — [Migration manuelle assistée : quand et pourquoi](04-migration-manuelle.md)

---

**Juin 2026** · .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Migration manuelle assistée : quand et pourquoi](/04-outils-migration/04-migration-manuelle.md)
