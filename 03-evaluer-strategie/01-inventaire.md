🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.1 Inventaire ⭐

> **Module 3 — Évaluer l'existant et choisir une stratégie**
> Avant de cartographier les dépendances (3.2), de chiffrer l'effort (3.3) ou de choisir une
> stratégie (3.4), il faut d'abord **savoir ce que l'on a**. Cette section établit la
> photographie de départ : tous les constituants du projet VB6, type par type.

---

## 🧭 Pourquoi commencer par l'inventaire

Un projet VB6 n'est pas un bloc monolithique : c'est un assemblage de fichiers de natures  
différentes, chacun porteur d'un **coût de migration distinct**. Un module standard `.bas` rempli  
de logique métier pure ne pose pas les mêmes problèmes qu'un formulaire `.frm` truffé de
*control arrays*, ou qu'une classe `.cls` qui s'appuie sur `Class_Terminate`.

L'inventaire répond à trois questions simples mais décisives :

1. **Combien y a-t-il à migrer ?** (volume brut : nombre de fichiers, de lignes)
2. **De quelle nature ?** (formulaires, logique, classes, composants externes)
3. **Où se concentrent les points durs ?** (OCX non maintenus, API Win32, `Variant` partout)

> ⚠️ **Un inventaire n'est pas un simple comptage de fichiers.** Compter « 142 `.frm` » ne dit
> rien de l'effort réel. Ce qui compte, c'est ce que chaque fichier **contient** : un formulaire
> de 30 contrôles avec dessin GDI et impression coûte dix fois plus qu'une boîte de dialogue
> « OK/Annuler ». L'inventaire qualitatif prime sur l'inventaire quantitatif.

---

## 📁 Anatomie d'un projet VB6 : les types de fichiers

Tout commence par le fichier projet. Voici les extensions que vous rencontrerez, et ce que  
chacune vous apprend.

### Les fichiers de structure du projet

| Extension | Rôle | Ce qu'il vous apprend |
|-----------|------|------------------------|
| **`.vbp`** | Fichier projet (*Visual Basic Project*) | **Le point d'entrée de l'inventaire** : il liste tous les fichiers, références et composants |
| **`.vbg`** | Groupe de projets (*Visual Basic Group*) | Plusieurs projets liés (ex. une appli + ses DLL ActiveX maison) |
| **`.vbw`** | Position des fenêtres dans l'IDE | Sans intérêt pour la migration (préférences d'affichage) |

Le fichier **`.vbp` est votre meilleur ami** : c'est un fichier texte que vous pouvez ouvrir dans  
un éditeur. Il recense, sous forme de lignes lisibles, l'intégralité des constituants du projet.  
Nous y revenons en détail plus bas.

### Les fichiers de code et d'interface

| Extension | Type | Contenu | Points d'attention migration |
|-----------|------|---------|------------------------------|
| **`.frm`** | Formulaire | Disposition **et** code d'un formulaire | *Control arrays*, twips, dessin GDI, cycle de vie ⚠️ |
| **`.frx`** | Ressources binaires d'un `.frm` | Images, icônes, données binaires des contrôles | À convertir en `.resx` (module 14.3) |
| **`.bas`** | Module standard | Procédures, fonctions, variables globales | Souvent le plus simple à migrer (logique pure) |
| **`.cls`** | Module de classe | Définition d'une classe | `Class_Initialize`/`Terminate`, finalisation ⚠️ |
| **`.ctl`** | UserControl (contrôle ActiveX maison) | Code et disposition d'un contrôle créé en VB6 | Cas avancé : à recréer ou à porter |
| **`.ctx`** | Ressources binaires d'un `.ctl` | Équivalent du `.frx` pour les UserControls | Idem `.frx` |
| **`.pag`** | Page de propriétés (*Property Page*) | UI de configuration d'un contrôle ActiveX | Lié aux UserControls |
| **`.dsr`** | Designer (DataEnvironment, etc.) | Composants à conception visuelle | Souvent lié à l'accès aux données (module 15) ⚠️ |
| **`.dob`** | Document ActiveX | Document hébergé (rare) | Cas particulier, à évaluer isolément |
| **`.res`** | Fichier de ressources | Ressources au niveau projet | À porter vers le modèle de ressources .NET |

> 💡 **Lecture rapide d'un projet** : la **proportion** entre ces fichiers vous donne déjà le
> profil de la migration. Beaucoup de `.frm` et peu de `.bas` → migration **à dominante UI**
> (module 13-14). Beaucoup de `.cls` → migration **à dominante objet** (module 12, attention à la
> finalisation). Présence de `.ctl`/`.dsr` → **points durs** à traiter à part.

---

## 🔗 Les références et les composants externes

C'est ici que se logent les **dépendances** — et donc une grande partie du risque. Le projet VB6  
s'appuie sur du code qu'il ne contient pas : bibliothèques COM, contrôles ActiveX, type  
libraries. La section 3.2 cartographiera ces dépendances en profondeur ; ici, on les **recense**.

Dans le fichier `.vbp`, ces éléments apparaissent sous deux formes :

- Les lignes **`Reference=`** : ce sont les **références** à des bibliothèques de types (COM/DLL)
  utilisées par *early binding* — par exemple ADO (`Microsoft ActiveX Data Objects`), la
  bibliothèque Scripting (`FileSystemObject`), ou des DLL métier maison.
- Les lignes **`Object=`** : ce sont les **composants** ActiveX, typiquement les contrôles
  visuels **OCX** déposés sur les formulaires (grilles, calendriers, barres d'outils tierces,
  contrôles communs Windows comme le `MSComctlLib`).

Pour chaque référence et chaque composant, l'inventaire doit noter :

1. **Le nom et l'éditeur** (Microsoft ? un tiers ? interne ?) ;
2. **Sa version** (le GUID et le numéro de version sont dans la ligne du `.vbp`) ;
3. **Son statut** : encore disponible et maintenu, ou abandonné ?
4. **Son rôle** : visuel (OCX sur formulaire) ou non visuel (bibliothèque appelée par code) ?

> ⚠️ **Le point dur classique : l'OCX tiers non maintenu.** Une grille de données d'un éditeur
> disparu, un contrôle de calendrier dont plus personne n'a la licence, un composant d'impression
> propriétaire… Ces éléments n'ont **pas d'équivalent direct** et exigent souvent un **remplacement**
> (module 13.6, 14.6) plutôt qu'une simple conversion. Les repérer **dès l'inventaire** évite la
> mauvaise surprise tardive. Un OCX oublié, c'est potentiellement plusieurs jours de travail
> imprévus.

---

## 🪟 Les déclarations d'API Windows (`Declare`)

Les appels directs au système d'exploitation — via l'instruction `Declare` qui pointe vers des  
DLL Win32 comme `kernel32`, `user32` ou `gdi32` — méritent un **recensement à part entière**. Ils  
ne se voient pas dans le `.vbp` : ils sont **disséminés dans le code**, généralement en tête des  
modules `.bas`.

Pourquoi les traiter spécifiquement ?

- Ils devront être **convertis de `Declare` vers `P/Invoke` (`DllImport`)** — module 16.4 ;
- Ils sont une **source majeure de pièges silencieux** : passage `ByRef`/`ByVal`, marshaling des
  chaînes **ANSI vs Unicode**, correspondance des types et des structures (`Type` → `Structure`
  avec `StructLayout`) — module 16.5 ⚠️ ;
- Leur nombre et leur complexité sont un **indicateur de difficulté** : une application qui
  multiplie les appels Win32 de bas niveau migre moins facilement qu'une application « pure VB ».

L'inventaire consiste à **lister toutes les déclarations `Declare`** du projet (un balayage du  
code à la recherche du mot-clé suffit pour un premier recensement), en notant pour chacune la DLL  
ciblée, la fonction appelée et la présence de structures (`Type`) dans la signature.

---

## 🛠️ Comment mener l'inventaire concrètement

L'inventaire combine **lecture de fichiers** et **balayage de code**. Plusieurs approches se  
complètent :

### Le fichier `.vbp` comme socle

Ouvert dans un éditeur de texte, le `.vbp` donne, gratuitement et immédiatement :

- la **liste exhaustive des fichiers** du projet (`Form=`, `Module=`, `Class=`, `UserControl=`…) ;
- la **liste des références** (`Reference=`) et **composants** (`Object=`) ;
- des **réglages de compilation** (`Startup=`, options diverses) utiles pour comprendre le projet.

C'est le **premier réflexe** : il fournit l'ossature de l'inventaire sans ouvrir l'IDE. Un extrait  
de `.vbp` ressemble à ceci (GUID abrégés pour la lisibilité) :

```ini
Type=Exe
Reference=*\G{00020430-…}#2.0#0#…\stdole2.tlb#OLE Automation
Reference=*\G{EF534B90-…}#2.8#0#…#Microsoft ActiveX Data Objects 2.8 Library
Object={831FDD16-…}#2.0#0; MSCOMCTL.OCX
Form=frmPrincipal.frm
Form=frmFacture.frm
Module=modGlobal; modGlobal.bas
Class=CFacture; CFacture.cls
Startup="Sub Main"
```

Chaque ligne est une **piste d'inventaire** : on y repère aussitôt les **références COM** (ici OLE  
Automation et ADO), le **contrôle ActiveX** `MSCOMCTL.OCX`, les **fichiers** du projet et le
**point d'entrée** (`Sub Main`).

### Le balayage du code source

Au-delà du `.vbp`, certaines informations ne se trouvent qu'**en lisant le code**. Un recensement  
par recherche textuelle (sur l'ensemble des `.frm`, `.bas`, `.cls`) permet de localiser les
éléments à risque que l'inventaire doit chiffrer :

| À rechercher | Pourquoi c'est important | Module concerné |
|--------------|--------------------------|-----------------|
| `Declare` | Appels d'API Win32 à porter en P/Invoke | 16.4 |
| `Class_Terminate` | Finalisation déterministe — **piège n°1** | 2.2, 12.2 |
| `On Error Resume Next` | Masque potentiellement des bugs | 11.3 |
| `Variant` / `As Variant` | Typage faible à clarifier | 7.1 |
| `Index` (sur contrôles) | Présence de *control arrays* | 13.6 |
| `Type ... End Type` | Structures à porter (`StructLayout`) | 12.7, 16.5 |
| `GoSub` / `On ... GoTo` | Constructions **supprimées** en VB.NET | 9.5 |
| `String * n` | Chaînes de longueur fixe | 7.4 |

> 💡 Ce balayage ne sert pas (encore) à corriger : il sert à **mesurer**. Combien de
> `Class_Terminate` ? Combien de `Declare` ? Combien de modules au `Variant` envahissant ? Ces
> chiffres alimentent directement l'estimation de l'effort (3.3) et le repérage des zones à
> risque (Annexe B).

### Les outils d'aide à l'inventaire

L'inventaire manuel atteint vite ses limites sur un gros projet. Plusieurs catégories d'outils  
peuvent l'assister :

- les **outils commerciaux de migration** (VBUC de Mobilize.Net, VB Migration Partner de Code
  Architects — module 4.3) proposent généralement une **phase d'analyse / d'évaluation** qui
  produit des rapports de constituants, de dépendances et de complexité ;
- l'**assistant de mise à niveau** de Visual Studio (module 4.1) signale en cours de route les
  éléments qu'il ne sait pas convertir — une forme indirecte d'inventaire des points durs ;
- des **scripts maison** (recherche textuelle, comptage de lignes) suffisent pour un premier
  débroussaillage des éléments du tableau ci-dessus.

---

## 📋 Ce que doit produire l'inventaire

À l'issue de cette étape, vous disposez d'un **état des lieux structuré** — un document de
référence qui servira à tout le reste du module. Il recense au minimum :

- **Les fichiers**, par type (nombre de `.frm`, `.bas`, `.cls`, `.ctl`, `.dsr`…) et idéalement
  leur taille (lignes de code), pour repérer les plus volumineux ;
- **Les références et composants externes** (COM, OCX, DLL), avec éditeur, version et **statut de
  maintenance** — la matière première de la cartographie 3.2 ;
- **Les déclarations d'API Win32** (`Declare`), avec DLL et fonction ciblées ;
- **Les éléments à risque repérés au balayage** (finalisation, `On Error Resume Next`, *control
  arrays*, `Variant`, structures, constructions supprimées…), avec leur **volume** ;
- **Les zones grises** : tout ce dont le statut reste à éclaircir (un composant non identifié, une
  DLL maison sans documentation).

> 📝 **Cet inventaire est un livrable, pas une note jetable.** Il sera relu à chaque étape :
> pour chiffrer (3.3), pour décider de la stratégie (3.4), pour définir le périmètre (3.6) et
> pour vérifier, en fin de migration, qu'on n'a **rien oublié** (Annexe E — Checklist).

---

## ⚠️ Les pièges de l'inventaire lui-même

Quelques erreurs classiques au moment de recenser l'existant :

- **Confondre comptage et compréhension.** « 200 fichiers » n'est pas une estimation d'effort.
  Deux projets de même taille brute peuvent différer d'un facteur dix selon leur contenu.
- **Oublier les dépendances invisibles.** Une DLL chargée dynamiquement, un composant instancié
  par `CreateObject` (late binding) n'apparaît pas forcément dans les références du `.vbp`. Le
  balayage du code est indispensable pour les débusquer.
- **Négliger le code mort.** Une partie du code inventorié n'est peut-être plus utilisée. Le
  repérer évite de migrer (et de tester) inutilement — c'est précisément l'objet du **nettoyage
  en amont** (module 5.1).
- **Sous-estimer les ressources binaires.** Les `.frx`/`.ctx`/`.res` (images, icônes) doivent
  être portés eux aussi ; on les oublie facilement car ils ne sont pas « du code ».
- **Faire l'inventaire une seule fois.** L'existant peut évoluer si le périmètre n'est pas gelé
  (module 5.6). Un inventaire fait six mois avant la migration mérite une **réactualisation**.

---

## ✅ En résumé

- L'inventaire est la **photographie de départ** : il transforme une impression vague en une
  connaissance précise de **quoi** migrer.
- Le fichier **`.vbp`** est le socle : il liste fichiers, références et composants en texte clair.
- Chaque **type de fichier** (`.frm`, `.bas`, `.cls`, `.ctl`, `.dsr`, `.frx`…) porte un coût et
  des pièges propres ; leur **proportion** révèle le profil de la migration.
- Les **références/composants externes** (COM, OCX, DLL) et les **déclarations d'API** (`Declare`)
  concentrent une grande part du risque — l'OCX tiers non maintenu en tête.
- L'inventaire est **qualitatif autant que quantitatif** : un balayage du code repère les
  éléments à risque (finalisation, `On Error Resume Next`, *control arrays*, `Variant`…) et en
  mesure le volume.
- Le résultat est un **livrable structuré**, réutilisé à chaque étape suivante du module et
  jusqu'à la checklist finale.

---

⬅️ Module 3 — [Évaluer l'existant et choisir une stratégie](README.md)  
➡️ Section 3.2 — [Cartographie des dépendances](02-dependances.md)

---

**Juin 2026** · .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Cartographie des dépendances (COM, ActiveX, bases de données, composants tiers)](/03-evaluer-strategie/02-dependances.md)
