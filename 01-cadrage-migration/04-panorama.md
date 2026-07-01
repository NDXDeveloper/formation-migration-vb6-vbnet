🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 1.4 Panorama complet d'une migration : langage, formulaires, données, COM, API

> **Chapitre 1 — Introduction : pourquoi et comment migrer**
> Avant de plonger, prenons de la hauteur : *de quoi est faite, au juste, une migration ?*

---

## Une migration n'est pas un chantier — c'est cinq chantiers

L'erreur de planification la plus fréquente consiste à voir la migration comme **une seule tâche**
(« convertir le code ») que l'on découvre ensuite, difficulté après difficulté, au fil de l'eau.
La réalité est qu'une application VB6 repose sur **cinq fronts distincts**, et que chacun a sa  
logique, son effort et ses pièges propres :

1. **Le langage** — le code lui-même (le cœur du travail) ;
2. **Les formulaires** — l'interface utilisateur ;
3. **Les données** — l'accès aux bases ;
4. **L'interopérabilité COM** — les composants ActiveX/OCX et la cohabitation ;
5. **Les API Windows** — les appels système.

Le but de cette section est de **dérouler la carte complète** afin que vous puissiez **anticiper  
l'ampleur réelle** du chantier, plutôt que de la subir par morceaux. Chaque front est ensuite  
traité en profondeur dans sa propre partie de la formation ; ici, on **cartographie**, on ne creuse  
pas encore.

> 💡 **Tous les fronts ne concernent pas toutes les applications.** Une application sans contrôle
> ActiveX allège le front COM ; une application sans `Declare` ignore le front API. C'est
> précisément le rôle de l'**inventaire** (section 3.1) que de déterminer **lesquels de ces cinq
> fronts** s'appliquent à *votre* code — et avec quel poids.

---

## 🗺️ La carte en un coup d'œil

| # | Front | De (VB6) | Vers (.NET Framework 4.7.2) | Détaillé dans |
|---|---|---|---|---|
| 1 | **Langage** | Types, syntaxe, POO, erreurs | VB.NET (types redimensionnés, exceptions, héritage) | **Partie 3** (chap. 7-12) |
| 2 | **Formulaires** | `.frm`/`.frx`, contrôles VB6 | **Windows Forms** + concepteur | **Partie 4** (chap. 13-14) |
| 3 | **Données** | DAO / RDO / **ADO**, `Recordset` | **ADO.NET** (`DataReader`/`DataSet`) | **Partie 5** (chap. 15) |
| 4 | **COM** | ActiveX/OCX, composants COM | **Interop** (RCW/CCW), cohabitation | **Partie 5** (chap. 16) |
| 5 | **API Windows** | `Declare` (Win32) | **P/Invoke** (`DllImport`) | **Partie 5** (chap. 16) |

Une façon utile de se représenter ces fronts : **le langage est la fondation**, et les quatre  
autres sont les murs qui s'y appuient.

```
   ┌────────────────────────────────────────────────────────┐
   │   FORMULAIRES        DONNÉES        COM       API      │
   │   (Windows Forms)    (ADO.NET)   (Interop)  (P/Invoke) │
   ├────────────────────────────────────────────────────────┤
   │                  LE LANGAGE (VB.NET)                   │  ← la fondation
   ├────────────────────────────────────────────────────────┤
   │              LE RUNTIME (CLR + ramasse-miettes)        │  ← change dessous (1.3)
   └────────────────────────────────────────────────────────┘
```

---

## 1️⃣ Le langage — le cœur, et le plus gros volume

C'est **la** partie centrale : convertir le code lui-même. C'est généralement le **plus gros volume  
de travail**, parce qu'elle touche **chaque ligne** de chaque module, classe et formulaire.

Ce front recouvre :

- les **types et déclarations** (`Variant`→`Object`, entiers redimensionnés, `Currency`→`Decimal`, dates, chaînes de longueur fixe) ;
- les **tableaux et collections** (tableaux désormais 0-based, `Collection`→`List`/`Dictionary`) ;
- les **opérateurs et structures de contrôle** (`GoSub`/`On…GoTo` supprimés, `While…Wend`…) ;
- les **procédures et paramètres** (le fameux **ByRef → ByVal par défaut**) ;
- la **gestion d'erreurs** (`On Error` → `Try`/`Catch`/`Finally`) ;
- la **POO** (constructeurs, héritage, finalisation, événements).

> ⚠️ **Rappel du 1.3** : c'est ici que se cachent la plupart des **changements silencieux** — ceux
> qui compilent mais changent le comportement. Le front « langage » n'est donc pas seulement le plus
> volumineux : c'est aussi celui qui concentre le plus de **risques de régression**.

👉 **Partie 3 (chapitres 7 à 12)**, avec l'aide-mémoire de l'**annexe A** et le catalogue de pièges
de l'**annexe B**.

---

## 2️⃣ Les formulaires — souvent le plus fastidieux

Vos formulaires VB6 (`.frm` et leurs ressources `.frx`) migrent vers **Windows Forms**. Si la
**parité** est bonne sur .NET Framework (c'est l'un des arguments du choix de 4.7.2, voir 1.2), ce
front réserve un travail **manuel, minutieux et parfois ingrat**, car il mêle conversion de code
**et** reconstruction d'interface.

Les points qui demandent le plus d'attention :

- **Twips → pixels** : VB6 positionne et dimensionne en *twips* (1/1440 de pouce), Windows Forms en pixels. Toute la géométrie (position, taille, ancrage) doit être **convertie**.
- **Les *control arrays* (tableaux de contrôles) : disparus.** Une construction très utilisée en VB6, sans équivalent direct — il faut la **recréer** (à la main, par partage de gestionnaire d'événements, ou via une bibliothèque dédiée).
- **Le cycle de vie** change (`Load`/`Unload`/`Activate` → `Load`/`FormClosing`/`Shown`), de même que les **formulaires MDI** et les **menus** (Menu Editor → `MenuStrip`).
- **Le dessin et l'impression** : le dessin VB6 (`Line`, `Circle`, `PSet`) passe à **GDI+**, et l'objet `Printer` à `System.Drawing.Printing`.

👉 **Partie 4 (chapitres 13 et 14)**, avec la correspondance des contrôles en **annexe C**.

---

## 3️⃣ Les données — un changement de paradigme, pas seulement d'API

L'accès aux données passe des technologies VB6 (**DAO**, **RDO**, **ADO**) à **ADO.NET**. Au-delà  
du remplacement d'API, c'est un **changement de modèle** qu'il faut assimiler.

- Le **`Recordset`** (curseur connecté, navigable) éclate en plusieurs notions distinctes : **`DataReader`** (flux connecté, rapide, en avant seulement) et **`DataSet`/`DataTable`** (modèle **déconnecté**). Choisir le bon outil selon l'usage est une décision de conception.
- La **liaison de données** évolue (`Data` control / *data binding* VB6 → `BindingSource`).
- Le **reporting** mérite une vigilance particulière : **Crystal Reports**, très présent dans le *legacy*, a sa propre histoire de compatibilité (et ses alternatives).
- Bonus appréciable : la migration est l'occasion d'introduire les **commandes paramétrées**, et donc de **corriger au passage** des vulnérabilités d'injection héritées.

> ⚠️ Le piège n'est pas tant l'API que la **sémantique** : un code qui présumait un curseur connecté
> navigable ne se comporte pas pareil avec un `DataReader` en avant seulement. *Changement
> silencieux, encore.*

👉 **Partie 5 (chapitre 15)**.

---

## 4️⃣ L'interopérabilité COM — le front qui rend la migration *progressive*

Beaucoup d'applications VB6 reposent sur des **composants COM** et des **contrôles ActiveX/OCX**
(grilles, calendriers, contrôles métier). Ce front répond à deux questions : *comment réutiliser
l'existant ?* et — surtout — *comment migrer sans tout basculer d'un coup ?*

- **Appeler du COM depuis .NET** se fait via un **RCW** (*Runtime Callable Wrapper*) : on réutilise un composant COM existant comme s'il était du .NET.
- **Exposer du .NET à du VB6** se fait via un **CCW** (*COM Callable Wrapper*) : c'est **la clé de la migration incrémentale**, car elle permet de faire **cohabiter** code VB6 et code .NET pendant toute la transition.
- Reste à arbitrer entre *early binding* et *late binding* (`Option Strict Off`), et à décider, contrôle par contrôle, ce qu'on **réutilise** par interop et ce qu'on **remplace**.

> 💡 **Le front stratégique.** C'est l'interop COM qui transforme un *big-bang* risqué en
> **migration par étapes validables**. Sa maîtrise conditionne la possibilité même d'une transition
> en douceur (voir la stratégie de cohabitation, section 3.5).

👉 **Partie 5 (chapitre 16)**.

---

## 5️⃣ Les API Windows — localisé, mais technique

Les applications VB6 déclarent souvent des fonctions de l'**API Win32** via l'instruction
**`Declare`**. En VB.NET, ces appels passent au **P/Invoke** (attribut **`DllImport`**). Le front
est généralement **circonscrit** (un nombre limité de déclarations), mais **techniquement pointu**.

- La **traduction `Declare` → `DllImport`** n'est pas qu'un changement de syntaxe : les **types** des paramètres doivent être réétudiés (rappel : les entiers ont changé de taille).
- Le **marshaling** et la distinction **ANSI / Unicode** deviennent critiques : une chaîne mal marshalée ou un mauvais jeu de caractères produit des bugs subtils.
- Les structures passées aux API (`Type…End Type` → **`Structure`**) exigent souvent un **`StructLayout`** explicite pour garantir une disposition mémoire correcte.

> ⚠️ Front à faible volume mais à **fort potentiel de bug silencieux** : un appel d'API mal traduit
> peut « presque » fonctionner, puis échouer dans des cas particuliers (chaînes longues, caractères
> accentués, plateformes 64 bits via WOW64).

👉 **Partie 5 (chapitre 16, sections 16.4 et 16.5)**.

---

## 🔗 Les fronts ne sont pas étanches

Présenter cinq fronts aide à structurer le travail, mais sur le terrain, **ils se recoupent  
constamment** :

- un **formulaire** qui héberge un contrôle **ActiveX** = front *formulaires* **+** front *COM* ;
- un écran de saisie **lié à une base** = front *formulaires* **+** front *données* ;
- un appel d'**API** depuis une classe = front *API* **+** front *langage* (types, structures) ;
- et **partout**, en dessous, les changements de **runtime** et de **GC** vus en 1.3.

C'est pourquoi on n'avance pas front par front en silos parfaitement isolés, mais par **tranches  
verticales** cohérentes (un module, un écran, une fonctionnalité), en s'appuyant sur la
**cohabitation COM** pour valider à chaque étape (stratégies détaillées au chapitre 3).

---

## 🛡️ Un fil rouge commun aux cinq fronts : la non-régression

Quel que soit le front, la même exigence revient : **prouver que le comportement n'a pas changé**.  
C'est le rôle, **transverse à toute la migration**, de deux dispositifs :

- le **harnais de tests de référence** (*golden master*, section 5.5 et annexe F), qui **capture le comportement de l'application VB6 avant** migration pour le **comparer après** ;
- le **catalogue des pièges silencieux** (annexe B), qui recense, front par front, les régressions classiques à traquer.

Ces deux ressources ne sont pas un front de plus : elles sont le **filet de sécurité** tendu sous  
les cinq.

---

## ✅ À retenir

- Une migration VB6 → VB.NET se décompose en **cinq fronts** : **langage**, **formulaires**, **données**, **COM**, **API Windows** — chacun avec son effort et ses pièges propres.
- Le **langage** est la **fondation** et le **plus gros volume** (et concentre les changements silencieux) ; les **formulaires** sont souvent les plus **fastidieux** ; les **données** changent de **paradigme** (connecté → déconnecté) ; le **COM** est le front **stratégique** qui rend la migration **progressive** ; les **API** sont **localisées** mais techniquement délicates.
- **Tous les fronts ne s'appliquent pas à toutes les applications** : c'est l'**inventaire** (3.1) qui le détermine.
- Les fronts **se recoupent** ; on migre par **tranches verticales**, pas en silos.
- Un **filet de sécurité transverse** (golden master + catalogue des pièges) accompagne les cinq fronts du début à la fin.

---

> 🧭 **Section suivante** → [1.5 Mythes et réalités](05-mythes-realites.md)

⏭️ [Mythes et réalités (« l'assistant fait tout », « ce n'est que de la syntaxe »)](/01-cadrage-migration/05-mythes-realites.md)
