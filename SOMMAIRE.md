# 📚 Sommaire — Formation Migration VB6 → VB.NET (.NET Framework 4.7.2)
**Migrer un code Visual Basic 6 vers VB.NET : le changement de langage, de runtime et de modèle objet — sans casser le comportement existant**  
**Juin 2026** · .NET Framework 4.7.2 · Visual Studio 2026

---

## 🧭 À lire avant de commencer — la philosophie de cette formation

Migrer du **Visual Basic 6** vers **VB.NET**, ce n'est **pas une simple conversion de syntaxe**.  
C'est un changement de **trois choses à la fois** :

1. **Le langage** (mots-clés, structures, types) ;
2. **Le runtime** (le moteur VB6 « VBVM » → le **CLR** de .NET, avec un *garbage collector*) ;
3. **Le modèle objet** (instanciation, durée de vie, finalisation, événements).

> ⚠️ **Le vrai danger n'est pas le code qui ne compile pas — c'est le code qui compile mais
> ne se comporte plus pareil.** ByRef qui devient ByVal, un `Integer` qui change de taille, un
> `Class_Terminate` qui ne se déclenche plus au bon moment : ces **changements silencieux** sont
> la première cause de régressions. Cette formation en fait une priorité (voir **Annexe B**).

### Pourquoi cibler .NET Framework 4.7.2 (et pas .NET moderne) ?

Parce que c'est le **« pont »** naturel : la cible qui **minimise le nombre de changements simultanés**.

- ✅ Surface d'API la plus **proche de VB6** (Windows, COM, GDI+, impression).
- ✅ Meilleure **compatibilité COM / ActiveX / OCX** et **Crystal Reports**.
- ✅ Parité du **concepteur Windows Forms** et des contrôles intrinsèques.
- ✅ **Runtime supporté** (lié au cycle de vie de Windows), adapté à du *legacy* d'entreprise.

> Le saut **4.7.2 → .NET moderne (8/10)** est une **migration distincte**, à faire **après**, et
> seulement si nécessaire. Faire les deux en même temps multiplie les risques. Le module 20 fait
> la passerelle vers la formation **« VB.NET avec .NET 10 LTS »**.

### Ce que cette formation est — et n'est pas

- ✅ Un **guide de migration pas à pas** : langage, formulaires, données, COM, API, validation.
- ✅ Un **catalogue des pièges** VB6 → VB.NET, avec corrections.
- ✅ Une approche **outillée ET manuelle** (assistants, outils commerciaux, IA), avec lucidité sur les limites de chacun.
- ❌ Ce n'est **pas** une formation VB.NET généraliste (pour cela, voir la formation .NET 10 LTS).
- ❌ Ce n'est **pas** une réécriture *from scratch* : on **migre** un existant, on ne le redessine pas.

> Repère honnête : aucun outil ne migre **tout** automatiquement. Les assistants font 60-90 % du
> travail mécanique ; le **jugement humain** (et l'IA, bien encadrée) traite le reste — surtout les
> pièges silencieux et l'UI.

---

## 🎯 Parcours de formation

| 👤 Profil | 📚 Modules recommandés | ⏱️ Durée estimée |
|-----------|------------------------|------------------|
| **Décideur / Cadrage** (évaluer, chiffrer, décider) | 1-4, 20 | 2-3 jours |
| **Migration du langage** (le cœur) ⭐ | 1-2, 5-12, Annexes A/B/D | 8-10 jours |
| **Migration de l'interface** (Forms) | 1-2, 13-14, Annexe C | 5-6 jours |
| **Migration données & COM** | 1-2, 15-16, Annexe A | 5-6 jours |
| **Chef de projet migration** | 1-4, 17-18, 20 | 4-5 jours |
| **Migration assistée par IA** 🤖 | 1-2, 4, 19 | 3-4 jours |
| **Formation complète** | 1-20 + annexes | 22-28 jours |

---

## **Partie 1 — Comprendre et cadrer la migration**

### 1. [Introduction : pourquoi et comment migrer](01-cadrage-migration/README.md) ⭐
- 1.1 [Pourquoi migrer VB6 en 2026 (IDE hors support, runtime, dette technique, recrutement, sécurité)](01-cadrage-migration/01-pourquoi-migrer.md)
- 1.2 [Pourquoi cibler .NET Framework 4.7.2 : le « pont » naturel](01-cadrage-migration/02-pourquoi-472.md) ⭐
- 1.3 [VB6 vs VB.NET : ce qui change vraiment (langage, runtime, modèle objet, GC)](01-cadrage-migration/03-ce-qui-change.md)
- 1.4 [Panorama complet d'une migration : langage, formulaires, données, COM, API](01-cadrage-migration/04-panorama.md)
- 1.5 [Mythes et réalités (« l'assistant fait tout », « ce n'est que de la syntaxe »)](01-cadrage-migration/05-mythes-realites.md) ⚠️

### 2. [Différences fondamentales VB6 ↔ VB.NET](02-differences-fondamentales/README.md) ⭐ ⚠️
- 2.1 [Deux runtimes, deux mondes (VBVM vs CLR)](02-differences-fondamentales/01-vbvm-vs-clr.md)
- 2.2 [La finalisation déterministe perdue — **le piège n°1**](02-differences-fondamentales/02-finalisation-deterministe.md) ⭐ ⚠️
    - VB6 : `Class_Terminate` se déclenche quand le compteur de références tombe à 0
    - .NET : le *garbage collector* décide **quand** ; pas de destruction déterministe → `IDisposable`/`Using`
- 2.3 [Le typage : `Variant`, entiers redimensionnés, `Currency`](02-differences-fondamentales/03-typage.md) ⚠️
- 2.4 [**ByRef par défaut → ByVal par défaut** (changements silencieux d'appel)](02-differences-fondamentales/04-byref-byval.md) ⭐ ⚠️
- 2.5 [Propriétés par défaut et mot-clé `Set` (disparu)](02-differences-fondamentales/05-default-properties-set.md) ⚠️
- 2.6 [Gestion d'erreurs : `On Error` vs exceptions structurées](02-differences-fondamentales/06-erreurs-apercu.md)
- 2.7 [Tableau de synthèse des incompatibilités majeures](02-differences-fondamentales/07-synthese.md)

### 3. [Évaluer l'existant et choisir une stratégie](03-evaluer-strategie/README.md) ⭐
- 3.1 [Inventaire (formulaires `.frm`, modules `.bas`, classes `.cls`, références, OCX, DLL, API)](03-evaluer-strategie/01-inventaire.md)
- 3.2 [Cartographie des dépendances (COM, ActiveX, bases de données, composants tiers)](03-evaluer-strategie/02-dependances.md)
- 3.3 [Estimer l'effort (taille, complexité cyclomatique, points durs)](03-evaluer-strategie/03-estimer-effort.md)
- 3.4 [Stratégies : *big-bang* vs incrémentale](03-evaluer-strategie/04-strategies.md)
- 3.5 [L'approche par **interopérabilité COM** : faire cohabiter VB6 et .NET pendant la transition](03-evaluer-strategie/05-cohabitation-com.md) 🔗
- 3.6 [Définir le périmètre, les critères de réussite et les indicateurs](03-evaluer-strategie/06-perimetre-criteres.md)

### 4. [Outils de migration](04-outils-migration/README.md) 🛠️
- 4.1 [L'assistant de mise à niveau Visual Studio (*Upgrade Wizard*) : capacités et **limites**](04-outils-migration/01-upgrade-wizard.md) ⚠️
- 4.2 [L'espace de noms `Microsoft.VisualBasic.Compatibility` : la **béquille à éviter** (et pourquoi)](04-outils-migration/02-compatibility-namespace.md) ⚠️
- 4.3 [Outils commerciaux : **VBUC** (Mobilize.Net) et **VB Migration Partner** (Code Architects)](04-outils-migration/03-outils-commerciaux.md)
    - Conversion vers VB.NET (ou C#), suppression des dépendances de compatibilité, gestion des *control arrays*
- 4.4 [Migration manuelle assistée : quand et pourquoi](04-outils-migration/04-migration-manuelle.md)
- 4.5 [Grille de décision : outil automatique vs manuel vs hybride](04-outils-migration/05-grille-decision.md)
- 4.6 [Migration assistée par IA → voir module 19](19-migration-ia/README.md) 🤖

---

## **Partie 2 — Préparer le terrain (avant de migrer)**

### 5. [Préparer le code VB6 avant la migration](05-preparer-code-vb6/README.md) ⭐
- 5.1 [Nettoyer (code mort, `Option Explicit` partout, supprimer le `Variant` superflu)](05-preparer-code-vb6/01-nettoyer.md)
- 5.2 [Réduire les pièges **en amont** (typage explicite, `ByVal` explicite côté VB6)](05-preparer-code-vb6/02-reduire-pieges-amont.md) ⭐
- 5.3 [Découpler l'UI de la logique métier (faciliter la migration progressive)](05-preparer-code-vb6/03-decoupler-ui-metier.md)
- 5.4 [Documenter les comportements critiques (arrondis, dates, ordre d'événements)](05-preparer-code-vb6/04-documenter-comportements.md)
- 5.5 [Mettre en place un **harnais de tests de référence** (*golden master*)](05-preparer-code-vb6/05-golden-master.md) ⭐
- 5.6 [Geler le périmètre (arrêter les évolutions fonctionnelles pendant la migration)](05-preparer-code-vb6/06-geler-perimetre.md)

### 6. [Préparer l'environnement .NET](06-preparer-environnement/README.md)
- 6.1 [Visual Studio et le ciblage **.NET Framework 4.7.2**](06-preparer-environnement/01-vs-ciblage-472.md)
- 6.2 [Structure de solution et de projets (`.frm`/`.bas`/`.cls` → fichiers `.vb`)](06-preparer-environnement/02-structure-projets.md)
- 6.3 [Référencer les composants COM/ActiveX nécessaires (RCW, *Primary Interop Assemblies*)](06-preparer-environnement/03-references-com.md) 🔗
- 6.4 [Les options de compilation : `Option Strict` / `Explicit` / `Infer` / `Compare`](06-preparer-environnement/04-options-compilation.md)
    - ⚠️ Démarrer en `Option Strict Off` (toléré pour le *late binding* hérité) puis l'activer **progressivement** (module 17)
- 6.5 [Stratégie de gestion de versions (Git, branche de migration, jalons)](06-preparer-environnement/05-gestion-versions.md)

---

## **Partie 3 — Migrer le langage (le cœur de la formation)**

### 7. [Types, variables et déclarations](07-types-variables/README.md) ⭐ ⚠️
- 7.1 [`Variant` → `Object` : conversions, boxing, `CType`/`DirectCast`](07-types-variables/01-variant-object.md)
- 7.2 [**Entiers redimensionnés** : `Integer`→`Short`, `Long`→`Integer`, et `Currency`→`Decimal`](07-types-variables/02-entiers-redimensionnes.md) ⭐ ⚠️
- 7.3 [`Date` (un `Double` en VB6 → `DateTime` en .NET) : conversions et pièges](07-types-variables/03-dates.md) ⚠️
- 7.4 [Chaînes de **longueur fixe** (`Dim s As String * 10`) : `VBFixedString` ou refactoring](07-types-variables/04-chaines-longueur-fixe.md) ⚠️
- 7.5 [Déclarations : `Dim` multiple, `DefInt`/`DefStr`…, portée, `As New` (instanciation implicite)](07-types-variables/05-declarations.md)
- 7.6 [Constantes, énumérations, et le booléen (`True` = -1 en VB6 ?)](07-types-variables/06-constantes-enums-bool.md) ⚠️
- 7.7 [`Nothing`, `Empty`, `Null`, `Missing` : la fin du `Variant` et de ses états](07-types-variables/07-nothing-empty-null.md) ⚠️

### 8. [Tableaux et collections](08-tableaux-collections/README.md) ⚠️
- 8.1 [Tableaux **toujours 0-based** : `Option Base`, bornes non nulles, `LBound`/`UBound`](08-tableaux-collections/01-tableaux-0-based.md) ⭐ ⚠️
- 8.2 [`ReDim`, `ReDim Preserve`, tableaux dynamiques](08-tableaux-collections/02-redim.md)
- 8.3 [Tableaux de `Variant` / d'objets](08-tableaux-collections/03-tableaux-variant.md)
- 8.4 [L'objet `Collection` VB6 → collections .NET (`List(Of T)`, `Dictionary(Of K,V)`)](08-tableaux-collections/04-collection-vers-net.md)
- 8.5 [`For Each` et énumération (différences de comportement)](08-tableaux-collections/05-foreach.md)

### 9. [Opérateurs, chaînes et structures de contrôle](09-operateurs-controle/README.md)
- 9.1 [Opérateurs (`&` vs `+`, `\` division entière, `Mod`, `Like`, `Is`/`IsNot`)](09-operateurs-controle/01-operateurs.md)
- 9.2 [Fonctions de chaînes (`Left`/`Right`/`Mid`, **instruction** `Mid`, `LSet`/`RSet`) : `Microsoft.VisualBasic` vs `String`](09-operateurs-controle/02-fonctions-chaines.md) ⚠️
- 9.3 [`If…Then`, `Select Case`, `IIf` → opérateur `If()` ternaire (et l'évaluation court-circuit)](09-operateurs-controle/03-conditions.md) ⚠️
- 9.4 [Boucles (`For…Next`, `For Each`, `Do…Loop`, `While…Wend` → `While…End While`)](09-operateurs-controle/04-boucles.md)
- 9.5 [`GoSub`/`Return`, `On…GoTo`/`On…GoSub`, numéros de ligne : **supprimés** → refactoring](09-operateurs-controle/05-gosub-ongoto-supprimes.md) ⚠️
- 9.6 [Instructions multiples, séparateur « `:` », étiquettes de ligne](09-operateurs-controle/06-instructions-etiquettes.md)

### 10. [Procédures, fonctions et paramètres](10-procedures-fonctions/README.md) ⭐ ⚠️
- 10.1 [`Sub`/`Function` ; appels avec/sans parenthèses, `Call`](10-procedures-fonctions/01-sub-function-appels.md)
- 10.2 [**ByRef par défaut → ByVal par défaut** : auditer chaque appel](10-procedures-fonctions/02-byref-byval-audit.md) ⭐ ⚠️
- 10.3 [`Optional` et **valeurs par défaut obligatoires** ; fin de `IsMissing`](10-procedures-fonctions/03-optional-ismissing.md) ⚠️
- 10.4 [`ParamArray` (passé ByRef en VB6 → tableau ByVal en .NET)](10-procedures-fonctions/04-paramarray.md) ⚠️
- 10.5 [Surcharge (nouveauté .NET) et valeurs de retour (`Return` vs `NomFonction = valeur`)](10-procedures-fonctions/05-surcharge-retour.md)
- 10.6 [`Property Get`/`Let`/`Set` → propriétés (`Let` **et** `Set` fusionnés en un unique accesseur `Set` — un faux-ami du `Set` VB6)](10-procedures-fonctions/06-property-get-let-set.md) ⚠️

### 11. [Gestion des erreurs — de `On Error` aux exceptions](11-gestion-erreurs/README.md) ⭐
- 11.1 [`On Error GoTo`, `Resume`, `Resume Next` : équivalents et logique de remplacement](11-gestion-erreurs/01-on-error.md)
- 11.2 [Convertir vers `Try`/`Catch`/`Finally` (et les filtres `When`)](11-gestion-erreurs/02-vers-try-catch.md)
- 11.3 [Le cas épineux de `On Error Resume Next` : pourquoi c'est dangereux à migrer](11-gestion-erreurs/03-on-error-resume-next.md) ⭐ ⚠️
- 11.4 [L'objet `Err` vs `Exception` ; `Err.Raise` → `Throw`](11-gestion-erreurs/04-err-vs-exception.md)
- 11.5 [Stratégie de migration **progressive** (cohabitation `On Error` / `Try` temporaire)](11-gestion-erreurs/05-migration-progressive.md)

### 12. [Programmation orientée objet](12-poo/README.md) ⭐ ⚠️
- 12.1 [Classes (`.cls`) → classes `.vb` : champs, propriétés, méthodes](12-poo/01-classes.md)
- 12.2 [`Class_Initialize` → constructeur ; `Class_Terminate` → **le problème de la finalisation**](12-poo/02-initialize-terminate.md) ⭐ ⚠️
- 12.3 [Finalisation déterministe perdue → `IDisposable` / `Using` / `Finalize` (le pattern Dispose)](12-poo/03-idisposable-using.md) ⭐
- 12.4 [Modules standards (`.bas`) → `Module` ; variables globales et leur portée](12-poo/04-modules-globales.md)
- 12.5 [Interfaces : `Implements` (différences) et **héritage** (nouveauté .NET)](12-poo/05-interfaces-heritage.md)
- 12.6 [`Public`/`Private`/`Friend` : `Friend` = portée **assembly** en .NET](12-poo/06-portee.md) ⚠️
- 12.7 [`Type…End Type` → `Structure` (sémantique de valeur, attention au comportement)](12-poo/07-type-vers-structure.md) ⚠️
- 12.8 [Événements : `Event`/`RaiseEvent`/`WithEvents`/`Handles` ; `AddHandler`/`RemoveHandler`](12-poo/08-evenements.md)

---

## **Partie 4 — Migrer l'interface utilisateur (les formulaires)**

### 13. [Des formulaires VB6 à Windows Forms](13-formulaires-winforms/README.md) ⭐ ⚠️
- 13.1 [Le modèle de formulaire VB6 (`.frm`/`.frx`) → Windows Forms (`.vb` + Designer)](13-formulaires-winforms/01-modele-formulaire.md)
- 13.2 [Cycle de vie : `Load`/`Unload`/`Activate` → `Load`/`FormClosing`/`Shown` ; `Show` vs `ShowDialog`](13-formulaires-winforms/02-cycle-de-vie.md) ⚠️
- 13.3 [Formulaires **MDI** : `MDIForm`/`MDIChild` → `IsMdiContainer`/`MdiParent`, `LayoutMdi` et fusion de menus](13-formulaires-winforms/03-formulaires-mdi.md) ⚠️
- 13.4 [Propriétés, ancrage et positionnement : **Twips → pixels** (conversion d'unités)](13-formulaires-winforms/04-twips-pixels-ancrage.md) ⚠️
- 13.5 [Correspondance des contrôles intrinsèques (`CommandButton`→`Button`, `TextBox`, `Label`, `Frame`, `OptionButton`…)](13-formulaires-winforms/05-correspondance-controles.md)
- 13.6 [Les **tableaux de contrôles** (*control arrays*) : **disparus** → solutions de remplacement](13-formulaires-winforms/06-control-arrays.md) ⭐ ⚠️
    - Recréer un tableau à la main, partage de gestionnaire d'événements, ou bibliothèque dédiée (outils tiers)
- 13.7 [Menus (Menu Editor → `MenuStrip`) et barres d'outils (`ToolStrip`)](13-formulaires-winforms/07-menus-toolbars.md)
- 13.8 [Boîtes de dialogue communes (`CommonDialog` → `OpenFileDialog`, `ColorDialog`…)](13-formulaires-winforms/08-dialogues-communs.md)
- 13.9 [`MsgBox`/`InputBox`, `DoEvents` (toujours là, à manier avec précaution)](13-formulaires-winforms/09-msgbox-doevents.md) ⚠️

### 14. [Graphismes, impression et contrôles ActiveX](14-graphismes-impression-activex/README.md) ⚠️
- 14.1 [Dessin VB6 (`Line`, `Circle`, `PSet`, `Print` sur formulaire) → **GDI+** / objet `Graphics`](14-graphismes-impression-activex/01-dessin-gdiplus.md) ⭐ ⚠️
- 14.2 [Système de coordonnées et `ScaleMode` : repenser le rendu](14-graphismes-impression-activex/02-coordonnees-scalemode.md)
- 14.3 [Images, icônes et ressources (`.frx`, `.res` → `.resx`)](14-graphismes-impression-activex/03-ressources.md)
- 14.4 [Impression : objet `Printer` → `System.Drawing.Printing` (`PrintDocument`)](14-graphismes-impression-activex/04-impression.md) ⚠️
- 14.5 [Contrôles **ActiveX/OCX** : réutiliser via interop ou remplacer](14-graphismes-impression-activex/05-activex-ocx.md) 🔗
- 14.6 [Contrôles tiers courants (grilles, calendriers…) : stratégies de remplacement](14-graphismes-impression-activex/06-controles-tiers.md)

---

## **Partie 5 — Migrer les données et l'interopérabilité**

### 15. [Accès aux données — d'ADO/DAO/RDO à ADO.NET](15-acces-donnees/README.md) ⭐
- 15.1 [Panorama : DAO, RDO, ADO (VB6) vs **ADO.NET**](15-acces-donnees/01-panorama.md)
- 15.2 [Connexions et **commandes paramétrées** (et la sécurité au passage)](15-acces-donnees/02-connexions-commandes.md)
- 15.3 [`Recordset` → `DataReader` / `DataSet` / `DataTable` (curseurs vs déconnecté)](15-acces-donnees/03-recordset-vers-adonet.md) ⚠️
- 15.4 [Liaison de données (`Data` control / *data binding* VB6 → `BindingSource`)](15-acces-donnees/04-data-binding.md)
- 15.5 [Bases Access (`.mdb`/`.accdb`) et SQL Server : chaînes de connexion et adaptations](15-acces-donnees/05-access-sqlserver.md)
- 15.6 [Transactions (de l'`ADO` au `DbTransaction`)](15-acces-donnees/06-transactions.md)
- 15.7 [Reporting : **Crystal Reports** (compatibilité) et alternatives](15-acces-donnees/07-reporting.md) ⚠️

### 16. [Interopérabilité COM et appels d'API Windows](16-interop-com-api/README.md) 🔗 ⭐
- 16.1 [Réutiliser des composants COM/DLL existants depuis .NET (**RCW**)](16-interop-com-api/01-rcw.md)
- 16.2 [Exposer du .NET à du VB6 (**CCW**) : la clé d'une migration **incrémentale**](16-interop-com-api/02-ccw.md) 🔗
- 16.3 [*Early binding* vs *late binding* (`Option Strict Off`) : où c'est nécessaire, où l'éliminer](16-interop-com-api/03-early-late-binding.md)
- 16.4 [Déclarations d'API Win32 : `Declare` → `P/Invoke` (`DllImport`)](16-interop-com-api/04-declare-vers-pinvoke.md) ⚠️
- 16.5 [Marshaling, **ANSI/Unicode**, types et structures (`Type` → `Structure` avec `StructLayout`)](16-interop-com-api/05-marshaling.md) ⚠️
- 16.6 [L'objet `App`, le registre, et `FileSystemObject` → `System.IO` / espace `My`](16-interop-com-api/06-app-registre-fso.md)

---

## **Partie 6 — Finaliser, valider et moderniser**

### 17. [Valider, fiabiliser et refactoriser après migration](17-valider-refactoriser/README.md) ⭐
- 17.1 [Activer `Option Strict` **progressivement** (éliminer le *late binding* résiduel)](17-valider-refactoriser/01-option-strict.md) ⭐
- 17.2 [Supprimer les dépendances à `Microsoft.VisualBasic.Compatibility`](17-valider-refactoriser/02-supprimer-compatibility.md) ⚠️
- 17.3 [Tests de **non-régression** : comparer avec le *golden master*](17-valider-refactoriser/03-non-regression.md) ⭐
- 17.4 [Traquer les **pièges silencieux** (ByRef, arrondis, dates, bornes de tableaux, finalisation)](17-valider-refactoriser/04-traquer-pieges.md) ⭐ ⚠️
- 17.5 [Refactoring idiomatique : LINQ, génériques, espace `My`, code VB.NET moderne](17-valider-refactoriser/05-refactoring-idiomatique.md)
- 17.6 [Revue de performance et de mémoire (`IDisposable`, `Using`, fuites de *handles*)](17-valider-refactoriser/06-performance-memoire.md)

### 18. [Déploiement et bascule en production](18-deploiement-bascule/README.md)
- 18.1 [Packaging .NET Framework (installeur, prérequis 4.7.2, redistribuables)](18-deploiement-bascule/01-packaging.md)
- 18.2 [Cohabitation VB6 / .NET pendant la transition (déploiement progressif)](18-deploiement-bascule/02-cohabitation.md) 🔗
- 18.3 [Stratégie de bascule : pilote, *parallel run*, plan de *rollback*](18-deploiement-bascule/03-bascule-rollback.md)
- 18.4 [Documentation et transfert de compétences à l'équipe](18-deploiement-bascule/04-documentation-transfert.md)

---

## **Partie 7 — IA, suite de parcours et ressources**

### 19. [Migrer avec l'aide de l'IA](19-migration-ia/README.md) 🤖 ⭐
- 19.1 [Pourquoi l'IA aide pour VB6 → VB.NET (et ses **limites**)](19-migration-ia/01-pourquoi-ia.md)
- 19.2 [Prompting : convertir un module/une classe, **expliquer** un comportement VB6 obscur](19-migration-ia/02-prompting.md)
    - ⚠️ Toujours préciser « **VB6** » vs « **VB.NET** » et la cible **.NET Framework 4.7.2** pour éviter les confusions
- 19.3 [Détecter les pièges avec l'IA (ByRef, `Variant`, finalisation, entiers redimensionnés)](19-migration-ia/03-detecter-pieges.md)
- 19.4 [Générer tests de non-régression et documentation XML](19-migration-ia/04-tests-documentation.md)
- 19.5 [Pièges de l'IA : hallucinations, **faux équivalents** VB6/VB.NET, validation systématique](19-migration-ia/05-limites-pieges.md) ⚠️

### 20. [Et après ? De 4.7.2 vers .NET moderne (optionnel)](20-apres-net-moderne/README.md) ⭐
- 20.1 [Pourquoi 4.7.2 d'abord, .NET moderne ensuite : **deux sauts séparés**](20-apres-net-moderne/01-deux-sauts.md) ⭐
- 20.2 [Évaluer un second saut vers .NET 8/10 (intérêt, coût, risques)](20-apres-net-moderne/02-evaluer-second-saut.md)
- 20.3 [Ce qui devra **encore** changer (APIs retirées, WinForms .NET, espace `My` partiel…)](20-apres-net-moderne/03-ce-qui-change-encore.md)
- 20.4 [Quand **rester** sur 4.7.2 (et c'est un choix légitime)](20-apres-net-moderne/04-rester-sur-472.md)
- 20.5 [Passerelle vers la formation « **VB.NET avec .NET 10 LTS** »](20-apres-net-moderne/05-passerelle-net10.md) 🔗

---

## **Annexes**

### A. [Tableau de correspondance VB6 → VB.NET](annexes/correspondance-vb6-vbnet/README.md) ⭐
Aide-mémoire de conversion : mots-clés, fonctions intrinsèques, instructions, opérateurs — la  
référence à garder ouverte pendant toute la migration.

### B. [**Catalogue des pièges silencieux** — le « top des bugs de migration »](annexes/pieges-silencieux/README.md) ⭐ ⚠️
L'annexe la plus importante. Pour chaque piège : **symptôme**, **cause**, **correction**.
- B.1 ByRef devenu ByVal (et l'inverse)
- B.2 `Integer`/`Long`/`Currency` redimensionnés et débordements
- B.3 Finalisation : `Class_Terminate` qui ne se déclenche plus
- B.4 Tableaux 0-based et `Option Base`
- B.5 `Variant` : `Empty`/`Null`/`Missing` perdus
- B.6 `On Error Resume Next` masquant des bugs
- B.7 Propriétés par défaut et `Set` disparu
- B.8 `True = -1` et conversions booléennes
- B.9 Dates (`Double` → `DateTime`) et arrondis monétaires
- B.10 Conversions sensibles à la culture (`CDbl`/`CDate`/`Format` : séparateur décimal et format de date selon les paramètres régionaux)

### C. [Correspondance des contrôles VB6 → Windows Forms](annexes/correspondance-controles/README.md)
Tableau contrôle par contrôle, avec propriétés/événements équivalents et points d'attention
(*control arrays*, contrôles intrinsèques sans équivalent direct).

### D. [Correspondance des types de données](annexes/correspondance-types/README.md)
`Variant`, `Integer`/`Long`, `Currency`, `Date`, `String * n`, `Boolean`… avec tailles, plages et
règles de conversion (`CType`, `DirectCast`, fonctions `CInt`/`CLng`/`CDec`).

### E. [Checklist de migration (avant / pendant / après)](annexes/checklist-migration/README.md)
Liste de vérification opérationnelle, étape par étape, pour ne rien oublier sur un projet réel.

### F. [Modèles de tests de non-régression (*golden master*)](annexes/golden-master/README.md)
Patrons pour capturer le comportement de référence de l'application VB6 et le comparer  
automatiquement après migration.

### G. [Glossaire et acronymes](annexes/glossaire/README.md)
Terminologie VB6, VB.NET, COM, CLR, *marshaling*, RCW/CCW, et vocabulaire de migration.

### H. [Ressources et outils](annexes/ressources/README.md)
Documentation officielle, support du runtime VB6 sur Windows, outils de migration, communauté.

---

## ✅ Ce que cette formation garantit

- **Réaliste** : on migre un existant qui doit continuer à tourner, pas une réécriture idéale.
- **Centrée sur les risques** : les changements **silencieux** (ByRef, types, finalisation) sont la priorité.
- **Outillée et lucide** : assistants, outils commerciaux et IA — avec leurs limites honnêtes.
- **Progressive** : cohabitation COM, activation graduelle d'`Option Strict`, tests de non-régression.
- **Bien cadrée** : 4.7.2 comme « pont », et le saut vers .NET moderne traité comme une étape **séparée** et optionnelle.

**Durée estimée** : ~30-40 h de lecture ; 22-28 jours pour le parcours complet en pratiquant  
**Public** : développeurs et équipes en charge d'un *legacy* VB6, chefs de projet de migration  
**Prérequis** : connaître VB6 (le code source à migrer) ; des bases en .NET aident mais ne sont pas indispensables

---

## 🏷️ Légende des indicateurs

- ⭐ **Point clé** : étape ou notion centrale de la migration
- ⚠️ **Piège** : source fréquente de régressions ou de bugs silencieux
- 🔗 **Interop** : cohabitation / interopérabilité VB6 ↔ .NET (COM)
- 🛠️ **Outils** : assistants et outils de migration
- 🤖 **IA** : migration assistée par intelligence artificielle

---

**Juin 2026**  
**Cible** : .NET Framework 4.7.2 (le « pont » depuis VB6)  
**Outils** : Visual Studio 2026 · assistants de migration · IA  
**Suite optionnelle** : migration 4.7.2 → .NET 10 LTS (formation dédiée)  
**Licence** : Creative Commons BY-NC-SA 4.0
