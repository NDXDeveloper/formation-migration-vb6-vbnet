🔝 Retour au [Sommaire](/SOMMAIRE.md)

# G. Glossaire et acronymes

Terminologie **VB6**, **VB.NET**, **COM**, **CLR**, *marshaling*, **RCW/CCW**, et **vocabulaire de  
migration**.

> 🧭 **Comment l'utiliser.** Un **tableau d'acronymes** (§1) pour décoder rapidement les sigles, puis
> des **définitions regroupées par thème** (§2 à §7), en ordre alphabétique à l'intérieur de chaque
> section. Quelques renvois (annexe / module) pointent vers le développement complet.

---

## 1. Acronymes (référence rapide)

| Sigle | Développé | En bref |
|---|---|---|
| **ADO** | ActiveX Data Objects | Accès aux données COM côté VB6. |
| **ADO.NET** | — | Pile d'accès aux données .NET (modèle **différent** d'ADO). |
| **ANSI** | American National Standards Institute (ici : encodage) | Chaînes mono-octet dépendantes de la page de code (API « A »). |
| **API** | Application Programming Interface | Souvent l'**API Win32** dans ce contexte. |
| **BCL** | Base Class Library | Bibliothèque de base de .NET (espaces `System.*`). |
| **BSTR** | Basic String | Type de chaîne COM (Unicode). |
| **CCW** | COM Callable Wrapper | Proxy exposant un objet **.NET** à un appelant **COM**. |
| **CIL / MSIL / IL** | (Microsoft) Common Intermediate Language | Langage intermédiaire produit par le compilateur .NET. |
| **CLR** | Common Language Runtime | Moteur d'exécution de .NET (remplace le VBVM). |
| **CLS** | Common Language Specification | Sous-ensemble du CTS pour l'interop entre langages. |
| **COM** | Component Object Model | Standard de composants binaires ; **fondation de VB6**. |
| **CTS** | Common Type System | Modèle de types commun aux langages .NET. |
| **DAO** | Data Access Objects | Accès aux données (Jet) sous VB6. |
| **DCOM** | Distributed COM | COM réparti sur le réseau. |
| **DLL** | Dynamic-Link Library | Bibliothèque liée dynamiquement. |
| **FSO** | FileSystemObject | API fichiers du *Scripting Runtime* (VB6). |
| **GAC** | Global Assembly Cache | Cache d'assemblys partagé de la machine. |
| **GC** | Garbage Collector | Ramasse-miettes (gestion **non déterministe** de la mémoire). |
| **GDI / GDI+** | Graphics Device Interface (+) | Couche de dessin Windows (.NET : `System.Drawing`). |
| **GUID** | Globally Unique Identifier | Identifiant unique (classes/interfaces COM, etc.). |
| **IDE** | Integrated Development Environment | Environnement de développement (VB6 IDE, Visual Studio). |
| **JIT** | Just-In-Time | Compilation IL → natif à l'exécution. |
| **LINQ** | Language-Integrated Query | Requêtes intégrées au langage (.NET). |
| **OCX** | OLE Control Extension | Fichier de contrôle **ActiveX**. |
| **OLE** | Object Linking and Embedding | Technologie d'incorporation (ancêtre, `OLE_COLOR`, conteneur OLE). |
| **P/Invoke** | Platform Invocation Services | Appel de fonctions de DLL natives depuis .NET. |
| **PIA** | Primary Interop Assembly | Assembly d'interop **officielle** d'une bibliothèque COM. |
| **RCW** | Runtime Callable Wrapper | Proxy permettant d'appeler un objet **COM** depuis **.NET**. |
| **RDO** | Remote Data Objects | Accès aux données distantes (VB6). |
| **STA / MTA** | Single/Multi-Threaded Apartment | Modèles de cantonnement de threads COM (STA = VB6/WinForms). |
| **TLB** | Type Library | Description des types d'un composant COM. |
| **VB6** | Visual Basic 6.0 | Le langage/IDE **source** de la migration. |
| **VB.NET** | Visual Basic .NET | Le langage **cible**. |
| **VBVM** | Visual Basic Virtual Machine | Le **runtime VB6** (`MSVBVM60.DLL`). |
| **WinForms** | Windows Forms | Framework d'IHM .NET (le plus proche des formulaires VB6). |
| **WPF** | Windows Presentation Foundation | Autre IHM .NET (hors champ de cette formation). |

---

## 2. Runtime, CLR et exécution .NET

- **Assembly** — Unité de déploiement .NET (`.dll`/`.exe`) contenant code IL, métadonnées et
  manifeste.
- **BCL (Base Class Library)** — Bibliothèque de base de .NET (espaces `System.*`).
- **Boxing / unboxing** — Encapsulation d'un **type valeur** dans un `Object` (sur le tas) / son
  extraction. Le déboxage par `DirectCast` exige le **type exact** (Annexe D §4.7).
- **CLR (Common Language Runtime)** — Moteur d'exécution de .NET : gestion mémoire (**GC**),
  compilation **JIT**, sécurité, types. **Remplace le VBVM** de VB6.
- **CLS (Common Language Specification)** — Règles d'interopérabilité entre langages .NET. Exemple :
  les **types non signés** (`UInteger`…) ne sont **pas conformes CLS**.
- **CTS (Common Type System)** — Système de types partagé par tous les langages .NET.
- **Espace de noms (*namespace*)** — Regroupement logique de types (ex. `System.IO`).
- **Finalisation déterministe** — Destruction à un **moment prévisible** (le `Class_Terminate` de
  VB6). **Perdue** en .NET → motif `IDisposable` (Annexe B.3).
- **Finaliseur / `Finalize`** — « Destructeur » .NET exécuté par le **GC**, à un moment **non
  déterministe**.
- **Generics (génériques)** — Types paramétrés (`List(Of T)`, `Dictionary(Of K, V)`).
- **Garbage Collector (GC)** — Récupération **automatique et non déterministe** de la mémoire.
  S'oppose au **comptage de références**.
- **IDisposable / `Dispose` / `Using`** — Motif de **libération déterministe** des ressources
  (remplace `Class_Terminate`).
- **IL / MSIL / CIL** — Langage intermédiaire émis par le compilateur, compilé en natif par le JIT.
- **JIT (Just-In-Time)** — Compilation de l'IL en code natif **à l'exécution**.
- **Manifeste** — Métadonnées décrivant une assembly (version, dépendances…).
- **Métadonnées / Réflexion** — Informations de type embarquées dans l'assembly / leur inspection à
  l'exécution.
- **Managed / unmanaged (géré / non géré)** — Code exécuté **sous le CLR** / code **natif** hors CLR
  (Win32, COM).
- **`My` (espace)** — Raccourcis VB.NET (`My.Application`, `My.Computer`, `My.Settings`…).
- **.NET Framework 4.7.2** — Runtime **cible** de cette formation, le **« pont »** depuis VB6 :
  Windows, forte compatibilité COM/WinForms (module 1.2).
- **`Nullable(Of T)`** — Type valeur pouvant aussi valoir `Nothing` (ex. `Integer?`), avec
  `.HasValue`/`.Value`.
- **Type valeur / type référence** — Structure ou primitif (**copié**) / classe (**référencée**).
- **Élargissement / restriction (*widening* / *narrowing*)** — Conversion **sûre** / potentiellement
  **avec perte** (Annexe D §4.3).

---

## 3. COM et interopérabilité

- **ActiveX** — Composants/contrôles fondés sur COM (fichiers **OCX**).
- **Apartment / STA / MTA** — Modèles de cantonnement de threads COM. VB6 et le thread d'IHM WinForms
  sont **STA** (`<STAThread>`).
- **BSTR** — Type de chaîne COM (Unicode).
- **CCW (COM Callable Wrapper)** — Proxy généré pour exposer un objet **.NET** à un appelant **COM**.
- **COM (Component Object Model)** — Standard binaire de composants ; **socle de VB6**.
- **DCOM** — COM distribué (appels inter-machines).
- **Early binding (liaison précoce)** — Appels résolus à la **compilation** (type connu) : plus
  rapide, vérifié.
- **IDispatch** — Interface COM de **liaison tardive** (utilisée par le *late binding* `Object`).
- **Interop COM** — Le pont entre le monde **géré** et le monde **COM** (RCW/CCW).
- **Assembly d'interop** — Wrapper généré (via `tlbimp`) à partir d'une **bibliothèque de types
  (TLB)** COM.
- **IUnknown** — Interface COM de base (`QueryInterface`/`AddRef`/`Release`).
- **Late binding (liaison tardive)** — Appels résolus à l'**exécution** (via `IDispatch` /
  `Option Strict Off`). À **éliminer** progressivement (modules 16.3, 17.1).
- **OLE / `OLE_COLOR`** — Technologie d'incorporation ; en VB6, type de **couleur** (`Long`) →
  `System.Drawing.Color` (Annexe C §7).
- **PIA (Primary Interop Assembly)** — Assembly d'interop **officielle et signée** d'une bibliothèque
  COM.
- **RCW (Runtime Callable Wrapper)** — Proxy permettant à du code **.NET** d'appeler un objet
  **COM**.
- **Type Library (TLB)** — Description des types exposés par un composant COM.
- **`VARIANT` / `VARIANT_BOOL`** — Type variant COM (base du `Variant` VB6) ; `VARIANT_BOOL` vaut
  **-1** pour vrai (Annexe B.8).

---

## 4. Marshaling, P/Invoke et interop natif

- **ANSI vs Unicode** — Variantes **« A »** (mono-octet, page de code) vs **« W »** (large) des API
  Win32 ; déterminant pour `Declare`/P/Invoke (Annexe A, module 16.4-16.5).
- **Blittable (type *blittable*)** — Type dont la représentation **mémoire est identique** en géré et
  non géré (ex. `Byte`, `Short`, `Integer`, `Double`) → **pas de conversion**. `Boolean`, `Char`,
  `String` ne sont **pas** blittables.
- **`Declare`** — Instruction VB (VB6 **et** VB.NET) pour appeler une fonction d'une DLL native ;
  **précurseur** de P/Invoke.
- **`DllImport` / `StructLayout` / `MarshalAs`** — Attributs .NET pilotant le P/Invoke et la
  **disposition mémoire** des structures (module 16.5).
- **Handle (`hWnd`, `hDC`)** — Identifiant opaque d'une ressource du système. En .NET :
  `Handle`/`IntPtr` (fenêtre), `Graphics` (contexte de dessin) — Annexe C §1.
- **`IntPtr`** — Entier de la **taille de la plateforme**, pour handles et pointeurs.
- **Marshaling** — **Conversion des données** entre représentations **gérée** et **non gérée** au
  passage d'une frontière (managed ↔ natif/COM).
- **P/Invoke (Platform Invoke)** — Mécanisme .NET d'appel de fonctions de **DLL natives**
  (`<DllImport>`).
- **Convention d'appel** — Protocole de passage des arguments/retours (`StdCall`, `Cdecl`…).

---

## 5. VB6 — langage et environnement

- **`ByRef` / `ByVal`** — Modes de passage des paramètres. ⚠️ **`ByRef` par défaut en VB6**
  (Annexe B.1).
- **`Class_Initialize` / `Class_Terminate`** — Événements de cycle de vie d'un objet VB6
  (construction / destruction **déterministe**).
- **Contrôle de données (`Data`)** — Contrôle VB6 de liaison aux données (DAO/Jet). **Sans
  équivalent** direct (Annexe C §4, module 15).
- **`Currency`** — Type monétaire 64 bits **mis à l'échelle** (4 décimales). → `Decimal`
  (Annexe D §3).
- **Fichiers `.frm` / `.bas` / `.cls`** — Formulaire / module standard / module de classe.
- **FileSystemObject (FSO)** — API fichiers du *Scripting Runtime* → `System.IO` (module 16.6).
- **`GoSub`/`Return`, `While`/`Wend`, `Eqv`/`Imp`** — Constructions **retirées** en VB.NET
  (Annexe A §3-4).
- **`Microsoft.VisualBasic.Compatibility`** — Espace de **compatibilité** émulant des comportements
  VB6 : la **« béquille à éviter »** (module 4.2, à retirer module 17.2).
- **`On Error` / `Resume Next`** — Gestion d'erreurs VB6. ⚠️ `Resume Next` **masque** les erreurs
  (Annexe B.6).
- **`Option Base` / `Option Explicit` / `Option Compare`** — Directives de module. `Option Base`
  **supprimé** (Annexe B.4).
- **`OLE_COLOR`** — Type de couleur VB6 (voir §3).
- **Propriété par défaut** — Propriété implicite (`Text1 = "x"` ⇒ `.Text`). **Supprimée** (sauf
  indexée) — Annexe B.7.
- **`ScaleMode`** — Sélecteur de système de coordonnées VB6. **Disparaît** (Annexe C §1).
- **String de longueur fixe (`String * n`)** — Chaîne à taille fixe. **Pas de type natif** en
  VB.NET (Annexe D §3).
- **Tableau de contrôles (*control array*)** — Contrôles partageant un `Name` et un `Index`.
  **Sans équivalent** WinForms (Annexe C §3).
- **Twip** — Unité de coordonnées par défaut de VB6 (**1/1440 de pouce**). → pixels.
- **Upgrade Wizard (assistant de mise à niveau)** — Outil Visual Studio convertissant VB6 → VB.NET
  (capacités **et limites**, module 4.1).
- **`Variant`** — Type « tout-venant » de VB6 (états `Empty`/`Null`/`Missing`). → `Object`
  (Annexe B.5).
- **VBVM** — Runtime VB6 (`MSVBVM60.DLL`) : comptage de références, fonctions intrinsèques, moteur de
  formulaires.

---

## 6. VB.NET — langage et frameworks

- **`Anchor` / `Dock`** — Propriétés de **disposition automatique** des contrôles (Annexe C §9).
- **`BindingSource` / `DataGridView` / `DataSet` / `DataTable`** — Briques de liaison **ADO.NET**
  (module 15).
- **`Char`** — Type d'un **caractère** Unicode (nouveau ; littéral `"A"c`).
- **`CType` / `DirectCast` / `TryCast`** — Opérateurs de **conversion / transtypage** (Annexe D §4.2).
- **`Decimal`** — Type décimal **exact** 128 bits (montants ; remplace `Currency`).
- **Délégué (*delegate*)** — Référence de fonction **typée** (base des événements ; `AddressOf`).
- **Événements : `Event` / `Handles` / `AddHandler` / `RaiseEvent`** — Modèle d'événements VB.NET
  (gestionnaires **partagés** utiles aux ex-*control arrays*, Annexe C §3).
- **Exception : `Try` / `Catch` / `Finally` / `Throw`** — Gestion d'erreurs **structurée**
  (remplace `On Error`).
- **`MenuStrip` / `ToolStrip` / `StatusStrip` / `ContextMenuStrip`** — Menus et barres modernes
  (Annexe C).
- **`Option Strict` / `Option Infer`** — Options de compilation. **`Option Strict On`** interdit les
  conversions implicites risquées et le *late binding* (modules 6.4, 17.1).
- **`Property` (`Get`/`Set`)** — Syntaxe **unifiée** des propriétés (fusionne `Get`/`Let`/`Set` de
  VB6).
- **`Short` / `Integer` / `Long`** — Entiers **16 / 32 / 64 bits** en VB.NET (⚠️ tailles différentes
  de VB6, Annexe D).
- **`Structure` / `Enum` / `Interface` / `Inherits` / `Overrides`** — Constructions de type et
  **orientées objet** (héritage absent de VB6).
- **Windows Forms (WinForms)** — Framework d'IHM .NET, **cible** de la migration des formulaires
  (modules 13-14).

---

## 7. Vocabulaire de migration

- **Approbation (*approval testing*, *received/approved*)** — Motif outillant le golden master :
  sortie **obtenue** comparée à une référence **approuvée** (Annexe F §6.3).
- **Arrondi au pair (*banker's rounding*)** — Arrondi **« au pair le plus proche »**, utilisé par
  `CInt`/`CLng`/`Math.Round` par défaut (Annexe D §4.4).
- ***Big-bang* vs incrémentale** — Tout migrer d'un coup vs progressivement (module 3.4).
- **Caractérisation (test de)** — Synonyme de **golden master** : figer le comportement actuel pour
  le préserver (Annexe F).
- **Cohabitation COM / pont d'interop** — Faire **coexister** VB6 et .NET pendant la transition
  (modules 3.5, 18.2).
- **Complexité cyclomatique** — Métrique de complexité du code (estimation d'effort, module 3.3).
- **Dette technique** — Coût de maintenance accumulé (un des motifs de migration, module 1.1).
- **Gel du périmètre** — **Arrêter les évolutions** fonctionnelles pendant la migration (module 5.6).
- **Golden master** — Capture du comportement de référence, comparée automatiquement après migration
  (Annexe F).
- **Invariant de culture** — Formatage/analyse **indépendants des paramètres régionaux** (parade au
  piège B.10, Annexe F §3).
- **LTS (Long-Term Support)** — Version à support long (ex. **.NET 10 LTS**, suite optionnelle,
  module 20).
- **Migration vs réécriture** — **Préserver** un existant vs le **redessiner** *from scratch* (cette
  formation **migre**, modules 1.5, intro).
- **Non-régression** — **Prouver** que le comportement n'a pas changé (module 17.3).
- ***Parallel run*** — Exécuter **ancien et nouveau en parallèle** pour comparer (module 18.3).
- **Pilote** — Premier déploiement **restreint** (module 18.3).
- **Pont (« le pont »)** — Le choix de cibler **4.7.2 d'abord** (modules 1.2, 20.1).
- **Refactoring (idiomatique)** — Amélioration de la **structure** du code **après** migration
  (module 17.5).
- **Régression / changement silencieux** — Code qui **compile mais ne se comporte plus pareil** :
  l'ennemi n°1 (Annexe B).
- **Rollback** — **Retour arrière** sur un déploiement (plan de *rollback*, module 18.3).

---

## 🔗 Annexes liées

- **Annexe A** — Correspondance des mots-clés/fonctions/opérateurs.
- **Annexe B** — Catalogue des pièges silencieux (le « pourquoi » de nombreux termes ⚠️).
- **Annexe C** — Correspondance des contrôles. **Annexe D** — Correspondance des types.
- **Annexe F** — Modèles de golden master (caractérisation / approbation / non-régression).

---

**Juin 2026**  
**Cible** : .NET Framework 4.7.2 (le « pont » depuis VB6)  
**Licence** : Creative Commons BY-NC-SA 4.0

⏭️ [Ressources et outils](/annexes/ressources/README.md)
