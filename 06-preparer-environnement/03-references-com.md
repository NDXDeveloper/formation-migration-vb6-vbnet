🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.3 Référencer les composants COM/ActiveX nécessaires (RCW, *Primary Interop Assemblies*) 🔗

> *Partie 2 — Préparer le terrain · Module 6 — Préparer l'environnement .NET*

---

## Objectif de cette section

Une application VB6 vit rarement seule : elle s'appuie sur des **contrôles ActiveX (OCX)**, des **DLL COM** (composants métier, accès aux données, *FileSystemObject*…) et des **bibliothèques de types**. Pour que le code migré continue de fonctionner, il faut **rétablir ces dépendances** côté .NET.

La bonne nouvelle, et c'est un pilier du choix de 4.7.2 : .NET Framework sait **dialoguer avec COM** via l'**interopérabilité**. On peut donc, dans l'esprit du « pont », **garder les composants COM existants au début** (effort minimal, comportement préservé) et ne les **remplacer par des équivalents natifs que plus tard**, sélectivement (§ 17.5). Un seul changement à la fois.

Cette section explique le mécanisme (RCW, assemblies d'interop, PIA), comment ajouter les références, et les **pièges** qui font perdre le plus de temps — la **bitness** en tête.

---

## Le point de départ : l'inventaire des dépendances COM

Tout part de l'**inventaire (§ 3.1)** et de la **cartographie des dépendances (§ 3.2)** : la liste des OCX, DLL COM, références et bibliothèques de types qu'utilise l'application. Pour **chacune**, une décision :

| | Réutiliser via **interop** (COM) | Remplacer par un **équivalent natif** .NET |
|---|---|---|
| Effort immédiat | **Faible** | Plus élevé |
| Fidélité comportementale | **Élevée** (même composant) | Risque de changement de comportement/apparence |
| Dépendance COM | **Conservée** (enregistrement, bitness…) | **Supprimée** |
| Quand le faire ? | **Pendant** la migration (le « pont ») | **Après** validation, au cas par cas (§ 17.5) |
| Exemple | Garder ADODB, un *grid* tiers | ADODB → ADO.NET, MSFlexGrid → DataGridView |

> 💡 **Par défaut pour la migration : réutiliser via interop.** On rétablit les dépendances telles quelles pour faire tourner l'application vite, on valide avec le *golden master*, **puis** on remplace ce qui mérite de l'être. Tout remplacer d'un coup, c'est cumuler migration du langage **et** changement de composants — exactement ce qu'on évite.

---

## Comment .NET parle à COM : RCW, CCW, assemblies d'interop

Quelques notions à poser, car elles reviennent partout :

| Terme | Rôle |
|---|---|
| **RCW** — *Runtime Callable Wrapper* | Un **proxy côté .NET** qui enveloppe un objet COM pour que le code managé l'appelle comme un objet .NET. C'est le sens **.NET → COM** (consommer du COM). |
| **CCW** — *COM Callable Wrapper* | Le **proxy inverse** : il permet à un client COM d'appeler un objet .NET. Sens **COM → .NET** (exposer du .NET). Utile en cohabitation (voir plus bas). |
| **Assembly d'interop** | Un assembly **managé** contenant les **métadonnées** miroir d'une bibliothèque de types COM. C'est **ce contre quoi votre code compile** ; à l'exécution, le CLR crée les RCW. |
| **PIA** — *Primary Interop Assembly* | L'assembly d'interop **officielle**, **signée par l'éditeur** et partagée (souvent dans le GAC). Elle garantit que **tout le monde utilise les mêmes types managés** pour une bibliothèque COM donnée. |

En résumé : on ajoute une **référence** à une bibliothèque COM → Visual Studio fournit un **assembly d'interop** (généré, ou la **PIA** si elle existe) → votre code compile contre lui → à l'exécution, le CLR fabrique des **RCW** pour parler aux vrais objets COM.

---

## PIA ou assembly d'interop généré ?

Deux situations :

- **Une PIA existe** (cas d'Office, d'ADODB, de nombreux composants Microsoft) → **utilisez-la**. Elle assure l'**identité de type** : tous les projets qui s'en servent partagent exactement les mêmes types managés, ce qui évite des erreurs d'incompatibilité.
- **Pas de PIA** (beaucoup de composants tiers ou maison) → Visual Studio **génère** un assembly d'interop pour vous au moment où vous ajoutez la référence (l'outil sous-jacent est `tlbimp.exe`, l'*importateur de bibliothèque de types*).

> ⚠️ **Ne mélangez pas, pour une même bibliothèque COM, une PIA et un interop généré localement** : vous obtiendrez des **types différents** vus comme incompatibles. Une bibliothèque = **une** source d'interop, partagée.

---

## Le cas des contrôles ActiveX (OCX) : **deux** assemblies

Un contrôle **visuel** (OCX) à poser sur un formulaire est un cas particulier : Visual Studio génère **deux** assemblies, pas une.

- **`Interop.XXX.dll`** — le **RCW** des interfaces/coclasses COM sous-jacentes du contrôle.
- **`AxInterop.XXX.dll`** — l'**enveloppe `AxHost`** (dérivée de `System.Windows.Forms.AxHost`) qui **héberge** le contrôle ActiveX à l'intérieur d'un conteneur Windows Forms.

Concrètement, le contrôle apparaît dans votre formulaire sous la forme d'une classe **`AxXXX`** (préfixe `Ax`). L'outil en ligne de commande correspondant est `aximp.exe` (*ActiveX importer*), pendant de `tlbimp.exe` pour les contrôles.

---

## Ajouter les références concrètement

Deux chemins selon la nature du composant :

- **DLL COM / bibliothèque de types** : *Projet → Ajouter une référence → onglet **COM*** → sélectionner la bibliothèque. Visual Studio crée l'assembly d'interop (ou utilise la **PIA** si elle est enregistrée).
- **Contrôle ActiveX (OCX)** : l'ajouter d'abord à la **Boîte à outils** (clic droit → *Choisir les éléments…* → onglet **Composants COM**), puis le **déposer sur un formulaire**. Visual Studio génère alors les références `Interop.XXX` **et** `AxInterop.XXX`.

> 💡 Pour des besoins avancés (signature forte, personnalisation), on peut générer les assemblies d'interop **à la main** avec `tlbimp.exe` / `aximp.exe`, puis y référer. Dans la grande majorité des cas, l'ajout via l'IDE suffit.

---

## « Embed Interop Types » (NoPIA)

Depuis .NET Framework 4 (donc disponible en 4.7.2), chaque référence COM possède une propriété **« Incorporer les types d'interop »** (*Embed Interop Types*). À `True`, les informations de type d'interop sont **embarquées dans votre propre assembly** : vous n'avez **plus à déployer la PIA** séparément (c'est la fonction d'**équivalence de types** / *NoPIA*).

- **Bibliothèques de types COM** : l'incorporation fonctionne et **simplifie le déploiement** (très courant pour Office).
- **Contrôles ActiveX (enveloppes `Ax`)** : l'incorporation **ne s'applique pas** de la même façon ; on **déploie** alors les assemblies `Interop.XXX` et `AxInterop.XXX` avec l'application.

---

## Le piège majeur : la **bitness** 32/64 bits

C'est **l'erreur d'interop la plus fréquente** en migration, et l'une des plus déroutantes.

Les composants COM hérités de VB6 sont **32 bits**. Or un projet .NET compilé en **AnyCPU** s'exécute en **64 bits** sur un Windows 64 bits — et un processus 64 bits **ne peut pas charger** un composant COM *in-process* 32 bits.

> ⚠️ **Symptôme typique :** à l'exécution, une erreur du type « *Retrieving the COM class factory… failed* » avec le code **`0x80040154`** (classe non enregistrée), ou un `BadImageFormatException`. Le composant est pourtant bien là et enregistré : c'est la **bitness** qui cloche.

**La parade :** forcer le projet en **32 bits**.

```xml
<!-- Dans le .vbproj : cibler explicitement x86 -->
<PlatformTarget>x86</PlatformTarget>
```

Cela se règle dans *Propriétés du projet → onglet « Compiler »* en choisissant **x86** comme cible de plateforme (alternative pour un exécutable AnyCPU : cocher **« Préférer 32 bits »**). Tant qu'on consomme du COM 32 bits, **on reste en 32 bits.**

---

## La contrainte d'enregistrement (*registration*)

Pour que Visual Studio trouve un composant COM — et pour que le RCW fonctionne à l'exécution — le composant doit être **enregistré** sur la machine :

- OCX et DLL COM *in-process* : enregistrés via **`regsvr32`** ;
- la **bibliothèque de types** doit être enregistrée également.

> ⚠️ **Cette contrainte vaut aussi pour les machines cibles.** L'interop COM **réintroduit** le besoin d'enregistrement (le fameux « *DLL hell* » que .NET cherchait justement à éviter). Les composants COM devront être présents **et enregistrés** sur les postes des utilisateurs.

L'**activation COM sans enregistrement** (*reg-free COM*, via manifestes côte à côte) est une option **avancée** pour éviter `regsvr32` au déploiement. Le packaging et ces choix sont traités au **§ 18.1**.

---

## Exemples fréquents (et leurs alternatives natives)

Quelques dépendances quasi universelles en VB6, avec le choix « garder via interop » **ou** « remplacer par du natif » :

- **ADODB** (accès aux données via ADO) → garder via sa **PIA**, **ou** migrer vers **ADO.NET**. Le sujet « données » a son propre module : voir **module 15**.
- **Microsoft Windows Common Controls** (`mscomctl.ocx` : TreeView, ListView, Toolbar, StatusBar, ProgressBar…) → garder via interop, **ou** remplacer par les contrôles **natifs** Windows Forms (`TreeView`, `ListView`, `StatusStrip`…).
- **Common Dialog** (`comdlg32.ocx`) → équivalents **natifs** `OpenFileDialog`, `SaveFileDialog`, `ColorDialog`, `FontDialog`, `PrintDialog`.
- **MSFlexGrid** (`msflxgrd.ocx`) → garder via interop, **ou** passer à **`DataGridView`**.
- **FileSystemObject** (*Scripting Runtime*) → garder via interop, **ou** passer à **`System.IO`** (traité aussi en § 16.6).

> 💡 Les **outils commerciaux** (VBUC, VB Migration Partner — module 4) proposent souvent de **remplacer automatiquement** certains contrôles intrinsèques par leurs équivalents natifs. Pratique, mais à **valider** : un remplacement est un **changement de comportement/apparence**, donc à passer au crible du *golden master* (§ 5.5). Voir aussi l'**Annexe C** (correspondance des contrôles).

---

## L'autre sens : exposer du .NET à COM (cohabitation)

En migration **incrémentale** avec cohabitation (§ 3.5, § 18.2), il arrive que du **VB6 doive appeler du .NET** déjà migré. C'est le sens inverse : on **expose** un composant .NET à COM via un **CCW**.

```vbnet
Imports System.Runtime.InteropServices

<ComVisible(True)>
Public Class MonComposant
    Public Function Calculer(montant As Decimal) As Decimal
        ' ...
    End Function
End Class
```

L'assembly est ensuite **enregistré pour COM** (`regasm.exe`, génération d'une bibliothèque de types), afin que le client VB6 le voie comme un composant COM ordinaire. Ce mécanisme est ce qui permet aux **deux mondes de cohabiter** pendant la transition.

---

## Points de vigilance

> ⚠️ **Bitness avant tout.** COM hérité = 32 bits → projet en **x86**. Le `0x80040154` « classe non enregistrée » sur un composant pourtant présent = quasi toujours un problème de bitness.

> ⚠️ **Enregistrement requis, machines cibles comprises.** L'interop COM réintroduit la dépendance à `regsvr32` ; à anticiper pour le déploiement (§ 18.1).

> ⚠️ **Une bibliothèque = une source d'interop.** PIA **si elle existe**, sinon un interop généré, mais pas les deux mélangés (erreurs d'identité de type).

> 💡 **Réutiliser d'abord, remplacer ensuite.** Garder les composants COM via interop pour avancer ; remplacer par du natif **après** validation, au cas par cas (§ 17.5) — chaque remplacement passe par le *golden master*.

---

## À retenir

> - .NET Framework **dialogue avec COM** : c'est ce qui permet de **garder les composants ActiveX/COM existants** pendant la migration (le « pont »), et de les remplacer **plus tard**.
> - **Mécanisme** : on ajoute une **référence** → un **assembly d'interop** (généré, ou la **PIA** si elle existe) → à l'exécution, le CLR crée des **RCW** (proxies .NET → COM).
> - Un **OCX** génère **deux** assemblies : `Interop.XXX` (RCW) **+** `AxInterop.XXX` (enveloppe `AxHost`), le contrôle apparaissant comme `AxXXX`.
> - **PIA** quand elle existe (identité de type) ; **« Embed Interop Types »** simplifie le déploiement des bibliothèques de types (pas des enveloppes `Ax`).
> - **Pièges clés** : la **bitness** (COM 32 bits → projet **x86**, sinon `0x80040154`) et l'**enregistrement** des composants, machines cibles incluses.
> - Sens inverse (cohabitation) : exposer du .NET à COM via un **CCW** (`<ComVisible(True)>`, `regasm`) pour que VB6 appelle du .NET migré.

---

## Liens utiles

- **§ 3.1 / 3.2** — Inventaire et cartographie des dépendances *(la liste des composants à rétablir)*
- **§ 3.5** — L'approche par interopérabilité COM *(faire cohabiter VB6 et .NET)*
- **§ 6.1 / 6.2** — IDE/ciblage et structure de projets *(le contexte de ces références)*
- **§ 5.5** — *Golden master* *(valider chaque remplacement de composant)*
- **§ 15** — Migration des données *(ADODB → ADO.NET)*
- **§ 16.6** — `App`, registre, `FileSystemObject` → `System.IO` *(remplacements natifs)*
- **§ 17.5** — Refactoring idiomatique *(remplacer le COM par du natif, après validation)*
- **§ 18.1 / 18.2** — Packaging et cohabitation *(enregistrement, déploiement, reg-free COM)*
- **Annexe C** — Correspondance des contrôles VB6 → Windows Forms

⏭️ [Les options de compilation : `Option Strict` / `Explicit` / `Infer` / `Compare`](/06-preparer-environnement/04-options-compilation.md)
