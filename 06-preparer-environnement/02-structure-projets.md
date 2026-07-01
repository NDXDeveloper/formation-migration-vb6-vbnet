🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.2 Structure de solution et de projets (`.frm`/`.bas`/`.cls` → fichiers `.vb`)

> *Partie 2 — Préparer le terrain · Module 6 — Préparer l'environnement .NET*

---

## Objectif de cette section

Une fois Visual Studio installé et 4.7.2 ciblé (§ 6.1), il faut **organiser l'intérieur** de la zone d'atterrissage : comment se structurent la solution et les projets .NET, et surtout **comment chaque fichier VB6 (`.frm`, `.bas`, `.cls`…) devient un ou plusieurs fichiers `.vb`**. Comprendre cette correspondance est essentiel : c'est elle que les outils de migration appliquent automatiquement, et c'est elle que vous devrez **relire et corriger** en connaissance de cause.

Le principe directeur reste celui du module : **refléter d'abord la structure VB6**, moderniser ensuite (§ 17.5). On ne réorganise pas l'architecture en même temps qu'on migre — ce serait rouvrir le périmètre qu'on vient de geler (§ 5.6).

---

## Deux modèles de projet

VB6 et .NET nomment et organisent les choses différemment. La correspondance de haut niveau :

| VB6 | VB.NET | Note |
|---|---|---|
| Groupe de projets (`.vbg`) | **Solution** (`.sln` / `.slnx`) | Le conteneur qui regroupe plusieurs projets |
| Projet (`.vbp`) | **Projet** (`.vbproj`) | L'unité de compilation (produit un `.exe` ou une `.dll`) |
| **Standard EXE** | **Windows Forms App** (.NET Framework) | → `.exe` |
| **ActiveX DLL** | **Class Library** (.NET Framework) | → `.dll` (composant *in-process*) |
| **ActiveX EXE** | Class Library | L'aspect « hors-processus » est à **repenser** |
| **ActiveX Control** (OCX) | **Windows Forms Control Library** (`UserControl`) | Contrôle visuel réutilisable |
| **ActiveX Document** | *(pas d'équivalent direct)* | À repenser entièrement |

> 💡 Un **groupe de projets VB6** (`.vbg`) devient une **solution** (`.sln`/`.slnx`) qui contient autant de **projets** (`.vbproj`) que le groupe en comptait — en **conservant les dépendances** entre eux. On garde la même topologie.

---

## La correspondance des fichiers VB6 → `.vb`

C'est le cœur de la section (et de son titre). Voici comment chaque type de fichier VB6 se transpose :

| Fichier VB6 | Ce qu'il contient | Devient en .NET |
|---|---|---|
| **`.frm`** | Formulaire : **disposition + code** dans un seul fichier | `Form1.vb` (**code**) **+** `Form1.Designer.vb` (**disposition**) **+** `Form1.resx` (ressources) |
| **`.frx`** | Ressources **binaires** du formulaire (images, icônes) | Fusionné dans le `.resx` / ressources embarquées |
| **`.bas`** | Module standard | `Xxx.vb` contenant `Module … End Module` |
| **`.cls`** | Classe | `Xxx.vb` contenant `Class … End Class` |
| **`.ctl`** | UserControl (source d'OCX) | `Xxx.vb` (UserControl) **+** `Xxx.Designer.vb` **+** `.resx` |
| **`.ctx`** | Ressources binaires du `.ctl` | Fusionné dans le `.resx` |
| **`.pag` / `.pgx`** | Property Page | Souvent **à repenser** (pas d'équivalent direct) |
| **`.dsr`** | Designer (DataEnvironment, DHTML…) | **Pas d'équivalent direct** → refactoring |
| **`.dob` / `.dox`** | ActiveX Document | **Pas d'équivalent direct** → à repenser |
| **`.res`** | Fichier de ressources du projet | `.resx` |

Trois familles se dégagent : celles qui se transposent **proprement** (`.bas`, `.cls`), celles qui se **réorganisent** (`.frm`/`.ctl` → fichiers multiples), et celles qui **n'ont pas d'équivalent** et demandent une refonte (`.dsr`, `.pag`, `.dob`).

---

## Le cas des formulaires : la classe partielle

C'est la transformation la plus importante à comprendre, parce qu'elle change la **façon de lire et d'éditer** un formulaire.

**En VB6**, un `.frm` mélange tout : en tête du fichier, une description sérialisée de la disposition (taille, position et propriétés des contrôles) ; en dessous, votre code. Un seul fichier pour deux préoccupations.

**En VB.NET**, un formulaire est **scindé en deux fichiers** reliés par le mécanisme de **classe partielle** (`Partial Class`) : le compilateur fusionne les deux moitiés en une seule classe.

```vbnet
' Form1.vb — VOTRE code (événements, logique métier de l'écran)
Partial Public Class Form1
    Inherits System.Windows.Forms.Form

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' votre code
    End Sub
End Class
```

```vbnet
' Form1.Designer.vb — GÉNÉRÉ par le concepteur (à ne PAS éditer à la main)
Partial Class Form1
    Private components As System.ComponentModel.IContainer

    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        ' ... position, taille et propriétés des contrôles ...
    End Sub

    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
```

À cela s'ajoute un **`Form1.resx`** qui héberge les ressources du formulaire (chaînes localisables, images embarquées) — c'est là qu'atterrit le contenu de l'ancien `.frx`.

> ⚠️ **Le fichier `.Designer.vb` appartient au concepteur.** On ne l'édite pas à la main : il est **régénéré** quand on manipule le formulaire dans l'éditeur visuel. Votre code, lui, vit dans `Form1.vb`. Confondre les deux (et bricoler le Designer) est une source d'ennuis classique en post-migration.

---

## Modules et classes : une correspondance plus directe

### `.bas` → `Module`

VB.NET **conserve** le mot-clé `Module`. Un `.bas` se transpose donc presque directement :

```vbnet
' Utilitaires.vb  (ex-Utilitaires.bas)
Module Utilitaires
    Public Function Arrondir(montant As Decimal) As Decimal
        ' ...
    End Function
End Module
```

Les membres d'un `Module` restent accessibles globalement dans l'espace de noms du projet — comportement proche de celui des modules VB6.

### `.cls` → `Class`

Un `.cls` devient une `Class`. Mais plusieurs notions VB6 doivent être retraitées :

- La propriété **`Instancing`** du module de classe VB6 (Private, PublicNotCreatable, MultiUse…) gouvernait sa **créabilité COM**. En .NET, cela se traduit par des **modificateurs d'accès** (et des attributs COM seulement si l'on **expose** la classe en COM — § 6.3).
- **`Class_Initialize`** → le **constructeur** `Sub New()`.
- **`Class_Terminate`** → **ne se transpose pas tel quel** : c'est **le piège n°1** de la migration (finalisation non déterministe). Le *garbage collector* décide **quand** détruire l'objet ; on passe donc à `IDisposable`/`Using`. Voir **§ 2.2** — à traiter avec le plus grand soin.

> 💡 **Convention : un type par fichier.** En .NET, on nomme le fichier d'après le type qu'il contient (`Client.vb` → `Class Client`). Les outils de migration produisent généralement **un `.vb` par fichier VB6 d'origine**, ce qui facilite la relecture côte à côte.

---

## Combien de projets ? Refléter d'abord la structure VB6

La tentation de « profiter de la migration pour mieux découper » est forte. **On y résiste** au moment de migrer.

- **Par défaut : un projet .NET par projet VB6.** Un `.vbp` → un `.vbproj`, un `.vbg` → une `.sln`, en **conservant les dépendances**. C'est le choix le plus **fidèle** et celui qui **minimise les changements simultanés**.
- **Ne fusionnez pas, ne redécoupez pas** la structure pendant la migration : ce serait modifier le comportement et l'architecture en même temps que le langage. Le refactoring d'architecture, c'est **après** validation (§ 17.5).
- En migration **incrémentale** avec cohabitation COM (§ 3.5), certains composants restent VB6 pendant que d'autres passent en .NET : la topologie de projets doit **accueillir cette transition**, pas la précéder.

---

## Espaces de noms : du monde « plat » de VB6 aux *namespaces*

VB6 n'a **pas de namespaces** : tout est global à l'intérieur d'un projet (et le **nom du projet** servait de bibliothèque/typelib pour les composants ActiveX, le ProgID d'une classe étant `NomProjet.NomClasse`).

VB.NET a des espaces de noms, pilotés par une propriété **« Espace de noms racine »** (*Root namespace*), qui vaut par défaut **le nom du projet**. Tous les types y vivent implicitement, sauf blocs `Namespace` explicites.

> 💡 **Fidélité d'abord.** Laissez l'espace de noms racine **égal au nom du projet** et n'introduisez pas une arborescence de namespaces sophistiquée pendant la migration. C'est un excellent chantier… **pour après** (§ 17.5). En migrant, on calque ; on ne redessine pas la structure logique.

---

## Organisation en dossiers

VB6 n'avait pas de vrais dossiers (une liste à plat, avec d'éventuels « groupes » d'IDE qui n'étaient pas des répertoires). Les projets .NET, eux, supportent de **vrais dossiers**.

Une **organisation légère** en dossiers (par exemple regrouper les formulaires, les modules utilitaires, les classes métier) est acceptable et améliore la lisibilité. Mais comme pour le reste : **on reste sobre**. Une réorganisation profonde en dossiers/namespaces relève de la modernisation post-migration, pas de la migration elle-même.

---

## Le point d'entrée : *Startup Object* et le framework d'application

VB6 désignait un **objet de démarrage** (« Startup Object ») : soit un **formulaire**, soit `Sub Main`. Cela se retrouve en VB.NET, avec une subtilité propre à Windows Forms :

- Le projet WinForms .NET Framework propose un **framework d'application** (case « Activer le framework d'application ») qui apporte `My.Application`, l'instance unique éventuelle, l'écran de démarrage, le mode d'arrêt, etc. On y choisit le **formulaire de démarrage**.
- Si on **désactive** ce framework, on fournit alors un **`Sub Main`** comme objet de démarrage — proche du modèle VB6 « Sub Main ».

Le choix se fait dans *Propriétés du projet → onglet « Application »*. À ce stade, il suffit de **reproduire le comportement de démarrage** de l'application VB6 d'origine.

---

## Cas particuliers qui ne se mappent pas 1:1

Quelques constructions VB6 n'ont **pas** d'équivalent direct et demandent une attention spécifique :

- **Tableaux de contrôles** (propriété `Index` sur un contrôle, gestionnaires d'événements partagés via `Index As Integer`) : **pas de notion native** en Windows Forms. Les outils commerciaux (VBUC, VB Migration Partner) fournissent un *helper* de compatibilité, ou l'on refactore vers une vraie collection de contrôles. Voir **Annexe C**.
- **DataEnvironment / `.dsr`**, **Property Pages / `.pag`**, **ActiveX Documents / `.dob`** : pas d'équivalent direct → refonte ciblée, à recenser dès l'inventaire.
- **Ressources** (`.res`, `.frx`, `.ctx`, `.pgx`) : regroupées dans des **`.resx`** managés.

> 🔗 Ces points durs doivent figurer dans l'**inventaire (§ 3.1)** et la **cartographie des dépendances (§ 3.2)** : ce sont eux qui font grimper l'effort, pas les `.bas`/`.cls` qui se transposent proprement.

---

## Ce que les outils font (et pourquoi comprendre quand même)

Les assistants et outils de migration (module 4) **génèrent automatiquement** cette structure : création de la `.sln` et des `.vbproj`, transformation de chaque `.frm` en `Form.vb` + `Designer.vb` + `.resx`, de chaque `.bas` en `Module`, de chaque `.cls` en `Class`.

Comprendre la correspondance ne dispense donc pas de l'automatisation — elle vous permet de **relire, valider et corriger** sa sortie : repérer un `.Designer.vb` malmené, un tableau de contrôles mal converti, un `Class_Terminate` transposé sans `IDisposable`. **L'outil produit la structure ; vous en jugez la justesse.**

---

## Points de vigilance

> ⚠️ **Ne pas éditer le `.Designer.vb` à la main.** Il est régénéré par le concepteur ; votre code va dans le fichier `Xxx.vb` jumeau.

> ⚠️ **`Class_Terminate` ≠ destructeur déterministe.** La transposition d'une classe doit traiter la finalisation (`IDisposable`/`Using`), pas la recopier naïvement (§ 2.2 — le piège n°1).

> 💡 **Refléter avant de réorganiser.** Un `.vbproj` par `.vbp`, espace de noms racine = nom du projet, organisation sobre. Le redécoupage en projets/namespaces/dossiers, c'est le module 17.5.

> 💡 **Recenser les fichiers « sans équivalent ».** `.dsr`, `.pag`, `.dob` et tableaux de contrôles sont les vrais points durs : à identifier tôt (inventaire § 3.1), pas à découvrir en cours de route.

---

## À retenir

> - **Correspondance des conteneurs** : groupe `.vbg` → **solution** `.sln`/`.slnx` ; projet `.vbp` → **projet** `.vbproj` (Standard EXE → WinForms App, ActiveX DLL → Class Library, OCX → Control Library).
> - **Correspondance des fichiers** : `.frm` → `Form.vb` **+** `Form.Designer.vb` **+** `.resx` ; `.bas` → `Module` ; `.cls` → `Class` ; ressources (`.frx`/`.ctx`/`.res`) → `.resx`.
> - **Le formulaire devient une classe partielle** : votre code dans `Xxx.vb`, la disposition (générée) dans `Xxx.Designer.vb` — qu'on **n'édite pas à la main**.
> - **`.bas`/`.cls` se transposent proprement** ; le point sensible d'une classe est la **finalisation** (`Class_Terminate` → `IDisposable`, § 2.2).
> - **Fidélité d'abord** : un projet par projet, espace de noms racine = nom du projet, organisation sobre ; la réorganisation est post-migration (§ 17.5).
> - **Les vrais points durs** (sans équivalent 1:1) : tableaux de contrôles, `.dsr`/`.pag`/`.dob` — à recenser dès l'inventaire (§ 3.1).

---

## Liens utiles

- **§ 6.1** — Visual Studio et le ciblage 4.7.2 *(l'étape précédente : l'IDE et le framework)*
- **§ 6.3** — Référencer les composants COM/ActiveX 🔗 *(les références que la structure doit accueillir)*
- **§ 3.1 / 3.2** — Inventaire et cartographie des dépendances *(où recenser les fichiers et points durs)*
- **§ 2.2** — La finalisation déterministe perdue *(le traitement de `Class_Terminate`)*
- **§ 5.6** — Geler le périmètre *(pourquoi on ne réorganise pas en migrant)*
- **§ 17.5** — Refactoring idiomatique *(la réorganisation projets/namespaces/dossiers, après validation)*
- **Annexe C** — Correspondance des contrôles VB6 → Windows Forms *(dont les tableaux de contrôles)*

⏭️ [Référencer les composants COM/ActiveX nécessaires (RCW, *Primary Interop Assemblies*)](/06-preparer-environnement/03-references-com.md)
