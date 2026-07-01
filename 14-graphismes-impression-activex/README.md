🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 🎨 14. Graphismes, impression et contrôles ActiveX

> **Quand le code compile, mais que plus rien ne s'affiche.**
> Le rendu graphique, l'impression et les composants visuels ne changent pas seulement de
> syntaxe : ils changent de **modèle**. C'est l'un des chapitres où l'on rencontre le plus de
> **changements silencieux** — d'où l'indicateur ⚠️.

**Partie 4 — Migrer l'interface utilisateur** · Cible : .NET Framework 4.7.2 · Windows Forms + GDI+

---

## 🧭 Où se situe ce chapitre

Le module 13 a migré **l'enveloppe** des formulaires : la fenêtre, son cycle de vie, les contrôles  
intrinsèques, les menus. Ce module 14 s'occupe de **tout ce qui vit dessus** et que VB6 traitait
« à sa façon » :

- ce que vous **dessinez** sur un formulaire ou une `PictureBox` ;
- les **images, icônes et ressources** embarquées ;
- ce que vous **imprimez** ;
- les contrôles **non intrinsèques** — les fameux **ActiveX/OCX** et composants tiers.

Ce sont quatre domaines où VB6 offrait des raccourcis très pratiques… qui n'ont **pas d'équivalent  
direct** en .NET. La migration y est moins une traduction qu'une **reconception ciblée**.

---

## ⚠️ Le piège central : du dessin *persistant* au dessin *redessiné*

C'est **le** point à comprendre avant tout le reste.

En VB6, vous pouviez dessiner **directement** sur un formulaire ou une `PictureBox` :

```vb
' VB6 — on dessine n'importe où, n'importe quand
Me.Line (0, 0)-(100, 100)
Me.Circle (50, 50), 30
Me.Print "Bonjour"
```

Et avec `AutoRedraw = True`, VB6 **mémorisait** le dessin dans une image interne : il restait  
affiché même après un recouvrement, un redimensionnement ou une réduction de la fenêtre.

En .NET / GDI+, **ce modèle n'existe plus**. L'interface est **redessinée en permanence** par le  
système. Tout ce qui n'est **pas (re)dessiné dans l'événement `Paint`** disparaît au premier  
rafraîchissement venu : une autre fenêtre passe devant, l'utilisateur redimensionne, réduit puis  
restaure…

> ⚠️ **Le bug le plus courant de ce chapitre** : du code qui **compile** et qui semble même
> fonctionner au premier affichage — puis dont le dessin « s'efface » dès que la fenêtre est
> manipulée. Rien n'a planté ; le rendu n'a simplement **pas été reproduit**. La règle .NET est :
> *on décrit comment se dessiner, on ne dessine pas une fois pour toutes.* (Détaillé en **14.1** et
> **14.2**.)

Retenez la bascule : **mode « je dessine et ça reste » → mode « je redessine à chaque demande de  
rendu ».** Elle gouverne tout le reste du chapitre, y compris l'impression.

---

## 🧩 Les autres ruptures de modèle

Au-delà du dessin, trois changements de **modèle** (et non de syntaxe) structurent ce chapitre.

**1. L'impression devient événementielle.**
L'objet `Printer` de VB6 était **séquentiel et à état** : on appelait `Printer.Print`,
`Printer.Line`, `Printer.NewPage`, puis `Printer.EndDoc`. En .NET, on configure un `PrintDocument`
et on **répond à l'événement `PrintPage`**, page par page, en indiquant s'il reste des pages à  
imprimer. Le résultat est plus robuste, mais la logique est **inversée**. (Module **14.4**.)

**2. Les ressources changent de format.**
Les ressources binaires des formulaires (`.frx`) et les fichiers `.res` cèdent la place aux  
fichiers **`.resx`** (XML) et à l'espace `My.Resources`. La migration récupère images et icônes,  
mais la **mécanique d'accès** diffère. (Module **14.3**.)

**3. Les ActiveX/OCX deviennent une décision d'architecture.**
Beaucoup d'applications VB6 reposent lourdement sur des contrôles **ActiveX/OCX** (grilles,  
calendriers, graphiques, éditeurs riches…). Deux voies : les **réutiliser via l'interopérabilité  
COM** (un emballage `AxHost` est généré automatiquement) ou les **remplacer** par des contrôles
.NET natifs. Ce choix se heurte souvent à des contraintes très concrètes — **licences**, composants
**32 bits uniquement**, éditeur disparu — qui en font, dans la pratique, **l'un des vrais points
durs d'une migration**, bien plus que le langage. (Modules **14.5** 🔗 et **14.6**.)

---

## 📋 Vue d'ensemble des correspondances

Un aperçu de haut niveau (chaque ligne est détaillée dans la section indiquée) :

| 🟦 VB6 | 🟩 .NET Framework 4.7.2 | Le vrai changement | Section |
|--------|--------------------------|--------------------|---------|
| `Line`, `Circle`, `PSet`, `Print` sur formulaire | Méthodes de l'objet `Graphics` (`DrawLine`, `DrawEllipse`, `DrawString`…) dans `Paint` | Dessin **persistant** → **redessiné** à chaque rendu | 14.1 ⚠️ |
| `ScaleMode`, twips, `ScaleLeft/Top/Width/Height` | Pixels par défaut + transformations (`PageUnit`, `Transform`) | Système de coordonnées à **repenser** | 14.2 |
| `.frx`, `.res` | `.resx` (XML), `My.Resources` | Format et **accès** aux ressources | 14.3 |
| Objet `Printer` (`Print`/`NewPage`/`EndDoc`) | `System.Drawing.Printing.PrintDocument` + `PrintPage` | Modèle **séquentiel** → **événementiel** | 14.4 ⚠️ |
| Contrôles **ActiveX/OCX** | Interop COM (`AxHost`) **ou** contrôle natif | Réutiliser **ou** remplacer (licence, 32 bits) | 14.5 🔗 |
| Contrôles tiers (grilles, calendriers…) | Équivalents .NET / composants commerciaux | Stratégie de **remplacement** | 14.6 |

---

## 📑 Ce que vous allez voir

- **14.1 — Dessin VB6 → GDI+ / objet `Graphics`** ⭐ ⚠️
  La rupture fondatrice : `Line`/`Circle`/`PSet`/`Print` n'écrivent plus « sur le formulaire » mais
  via un objet `Graphics`, dans l'événement `Paint`. Le cœur du chapitre.
- **14.2 — Système de coordonnées et `ScaleMode`**
  De twips et d'échelles personnalisables vers des pixels et des transformations : comment retrouver
  le bon repère de rendu.
- **14.3 — Images, icônes et ressources (`.frx`, `.res` → `.resx`)**
  Récupérer les ressources embarquées des formulaires et y accéder « à la .NET ».
- **14.4 — Impression : `Printer` → `PrintDocument`** ⚠️
  Passer d'une impression séquentielle à une impression pilotée par événements, page par page.
- **14.5 — Contrôles ActiveX/OCX** 🔗
  Réutiliser via interop COM ou remplacer : critères, contraintes (licences, 32 bits) et pièges.
- **14.6 — Contrôles tiers courants**
  Grilles, calendriers et autres composants commerciaux : stratégies de remplacement réalistes.

---

## ♻️ Un réflexe à acquérir tout de suite : `IDisposable` et GDI+

Les objets GDI+ que vous créez — `Pen`, `Brush`, `Font`, `Bitmap`, et tout `Graphics` que vous  
obtenez vous-même — encapsulent des **ressources natives** (des *handles* GDI). En VB6, on ne s'en  
souciait pas. En .NET, on **doit les libérer**, idéalement avec `Using` :

```vb
Using p As New Pen(Color.Blue, 2)
    e.Graphics.DrawLine(p, 0, 0, 100, 100)
End Using
```

> 🔗 C'est une application directe du **problème de la finalisation déterministe** vu aux modules 2
> et 12 : le *garbage collector* finira par récupérer ces objets, mais **trop tard**. Ne pas les
> libérer provoque des **fuites de handles GDI** — un classique des applications migrées « qui
> ralentissent puis n'arrivent plus à dessiner » après quelques heures. (Voir aussi **Annexe B**,
> piège de la finalisation.)

---

## 🎯 Objectifs du chapitre

À l'issue de ce module, vous saurez :

- ✅ Reproduire un rendu graphique VB6 **correctement**, via l'événement `Paint` et l'objet
  `Graphics`, sans « dessin qui s'efface ».
- ✅ **Convertir les coordonnées** (twips → pixels) et choisir le bon système de rendu.
- ✅ Migrer **images, icônes et ressources** de `.frx`/`.res` vers `.resx`.
- ✅ Reconstruire une **impression** sur le modèle `PrintDocument`/`PrintPage`.
- ✅ **Décider** entre interop COM et remplacement natif pour chaque contrôle ActiveX/OCX, en
  connaissant les contraintes (licence, 32 bits, support).
- ✅ Libérer les objets GDI+ avec `Using` pour éviter les **fuites de handles**.

---

## 🔑 Prérequis et liens utiles

- **Module 12 — POO** (`IDisposable`/`Using`/`Finalize`) : indispensable pour gérer proprement les
  objets GDI+.
- **Module 13 — Formulaires → Windows Forms** : le cadre dans lequel ce dessin et ces contrôles
  s'inscrivent.
- **Module 16 — Interopérabilité COM** (🔗 RCW) : socle technique de la réutilisation des
  ActiveX/OCX (section 14.5).
- **Annexe C — Correspondance des contrôles VB6 → Windows Forms** : à garder ouverte pour 14.5 et
  14.6.
- **Annexe B — Catalogue des pièges silencieux** : finalisation (handles GDI), entiers et
  coordonnées.

---

> **En résumé**, ce chapitre n'est pas une affaire de mots-clés mais de **modèles** : on passe d'un
> dessin et d'une impression « immédiats et persistants » à un rendu **événementiel et
> reproductible**, et l'on transforme la dépendance aux ActiveX en **décision d'architecture**.
> Gardez en tête le piège central — *ce qui n'est pas redessiné dans `Paint` disparaît* — et la
> plupart des régressions de ce chapitre s'évitent d'elles-mêmes.

Commençons par la rupture fondatrice : **14.1 — Dessin VB6 (`Line`, `Circle`, `PSet`, `Print`) →  
GDI+ / objet `Graphics`**.

⏭️ [Dessin VB6 (`Line`, `Circle`, `PSet`, `Print` sur formulaire) → **GDI+** / objet `Graphics`](/14-graphismes-impression-activex/01-dessin-gdiplus.md)
