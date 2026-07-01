🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.3 — Formulaires MDI : `MDIForm`/`MDIChild` → `IsMdiContainer`/`MdiParent`, `LayoutMdi` et fusion de menus ⚠️

> **Chapitre 13 — Des formulaires VB6 à Windows Forms** · Section 13.3
> Le MDI cesse d'être un **type de formulaire spécial** pour devenir une **propriété** — et la **fusion de menus** devient plus explicite.

---

## 🧭 Périmètre

Le **MDI** (interface à documents multiples) — un conteneur abritant plusieurs fenêtres enfants — existe dans les deux mondes, mais son **modèle** change. VB6 a un **type** `MDIForm` dédié et une **propriété** `MDIChild` ; Windows Forms n'a **ni l'un ni l'autre** : tout repose sur les propriétés **`IsMdiContainer`** et **`MdiParent`**.

---

## 1. Le modèle MDI, de VB6 à Windows Forms

En **VB6** :
- le conteneur est un **`MDIForm`** — un **type de formulaire spécial**, **au plus un par projet** ;
- une fenêtre enfant est un **formulaire ordinaire** avec **`MDIChild = True`**.

En **Windows Forms** :
- le conteneur est un **`Form` ordinaire** dont on met **`IsMdiContainer = True`** ;
- une fenêtre devient enfant en affectant sa propriété **`MdiParent`** au conteneur, **en code**.

| VB6 | Windows Forms |
|-----|---------------|
| `MDIForm` (type spécial, ≤ 1 par projet) | `Form` ordinaire **+ `IsMdiContainer = True`** |
| `MDIChild = True` (propriété, design) | `child.MdiParent = conteneur` (**code**, avant `Show`) |
| `MDIForm.Arrange vbCascade/…` | `conteneur.LayoutMdi(MdiLayout.Cascade/…)` |
| `MDIForm.ActiveForm` | `conteneur.ActiveMdiChild` |
| *(pas de collection directe)* | `conteneur.MdiChildren` (tableau des enfants) |
| `WindowList = True` (menu) | `MenuStrip.MdiWindowListItem` |
| fusion **automatique** des menus enfant/parent | fusion **explicite** (`MergeAction`/`MergeIndex`/`AllowMerge`) |

> ⚠️ **Changement conceptuel** : le MDI n'est plus un *type*, c'est une *propriété*. Et le lien parent-enfant, qui était une propriété **de conception** en VB6 (`MDIChild`), se fixe désormais **à l'exécution** (`MdiParent`). L'assistant gère la bascule, mais cette différence explique plusieurs pièges ci-dessous.

---

## 2. Créer et gérer les enfants

En VB6, un formulaire `MDIChild = True` apparaît **automatiquement** dans le `MDIForm` dès qu'on l'affiche :

```vb
' VB6
Dim f As New frmDocument        ' frmDocument.MDIChild = True
f.Show                          ' apparaît dans le MDIForm
```

En Windows Forms, on **affecte `MdiParent`** — impérativement **avant `Show`** :

```vb
' VB.NET — instance EXPLICITE par document
Dim f As New frmDocument()
f.MdiParent = Me                ' AVANT Show, sinon fenêtre indépendante
f.Show()
```

> ⚠️ **Le piège de l'instance par défaut.** Utiliser l'instance par défaut (`frmDocument.Show()`, voir **[13.1](01-modele-formulaire.md)**) ne donne qu'**un seul** enfant **partagé** — incompatible avec le MDI, où l'on veut **plusieurs** documents ouverts. Pour un MDI, **instanciez explicitement** (`New frmDocument()`) à **chaque** nouvelle fenêtre.

```vb
' ⚠️ À ÉVITER en MDI : un seul document, réutilisé
frmDocument.MdiParent = Me
frmDocument.Show()              ' toujours la même instance

' ✅ Correct : une nouvelle instance par document
Dim doc As New frmDocument()
doc.MdiParent = Me
doc.Show()
```

L'enfant actif et la liste des enfants se lisent via **`ActiveMdiChild`** et **`MdiChildren`** :

```vb
' VB.NET
Dim actif = Me.ActiveMdiChild
For Each enfant In Me.MdiChildren
    ' …
Next
```

---

## 3. Disposer les fenêtres : `Arrange` → `LayoutMdi`

La méthode `Arrange` du `MDIForm` devient **`LayoutMdi`** sur le conteneur :

| VB6 (`Arrange`) | Windows Forms (`LayoutMdi`) |
|-----------------|------------------------------|
| `vbCascade` | `MdiLayout.Cascade` |
| `vbTileHorizontal` | `MdiLayout.TileHorizontal` |
| `vbTileVertical` | `MdiLayout.TileVertical` |
| `vbArrangeIcons` | `MdiLayout.ArrangeIcons` |

```vb
' VB.NET
Me.LayoutMdi(MdiLayout.Cascade)
```

---

## 4. ⚠️ La fusion de menus (le point le plus délicat)

En **VB6**, la fusion est **automatique** : quand un enfant a son propre menu, celui-ci **fusionne** avec (ou **remplace**) le menu du `MDIForm` selon des règles de **négociation** intégrées. Le menu « Fenêtre » avec **`WindowList = True`** liste automatiquement les enfants ouverts.

En **Windows Forms**, tout passe par **`MenuStrip`** et devient **plus explicite** :

- la **liste des fenêtres** s'obtient en désignant l'élément de menu concerné via **`MdiWindowListItem`** :

```vb
' VB.NET — menu "Fenêtre" qui liste les enfants ouverts
menuStripPrincipal.MdiWindowListItem = mnuFenetre   ' équivalent de WindowList = True
```

- la **fusion** d'un menu d'enfant dans celui du parent se **configure** via les propriétés **`AllowMerge`** (sur le `MenuStrip`), **`MergeAction`** et **`MergeIndex`** (sur chaque `ToolStripMenuItem`) — bien plus **manuel** que la négociation automatique de VB6.

> ⚠️ **Deux stratégies de migration.**
> - **Reproduire** les menus par enfant **avec fusion** : fidèle, mais **laborieux** à régler (chaque élément doit indiquer son `MergeAction`/`MergeIndex`).
> - **Consolider** vers **un seul menu** sur le conteneur, activé/désactivé selon le contexte : **plus simple**, souvent **préférable** pour une première migration.
>
> Le détail des menus (`Menu Editor` → `MenuStrip`) est en **[13.7](07-menus-toolbars.md)**.

---

## 5. Contrôles sur le conteneur

En VB6, on ne pouvait poser sur un `MDIForm` que des contrôles dotés d'une propriété **`Align`** (ex. `PictureBox`, barres d'outils) ; les autres devaient aller sur les enfants.

Windows Forms est **plus permissif**, mais la pratique reste d'**ancrer** (`Dock`) les accessoires sur les bords : **`MenuStrip`**, **`ToolStrip`**, **`StatusStrip`**. La **zone client MDI** est elle-même un contrôle (`MdiClient`) ; les contrôles ajoutés cohabitent avec elle — à positionner avec soin pour ne pas la masquer.

---

## 🔁 Récapitulatif VB6 → VB.NET

| VB6 | Windows Forms |
|-----|---------------|
| `MDIForm` | `Form` + `IsMdiContainer = True` |
| `MDIChild = True` | `child.MdiParent = conteneur` (avant `Show`) |
| `New enfant : enfant.Show` | `New enfant() : enfant.MdiParent = Me : enfant.Show()` |
| instance par défaut pour un enfant | ⚠️ **`New` explicite** par document |
| `MDIForm.Arrange vb…` | `LayoutMdi(MdiLayout.…)` |
| `MDIForm.ActiveForm` | `ActiveMdiChild` |
| `WindowList = True` | `MenuStrip.MdiWindowListItem` |
| fusion automatique des menus | fusion **explicite** (`AllowMerge`/`MergeAction`/`MergeIndex`) ou menu unique |
| contrôles `Align` sur le `MDIForm` | contrôles **`Dock`** (`MenuStrip`/`ToolStrip`/`StatusStrip`) |

---

## ✅ Points clés

- Le MDI devient une **propriété** : conteneur via **`IsMdiContainer`**, enfant via **`MdiParent`** (fixé **en code, avant `Show`**).
- ⚠️ Pour ouvrir **plusieurs** documents, **instanciez explicitement** chaque enfant (`New`) : l'**instance par défaut** n'en donne **qu'un**.
- `Arrange` → **`LayoutMdi`** ; `ActiveForm` → **`ActiveMdiChild`** ; ajout de **`MdiChildren`**.
- ⚠️ La **fusion de menus** passe d'**automatique** à **explicite** (`MergeAction`/`MergeIndex`) ; envisagez un **menu unique** pour simplifier. `WindowList` → **`MdiWindowListItem`**.
- Sur le conteneur, **ancrez** (`Dock`) les barres ; la zone client MDI est un contrôle à part.

---

## 🔗 Renvois

- **[13.1](01-modele-formulaire.md)** — instance par défaut (le piège des enfants MDI). · **[13.2](02-cycle-de-vie.md)** — cycle de vie.
- **[13.7](07-menus-toolbars.md)** — menus (`MenuStrip`) et barres d'outils (`ToolStrip`) : le détail de la fusion.
- **Suite → [13.4 — Propriétés, ancrage et positionnement : Twips → pixels (conversion d'unités)](04-twips-pixels-ancrage.md)** ⚠️

⏭️ [Propriétés, ancrage et positionnement : **Twips → pixels** (conversion d'unités)](/13-formulaires-winforms/04-twips-pixels-ancrage.md)
