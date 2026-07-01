🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.7 — Menus (Menu Editor → `MenuStrip`) et barres d'outils (`ToolStrip`)

> **Chapitre 13 — Des formulaires VB6 à Windows Forms** · Section 13.7
> Les **menus** (intrinsèques en VB6) et les **barres d'outils** (un **ActiveX** en VB6) trouvent des équivalents **natifs** : `MenuStrip` et `ToolStrip`.

---

## 🧭 Périmètre

Deux familles, à l'origine **de natures différentes** en VB6 :

- les **menus** sont **intrinsèques** (créés dans le **Menu Editor**) → **`MenuStrip`** ;
- les **barres d'outils** ne le sont **pas** : la `Toolbar` vient de l'**OCX « Microsoft Windows Common Controls »** → on la **remplace** par le **`ToolStrip`** natif.

---

## 1. Les menus : Menu Editor → `MenuStrip`

### 1.1 Le modèle

En **VB6**, les menus se construisent dans le **Menu Editor** : une hiérarchie de **contrôles `Menu`** rattachée à la barre de menus (implicite) du formulaire. Chaque élément a un événement **`Click`**.

En **Windows Forms**, on dépose un **`MenuStrip`** sur le formulaire (ancré en haut) et l'on saisit les éléments **directement en place** ; chaque élément est un **`ToolStripMenuItem`**.

| VB6 (Menu Editor) | Windows Forms |
|-------------------|---------------|
| barre de menus (implicite) | **`MenuStrip`** |
| contrôle `Menu` (élément) | **`ToolStripMenuItem`** |
| `Caption` | `Text` |
| séparateur (`Caption = "-"`) | **`ToolStripSeparator`** |
| `Shortcut` | **`ShortcutKeys`** (énum `Keys`) |
| `Checked` | `Checked` / `CheckOnClick` |
| tableau de menus (`Index`) | éléments **dynamiques** (**[13.6](06-control-arrays.md)**) |
| `WindowList` (MDI) | `MdiWindowListItem` (**[13.3](03-formulaires-mdi.md)**) |
| événement `Click` | événement `Click` |

```vb
' VB6 — éléments créés dans le Menu Editor ; on n'écrit que les Click
'   mnuFichierOuvrir  (Caption "&Ouvrir...", Shortcut Ctrl+O)
'   mnuFichierQuitter (Caption "&Quitter")
Private Sub mnuFichierOuvrir_Click()
    ' …
End Sub
Private Sub mnuFichierQuitter_Click()
    Unload Me
End Sub
```

```vb
' VB.NET — MenuStrip construit dans le concepteur ; on écrit les handlers
Private Sub mnuFichierOuvrir_Click(sender As Object, e As EventArgs) _
        Handles mnuFichierOuvrir.Click
    ' …
End Sub
Private Sub mnuFichierQuitter_Click(sender As Object, e As EventArgs) _
        Handles mnuFichierQuitter.Click
    Me.Close()
End Sub
```

### 1.2 Les détails qui changent

- `Caption` → **`Text`** ; le **`&`** des raccourcis clavier (mnémoniques, ex. `&Fichier` → Alt+F) **fonctionne à l'identique**.
- `Shortcut` → **`ShortcutKeys`** avec l'énum `Keys` (ex. `Keys.Control Or Keys.O`), réglé dans le concepteur ; `ShowShortcutKeys` contrôle l'affichage.
- Séparateur : `Caption = "-"` → **`ToolStripSeparator`** (un type d'élément, plus une convention de texte).
- Coches : `Checked` → `Checked` / `CheckOnClick` (et `CheckState` pour trois états).
- **Images** : un `ToolStripMenuItem` peut porter une **icône** (`Image`) — ce que les menus VB6 ne faisaient pas nativement.

> 💡 Utilisez **`MenuStrip`** (moderne, depuis .NET 2.0), **pas** l'ancien **`MainMenu`** (.NET 1.x). La fusion de menus en MDI (`MergeAction`/`MergeIndex`) est traitée en **[13.3](03-formulaires-mdi.md)**.

### 1.3 Menus contextuels : `PopupMenu` → `ContextMenuStrip`

En VB6, on affichait un menu contextuel via la **méthode `PopupMenu`**, déclenchée au clic droit. En Windows Forms, on **assigne** un **`ContextMenuStrip`** à la propriété `ContextMenuStrip` d'un contrôle : il s'affiche **automatiquement** au clic droit, **sans code de déclenchement**.

```vb
' VB6 — afficher un menu contextuel à la main
Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu mnuContexte
End Sub
```

```vb
' VB.NET — ContextMenuStrip assigné dans le concepteur :
'   listBox1.ContextMenuStrip = cmsContexte
' → s'affiche tout seul au clic droit (aucun code requis)
```

---

## 2. Les barres d'outils : `Toolbar` (OCX) → `ToolStrip`

### 2.1 Un changement de nature

⚠️ La **`Toolbar` de VB6 n'est pas intrinsèque** : elle provient de l'**OCX « Microsoft Windows Common Controls »** (`MSComctl`), tout comme `StatusBar`, `ImageList`, `TreeView`, `ListView`, `ProgressBar`… Windows Forms fournit un **`ToolStrip` natif** : on **remplace** l'OCX, on ne l'**interopère** pas.

> 💡 Bien qu'on **puisse** techniquement réutiliser l'OCX par interop (cas général en **[14.5](../14-graphismes-impression-activex/README.md)**), le **remplacement** par `ToolStrip` est **préférable** : il évite de traîner une **dépendance ActiveX** et donne un rendu/comportement natif.

### 2.2 Le modèle

| VB6 (Common Controls OCX) | Windows Forms (natif) |
|---------------------------|------------------------|
| `Toolbar` | **`ToolStrip`** |
| `Button` (de la `Toolbar`) | **`ToolStripButton`** |
| séparateur, liste déroulante | `ToolStripSeparator`, `ToolStripDropDownButton`, `ToolStripComboBox`… |
| `ImageList` (OCX, requis) | composant **`ImageList`** natif, **ou** `Image` par bouton |
| événement `ButtonClick` (avec `Button`) | **`Click`** par bouton (ou `ItemClicked`) |
| `StatusBar` (OCX) | **`StatusStrip`** (+ `ToolStripStatusLabel`) |

```vb
' VB6 — Toolbar (OCX) : un ButtonClick global, on teste la clé du bouton
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "ouvrir" : OuvrirFichier
        Case "enreg"  : EnregistrerFichier
    End Select
End Sub
```

```vb
' VB.NET — ToolStrip natif : un Click par bouton (plus lisible)
Private Sub tsbOuvrir_Click(sender As Object, e As EventArgs) Handles tsbOuvrir.Click
    OuvrirFichier()
End Sub
Private Sub tsbEnreg_Click(sender As Object, e As EventArgs) Handles tsbEnreg.Click
    EnregistrerFichier()
End Sub
```

Quelques apports du `ToolStrip` :

- chaque `ToolStripButton` règle son rendu via **`DisplayStyle`** (`Image` / `Text` / `ImageAndText`) et porte directement son **`Image`** et sa **`ToolTipText`** ;
- des **types d'éléments riches** (zone de saisie, liste déroulante, étiquette) cohabitent dans la même barre ;
- un **`ToolStripContainer`** permet à l'utilisateur de **déplacer/ancrer** les barres.

### 2.3 Barre d'état : `StatusBar` → `StatusStrip`

Dans la même logique, la `StatusBar` (OCX) est remplacée par le **`StatusStrip`** natif, dont les panneaux sont des **`ToolStripStatusLabel`**.

---

## 🔁 Récapitulatif VB6 → VB.NET

| VB6 | Windows Forms |
|-----|---------------|
| menus (Menu Editor) | `MenuStrip` + `ToolStripMenuItem` |
| `Caption` (menu) | `Text` |
| `Shortcut` | `ShortcutKeys` (`Keys`) |
| `PopupMenu` | `ContextMenuStrip` (assigné, automatique) |
| `Toolbar` (OCX) | `ToolStrip` (natif — **remplacer**) |
| `Button` (Toolbar) | `ToolStripButton` |
| `ImageList` (OCX) | `ImageList` natif / `Image` par bouton |
| `StatusBar` (OCX) | `StatusStrip` |

---

## ✅ Points clés

- Les **menus** passent de **Menu Editor** à **`MenuStrip`** / **`ToolStripMenuItem`** (`Caption`→`Text`, `Shortcut`→`ShortcutKeys`, séparateur → `ToolStripSeparator`).
- Les **menus contextuels** se simplifient : **`PopupMenu` → `ContextMenuStrip`** assigné, qui s'affiche **automatiquement**.
- ⚠️ La **`Toolbar` VB6 est un ActiveX** (Common Controls) : on la **remplace** par le **`ToolStrip` natif** plutôt que de l'interopérer.
- Le `ToolStrip` apporte images par bouton, types d'éléments variés et barres déplaçables ; la `StatusBar` devient **`StatusStrip`**.
- Préférez **`MenuStrip`** (moderne) à l'ancien `MainMenu`.

---

## 🔗 Renvois

- **[13.3](03-formulaires-mdi.md)** — fusion de menus en MDI (`MergeAction`/`MergeIndex`), `MdiWindowListItem`. · **[13.6](06-control-arrays.md)** — tableaux de menus (éléments dynamiques).
- **[14.5](../14-graphismes-impression-activex/README.md)** — contrôles ActiveX/OCX (interop, quand on ne remplace pas). · **[14.3](../14-graphismes-impression-activex/README.md)** — images et `ImageList`/ressources.
- **Suite → [13.8 — Boîtes de dialogue communes (`CommonDialog` → `OpenFileDialog`, `ColorDialog`…)](08-dialogues-communs.md)**

⏭️ [Boîtes de dialogue communes (`CommonDialog` → `OpenFileDialog`, `ColorDialog`…)](/13-formulaires-winforms/08-dialogues-communs.md)
