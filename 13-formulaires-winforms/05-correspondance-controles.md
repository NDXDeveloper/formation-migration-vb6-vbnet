🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.5 — Correspondance des contrôles intrinsèques (`CommandButton`→`Button`, `TextBox`, `Label`, `Frame`, `OptionButton`…)

> **Chapitre 13 — Des formulaires VB6 à Windows Forms** · Section 13.5
> Le **catalogue** des contrôles intrinsèques : à quel contrôle Windows Forms correspond chaque contrôle VB6, et **quelles propriétés/événements changent de nom** au passage.

---

## 🧭 Périmètre

La plupart des contrôles intrinsèques VB6 ont un **équivalent direct** en Windows Forms. La difficulté n'est pas de **trouver** le contrôle cible, mais de connaître les **renommages** de propriétés et d'événements — sources de **changements silencieux**. Cette section donne l'**essentiel** et les **pièges** ; le tableau **exhaustif** (propriété par propriété) est en **annexe C**.

---

## 1. Vue d'ensemble : le tableau de correspondance

| Contrôle VB6 | Contrôle Windows Forms | Remarque |
|--------------|------------------------|----------|
| `CommandButton` | `Button` | — |
| `TextBox` | `TextBox` | plusieurs renommages (voir §3) |
| `Label` | `Label` | `Alignment` → `TextAlign` |
| `Frame` | `GroupBox` (ou `Panel`) | `GroupBox` = bord + titre ; `Panel` = conteneur nu |
| `OptionButton` | `RadioButton` | `Value` → `Checked` |
| `CheckBox` | `CheckBox` | `Value` (0/1/2) → `Checked` / `CheckState` |
| `ListBox` | `ListBox` | modèle `Items` (voir §3) |
| `ComboBox` | `ComboBox` | `Style` → `DropDownStyle` |
| `PictureBox` | `PictureBox` / `Panel` | image **vs** conteneur **vs** dessin (voir §3) |
| `Image` (contrôle) | `PictureBox` | pas de contrôle « Image » léger distinct |
| `Timer` | `Timer` | événement `Timer` → **`Tick`** ; dans le bac à composants |
| `HScrollBar` / `VScrollBar` | `HScrollBar` / `VScrollBar` | — |
| `Shape`, `Line` | *(aucun)* | → GDI+ (14) ou `Panel`/`Label` bordé |
| `Data` | *(aucun)* | → liaison de données (`BindingSource`, 15) |
| `DriveListBox`/`DirListBox`/`FileListBox` | *(aucun)* | → boîtes de dialogue (13.8) ou à reconstruire |
| `OLE` (conteneur) | *(aucun)* | rarement utilisé |

---

## 2. Les renommages qui reviennent **partout**

Avant le détail par contrôle, voici les changements **transversaux** — ils touchent **de nombreux** contrôles :

| VB6 | Windows Forms | Note |
|-----|---------------|------|
| `Caption` | **`Text`** | titre (Button, Label, CheckBox, RadioButton, GroupBox, Form) — déjà vu en **[13.1](01-modele-formulaire.md)** |
| `Value` (option/case) | **`Checked`** | booléen |
| `BackColor` / `ForeColor` | `BackColor` / `ForeColor` | ⚠️ **type** différent : `OLE_COLOR`/`Long` → `System.Drawing.Color` ; `vbButtonFace` → `SystemColors.Control` |
| `FontName`/`FontSize`/`FontBold`… | un seul objet **`Font`** | `FontBold` → `Font.Bold` ; le `Font` se construit en bloc |
| `hWnd` | **`Handle`** | poignée de fenêtre (appels d'API) |
| `Tag` | `Tag` | mais de type **`Object`** (et non `String`) |
| `ToolTipText` | composant **`ToolTip`** | ⚠️ plus de propriété par contrôle : on ajoute **un** composant `ToolTip` partagé |
| `MousePointer` / `MouseIcon` | **`Cursor`** | `Cursors.Hand`, etc. |
| `Index` (control array) | *(supprimé)* | → **[13.6](06-control-arrays.md)** |

> ⚠️ **Couleurs et polices** changent de **type**, pas seulement de nom. Une couleur VB6 (un `Long`) et une police VB6 (propriétés éclatées) deviennent des **objets .NET** (`Color`, `Font`). Les **constantes système** (`vbButtonFace`, `vbWindowText`…) se traduisent par `SystemColors.*`.

---

## 3. Contrôle par contrôle : les points qui piègent

### `CommandButton` → `Button`
- `Caption` → `Text`.
- ⚠️ **`Default` et `Cancel` migrent vers le FORMULAIRE** : `Default = True` → `Form.AcceptButton = btn` (bouton activé par **Entrée**) ; `Cancel = True` → `Form.CancelButton = btn` (activé par **Échap**). Ce ne sont **plus** des propriétés du bouton.
- Bouton graphique (`Style`/`Picture`) → `Image` + `FlatStyle`.

### `TextBox` → `TextBox`
- ⚠️ `MultiLine` → **`Multiline`** (changement de **casse**).
- ⚠️ `Locked` → **`ReadOnly`**.
- ⚠️ `SelStart` / `SelLength` / `SelText` → **`SelectionStart`** / **`SelectionLength`** / **`SelectedText`**.
- ⚠️ événement `Change` → **`TextChanged`**.
- `PasswordChar` → `PasswordChar` (ou `UseSystemPasswordChar`).

### `Label` → `Label`
- `Caption` → `Text`.
- ⚠️ `Alignment` (0/1/2) → **`TextAlign`** (énum `ContentAlignment` à **9** valeurs).

### `OptionButton` → `RadioButton`
- `Caption` → `Text`.
- ⚠️ `Value` (booléen) → **`Checked`**.
- Regroupement : par **conteneur** (`GroupBox`/`Panel`/formulaire), comme en VB6 — mais le regroupement par **control array** disparaît (**13.6**).
- Pour réagir à la sélection : **`CheckedChanged`** (plutôt que `Click`).

### `CheckBox` → `CheckBox`
- `Caption` → `Text`.
- ⚠️ `Value` **numérique** (0/1/2) → **`Checked`** (booléen) pour deux états, ou **`CheckState`** (`Unchecked`/`Checked`/`Indeterminate`) pour trois états.
- Événement de bascule : **`CheckedChanged`**.

### `Frame` → `GroupBox` (ou `Panel`)
- `Caption` → `Text`.
- Choix : **`GroupBox`** si l'on veut le **bord + titre** ; **`Panel`** si le cadre servait juste de **conteneur** (éventuellement sans bordure).

### `ListBox` / `ComboBox`
- ⚠️ Les **méthodes** `AddItem` / `RemoveItem` / `Clear` → le modèle **`Items`** : `Items.Add(...)` / `Items.RemoveAt(i)` / `Items.Clear()`.
- ⚠️ `List(i)` → `Items(i)` ; `ListIndex` → **`SelectedIndex`** ; `ListCount` → **`Items.Count`**.
- ⚠️ `ItemData` (tableau parallèle de `Long`) → **pas d'équivalent direct** : stockez des **objets** dans `Items` (avec `DisplayMember`/`ValueMember`) ou une structure parallèle.
- Sélection : événement `Click` → **`SelectedIndexChanged`**.
- ⚠️ ComboBox `Style` (0/1/2) → **`DropDownStyle`** (`DropDown` / `Simple` / `DropDownList`).
- `MultiSelect` (ListBox) → `SelectionMode`.

### `PictureBox` → `PictureBox` / `Panel`
Le `PictureBox` VB6 était **polyvalent** ; en WinForms, ses rôles se **séparent** :
- **affichage d'image** → `PictureBox` (`Picture` → **`Image`** ; `StdPicture` → `Image` ; `SizeMode`) ;
- **conteneur** → **`Panel`** ;
- ⚠️ **dessin** (`Line`, `Circle`, `PSet`, `Print`, `Cls`, `AutoRedraw`) → **GDI+** sur l'événement `Paint`, **sans équivalent direct** → **[chapitre 14](../14-graphismes-impression-activex/README.md)**.

### `Timer` → `Timer`
- `Interval` → `Interval` (en **ms**) ; plus de limite à 65535 ms.
- ⚠️ événement `Timer` → **`Tick`**.
- C'est un **composant** (bac à composants), pas un contrôle visible ; `Enabled`, ou `Start()`/`Stop()`.

---

## 4. Les contrôles **sans équivalent direct**

Certains contrôles VB6 **n'ont pas** de pendant intrinsèque en Windows Forms et imposent une **autre approche** :

| VB6 | Solution en Windows Forms |
|-----|---------------------------|
| `Shape`, `Line` | **dessin GDI+** (14.1), ou un `Panel`/`Label` **bordé** pour les cas simples |
| `Data` (contrôle) | **liaison de données** via `BindingSource` (**[chapitre 15](../15-acces-donnees/README.md)**) |
| `DriveListBox`, `DirListBox`, `FileListBox` | **boîtes de dialogue** (`OpenFileDialog`, `FolderBrowserDialog`, **[13.8](08-dialogues-communs.md)**), ou reconstruction (système de fichiers + `TreeView`/`ListView`) |
| `OLE` (conteneur) | *(pas d'équivalent ; usage devenu rare)* |

---

## 🔁 Récapitulatif (renommages les plus fréquents)

| VB6 | Windows Forms |
|-----|---------------|
| `Caption` | `Text` |
| `Value` (option/case) | `Checked` / `CheckState` |
| `AddItem` / `List(i)` / `ListIndex` / `ListCount` | `Items.Add` / `Items(i)` / `SelectedIndex` / `Items.Count` |
| `Change` | `TextChanged` |
| `Click` (liste/option/case) | `SelectedIndexChanged` / `CheckedChanged` |
| `Locked` (TextBox) | `ReadOnly` |
| `MultiLine` | `Multiline` |
| `SelStart`/`SelLength`/`SelText` | `SelectionStart`/`SelectionLength`/`SelectedText` |
| `Default`/`Cancel` (bouton) | `Form.AcceptButton`/`Form.CancelButton` |
| `Timer` (événement) | `Tick` |
| `ToolTipText` | composant `ToolTip` |
| `MousePointer` | `Cursor` |

---

## ✅ Points clés

- La plupart des contrôles ont un **équivalent direct** ; l'enjeu est dans les **renommages** de propriétés/événements.
- Transversaux à retenir : **`Caption`→`Text`**, **`Value`→`Checked`**, **`Change`→`TextChanged`**, et les **couleurs/polices** qui deviennent des **objets .NET**.
- Listes/combos passent au **modèle `Items`** (`AddItem`→`Items.Add`, `ListIndex`→`SelectedIndex`) ; `ItemData` est à **repenser**.
- Le `PictureBox` VB6 se **scinde** : image (`PictureBox`), conteneur (`Panel`), **dessin** (GDI+, **chapitre 14**).
- Quelques contrôles **disparaissent** (`Shape`/`Line`, `Data`, `Drive/Dir/File ListBox`) : prévoir une **autre approche**.
- Pour le détail **exhaustif**, garder ouverte l'**annexe C**.

---

## 🔗 Renvois

- **[13.1](01-modele-formulaire.md)** — `Caption`→`Text`, anatomie. · **[13.6](06-control-arrays.md)** — `Index` / *control arrays*.
- **[13.8](08-dialogues-communs.md)** — boîtes de dialogue (remplacent Drive/Dir/File). · **[Chapitre 14](../14-graphismes-impression-activex/README.md)** — dessin (PictureBox, Shape, Line). · **[Chapitre 15](../15-acces-donnees/README.md)** — `Data` / liaison.
- **Annexe C** — *Correspondance des contrôles VB6 → Windows Forms* (référence complète).
- **Suite → [13.6 — Les tableaux de contrôles (*control arrays*) : disparus → solutions de remplacement](06-control-arrays.md)** ⭐ ⚠️

⏭️ [Les **tableaux de contrôles** (*control arrays*) : **disparus** → solutions de remplacement](/13-formulaires-winforms/06-control-arrays.md)
