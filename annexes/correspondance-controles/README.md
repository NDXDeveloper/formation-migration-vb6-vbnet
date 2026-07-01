🔝 Retour au [Sommaire](/SOMMAIRE.md)

# C. Correspondance des contrôles VB6 → Windows Forms

**Tableau contrôle par contrôle**, avec propriétés/événements équivalents et points d'attention
(*control arrays*, contrôles intrinsèques sans équivalent direct).

> 🎯 **Cadre.** Cette annexe couvre la migration de l'IHM **VB6 → Windows Forms** sur
> **.NET Framework 4.7.2** — la cible qui offre la **meilleure parité du concepteur** et des
> contrôles intrinsèques (module 1.2). On *migre* des formulaires existants, on ne les redessine
> pas (modules 13-14).

> ⚠️ **Quatre changements transversaux** à garder en tête avant tout tableau :
> 1. **`Caption` → `Text`** sur la plupart des contrôles porteurs de libellé.
> 2. **Unités : twips → pixels** (le `ScaleMode` disparaît).
> 3. **Couleurs : `OLE_COLOR` (`Long`) → `System.Drawing.Color`**.
> 4. **Pas de *tableaux de contrôles*** (control arrays) : voir §3, le point le plus structurant.

---

## 1. Propriétés transversales (communes à la plupart des contrôles)

> Ces correspondances valent pour `Form`, `Label`, `TextBox`, `Button`, etc. Elles ne sont **pas
> répétées** dans les tableaux par contrôle (§4-§5).

| VB6 | Windows Forms | ⚠️ Point d'attention |
|---|---|---|
| `Caption` | `Text` | Sur `Form`, `Label`, `Button`, `CheckBox`, `RadioButton`, `GroupBox`, menus. |
| `Text` (TextBox) | `Text` | Conservé. |
| `Name` | `Name` | Conservé. |
| `Index` (control array) | *(aucun)* | ⚠️ Plus de tableaux de contrôles. Voir **§3**. |
| `Enabled` / `Visible` | `Enabled` / `Visible` | Conservés. |
| `Left` / `Top` / `Width` / `Height` | identiques **ou** `Location` (`Point`) / `Size` (`Size`) | ⚠️ En **pixels**, plus en twips. Préférer `Location`/`Size`/`Bounds`. |
| `Move l, t, w, h` (méthode) | `SetBounds(...)` / affecter `Bounds` | — |
| `BackColor` / `ForeColor` | identiques | ⚠️ Type **`Color`**, plus `OLE_COLOR`. Voir **§7**. |
| `Font.Name` / `Font.Size` / `Font.Bold` | `Font` (objet `Font`) | ⚠️ Le `Font` est **immuable** : affecter un **nouvel** objet `Font(...)`, pas ses sous-propriétés. |
| `FontName` / `FontSize` (ancien VB) | `Font` | Idem. |
| `Tag` | `Tag` | ⚠️ Désormais **`Object`** (plus `String`) — pratique pour porter un identifiant (voir §3). |
| `TabIndex` / `TabStop` | `TabIndex` / `TabStop` | Conservés. |
| `ToolTipText` | composant `ToolTip` → `tt.SetToolTip(ctl, "…")` | ⚠️ **Plus une propriété directe** : on ajoute un composant `ToolTip` et on appelle `SetToolTip`. |
| `MousePointer` / `MouseIcon` | `Cursor` (+ classe `Cursors`) | — |
| `BorderStyle` | `BorderStyle` | Énumération différente (`None`/`FixedSingle`/`Fixed3D`). |
| `Appearance` (Flat/3D) | `FlatStyle` (selon contrôle) | — |
| `Alignment` (libellé/texte) | `TextAlign` | — |
| `hWnd` | `Handle` (`IntPtr`) | — |
| `hDC` | `CreateGraphics()` / `e.Graphics` (Paint) | ⚠️ Plus de `hDC` persistant : dessin via **GDI+**. Voir **§8**. |
| `Container` | `Parent` | — |
| `ZOrder` (méthode) | `BringToFront()` / `SendToBack()` | — |
| `Refresh` (méthode) | `Refresh()` / `Invalidate()` | — |
| `SetFocus` (méthode) | `Focus()` / `Select()` | — |

---

## 2. Événements transversaux

| VB6 | Windows Forms | ⚠️ Point d'attention |
|---|---|---|
| `Click` | `Click` | Conservé. |
| `DblClick` | `DoubleClick` | — |
| `GotFocus` / `LostFocus` | `Enter` / `Leave` (et `GotFocus`/`LostFocus` existent) | ⚠️ Préférer `Enter`/`Leave` ; l'**ordre et la sémantique** diffèrent. La validation passe par `Validating`/`Validated`. |
| `Change` | événement **spécifique** : `TextChanged`, `ValueChanged`, `SelectedIndexChanged`, `CheckedChanged`… | ⚠️ **`Change` n'existe plus tel quel** : il se décline selon le contrôle. |
| `KeyDown` / `KeyUp` | `KeyDown` / `KeyUp` (`KeyEventArgs`) | `KeyCode`/`Shift` → `e.KeyCode` / `e.Modifiers`. |
| `KeyPress` | `KeyPress` (`KeyPressEventArgs`) | ⚠️ `KeyAscii` (`Integer`) → **`e.KeyChar`** (`Char`). Annuler une touche : VB6 `KeyAscii = 0` → **`e.Handled = True`**. |
| `MouseDown` / `MouseUp` / `MouseMove` | mêmes noms (`MouseEventArgs`) | `Button`/`X`/`Y` → `e.Button` / `e.X` / `e.Y` ; état Shift via `Control.ModifierKeys`. |
| `Paint` | `Paint` (`PaintEventArgs`, `e.Graphics`) | Voir **§8**. |
| `Resize` | `Resize` / `SizeChanged` | ⚠️ Souvent **remplaçable** par `Anchor`/`Dock` (voir §9). |
| `Validate(Cancel)` | `Validating` (`CancelEventArgs`) / `Validated` | Nouveau modèle robuste ; propriété `CausesValidation`. |
| `Initialize` (form/classe) | constructeur `Sub New()` | — |
| `Terminate` | `Dispose` / `Finalize` | ⚠️ Plus de destruction déterministe. Voir **Annexe B.3**. |

> 📌 **Cycle de vie du formulaire** (différences) :
> VB6 `Initialize → Load → Activate → … → QueryUnload → Unload → Terminate`
> WinForms `New() → Load → Activated → Shown → … → FormClosing → FormClosed → Dispose`
> - `QueryUnload` (paramètre `Cancel`, raison) → **`FormClosing`** (`e.Cancel`, `e.CloseReason`).
> - `Unload` → **`FormClosed`** ; `Unload Me` → `Me.Close()`.
> - `Form1.Show vbModal` → **`f.ShowDialog()`** ; `Form1.Show` (modeless) → `f.Show()`.
> - ⚠️ VB6 **instancie automatiquement** un formulaire dès qu'on le référence. En VB.NET, préférer
>   `Dim f As New Form1` explicite (les *instances par défaut* existent mais se comportent
>   différemment).

---

## 3. ⚠️ Les *tableaux de contrôles* (control arrays) — le point le plus structurant

> **Aucun équivalent direct en Windows Forms.** C'est souvent le plus gros chantier d'une migration
> d'IHM.

### Le mécanisme VB6
Plusieurs contrôles partagent **le même `Name`** avec un **`Index`** différent ; on pouvait en
**créer/détruire à l'exécution** (`Load`/`Unload`) ; **un seul gestionnaire** d'événement recevait
le paramètre `Index`.

### Les équivalents .NET
- **Gestionnaire partagé** : associer le **même** gestionnaire aux événements de plusieurs contrôles
  (via `Handles … , … , …` ou `AddHandler`), puis identifier l'émetteur avec **`sender`**
  (le caster, lire `.Name` ou `.Tag`).
- **Collection explicite** : ranger les contrôles dans un `List(Of Control)` / tableau que l'on gère
  soi-même.
- **Création dynamique** en code : `New Button()`, ajout à `Controls`, câblage par `AddHandler`.
- ⚠️ L'assistant de mise à niveau **émule** les tableaux de contrôles via
  `Microsoft.VisualBasic.Compatibility` (classes `…Array` générées) : c'est la **béquille à
  retirer** (module 4.2).

```vb
' VB6 — tableau de contrôles : un seul gestionnaire, paramètre Index
Private Sub cmdNum_Click(Index As Integer)
    txtAffichage.Text = txtAffichage.Text & CStr(Index)
End Sub
```

```vbnet
' VB.NET — gestionnaire PARTAGÉ + sender (le .Tag porte le "numéro")
Private Sub cmdNum_Click(sender As Object, e As EventArgs) _
        Handles cmd0.Click, cmd1.Click, cmd2.Click
    Dim b = DirectCast(sender, Button)
    txtAffichage.Text &= b.Tag.ToString()
End Sub

' Création dynamique équivalente à Load cmdNum(i)
For i As Integer = 0 To 9
    Dim b As New Button With {.Text = i.ToString(), .Tag = i}
    AddHandler b.Click, AddressOf cmdNum_Click
    Me.Controls.Add(b)
Next
```

---

## 4. Contrôles intrinsèques VB6 → Windows Forms

| Contrôle VB6 | Équivalent WinForms | Correspondances clés | ⚠️ Point d'attention |
|---|---|---|---|
| `Form` | `Form` | `Caption`→`Text` ; `WindowState`→`WindowState` ; `BorderStyle`→`FormBorderStyle` ; `ControlBox`/`MinButton`/`MaxButton`→`ControlBox`/`MinimizeBox`/`MaximizeBox` ; `StartUpPosition`→`StartPosition` | Cycle de vie différent (voir §2). `ScaleMode`/`Scale*` **supprimés** → `ClientSize`. `KeyPreview` conservé. |
| `Label` | `Label` | `Caption`→`Text` ; `Alignment`→`TextAlign` ; `AutoSize`→`AutoSize` | `BackStyle` (Transparent/Opaque) → `BackColor = Color.Transparent`. |
| `TextBox` | `TextBox` | `Text`→`Text` ; `MultiLine`→`Multiline` ; `Locked`→`ReadOnly` ; `MaxLength`→`MaxLength` ; `PasswordChar`→`PasswordChar` ; `ScrollBars`→`ScrollBars` | `Change`→**`TextChanged`** ; `SelStart`/`SelLength`/`SelText`→`SelectionStart`/`SelectionLength`/`SelectedText`. |
| `CommandButton` | `Button` | `Caption`→`Text` ; bouton par défaut → `Form.AcceptButton` ; bouton annulation → `Form.CancelButton` ; `Picture`→`Image` | `Default`/`Cancel` ne sont **plus** des propriétés du bouton mais du **formulaire**. `Style` (Standard/Graphical) → `FlatStyle`/`Image`. |
| `CheckBox` | `CheckBox` | `Caption`→`Text` ; `Value`→`Checked` (+ `CheckState`) | ⚠️ `Value` (**0/1/2**, `Integer`) → `Checked` (**`Boolean`**). Tri-état : `ThreeState=True` + `CheckState`. `Click`→`CheckedChanged`. |
| `OptionButton` | `RadioButton` | `Caption`→`Text` ; `Value`→`Checked` | Regroupement par **conteneur** (`GroupBox`/`Panel`), comme le `Frame` VB6. `Click`→`CheckedChanged`. |
| `Frame` | `GroupBox` (ou `Panel`) | `Caption`→`Text` | ⚠️ `GroupBox` a un libellé ; `Panel` **non** (mais défile, `AutoScroll`). |
| `ListBox` | `ListBox` | `AddItem`→`Items.Add` ; `RemoveItem`→`Items.RemoveAt` ; `Clear`→`Items.Clear` ; `List`/`ListCount`→`Items`/`Items.Count` ; `ListIndex`→`SelectedIndex` ; `Text`→`SelectedItem`/`Text` ; `MultiSelect`→`SelectionMode` ; `Selected(i)`→`GetSelected`/`SetSelected` ; `Sorted`→`Sorted` | ⚠️ **`ItemData` sans équivalent** : stocker des **objets** dans `Items` (avec `ToString`), ou une liste parallèle. `Click`→`SelectedIndexChanged` ; `DblClick`→`DoubleClick`. |
| `ComboBox` | `ComboBox` | comme `ListBox` + `Style`→`DropDownStyle` | `Style` 0/1/2 → `DropDown`/`Simple`/`DropDownList`. `Change`→`TextChanged` ; `Click`→`SelectedIndexChanged`. |
| `PictureBox` | `PictureBox` **et/ou** `Panel` | `Picture`→`Image` ; `AutoSize`→`SizeMode` | ⚠️ **Rôle scindé** : le `PictureBox` VB6 était *conteneur* **et** *surface de dessin* **et** *image*. En WinForms : image → `PictureBox` ; **conteneur** → `Panel` ; **dessin** → `Paint`/`Graphics` (§8). |
| `Image` | `PictureBox` (avec `SizeMode`) | `Stretch`→`SizeMode.StretchImage` | ⚠️ Pas de contrôle *Image* léger distinct : utiliser `PictureBox`. |
| `Timer` | `Timer` (`System.Windows.Forms.Timer`) | `Interval`→`Interval` ; `Enabled`→`Enabled` ; `Timer`→**`Tick`** | Devient un **composant** (bac à composants), non visible sur le formulaire. |
| `HScrollBar` / `VScrollBar` | `HScrollBar` / `VScrollBar` | `Min`/`Max`/`Value`/`SmallChange`/`LargeChange`→identiques | `Change`/`Scroll`→`Scroll`/`ValueChanged`. Souvent remplaçables par `AutoScroll` du conteneur ou par `TrackBar`. |
| `Menu` (Éditeur de menus) | `MenuStrip` + `ToolStripMenuItem` | `Caption`→`Text` ; `Checked`→`Checked` ; `Shortcut`→`ShortcutKeys` ; séparateur (`-`)→`ToolStripSeparator` | ⚠️ L'**Éditeur de menus** est remplacé par le concepteur `MenuStrip`. `mnuX_Click`→`ToolStripMenuItem.Click`. Menus contextuels : `PopupMenu`→`ContextMenuStrip`. Menus dynamiques (control arrays de menus) → ajout en code. |
| `Line` | *(aucun contrôle)* | — | ⚠️ **Sans équivalent.** Dessiner avec `e.Graphics.DrawLine` dans `Paint` (§8), ou *bidouille* `Panel`/`Label` à bordure. |
| `Shape` | *(aucun contrôle)* | — | ⚠️ **Sans équivalent.** Dessiner (rectangle/ellipse) avec `Graphics` dans `Paint` (§8). |
| `Data` (DAO/Jet lié) | *(aucun équivalent direct)* | — | ⚠️ **Modèle de liaison entièrement différent** : `BindingSource` + ADO.NET (`DataSet`/`DataTable`), voir §10 et **module 15**. |
| `OLE` (conteneur OLE) | *(aucun équivalent simple)* | — | ⚠️ Incorporation OLE largement obsolète. Recourir à un composant spécifique / `WebBrowser` selon le besoin. |
| `DriveListBox` | *(aucun)* | — | ⚠️ **Sans équivalent.** Utiliser `OpenFileDialog`/`FolderBrowserDialog`, ou reconstruire avec `ComboBox` + `DriveInfo.GetDrives()`. |
| `DirListBox` | *(aucun)* | — | ⚠️ **Sans équivalent.** `FolderBrowserDialog`, ou `TreeView` + `Directory.GetDirectories`. |
| `FileListBox` | *(aucun)* | — | ⚠️ **Sans équivalent.** `OpenFileDialog`, ou `ListView` + `Directory.GetFiles`. |

---

## 5. Contrôles ActiveX courants (Common Controls et autres OCX fréquents)

> Ces contrôles étaient livrés en **OCX** (souvent *Microsoft Windows Common Controls*). La plupart
> ont un **équivalent natif** WinForms — mais ce ne sont pas des remplacements « au pixel près ».

| OCX VB6 | Équivalent WinForms | ⚠️ Point d'attention |
|---|---|---|
| `Toolbar` | `ToolStrip` | Modèle de boutons (`ToolStripButton`) différent. |
| `StatusBar` | `StatusStrip` (+ `ToolStripStatusLabel`) | Panneaux gérés autrement. |
| `ProgressBar` | `ProgressBar` | `Min`/`Max`/`Value` conservés. |
| `Slider` | `TrackBar` | — |
| `TabStrip` | `TabControl` (+ `TabPage`) | ⚠️ Le `TabStrip` VB6 ne fournissait **que les onglets** (on gérait les pages avec des `Frame`). Le `TabControl` **inclut** les `TabPage` : remaniement conceptuel. |
| `SSTab` (Tabbed Dialog) | `TabControl` | Idem. |
| `ImageList` | `ImageList` | Associé aux contrôles par propriété (`ImageList`/`ImageIndex`). |
| `TreeView` | `TreeView` | API de nœuds (`Nodes`) proche mais distincte. |
| `ListView` | `ListView` | `ListItems`/`SubItems`→`Items`/`SubItems` ; modes d'affichage. |
| `MonthView` | `MonthCalendar` | — |
| `DTPicker` | `DateTimePicker` | `Value` (`Date`). Voir formats de date **Annexe B.10**. |
| `UpDown` (buddy) | `NumericUpDown` | ⚠️ Le `UpDown` VB6 était un contrôle *associé* (buddy) ; `NumericUpDown` **fusionne** valeur + boutons. |
| `RichTextBox` (RichTx32) | `RichTextBox` | `TextRTF`→`Rtf`. |
| `MaskedEdit` | `MaskedTextBox` | Syntaxe de masque différente. |
| `MSFlexGrid` / `MSHFlexGrid` | `DataGridView` | ⚠️ Pas un remplacement direct : repenser remplissage et liaison. |
| `DataGrid` (lié) | `DataGridView` | Liaison via `BindingSource` (§10, module 15). |
| `CoolBar` | `ToolStripContainer` / `ToolStrip` | Approximatif. |
| `Animation` | *(aucun direct)* | `PictureBox` avec GIF animé, ou personnalisé. |
| `Winsock` | `System.Net.Sockets` (`TcpClient`/`TcpListener`/`UdpClient`) | ⚠️ **Plus un contrôle** : API réseau (module 16). |
| `Inet` (Internet Transfer) | `WebClient` / `HttpClient` | ⚠️ Plus un contrôle. |
| `MSComm` (série) | `System.IO.Ports.SerialPort` | ⚠️ Plus un contrôle. |
| `CommonDialog` (ComDlg32) | **Plusieurs** classes | Voir **§6** (scission). |

---

## 6. La boîte de dialogue commune (`CommonDialog`) — scindée

Le contrôle unique `CommonDialog` (méthodes `ShowOpen`/`ShowSave`/`ShowColor`/`ShowFont`/
`ShowPrinter`/`ShowHelp`) éclate en **classes distinctes** :

| VB6 (`CommonDialog`) | Windows Forms | Correspondances |
|---|---|---|
| `ShowOpen` | `OpenFileDialog` | `Filter`→`Filter` ; `FilterIndex`→`FilterIndex` ; `FileName`→`FileName` ; `InitDir`→`InitialDirectory` ; `DialogTitle`→`Title` |
| `ShowSave` | `SaveFileDialog` | idem |
| `ShowColor` | `ColorDialog` | `Color`→`Color` |
| `ShowFont` | `FontDialog` | `Font`→`Font` |
| `ShowPrinter` | `PrintDialog` / `PrintPreviewDialog` | lié à `PrintDocument` (impression : module dédié) |
| *(parcourir un dossier)* | `FolderBrowserDialog` | nouveau, pratique |
| `ShowHelp` | *(aucun direct)* | aide hors champ WinForms |

> ⚠️ `ShowXxx` renvoyait via le contrôle ; en WinForms, `dlg.ShowDialog()` renvoie un
> **`DialogResult`** (`OK`/`Cancel`). Les `Flags` (entier binaire) deviennent des **propriétés
> booléennes** nommées (ex. `Multiselect`, `CheckFileExists`).

---

## 7. Couleurs : `OLE_COLOR` → `System.Drawing.Color`

| VB6 | Windows Forms | ⚠️ Point d'attention |
|---|---|---|
| Propriété typée `OLE_COLOR` (`Long`) | propriété typée `Color` | Le `Long` portait soit une valeur **RGB**, soit une **couleur système** (ex. `&H8000000F`). |
| `RGB(r, g, b)` | `Color.FromArgb(r, g, b)` | — |
| Constantes `vbButtonFace`, `vbWindowBackground`… | `SystemColors.Control`, `SystemColors.Window`… | Couleurs système via la classe **`SystemColors`**. |
| Valeur `OLE_COLOR` héritée (stockée/configurée) | `ColorTranslator.FromOle(valeur)` | ⚠️ Pour convertir une valeur **OLE** existante. `ColorTranslator.ToOle` pour l'inverse. |
| `vbRed`, `vbBlue`… | `Color.Red`, `Color.Blue`… | Voir **Annexe A**, §10 (constantes). |

---

## 8. Dessin sur formulaire / `PictureBox`

Les **méthodes graphiques** VB6 (`Line`, `Circle`, `PSet`, `Point`, `Print`, `Cls`) et les  
propriétés associées (`AutoRedraw`, `DrawWidth`, `DrawStyle`, `FillColor`, `FillStyle`,
`CurrentX`/`CurrentY`) **disparaissent**. Le dessin passe par **GDI+**, dans l'événement `Paint`.

| VB6 | Windows Forms (GDI+) |
|---|---|
| `Me.Line (x1,y1)-(x2,y2)` | `e.Graphics.DrawLine(pen, x1, y1, x2, y2)` |
| `Me.Circle (x,y), r` | `e.Graphics.DrawEllipse(pen, x-r, y-r, 2*r, 2*r)` |
| `Me.PSet (x,y)` | `e.Graphics.FillRectangle(brush, x, y, 1, 1)` |
| `Me.Print "texte"` | `e.Graphics.DrawString("texte", Font, brush, x, y)` |
| `Me.Cls` | `e.Graphics.Clear(BackColor)` |
| `AutoRedraw = True` | redessiner dans `Paint` (et/ou `DoubleBuffered = True`) |

> ⚠️ En WinForms, **rien n'est persistant** : tout dessin doit être **redessiné** à chaque `Paint`
> (sinon il disparaît au moindre rafraîchissement). Alternative : dessiner dans un `Bitmap` que l'on
> affiche. Voir aussi **Annexe A**, §11.

```vbnet
' VB.NET — dessin dans l'événement Paint
Private Sub Form1_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
    Using p As New Pen(Color.Blue, 2)
        e.Graphics.DrawLine(p, 10, 10, 200, 100)
        e.Graphics.DrawEllipse(p, 50, 50, 80, 80)
    End Using
End Sub
```

---

## 9. Nouveautés utiles : ancrage et docking

VB6 n'avait **pas** de redimensionnement automatique : on repositionnait les contrôles à la main  
dans l'événement `Resize`. Windows Forms apporte deux propriétés qui **remplacent** souvent ce code :

| Besoin | Windows Forms |
|---|---|
| Garder un contrôle collé à un/des bord(s) lors du redimensionnement | propriété **`Anchor`** |
| Remplir une zone / s'amarrer à un bord du conteneur | propriété **`Dock`** |

> 💡 **Opportunité de migration** : du code `Resize` manuel peut souvent être **supprimé** au profit
> de `Anchor`/`Dock`. L'assistant ne le fait **pas** automatiquement — c'est un gain de
> simplification à faire lors du refactoring (module 17.5).

---

## 10. Liaison de données (aperçu)

La liaison VB6 (contrôle `Data`, `DataGrid`/contrôles liés via `DataSource`/`DataField`,  
DAO/RDO/ADO Data) **n'a pas d'équivalent direct** : le modèle WinForms repose sur **ADO.NET**
(`DataSet`/`DataTable`), un **`BindingSource`** et `DataGridView`, avec
`control.DataBindings.Add(...)`.

> ⚠️ Chantier à part entière, traité au **module 15** (données). Ne pas tenter de « mapper » le
> contrôle `Data` ligne à ligne : repenser la couche d'accès.

---

## 🔗 Annexes et modules liés

- **Annexe A** — Tableau de correspondance VB6 → VB.NET (langage, `Caption`/`Text` hors contrôles,
  constantes de couleur).
- **Annexe B** — Pièges silencieux (B.3 finalisation/`Terminate`, B.7 propriétés par défaut et
  `Set`, B.8 `Value` booléen, B.10 formats de date des sélecteurs).
- **Modules 13-14** — Migration de l'interface (formulaires, contrôles, événements).
- **Module 15** — Données (remplace le contrôle `Data` et la liaison VB6).
- **Module 16** — Interop (`Winsock`/`Inet`/`MSComm` → API .NET).
- **Module 17.5** — Refactoring idiomatique (ex. `Resize` manuel → `Anchor`/`Dock`).

---

**Juin 2026**  
**Cible** : .NET Framework 4.7.2 (le « pont » depuis VB6) · Windows Forms  
**Licence** : Creative Commons BY-NC-SA 4.0

⏭️ [Correspondance des types de données](/annexes/correspondance-types/README.md)
