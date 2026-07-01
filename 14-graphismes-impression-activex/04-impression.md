🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.4 Impression : objet `Printer` → `System.Drawing.Printing` (`PrintDocument`) ⚠️

> En VB6, **vous poussiez** le contenu vers l'imprimante, page après page, puis `EndDoc`. En .NET,
> **le framework vous tire** les pages : il déclenche l'événement `PrintPage` et vous lui répondez
> « voici cette page, et oui/non il en reste une autre ». Cette **inversion de logique** est le vrai
> chantier de la section — bien plus que la correspondance des méthodes.

**Module 14 — Graphismes, impression et contrôles ActiveX** · Cible : .NET Framework 4.7.2 ·
`System.Drawing.Printing`

---

## 1. Le changement de modèle : séquentiel/impératif → événementiel

| | VB6 (`Printer`) | .NET (`PrintDocument`) |
|---|-----------------|------------------------|
| **Nature** | Objet **global**, séquentiel, à état | **Instance** que vous créez |
| **Logique** | Vous **poussez** le contenu (`Print`, `NewPage`, `EndDoc`) | Le framework **tire** les pages via `PrintPage` |
| **Saut de page** | `Printer.NewPage` (en plein code) | `e.HasMorePages = True` puis on rend la main |
| **Fin du document** | `Printer.EndDoc` | `PrintPage` retourne avec `HasMorePages = False` |
| **Dessin** | `Printer.Line/Circle/Print…` | GDI+ sur `e.Graphics` (**section 14.1 s'applique telle quelle**) |

> ⚠️ **La conséquence structurelle.** Le code VB6 d'impression est souvent une **procédure linéaire**
> (une boucle qui imprime, avec un `NewPage` au milieu). En .NET, il faut le **restructurer en machine
> à états** : `PrintPage` est appelé **une fois par page**, et vous devez **reprendre** là où vous
> vous étiez arrêté (quelle ligne, quel enregistrement). C'est le point délicat (§4).

Bonne nouvelle : **le dessin est identique à celui de l'écran**. `DrawString`, `DrawLine`,
`DrawImage`… s'utilisent sur `e.Graphics` exactement comme dans la section 14.1.

---

## 2. Le modèle d'objets

L'unique objet `Printer` de VB6 éclate en un **document** entouré d'objets de configuration :

| VB6 | .NET | Rôle |
|-----|------|------|
| `Printer` (objet global) | `PrintDocument` (instance) | Le document à imprimer |
| `Printer.DeviceName`, `Copies`, `Duplex`… | `PrinterSettings` | **Quelle** imprimante, copies, recto-verso, plage |
| `Printer.PaperSize`, `Orientation`, marges | `PageSettings` (et `Margins`) | **Mise en page** : papier, orientation, marges |
| (dans la boucle d'impression) | `PrintPageEventArgs e` | Le **contexte d'une page** |

### Propriétés courantes

| VB6 | .NET |
|-----|------|
| `Printer.Orientation` (1=portrait, 2=paysage) | `PageSettings.Landscape` (`Boolean`) |
| `Printer.PaperSize` | `PageSettings.PaperSize` (`PaperSize` / `PaperKind`) |
| `Printer.Copies` | `PrinterSettings.Copies` |
| `Printer.Duplex` | `PrinterSettings.Duplex` |
| `Printer.ColorMode` | `PageSettings.Color` (`Boolean`) |
| `Printer.DeviceName` | `PrinterSettings.PrinterName` |
| `Printers` (collection) | `PrinterSettings.InstalledPrinters` |
| `Printer.hDC` | `e.Graphics.GetHdc()` (rarement utile) |
| `Printer.ScaleWidth/Height` | `e.MarginBounds` / `e.PageBounds` |

---

## 3. Méthodes et dessin

| VB6 | .NET (`e As PrintPageEventArgs`) | Remarque |
|-----|----------------------------------|----------|
| `Printer.Print "txt"` | `e.Graphics.DrawString("txt", police, pinceau, x, y)` | Pas de curseur (§4) |
| `Printer.Line/Circle/PSet` | `e.Graphics.DrawLine/DrawEllipse/…` | Comme **14.1** |
| `Printer.PaintPicture` | `e.Graphics.DrawImage(...)` | — |
| `Printer.CurrentX/Y` | Variables suivies **à la main** | — |
| `Printer.NewPage` | `e.HasMorePages = True` **puis** sortir | **Changement majeur** ⚠️ |
| `Printer.EndDoc` | Retourner avec `HasMorePages = False` | C'est la valeur par défaut |
| `Printer.KillDoc` | `e.Cancel = True` (dans `PrintPage`) | Annulation |
| `Printer.TextWidth/TextHeight` | `e.Graphics.MeasureString("txt", police)` | **Indispensable** pour paginer |
| `Printer.Print` (lancement) | `doc.Print()` | Démarre le cycle `PrintPage` |

> ℹ️ **Cas particulier : `Form.PrintForm`.** VB6 offrait `Me.PrintForm` pour imprimer d'un coup une
> **image du formulaire** (capture « écran ») sur l'imprimante par défaut. Il **n'a pas d'équivalent**
> dans le modèle `PrintDocument`. Reproduisez-le en **capturant** le formulaire dans un `Bitmap` via
> **`Control.DrawToBitmap`**, puis en imprimant ce bitmap dans `PrintPage` :
> ```vb
> ' VB.NET — équivalent de PrintForm : capturer le formulaire puis l'imprimer
> Private Sub _doc_PrintPage(sender As Object, e As PrintPageEventArgs) Handles _doc.PrintPage
>     Using bmp As New Bitmap(Me.ClientSize.Width, Me.ClientSize.Height)
>         Me.DrawToBitmap(bmp, Me.ClientRectangle)          ' capture du formulaire
>         e.Graphics.DrawImage(bmp, e.MarginBounds.Location) ' impression de l'image
>     End Using
> End Sub
> ```
> Si le formulaire **dépasse** la page, passez un **rectangle de destination** à `DrawImage`
> (par ex. `e.MarginBounds`) pour le **mettre à l'échelle**.
> Le composant `Microsoft.VisualBasic.PowerPacks.Printing.PrintForm` **émule** aussi `PrintForm`, au
> prix d'une **dépendance** aux *PowerPacks* — préférez `DrawToBitmap`, sans dépendance.

---

## 4. ⚠️ Le vrai défi : la pagination en **machine à états**

Il ne faut **plus** appeler `NewPage` au fil du code. À la place :

1. on imprime tant que **le contenu tient dans la page** (`y <= e.MarginBounds.Bottom`) **et** qu'il
   reste des éléments ;
2. on indique s'il faut **une page de plus** via `e.HasMorePages` ;
3. on **mémorise l'index** de reprise dans un **champ**, car `PrintPage` sera rappelé pour la page
   suivante.

```vb
Private WithEvents _doc As New PrintDocument()
Private _lignes As List(Of String)
Private _index As Integer            ' où reprendre entre deux pages

Private Sub Imprimer()
    _doc.DocumentName = "Liste"
    _doc.Print()                     ' déclenche PrintPage autant de fois que nécessaire
End Sub

Private Sub _doc_BeginPrint(sender As Object, e As PrintEventArgs) Handles _doc.BeginPrint
    _index = 0                       ' réinitialiser l'état AVANT chaque impression
End Sub

Private Sub _doc_PrintPage(sender As Object, e As PrintPageEventArgs) Handles _doc.PrintPage
    Using police As New Font("Segoe UI", 10)
        Dim hauteurLigne As Single = police.GetHeight(e.Graphics)
        Dim y As Single = e.MarginBounds.Top
        ' Imprimer tant qu'on tient dans la page ET qu'il reste des lignes
        While _index < _lignes.Count AndAlso y + hauteurLigne <= e.MarginBounds.Bottom
            e.Graphics.DrawString(_lignes(_index), police, Brushes.Black,
                                  e.MarginBounds.Left, y)
            y += hauteurLigne
            _index += 1
        End While
        ' Reste-t-il des lignes ? -> demander une page supplémentaire
        e.HasMorePages = (_index < _lignes.Count)
    End Using
End Sub
```

> ℹ️ `WithEvents` + `Handles` est le câblage d'événements idiomatique de VB.NET (cf. module 12.8). On
> peut aussi utiliser `AddHandler`.

---

## 5. ⚠️ Unités et marges : un repère **différent** de l'écran et de VB6

C'est un piège de coordonnées propre à l'impression :

- L'unité **par défaut** de `e.Graphics` à l'impression est le **centième de pouce**
  (`GraphicsUnit.Display` sur une imprimante) — **ni twips** (VB6) **ni pixels** (écran).
- `e.MarginBounds` et `e.PageBounds` sont des `Rectangle` exprimés en **centièmes de pouce** (marge
  par défaut : **1 pouce = 100 unités** sur les quatre côtés).
- Par défaut (`OriginAtMargins = False`), l'**origine (0,0)** est au **coin supérieur gauche de la
  zone imprimable** — au bord des marges **matérielles** de l'imprimante (`HardMarginX/Y`), et
  **non** au coin de **vos** marges (`PageSettings.Margins` est alors **ignoré**). Réglez
  `PrintDocument.OriginAtMargins = True` pour que l'origine tienne compte de **vos** marges.

> ⚠️ **Deux conséquences concrètes.**
> 1. Si vous portez des coordonnées **en twips**, l'échelle est fausse (cf. **section 14.2** pour la
>    conversion).
> 2. Si vous changez `e.Graphics.PageUnit` (par ex. en millimètres), votre **dessin** passe en mm
>    mais `e.MarginBounds` **reste en centièmes de pouce** → désalignement. Tenez-vous-en à une seule
>    convention, et dessinez de préférence **à l'intérieur de `e.MarginBounds`**.

Pour mémoire, la classe `Margins` (`PageSettings.Margins`) est elle aussi en centièmes de pouce.

---

## 6. Texte imprimé : `DrawString` (et non `TextRenderer`)

Pour l'écran, la section 14.1 proposait `TextRenderer.DrawText` pour coller au rendu des contrôles.
**À l'impression, c'est l'inverse** : utilisez **`Graphics.DrawString`** (GDI+), le chemin conçu pour
l'impression. `TextRenderer` (GDI) se comporte de façon moins fiable sur un contexte d'imprimante.

Une `Font` exprimée en **points** s'imprime à sa **taille physique** réelle ; mesurez les blocs avec
`MeasureString` pour décider des retours à la ligne et des sauts de page :

```vb
Dim taille As SizeF = e.Graphics.MeasureString(texte, police, e.MarginBounds.Width)
' taille.Height -> hauteur nécessaire (avec retour à la ligne dans la largeur de marge)
```

---

## 7. Cycle de vie : `BeginPrint` / `QueryPageSettings` / `EndPrint`

`PrintDocument` expose plusieurs événements, utiles pour **initialiser** et **nettoyer** :

| Événement | Usage typique |
|-----------|---------------|
| `BeginPrint` | **Réinitialiser l'état** (l'index de reprise), ouvrir des ressources |
| `QueryPageSettings` | Ajuster la mise en page **page par page** (marges différentes…) |
| `PrintPage` | Dessiner la page courante (le cœur) |
| `EndPrint` | **Libérer** les ressources (polices créées une seule fois, fichiers…) |

> ℹ️ Par défaut, `doc.Print()` affiche une petite **boîte d'état** « Impression de la page N… ». Pour
> l'imprimer **silencieusement**, affectez un contrôleur standard :
> `doc.PrintController = New Printing.StandardPrintController()`.

---

## 8. ➕ Les cadeaux de .NET : dialogues et **aperçu** intégrés

VB6 offrait `CommonDialog.ShowPrinter` mais **aucun aperçu** intégré. Windows Forms fournit, prêts à  
l'emploi et **réutilisant le même gestionnaire `PrintPage`** :

```vb
' Choix de l'imprimante / copies / plage
Using dlg As New PrintDialog()
    dlg.Document = _doc
    If dlg.ShowDialog() = DialogResult.OK Then _doc.Print()
End Using

' Aperçu avant impression — gratuit, sans réécrire le rendu
Using apercu As New PrintPreviewDialog()
    apercu.Document = _doc
    apercu.ShowDialog()
End Using
```

`PageSetupDialog` complète l'ensemble (marges, orientation, format papier). **L'aperçu sans effort**
est l'un des vrais gains de la migration.

---

## 9. Exemple de migration (avant / après)

**VB6** — impression **linéaire**, `NewPage` au milieu de la boucle :

```vb
Private Sub ImprimerListe(lignes() As String)
    Dim i As Long
    Printer.CurrentY = 0
    For i = 0 To UBound(lignes)
        If Printer.CurrentY + Printer.TextHeight("X") > Printer.ScaleHeight Then
            Printer.NewPage
        End If
        Printer.Print lignes(i)
    Next
    Printer.EndDoc
End Sub
```

**VB.NET** — même résultat, mais piloté par `PrintPage` (voir le squelette complet du §4) : la boucle
`For…NewPage…EndDoc` devient un **`PrintPage`** qui imprime ce qui tient, mémorise l'index dans
`_index`, et positionne `e.HasMorePages`. La logique « je pousse jusqu'à la fin » se transforme en
« je réponds page par page ».

---

## 10. ⚠️ Pièges silencieux à retenir

- **Garder une logique linéaire** (`NewPage`/`EndDoc` en plein code) → ne se transpose pas. Il faut
  une **machine à états** avec index de reprise dans un **champ**.
- **Oublier de réinitialiser l'index** entre deux impressions → seconde impression incomplète.
  Réinitialiser dans **`BeginPrint`**.
- **Boucle infinie** si l'on ne met jamais `e.HasMorePages = False`, ou si une ligne **ne tient
  jamais** dans la page (vérifier qu'au moins un élément avance par page).
- **Unité par défaut = 1/100 de pouce** (ni twips, ni pixels). Coordonnées twips portées telles
  quelles → échelle fausse (section 14.2).
- **`PageUnit` modifié** mais `e.MarginBounds` reste en 1/100 de pouce → **désalignement**.
- **Origine (0,0) au coin de la zone imprimable** (marges matérielles), **pas** à vos marges, par
  défaut → dessiner dans `e.MarginBounds`, ou régler `OriginAtMargins`.
- **`TextRenderer` à l'impression** → préférer **`DrawString`**.
- **`Pen`/`Brush`/`Font` non libérés** → fuites de handles ; créer/libérer proprement (`Using`,
  `BeginPrint`/`EndPrint`), cf. **module 12**.
- **Boîte d'état d'impression** inattendue → `StandardPrintController` pour imprimer en silence.

> 🔗 Voir la **section 14.1** (dessin GDI+, réutilisé tel quel à l'impression), la **section 14.2**
> (unités/coordonnées), le **module 12** (`IDisposable`) et l'**Annexe B** (pièges silencieux).

---

**Section suivante → 14.5 — Contrôles ActiveX/OCX : réutiliser via interop ou remplacer** 🔗, où la
dépendance aux composants binaires devient une **décision d'architecture**.

⏭️ [Contrôles **ActiveX/OCX** : réutiliser via interop ou remplacer](/14-graphismes-impression-activex/05-activex-ocx.md)
