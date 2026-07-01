🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.1 Dessin VB6 (`Line`, `Circle`, `PSet`, `Print`) → **GDI+** / objet `Graphics` ⭐ ⚠️

> **La section fondatrice du chapitre.** En VB6, on dessinait *sur* un formulaire. En .NET, on
> dessine *via* un objet `Graphics`, et toujours **au bon moment** : dans l'événement `Paint`.
> Comprendre cette bascule, c'est éviter le bug n°1 de la migration graphique — *le dessin qui
> s'efface*.

**Module 14 — Graphismes, impression et contrôles ActiveX** · Cible : .NET Framework 4.7.2 · GDI+
(`System.Drawing`)

---

## 🧭 L'idée centrale en une phrase

En VB6, `Me.Line`, `Me.Circle`, `Me.PSet` et `Me.Print` sont des **méthodes du formulaire** qui  
peignent **immédiatement** sur sa surface, et — avec `AutoRedraw = True` — y **restent**.  
En .NET/GDI+, ces opérations passent par un objet **`Graphics`** que l'on **n'obtient que le temps  
d'un rendu** : on décrit *comment* se dessiner, dans `Paint`, et le système rappelle ce code à
**chaque** rafraîchissement.

> ⚠️ **À retenir absolument** : un dessin réalisé **ailleurs que dans `Paint`** (par exemple dans
> `Form_Load` ou dans un gestionnaire de clic) **disparaîtra** dès que la fenêtre est redimensionnée,
> recouverte, réduite puis restaurée. Ce n'est pas un plantage : le rendu n'a simplement pas été
> **reproduit**.

---

## 1. Mode « immédiat et persistant » → mode « à la demande »

| | VB6 | .NET / GDI+ |
|---|-----|-------------|
| **Sur quoi** | Le formulaire ou la `PictureBox` directement (`Me.Line…`) | Un objet `Graphics` |
| **Quand** | N'importe quand, n'importe où | Dans l'événement `Paint` (idéalement) |
| **Persistance** | `AutoRedraw = True` mémorise le dessin | Aucune : il faut **redessiner** à chaque rendu |
| **Déclenchement d'un nouveau rendu** | Automatique | `Invalidate()` / `Refresh()` |
| **Curseur de dessin** | `CurrentX` / `CurrentY` | Inexistant : coordonnées **explicites** à chaque appel |

La règle .NET tient en une formule : **on ne dessine pas une fois pour toutes, on indique comment se  
redessiner.**

---

## 2. Obtenir un objet `Graphics` : trois sources, un seul bon réflexe

| Source | Usage | Précautions |
|--------|-------|-------------|
| **`e.Graphics`** (dans `Paint`) | ✅ Le cas normal : tout dessin **persistant** | Ne **pas** le `Dispose` (il appartient au système) |
| **`Control.CreateGraphics()`** | Dessin **transitoire** très ponctuel (rare) | À **`Dispose`** ; le dessin **ne survit pas** au prochain `Paint` |
| **`Graphics.FromImage(bmp)`** | Dessiner **hors écran**, dans un `Bitmap` tampon | À **`Dispose`** ; sert à imiter `AutoRedraw` (voir §8) |

```vb
Private Sub Canevas_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
    ' e.Graphics est fourni par le système : on dessine ICI.
    e.Graphics.DrawLine(Pens.Black, 0, 0, 100, 100)
End Sub
```

> ⚠️ **Piège fréquent** : utiliser `CreateGraphics()` pour un dessin censé rester affiché. Cela
> « marche » à l'écran… jusqu'au premier rafraîchissement, qui efface tout. `CreateGraphics()` n'est
> **pas** un substitut à `AutoRedraw`.

---

## 3. Le modèle d'objets GDI+ : `Pen`, `Brush`, `Font`, `Color`

VB6 pilotait le dessin par des **propriétés** du formulaire (`ForeColor`, `DrawWidth`, `FillColor`…).  
GDI+ utilise des **objets explicites**, passés en paramètre à chaque méthode :

| Propriété VB6 | Objet / membre GDI+ | Espace de noms |
|---------------|---------------------|----------------|
| `ForeColor` (trait) | Un `Pen` (couleur + épaisseur) | `System.Drawing` |
| `DrawWidth` | `Pen.Width` | `System.Drawing` |
| `DrawStyle` (trait plein, tirets…) | `Pen.DashStyle` (`Solid`, `Dash`, `Dot`, `DashDot`…) | `System.Drawing.Drawing2D` |
| `FillColor` + `FillStyle` (plein) | Un `SolidBrush` | `System.Drawing` |
| `FillStyle` (hachures) | Un `HatchBrush` (`HatchStyle.Cross`…) | `System.Drawing.Drawing2D` |
| `FontName`, `FontSize`, `FontBold`… | Un `Font` | `System.Drawing` |
| `RGB(r, g, b)` | `Color.FromArgb(r, g, b)` ou couleurs nommées | `System.Drawing` |

```vb
Using stylo As New Pen(Color.Blue, 2),
      pinceau As New SolidBrush(Color.LightYellow),
      police As New Font("Segoe UI", 9)
    e.Graphics.FillRectangle(pinceau, 10, 10, 200, 100)
    e.Graphics.DrawRectangle(stylo, 10, 10, 200, 100)
    e.Graphics.DrawString("Texte", police, Brushes.Black, 14, 14)
End Using
```

> 🔗 **`Pen`, `Brush`, `Font`, `Bitmap` sont `IDisposable`.** Ils encapsulent des *handles* GDI
> natifs. C'est l'application directe du **problème de la finalisation déterministe** (modules 2 et
> 12) : si vous ne les libérez pas (idéalement avec `Using`), vous provoquez des **fuites de handles
> GDI** — l'application « ralentit puis ne dessine plus » après un temps d'usage. Pour les couleurs,
> épaisseurs et pinceaux **standards**, préférez les objets partagés `Pens.Black`, `Brushes.Red`,
> `SystemBrushes.Control`… : ils sont gérés par le framework et ne se libèrent **pas**.

---

## 4. Correspondance des méthodes de dessin

| VB6 | GDI+ (`Graphics g`) | Point d'attention |
|-----|---------------------|-------------------|
| `Line (x1,y1)-(x2,y2)` | `g.DrawLine(stylo, x1, y1, x2, y2)` | — |
| `Line (x1,y1)-(x2,y2), , B` | `g.DrawRectangle(stylo, x, y, l, h)` | **Deux coins** → origine + largeur/hauteur |
| `Line (x1,y1)-(x2,y2), , BF` | `g.FillRectangle(pinceau, x, y, l, h)` | Rempli → `Brush`, pas `Pen` |
| `Circle (x,y), r` | `g.DrawEllipse(stylo, x-r, y-r, 2*r, 2*r)` | **Centre+rayon** → **boîte englobante** ⚠️ |
| `Circle …, , début, fin` (arc) | `g.DrawArc(stylo, …, angleDébut, balayage)` | **Radians → degrés**, sens inversé ⚠️ |
| `PSet (x,y), couleur` | `bmp.SetPixel(x, y, couleur)` | Pas d'API « pixel » sur `Graphics` ⚠️ |
| `Point(x,y)` (lire) | `bmp.GetPixel(x, y)` | Nécessite un `Bitmap` |
| `Print "texte"` | `g.DrawString(texte, police, pinceau, x, y)` | Pas de curseur `CurrentX/Y` ⚠️ |
| `Cls` | `g.Clear(couleurFond)` | Efface toute la surface |
| `PaintPicture img, …` | `g.DrawImage(image, …)` | — |

Les quatre cas marqués ⚠️ méritent un détail.

### `Line … B` / `BF` : du couple de coins aux dimensions

VB6 prend **deux coins opposés** ; GDI+ prend **origine + largeur + hauteur**. La conversion  
calcule les dimensions (et doit gérer le cas où les coins ne sont pas dans l'ordre attendu) :

```vb
' VB6 : Line (x1,y1)-(x2,y2), , B
Dim x = Math.Min(x1, x2), y = Math.Min(y1, y2)
Dim l = Math.Abs(x2 - x1), h = Math.Abs(y2 - y1)
g.DrawRectangle(stylo, x, y, l, h)        ' contour
' ... ou FillRectangle(pinceau, x, y, l, h) pour BF
```

### `Circle` : centre+rayon → boîte englobante (et le piège des arcs)

`DrawEllipse` ne prend **pas** un centre et un rayon, mais le **rectangle englobant**. Pour un
`Circle (cx, cy), r` :

```vb
g.DrawEllipse(stylo, cx - r, cy - r, 2 * r, 2 * r)   ' contour
g.FillEllipse(pinceau, cx - r, cy - r, 2 * r, 2 * r) ' disque plein
```

> ⚠️ **Arcs.** `Circle` accepte des angles **de début et de fin en radians**, mesurés dans le sens
> **trigonométrique**. `DrawArc` attend un **angle de départ et un angle de balayage en degrés**,
> dans le sens **horaire**. Il faut donc convertir (`degrés = radians × 180 / π`) **et** ajuster les
> signes/sens — **vérifiez toujours le rendu visuellement**. De plus, l'astuce VB6 de l'**angle
> négatif** (qui trace en prime un rayon du centre vers l'extrémité) n'a **pas d'équivalent** :
> ajoutez un `DrawLine` explicite si besoin.

### `PSet` / `Point` : pas de pixel sur `Graphics` (et un vrai piège de performance)

`Graphics` n'expose pas de méthode « poser un pixel ». On passe par un `Bitmap` :

```vb
bmp.SetPixel(x, y, couleur)          ' équivaut à PSet
Dim c As Color = bmp.GetPixel(x, y)  ' équivaut à Point()
```

> ⚠️ **Performance.** `GetPixel`/`SetPixel` sont **très lents** (verrouillage interne à chaque
> appel). Un code VB6 qui trace **pixel par pixel** (courbes, fractales, traitement d'image) sera
> inutilisable tel quel. Migrez ce type de boucle vers un dessin **dans un `Bitmap`** avec
> **`LockBits`** (accès direct à la mémoire), ou repensez-le avec des primitives GDI+ (`DrawLine`
> entre points successifs, etc.). Pour un `PSet` avec `DrawWidth > 1`, VB6 dessinait un petit pavé :
> utilisez `FillRectangle`/`FillEllipse` de la taille voulue.

### `Print` : du curseur de texte à `DrawString` (ou `TextRenderer`)

`Print` avançait un curseur (`CurrentX`/`CurrentY`) ; GDI+ exige des **coordonnées explicites** et
un **`Brush`** pour la couleur :

```vb
g.DrawString("Bonjour", police, Brushes.Black, x, y)
```

> ℹ️ **Deux rendus de texte coexistent.** `Graphics.DrawString` (GDI+) gère rotation, dégradés et
> rectangles de mise en page. Mais les contrôles Windows Forms affichent leur texte via **GDI**
> (`TextRenderer.DrawText`), au rendu légèrement différent. Pour que votre texte dessiné **ressemble**
> à celui des `Label`/`TextBox`, préférez :
> ```vb
> TextRenderer.DrawText(g, "Bonjour", police, New Point(x, y), Color.Black)
> ```
> Réservez `DrawString` aux besoins propres à GDI+ (texte tourné, `StringFormat`, retour à la ligne
> dans un rectangle…).

---

## 5. Couleurs : `RGB`, `QBColor`, et les valeurs `OLE_COLOR`

Deux situations, deux conversions :

- **Vous avez les composantes** r, g, b → `Color.FromArgb(r, g, b)`.
- **Vous héritez d'une valeur `Long`** issue de `RGB(...)`, de `QBColor(...)`, d'une propriété de
  contrôle, ou d'une **couleur système** (`vbButtonFace`…) → ce sont des valeurs **`OLE_COLOR`**
  (ordre **BGR**, bit de poids fort pour les couleurs système). Convertissez avec
  **`ColorTranslator.FromOle`** :

```vb
Dim c As Color = ColorTranslator.FromOle(valeurOleLong)
```

> ⚠️ **Piège silencieux.** Passer directement une valeur `RGB()` de VB6 là où .NET attend un entier
> ARGB inverse le rouge et le bleu (BGR ≠ RGB). Pour toute valeur de couleur **héritée**, passez par
> `ColorTranslator.FromOle`. Pour les couleurs **système**, préférez la classe `SystemColors`
> (`SystemColors.Control`, `SystemColors.Highlight`…).

---

## 6. `CurrentX` / `CurrentY` : le curseur de dessin a disparu

VB6 mémorisait une position courante : `Print` la faisait avancer ligne après ligne, et
`Line -(x, y)` poursuivait le tracé depuis elle. **GDI+ n'a aucun curseur** : chaque appel est
**absolu**. Pour migrer un code « à curseur », **maintenez vos propres variables** :

```vb
Dim curseurY As Single = 10
g.DrawString("Ligne 1", police, Brushes.Black, 10, curseurY)
curseurY += police.GetHeight(g)            ' on avance "à la main"
g.DrawString("Ligne 2", police, Brushes.Black, 10, curseurY)
```

---

## 7. `DrawMode` et le tracé en `XOR` (élastique / *rubber-band*)

VB6 permettait, via `DrawMode = vbInvert`, de tracer une ligne ou un cadre **réversible** (le  
classique rectangle de sélection « élastique » que l'on dessine puis efface en le retraçant). GDI+  
ne fournit pas ce mode d'inversion sur `Graphics`. Deux options :

- ✅ **Idiomatique** : ne pas inverser du tout — stocker l'état (le rectangle en cours) et
  **`Invalidate`** pour redessiner proprement (voir §8).
- 🔧 **Au plus proche** : `ControlPaint.DrawReversibleLine(...)` et
  `ControlPaint.DrawReversibleFrame(...)`, qui travaillent en **coordonnées écran**.

> ⚠️ Le tracé `XOR` est un **faux-ami** : il « marche » mais devient illisible sur fonds complexes
> et se comporte mal avec le double-tampon. Sur du code migré, l'approche **état + `Invalidate`** est
> presque toujours préférable.

---

## 8. Remplacer `AutoRedraw` : faire « rester » le dessin

C'est le **cœur pratique** de la migration. Deux stratégies, selon votre objectif.

### Stratégie A — Redessiner depuis un **état** (recommandée, idiomatique)

On ne conserve pas l'**image**, on conserve **ce qu'il faut dessiner** (un modèle), et `Paint` le  
reproduit intégralement :

```vb
Private _formes As New List(Of Rectangle)   ' l'ÉTAT

Private Sub Canevas_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
    Using stylo As New Pen(Color.Blue, 2)
        For Each r In _formes
            e.Graphics.DrawRectangle(stylo, r)   ' tout est (re)dessiné
        Next
    End Using
End Sub

Private Sub AjouterForme(r As Rectangle)
    _formes.Add(r)
    Invalidate()      ' demande un nouveau rendu -> Paint sera rappelé
End Sub
```

### Stratégie B — Dessiner dans un **`Bitmap` tampon** (au plus proche de `AutoRedraw = True`)

Quand le dessin est **accumulé** et coûteux à reconstruire (un éditeur de type « peinture »), on  
imite la mémoire d'image de VB6 : on peint dans un `Bitmap`, et `Paint` se contente de le recopier.

```vb
Private _tampon As Bitmap

Private Sub Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    _tampon = New Bitmap(ClientSize.Width, ClientSize.Height)
    Using g = Graphics.FromImage(_tampon) : g.Clear(BackColor) : End Using
End Sub

Private Sub DessinerLigne(p1 As Point, p2 As Point)
    Using g = Graphics.FromImage(_tampon), stylo As New Pen(ForeColor)
        g.DrawLine(stylo, p1, p2)        ' on écrit dans le tampon
    End Using
    Invalidate()
End Sub

Private Sub Form_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
    e.Graphics.DrawImage(_tampon, 0, 0)  ' comportement "AutoRedraw"
End Sub
```

> ⚠️ Avec la stratégie B, pensez à **recréer le `Bitmap`** lors d'un redimensionnement (sinon le
> dessin est tronqué) et à **`Dispose`** l'ancien tampon.

### Confort de rendu : `Invalidate` et le double-tampon

- `Invalidate()` salit toute la surface ; `Invalidate(rect)` n'en redemande qu'une **zone**
  (plus efficace). `Refresh()` force un rendu **immédiat**.
- Contre le **scintillement** (que VB6 atténuait via `ClipControls`/`AutoRedraw`), activez le
  **double-tampon** : `Me.DoubleBuffered = True` sur un formulaire ; pour un contrôle personnalisé,
  `SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True)`.

---

## 9. Exemple de migration (avant / après)

**VB6** (avec `AutoRedraw = True`, le dessin reste affiché) :

```vb
Private Sub Form_Click()
    Me.ForeColor = RGB(0, 0, 255)
    Me.DrawWidth = 2
    Me.Line (10, 10)-(210, 110), , B     ' rectangle (deux coins)
    Me.Line (10, 10)-(210, 110)          ' diagonale
    Me.Circle (110, 60), 40              ' cercle (centre, rayon)
    Me.CurrentX = 10 : Me.CurrentY = 130
    Me.Print "Surface dessinee"
End Sub
```

**VB.NET** (l'état déclenche le rendu ; tout est dessiné dans `Paint`) :

```vb
Private _dessiner As Boolean = False

Private Sub Form1_Click(sender As Object, e As EventArgs) Handles Me.Click
    _dessiner = True
    Invalidate()
End Sub

Private Sub Form1_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
    If Not _dessiner Then Return
    Dim g = e.Graphics
    Using stylo As New Pen(Color.Blue, 2),
          police As New Font("Segoe UI", 9)
        g.DrawRectangle(stylo, 10, 10, 200, 100)   ' coins (10,10)-(210,110) -> 200 x 100
        g.DrawLine(stylo, 10, 10, 210, 110)         ' diagonale
        g.DrawEllipse(stylo, 70, 20, 80, 80)        ' Circle (110,60),40 -> (70,20,80,80)
        g.DrawString("Surface dessinée", police, Brushes.Blue, 10, 130)
    End Using
End Sub
```

> ℹ️ Les coordonnées de l'exemple supposent un repère **en pixels**. Si le formulaire VB6 d'origine
> raisonnait en **twips** (le défaut), les valeurs doivent être **converties** — c'est l'objet de la
> **section 14.2**.

---

## 10. ⚠️ Pièges silencieux à retenir

- **Dessiner hors de `Paint`** (dans `Load`, un clic…) : le dessin **disparaît** au premier
  rafraîchissement. → Tout dessin persistant va dans `Paint`, déclenché par `Invalidate()`.
- **Oublier `Dispose`** sur `Pen`/`Brush`/`Font`/`Bitmap` créés : **fuites de handles GDI**.
  → `Using`, ou objets partagés (`Pens.*`, `Brushes.*`, `SystemColors.*`).
- **`Circle` → `DrawEllipse`** : centre+rayon ≠ boîte englobante. Erreur de position/taille
  classique.
- **Arcs** : radians/degrés et **sens** différents ; l'angle négatif VB6 n'existe pas.
- **`PSet` en boucle** : `SetPixel` est lent → `LockBits` ou primitives GDI+.
- **Couleurs héritées** (`RGB`/`OLE_COLOR`) passées telles quelles : **rouge/bleu inversés**.
  → `ColorTranslator.FromOle`.
- **`CurrentX`/`CurrentY`** : curseur disparu → coordonnées explicites, position suivie « à la main ».
- **Tracé `XOR`** (`DrawMode = vbInvert`) : pas d'équivalent direct → `ControlPaint.DrawReversible*`
  ou, mieux, état + `Invalidate`.
- **Rendu de texte** : `DrawString` (GDI+) ≠ `TextRenderer` (GDI) ≠ `Print` VB6 → léger décalage
  visuel. Choisir selon le besoin de cohérence avec l'UI.

> 🔗 Ces pièges figurent au **catalogue de l'Annexe B** (finalisation/handles, conversions). Gardez
> aussi en tête le lien avec le module 12 (`IDisposable`/`Using`).

---

**Section suivante → 14.2 — Système de coordonnées et `ScaleMode` : repenser le rendu**, pour
convertir proprement les unités (twips → pixels) et retrouver le bon repère.

⏭️ [Système de coordonnées et `ScaleMode` : repenser le rendu](/14-graphismes-impression-activex/02-coordonnees-scalemode.md)
