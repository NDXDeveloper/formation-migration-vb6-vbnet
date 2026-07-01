🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.2 Système de coordonnées et `ScaleMode` : repenser le rendu

> En VB6, un formulaire vierge mesurait quelque chose comme **4800 sur 3600** ; le même formulaire,
> migré, en fait **320 sur 240**. Rien n'a rétréci : on a changé d'**unité**. Tant qu'on s'en tient
> au cas par défaut, la migration des coordonnées se résume à **une conversion d'unité**. Les
> systèmes de coordonnées *personnalisés* de VB6, eux, demandent un vrai changement d'approche.

**Module 14 — Graphismes, impression et contrôles ActiveX** · Cible : .NET Framework 4.7.2 · GDI+
(`System.Drawing`)

---

## 🧭 La bonne nouvelle d'abord : l'orientation ne change pas

VB6 (par défaut) et Windows Forms partagent **le même repère** : origine en **haut-gauche**, axe X  
vers la droite, axe **Y vers le bas**. Contrairement à d'autres systèmes graphiques, **aucune  
inversion d'axe** n'est nécessaire dans le cas courant.

Ce qui change, c'est l'**unité** :

| | VB6 | Windows Forms / GDI+ |
|---|-----|----------------------|
| **Unité par défaut** | **Twips** (1/1440 de pouce) | **Pixels** |
| **Réglage de l'unité** | Propriété `ScaleMode` | `Graphics.PageUnit` (+ `PageScale`) |
| **Repère personnalisable** | `ScaleLeft/Top/Width/Height`, `Scale` | Transformations (`Matrix`, `TranslateTransform`…) |
| **Position des contrôles** | Dans l'unité du **conteneur** | Toujours en **pixels** |
| **Coordonnées souris** | Dans l'unité du conteneur (twips par défaut) | En **pixels** |

> ⚠️ **Le piège n°1 de cette section** : du code de dessin qui contient des **constantes en twips**
> (`Me.Line (1440, 720)…`) et qu'on porte tel quel. Les coordonnées étant désormais en **pixels**, le
> dessin apparaît **~15 fois trop petit**, tassé en haut à gauche — voire **hors champ**. Le code
> compile, le rendu est faux.

---

## 1. Le modèle VB6 : twips, `ScaleMode` et repères sur mesure

### Les twips et les modes prédéfinis

Le **twip** vaut 1/20 de point, soit **1/1440 de pouce**. La propriété `ScaleMode` (sur `Form`,
`PictureBox` et `Printer`) choisit l'unité :

| `ScaleMode` | Valeur | Unité |
|-------------|--------|-------|
| `vbUser` | 0 | Définie par `ScaleLeft/Top/Width/Height` |
| `vbTwips` | 1 | Twips (défaut) |
| `vbPoints` | 2 | Points |
| `vbPixels` | 3 | Pixels |
| `vbCharacters` | 4 | Caractères |
| `vbInches` | 5 | Pouces |
| `vbMillimeters` | 6 | Millimètres |
| `vbCentimeters` | 7 | Centimètres |

### Les repères personnalisés

VB6 permettait de **redéfinir entièrement** le système de coordonnées d'une surface :

- `ScaleLeft`, `ScaleTop` : coordonnées du coin haut-gauche ;
- `ScaleWidth`, `ScaleHeight` : « taille logique » de la zone (valeurs **négatives** autorisées → axe
  inversé) ;
- l'instruction `Scale (x1, y1)-(x2, y2)` fait tout en une fois ;
- les fonctions `ScaleX`, `ScaleY` convertissent une mesure d'un mode vers un autre.

C'est ainsi qu'on obtenait, par exemple, un repère mathématique (origine en bas, Y vers le haut) pour  
tracer une courbe :

```vb
' VB6 : (0,0) en bas-gauche, (100,100) en haut-droite, Y vers le HAUT
Me.Scale (0, 100)-(100, 0)
```

> ℹ️ Le facteur twips↔pixels **dépend du DPI** de l'écran. VB6 l'exposait via
> `Screen.TwipsPerPixelX` / `Screen.TwipsPerPixelY` (et `Printer.TwipsPerPixelX/Y`). À **96 DPI**,
> 1 pouce = 1440 twips = 96 pixels, soit **15 twips par pixel** — mais ce 15 n'est vrai **qu'**à
> 96 DPI.

---

## 2. Le modèle .NET : pixels, `PageUnit` et transformations

En Windows Forms, **tout est en pixels** par défaut : `Location`, `Size`, `Bounds`, les coordonnées  
souris (`MouseEventArgs.X/Y`) et le dessin GDI+.

Pour retrouver un équivalent de `ScaleMode`, GDI+ offre **deux leviers** sur l'objet `Graphics` :

- **`PageUnit`** (`GraphicsUnit`) : l'unité de dessin — `Pixel` (défaut), `Point`, `Inch`,
  `Millimeter`, `Document` (1/300 pouce), `Display`.
- **`PageScale`** : un facteur d'échelle appliqué aux coordonnées.

```vb
' Dessiner en millimètres (utile surtout pour l'impression, cf. 14.4)
g.PageUnit = GraphicsUnit.Millimeter
g.DrawRectangle(Pens.Black, 10, 10, 50, 30)   ' rectangle de 50 x 30 mm
```

> ⚠️ Quand `PageUnit` n'est **pas** `Pixel`, l'épaisseur d'un `Pen` s'exprime dans cette unité : un
> `Pen` de largeur **1** devient **1 mm** (épais !). Pour un trait fin d'**un pixel** quel que soit le
> mode, utilisez un **`Pen` de largeur 0** (GDI+ le rend toujours sur 1 pixel device).

Pour les repères **sur mesure** (décalage d'origine, mise à l'échelle, axe inversé), on n'utilise pas
`PageUnit` mais les **transformations** du monde : `TranslateTransform`, `ScaleTransform`,
`RotateTransform`, ou une `Matrix` complète.

---

## 3. Tableau de correspondance

| VB6 | .NET / GDI+ | Remarque |
|-----|-------------|----------|
| `ScaleMode = vbPixels` | `Graphics.PageUnit = GraphicsUnit.Pixel` (défaut) | Cas le plus simple |
| `ScaleMode = vbPoints/Inches/Millimeters` | `PageUnit = Point/Inch/Millimeter` | Unités physiques |
| `ScaleMode = vbTwips` | Convertir en pixels **ou** `PageUnit=Point` + `PageScale=1/20` | Pas de `GraphicsUnit` « twip » |
| `ScaleMode = vbCentimeters` | `PageUnit=Millimeter` (×10) ou transformation | Pas de `GraphicsUnit` « cm » |
| `ScaleMode = vbCharacters` | **Aucun équivalent** → mesurer le texte | À repenser (voir §6) |
| `ScaleLeft/Top` | `g.TranslateTransform(dx, dy)` | Décalage d'origine |
| `ScaleWidth/Height` (logique) | `g.ScaleTransform(sx, sy)` | Mise à l'échelle (sy<0 = Y inversé) |
| `Scale (x1,y1)-(x2,y2)` | Combinaison `Translate` + `Scale` | Voir §4 |
| `ScaleX(v, de, vers)` / `ScaleY` | Helper basé sur le **DPI** | Voir §5 |
| `Screen.TwipsPerPixelX/Y` | `1440 / Graphics.DpiX` (ou `/DpiY`) | DPI réel, pas la constante 15 |
| `ScaleWidth`/`ScaleHeight` (lire la zone) | `ClientSize.Width` / `ClientSize.Height` | En pixels |
| Souris `X, Y` (unités conteneur) | `MouseEventArgs.X/Y` (pixels) | + transformation inverse si besoin (§7) |

---

## 4. Repères personnalisés et axe inversé → transformations

Pour reproduire un `Scale (0, 100)-(100, 0)` (origine en bas-gauche, Y vers le haut, plage 0–100),  
on **compose** un décalage et une mise à l'échelle **négative** sur `Graphics` :

```vb
' Équivalent .NET du repère mathématique 0..100, Y vers le haut
Private Sub AppliquerRepereLogique(g As Graphics)
    g.TranslateTransform(0, ClientSize.Height)                 ' origine en bas
    g.ScaleTransform(ClientSize.Width / 100.0F,
                     -ClientSize.Height / 100.0F)              ' échelle + Y inversé
    ' À partir d'ici, on dessine en coordonnées 0..100 avec Y croissant vers le HAUT.
End Sub
```

> ℹ️ **L'ordre des transformations compte.** Par défaut, GDI+ les *préfixe* (`MatrixOrder.Prepend`) :
> ci-dessus, l'échelle est appliquée **avant** la translation, ce qui est bien ce que l'on veut. En
> cas de doute, raisonnez « du repère logique vers les pixels » et **vérifiez visuellement**.
> Pensez à `g.ResetTransform()` pour revenir au repère pixel.

> ⚠️ Avec une échelle ≠ 1, **les épaisseurs de trait sont mises à l'échelle elles aussi** (et une
> échelle négative peut « inverser » certains effets). Pour des traits d'épaisseur constante à
> l'écran, utilisez un `Pen` de **largeur 0**.

---

## 5. Convertir twips ↔ pixels **proprement** (selon le DPI)

La conversion ne doit **jamais** coder en dur le « 15 » : il n'est valable qu'à 96 DPI. Récupérez le  
DPI réel via `Graphics.DpiX`/`DpiY` (ou la propriété `Control.DeviceDpi`, disponible depuis .NET  
Framework 4.7) :

```vb
' twips -> pixels, dépendant du DPI réel
Private Function TwipsVersPixels(twips As Integer, g As Graphics) As Single
    Return twips * g.DpiX / 1440.0F      ' à 96 DPI : twips / 15
End Function

' pixels -> twips
Private Function PixelsVersTwips(px As Single, g As Graphics) As Integer
    Return CInt(px * 1440.0F / g.DpiX)
End Function
```

Pour mémoire, l'équivalent des fonctions `ScaleX`/`ScaleY` se construit à partir du même principe
(1 pouce = 1440 twips = 72 points = 25,4 mm = `DpiX` pixels).

| Unité | Twips | Repère |
|-------|-------|--------|
| Pixel | `1440 / DPI` (15 à 96 DPI) | Dépend du DPI |
| Point | 20 | 1/72 pouce |
| Pouce | 1440 | — |
| Millimètre | ≈ 56,7 | 1440 / 25,4 |
| Centimètre | 567 | — |

---

## 6. Le mode « caractères » (`vbCharacters`) : à reconcevoir

`ScaleMode = vbCharacters` positionnait en **largeurs/hauteurs de caractère** — pratique en VB6, sans
équivalent en GDI+. À la migration, remplacez ces positions par une **mesure de texte réelle** :

```vb
Dim taille As SizeF = g.MeasureString("M", police)   ' ou TextRenderer.MeasureText(...)
Dim x As Single = colonne * taille.Width
Dim y As Single = ligne * taille.Height
```

> ℹ️ Préférez `TextRenderer.MeasureText` si le texte est ensuite rendu avec `TextRenderer.DrawText`
> (cohérence GDI), et `Graphics.MeasureString` avec `DrawString` (GDI+). Cf. **14.1** sur les deux
> moteurs de rendu de texte.

---

## 7. Coordonnées souris : le piège du **mélange d'unités** et de la transformation

Deux écueils très concrets, souvent silencieux :

**a) Souris en pixels, dessin en twips.** En VB6, la souris (`MouseMove`, etc.) et le dessin
partageaient la même unité (twips). En .NET, la souris est en **pixels**. Si votre logique compare  
une position souris à des **constantes de dessin restées en twips**, le test devient faux. → Tout  
ramener à **une seule unité** (les pixels, idéalement).

**b) La transformation du dessin ne s'applique pas à la souris.** Une transformation posée sur
`Graphics` n'affecte **que** le dessin de cet objet `Graphics`, **pas** les coordonnées d'entrée. Si
vous dessinez dans un repère logique (via `Translate`/`Scale`), une position souris en pixels doit
être ramenée dans ce repère par la **transformation inverse** :

```vb
Private _matrice As New Matrix()   ' la transformation appliquée au dessin

Private Function SourisVersLogique(p As Point) As PointF
    Dim pts() As PointF = {New PointF(p.X, p.Y)}
    Using inv = _matrice.Clone()
        inv.Invert()
        inv.TransformPoints(pts)    ' pixels -> coordonnées logiques
    End Using
    Return pts(0)
End Function
```

---

## 8. DPI : une vigilance moderne

Les applications VB6 tournaient en général **sans conscience du DPI** (96 DPI, mise à l'échelle  
bitmap assurée par Windows sur écrans haute densité). En Windows Forms :

- les **contrôles** se mettent à l'échelle via `AutoScaleMode` (`Font`, le défaut des formulaires de
  Visual Studio, ou `Dpi`) ;
- le **dessin GDI+ personnalisé ne se met *pas* à l'échelle tout seul** : à haute densité, des
  tailles codées en pixels apparaissent trop petites. Mettez votre dessin à l'échelle par
  `DeviceDpi / 96`, ou raisonnez en unités physiques (`PageUnit`).

> ℹ️ .NET Framework 4.7+ améliore la prise en charge **par moniteur** (activation via le fichier de
> configuration de l'application). Le sujet est riche ; pour cette formation, retenez surtout que
> **votre dessin est sous votre responsabilité** en matière de DPI. Le saut vers .NET moderne (module
> 20) affine encore ce comportement.

---

## 9. 🧭 Approche recommandée pour la migration

1. **Travaillez en pixels** (le défaut .NET). C'est l'option la plus simple et la plus robuste.
2. Laissez l'**outil de migration** convertir la **mise en page du concepteur** (positions/tailles
   des contrôles) — c'est le point fort des assistants (`Upgrade Wizard`, VBUC ; cf. module 4). Le
   **risque réel** est dans le **code de dessin et de mise en page écrit à la main**.
3. **Convertissez une fois** toutes les constantes de coordonnées **en twips** présentes dans le code
   de dessin (helper du §5), au lieu de les laisser passer.
4. N'utilisez `PageUnit`/`PageScale` ou une **transformation** que si la logique d'origine dépendait
   vraiment d'unités **physiques** (impression en mm, cf. 14.4) ou d'un **repère personnalisé** (axe
   inversé, plage logique). Dans ces cas, une transformation est **plus propre** que de convertir
   chaque nombre à la main.

---

## 10. ⚠️ Pièges silencieux à retenir

- **Constantes en twips** laissées dans le dessin → rendu ~15× trop petit, tassé en haut-gauche, ou
  hors champ.
- **Coder en dur « 15 » twips/pixel** → faux dès que le DPI n'est pas 96. Utiliser `Graphics.DpiX`.
- **Souris en pixels mélangée à des constantes en twips** → tests de position faux.
- **Transformation du dessin non appliquée à la souris** → il faut la transformation **inverse**.
- **`vbCharacters` / `vbUser`** simplement « recopiés » en nombres → repère perdu (mesurer le texte,
  poser une transformation).
- **Épaisseur de trait** sous `PageUnit` non-pixel ou échelle ≠ 1 → traits trop épais ; préférer un
  `Pen` de **largeur 0**.
- **Position des contrôles** : l'unité du conteneur comptait en VB6 ; en .NET c'est **toujours des
  pixels**.
- **Haute densité (DPI)** : le dessin personnalisé **ne se met pas à l'échelle** automatiquement.

> 🔗 Voir l'**Annexe B** (catalogue des pièges silencieux) pour les conversions, et l'**Annexe D**
> pour les correspondances de types impliquées dans les calculs de coordonnées.

---

**Section suivante → 14.3 — Images, icônes et ressources (`.frx`, `.res` → `.resx`)**, pour récupérer
les ressources embarquées des formulaires et y accéder « à la .NET ».

⏭️ [Images, icônes et ressources (`.frx`, `.res` → `.resx`)](/14-graphismes-impression-activex/03-ressources.md)
