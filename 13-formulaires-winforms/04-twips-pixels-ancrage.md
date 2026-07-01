🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.4 — Propriétés, ancrage et positionnement : Twips → pixels (conversion d'unités) ⚠️

> **Chapitre 13 — Des formulaires VB6 à Windows Forms** · Section 13.4
> Le piège de géométrie **le plus diffus** du chapitre : VB6 mesure en **twips**, Windows Forms en **pixels** — et le danger se cache surtout dans le **code d'exécution**.

---

## 🧭 Périmètre

Deux sujets : le **changement d'unité** (twips → pixels), qui touche **toutes** les tailles et positions, et le **modèle de disposition** (`Anchor`/`Dock`), qui **remplace** le redimensionnement manuel de VB6. L'outillage convertit la **disposition de conception**, mais le **code** qui manipule la géométrie réclame une **revue attentive**.

---

## 1. Deux unités, deux mondes

VB6 mesure par défaut en **twips** (1 twip = 1/1440 de pouce). Windows Forms mesure **toujours** en **pixels**. À la résolution standard de **96 ppp** :

| Unité | En twips | En pixels (96 ppp) |
|-------|----------|---------------------|
| 1 pouce | 1440 | 96 |
| 1 point | 20 | 1,33 |
| **1 pixel** | **15** | **1** |

La règle d'or : **1 pixel = 15 twips** (à 96 ppp). Pour convertir une valeur VB6 : **pixels = twips ÷ 15**.

> 💡 **Nuance `ScaleMode`.** En VB6, l'unité par défaut est le twip, mais un formulaire pouvait fixer `ScaleMode = vbPixels` (3). Dans ce cas, ses valeurs étaient **déjà** en pixels et **ne nécessitent aucune conversion**. La règle de conversion **dépend donc du `ScaleMode` d'origine** — méfiez-vous des projets à `ScaleMode` **mixtes**.

---

## 2. La conversion à la migration (outillée, mais à surveiller)

L'assistant de mise à niveau **convertit les valeurs de conception** : les tailles/positions des blocs `.frm` (en twips) deviennent des valeurs en **pixels** dans `InitializeComponent` (rappel **[13.1](01-modele-formulaire.md)**).

Correspondance des propriétés :

| VB6 (twips par défaut) | Windows Forms (pixels) |
|------------------------|------------------------|
| `Top`, `Left` | `Top`, `Left` / `Location` (`Point`) |
| `Height`, `Width` | `Height`, `Width` / `Size` (`Size`) |
| `Move l, t, w, h` | `SetBounds(x, y, w, h)` |
| `ScaleHeight`, `ScaleWidth` | `ClientSize.Height` / `ClientSize.Width` |
| `ScaleMode` | *(aucun : pixels)* |
| `ScaleX`/`ScaleY`, `TwipsPerPixelX/Y` | *(aucun : à réécrire)* |

Deux limites de la conversion automatique :

- ⚠️ **Arrondis et dérive.** `twips ÷ 15` n'est pas toujours entier : l'arrondi provoque de **légers décalages** de pixels. Les dispositions denses peuvent demander une **retouche manuelle**.
- ⚠️ **Hypothèse 96 ppp.** La conversion suppose 96 ppp ; en **haute résolution**, le rendu peut différer (voir §5).

### ⚠️ Le vrai piège : les littéraux de twips **dans le code**

L'assistant convertit la disposition **de conception**, mais il peut **manquer** le code qui **manipule la géométrie à l'exécution** avec des **constantes en twips**. Ces valeurs deviennent alors des **pixels** — un **bug silencieux** spectaculaire.

```vb
' VB6 — 1500 twips ≈ 100 pixels
Text1.Width = 1500
```

```vb
' VB.NET — ⚠️ 1500 PIXELS (≈ 15× trop large !)
TextBox1.Width = 1500        ' littéral en twips non converti = bug silencieux
' Correct (si l'intention était ~100 px) :
TextBox1.Width = 100
```

Les fonctions de conversion VB6 (`ScaleX`, `ScaleY`, `TwipsPerPixelX/Y`) **n'ont pas d'équivalent** : tout code qui s'en sert doit être **réécrit** en raisonnant directement en pixels.

---

## 3. Le modèle de disposition : `Anchor` et `Dock` (un vrai progrès)

VB6 **n'avait aucune** disposition automatique : pour qu'un formulaire se redimensionne proprement, on écrivait du code dans **`Form_Resize`** qui **repositionnait/redimensionnait** les contrôles à la main (beaucoup d'arithmétique, en twips).

Windows Forms introduit deux propriétés qui **automatisent** cela :

**`Anchor`** — ancre les **bords** d'un contrôle à ceux de son conteneur :

| `Anchor` | Effet au redimensionnement |
|----------|----------------------------|
| `Top, Left` (défaut) | position **fixe** (coin haut-gauche) |
| `Top, Left, Right` | s'**étire horizontalement** |
| `Top, Bottom, Left, Right` | s'étire **dans les deux sens** |
| `Bottom, Right` | reste **collé** au coin bas-droit |

**`Dock`** — colle un contrôle à un **bord** (ou remplit l'espace) :

| `Dock` | Effet |
|--------|-------|
| `Top` / `Bottom` | collé en haut/bas, pleine largeur (barres) |
| `Left` / `Right` | collé à gauche/droite, pleine hauteur |
| `Fill` | occupe **tout** l'espace restant |
| `None` (défaut) | position libre (`Anchor` s'applique) |

Pour des dispositions plus riches, des **conteneurs de mise en page** existent : `Panel`, `GroupBox`, `TableLayoutPanel`, `FlowLayoutPanel`, `SplitContainer`.

### Migration : conserver le `Resize`, ou adopter `Anchor`/`Dock` ?

```vb
' VB6 — redimensionnement MANUEL dans Form_Resize (marges en twips)
Private Sub Form_Resize()
    Text1.Width = Me.ScaleWidth - 240
End Sub
```

- **Option fidèle** : garder un gestionnaire `Resize`, mais **convertir les unités** (raisonner en pixels via `ClientSize`).
- **Option idiomatique** (souvent préférable) : **supprimer** le code `Resize` et poser un **`Anchor`** dans le concepteur :

```vb
' VB.NET — TextBox1.Anchor = Top, Left, Right
'   → s'étire tout seul avec le formulaire ; plus AUCUN code Resize
```

> 💡 `Anchor`/`Dock` sont si simples et si **fiables** (ils éliminent une arithmétique source d'erreurs) que leur adoption est **raisonnable dès la migration** pour les cas simples. Pour des dispositions complexes, convertissez d'abord, refactorez ensuite (**17.5**).

---

## 4. Positionnement du formulaire et autres propriétés

- **Centrage** : l'arithmétique manuelle de VB6 devient **une propriété**.

```vb
' VB6 — centrer à la main (twips)
Me.Left = (Screen.Width - Me.Width) \ 2
Me.Top = (Screen.Height - Me.Height) \ 2
```

```vb
' VB.NET — une propriété suffit
Me.StartPosition = FormStartPosition.CenterScreen
```

- **État de la fenêtre** : `WindowState` (vbNormal/vbMinimized/vbMaximized) → `WindowState` de type `FormWindowState` (`Normal`/`Minimized`/`Maximized`) — correspondance directe.
- **Zone client** : `ScaleHeight`/`ScaleWidth` → **`ClientSize`** (en pixels).

---

## 5. Un mot sur la haute résolution (DPI)

Les twips visaient l'**indépendance vis-à-vis de la résolution** : un formulaire conservait sa taille **physique** quelle que soit la définition de l'écran. Windows Forms atteint un objectif voisin **autrement**, via la propriété **`AutoScaleMode`** :

- le concepteur règle par défaut **`AutoScaleMode = Font`** (mise à l'échelle selon la police), ce qui **approxime** l'ancien comportement ;
- en **haute résolution (high-DPI)**, un formulaire migré naïvement peut **mal s'afficher** ; la mise à l'échelle DPI est un **sujet à part**, à traiter au cas par cas après la migration de base.

> 💡 Retenez surtout que la **promesse** (indépendance de résolution) survit, mais que le **mécanisme** change : ce ne sont plus les twips, c'est **`AutoScaleMode`**.

---

## 🔁 Récapitulatif VB6 → VB.NET

| VB6 | Windows Forms |
|-----|---------------|
| **twips** (défaut) | **pixels** (toujours) — **1 px = 15 twips** (96 ppp) |
| `Move l, t, w, h` | `SetBounds(x, y, w, h)` |
| `ScaleHeight`/`ScaleWidth` | `ClientSize` |
| `ScaleX`/`ScaleY`, `TwipsPerPixel…` | *(à réécrire en pixels)* |
| littéraux en twips dans le code | ⚠️ deviennent des **pixels** → convertir |
| `Form_Resize` manuel | **`Anchor`** / **`Dock`** (souvent : supprimer le code) |
| centrage arithmétique | `StartPosition = CenterScreen` |
| indépendance via twips | **`AutoScaleMode`** (Font par défaut) |

---

## ✅ Points clés

- VB6 mesure en **twips**, WinForms en **pixels** : **1 px = 15 twips** à 96 ppp (mais **0 conversion** si le `ScaleMode` d'origine était déjà `vbPixels`).
- L'assistant convertit la **disposition de conception** ; ⚠️ il peut **manquer** les **littéraux de twips dans le code** (ex. `ctrl.Width = 1500`), qui deviennent des **pixels**.
- `ScaleX`/`ScaleY`/`TwipsPerPixel…` **n'ont pas d'équivalent** : raisonner directement en pixels.
- **`Anchor`/`Dock`** remplacent le redimensionnement manuel de `Form_Resize` — un **progrès** à adopter pour les cas simples.
- Le **centrage** devient `StartPosition` ; l'**indépendance de résolution** passe par **`AutoScaleMode`**, pas par les twips.

---

## 🔗 Renvois

- **[13.1](01-modele-formulaire.md)** — `InitializeComponent` (où atterrissent les valeurs converties). · **[13.2](02-cycle-de-vie.md)** — `Resize`/cycle de vie.
- **[13.5](05-correspondance-controles.md)** — propriétés des contrôles (au-delà de la géométrie). · **[14.2](../14-graphismes-impression-activex/README.md)** — coordonnées et `ScaleMode` côté **dessin**.
- **[17.5](../17-valider-refactoriser/README.md)** — refactoring (passage complet à `Anchor`/`Dock`).
- **Suite → [13.5 — Correspondance des contrôles intrinsèques (`CommandButton`→`Button`, `TextBox`, `Label`, `Frame`, `OptionButton`…)](05-correspondance-controles.md)**

⏭️ [Correspondance des contrôles intrinsèques (`CommandButton`→`Button`, `TextBox`, `Label`, `Frame`, `OptionButton`…)](/13-formulaires-winforms/05-correspondance-controles.md)
