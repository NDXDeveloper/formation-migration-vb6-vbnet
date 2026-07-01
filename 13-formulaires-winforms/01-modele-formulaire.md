🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.1 — Le modèle de formulaire VB6 (`.frm`/`.frx`) → Windows Forms (`.vb` + Designer)

> **Chapitre 13 — Des formulaires VB6 à Windows Forms** · Section 13.1
> L'**anatomie** du formulaire : les fichiers, la **classe partielle**, le **Designer** et les **ressources**. Le pendant « conteneur » de la [12.1](../12-poo/01-classes.md), côté interface.

---

## 🧭 Périmètre

On s'intéresse ici à la **structure** d'un formulaire — comment un `.frm` (+ `.frx`) devient un trio `.vb` / `.Designer.vb` / `.resx`. On laisse de côté :

- le **cycle de vie** (`Load`/`Unload`…) → **[13.2](02-cycle-de-vie.md)** ;
- la **géométrie** (twips → pixels) → **[13.4](04-twips-pixels-ancrage.md)** ;
- la **correspondance des contrôles** → **[13.5](05-correspondance-controles.md)**.

> 💡 Comme pour les classes (12.1), la traduction du **conteneur** est **largement outillée**. L'enjeu est de **comprendre la nouvelle anatomie** pour savoir où va quoi — et où **ne pas** intervenir à la main.

---

## 1. Le formulaire VB6 : un fichier `.frm` (+ `.frx`)

En VB6, **un** formulaire = **un** fichier `.frm`, un fichier **texte** qui **mêle** trois choses :

1. la **définition visuelle** — un bloc `Begin VB.Form … End` décrivant le formulaire **et tous ses contrôles** avec leurs propriétés ;
2. des **attributs** (dont `VB_PredeclaredId`) ;
3. le **code** (gestionnaires d'événements, méthodes).

```vb
' VB6 — Form1.frm (extrait)
VERSION 5.00
Begin VB.Form Form1
   Caption         =   "Bonjour"
   ClientHeight    =   3000
   ClientWidth     =   4500
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Height          =   495          ' ← twips (voir 13.4)
      Left            =   1680
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtNom
      Height          =   285
      Left            =   1200
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_PredeclaredId = True       ' ← instance par défaut globale
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    MsgBox "Bonjour " & txtNom.Text
End Sub
```

Deux éléments à retenir :

- le fichier **`.frx`** (compagnon **binaire**) stocke les ressources non textuelles (images, icônes), **référencées par décalage** depuis le `.frm` (ex. `Picture = "Form1.frx":0000`) ;
- **`VB_PredeclaredId = True`** dote le formulaire d'une **instance par défaut globale** : `Form1.Show` fonctionne **sans `New`** (voir §4).

---

## 2. Le formulaire Windows Forms : une **classe partielle** + Designer

En VB.NET, le même formulaire est **réparti** sur des fichiers, grâce aux **classes partielles** :

| Fichier | Contenu | Qui le maintient |
|---------|---------|------------------|
| **`Form1.vb`** | votre **code-behind** (handlers, logique) | **vous** |
| **`Form1.Designer.vb`** | `Inherits …Form`, `InitializeComponent`, déclaration des contrôles | le **concepteur** |
| **`Form1.resx`** | ressources (images, icônes, chaînes) | le concepteur |

**Le code-behind** ne contient que votre logique :

```vb
' VB.NET — Form1.vb (code-behind)
Public Class Form1

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        MsgBox("Bonjour " & txtNom.Text)
    End Sub

End Class
```

**Le fichier Designer** porte l'héritage, la création des contrôles et leur câblage. Détail clé : la **disposition visuelle devient du code**, dans `InitializeComponent` — on passe d'une description **déclarative** (les blocs `.frm`) à du code **impératif**.

```vb
' VB.NET — Form1.Designer.vb (extrait, simplifié)
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    ' NE PAS modifier à la main : régénéré par le concepteur
    Private Sub InitializeComponent()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.txtNom = New System.Windows.Forms.TextBox()
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(112, 120)   ' ← pixels (voir 13.4)
        Me.cmdOK.Size = New System.Drawing.Size(81, 33)
        Me.cmdOK.Text = "OK"
        '
        'Form1
        '
        Me.Text = "Bonjour"                                      ' Caption → Text
        Me.Controls.Add(Me.txtNom)
        Me.Controls.Add(Me.cmdOK)
    End Sub

    Friend WithEvents cmdOK As System.Windows.Forms.Button       ' WithEvents → Handles
    Friend WithEvents txtNom As System.Windows.Forms.TextBox
End Class
```

Trois points structurants :

- **`Partial`** permet aux deux moitiés (`Form1.vb` et `Form1.Designer.vb`) de former **une seule** classe `Form1` ; il suffit qu'**une** déclaration porte le mot-clé.
- Les contrôles sont déclarés **`Friend WithEvents`** : c'est ce `WithEvents` qui rend possible la clause **`Handles`** du code-behind (rappel **[12.8](../12-poo/08-evenements.md)**).
- Le formulaire **`Inherits System.Windows.Forms.Form`** — l'héritage entre enfin en scène (rappel **[12.5](../12-poo/05-interfaces-heritage.md)**).

---

## 3. La traduction : du déclaratif à l'impératif (outillé)

L'assistant de mise à niveau **analyse les blocs de propriétés** du `.frm` et **génère** : la classe partielle Designer, le code `InitializeComponent` équivalent, et **extrait** les ressources du `.frx` vers le `.resx`.

| VB6 | Windows Forms |
|-----|---------------|
| `.frm` — blocs de propriétés / contrôles | `Form1.Designer.vb` → `InitializeComponent()` |
| `.frm` — code | `Form1.vb` (code-behind) |
| `.frx` — binaire | `Form1.resx` — XML |
| `Attribute VB_Name = "Form1"` | nom de la classe partielle |
| `VB_PredeclaredId = True` | instance par défaut via `My.Forms` (§4) |
| contrôle dans un bloc `Begin … End` | `Friend WithEvents …` **+** création dans `InitializeComponent` |

> ⚠️ **Ne modifiez pas le code du Designer à la main.** `InitializeComponent` est **régénéré** par le concepteur ; une édition manuelle peut être **écrasée** ou **casser** le concepteur. Votre code va dans le **code-behind** (`Form1.vb`). C'est une **règle d'hygiène** essentielle de la migration.

---

## 4. L'instance par défaut : `Form1.Show` sans `New`

En **VB6**, l'instance par défaut globale (via `VB_PredeclaredId`) permet `Form1.Show`, `Form1.Caption = …`, etc., **sans instancier**.

En **VB.NET**, ce confort est **préservé** — mais par un mécanisme **différent** : l'**instance par défaut** exposée par **`My.Forms`**. `Form1.Show()` reste donc possible. À connaître, toutefois :

- l'instance par défaut est **par thread** (chaque thread a la sienne) ;
- son **comportement diffère subtilement** de VB6 (moment de création ; refermer puis ré-accéder **recrée** une nouvelle instance) ;
- en .NET **idiomatique**, on lui préfère l'**instanciation explicite** :

```vb
' VB.NET — instanciation explicite (recommandée à terme)
Dim f As New Form1()
f.Show()                 ' (ou f.ShowDialog() pour du modal — voir 13.2)
```

> 💡 Rappel de **[12.1](../12-poo/01-classes.md)** : pour une **classe ordinaire**, `VB_PredeclaredId` **n'a pas** d'équivalent et impose un refactoring. Les **formulaires** sont l'**exception** : ils récupèrent une instance par défaut via `My.Forms`. L'assistant garde donc le style `Form1.Show` fonctionnel.

---

## 5. Détails du conteneur à connaître

- **`Caption` → `Text`.** La propriété de titre du formulaire (et de nombreux contrôles) est **renommée** `Text`. La correspondance complète des contrôles est en **[13.5](05-correspondance-controles.md)**.
- **`Me`** désigne l'instance du formulaire — concept inchangé.
- **Démarrage** : le **formulaire de démarrage** se définit dans les propriétés du projet ; sous le capot, `Application.Run(New Form1())` lance la boucle de messages. On peut aussi démarrer par **`Sub Main`** dans un `Module` (voir **[12.4](../12-poo/04-modules-globales.md)**).
- **Ressources** : le passage `.frx` → `.resx` est traité plus en détail (images, icônes) en **[14.3](../14-graphismes-impression-activex/README.md)**.
- **Unités** : les valeurs de taille/position changent d'**échelle** (twips → pixels) — sujet entier de la **[13.4](04-twips-pixels-ancrage.md)**.

---

## 🔁 Récapitulatif VB6 → VB.NET

| VB6 | Windows Forms |
|-----|---------------|
| `Form1.frm` (visuel **+** code, **un** fichier) | `Form1.vb` **+** `Form1.Designer.vb` (**séparés**, classe **partielle**) |
| `Form1.frx` (binaire) | `Form1.resx` (XML) |
| blocs `Begin … End` (déclaratif) | `InitializeComponent()` (**impératif**, généré) |
| formulaire = classe à **instance par défaut** | classe qui **`Inherits …Form`** ; instance par défaut via **`My.Forms`** |
| `Caption` | `Text` |
| contrôle déclaré dans le `.frm` | `Friend WithEvents …` (câblé par `Handles`) |

---

## ✅ Points clés

- Un `.frm` **unique** (visuel + code) éclate en une **classe partielle** : **`Form1.vb`** (votre code) **+** **`Form1.Designer.vb`** (généré) **+** **`Form1.resx`**.
- La disposition passe du **déclaratif** (blocs `.frm`) à l'**impératif** (`InitializeComponent`), **maintenu par le concepteur** — **à ne pas éditer à la main**.
- Le formulaire **`Inherits …Form`** et déclare ses contrôles en **`Friend WithEvents`** (d'où les clauses `Handles`).
- L'**instance par défaut** (`Form1.Show` sans `New`) est **préservée** via **`My.Forms`**, mais **par thread** et au comportement **subtilement différent** — l'instanciation explicite est préférable à terme.
- La conversion est **outillée** ; restez attentif aux **ressources**, à la **géométrie** et aux **règles d'édition** du Designer.

---

## 🔗 Renvois

- **[12.1](../12-poo/01-classes.md)** — anatomie d'une classe (et `VB_PredeclaredId`). · **[12.5](../12-poo/05-interfaces-heritage.md)** — héritage (`Inherits …Form`). · **[12.8](../12-poo/08-evenements.md)** — `WithEvents`/`Handles`.
- **[12.4](../12-poo/04-modules-globales.md)** — `Sub Main` (démarrage).
- **[13.4](04-twips-pixels-ancrage.md)** — twips → pixels. · **[13.5](05-correspondance-controles.md)** — contrôles (`Caption`→`Text`). · **[14.3](../14-graphismes-impression-activex/README.md)** — ressources `.frx`/`.res` → `.resx`.
- **Suite → [13.2 — Cycle de vie : `Load`/`Unload`/`Activate` → `Load`/`FormClosing`/`Shown` ; `Show` vs `ShowDialog`](02-cycle-de-vie.md)** ⚠️

⏭️ [Cycle de vie : `Load`/`Unload`/`Activate` → `Load`/`FormClosing`/`Shown` ; `Show` vs `ShowDialog`](/13-formulaires-winforms/02-cycle-de-vie.md)
