🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.2 — Cycle de vie : `Load`/`Unload`/`Activate` → `Load`/`FormClosing`/`Shown` ; `Show` vs `ShowDialog` ⚠️

> **Chapitre 13 — Des formulaires VB6 à Windows Forms** · Section 13.2
> Les événements de cycle de vie **ne se correspondent pas un pour un**, et le comportement de **fermeture** change — au point de **casser** un idiome VB6 courant.

---

## 🧭 Périmètre

Deux sujets liés : la **séquence d'événements** d'un formulaire (de sa création à sa destruction), et la façon de l'**afficher** (`Show` vs `ShowDialog`). Les deux recèlent des **changements silencieux** : un événement qui se déclenche **à un autre moment**, une fermeture qui **détruit** l'objet, un résultat modal qui se lit **autrement**.

---

## 1. Le cycle de vie, de VB6 à Windows Forms

L'ordre des événements au premier affichage :

```
VB6      :  Initialize → Load → Activate → … → QueryUnload → Unload → Terminate
WinForms :  New        → Load → Activated → Shown → …
                                         … → FormClosing → FormClosed → Disposed
```

Correspondance :

| VB6 | Windows Forms | Remarque |
|-----|---------------|----------|
| `Form_Initialize` | `Sub New` (constructeur) | création (exécute `InitializeComponent`) |
| `Form_Load` | `Load` | **avant** l'affichage |
| `Form_Activate` | `Activated` | à **chaque** prise de focus (se **répète**) |
| *(aucun équivalent)* | **`Shown`** | **après** le premier affichage, **une seule fois** (après `Activated`) |
| `Form_Deactivate` | `Deactivate` | perte de focus |
| `Form_QueryUnload` | **`FormClosing`** | **annulable** (`Cancel`, `CloseReason`) |
| `Form_Unload` | `FormClosed` | **après** fermeture (non annulable) |
| `Form_Terminate` | `Dispose` / `Disposed` | libération (voir **[12.3](../12-poo/03-idisposable-using.md)**) |

---

## 2. Les correspondances qui piègent ⚠️

### 2.1 `Load` vs `Shown` : avant ou après l'affichage

En VB6, on plaçait souvent dans `Load` (ou dans `Activate`) du code censé s'exécuter « **une fois le formulaire visible** ». Or :

- **`Load`** se déclenche **avant** que le formulaire soit affiché ;
- **`Activate`** (→ `Activated`) se **répète** à chaque prise de focus.

Windows Forms fournit le bon point d'accroche : **`Shown`**, qui survient **après** le premier affichage et **une seule fois**. C'est là que doit aller la logique « après affichage » (par ex. ouvrir un message **par-dessus** le formulaire désormais visible).

> ⚠️ Mettre du code « post-affichage » dans `Load` peut échouer (le formulaire n'est **pas encore** à l'écran). Déplacez-le dans **`Shown`**.

### 2.2 `Activate` → `Activated` : un événement qui se répète

`Activated` se déclenche **à chaque fois** que le formulaire **regagne le focus** — comme `Form_Activate` en VB6. Si votre code en dépendait pour ré-initialiser à chaque réactivation, la correspondance est directe. Mais pour une action **« une fois, au premier affichage »**, utilisez **`Shown`**, pas `Activated`.

### 2.3 `QueryUnload` → `FormClosing` ; `Unload` → `FormClosed`

Le crochet **« voulez-vous vraiment fermer ? »** de VB6 (`QueryUnload`, avec `Cancel` et `UnloadMode`) devient **`FormClosing`** (avec `Cancel` et `CloseReason`).

```vb
' VB.NET — confirmer/annuler la fermeture
Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    If MessageBox.Show("Quitter sans enregistrer ?", "Confirmation",
                       MessageBoxButtons.YesNo) = DialogResult.No Then
        e.Cancel = True          ' annule la fermeture
    End If
End Sub
```

> ⚠️ **L'annulation doit être dans `FormClosing`.** En VB6, `Unload` aussi possédait un `Cancel` ; en WinForms, **`FormClosed` est trop tard** (la fermeture a eu lieu, on ne peut plus l'annuler). Tout code d'annulation migre vers **`FormClosing`**. Le `CloseReason` indique la **cause** (utilisateur, code, arrêt de Windows, parent MDI…), avec des valeurs d'énumération **différentes** de l'`UnloadMode` de VB6.

---

## 3. ⚠️ Le piège majeur : `Hide` vs `Close` (et la **disposition**)

C'est le changement de comportement **le plus dangereux** de cette section.

En **VB6** :
- `Hide` **masque** le formulaire mais le **garde en mémoire** (avec son état) → on peut le **réafficher** sans relancer `Load` ;
- `Unload` le **retire** de la mémoire ; mais via l'instance par défaut, un **nouvel** accès le **recharge**.

En **Windows Forms** :
- `Hide()` (ou `Visible = False`) **garde** le formulaire **réutilisable** ;
- **`Close()`** ferme **et DÉTRUIT** (dispose) un formulaire **non modal** → on ne peut **plus** le réafficher.

```vb
' VB6 — masquer puis réafficher la MÊME instance
Form2.Hide
' …
Form2.Show           ' réaffiche la même instance (Load NON re-déclenché)
```

```vb
' VB.NET — ⚠️ Close() DISPOSE le formulaire
_form2.Close()       ' (non modal) → disposé
_form2.Show()        ' ❌ ObjectDisposedException

' Pour réafficher la même instance : utiliser Hide()
_form2.Hide()
' …
_form2.Show()        ' ✅ réaffiche la même instance
```

> ⚠️ L'idiome VB6 **« `Unload` puis ré-afficher »** **casse** en WinForms si l'on emploie `Close()` : le formulaire est **disposé**. Deux solutions : **`Hide()`** pour conserver l'instance, ou **ré-instancier** (`New`) si l'on veut repartir d'un état neuf.

---

## 4. `Show` vs `ShowDialog` : le modèle modal

En **VB6**, `Show` portait un argument de modalité : `Show vbModal` (bloquant) ou `Show vbModeless` / `Show` (non bloquant). En **Windows Forms**, ce sont **deux méthodes** :

| VB6 | Windows Forms |
|-----|---------------|
| `Show vbModeless` / `Show` | `Show()` — non modal |
| `Show vbModal` | `ShowDialog()` — modal, **retourne un `DialogResult`** |
| `Hide` | `Hide()` / `Visible = False` (réutilisable) |
| `Unload Form1` | `Close()` (⚠️ dispose en non modal) |
| drapeau public pour le résultat | **`DialogResult`** + lecture des propriétés |

La grande différence : **`ShowDialog()` renvoie un résultat** (`DialogResult.OK`, `Cancel`, `Yes`…). On n'a plus besoin du **drapeau public** que VB6 imposait pour communiquer le choix de l'utilisateur. Un bouton peut même **fixer** `DialogResult` (propriété du bouton dans le concepteur), ce qui **ferme** le dialogue automatiquement.

```vb
' VB6 — modal + variable publique sur le formulaire
FormSaisie.Show vbModal
If FormSaisie.OK Then              ' "OK" = propriété/variable publique du formulaire
    valeur = FormSaisie.txtNom.Text
End If
Unload FormSaisie
```

```vb
' VB.NET — modal via ShowDialog + DialogResult
Using dlg As New FormSaisie()
    If dlg.ShowDialog() = DialogResult.OK Then
        Dim valeur = dlg.txtNom.Text   ' lisible : pas encore disposé (on est dans le Using)
    End If
End Using                              ' Dispose() ici, déterministe (voir 12.3)
```

> ⚠️ **`ShowDialog` ne dispose PAS automatiquement** le formulaire (contrairement à `Close()` sur un non-modal). C'est **voulu** : on peut **lire ses propriétés après** sa fermeture. En contrepartie, **vous** devez le **libérer** — d'où le **`Using`** ci-dessus, qui appelle `Dispose` à la sortie du bloc.

---

## 5. `Form_Initialize` et `Form_Terminate`

- **`Form_Initialize`** → le **constructeur `Sub New`** (qui exécute `InitializeComponent`). Rarement migré tel quel.
- **`Form_Terminate`** → la **disposition** du formulaire. Les formulaires sont **`IDisposable`** : le concepteur **génère** un `Protected Overrides Sub Dispose(disposing As Boolean)` (qui libère les composants). Tout nettoyage de `Form_Terminate` migre donc vers ce `Dispose` (ou vers `FormClosed`) — exactement la logique de la **[12.3](../12-poo/03-idisposable-using.md)**.

---

## 🔁 Récapitulatif VB6 → VB.NET

| Aspect | VB6 | Windows Forms |
|--------|-----|---------------|
| Avant affichage | `Load` | `Load` |
| Après premier affichage | *(via `Activate`, bancal)* | **`Shown`** (une fois) |
| Réactivation | `Activate` | `Activated` (se répète) |
| Confirmer la fermeture | `QueryUnload` (`Cancel`) | **`FormClosing`** (`Cancel`, `CloseReason`) |
| Après fermeture | `Unload` | `FormClosed` (non annulable) |
| Masquer / réafficher | `Hide` / `Show` | `Hide()` / `Show()` |
| Fermer (détruire) | `Unload` | `Close()` (⚠️ dispose en non modal) |
| Modal | `Show vbModal` | `ShowDialog()` → `DialogResult` |
| Nettoyage final | `Terminate` | `Dispose` (généré) / `Disposed` |

---

## ✅ Points clés

- Les événements **ne se correspondent pas un pour un** : `Load` reste, mais **`Shown`** (après affichage, une fois) **remplace** l'usage bancal de `Activate`.
- L'**annulation** de fermeture migre vers **`FormClosing`** (`FormClosed` est trop tard).
- ⚠️ **`Close()` dispose** un formulaire non modal : l'idiome **« `Unload` puis ré-afficher »** casse → utiliser **`Hide()`** ou ré-instancier.
- `Show vbModal` → **`ShowDialog()`**, qui **retourne un `DialogResult`** : fini le **drapeau public** pour transmettre le choix.
- ⚠️ **`ShowDialog` ne dispose pas** automatiquement : lisez les propriétés **après**, puis **`Dispose`** (via **`Using`**).
- `Form_Terminate` → le **`Dispose` généré** du formulaire (rappel 12.3).

---

## 🔗 Renvois

- **[13.1](01-modele-formulaire.md)** — instance par défaut (`My.Forms`), `Close`/réaffichage. · **[13.3](03-formulaires-mdi.md)** — cycle de vie des MDI.
- **[12.3](../12-poo/03-idisposable-using.md)** — `IDisposable` / `Using` / `Dispose` (fermeture et libération).
- **[13.9](09-msgbox-doevents.md)** — `MsgBox`/`DoEvents` (interactions pendant le cycle de vie).
- **Suite → [13.3 — Formulaires MDI : `MDIForm`/`MDIChild` → `IsMdiContainer`/`MdiParent`, `LayoutMdi` et fusion de menus](03-formulaires-mdi.md)** ⚠️

⏭️ [Formulaires **MDI** : `MDIForm`/`MDIChild` → `IsMdiContainer`/`MdiParent`, `LayoutMdi` et fusion de menus](/13-formulaires-winforms/03-formulaires-mdi.md)
