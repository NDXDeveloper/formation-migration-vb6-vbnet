🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.8 — Boîtes de dialogue communes (`CommonDialog` → `OpenFileDialog`, `ColorDialog`…)

> **Chapitre 13 — Des formulaires VB6 à Windows Forms** · Section 13.8
> Un **contrôle multiplexé** (`CommonDialog`, un ActiveX) éclate en **plusieurs composants natifs** — et la détection de l'annulation change de **mécanisme**.

---

## 🧭 Périmètre

VB6 expose **une seule** `CommonDialog` (issue de l'OCX `comdlg32.ocx`) qui donne accès à **plusieurs** boîtes de dialogue Windows via des méthodes différentes. Windows Forms fournit, à la place, **un composant natif par dialogue**. Deux conséquences : la **répartition** en plusieurs composants, et le passage de l'idiome `CancelError`/`On Error` au modèle **`DialogResult`**.

---

## 1. Un contrôle multiplexé → plusieurs composants

En **VB6**, on pose **une** `CommonDialog` et l'on appelle la méthode correspondant au dialogue voulu : `ShowOpen`, `ShowSave`, `ShowColor`, `ShowFont`, `ShowPrinter`.

En **Windows Forms**, chaque dialogue est un **composant distinct** (tous dérivent de la classe `CommonDialog` et exposent `ShowDialog()` retournant un `DialogResult`) :

| VB6 (`CommonDialog`) | Windows Forms |
|----------------------|---------------|
| `ShowOpen` | **`OpenFileDialog`** |
| `ShowSave` | **`SaveFileDialog`** |
| `ShowColor` | **`ColorDialog`** |
| `ShowFont` | **`FontDialog`** |
| `ShowPrinter` | **`PrintDialog`** |
| *(aucun équivalent)* | **`FolderBrowserDialog`** (choix de dossier) |
| *(aucun équivalent)* | `PrintPreviewDialog`, `PageSetupDialog` |

> ⚠️ **OCX → natif.** `CommonDialog` est un **ActiveX** ; les dialogues Windows Forms sont **natifs**. Comme pour la `Toolbar` (**[13.7](07-menus-toolbars.md)**), on **remplace** plutôt qu'on interopère — ce qui supprime une **dépendance ActiveX**. L'assistant **scinde** chaque appel `ShowXxx` vers le bon composant.

---

## 2. Le changement de pattern : `CancelError`/`On Error` → `DialogResult`

C'est **le** changement de code à retenir. En **VB6**, on détectait l'annulation en mettant **`CancelError = True`** puis en interceptant l'**erreur** levée à l'appui sur *Annuler* (`On Error`). En **Windows Forms**, **`ShowDialog()` retourne un `DialogResult`** que l'on **teste** simplement.

```vb
' VB6 — Annuler détecté via CancelError + On Error
On Error GoTo Annule
CommonDialog1.CancelError = True
CommonDialog1.Filter = "Texte (*.txt)|*.txt|Tous (*.*)|*.*"
CommonDialog1.ShowOpen
OuvrirFichier CommonDialog1.FileName
Exit Sub
Annule:
    ' l'utilisateur a annulé
```

```vb
' VB.NET — Annuler détecté via DialogResult
Using dlg As New OpenFileDialog()
    dlg.Filter = "Texte (*.txt)|*.txt|Tous (*.*)|*.*"   ' même format de filtre
    If dlg.ShowDialog() = DialogResult.OK Then
        OuvrirFichier(dlg.FileName)
    End If
End Using                                                ' Dispose (voir 12.3)
```

> 💡 Plus de `On Error` pour gérer l'annulation : c'est plus **lisible** et cohérent avec le reste de Windows Forms (rappel de `ShowDialog`/`DialogResult` en **[13.2](02-cycle-de-vie.md)**).

---

## 3. Propriétés : ce qui se transpose

| VB6 | Windows Forms |
|-----|---------------|
| `FileName` | `FileName` (et `FileNames` si `Multiselect`) |
| `Filter` | `Filter` — **même format** : `"Desc (*.ext)\|*.ext\|…"` |
| `FilterIndex` | `FilterIndex` (1-based dans les deux) |
| `InitDir` | **`InitialDirectory`** |
| `DialogTitle` | **`Title`** |
| `DefaultExt` | `DefaultExt` |
| `Color` (ShowColor) | `Color` — ⚠️ **type** différent (`Long` → `System.Drawing.Color`) |
| `FontName`/`FontSize`/`FontBold`… | objet **`Font`** |
| `CancelError` + `On Error` | **`ShowDialog() = DialogResult.OK`** |

Points utiles :

- le **format du filtre** est **identique** (séparateur `|`), ce qui limite les changements ;
- `OpenFileDialog.Multiselect = True` peuple **`FileNames`** (tableau) — plus net que les *flags* de VB6 ;
- `FontDialog` n'exige **pas** les *Flags* obligatoires de la `ShowFont` VB6 (qui, sans `cdlCFScreenFonts`/`cdlCFBoth`, levait une erreur).

---

## 4. Les nouveaux dialogues (sans équivalent VB6)

- **`FolderBrowserDialog`** : choisir un **dossier** — ce que VB6 imposait de bricoler (contrôle `DirListBox`, ou API `SHBrowseForFolder`). C'est le remplacement naturel des contrôles de système de fichiers disparus (**[13.5](05-correspondance-controles.md)**).
- **`PrintPreviewDialog`** et **`PageSetupDialog`** : aperçu avant impression et mise en page, **inédits** côté boîtes communes (l'impression complète est traitée en **[14.4](../14-graphismes-impression-activex/README.md)**).

---

## 5. Libération

Les composants de dialogue sont **`IDisposable`**. Deux cas :

- **créés en code** → enveloppez-les dans un **`Using`** (comme ci-dessus) pour une libération **déterministe** ;
- **déposés sur le formulaire** (bac à composants) → le **formulaire** les libère via son `Dispose` généré.

C'est exactement la logique de la **[12.3](../12-poo/03-idisposable-using.md)**.

---

## 🔁 Récapitulatif VB6 → VB.NET

| VB6 | Windows Forms |
|-----|---------------|
| `CommonDialog` (OCX, multiplexé) | composants **natifs distincts** |
| `ShowOpen` / `ShowSave` | `OpenFileDialog` / `SaveFileDialog` |
| `ShowColor` / `ShowFont` / `ShowPrinter` | `ColorDialog` / `FontDialog` / `PrintDialog` |
| `CancelError` + `On Error` | `ShowDialog() = DialogResult.OK` |
| `InitDir` / `DialogTitle` | `InitialDirectory` / `Title` |
| `Filter` (`a\|*.a\|b\|*.b`) | `Filter` (**même format**) |
| `Color` (`Long`) | `Color` (`System.Drawing.Color`) |
| *(bricolage pour un dossier)* | `FolderBrowserDialog` |

---

## ✅ Points clés

- Le contrôle **unique** `CommonDialog` (un **ActiveX**) devient **plusieurs composants natifs** : on **remplace**, on n'interopère pas.
- Changement de code central : **`CancelError`/`On Error` → `ShowDialog() = DialogResult.OK`**.
- Le **format du filtre** reste **identique** ; `FilterIndex` reste 1-based ; `InitDir`→`InitialDirectory`, `DialogTitle`→`Title`.
- Nouveaux venus utiles : **`FolderBrowserDialog`** (remplace `DirListBox`/API), `PrintPreviewDialog`, `PageSetupDialog`.
- Dialogues `IDisposable` : **`Using`** si créés en code, sinon libérés par le formulaire (12.3).

---

## 🔗 Renvois

- **[13.2](02-cycle-de-vie.md)** — `ShowDialog`/`DialogResult`. · **[12.3](../12-poo/03-idisposable-using.md)** — `Using`/`Dispose`.
- **[13.5](05-correspondance-controles.md)** — contrôles de système de fichiers disparus. · **[13.7](07-menus-toolbars.md)** — OCX remplacés par du natif.
- **[14.4](../14-graphismes-impression-activex/README.md)** — impression (`PrintDocument`, `PrintDialog`).
- **Suite → [13.9 — `MsgBox`/`InputBox`, `DoEvents` (toujours là, à manier avec précaution)](09-msgbox-doevents.md)** ⚠️

⏭️ [`MsgBox`/`InputBox`, `DoEvents` (toujours là, à manier avec précaution)](/13-formulaires-winforms/09-msgbox-doevents.md)
