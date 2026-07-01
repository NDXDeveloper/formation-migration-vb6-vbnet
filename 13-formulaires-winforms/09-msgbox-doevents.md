🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.9 — `MsgBox`/`InputBox`, `DoEvents` (toujours là, à manier avec précaution) ⚠️

> **Chapitre 13 — Des formulaires VB6 à Windows Forms** · Section 13.9 *(dernière section du chapitre)*
> Trois fonctions VB6 **survivent** en VB.NET (espace `Microsoft.VisualBasic`). Pratique pour migrer… mais deux d'entre elles demandent de la **prudence**, et `DoEvents` est un **signal d'alerte**.

---

## 🧭 Périmètre

`MsgBox`, `InputBox` et `DoEvents` **existent toujours** en VB.NET — elles « fonctionnent telles quelles », ce qui facilite une migration **fidèle**. Mais : `MsgBox` a une **contrepartie idiomatique** au piège d'**ordre des arguments**, `InputBox` conserve une **ambiguïté** ancienne, et `DoEvents` reste **dangereux**.

---

## 1. `MsgBox` : survit, mais `MessageBox.Show` est idiomatique

`MsgBox` fonctionne en VB.NET (il renvoie un `MsgBoxResult`) : le code VB6 passe **presque tel quel**. La forme **idiomatique** .NET est cependant **`MessageBox.Show`** (qui renvoie un `DialogResult` et utilise des **énumérations**).

```vb
' VB6 (et VB.NET : MsgBox survit) — 2e arg = BOUTONS, 3e = TITRE
Dim r = MsgBox("Enregistrer les modifications ?", vbYesNo + vbQuestion, "Confirmation")
If r = vbYes Then Enregistrer()
```

```vb
' VB.NET idiomatique — MessageBox.Show : 2e arg = TITRE, 3e = BOUTONS
Dim r = MessageBox.Show("Enregistrer les modifications ?", "Confirmation",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question)
If r = DialogResult.Yes Then Enregistrer()
```

> ⚠️ **Piège de l'ordre des arguments.** Dans `MsgBox`, le **2ᵉ** argument est les **boutons** et le **3ᵉ** le **titre**. Dans `MessageBox.Show`, c'est **l'inverse** : le **2ᵉ** est le **titre**, les **boutons** viennent en **3ᵉ**. Convertir naïvement `MsgBox` en `MessageBox.Show` **sans réordonner** met le titre à la place des boutons — un **bug silencieux**. Les valeurs de retour diffèrent aussi (`MsgBoxResult` ↔ `DialogResult`), mais leurs membres se correspondent (`vbYes` ↔ `DialogResult.Yes`).

---

## 2. `InputBox` : survit (propre à VB), mais limité

`InputBox` existe toujours (c'est une fonction **propre à VB**, sans équivalent natif Windows Forms) et renvoie la **chaîne saisie**.

```vb
' VB6 (et VB.NET : InputBox survit)
Dim nom = InputBox("Votre nom ?", "Saisie", "Anonyme")
```

> ⚠️ **Ambiguïté Annuler / vide.** `InputBox` renvoie **`""`** si l'utilisateur **annule**… **mais aussi** s'il valide un **champ vide**. Impossible de **distinguer** les deux. Cette limite VB6 **persiste**.

Pour une détection **fiable** de l'annulation (ou une saisie plus riche), construisez un **petit formulaire de dialogue** dédié : un `TextBox` + boutons OK/Annuler, le bouton OK fixant `DialogResult.OK`. L'appelant teste alors `ShowDialog()` **et** lit la propriété de saisie — une approche **propre** (et c'est le modèle `ShowDialog`/`DialogResult` de **[13.2](02-cycle-de-vie.md)**).

---

## 3. `DoEvents` : survit, mais c'est un **signal d'alerte** ⚠️

`Application.DoEvents()` existe toujours (et `DoEvents` de `Microsoft.VisualBasic` aussi). En VB6 **monothread**, c'était **le** moyen de garder l'interface réactive pendant une boucle longue. Mais il **conserve ses dangers**.

```vb
' VB6 (et VB.NET) — garder l'UI réactive avec DoEvents
For i = 1 To 1000000
    Traiter(i)
    DoEvents          ' ⚠️ réentrance possible
Next
```

> ⚠️ **Réentrance.** Pendant que `DoEvents` traite les messages en attente, l'utilisateur peut **recliquer** (déclenchant **à nouveau** le même gestionnaire, en réentrance), **fermer** le formulaire, ou modifier l'état en cours. D'où des bugs subtils, déjà présents en VB6 — et toujours là en .NET.

### Les alternatives modernes (inexistantes en VB6)

.NET offre de **meilleures** réponses pour garder l'UI réactive **sans** `DoEvents` :

- **`Async`/`Await`** (sur `Task`) — la voie moderne, disponible sur **.NET Framework 4.5+** (donc **4.7.2**) ;
- **`BackgroundWorker`** — composant Windows Forms : travail en arrière-plan, rapport de progression, mises à jour d'UI encadrées ;
- **`Task.Run`** / threads — déporter le calcul.

```vb
' VB.NET — alternative moderne : async/await (UI réactive, sans DoEvents)
Private Async Sub btnTraiter_Click(sender As Object, e As EventArgs) Handles btnTraiter.Click
    btnTraiter.Enabled = False
    Await Task.Run(Sub() TraiterTout())   ' calcul HORS du thread UI
    btnTraiter.Enabled = True             ' de retour sur le thread UI (sûr)
End Sub
```

> ⚠️ **Accès inter-threads.** Si vous déportez du travail sur un autre thread, vous **ne devez pas** toucher les contrôles depuis ce thread (exception inter-thread) : utilisez **`Invoke`/`BeginInvoke`**, ou laissez `Await` vous **ramener** sur le thread UI (comme ci-dessus). C'est une contrainte **absente** du monde VB6 monothread.

**Posture de migration** : **conservez** `DoEvents` pour rester **fidèle** au comportement initial, mais **signalez-le** comme cible de **refactoring** (**[17.5](../17-valider-refactoriser/README.md)**). Et n'**ajoutez pas** de nouveaux `DoEvents`.

---

## 4. Le point commun : `Microsoft.VisualBasic`

Ces trois fonctions proviennent de l'espace **`Microsoft.VisualBasic`**, **importé par défaut** : c'est ce qui les rend disponibles **sans effort**. Avantage pour la migration ; en contrepartie, on **garde une dépendance** au runtime VB et des **patterns non idiomatiques** — à arbitrer au refactoring (rappel de la logique vue tout au long du chapitre).

---

## 🔁 Récapitulatif VB6 → VB.NET

| VB6 | VB.NET — fidèle | VB.NET — idiomatique |
|-----|-----------------|----------------------|
| `MsgBox(p, boutons, titre)` | `MsgBox(...)` (survit) | `MessageBox.Show(p, titre, boutons, …)` ⚠️ **ordre différent** |
| `InputBox(...)` | `InputBox(...)` (survit) | **formulaire dédié** (`DialogResult` + saisie) |
| `DoEvents` | `DoEvents` / `Application.DoEvents()` (survit) | **`Async`/`Await`**, `BackgroundWorker`, `Task.Run` |

---

## ✅ Points clés

- `MsgBox`, `InputBox` et `DoEvents` **survivent** (espace `Microsoft.VisualBasic`) : utile pour une migration **fidèle**.
- ⚠️ `MessageBox.Show` est idiomatique mais **inverse l'ordre** boutons/titre par rapport à `MsgBox` — réordonner en convertissant.
- ⚠️ `InputBox` ne distingue pas **Annuler** d'une **saisie vide** : pour une détection fiable, un **formulaire dédié**.
- ⚠️ `DoEvents` reste **réentrant** et **dangereux** : le **conserver** pour la fidélité, mais le **refactorer** vers **`Async`/`Await`** ou `BackgroundWorker` — sans en ajouter de nouveaux.
- Déporter du travail impose de gérer l'**accès inter-threads** (`Invoke`/`BeginInvoke` ou retour via `Await`).

---

## 🔗 Renvois

- **[13.2](02-cycle-de-vie.md)** — `ShowDialog`/`DialogResult` (formulaire de saisie dédié). · **[12.4](../12-poo/04-modules-globales.md)** — fonctions « globales » de `Microsoft.VisualBasic`.
- **[17.5](../17-valider-refactoriser/README.md)** — refactoring (retrait de `DoEvents`, passage à l'asynchrone).
- **Fin du chapitre 13.** Suite → **[14. Graphismes, impression et contrôles ActiveX](../14-graphismes-impression-activex/README.md)** ⚠️

⏭️ [Graphismes, impression et contrôles ActiveX](/14-graphismes-impression-activex/README.md)
