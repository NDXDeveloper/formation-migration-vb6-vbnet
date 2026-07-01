🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.6 — Les tableaux de contrôles (*control arrays*) : disparus → solutions de remplacement ⭐ ⚠️

> **Chapitre 13 — Des formulaires VB6 à Windows Forms** · Section 13.6
> Les *control arrays* **n'existent pas** en Windows Forms. C'est l'un des changements les plus **structurants** du chapitre — et un **point de retouche manuelle** fréquent.

---

## 🧭 Pourquoi cette section est importante

Le tableau de contrôles est un mécanisme **omniprésent** dans le code VB6, et il **disparaît** sans remplacement automatique satisfaisant. Comprendre **ce qu'il faisait** et **comment le remplacer** est indispensable — d'autant que la traduction naïve s'appuie souvent sur une **béquille** à retirer (voir §2).

---

## 1. Rappel : qu'est-ce qu'un *control array* VB6 ?

Un tableau de contrôles regroupe plusieurs contrôles **de même type** qui partagent :

- le **même nom** (ex. plusieurs `txtChamp`) ;
- une propriété **`Index`** qui les distingue (`txtChamp(0)`, `txtChamp(1)`…) ;
- **un seul jeu** de gestionnaires d'événements — la procédure reçoit l'**`Index`** en paramètre.

Il servait **trois** usages :

1. **Partage d'événements** : un handler pour plusieurs contrôles, l'`Index` indiquant lequel a réagi (ex. un pavé numérique) ;
2. **Création dynamique à l'exécution** : `Load txtChamp(2)` / `Unload txtChamp(2)` — c'était **le** moyen VB6 de créer des contrôles au runtime ;
3. **Itération** sur un groupe : `For i = 0 To txtChamp.Count - 1 … Next`.

---

## 2. Pourquoi ils disparaissent en Windows Forms

Windows Forms **ne connaît pas** les tableaux de contrôles : pas d'`Index`, pas de `Load`/`Unload`. **Chaque contrôle est un objet indépendant**, et la création dynamique se fait simplement avec **`New`**. La béquille n'est donc plus nécessaire — mais les **trois usages** ci-dessus doivent être **réimplémentés**.

> ⚠️ **La traduction naïve s'appuie sur une béquille.** L'assistant de mise à niveau émule souvent les *control arrays* via l'espace **`Microsoft.VisualBasic.Compatibility`** — précisément la **béquille à éviter** (voir **[4.2](../04-outils-migration/README.md)**). Mieux vaut adopter une **vraie** solution de remplacement (ci-dessous) et planifier le retrait de cette dépendance.

Trois stratégies, selon l'usage d'origine.

---

## 3. Stratégie 1 — Le **handler partagé** (pour le partage d'événements)

Pour le cas « un handler pour plusieurs contrôles », Windows Forms permet à **plusieurs** contrôles de partager **un** gestionnaire — via une clause **`Handles`** multiple (ou `AddHandler`). On identifie le contrôle qui a réagi grâce au paramètre **`sender`**, et l'on porte la donnée d'identité dans **`Tag`** — qui **remplace l'`Index`**.

```vb
' VB6 — un handler pour 10 boutons (Index)
Private Sub cmdChiffre_Click(Index As Integer)
    txtAffichage.Text = txtAffichage.Text & CStr(Index)
End Sub
```

```vb
' VB.NET — handler partagé ; on distingue via sender + Tag
' (chaque bouton porte son chiffre dans Tag, fixé dans le concepteur)
Private Sub cmdChiffre_Click(sender As Object, e As EventArgs) _
        Handles cmd0.Click, cmd1.Click, cmd2.Click, cmd3.Click   ' … etc.
    Dim b = DirectCast(sender, Button)
    txtAffichage.Text &= b.Tag.ToString()      ' Tag remplace Index
End Sub
```

> 💡 C'est la solution **la plus simple et idiomatique** pour un **groupe fixe** de contrôles. Elle s'appuie sur le modèle d'événements vu en **[12.8](../12-poo/08-evenements.md)**.

---

## 4. Stratégie 2 — Recréer un **tableau à la main** (pour la création dynamique et l'itération)

Pour les cas « créer dynamiquement » et « itérer », on maintient **sa propre** collection de contrôles (`List(Of T)` ou tableau), et l'on **crée/supprime** les contrôles soi-même — ce qui **remplace `Load`/`Unload`**.

```vb
' VB6 — créer dynamiquement via Load
Load txtChamp(i)
txtChamp(i).Top = txtChamp(i - 1).Top + 400   ' twips
txtChamp(i).Visible = True
```

```vb
' VB.NET — créer dynamiquement via New + Controls.Add (remplace Load)
Private _champs As New List(Of TextBox)()

Private Sub AjouterChamp()
    Dim tb As New TextBox()
    tb.Location = New Point(10, 10 + _champs.Count * 28)   ' pixels (voir 13.4)
    tb.Width = 200
    AddHandler tb.TextChanged, AddressOf Champ_Change       ' câblage dynamique
    Me.Controls.Add(tb)
    _champs.Add(tb)
End Sub

Private Sub Champ_Change(sender As Object, e As EventArgs)
    ' … traiter le champ qui a changé, identifié via sender …
End Sub
```

Et la **suppression** — avec **désabonnement** et **libération** (rappel **[12.3](../12-poo/03-idisposable-using.md)**) :

```vb
' VB.NET — supprimer (remplace Unload)
Private Sub SupprimerDernierChamp()
    If _champs.Count = 0 Then Return
    Dim tb = _champs(_champs.Count - 1)
    RemoveHandler tb.TextChanged, AddressOf Champ_Change     ' éviter la fuite (12.8)
    Me.Controls.Remove(tb)
    tb.Dispose()                                             ' libérer (12.3)
    _champs.RemoveAt(_champs.Count - 1)
End Sub
```

L'**itération** se fait sur votre liste (ou en filtrant les `Controls`) :

```vb
' VB.NET — itérer
For Each tb In _champs
    tb.Clear()
Next
' (ou, sans liste dédiée : For Each tb In Me.Controls.OfType(Of TextBox)() … )
```

> 💡 Pour éviter le calcul manuel des positions, déposez les contrôles dynamiques dans un **`FlowLayoutPanel`** ou un **`TableLayoutPanel`** : ils s'**arrangent automatiquement**. C'est souvent la version la plus propre pour des **lignes dynamiques**.

---

## 5. Stratégie 3 — Une **bibliothèque d'émulation** (outils tiers)

Les outils commerciaux **VB Migration Partner** (Code Architects) et **VBUC** (Mobilize.Net) fournissent une **bibliothèque d'émulation** de *control arrays* : elle recrée la sémantique VB6 (`txtChamp(i)`, `Index`, `Load`/`Unload`) en .NET, ce qui permet de garder le code **quasi inchangé**.

- ✅ **Avantages** : **changement minimal**, fidélité à l'idiome VB6, migration **rapide** — utile pour un **gros volume** de code.
- ⚠️ **Inconvénients** : une **dépendance** au runtime de l'outil (comme l'espace `Compatibility`), un code **non idiomatique**, et une **dette technique** à retirer ensuite (refactoring vers les stratégies 1/2).

> 💡 Position raisonnable : utiliser l'émulation pour **passer le cap** rapidement, puis **migrer progressivement** vers les patterns natifs (**[17.5](../17-valider-refactoriser/README.md)**) — exactement la logique qu'on applique à l'espace `Compatibility`.

---

## 6. Quelle stratégie choisir ?

| Usage VB6 du *control array* | Stratégie .NET recommandée |
|------------------------------|----------------------------|
| Partage d'événements, groupe **fixe** | **Handler partagé** (`Handles …` + `Tag`) — §3 |
| Création **dynamique** (`Load`/`Unload`) | **Tableau manuel** (`List` + `New` + `Controls.Add` + `AddHandler`) — §4 |
| **Itération** sur un groupe | tableau manuel, ou `Controls.OfType(Of …)()` — §4 |
| Gros code, **changement minimal** voulu | **bibliothèque d'émulation** (outil tiers) — §5 *(dette à retirer)* |

---

## 7. Cas connexe : les **menus** en tableau

VB6 connaissait aussi des **tableaux d'éléments de menu** (menus avec `Index`). Ils se remplacent de la même façon : créer des **`ToolStripMenuItem`** dynamiquement et les relier à un **handler partagé** (via `AddHandler`). Le détail des menus est en **[13.7](07-menus-toolbars.md)**.

---

## 🔁 Récapitulatif VB6 → VB.NET

| VB6 | Windows Forms |
|-----|---------------|
| `ctrl(Index)` (même nom + `Index`) | contrôles **indépendants** (pas d'`Index`) |
| handler avec paramètre `Index` | **handler partagé** (`Handles …`) + **`sender`/`Tag`** |
| `Load ctrl(i)` / `Unload ctrl(i)` | **`New`** + `Controls.Add` / `Controls.Remove` + `Dispose` |
| `ctrl.Count`, boucle sur `ctrl(i)` | **`List(Of T)`** dédiée, ou `Controls.OfType(Of T)()` |
| émulation via `Compatibility` (assistant) | ⚠️ **béquille** → stratégie native, ou **outil tiers** assumé comme dette |

---

## ✅ Points clés

- Les *control arrays* **n'existent pas** en Windows Forms : **chaque contrôle est indépendant**, et `New` remplace `Load`.
- ⚠️ La traduction naïve s'appuie sur la **béquille `Compatibility`** ([4.2](../04-outils-migration/README.md)) : préférez une **vraie** solution de remplacement.
- **Trois stratégies** : **handler partagé** (`Handles` + `Tag`) pour un groupe fixe ; **tableau manuel** (`List` + `New` + `Controls.Add` + `AddHandler`) pour le dynamique ; **bibliothèque d'émulation** (outil tiers) pour un changement minimal, à **retirer** ensuite.
- À la création dynamique, pensez **désabonnement** (`RemoveHandler`) et **libération** (`Dispose`) à la suppression ; `FlowLayoutPanel`/`TableLayoutPanel` simplifient la disposition.
- **`Tag`** (et `sender`) **remplacent l'`Index`** pour identifier le contrôle.

---

## 🔗 Renvois

- **[12.8](../12-poo/08-evenements.md)** — événements (`Handles`, `AddHandler`/`RemoveHandler`). · **[12.3](../12-poo/03-idisposable-using.md)** — `Dispose` (suppression de contrôles).
- **[13.5](05-correspondance-controles.md)** — propriété `Index` / contrôles. · **[13.7](07-menus-toolbars.md)** — menus (tableaux de menus).
- **[4.2](../04-outils-migration/README.md)** — `Microsoft.VisualBasic.Compatibility` (la béquille). · **[17.5](../17-valider-refactoriser/README.md)** — refactoring (retrait de l'émulation).
- **Suite → [13.7 — Menus (Menu Editor → `MenuStrip`) et barres d'outils (`ToolStrip`)](07-menus-toolbars.md)**

⏭️ [Menus (Menu Editor → `MenuStrip`) et barres d'outils (`ToolStrip`)](/13-formulaires-winforms/07-menus-toolbars.md)
