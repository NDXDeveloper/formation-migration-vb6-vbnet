🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 🪟 Chapitre 13 — Des formulaires VB6 à Windows Forms ⭐ ⚠️
**Migrer l'interface : du concepteur VB6 (`.frm`/`.frx`) au modèle Windows Forms (Designer + code-behind) — en préservant l'apparence *et* le comportement**

> **Partie 4 — Migrer l'interface utilisateur (les formulaires)**
> Indicateurs : ⭐ chapitre central · ⚠️ pièges fréquents

---

## 🧭 De quoi parle ce chapitre

La **Partie 3** a migré le **langage** — comment le code se *comporte*. La **Partie 4** s'attaque à un sujet de nature différente : le **modèle d'interface**. Ici, le formulaire n'est plus seulement du code, c'est un **objet visuel** dont il faut préserver à la fois l'**aspect** et les **réactions**.

C'est aussi, très concrètement, **le premier endroit où l'héritage surgit** (annoncé en **[12.5](../12-poo/05-interfaces-heritage.md)**) : tout formulaire migré **`Inherits System.Windows.Forms.Form`**.

> ⚠️ Le fil rouge reste le même — *ça compile, mais ça ne se comporte (ni ne s'affiche) plus pareil.* Pour les formulaires, le danger se déplace vers la **géométrie** (twips → pixels), l'**ordre des événements** (cycle de vie), et quelques **mécanismes disparus** (les *control arrays*).

---

## 🔀 Le grand écart : deux modèles de formulaire

VB6 mêle, dans un format propriétaire, la **description visuelle** et le **code**. Windows Forms **sépare** le code généré par le concepteur du code que vous écrivez.

| Aspect | VB6 | Windows Forms (VB.NET) |
|--------|-----|------------------------|
| Fichier(s) | `.frm` (texte : *layout* **+** code) **+** `.frx` (binaire) | `.vb` (code-behind) **+** `.Designer.vb` (généré) **+** `.resx` |
| Nature de la classe | formulaire à **instance par défaut** globale | classe qui **`Inherits …Form`** |
| Instanciation | `Form1.Show` (sans `New`) | `New Form1()` (instance par défaut via `My.Forms`, **nuancée**) |
| Unité de mesure | **twips** (1/1440 de pouce) | **pixels** |
| Design et code | **mêlés** dans le `.frm` | **séparés** (Designer ↔ code-behind) |
| Ressources | `.frx` (binaire) | `.resx` (XML) |
| Cycle de vie | `Load` / `Activate` / `Unload` | `Load` / `Shown` / `Activated` / `FormClosing` |

À retenir : chaque ligne cache soit une **conversion mécanique** (que l'outillage gère), soit un **changement de comportement** (qu'il faut traquer). Ce chapitre apprend à distinguer les deux.

---

## 🎯 Les pièges qui parcourent ce chapitre

Quatre familles de pièges reviennent d'une section à l'autre :

- **Géométrie — twips → pixels** ⚠️ : VB6 mesure en **twips**, Windows Forms en **pixels**. **Toutes** les coordonnées et tailles changent d'échelle — un décalage **silencieux** et omniprésent (**13.4**).
- **Cycle de vie** ⚠️ : `Load`/`Unload`/`Activate` ne correspondent pas un pour un à `Load`/`FormClosing`/`Shown`, et `Show` **modal** de VB6 se scinde en `Show` (non modal) **et** `ShowDialog` (modal) (**13.2**).
- **Instance par défaut** : `Form1.Show` sans instanciation existait en VB6 (instance globale) ; VB.NET en propose un équivalent via `My.Forms`, mais **par thread** et au comportement **subtilement différent** (**13.1**/**13.2**).
- **Mécanismes disparus** ⭐ ⚠️ : les **tableaux de contrôles** (*control arrays*) **n'existent plus** — un point structurant qui demande une **solution de remplacement** (**13.6**). Et `DoEvents`, **toujours là**, reste **dangereux** (réentrance) (**13.9**).

---

## 🛠️ Outillage et part manuelle

Bonne nouvelle : la conversion des formulaires est **fortement outillée**. L'assistant de mise à niveau et les outils commerciaux (**VBUC**, **VB Migration Partner** — voir **[chapitre 4](../04-outils-migration/README.md)**) savent **analyser le `.frm`**, **générer le code Designer** et **mapper les contrôles** intrinsèques.

> Repère honnête (rappel du cadrage) : l'outil fait le gros œuvre, mais **le jugement reste nécessaire** pour les *control arrays*, l'**ajustement de la géométrie** (twips/pixels, ancrage), le **dessin personnalisé** et les **contrôles tiers (OCX)** — ces deux derniers relevant du **[chapitre 14](../14-graphismes-impression-activex/README.md)**.

---

## ⚖️ Ce que ce chapitre couvre — et ce qu'il ne couvre pas

- ✅ **Couvre** : le **modèle de formulaire** (Designer/code-behind), le **cycle de vie**, les **MDI**, la **géométrie**, les **contrôles intrinsèques**, les **control arrays**, les **menus et barres d'outils**, les **boîtes de dialogue communes**, et `MsgBox`/`InputBox`/`DoEvents`.
- ❌ **Ne couvre pas** : le **dessin** (GDI+), l'**impression** et les **ActiveX/OCX** → **[chapitre 14](../14-graphismes-impression-activex/README.md)** ; la **liaison de données** → **[chapitre 15](../15-acces-donnees/README.md)**.

---

## 🗺️ Plan du chapitre

Le chapitre suit le parcours « **le formulaire d'abord, puis ses contrôles, puis ses accessoires** » :

1. **13.1 — [Le modèle de formulaire : `.frm`/`.frx` → `.vb` + Designer](01-modele-formulaire.md)** : le conteneur, la classe partielle générée, les ressources.
2. **13.2 — [Cycle de vie : `Load`/`Unload`/`Activate` → `Load`/`FormClosing`/`Shown` ; `Show` vs `ShowDialog`](02-cycle-de-vie.md)** ⚠️ : l'ordre et la nature des événements.
3. **13.3 — [Formulaires MDI](03-formulaires-mdi.md)** ⚠️ : `MDIForm`/`MDIChild` → `IsMdiContainer`/`MdiParent`, `LayoutMdi`, fusion de menus.
4. **13.4 — [Propriétés, ancrage et positionnement : Twips → pixels](04-twips-pixels-ancrage.md)** ⚠️ : la conversion d'unités, partout.
5. **13.5 — [Correspondance des contrôles intrinsèques](05-correspondance-controles.md)** : `CommandButton`→`Button`, `TextBox`, `Label`, `Frame`, `OptionButton`…
6. **13.6 — [Les tableaux de contrôles (*control arrays*) : disparus → remplacements](06-control-arrays.md)** ⭐ ⚠️ : tableau manuel, handler partagé, ou bibliothèque dédiée.
7. **13.7 — [Menus et barres d'outils](07-menus-toolbars.md)** : Menu Editor → `MenuStrip`, `ToolStrip`.
8. **13.8 — [Boîtes de dialogue communes](08-dialogues-communs.md)** : `CommonDialog` → `OpenFileDialog`, `ColorDialog`…
9. **13.9 — [`MsgBox`/`InputBox`, `DoEvents`](09-msgbox-doevents.md)** ⚠️ : toujours présents, à manier avec précaution.

---

## ✅ Messages clés à emporter

- La migration d'interface est **de nature différente** de celle du langage : on préserve un **rendu** et des **réactions**, pas seulement une sémantique.
- Un formulaire devient une **classe qui `Inherits …Form`**, avec **séparation** Designer / code-behind — c'est ici que l'**héritage** apparaît pour de bon.
- Le piège **le plus diffus** est la **géométrie** : **twips → pixels** touche toutes les tailles et positions.
- Le piège **le plus structurant** est la **disparition des *control arrays*** : prévoir une **stratégie de remplacement**.
- L'**outillage** convertit beaucoup, mais **géométrie, cycle de vie et control arrays** réclament une **revue manuelle**.

---

## 🔗 À relire en parallèle

- **[12.5](../12-poo/05-interfaces-heritage.md)** — l'héritage (d'où vient `Inherits …Form`). · **[12.8](../12-poo/08-evenements.md)** — événements (handlers partagés, *control arrays*).
- **[Chapitre 4](../04-outils-migration/README.md)** — outils de migration (assistant, VBUC, VB Migration Partner).
- **[Chapitre 14](../14-graphismes-impression-activex/README.md)** — dessin, impression, ActiveX. · **[Chapitre 15](../15-acces-donnees/README.md)** — liaison de données.
- **Annexe C** — *Correspondance des contrôles VB6 → Windows Forms* (l'aide-mémoire à garder ouvert).

---

> Une fois les sections 13.1 à 13.9 parcourues : **[14. Graphismes, impression et contrôles ActiveX](../14-graphismes-impression-activex/README.md)** ⚠️ — la suite de la migration de l'interface.

⏭️ [Le modèle de formulaire VB6 (`.frm`/`.frx`) → Windows Forms (`.vb` + Designer)](/13-formulaires-winforms/01-modele-formulaire.md)
