🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.2 — Supprimer les dépendances à `Microsoft.VisualBasic.Compatibility` ⚠️

**Retirer l'échafaudage de compatibilité injecté par l'assistant de migration : remplacer les appels à `Microsoft.VisualBasic.Compatibility` par du code .NET idiomatique, pour ne pas figer la migration sur une couche dépréciée et sans avenir.**

> ⚠️ La « béquille » annoncée au module 4.2, retirée ici. Si vous avez utilisé l'**assistant de mise à
> niveau** (module 4.1), votre code en est probablement truffé : l'assistant produit du code qui
> *tourne*, mais au prix d'une dépendance massive à cet espace de compatibilité. Le solder, c'est faire
> passer ce code du statut « converti par un outil » à celui de « réellement migré ».

---

## Qu'est-ce que cet espace de noms ?

`Microsoft.VisualBasic.Compatibility` (assemblies `Microsoft.VisualBasic.Compatibility.dll` et
`Microsoft.VisualBasic.Compatibility.Data.dll`) est une **couche d'émulation** fournie par Microsoft
pour faciliter la conversion VB6 → VB.NET. L'**assistant de mise à niveau** l'utilise abondamment pour
**reproduire au runtime** des comportements VB6 qui ne se mappent pas proprement sur .NET :

- les **tableaux de contrôles** (*control arrays*, module 13.6), enveloppés dans des classes comme
  `…Compatibility.VB6.ButtonArray`, `LabelArray`, `TextBoxArray`… ;
- des conversions **twips ↔ pixels** (module 13.4) ;
- une émulation de l'objet **`Printer`** (module 14.4) ;
- des **helpers de tableaux** (copie, bornes — module 8.1) ;
- l'émulation de propriétés VB6 comme **`ItemData`** sur les listes/combos ;
- divers utilitaires de dessin et de données.

En somme : un **vernis** qui rejoue la sémantique VB6 pour que la sortie de l'assistant compile et  
s'exécute **sans correction manuelle** — au prix décrit ci-dessous.

---

## ⚠️ À ne **pas** confondre avec `Microsoft.VisualBasic`

C'est la confusion la plus dangereuse de cette section :

| Espace | Rôle | À faire |
|--------|------|---------|
| **`Microsoft.VisualBasic`** | La **bibliothèque cœur** de VB.NET : `MsgBox`, `Left`/`Mid`/`Right`, `CStr`, `IIf`, `GetSetting`/`SaveSetting`, l'espace `My`… | ✅ **Conserver** — supportée, légitime (et même disponible sur .NET moderne) |
| **`Microsoft.VisualBasic.Compatibility`** | La **couche d'émulation de migration** (l'objet de cette section) | ⚠️ **Supprimer** — dépréciée, sans avenir |

> ⚠️ **Ne supprimez surtout pas `Microsoft.VisualBasic`** en croyant bien faire. Seul l'espace
> **`.Compatibility`** est visé ici. (Remplacer certains appels de la bibliothèque cœur par leurs
> équivalents .NET relève d'un choix **stylistique**, traité au module 17.5 — pas d'une obligation.)

---

## Pourquoi le retirer **maintenant**

`Microsoft.VisualBasic.Compatibility` **fonctionne** sur .NET Framework 4.7.2 (il y est livré) — et
c'est précisément le **piège** : rien ne vous **force** à le retirer, alors qu'il faut le faire. Quatre  
raisons :

1. **C'est un cul-de-sac.** Cet espace est **déprécié** et **absent du .NET moderne** (.NET 5/8/10).
   Toute dépendance à `.Compatibility` est donc un **bloqueur dur** pour le **second saut** vers .NET
   moderne (module 20). Même si vous restez sur 4.7.2, vous bâtissez sur une couche en fin de vie.
2. **Ce n'est pas du .NET idiomatique.** C'est une émulation des **quirks** de VB6 — elle vous maintient
   mentalement et structurellement **arrimés au VB6**, au lieu d'adopter les mécanismes natifs.
3. **Elle masque le vrai travail de migration.** En reproduisant VB6, elle **cache** les endroits où il
   faudrait prendre une vraie décision (comment remplacer correctement un tableau de contrôles, par
   exemple). « Ça compile et ça tourne » sans que le code soit réellement migré.
4. **Maintenabilité et opacité.** Peu de développeurs connaissent cette couche ; le débogage la traverse
   difficilement, et certains comportements émulés diffèrent subtilement du natif.

---

## Les trouver

Méthode (proche de celle du module 17.1) :

- **Rechercher** dans tout le code les `Imports Microsoft.VisualBasic.Compatibility` et les références
  qualifiées `Microsoft.VisualBasic.Compatibility.VB6.…`.
- **Inspecter les références** du projet : `Microsoft.VisualBasic.Compatibility.dll` et
  `Microsoft.VisualBasic.Compatibility.Data.dll`.
- **Utiliser le compilateur comme liste de travail** : une fois les usages remplacés, **retirer la
  référence** à l'assembly transforme chaque usage restant en **erreur de compilation** → la garantie de
  n'avoir rien oublié.

---

## Les remplacer, catégorie par catégorie

### Tableaux de contrôles — le gros morceau (→ module 13.6)

C'est généralement l'usage le plus volumineux. L'assistant produit un « tableau de contrôles » émulé,  
avec un paramètre `Index` ajouté aux gestionnaires :

```vbnet
' ❌ Sortie typique de l'assistant : tableau de contrôles émulé
Friend WithEvents Command1 As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray

Private Sub Command1_Click(eventSender As Object, eventArgs As EventArgs, Index As Short) _
        Handles Command1.Click
    ' Index identifie le bouton cliqué
End Sub
```

```vbnet
' ✅ Idiomatique : de vrais Button partageant un gestionnaire (clause Handles multiple)
Private Sub Boutons_Click(sender As Object, e As EventArgs) _
        Handles btnUn.Click, btnDeux.Click, btnTrois.Click
    Dim bouton = DirectCast(sender, Button)
    ' sender (ou bouton.Name / bouton.Tag) remplace l'ancien Index
End Sub
```

Le module **13.6** détaille les stratégies complètes (gestionnaire partagé, collection de contrôles,  
recréation dynamique).

### `ItemData` des listes/combos → un vrai modèle de données

VB6 associait un `Long` à chaque élément via `ItemData` ; l'assistant l'émule par
`VB6.GetItemData`/`SetItemData`. La bonne réponse .NET : **stocker de vrais objets** comme éléments.

```vbnet
' ❌ Émulation : un entier parallèle par élément
Dim id As Integer = Microsoft.VisualBasic.Compatibility.VB6.GetItemData(ListBox1, i)
```

```vbnet
' ✅ Idiomatique : l'objet métier EST l'élément (ToString surchargé, ou liaison de données)
ListBox1.Items.Add(New Client With {.Id = 42, .Nom = "Durand"})
' ...
Dim client = DirectCast(ListBox1.SelectedItem, Client)
Dim id As Integer = client.Id
```

### Les autres catégories

- **Twips ↔ pixels** (module 13.4) : remplacer les helpers par une conversion explicite (1 pouce =
  1440 twips, via la résolution `Graphics.DpiX`) — ou, mieux, **refaire la mise en page en pixels** avec
  ancrage/docking, plutôt que de porter des coordonnées en twips.
- **Objet `Printer`** (module 14.4) : remplacer par `System.Drawing.Printing.PrintDocument`.
- **Helpers de tableaux** (`VB6.CopyArray`, bornes) : remplacer par le natif (`Array.Copy`, indices
  **0-based** du module 8.1).
- **Dessin** (module 14.1) → **GDI+** / objet `Graphics` ; **données/liaison** (module 15) → **ADO.NET**
  / `BindingSource`.

---

## Méthode et critère de sortie

La boucle, sous la protection du filet de sécurité :

1. **Inventorier** les usages (recherche + références).
2. **Remplacer** catégorie par catégorie — en commençant souvent par les **tableaux de contrôles**
   (le plus gros), et en prenant à chaque fois une **vraie décision .NET** via le module concerné (13.x,
   14.x, 15, 8.1).
3. **Valider** chaque changement contre le **golden master** (module 17.3) — l'UI surtout : la gestion
   d'événements des ex-tableaux de contrôles et la mise en page sont sensibles.
4. **Retirer la référence** à l'assembly dès qu'aucun usage ne subsiste.

> 🎯 **Critère de sortie — binaire et vérifiable** : le projet **ne référence plus**
> `Microsoft.VisualBasic.Compatibility` (ni `.Compatibility.Data`). Tant que la référence est là, la
> dette n'est pas soldée.

---

## ⚠️ Pièges à connaître

- **Confondre avec `Microsoft.VisualBasic`** : ne retirez **que** l'espace `.Compatibility` ; la
  bibliothèque cœur reste légitime. ⚠️
- **Croire que « ça compile sur 4.7.2 » suffit** : la couche y fonctionne, mais elle est **dépréciée** et
  **bloque le module 20** (.NET moderne). Le retrait est **délibéré**, sans erreur de compilation pour
  vous y forcer. ⚠️
- **Sous-estimer le changement de comportement UI** : remplacer un tableau de contrôles modifie la
  gestion d'événements et parfois la mise en page → **valider au golden master**. ⚠️
- **Remplacer mécaniquement** sans repenser le modèle (réémuler `ItemData` avec un dictionnaire parallèle
  au lieu d'utiliser de vrais objets, par exemple) : on échange une béquille contre une autre.
- **Oublier une référence transitive** (`.Compatibility.Data`) : vérifier les **deux** assemblies.

---

## En résumé

- `Microsoft.VisualBasic.Compatibility` est la **couche d'émulation** injectée par l'assistant de
  migration (tableaux de contrôles, `ItemData`, twips, `Printer`, helpers de tableaux…). Elle fait
  « tourner » la sortie de l'outil **sans la migrer vraiment**.
- **À ne pas confondre** avec `Microsoft.VisualBasic` (bibliothèque cœur, **à conserver**). Seul
  l'espace **`.Compatibility`** est visé.
- **Pourquoi maintenant** : déprécié et **absent du .NET moderne** (bloqueur pour le module 20), non
  idiomatique, et masque le vrai travail — alors que rien ne vous y force sur 4.7.2.
- **Remplacer catégorie par catégorie** (tableaux de contrôles d'abord → 13.6 ; `ItemData` → vrais
  objets ; twips → 13.4 ; `Printer` → 14.4 ; tableaux → natif/8.1), **valider** chaque lot au **golden
  master** (17.3).
- **Critère de sortie binaire** : plus **aucune** référence à `.Compatibility` (ni `.Compatibility.Data`).

> ➡️ **Le filet de sécurité invoqué à chaque étape mérite enfin son chapitre dédié : les tests de
> non-régression contre le *golden master*, pièce maîtresse de la validation.** (Module 17.3)

---

🏷️ **Indicateurs** : ⚠️ Piège (couche dépréciée, changements UI) · lié aux modules 4.1/4.2 (assistant, béquille), 13.4/13.6, 14.1/14.4, 15, 8.1, 17.3 (golden master), 20 (second saut)
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Tests de **non-régression** : comparer avec le *golden master*](/17-valider-refactoriser/03-non-regression.md)
