🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.6 Contrôles tiers courants (grilles, calendriers…) : stratégies de remplacement

> Les contrôles **Microsoft** ont presque tous un équivalent natif (section 14.5). Les contrôles
> **tiers commerciaux** — grilles puissantes, plannings, graphiques avancés, éditeurs riches — sont
> une autre affaire : souvent **sans équivalent direct**, avec une **API propriétaire** et **beaucoup
> de code** qui en dépend. Les remplacer, c'est **réimplémenter un comportement**, pas échanger un
> contrôle. La clé pour rendre l'exercice tenable : **migrer l'usage *réel*, pas la fiche technique.**

**Module 14 — Graphismes, impression et contrôles ActiveX** · Cible : .NET Framework 4.7.2 ·
Windows Forms

---

## 1. Pourquoi les contrôles tiers sont une catégorie à part

C'étaient les **outils de puissance** des applications VB6 : grilles type tableur (FarPoint Spread,  
VSFlexGrid, True DBGrid, Sheridan Data Widgets…), plannings de rendez-vous, graphiques (ChartFX,  
Olectra…), éditeurs masqués, arbres enrichis. Par rapport aux common controls :

- pas de **correspondance 1:1** native → le remplacement est une **réécriture de comportement** ;
- une **API propriétaire** très étendue, souvent **profondément imbriquée** dans la logique métier ;
- les mêmes risques qu'au §14.5 (licence, éditeur, 32 bits), **aggravés** par les écarts de
  fonctionnalités et le volume de code dépendant.

---

## 2. L'éventail des stratégies

Pour **chaque** contrôle tiers, on choisit dans un spectre — du moins coûteux au plus coûteux :

| Stratégie | Quand | Compromis |
|-----------|-------|-----------|
| **1. Édition .NET de l'éditeur d'origine** | L'éditeur existe encore et fournit une suite WinForms | API **différente** (pas un *drop-in*), coût de licence ; voie souvent la plus rapide |
| **2. Autre contrôle commercial .NET** | Éditeur disparu ou trop coûteux | Réapprentissage d'API, mais **support** et **pérennité** |
| **3. Contrôle natif (.NET), éventuellement étendu** | Le natif couvre l'essentiel de l'usage réel | **Gratuit**, peu de dépendances, idéal pour le saut .NET moderne ; travail custom pour les manques |
| **4. Interop COM temporaire** (voie A du §14.5) | Dé-risquer la bascule | Pont COM à résorber ensuite |
| **5. Contrôle personnalisé** | Besoin de niche sans bonne option | **Coût élevé** : dernier recours |

> ℹ️ Beaucoup d'éditeurs VB6 proposent une **édition WinForms** (Infragistics — héritier de Sheridan
> —, ComponentOne, FarPoint Spread, DevExpress, Telerik, Syncfusion…). Les **rachats** ont été
> nombreux au fil des ans : **vérifiez l'existence actuelle de l'éditeur, la licence et la voie de
> migration proposée** au moment de l'inventaire. Ne présumez pas qu'une édition .NET est compatible
> ligne à ligne.

---

## 3. La méthode : migrer l'**usage réel**, pas la fiche technique

C'est l'idée qui rend le remplacement **tractable**. Une grille tierce expose peut-être 300 propriétés
— mais votre application n'en utilise souvent que **vingt**.

1. **Inventorier l'usage effectif** : où le contrôle est employé, comment il est **configuré**
   (concepteur vs code), **lié ou non lié** aux données, quels **propriétés/méthodes/événements** sont
   réellement appelés.
2. **Cartographier** chaque fonctionnalité **utilisée** vers l'équivalent du contrôle cible.
3. **Repérer les écarts** : ce que le natif (ou le nouveau contrôle) ne fait pas comme avant.
4. **Traiter les écarts** : étendre (dessin propriétaire, colonnes/cellules personnalisées), accepter
   un **léger changement d'UX**, ou monter en gamme (contrôle commercial plus riche).

> ⚠️ **Piège classique** : vouloir reproduire **100 %** des capacités du contrôle d'origine au lieu
> de ce qui est **vraiment utilisé**. C'est la première cause de dérapage de planning sur cette
> partie.

---

## 4. Encapsuler derrière une **abstraction** (le pattern clé)

Pour que le reste de l'application **ne dépende pas** du contrôle choisi, placez-le derrière une petite
**abstraction** (interface + `UserControl`/adaptateur) exposant **les opérations dont vous avez
besoin** :

```vb
' Le code métier parle à CETTE abstraction, pas à DataGridView/DevExpress/etc.
Public Interface IGrilleArticles
    Sub Charger(articles As IEnumerable(Of Article))
    Function ArticleSelectionne() As Article
    Event SelectionChangee(article As Article)
End Interface

' Une implémentation possible, ici sur DataGridView (interchangeable plus tard)
Public Class GrilleArticlesDgv
    Implements IGrilleArticles
    ' encapsule le DataGridView et traduit ses événements vers SelectionChangee
End Class
```

> 🔗 Bénéfices : on peut **changer d'implémentation** (interop → natif → commercial) en touchant **un
> seul endroit**, ce qui colle à la **migration incrémentale** (module 3) et au **découplage UI /
> logique** (module 5.3). C'est aussi une application directe des **interfaces** (module 12.5), et
> cela **prépare** le saut vers .NET moderne (module 20).

---

## 5. Les **grilles** : le cas le plus fréquent (et le plus douloureux)

### Le paysage VB6 et la cible native

Grilles d'**affichage** (VSFlexGrid, MSFlexGrid) et grilles **liées aux données** (True DBGrid,  
DataGrid, FarPoint Spread, Sheridan). La cible native est presque toujours **`DataGridView`**, qui  
couvre l'essentiel : colonnes typées (texte, liste, case à cocher, bouton, image, lien), **liaison de  
données** via `BindingSource`, tri, mise en forme, colonnes/cellules **personnalisées** (dessin  
propriétaire, sous-classes de `DataGridViewColumn`), colonnes/lignes **figées**, lecture seule, et
**mode virtuel** pour les très gros volumes.

### Les écarts à anticiper

`DataGridView` **ne fait pas** nativement : **cellules fusionnées**, grilles **hiérarchiques/à
bandes**, **éditeurs en cellule** complexes, en-têtes **multi-niveaux**, ou **formules** de type  
tableur (Spread). Pour ces besoins : **étendre**, accepter un changement, ou choisir une **grille  
commerciale .NET**.

### Le changement de paradigme d'accès aux cellules

VB6 poussait souvent des valeurs **cellule par cellule** (grille **non liée**). L'équivalent existe,  
mais la liaison de données est plus idiomatique :

```vb
' VB6 (non liée) :        grille.TextMatrix(ligne, colonne) = "Valeur"
' DataGridView (non liée) :  dgv.Rows(ligne).Cells(colonne).Value = "Valeur"

' Lecture : la valeur est un Object -> transtypage (Option Strict)
Dim v As String = CStr(dgv.Rows(ligne).Cells(colonne).Value)
```

Et les **événements** changent de forme :

```vb
' grille.Click            -> dgv.CellClick / dgv.CellContentClick
' e.RowIndex / e.ColumnIndex donnent la cellule concernée
```

> ℹ️ Pour les **gros jeux de données**, préférez la **liaison** (`BindingSource`, cf. **15.4**) ou le
> **mode virtuel** plutôt que de remplir les cellules une à une — meilleures performances et code plus
> simple.

---

## 6. Calendriers, saisie et plannings

- **Dates** : `DTPicker`/`MonthView` ont des équivalents natifs (`DateTimePicker`/`MonthCalendar`,
  déjà vus au §14.5). En revanche, un **planning** de type agenda (vue Outlook) **n'a pas**
  d'équivalent natif → suite **commerciale** ou contrôle personnalisé.
- **Saisie masquée** : `MaskEdBox` (MSMASK32) → **`MaskedTextBox`** natif. ⚠️ La **syntaxe des
  masques diffère** : retravailler les masques de saisie.
- **Compteurs numériques / monétaires** : → `NumericUpDown` (et liaison/format culture, cf. Annexe
  B.10).

---

## 7. Graphiques

Les contrôles graphiques tiers (`MSChart`, ChartFX, Olectra…) se remplacent par le contrôle natif
**`Chart`** (`System.Windows.Forms.DataVisualization.Charting`, fourni avec .NET Framework et
disponible en 4.7.2) — nombreux types de graphiques et **liaison de données** — ou par une
**bibliothèque de graphiques commerciale**.

> 🔗 Le contrôle `Chart` a **évolué** dans .NET moderne : si une migration ultérieure est envisagée,
> tenez-en compte (**module 20**).

---

## 8. Coût, risque et pérennité

- **Suites commerciales .NET** : coût de **licence** et courbe d'apprentissage, mais **support**,
  **longévité**, **64 bits** et compatibilité **DPI** — et un terrain plus sûr pour la suite.
- **Contrôles natifs (étendus)** : gratuits, peu de dépendances, **meilleure trajectoire** vers .NET
  moderne ; coût en **travail custom** pour les fonctionnalités avancées.
- **L'effort augmente** avec l'ampleur des fonctionnalités **réellement utilisées** et le volume de
  code dépendant.
- **Dé-risquer** : garder en **interop** (§14.5) les contrôles les plus complexes pour la première
  bascule, remplacer **ensuite**.
- **Déploiement** : les contrôles commerciaux ont leurs **redistribuables et licences** à gérer
  (module 18).

---

## 9. Ce que font les outils (renvoi module 4)

Les outils commerciaux (**VBUC**, **VB Migration Partner**) **mappent** une partie des contrôles  
courants (y compris certains tiers répandus) vers des équivalents .NET et gèrent les **tableaux de  
contrôles** (cf. 13.6). Certains **éditeurs** fournissent aussi leurs propres **assistants de  
migration** de leur version COM vers leur version .NET — à vérifier au cas par cas.

---

## 10. 🧭 Approche recommandée et ⚠️ pièges

**Recommandé**

1. **Inventorier** chaque contrôle tiers et son **usage effectif** (modules 3.1/3.2).
2. Choisir **par contrôle** dans l'éventail du §2 ; privilégier le **natif étendu** quand il couvre
   l'usage réel, la **suite commerciale** sinon.
3. **Encapsuler** derrière une abstraction (§4) pour pouvoir **réviser le choix** plus tard.
4. **Lier les données** plutôt que remplir cellule par cellule ; traiter la **culture** (Annexe B.10).
5. **Dé-risquer** via interop temporaire, puis **résorber** la dette.

**Pièges silencieux**

- **Viser 100 % des fonctionnalités** au lieu de l'usage réel → dérapage.
- Croire qu'une **édition .NET de l'éditeur** est un *drop-in* → l'API diffère.
- **Disséminer** l'API du nouveau contrôle partout (pas d'abstraction) → coûteux si l'on change encore.
- **Accès cellules/événements** : `TextMatrix` → `Cells.Value` (Object à transtyper), `Click` →
  `CellClick`.
- **Non lié vs lié** et **performances** : préférer la liaison / le mode virtuel pour les gros volumes.
- **Masques** : syntaxe différente (`MaskedTextBox`).
- **Licences/redistribuables** commerciaux oubliés au déploiement (module 18).
- **Impact sur le saut .NET moderne** : certains contrôles (dont `Chart`) ont changé (module 20).

> 🔗 Voir la **section 14.5** (décision interop/remplacement, licences, bitness), l'**Annexe C**
> (correspondance des contrôles), la **section 15.4** (liaison de données), le **module 5.3**
> (découplage), le **module 3** (incrémental) et le **module 20** (.NET moderne).

---

## ✅ Fin du module 14

Ce chapitre n'a pas été une affaire de mots-clés mais de **modèles** :

- le **dessin** passe d'un rendu *persistant* à un rendu *redessiné* dans `Paint` (14.1), avec un
  **système de coordonnées** à repenser (twips → pixels, 14.2) ;
- les **ressources** changent de format et de mode d'accès (`.frx`/`.res` → `.resx`/`My.Resources`,
  14.3) ;
- l'**impression** devient *événementielle* (`Printer` → `PrintDocument`/`PrintPage`, 14.4) ;
- et la dépendance aux **ActiveX/OCX** se transforme en **décision d'architecture** — réutiliser via
  interop ou remplacer (14.5), les contrôles **tiers** appelant des **stratégies de remplacement**
  dédiées (14.6).

Gardez le fil rouge du chapitre : *ce qui n'est pas redessiné dans `Paint` disparaît*, et **chaque  
binaire COM est un choix**, pas une simple conversion.

---

**Partie suivante → Partie 5 — Migrer les données et l'interopérabilité**, en commençant par le  
**module 15 — Accès aux données : d'ADO/DAO/RDO à ADO.NET**.

⏭️ [Accès aux données — d'ADO/DAO/RDO à ADO.NET](/15-acces-donnees/README.md)
