🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.4 — Traquer les pièges silencieux (ByRef, arrondis, dates, bornes de tableaux, finalisation) ⭐ ⚠️

**Chasser méthodiquement les différences de comportement qui *compilent proprement* mais cassent l'équivalence : le compilateur ne peut pas les voir — il faut les traquer délibérément, à l'aide du catalogue de l'annexe B.**

> ⭐ ⚠️ Le pendant **diagnostic** du module 17.3. Le *golden master* **détecte** qu'un scénario a
> divergé ; cette section apprend à **identifier pourquoi** — et à débusquer les pièges de façon
> **proactive**, avant même qu'un test ne les révèle.

---

## Pourquoi une chasse **dédiée**

Le module 17.1 (`Option Strict On`) a déjà fait remonter une classe de défauts **à la compilation** :  
liage tardif et conversions implicites. Mais les pièges visés ici sont d'une autre nature : **ils  
compilent sans la moindre erreur** et **s'exécutent sans exception** — ils produisent simplement un  
résultat **différent**. C'est exactement ce qui les rend **silencieux**, et c'est pourquoi ni le  
compilateur, ni `Option Strict`, ne les attrapent. Seule une **traque délibérée** y parvient.

> 🎯 **Division du travail** : le module 17.1 neutralise ce que le **compilateur** peut voir ; le module
> 17.4 traque ce qu'il **ne peut pas** voir. Les deux sont nécessaires.

---

## Deux modes de chasse, complémentaires

1. **Chasse réactive (diagnostique)** — quand le *golden master* (17.3) signale un écart, remonter du
   **symptôme** au **piège probable** (table ci-dessous).
2. **Chasse proactive (statique)** — parcourir le code à la recherche des **motifs** caractéristiques de
   chaque piège, **indépendamment** de la couverture des tests (liste d'audit plus bas).

> 💡 **Pourquoi les deux ?** Le *golden master* ne voit que les chemins que le corpus **exerce** (ses
> angles morts). La chasse statique, elle, repère les motifs **partout** — mais ne confirme pas le
> comportement. Ensemble, ils sont complets : la chasse statique trouve les **candidats**, le *golden
> master* **confirme** la correction.

---

## Diagnostic : symptôme → piège probable

| Symptôme observé (souvent via le *golden master*) | Piège probable | Origine |
|----|----|----|
| Total monétaire faux d'un centime ; dernier décimal différent | Arrondis / `Currency`→`Decimal` | B.9 · 7.2/7.3 |
| Décalage d'un cran ; `IndexOutOfRangeException` | Bornes de tableaux (0-based) | B.4 · 8.1 |
| Une modification d'argument est **« perdue »** | `ByRef` devenu `ByVal` | B.1 · 2.4/10.2 |
| **Corruption mémoire** à un appel d'API | `ByRef`/`ByVal` à la frontière native | B.1 · 16.4 |
| Fichier resté **verrouillé** ; ordre d'effets changé | Finalisation non déterministe | B.3 · 2.2/12.3 |
| Date fausse après un calcul | Dates (`Double` → `DateTime`) | B.9 · 7.3 |
| Résultat faux **seulement sous certaines locales** | Conversions sensibles à la culture | B.10 |
| Erreurs **silencieusement ignorées** | `On Error Resume Next` | B.6 · 11.3 |
| `True` converti en 1 au lieu de -1 ; logique booléenne inversée | `True = -1` / conversions booléennes | B.8 · 7.6 |
| Comparaison d'objets qui « marche par hasard » | Propriétés par défaut / `Set` | B.7 · 2.5/10.6 |
| `Null` / `Empty` / `Missing` mal gérés | États du `Variant` perdus | B.5 · 7.7 |

---

## Les cinq pièges du titre

L'annexe **B** donne pour chacun le triptyque **symptôme / cause / correction** détaillé ; voici  
l'essentiel et, surtout, **comment les traquer**.

### ByRef devenu ByVal (B.1)

- **Symptôme** : une procédure censée modifier la variable de l'appelant ne le fait plus (la
  modification est « perdue ») ; à la frontière d'API, **corruption mémoire**.
- **Cause** : VB6 passait `ByRef` **par défaut**, VB.NET passe `ByVal` par défaut. Toute migration
  naïve **inverse** la sémantique des paramètres qui s'appuyaient sur ce défaut (modules 2.4, 10.2 ;
  frontière native au 16.4).

```vbnet
' VB6 : "valeur" est ByRef par défaut → la procédure modifie bien l'appelant
Sub Doubler(valeur As Long)
    valeur = valeur * 2
End Sub

' ❌ Migration naïve : ByVal par défaut → la modification est PERDUE
Sub Doubler(valeur As Integer) ' ...

' ✅ Rétablir l'intention explicitement
Sub Doubler(ByRef valeur As Integer)
    valeur = valeur * 2
End Sub
```

- **Où chercher** : tout paramètre **affecté dans le corps** d'une procédure ; **chaque**
  `Declare`/`DllImport` ; les paramètres sans `ByVal`/`ByRef` explicite dans le VB6 d'origine.

### Arrondis monétaires (B.9)

- **Symptôme** : totaux faux d'un centime, écarts sur le dernier décimal.
- **Nuance importante** : VB6 **et** VB.NET utilisent par défaut l'**arrondi bancaire** (au pair le plus
  proche) dans leurs fonctions de conversion/arrondi (`CInt`, `Round`…). Une migration fidèle le
  **préserve** donc — ⚠️ **ne le « corrigez » pas** en croyant à un bug. Les vrais dangers sont
  ailleurs : (a) `Currency` mappé en **`Double`/`Single`** (erreur d'arrondi flottant) au lieu de
  **`Decimal`** (module 7.2) ; (b) un **mode d'arrondi changé** (`MidpointRounding.AwayFromZero`,
  `Format`…) ; (c) la **précision des calculs intermédiaires** qui diffère.
- **Où chercher** : tous les calculs **monétaires** ; `Currency` non mappé en `Decimal` ; les appels
  d'arrondi et leur **mode**.

### Dates (B.9)

- **Symptôme** : dates fausses après calcul ; analyse/format de date divergent.
- **Cause** : la `Date` VB6 est un **`Double`** (date sérielle, époque **1899-12-30**) ; `DateTime` est
  un type distinct. L'**arithmétique brute** sur dates ne se transpose pas (module 7.3).

```vbnet
' VB6 : on ajoutait des jours via la nature Double de Date
dateFin = dateDebut + 1.5      ' 1,5 jour

' ✅ VB.NET : méthodes de DateTime (pas d'addition brute de Double)
dateFin = dateDebut.AddDays(1.5)
```

- Surveillez aussi : l'**interprétation des années à deux chiffres**, l'**analyse sensible à la culture**
  (`CDate`, recoupe B.10), et les **dates nulles** (une `Date` ne peut pas être `Null` → `Nullable(Of
  DateTime)`, recoupe B.5). Pour les valeurs sérielles, `ToOADate`/`FromOADate` (module 16.5) rendent la
  conversion explicite.
- **Où chercher** : **arithmétique** sur dates (`+`/`-` bruts) ; `CDate`/`Format` dépendants de la
  locale ; dates non initialisées / nullables.

### Bornes de tableaux (B.4)

- **Symptôme** : décalage d'un cran, `IndexOutOfRangeException`, ou résultat silencieusement faux.
- **Nuance** : `Dim a(10)` crée **11 éléments** (0 à 10) en VB.NET **comme** en VB6 (`Option Base 0`) —
  cas **cohérent**. Le piège vient des tableaux **1-based** : `Option Base 1`, ou `Dim a(1 To 10)`. Or
  **VB.NET impose des tableaux 0-based** (`Dim a(1 To 10)` ne compile pas), et **`LBound` renvoie
  toujours 0**. Tout code 1-based doit être réécrit (module 8.1).
- **Où chercher** : `Option Base 1` ; bornes basses explicites (`1 To n`) ; `LBound`/`UBound` ; boucles
  `For i = 1 To …` ; `ReDim`.

### Finalisation (B.3)

- **Symptôme** : ressources libérées trop tard ou jamais (fichiers verrouillés, connexions épuisées,
  *handles* fuités) ; effets de bord d'un `Class_Terminate` qui apparaissent au **mauvais moment** ou
  dans le **mauvais ordre**.
- **Cause** : finalisation **déterministe** VB6 (`Class_Terminate` au passage du compteur à 0) →
  finalisation **non déterministe** par le GC (modules 2.2, 12.2). Toute logique qui dépendait du
  **moment** de la destruction est désormais fausse.
- **Correction** : `IDisposable` / `Using` (module 12.3) ; déplacer la logique de `Class_Terminate` dans
  `Dispose`. Le **conséquence ressource** (fuites) est traitée au module 17.6 ; ici, on vise la
  **différence de comportement** (timing/ordre).
- **Où chercher** : **chaque** `Class_Terminate` du VB6 d'origine ; objets tenant des ressources rares
  ou non managées.

---

## Ne pas oublier le reste du catalogue (annexe B)

Les cinq ci-dessus sont les plus emblématiques, mais l'audit doit couvrir **tout** l'annexe B :

- **B.5 — `Variant` : `Empty`/`Null`/`Missing` perdus** (module 7.7) : divergences de gestion du « rien ».
- **B.6 — `On Error Resume Next`** (module 11.3) : il **avalait** les erreurs en VB6 ; reconduit
  naïvement, il **masque** des bugs — et l'ensemble des erreurs levées diffère (.NET lève des exceptions
  là où VB6 pouvait ne rien signaler). **Auditez chaque occurrence.**
- **B.7 — Propriétés par défaut et `Set` disparu** (modules 2.5, 10.6).
- **B.8 — `True = -1` et conversions booléennes** (modules 7.6, 16.5).
- **B.10 — Conversions sensibles à la culture** (`CDbl`/`CDate`/`Format`/`Val`) : recoupe les dates et
  les arrondis ; un résultat faux **seulement sous certaines locales** en est la signature.

---

## La chasse proactive : liste d'audit « quoi rechercher »

À utiliser comme **grille de recherche** sur le VB6 d'origine **et** le code migré :

| Piège | Motif à rechercher |
|----|----|
| **ByRef/ByVal** (B.1) | Paramètres **affectés dans le corps** ; tout `Declare`/`DllImport` ; paramètres sans `ByVal`/`ByRef` explicite côté VB6 |
| **Arrondis / Currency** (B.9) | Calculs **monétaires** ; `Currency` mappé en `Double`/`Single` ; appels d'arrondi (`Math.Round`, `Format`) et leur **mode** |
| **Dates** (B.9) | **Arithmétique** sur dates (`+`/`-` bruts) ; `CDate`/`Format` sans culture ; dates non initialisées/nullables |
| **Bornes de tableaux** (B.4) | `Option Base 1` ; bornes basses (`1 To n`) ; `LBound`/`UBound` ; boucles `For i = 1 To …` ; `ReDim` |
| **Finalisation** (B.3) | Chaque `Class_Terminate` ; objets à ressources rares/non managées |
| **Variant** (B.5) | `IsNull`/`IsEmpty`/`IsMissing` ; `Variant`/`Object` recevant `Null`/`Empty` |
| **On Error Resume Next** (B.6) | **Chaque** `On Error Resume Next` |
| **Propriétés par défaut** (B.7) | Objets utilisés sans membre explicite ; ancien mot-clé `Set` |
| **Booléens** (B.8) | Conversions `Boolean` ↔ entier ; comparaisons à `-1`/`True` |
| **Culture** (B.10) | `CDbl`/`CDate`/`Format`/`Val` sans culture explicite |

> 🔄 **La boucle complète** : la chasse statique repère un **candidat** → on corrige → le *golden
> master* (17.3) **confirme** que le comportement est rétabli (et qu'on n'a rien cassé d'autre).

---

## En résumé

- Ces pièges **compilent proprement** et **ne lèvent pas d'exception** : ni le compilateur ni
  `Option Strict` (17.1) ne les voient. Il faut les **traquer délibérément**.
- **Deux modes complémentaires** : **réactif** (symptôme → piège, à partir des écarts du *golden
  master*) et **proactif** (recherche des motifs, indépendamment de la couverture des tests).
- **Les cinq pièges du titre** : `ByRef`→`ByVal` (modification perdue / corruption), **arrondis** (ne pas
  « corriger » l'arrondi bancaire ; mapper `Currency` en `Decimal`), **dates** (pas d'arithmétique brute
  → `AddDays`…), **bornes** (tout 1-based à réécrire ; `LBound` = 0), **finalisation** (`Class_Terminate`
  → `IDisposable`/`Using`).
- **L'audit couvre tout l'annexe B** (B.1–B.10), `On Error Resume Next` (B.6) en tête des oublis
  dangereux.
- La **boucle** : chasse statique → correction → confirmation par le *golden master* (17.3).

> ➡️ **Une fois la correction prouvée, place à la récompense : passer d'une traduction littérale à du
> VB.NET idiomatique — LINQ, génériques, espace `My`, code moderne.** (Module 17.5)

---

🏷️ **Indicateurs** : ⭐ Point clé · ⚠️ Pièges (le cœur du risque) · application du catalogue de l'**annexe B** ; lié aux modules 2.2/2.4/2.5, 7.2/7.3/7.6/7.7, 8.1, 10.2/10.6, 11.3, 12.2/12.3, 16.4/16.5, 17.3
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Refactoring idiomatique : LINQ, génériques, espace `My`, code VB.NET moderne](/17-valider-refactoriser/05-refactoring-idiomatique.md)
