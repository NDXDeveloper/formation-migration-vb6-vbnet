🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7. Types, variables et déclarations ⭐ ⚠️

> **Partie 3 — Migrer le langage (le cœur de la formation)**
>
> Premier chapitre du cœur de la formation. Avant les procédures, les classes ou les
> formulaires, on s'attaque à la **brique de base** de tout programme : les **types**, les
> **variables** et leurs **déclarations**.

---

## 🧭 Pourquoi commencer par les types ?

Parce que **tout le reste en dépend**. Une procédure manipule des variables ; une classe expose  
des propriétés typées ; un formulaire stocke un état. Si le **socle des types** se comporte  
différemment après migration, ce sont des **dizaines de bugs en cascade** — souvent invisibles à  
la compilation.

Et c'est précisément là que se niche la spécificité de cette migration : en VB6 comme en VB.NET,  
on écrit `Dim`, `Integer`, `Date`, `Variant`… **les mêmes mots-clés**. Mais derrière ces mots  
identiques, la **taille**, la **sémantique** et le **comportement à l'exécution** ont parfois  
changé.

> ⚠️ **Le code qui ne compile pas, on le voit. Le code qui compile mais ne calcule plus pareil,
> on ne le voit pas.**
> Ce chapitre est l'un des principaux fournisseurs de ces *changements silencieux*
> (voir **Annexe B — Catalogue des pièges silencieux**).

---

## 🧨 Le piège fondateur : `Integer` n'a pas la même taille

S'il ne fallait retenir qu'une seule chose de ce chapitre, ce serait celle-ci :

| Mot-clé | En VB6 | En VB.NET |
|---------|--------|-----------|
| `Integer` | entier **16 bits** (-32 768 → 32 767) | entier **32 bits** (≈ ±2,1 milliards) |
| `Long` | entier **32 bits** | entier **64 bits** |

Autrement dit : un `Dim compteur As Integer` **recopié tel quel** compile sans broncher… mais ne  
désigne plus le même type. La plupart du temps, « ça marche » — jusqu'au jour où un calcul, un  
appel d'API Win32 ou une lecture binaire dépend de la **taille exacte**. Le bon réflexe de  
migration est de raisonner en **équivalence de taille**, pas en équivalence de **nom** :

- VB6 `Integer` (16 bits) → VB.NET **`Short`**
- VB6 `Long` (32 bits) → VB.NET **`Integer`**

> 🔎 C'est le sujet de la section **7.2**, signalée ⭐ ⚠️ : à elle seule, elle justifie une
> relecture attentive de **chaque** déclaration numérique.

---

## 🗺️ Les trois familles de changements

Tout au long du chapitre, les difficultés se rangent dans **trois catégories**. Les reconnaître  
aide à savoir *à quel point* se méfier :

1. **Même nom, taille ou sémantique différente** — *le plus dangereux*, car silencieux.
   `Integer`, `Long`, le booléen `True`…
2. **Un concept qui disparaît** — la migration *force* un choix (compatibilité ou refactoring).
   `Variant`, `Currency`, les chaînes de **longueur fixe** (`String * n`).
3. **Des états qui fusionnent** — plusieurs notions VB6 se ramènent à une seule en .NET.
   `Empty`, `Null`, `Missing`, `Nothing`.

---

## 📚 Ce que couvre ce chapitre

Les sections suivantes prennent chaque pièce du socle, une à une :

- **[7.1 — `Variant` → `Object`](01-variant-object.md).** Le conteneur universel de VB6 devient
  `Object`. Conversions, *boxing* des types valeur, et le choix entre `CType` (qui **convertit**)
  et `DirectCast` (qui **caste** strictement).
- **[7.2 — Entiers redimensionnés](02-entiers-redimensionnes.md).** ⭐ ⚠️ `Integer`→`Short`,
  `Long`→`Integer`, `Currency`→`Decimal` : tailles, débordements et la bonne table d'équivalence.
- **[7.3 — Le type `Date`](03-dates.md).** ⚠️ En VB6, une date **est** un `Double` ; en .NET,
  c'est un vrai type `DateTime`. Conséquences sur les conversions et l'arithmétique de dates.
- **[7.4 — Chaînes de longueur fixe](04-chaines-longueur-fixe.md).** ⚠️ `Dim s As String * 10`
  n'a **pas** d'équivalent natif : attribut `VBFixedString` (compatibilité) ou refactoring vers
  une `String` ordinaire.
- **[7.5 — Déclarations](05-declarations.md).** Le `Dim` multiple (un faux-ami :
  `Dim a, b As Integer` ne type pas `a` de la même façon dans les deux langages !), les directives
  `DefInt`/`DefStr`, la portée, et le `As New` (instanciation implicite).
- **[7.6 — Constantes, énumérations et booléen](06-constantes-enums-bool.md).** ⚠️ `Const`, `Enum`
  (désormais adossé à un type entier), et la fameuse question du `True = -1`.
- **[7.7 — `Nothing`, `Empty`, `Null`, `Missing`](07-nothing-empty-null.md).** ⚠️ La fin du
  `Variant` et de ses états distincts : ce qui se ramène à `Nothing`, et où réapparaît la notion
  de « valeur nulle » (`DBNull`, côté données).

---

## 🧾 Aide-mémoire express

À garder sous les yeux pendant la lecture (version détaillée en **Annexe D — Correspondance des
types de données**) :

| Type VB6 | Cible VB.NET conseillée | Point de vigilance |
|----------|-------------------------|--------------------|
| `Variant` | `Object` | conversions implicites, *boxing* |
| `Integer` (16 bits) | `Short` | ⚠️ `Integer` en .NET = 32 bits |
| `Long` (32 bits) | `Integer` | débordement / taille mémoire |
| `Currency` | `Decimal` | pas une copie bit à bit ; précision monétaire |
| `Date` | `DateTime` | vrai type, ce n'est plus un `Double` |
| `String * n` | `String` (+ `VBFixedString` si besoin) | longueur fixe non native |
| `Boolean` (`True` = -1) | `Boolean` | conversions numériques |
| `Empty` / `Null` / `Missing` | `Nothing` (et `DBNull`) | états fusionnés |

> ⚠️ Ce tableau donne la **cible**, pas une recette automatique. Le bon type dépend de l'**usage
> réel** de la variable (calcul, interop, stockage, affichage).

---

## 🎯 Objectifs du chapitre

À l'issue de ce chapitre, tu sauras :

- **identifier** les types VB6 dont le nom est trompeur en .NET, et choisir la bonne équivalence
  par la **taille** plutôt que par le mot-clé ;
- **anticiper** les pièges silencieux liés aux types (débordements, dates, booléens, états du
  `Variant`) avant qu'ils ne deviennent des régressions ;
- **décider**, pour chaque cas sans équivalent direct (`Variant`, `Currency`, chaînes de longueur
  fixe), entre **compatibilité** (béquille temporaire) et **refactoring** propre ;
- **relire** une déclaration VB6 (`Dim`, `Const`, `Enum`, `As New`…) en sachant exactement ce qui
  change de sens une fois en VB.NET.

---

## 🔗 À garder ouvert en parallèle

- **Annexe B — Catalogue des pièges silencieux** : les entrées **B.2** (entiers redimensionnés),
  **B.5** (`Variant` : `Empty`/`Null`/`Missing`), **B.8** (`True = -1`) et **B.9** (dates et
  arrondis) recoupent directement ce chapitre.
- **Annexe D — Correspondance des types de données** : la table de référence complète (tailles,
  plages, fonctions de conversion `CType`/`DirectCast`/`CInt`/`CLng`/`CDec`).
- **Chapitre 5 — Préparer le code VB6** : une partie de ces pièges se **désamorce en amont**
  (typage explicite côté VB6) avant même de migrer.

> 💡 Et n'oublie pas le **golden master** (chapitre 5.5) : c'est le filet qui révèle, *chiffres à
> l'appui*, qu'un type a silencieusement changé le résultat d'un calcul.

---

## ➡️ Pour commencer

On attaque par le cas le plus emblématique du « monde VB6 » : le **`Variant`**, ce type qui  
pouvait tout contenir, et sa traduction en `Object`.

➡️ **[7.1 — `Variant` → `Object` : conversions, boxing, `CType`/`DirectCast`](01-variant-object.md)**

⏭️ [`Variant` → `Object` : conversions, boxing, `CType`/`DirectCast`](/07-types-variables/01-variant-object.md)
