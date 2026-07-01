🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.4 Les options de compilation : `Option Strict` / `Explicit` / `Infer` / `Compare`

> *Partie 2 — Préparer le terrain · Module 6 — Préparer l'environnement .NET*

---

## Objectif de cette section

Quatre directives `Option` règlent à quel point le compilateur VB.NET est **strict**, **typé** et **sensible à la casse**. Bien réglées **dès le départ**, elles évitent deux écueils opposés : un compilateur **trop sévère** qui refuse de compiler le code hérité (riche en liaison tardive), ou **trop laxiste** qui laisse passer des erreurs et **change silencieusement** le comportement.

La ligne directrice de cette section tient en une phrase : **démarrer en `Option Strict Off`** pour accueillir le code existant, puis **l'activer progressivement** (module 17). Et au passage, ne pas se faire piéger par `Option Compare`, dont un mauvais réglage provoque des **régressions invisibles**.

---

## Quatre directives, deux niveaux de réglage

En VB6, ces options se posaient **en tête de chaque module** (`Option Explicit`, `Option Compare Text`…). En VB.NET, elles existent à **deux niveaux** :

1. **Au niveau du projet** : *Propriétés du projet → onglet « Compiler »* (réglage par défaut pour **tous** les fichiers).
2. **Au niveau d'un fichier** : une ligne tout en **haut** du fichier `.vb`, **avant** les `Imports`, qui **surcharge** le réglage du projet pour ce fichier seulement.

```vbnet
' Tout en haut d'un fichier .vb — surcharge le réglage du projet pour CE fichier
Option Strict On
Option Compare Text
```

> 💡 Cette surcharge **par fichier** est précisément le mécanisme qui permettra d'**activer `Option Strict` progressivement** (fichier par fichier) au module 17, sans devoir basculer tout le projet d'un coup.

---

## `Option Explicit` : déclarer ses variables

**Ce qu'elle fait :** impose de **déclarer** (`Dim`) toute variable avant usage. Sans elle, une variable mal orthographiée est créée à la volée — source classique de bugs silencieux.

- **VB6 :** facultative, mais bonne pratique (on la mettait en tête de module). Le § 5.1 recommandait déjà de l'**ajouter partout** dans le code VB6 **avant** de migrer.
- **VB.NET :** **activée par défaut**. On la **garde `On`**.
- **À la migration :** le code VB6 qui n'avait **pas** `Option Explicit` contient des variables implicites qui **ne compileront pas** sous `Explicit On` — et c'est tant mieux : on les déclare, et on élimine au passage les bugs de frappe.

---

## `Option Strict` : le réglage central

C'est **la** directive structurante. Activée (`On`), elle interdit trois choses que VB6 pratiquait abondamment :

1. **Les conversions implicites avec perte** (*narrowing*) : `Double → Integer`, `Long → Integer`, `Object → String`… doivent devenir **explicites** (`CInt`, `CType`, `DirectCast`…).
2. **La liaison tardive** (*late binding*) : appeler un membre sur un `Object` dont le type n'est pas connu à la compilation est **interdit**.
3. **Les déclarations sans type** : un `Dim x` sans clause `As` (qui deviendrait `Object`) est refusé — sauf si `Option Infer` peut déduire le type d'une valeur initiale (voir plus bas).

- **VB6 :** s'appuie massivement sur la liaison tardive (`Variant`, `Object`, `CreateObject`, propriétés par défaut) et les conversions implicites. Le code fraîchement migré présente donc **énormément** de violations de `Strict`.
- **VB.NET :** **désactivée (`Off`) par défaut**… mais l'**objectif final** est `On` : détection des erreurs **à la compilation** plutôt qu'à l'exécution, meilleures performances, plus de conversions qui échouent par surprise.

```vbnet
' Sous Option Strict Off : COMPILE (conversions et liaison tardive implicites)
Dim total As Integer
total = 3.7                 ' conversion implicite Double → Integer (perte !)
Dim doc As Object = OuvrirDocument()
doc.Enregistrer()           ' liaison tardive : résolue à l'exécution

' Sous Option Strict On : ERREURS de compilation
'   • « Option Strict On interdit les conversions implicites de 'Double' en 'Integer' »
'   • « Option Strict On interdit la liaison tardive »
' → corriger : total = CInt(3.7)  …et typer 'doc' (CType/DirectCast) avant l'appel
```

---

## La stratégie ⚠️ : démarrer `Off`, activer **progressivement**

C'est le point à retenir de cette section.

**Pourquoi démarrer en `Option Strict Off` ?** Parce que le code hérité **dépend** de la liaison tardive et des conversions implicites. En `Off`, ce code **compile et s'exécute** : on peut faire tourner l'application migrée vite, et la **valider** contre le *golden master* (§ 5.5). C'est le « pont ».

**Pourquoi ne pas activer `On` tout de suite ?** Parce qu'on obtiendrait **des milliers d'erreurs** d'un coup, et surtout on **mélangerait** deux chantiers : la migration du langage **et** le durcissement du typage. Or on ne change **qu'une variable à la fois**.

**Comment activer ensuite ?** **Progressivement**, fichier par fichier (grâce à la surcharge `Option Strict On` en tête de fichier), en corrigeant les violations et en **éliminant la liaison tardive résiduelle**. C'est l'objet du **§ 17.1**.

> ⚠️ **`Option Strict Off` est un compromis temporaire et assumé, pas un renoncement.** On démarre `Off` pour avancer, on finit `On` partout. Rester `Off` indéfiniment, ce serait conserver durablement les risques (conversions silencieuses, erreurs d'exécution) que `Strict On` est fait pour éliminer.

> 💡 Les outils de migration (module 4) génèrent d'ailleurs en général le projet avec **`Option Strict Off`** et **`Option Explicit On`** — exactement le point de départ recommandé ici.

---

## `Option Infer` : l'inférence de type local

**Ce qu'elle fait :** permet d'écrire `Dim x = 5` et de laisser le compilateur **inférer** le type (`Integer`) à partir de la valeur initiale. La variable reste **fortement typée** — le type est simplement déduit, pas absent. Notion **nouvelle** en VB.NET (elle n'existe pas en VB6).

- **VB.NET :** **activée par défaut** dans les projets créés par Visual Studio.

Son effet dépend de son interaction avec `Option Strict`, pour une déclaration **sans type** munie d'une valeur initiale :

| `Dim x = 5` | `Infer On` | `Infer Off` |
|---|---|---|
| **`Strict On`** | `x` inféré `Integer` ✅ | **Erreur** (clause `As` requise) |
| **`Strict Off`** | `x` inféré `Integer` | `x` typé `Object` (liaison tardive) |

> 💡 En pratique, le **vrai levier reste `Option Strict`**. On peut **laisser `Infer On`** (c'est le défaut moderne, sans danger pour le code migré déjà typé) ; l'inférence concerne surtout le **code idiomatique** que vous écrirez **après** la migration.

---

## `Option Compare` : Binary vs Text — **un piège silencieux !**

**Ce qu'elle fait :** règle la **comparaison de chaînes** (et le tri, l'opérateur `Like`, et les fonctions de chaînes) :

- **`Binary`** : comparaison **sensible à la casse**, fondée sur la valeur des codes de caractères. `"Dupont" = "DUPONT"` → **False**.
- **`Text`** : comparaison **insensible à la casse**, selon les règles de la culture. `"Dupont" = "DUPONT"` → **True**.

```vbnet
If nomClient = "DUPONT" Then ...   ' le RÉSULTAT dépend de Option Compare !
```

- **VB6 :** défaut **`Binary`**, sauf `Option Compare Text` posé en tête de module.
- **VB.NET :** défaut **`Binary`** également (réglable au projet, surchargeable par fichier).

> ⚠️ **Le danger de migration.** Un module VB6 qui reposait sur **`Option Compare Text`** (comparaisons insensibles à la casse, fréquentes en logique métier) migré dans un projet par défaut **`Binary`** voit ses comparaisons **passer silencieusement** d'insensibles à **sensibles** à la casse. Aucune erreur de compilation — juste un **comportement qui change**. C'est exactement le genre de **régression invisible** que cette formation traque (Annexe B). 

**La règle :** **préserver le `Option Compare` effectif de chaque module d'origine.** Là où un module VB6 utilisait `Option Compare Text`, on remet `Option Compare Text` (au niveau du fichier, ou du projet selon la majorité des cas).

> 🔗 Au-delà de l'opérateur `=`, `Option Compare` influence `Like`, et les fonctions `InStr`, `StrComp`, `Replace`, `Split`/`Filter`… dont l'argument de comparaison **par défaut** suit ce réglage. Pensez aussi aux conversions sensibles à la **culture** (séparateur décimal, dates) — piège **B.10**. Le *golden master* (§ 5.5) est votre filet pour détecter ces écarts.

---

## Réglages recommandés pour démarrer

| Option | Réglage **de départ** | Réglage **cible** (après module 17) |
|---|---|---|
| `Explicit` | **On** *(garder)* | On |
| `Strict` | **Off** *(temporaire)* | **On** partout |
| `Infer` | **On** | On |
| `Compare` | **= comportement VB6 d'origine** *(souvent `Binary` ; `Text` là où le module VB6 l'utilisait)* | identique |

L'état de départ accueille le code hérité ; l'état cible, atteint **après validation** via l'activation progressive de `Strict` (§ 17.1), élimine la liaison tardive résiduelle.

---

## Où régler tout ça

Le réglage se fait dans *Propriétés du projet → onglet « Compiler »*, et se reflète dans le fichier projet :

```xml
<OptionExplicit>On</OptionExplicit>
<OptionStrict>Off</OptionStrict>
<OptionInfer>On</OptionInfer>
<OptionCompare>Binary</OptionCompare>
```

> 💡 Comme pour la cible du framework (§ 6.1), ces lignes sont **lisibles en gestion de versions** (§ 6.5) : on peut ainsi **suivre dans l'historique** le moment où `Option Strict` bascule de `Off` à `On`, projet ou fichier par fichier.

---

## Points de vigilance

> ⚠️ **`Strict Off` au départ, `On` à l'arrivée.** Compromis temporaire pour accueillir la liaison tardive héritée ; activation **progressive** ensuite (§ 17.1). Ne pas rester `Off` par confort.

> ⚠️ **`Option Compare` : préserver le comportement d'origine.** Un module passé de `Text` à `Binary` change silencieusement de casse → régression. À reproduire fidèlement, fichier par fichier si besoin.

> 💡 **`Explicit On` partout.** À garder activée ; elle révèle les variables implicites du VB6 sans `Option Explicit` (à déclarer — § 5.1).

> 💡 **`Infer` n'est pas le sujet principal.** On la laisse `On` (défaut moderne) ; le vrai levier de rigueur est `Strict`.

---

## À retenir

> - Quatre directives `Option` réglées **au projet** (onglet « Compiler ») et **surchargeables par fichier** (ligne en tête, avant `Imports`).
> - **`Explicit` : On** (garder) — révèle et élimine les variables implicites (§ 5.1).
> - **`Strict` : le réglage central.** Démarrer **`Off`** pour accueillir la liaison tardive héritée (le « pont »), puis l'activer **progressivement** (`On` par fichier) au **§ 17.1** — un seul changement à la fois.
> - **`Infer` : On** (défaut moderne, sans danger) ; concerne surtout le code idiomatique **post-migration**. Son effet dépend de `Strict` pour les `Dim` sans type.
> - **`Compare` : piège silencieux.** `Binary` (sensible à la casse) vs `Text` (insensible). **Préserver le réglage de chaque module VB6**, sinon régression invisible (Annexe B) ; le *golden master* (§ 5.5) la détecte.

---

## Liens utiles

- **§ 5.1** — Nettoyer le code VB6 (`Option Explicit` partout) *(en amont de la migration)*
- **§ 6.1 / 6.5** — Ciblage du framework et gestion de versions *(réglages lisibles dans le `.vbproj`)*
- **§ 17.1** — Activer `Option Strict` **progressivement** *(la suite directe de cette section)*
- **§ 5.5** — *Golden master* *(détecter les régressions de comparaison et de conversion)*
- **§ 2.3** — Le typage (`Variant`, entiers redimensionnés, `Currency`) *(ce que `Strict` met en lumière)*
- **Annexe B** — Catalogue des pièges silencieux *(dont B.10 : conversions sensibles à la culture)*

⏭️ [Stratégie de gestion de versions (Git, branche de migration, jalons)](/06-preparer-environnement/05-gestion-versions.md)
