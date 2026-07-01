🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.3 — `Optional` et valeurs par défaut obligatoires ; fin de `IsMissing` ⚠️

> **`Optional` survit, mais sous contrainte.** En VB.NET, tout paramètre optionnel **doit** porter une valeur par défaut, et `IsMissing` **disparaît**. Conséquence sournoise : on ne peut plus distinguer « argument **omis** » de « argument **fourni égal au défaut** ».

📍 *Module 10 · § 10.3 · [↑ Introduction du chapitre](README.md) · [← § 10.2](02-byref-byval-audit.md)*

---

## ⚡ En bref

- **Deux changements liés.** (1) En VB.NET, un paramètre `Optional` **doit** déclarer une valeur par défaut — sinon **erreur de compilation**. (2) La fonction `IsMissing` **n'existe plus**.
- **Le vrai piège est sémantique** : dès qu'un paramètre a un défaut, l'intérieur de la procédure ne sait plus si l'appelant a **omis** l'argument ou s'il a **passé la valeur du défaut**. VB6 le distinguait (pour les `Variant`) via `IsMissing`.
- **Contrainte supplémentaire** : la valeur par défaut doit être une **constante de compilation** (`Now`, `New …()`, un appel de fonction sont **interdits**).
- **Trois remplacements propres** : type **`Nullable(Of T)`** (recommandé pour les valeurs), **surcharge** (la plus explicite), ou **valeur sentinelle** (simple mais piégeuse).

---

## 1. Deux changements, à traiter ensemble

| Aspect | VB6 | VB.NET (4.7.2) |
|---|---|---|
| `Optional` sans valeur par défaut | **Autorisé** | ❌ **Erreur de compilation** |
| Valeur par défaut | Facultative | **Obligatoire** |
| Forme du défaut | littéral | **constante de compilation uniquement** |
| `IsMissing(arg)` | Disponible (utile pour `Optional … As Variant`) | **Supprimé** |
| Distinguer « omis » de « fourni = défaut » | Possible (via `IsMissing`, pour `Variant`) | **Impossible** avec un simple défaut |
| « Tous les paramètres après un `Optional` sont `Optional` » | Oui | Oui *(inchangé)* |

Les deux premiers changements sont **« bruyants »** (le code ne compile pas, vous êtes prévenu). Le changement sur `IsMissing` est **« silencieux »** : le code peut compiler tout en ayant perdu une distinction dont la logique dépendait.

---

## 2. La valeur par défaut devient obligatoire

En VB6, on pouvait déclarer un paramètre `Optional` **sans** valeur par défaut :

```vb
' VB6 — Optional sans valeur par défaut : autorisé
Sub Afficher(Optional titre As String)
    ' 'titre' vaut "" s'il est omis (type spécifique → IsMissing inutile ici)
End Sub
```

En VB.NET, la même déclaration **ne compile pas** :

```vb
' VB.NET — ❌ erreur : « une valeur par défaut est attendue »
Sub Afficher(Optional ByVal titre As String)
End Sub

' VB.NET — ✅ corrigé : une valeur par défaut explicite
Sub Afficher(Optional ByVal titre As String = "")
End Sub
```

### ⚠️ Le défaut doit être une **constante de compilation**

C'est le piège qui surprend dès les premiers paramètres optionnels migrés. La valeur par défaut doit être **calculable à la compilation** : un littéral, une constante, un membre d'énumération, `Nothing`. **Sont interdits** : un appel de fonction, `Now`, `New …()`, ou toute expression évaluée à l'exécution.

```vb
' VB.NET — ❌ ne compile pas : le défaut n'est pas constant
Sub Journaliser(Optional ByVal horodatage As Date = Now)            ' ❌ Now
Sub Creer(Optional ByVal options As Config = New Config())          ' ❌ New
```

Le contournement consiste à donner une **sentinelle constante** comme défaut, puis à **calculer la vraie valeur dans le corps** :

```vb
' VB.NET — ✅ sentinelle + calcul à l'exécution
Sub Journaliser(Optional ByVal horodatage As Date = Nothing)
    If horodatage = Nothing Then horodatage = Now   ' Nothing → Date.MinValue (#0001-01-01#)
    ' ...
End Sub
```

> 💡 `Nothing` affecté à un **type valeur** vaut le **défaut du type** (`0`, `False`, `Date.MinValue`…). C'est une sentinelle pratique — mais elle confond « omis » avec « l'appelant a passé cette valeur par défaut » (voir § 3 et § 4).

---

## 3. La fin de `IsMissing` : le vrai changement sémantique

En VB6, `IsMissing` permettait — **pour un `Optional … As Variant`** — de savoir si l'appelant avait **réellement omis** l'argument, indépendamment de toute valeur par défaut :

```vb
' VB6 — IsMissing distingue "omis" de "fourni"
Sub Configurer(Optional niveau As Variant)
    If IsMissing(niveau) Then
        niveau = 3              ' défaut "calculé" seulement si réellement omis
    End If
    ' ... utiliser niveau ...
End Sub
```

En VB.NET, `IsMissing` **disparaît**, et avec un simple défaut, la distinction est **perdue** :

```vb
' VB.NET — avec un défaut nu, deux appels deviennent indistinguables
Sub Configurer(Optional ByVal niveau As Integer = 3)
    ' Configurer()    → niveau = 3
    ' Configurer(3)   → niveau = 3   ← IMPOSSIBLE de savoir lequel des deux
End Sub
```

Tant que la logique se contente d'« utiliser la valeur », cette perte est **sans conséquence**. Elle devient un **bug** dès que le code faisait quelque chose de **spécifique au cas « omis »** (journaliser l'absence, déclencher un comportement différent, calculer un défaut dynamique…).

> 🔗 Cette disparition fait partie de la fin plus large des états du `Variant` (`Empty`/`Null`/`Missing`), traitée au **[§ 7.7](../07-types-variables/07-nothing-empty-null.md)** et cataloguée en **[Annexe B.5](../annexes/pieges-silencieux/README.md)**.

---

## 4. Les stratégies de remplacement

Quand la distinction « omis vs fourni » **compte**, il faut la **reconstruire** explicitement. Trois approches propres, plus une transcription minimale.

### 4.1 Type `Nullable(Of T)` — recommandé pour les valeurs

Pour les types valeur (`Integer`, `Date`, `Boolean`, `Decimal`…), un paramètre **nullable** avec défaut `Nothing` rend « non fourni » à nouveau **distinguable**, **sans collision** :

```vb
' VB.NET — Nullable : "non fourni" redevient explicite
Sub Configurer(Optional ByVal niveau As Integer? = Nothing)
    If Not niveau.HasValue Then
        niveau = 3                  ' réellement omis
    End If
    ' ... niveau.Value ...
End Sub
```

`Integer?` (= `Nullable(Of Integer)`) est disponible sous .NET Framework 4.7.2. C'est l'équivalent le plus **fidèle et robuste** du « missing » d'un `Variant`, tout en **gardant le typage fort**.

### 4.2 Surcharge — la plus explicite

Supprimer carrément le paramètre optionnel au profit de **plusieurs surcharges** : l'intention « avec » et « sans » devient deux signatures distinctes, sans aucune sentinelle.

```vb
' VB.NET — surcharge (nouveauté .NET → § 10.5)
Sub Configurer()
    Configurer(3)                   ' "sans argument" = comportement par défaut
End Sub

Sub Configurer(ByVal niveau As Integer)
    ' ...
End Sub
```

> La **surcharge** n'existait pas en VB6 ; elle est introduite au **[§ 10.5](05-surcharge-retour.md)**. C'est souvent la solution la plus lisible quand « avec » et « sans » correspondent à deux comportements clairs.

### 4.3 Valeur sentinelle — simple, mais piégeuse

Choisir comme défaut une valeur **hors du domaine légitime** et la traiter comme « non fourni » :

```vb
' VB.NET — sentinelle : attention à la collision
Sub Configurer(Optional ByVal niveau As Integer = -1)
    If niveau = -1 Then niveau = 3
End Sub
```

⚠️ Ne fonctionne que si la sentinelle (`-1` ici) **ne peut jamais** être une valeur réelle. Si `-1` est une entrée valide, le code se trompe **silencieusement**.

### 4.4 `Object = Nothing` — la transcription minimale du `Variant`

Puisque `IsMissing` portait sur des `Variant`, la traduction littérale conserve `Object` :

```vb
' VB.NET — proche du Variant, mais perd le typage fort
Sub Configurer(Optional ByVal niveau As Object = Nothing)
    If niveau Is Nothing Then niveau = 3
    ' ...
End Sub
```

⚠️ Double inconvénient : on **perd le type** (retour à `Object`, donc *late binding* et conversions → voir **[§ 7.1](../07-types-variables/01-variant-object.md)**), et `Nothing` redevient une **sentinelle ambiguë** si l'appelant peut légitimement passer `Nothing`. À réserver à une transcription rapide, **à raffiner ensuite**.

### Comparatif

| Stratégie | Typage | Robustesse | Quand l'employer |
|---|---|---|---|
| **`Nullable(Of T)`** (`Integer?`) | Fort | Élevée (`Nothing` = non fourni, sans collision) | Types valeur — **recommandé** |
| **Surcharge** | Fort | Élevée (aucune sentinelle) | Quand « avec / sans » sont deux comportements nets (→ § 10.5) |
| **Sentinelle** | Fort | Moyenne (collision possible) | Quand une valeur hors domaine existe vraiment |
| **`Object = Nothing`** | **Faible** | Moyenne (collision si `Nothing` est légitime) | Transcription express d'un `Variant`, à raffiner |

---

## 5. Détails qui complètent

- **Mode de passage des paramètres optionnels.** Un `Optional` doit aussi recevoir un `ByVal`/`ByRef` **explicite** — la plupart des entrées optionnelles sont `ByVal`, mais le choix relève de la même discipline qu'au **[§ 10.2](02-byref-byval-audit.md)**.
- **`Optional` et `ParamArray` ne se combinent pas.** Une signature ne peut pas avoir à la fois un paramètre `Optional` et un `ParamArray` (vrai en VB6 **comme** en VB.NET) → voir **[§ 10.4](04-paramarray.md)**.
- **Règle de position inchangée.** Tout paramètre situé **après** un `Optional` doit lui aussi être `Optional`, dans les deux langages.

---

## 6. Migration : outils et travail manuel

- Les **assistants** (l'**[*Upgrade Wizard*](../04-outils-migration/01-upgrade-wizard.md)**, les outils commerciaux) **ajoutent une valeur par défaut** là où elle manque et **signalent** les usages de `IsMissing`. Ils rendent le code **compilable**.
- Mais le **choix sémantique** — « ce paramètre doit-il distinguer omis vs fourni ? », « quelle stratégie de remplacement ? » — reste **humain**. Un défaut ajouté machinalement peut compiler tout en ayant **effacé** une distinction dont dépendait la logique.
- **Recommandation par défaut** : pour un type valeur dont l'absence a un sens, privilégier **`Nullable(Of T)`** ; pour deux comportements distincts, **la surcharge** ; n'utiliser la **sentinelle** que si une valeur réellement impossible existe.

---

## 📋 Tableau de synthèse

| Aspect | VB6 | VB.NET (4.7.2) | Nature du changement |
|---|---|---|---|
| `Optional` sans défaut | Autorisé | ❌ Erreur de compilation | Bruyant |
| Forme du défaut | Littéral | **Constante de compilation** | Bruyant (`Now`/`New` interdits) |
| `IsMissing` | Disponible (`Variant`) | **Supprimé** | ⚠️ Sémantique silencieuse |
| Distinguer omis / fourni | Oui (via `IsMissing`) | À **reconstruire** (Nullable / surcharge / sentinelle) | ⚠️ |
| Position des `Optional` | Tous en fin | Identique | ✅ Stable |
| Combinaison avec `ParamArray` | Interdite | Interdite | ✅ Stable |

---

## ✅ À retenir

1. **Valeur par défaut obligatoire** en VB.NET : le `Optional … As Variant` nu de VB6 ne compile plus.
2. **Le défaut doit être une constante** : `Now`, `New …()`, un appel de fonction sont interdits → sentinelle + calcul dans le corps.
3. **`IsMissing` disparaît**, et avec un défaut nu on **ne distingue plus** « omis » de « fourni égal au défaut » — c'est le piège *silencieux* de cette section.
4. **Reconstruire la distinction quand elle compte** : `Nullable(Of T)` (recommandé pour les valeurs), **surcharge** (la plus explicite), ou **sentinelle** (à manier avec précaution).
5. **Les outils rendent le code compilable ; ils ne décident pas du sens.** Le choix de la stratégie reste un jugement à valider.

---

*Section suivante : **[§ 10.4 — `ParamArray` (passé ByRef en VB6 → tableau ByVal en .NET)](04-paramarray.md)** ⚠️*

**Cible** : .NET Framework 4.7.2 · **Indicateurs** : ⭐ point clé · ⚠️ piège · 🛠️ outils

⏭️ [`ParamArray` (passé ByRef en VB6 → tableau ByVal en .NET)](/10-procedures-fonctions/04-paramarray.md)
