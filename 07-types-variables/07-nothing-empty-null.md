🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.7 — `Nothing`, `Empty`, `Null`, `Missing` : la fin du `Variant` et de ses états ⚠️

> **Chapitre 7 — Types, variables et déclarations** · Partie 3 — Migrer le langage
>
> Dernière section du chapitre, et celle qui **rassemble les fils** : le `Variant` de VB6 portait
> **quatre états « non-valeur » distincts** — `Empty`, `Null`, `Nothing`, `Missing` — chacun avec
> sa sémantique et son test. .NET **démantèle** cette taxonomie au profit d'un jeu plus restreint
> et **typé** (`Nothing`, `DBNull`, `Nullable(Of T)`). La correspondance n'est pas « 1 pour 1 » :
> c'est tout l'enjeu de cette section.

---

## 🧭 Les quatre états du `Variant` en VB6

```vb
' VB6 — quatre états distincts (+ Nothing pour les objets)
Dim v As Variant          ' v est EMPTY (non initialisé)
Debug.Print IsEmpty(v)    ' True
Debug.Print v + 1         ' 1   (Empty se comporte comme 0, et comme "" en contexte chaîne)

v = Null                  ' NULL : « pas de donnée valide » (ex. champ de base NULL)
Debug.Print IsNull(v)     ' True
Debug.Print v + 5         ' Null  ⚠️ le Null se PROPAGE en arithmétique
Debug.Print v = Null      ' Null  ⚠️ jamais True : '=' ne teste PAS le Null

Dim o As Object
Set o = Nothing           ' NOTHING : pas de référence objet (et libère la réf. COM)
Debug.Print o Is Nothing  ' True
```

| État | Sens | Test VB6 | Particularité |
|------|------|----------|---------------|
| **`Empty`** | `Variant` **non initialisé** | `IsEmpty` | vaut **0** (numérique) ou **`""`** (chaîne) |
| **`Null`** | **pas de donnée valide** (BD) | `IsNull` | se **propage** ; `= Null` ne marche pas |
| **`Nothing`** | **pas de référence** objet | `Is Nothing` | `Set … = Nothing` **libère** la réf. COM |
| **`Missing`** | argument optionnel **omis** | `IsMissing` | seulement pour `Optional … As Variant` |

> 🪤 **Deux pièges VB6 à connaître avant même de migrer :**
> - **`= Null` ne teste rien** : `If v = Null` renvoie `Null` (jamais `True`). Il fallait `IsNull(v)`.
> - **Le `Null` contamine** les calculs : `Null + 5` = `Null`. (Subtilité : même le `&` traite le
>   `Null` comme une chaîne **vide**, alors que `+` et l'arithmétique le **propagent**.)
>
> Et `Missing` : seul un paramètre `Optional As Variant` **sans valeur par défaut** pouvait être
> « manquant » (voir l'idiome `If IsMissing(arg) Then …`).

---

## 🎯 Le jeu d'outils .NET (plus petit, mais typé)

.NET remplace ces quatre états par **trois** notions distinctes — qu'il ne faut **surtout pas
confondre** :

### `Nothing` — la **valeur par défaut** du type (sens élargi)

En VB6, `Nothing` ne concernait que les **objets**. En VB.NET, `Nothing` désigne **la valeur par  
défaut de n'importe quel type** :

```vb
Dim o As Client   = Nothing    ' référence nulle (comme VB6)
Dim i As Integer  = Nothing    ' 0   (et non une erreur !)
Dim b As Boolean  = Nothing    ' False
Dim d As DateTime = Nothing    ' DateTime.MinValue = 0001-01-01 (voir 7.3)
```

> ⚠️ Conséquence : **`Is Nothing` ne s'applique qu'aux types référence** (et nullables). Sur un
> type valeur, `i Is Nothing` est une **erreur de compilation** — il faut comparer à la valeur par
> défaut, ou utiliser un type nullable.

### `DBNull.Value` — le **null des bases de données**

Le `Null` « donnée absente » de VB6 (typiquement un champ d'`ADO`) devient, au **bord de l'accès  
aux données**, le singleton **`System.DBNull.Value`** :

```vb
If IsDBNull(ligne("Nom")) Then ...     ' un champ de base est NULL
```

> 🚨 **Piège fréquent :** un `DataReader`/`DataRow` renvoie **`DBNull.Value`**, **pas** `Nothing`.
> Comparer un champ de base à `Nothing` au lieu de `DBNull.Value` est un bug classique. Et
> **`DBNull` ne se propage pas** comme le `Null` VB6 : `DBNull.Value + 5` **lève une exception** —
> il faut **tester** `IsDBNull` **avant** d'utiliser la valeur.

### `Nullable(Of T)` — l'**absence typée** (et la propagation retrouvée)

Pour un **type valeur** qui peut être « non renseigné », .NET offre `Nullable(Of T)` (`Integer?`,
`DateTime?`…). Bonne surprise : ses opérateurs sont **« levés »**, donc la **propagation du Null
revient** — proprement et de façon typée :

```vb
Dim age As Integer? = Nothing
Dim resultat = age + 5         ' Nothing  ✅ la propagation est restaurée
If age.HasValue Then Traiter(age.Value)
```

### Et pour les chaînes : `Nothing` vs `""`

VB6 distinguait `Empty` (Variant), `""` (chaîne vide) et `vbNullString` (pointeur nul). En .NET,  
une `String` est soit **`Nothing`** (nulle), soit **`""`** (vide) — `vbNullString` devient
**`Nothing`** (voir 7.6). Pour couvrir les deux cas :

```vb
If String.IsNullOrEmpty(s) Then ...
```

---

## 🔁 Tableau de correspondance

| État **VB6** | Test VB6 | Cible **.NET** | Test **.NET** |
|--------------|----------|----------------|---------------|
| `Empty` | `IsEmpty` | **valeur par défaut** (`Nothing` pour les objets ; `0`/`False`/`MinValue`…) | selon le type |
| `Null` (donnée BD) | `IsNull` | **`DBNull.Value`** | `IsDBNull` |
| `Null` (calcul « sans valeur ») | `IsNull` | **`Nullable(Of T)`** (propage) | `.HasValue` |
| `Nothing` (objet) | `Is Nothing` | **`Nothing`** (référence) | `Is Nothing` |
| `Missing` (arg. omis) | `IsMissing` | **défaut obligatoire** (sentinelle) | comparer au défaut (voir **10.3**) |

---

## 🧭 Migration, état par état

### `Empty` → valeur par défaut

L'état `Empty` **disparaît** : une variable non initialisée prend la **valeur par défaut** de son  
type (`Nothing` pour un objet, `0`/`False`/`MinValue` pour un type valeur). Le code qui s'appuyait  
sur `IsEmpty`/le comportement « 0 ou "" » doit **tester explicitement** la valeur par défaut  
attendue.

### `Null` → `DBNull` **ou** `Nullable(Of T)` **ou** `Nothing`

C'est la traduction la plus **contextuelle**. Choisir selon l'**intention** :

- **donnée de base** (champ NULL) → **`DBNull.Value`** + `IsDBNull` ;
- **valeur métier optionnelle** (calculs) → **`Nullable(Of T)`** + `.HasValue` (propagation typée) ;
- **référence absente** → **`Nothing`**.

Et remplacer impérativement les tests `= Null` (qui ne marchaient pas) par le **bon prédicat**.

### `Nothing` → `Nothing` (références), mais un détail capital

Le `Nothing` « objet » se migre directement. Mais en VB6, `Set obj = Nothing` **libérait** la  
référence COM et déclenchait potentiellement `Class_Terminate` (**finalisation déterministe**). En
.NET, affecter `Nothing` **ne libère rien immédiatement** : le *garbage collector* décide **quand**.
Pour les objets détenant des ressources, c'est **`IDisposable`/`Using`** qui prend le relais.

> 🔗 C'est le **piège n°1** de la migration, traité en profondeur aux **chapitres 2.2, 12.2 et
> 12.3**. Affecter `Nothing` n'est **pas** une libération.

### `Missing` → valeur par défaut obligatoire

`IsMissing` **disparaît** : en VB.NET, un paramètre **`Optional` doit avoir une valeur par défaut**.
Un argument omis prend donc son **défaut déclaré** — il n'y a plus d'état « manquant ».

```vb
' VB6
Sub Afficher(Optional Titre As Variant)
    If IsMissing(Titre) Then Titre = "(par défaut)"
End Sub

' VB.NET — défaut OBLIGATOIRE ; plus de IsMissing
Sub Afficher(Optional Titre As String = Nothing)
    If Titre Is Nothing Then Titre = "(par défaut)"
End Sub
```

> 🔗 `Optional`, les valeurs par défaut obligatoires et la fin de `IsMissing` sont détaillés au
> **chapitre 10.3**.

---

## ⚠️ Pièges silencieux — récapitulatif

- **Trois notions à ne pas confondre** : `Nothing` (défaut/réf. nulle), `DBNull.Value` (null de
  base), `Nullable(Of T)` sans valeur. **Ce ne sont pas des synonymes.**
- **Un champ de base est `DBNull.Value`, pas `Nothing`** : le comparer à `Nothing` est un bug.
- **Plus de propagation par défaut** : `DBNull` **lève** en arithmétique — mais **`Nullable(Of T)`
  propage** (opérateurs levés).
- **`Is Nothing`** ne vaut que pour les **références/nullables** (erreur de compilation sur un type
  valeur).
- **`Nothing` ≠ libération** : ne déclenche pas de finalisation déterministe (→ `IDisposable`).
- **Chaînes** : `Nothing` (nulle) **vs** `""` (vide) → `String.IsNullOrEmpty`.

---

## 🧭 Stratégie de migration

1. **Cartographier l'intention** derrière chaque `Empty`/`Null`/`Nothing`/`Missing` avant de
   traduire — c'est l'intention, pas le mot, qui dicte la cible.
2. **Router les `Null`** : `DBNull.Value` au bord des données, `Nullable(Of T)` dans la logique
   métier, `Nothing` pour les références.
3. **Réécrire tous les tests `= Null`/`IsEmpty`/`IsMissing`** en prédicats explicites
   (`IsDBNull`, `.HasValue`, `Is Nothing`, comparaison au défaut).
4. **Ne jamais traiter `Nothing` comme une libération** : pour les ressources, `Using`/`IDisposable`
   (chapitres 12.2–12.3).
5. **Filet de sécurité** : le ***golden master*** (chapitre 5.5) révèle un champ NULL mal testé, une
   propagation perdue, ou un défaut d'argument optionnel qui a changé.

---

## ✅ À retenir

- Les **quatre états** du `Variant` (`Empty`, `Null`, `Nothing`, `Missing`) **disparaissent** au
  profit de **`Nothing`** (défaut), **`DBNull.Value`** (BD) et **`Nullable(Of T)`** (absence typée,
  qui **propage**).
- **`Nothing` est élargi** : valeur par défaut de **tout** type (`0`/`False`/`MinValue`), pas
  seulement « objet nul » → **`Is Nothing` réservé aux références**.
- **Le `Null` de base = `DBNull.Value`** (≠ `Nothing`) et **ne se propage pas**.
- **`= Null`/`IsEmpty`/`IsMissing`** → **prédicats explicites** (`IsDBNull`/`.HasValue`/`Is Nothing`).
- **`Nothing` ne libère rien** : la finalisation déterministe passe par **`IDisposable`/`Using`**.

---

## 🔗 Pour aller plus loin

- **7.1 — `Variant` → `Object`** : le conteneur dont ces états dépendaient.
- **7.3 — `Date`** : la « date vide » = `DateTime.MinValue`, et `DateTime?` pour l'absence.
- **7.6 — Booléen / constantes** : `vbNullString` → `Nothing`.
- **Chapitre 10.3** — `Optional` et fin de `IsMissing`.
- **Chapitres 2.2, 12.2 et 12.3** — finalisation déterministe perdue, `IDisposable`/`Using`.
- **Chapitre 15** — accès aux données et `DBNull` côté `ADO.NET`.
- **Annexe B.5** (`Variant` : `Empty`/`Null`/`Missing` perdus) ; **Annexe D** (correspondance des
  types).

---

🏁 **Fin du chapitre 7 — Types, variables et déclarations.** Le socle des types étant posé, le
chapitre suivant aborde les **tableaux et collections** — à commencer par un autre changement de  
fond : les tableaux **toujours 0-based**.

➡️ Chapitre suivant : **[8. Tableaux et collections](../08-tableaux-collections/README.md)**  
⬆️ Retour au **[sommaire du chapitre 7](README.md)**

⏭️ [Tableaux et collections](/08-tableaux-collections/README.md)
