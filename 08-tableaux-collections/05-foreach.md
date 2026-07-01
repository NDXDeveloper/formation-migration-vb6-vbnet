🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.5 — `For Each` et énumération (différences de comportement)

> **Chapitre 8 — Tableaux et collections** · Partie 3 — Migrer le langage
>
> Dernière section du chapitre. `For Each` existe dans les deux langages, mais .NET est **plus
> strict** et **typé**. La différence la plus visible : **modifier une collection pendant qu'on la
> parcourt** était *toléré* (mais bugué) en VB6, et **lève une exception** en .NET. S'y ajoutent le
> **type de la variable de boucle**, l'**ordre** et le piège de la **copie**.

---

## 🧭 VB6 : `For Each`

```vb
' VB6 — variable Variant/Object, parcours d'un tableau ou d'une Collection
Dim element As Variant
For Each element In maCollection
    Debug.Print element
Next

' Modifier pendant le parcours : "toléré" mais BUGUÉ (éléments sautés)
Dim i As Variant
For Each i In c
    If DoitEtreSupprime(i) Then c.Remove i   ' ⚠️ fait sauter l'élément suivant
Next
```

- La variable de boucle est typiquement un **`Variant`** (ou un `Object`).
- Fonctionne sur les **tableaux**, l'objet **`Collection`** et les **collections COM** (tout objet
  exposant un énumérateur par défaut, `_NewEnum`).
- **Lenient mais peu fiable** : retirer des éléments en cours de parcours ne **plantait** pas — ça
  **sautait** silencieusement des éléments (le bug classique).

---

## 🎯 VB.NET : `For Each` sur `IEnumerable`

```vb
' VB.NET — variable TYPÉE, parcours de génériques
For Each nom As String In noms
    Console.WriteLine(nom)
Next
```

- Parcourt tout ce qui implémente **`IEnumerable`/`IEnumerable(Of T)`** (motif
  `GetEnumerator`/`MoveNext`/`Current`).
- La variable de boucle peut — et devrait, sous **`Option Strict On`** — être **typée**. Sur une
  collection **non générique** (la `Collection` VB, un `SAFEARRAY`…), les éléments sont des `Object`
  à **caster** (cf. 7.1).

---

## 🧨 Le piège phare : modifier la collection pendant le parcours

En .NET, l'énumérateur **détecte toute modification de structure** et **échoue immédiatement** :

```vb
' ❌ Ajouter/retirer pendant For Each → InvalidOperationException
For Each n As String In noms
    If DoitEtreSupprime(n) Then noms.Remove(n)   ' « Collection was modified… »
Next
```

> ⚠️ Là où VB6 **sautait** des éléments en silence, .NET lève **`InvalidOperationException`** au
> premier `MoveNext` suivant la modification. **Plus sûr** (le bug devient visible), mais c'est un
> changement de comportement à traiter.

**Trois corrections**, selon le besoin :

```vb
' (1) Parcourir une COPIE
For Each n As String In noms.ToList()
    If DoitEtreSupprime(n) Then noms.Remove(n)
Next

' (2) Boucle FOR À REBOURS (idéale pour la suppression par indice)
For i As Integer = noms.Count - 1 To 0 Step -1
    If DoitEtreSupprime(noms(i)) Then noms.RemoveAt(i)
Next

' (3) Collecter d'abord, supprimer ENSUITE
Dim aSupprimer = noms.Where(Function(n) DoitEtreSupprime(n)).ToList()
For Each n In aSupprimer : noms.Remove(n) : Next
```

> 💡 La boucle **à rebours** (2) est souvent la plus sûre pour supprimer par indice : retirer
> l'élément `i` ne décale pas les indices **déjà parcourus**.

---

## 🧩 La variable de boucle est une **copie**

Affecter la variable de `For Each` **ne modifie pas** l'élément sous-jacent. C'est vrai dans les deux  
langages, mais le piège est aigu avec les **types valeur** (structures, cf. 12.7) :

```vb
' Point est une Structure (type valeur)
For Each p As Point In listePoints
    p.X = 0                 ' ⚠️ modifie une COPIE, pas l'élément de la liste
Next

' Pour modifier EN PLACE → boucle FOR indexée
For i As Integer = 0 To listePoints.Count - 1
    Dim p = listePoints(i)
    p.X = 0
    listePoints(i) = p
Next
```

---

## 🔑 Énumérer un `Dictionary` (et l'ordre)

Le parcours d'un dictionnaire fournit des **`KeyValuePair(Of K,V)`**, pas seulement des valeurs
(cf. 8.4) :

```vb
For Each kv As KeyValuePair(Of String, String) In annuaire
    Console.WriteLine($"{kv.Key} = {kv.Value}")
Next
For Each cle As String In annuaire.Keys     ' …les clés seules
For Each valeur As String In annuaire.Values ' …les valeurs seules
```

> 🪤 **L'ordre n'est pas garanti** pour un `Dictionary` (alors que la `Collection` VB6 était
> ordonnée). Du code qui s'appuyait sur l'ordre d'insertion via `For Each` **change de comportement
> silencieusement** une fois passé au dictionnaire (voir 8.4 pour les alternatives ordonnées).
> Les **tableaux** et **`List(Of T)`**, eux, s'énumèrent toujours dans l'**ordre des indices**.

---

## 🧱 Cas particuliers

- **`For Each` sur `Nothing`** : lève une **`NullReferenceException`** en .NET. **Garder** :
  `If maCollection IsNot Nothing Then For Each …`.
- **Classe énumérable personnalisée** : VB6 exposait `_NewEnum` (`IEnumVARIANT`) pour rendre une
  classe parcourable par `For Each`. En .NET, on **implémente `IEnumerable(Of T)`** :

  ```vb
  Class Equipe : Implements IEnumerable(Of Joueur)
      ' … Public Function GetEnumerator() As IEnumerator(Of Joueur) …
  End Class
  ```

  (Détails côté POO au **chapitre 12** et interop COM au **module 16**.)

> 💡 **Au-delà de `For Each`.** `IEnumerable` ouvre la porte à **`LINQ`** (`Where`, `Select`,
> `OrderBy`…), qui remplace souvent une boucle entière par une expression lisible — une
> modernisation à envisager lors de la fiabilisation (**chapitre 17.5**).

---

## ⚠️ Pièges silencieux — récapitulatif

- **Modifier pendant `For Each`** : VB6 **sautait** des éléments ; .NET **lève
  `InvalidOperationException`** → parcourir une copie, boucler à rebours, ou différer la suppression.
- **Variable de boucle = copie** : la modifier ne change pas la collection (aigu avec les
  **structures**) → boucle `For` indexée.
- **`Dictionary` non ordonné** : l'ordre de `For Each` n'est **pas** garanti (≠ `Collection` VB6).
- **`For Each` sur `Nothing`** → `NullReferenceException`.
- Collection **non générique** → éléments `Object` à **caster** sous `Option Strict On`.

---

## 🧭 Stratégie de migration

1. **Repérer les modifications en cours de parcours** (`Add`/`Remove` dans un `For Each`) et les
   réécrire (copie, boucle à rebours, suppression différée).
2. **Typer** la variable de boucle ; **caster** les éléments des collections non génériques.
3. **Vérifier les dépendances à l'ordre** lorsqu'une `Collection` devient un `Dictionary`.
4. **Protéger** les `For Each` contre les collections `Nothing`.
5. **Convertir `_NewEnum`** en `IEnumerable(Of T)` pour les classes énumérables.
6. **Filet de sécurité** : le ***golden master*** (chapitre 5.5) révèle un élément sauté, une
   structure modifiée « dans le vide » ou un ordre de parcours changé.

---

## ✅ À retenir

- **`For Each` est plus strict en .NET** : modifier la collection pendant le parcours **lève une
  exception** (au lieu de sauter des éléments).
- La **variable de boucle est une copie** : pour modifier en place (surtout des **structures**),
  utiliser une boucle **`For` indexée**.
- **`Dictionary`** s'énumère en **`KeyValuePair`** et **sans ordre garanti** ; tableaux et `List` en
  **ordre d'indices**.
- **`Nothing`** → `NullReferenceException` ; collections non génériques → **casts**.
- **`IEnumerable`** remplace `_NewEnum` et ouvre la voie à **`LINQ`**.

---

## 🔗 Pour aller plus loin

- **8.1 — Tableaux 0-based** : la boucle `For` indexée et ses bornes.
- **8.4 — `Collection` → collections .NET** : `KeyValuePair`, ordre, et alternatives ordonnées.
- **Chapitre 12.7** — `Type` → `Structure` : sémantique de valeur (le piège de la copie).
- **Chapitre 11 — Gestion des erreurs** : exceptions comme `InvalidOperationException`.
- **Chapitre 17.5 — Refactoring idiomatique** : `LINQ` au-delà de `For Each`.
- **Annexe D** — correspondance des types de données.

---

🏁 **Fin du chapitre 8 — Tableaux et collections.** Les structures de données étant traitées, le
chapitre suivant aborde les **opérateurs, les chaînes et les structures de contrôle** — où d'autres  
faux-amis attendent (`&` vs `+`, la division entière `\`, `IIf` vs `If()`…).

➡️ Chapitre suivant : **[9. Opérateurs, chaînes et structures de contrôle](../09-operateurs-controle/README.md)**  
⬆️ Retour au **[sommaire du chapitre 8](README.md)**

⏭️ [Opérateurs, chaînes et structures de contrôle](/09-operateurs-controle/README.md)
