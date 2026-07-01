🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.5 — Refactoring idiomatique : LINQ, génériques, espace `My`, code VB.NET moderne

**La récompense : une fois la correction prouvée, passer d'une traduction VB6 littérale à du VB.NET expressif et maintenable — par petits pas, sous la protection du filet de sécurité.**

> 🎁 L'étape **finale** du module, et la seule qui soit du « confort » plutôt que de la nécessité. Mais
> elle est précieuse : un code encore *façonné comme du VB6* est verbeux, plus risqué, et difficile à
> faire évoluer — y compris vers le .NET moderne (module 20). On la mène **après** la validation
> (17.3) et la chasse aux pièges (17.4), jamais avant.

---

## La règle : refactoriser **après**, et **avec le filet**

Trois garde-fous, dans le prolongement du **principe cardinal** du module :

- **Refactoring ≠ réécriture ≠ ajout de fonctionnalités.** Refactoriser, c'est améliorer la **structure
  interne** en **préservant le comportement**. Le périmètre reste **gelé** (module 5.6) ; le *golden
  master* (17.3) doit rester **vert** d'un bout à l'autre.
- **Jamais sans le filet.** Chaque modification re-validée contre le *golden master* : toute « belle »
  réécriture peut introduire une régression.
- **Par petits pas.** Un refactoring à la fois, des *commits* courts (Git, module 6.5), pour qu'une
  régression soit triviale à localiser. Le refactoring *big-bang* est aussi risqué que la migration
  *big-bang*.

> 🛡️ **`Option Strict On` (module 17.1) est désormais votre rail.** Le compilateur vérifie les types à
> chaque transformation : un refactoring qui casse un contrat se voit immédiatement.

---

## Pourquoi refactoriser (et jusqu'où)

Une traduction mécanique « fonctionne », mais traîne des idiomes VB6 verbeux et peu familiers à un  
développeur .NET. Refactoriser améliore la **lisibilité**, la **sûreté** (certains idiomes éliminent des  
classes entières de bugs), parfois la **performance**, et **prépare l'avenir** (module 20).

> ⚖️ **Mais ne sur-polissez pas.** Refactoriser *partout*, par pur souci de pureté, coûte du temps et
> ajoute du risque. Concentrez l'effort sur le code **à fort trafic**, **à forte valeur** ou **difficile
> à comprendre** ; laissez tranquille le code stable et rarement touché. *« Si ça marche et qu'on n'y
> touchera pas, mieux vaut souvent le laisser. »*

---

## LINQ : remplacer les boucles manuelles

Une grande partie du code porté manipule les collections par des **boucles manuelles** (filtrer,  
transformer, agréger). LINQ exprime la même chose de façon **déclarative** :

```vbnet
' ❌ Traduction littérale : boucle manuelle (et risque d'erreur de bornes — B.4)
Dim total As Decimal = 0D
For i As Integer = 0 To commandes.Count - 1
    If commandes(i).Statut = "Validée" Then
        total += commandes(i).Montant
    End If
Next

' ✅ Idiomatique : LINQ, intention déclarative
Dim total As Decimal = commandes.
    Where(Function(c) c.Statut = "Validée").
    Sum(Function(c) c.Montant)
```

Moins de code, **moins d'erreurs d'indices** (recoupe le piège B.4), intention claire. (VB.NET offre  
aussi la **syntaxe de requête** : `Aggregate c In commandes Where … Into Sum(…)`.)

> ⚠️ **Trois précautions** : (1) **exécution différée** — une requête s'exécute à l'**énumération**, pas
> à sa définition (attention si la source change entre-temps, ou pour des effets de bord) ; (2)
> **performance** — dans une boucle chaude, une boucle serrée peut battre LINQ (à mesurer, module 17.6) ;
> (3) **données** — n'aspirez pas une table entière en mémoire pour la « LINQ-er » ; poussez le filtrage
> vers SQL (module 15).

---

## Génériques : remplacer le non-typé et les transtypages

Le code porté regorge souvent de collections **non typées** (`Collection` VB6, parfois `ArrayList`,
`Hashtable`) et donc de **transtypages** partout. Les **génériques** (introduits au module 8.4) y
substituent des types **forts** :

```vbnet
' ❌ Collection non typée (héritée de VB6) → transtypage, index 1-based, pas de vérification
Dim clients As New Collection
clients.Add(unClient)
Dim premier As Client = CType(clients(1), Client)

' ✅ Générique fortement typé → pas de cast, vérifié à la compilation, 0-based
Dim clients As New List(Of Client)
clients.Add(unClient)
Dim premier As Client = clients(0)
```

> 🔗 **Ce refactoring *achève* le travail du module 17.1** : remplacer `Object`/`Collection` par
> `List(Of T)` / `Dictionary(Of K, V)` **élimine** d'un coup les transtypages **et** le liage tardif que
> `Option Strict` signalait. Bénéfice : sûreté de type, clarté, et performance (pas de *boxing* pour les
> types valeur).

---

## L'espace `My` : la concision pour le courant

Adoptez l'espace **`My`** (déjà rencontré au module 16.6) là où il **simplifie** réellement : lecture de  
fichiers (`My.Computer.FileSystem`), préférences (`My.Settings`), ressources (`My.Resources`),  
informations d'application (`My.Application.Info`)…

```vbnet
' Concision idiomatique VB
Dim texte As String = My.Computer.FileSystem.ReadAllText(chemin)
```

> ⚖️ **Équilibre** : `My` est commode mais c'est une abstraction **spécifique à VB**. Pour du code
> destiné à un style plus portable, ou nécessitant un contrôle fin, les API .NET sous-jacentes
> (`System.IO`, `ConfigurationManager`) peuvent être préférables. À noter : `My` n'est que
> **partiellement** disponible sur .NET moderne (module 20.3) — une dépendance lourde est donc une
> légère considération si un second saut est envisagé. À utiliser **à bon escient**, ni dogmatiquement
> ni en l'évitant par principe.

---

## Code VB.NET moderne : la panoplie

Au-delà des trois axes nommés, plusieurs transformations **achèvent** des migrations entamées plus tôt  
ou rendent le code nettement plus clair :

- **`Try`/`Catch`/`Finally`** — finir la migration de `On Error` vers les exceptions structurées
  (module 11) ; retirer le résiduel `On Error`.
- **`Using`** — généraliser la libération déterministe des objets `IDisposable` (module 12.3) ; recoupe
  directement le module 17.6 (fuites).
- **Propriétés auto-implémentées** — `Public Property Nom As String` au lieu du *boilerplate* `Get`/`Set`
  ou des champs publics.
- **Initialiseurs d'objet et de collection** — `New Client With {.Nom = "Durand"}` ;
  `New List(Of Integer) From {1, 2, 3}`.
- **Interpolation de chaînes** — `$"Client : {nom}"` au lieu de la concaténation par `&`.
- **Types valeur nullables** — `Integer?` / `Nullable(Of T)` pour le « peut-être absent » (remplace les
  états `Empty`/`Null` du `Variant`, recoupe B.5).
- **Opérateur conditionnel nul `?.`** — `client?.Nom` renvoie `Nothing` au lieu de planter.
- **`NameOf(...)`** — des noms de symboles **sûrs au renommage**.
- **`For Each`** plutôt que les boucles indexées quand l'indice est inutile (évite les erreurs de bornes,
  B.4) ; **énumérations** au lieu de constantes « nombres magiques ».

### Un cas à manier avec conscience : `IIf` → `If()`

Le ternaire `If(...)` remplace `IIf(...)` — mais ce n'est **pas** qu'un changement de style : `IIf`
évalue **les deux** arguments, `If()` **court-circuite** (modules 9.3) :

```vbnet
' IIf évalue les DEUX branches → n \ d est calculé même si d = 0 → erreur
x = IIf(d <> 0, n \ d, 0)

' If() n'évalue que la branche retenue → sûr
x = If(d <> 0, n \ d, 0)
```

> ⚠️ C'est à la fois plus **idiomatique** et plus **sûr** — mais c'est un **changement de comportement**
> (l'un plantait, l'autre non). Comme tout refactoring : **validez-le au *golden master*** (et il se peut
> qu'il *révèle* ou *corrige* un cas limite — à classer sciemment, cf. 17.3).

---

## Comment refactoriser sûrement (méthode)

La boucle, répétée par petites touches :

1. **Choisir une transformation** ciblée et **une seule** (ne mélangez pas deux refactorings, ni un
   refactoring avec un changement de comportement).
2. **Appliquer**, en s'appuyant sur `Option Strict On` (le rail) et les **outils de refactoring** de
   l'IDE (renommage, extraction de méthode).
3. **Revalider** au *golden master* (17.3) ; **committer** petit (6.5).
4. **Prioriser** : le code à fort trafic / forte valeur / difficile à comprendre d'abord ; le stable et
   figé en dernier (ou pas du tout).

---

## ⚠️ Précautions

- **Ne pas confondre refactoring et changement de comportement** : le périmètre reste gelé (5.6).
- **Jamais sans filet** : chaque lot revalidé au *golden master*.
- **Réécritures qui changent la sémantique** (`IIf`→`If()`, arrondis, ordre d'énumération…) : les
  traiter comme des changements **conscients**, validés. ⚠️
- **LINQ** : exécution différée, coût en boucle chaude, ne pas matérialiser de gros volumes (pousser vers
  SQL).
- **`My` et avenir** : dépendance spécifique à VB, partielle sur .NET moderne (module 20.3).
- **Ne pas sur-polir** : refactoriser là où ça paie, pas partout.
- **Performance** : « idiomatique » n'est pas toujours « plus rapide » — mesurer (17.6).

---

## La récompense, au-delà du présent

Ce refactoring n'embellit pas seulement le code d'aujourd'hui : un VB.NET **idiomatique et moderne** sur
4.7.2 est le **meilleur point de départ possible** pour le **second saut** optionnel vers .NET 10
(module 20). Un code resté façonné comme du VB6 (`Collection`, `On Error`, motifs `Variant`,
`Compatibility`) est bien plus dur à déplacer. Même si vous restez sur 4.7.2, vous **préparez le
terrain**.

---

## En résumé

- Le refactoring idiomatique est la **récompense** : on y vient **après** la validation (17.3) et la
  chasse aux pièges (17.4), **par petits pas**, chaque lot **revalidé** au *golden master*, sous le rail
  d'`Option Strict On`.
- **Refactoring ≠ réécriture ≠ fonctionnalités** : on **préserve le comportement** (périmètre gelé, 5.6).
- **Quatre axes** : **LINQ** (boucles manuelles → requêtes déclaratives), **génériques** (`Collection`/
  casts → `List(Of T)`, ce qui **achève** le travail de 17.1), **`My`** (concision, à bon escient), et le
  **code moderne** (`Try`/`Catch`, `Using`, propriétés auto, initialiseurs, interpolation, nullables,
  `?.`, `NameOf`, `For Each`, `If()`).
- Certaines réécritures « idiomatiques » **changent la sémantique** (`IIf`→`If()`) : à traiter
  consciemment et à **valider**.
- **Ne pas sur-polir** ; concentrer l'effort là où il paie. Bénéfice durable : préparer le **second
  saut** (module 20).

> ➡️ **Dernière étape du module : la revue de performance et de mémoire — `IDisposable`, `Using`, et la
> traque des fuites de *handles*, particulièrement sensibles après de l'interop.** (Module 17.6)

---

🏷️ **Indicateurs** : modernisation (« la récompense ») · achève les migrations des modules 8.4 (collections), 11 (erreurs), 12.3 (Using), 9.3 (If) ; sous filet du module 17.3 ; prépare le module 20
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Revue de performance et de mémoire (`IDisposable`, `Using`, fuites de *handles*)](/17-valider-refactoriser/06-performance-memoire.md)
