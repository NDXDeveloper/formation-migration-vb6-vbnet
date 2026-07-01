🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.2 — `ReDim`, `ReDim Preserve`, tableaux dynamiques

> **Chapitre 8 — Tableaux et collections** · Partie 3 — Migrer le langage
>
> `ReDim` et `ReDim Preserve` **existent toujours** en VB.NET : le code compile. Mais la **mécanique
> change**. En .NET, un tableau est un **objet à taille fixe** : on ne l'agrandit pas sur place —
> `ReDim` **alloue un nouveau tableau**. Conséquences : une **sémantique de référence** qui surprend,
> et surtout un **coût caché** qui pousse, à terme, vers `List(Of T)`.

---

## 🧭 VB6 : les tableaux dynamiques

```vb
' VB6 — tableau dynamique
Dim a() As Integer       ' déclaré non dimensionné (vide)
ReDim a(9)               ' dimensionne : 0..9 (Option Base 0)
ReDim a(19)              ' redimensionne — contenu PERDU (remis aux valeurs par défaut)
ReDim Preserve a(29)     ' redimensionne en CONSERVANT le contenu existant
Erase a                  ' libère le tableau

' Preserve : seule la DERNIÈRE dimension d'un tableau multi-D peut changer
Dim m() As Integer
ReDim m(2, 4)
ReDim Preserve m(2, 9)   ' ✅ on agrandit la dernière dimension
' ReDim Preserve m(5, 9) ' ❌ erreur : la 1re dimension ne peut pas changer
```

- Un tableau **dynamique** se déclare avec des **parenthèses vides** (`Dim a()`) puis se dimensionne
  par `ReDim`.
- **`ReDim`** sans `Preserve` **réinitialise** le contenu ; **`ReDim Preserve`** le **conserve**.
- **`Erase`** libère un tableau dynamique (et réinitialise un tableau fixe).
- **`Preserve`** ne peut redimensionner que la **dernière dimension** d'un tableau multidimensionnel.

---

## 🎯 VB.NET : mêmes mots-clés, mécanique différente

Parce qu'un tableau .NET a une **taille immuable**, `ReDim` ne redimensionne pas : il **réaffecte un  
nouvel objet tableau**.

```vb
' VB.NET — les tableaux sont des objets à taille fixe
Dim a() As Integer
ReDim a(9)               ' alloue un NOUVEAU tableau (0..9)   ≡  a = New Integer(9) {}
ReDim a(19)              ' alloue ENCORE un nouveau tableau — contenu perdu
ReDim Preserve a(29)     ' nouveau tableau + COPIE de l'ancien  ≈  Array.Resize(a, 30)
Erase a                  ' a = Nothing  (et non « remis à 0 »)
```

- **0-based** comme partout (cf. 8.1) : `ReDim a(20)` → 21 éléments ; `ReDim a(1 To 20)` → **erreur**.
- **`ReDim Preserve a(n)`** prend une **borne haute** `n`, alors que **`Array.Resize(a, longueur)`**
  prend un **nombre d'éléments** : `ReDim Preserve a(n)` ≈ `Array.Resize(a, n + 1)`. Ne pas
  confondre les deux conventions.
- VB.NET est plus permissif : on peut `ReDim` n'importe quelle variable tableau (toutes sont des
  références), pas seulement celles déclarées « dynamiques ».

---

## ⚠️ Piège : sémantique de référence (alias)

Comme `ReDim` crée un **nouvel objet**, une autre variable qui désignait l'**ancien** tableau **ne  
suit pas**. Et l'**affectation** elle-même ne se comporte plus pareil :

```vb
' VB6 : affecter un tableau le COPIE
Dim a(2) As Integer, b() As Integer
b = a                    ' b est une COPIE indépendante
b(0) = 99                ' a(0) reste inchangé

' VB.NET : affecter copie la RÉFÉRENCE (alias)
Dim a(2) As Integer, b() As Integer
b = a                    ' a et b désignent le MÊME tableau
b(0) = 99                ' a(0) vaut aussi 99 !
ReDim a(5)               ' a pointe désormais vers un NOUVEAU tableau ; b garde l'ancien → divergence
```

> 🪤 En VB6, `b = a` **dupliquait** le tableau ; en VB.NET, `b = a` crée un **alias** (deux noms, un
> seul objet). Pour obtenir une vraie copie en .NET, utiliser `a.Clone()`, `Array.Copy` ou
> `liste.ToArray()`. C'est un changement de comportement **silencieux** sur du code qui s'appuyait
> sur la copie implicite.

---

## ⚠️ Piège : le coût caché (`ReDim Preserve` en boucle = O(n²))

Chaque `ReDim Preserve` **alloue un nouveau tableau et recopie tout l'existant**. Le redimensionner à
**chaque ajout** — l'idiome VB6 du « tableau qui grandit » — rend la construction **quadratique**
(≈ n²/2 copies pour n éléments).

```vb
' Anti-pattern : ReDim Preserve à chaque tour → O(n²)
Dim a() As Integer
For i = 0 To n - 1
    ReDim Preserve a(i)   ' recopie tout le tableau à CHAQUE itération
    a(i) = Calcul(i)
Next
```

La réponse idiomatique .NET est **`List(Of T)`** : ajout **amorti O(1)** (capacité interne doublée  
automatiquement).

```vb
' Idiomatique .NET : List(Of T)
Dim liste As New List(Of Integer)
For i = 0 To n - 1
    liste.Add(Calcul(i))
Next
Dim a() As Integer = liste.ToArray()   ' si un tableau est requis au final
```

> 🎯 **`ReDim Preserve` dans une boucle est un signal de refactoring** vers `List(Of T)` (ou un autre
> générique — voir **8.4**). On peut migrer d'abord à l'identique (le `ReDim Preserve` compile),
> puis remplacer par une collection lors de la fiabilisation (**chapitres 17.5 et 17.6**).

---

## 🧩 `Preserve` et le multidimensionnel (contrainte inchangée)

La règle VB6 demeure : `ReDim Preserve` ne redimensionne que la **dernière dimension**.

```vb
Dim m(,) As Integer
ReDim m(2, 4)
ReDim Preserve m(2, 9)   ' ✅ dernière dimension
' ReDim Preserve m(5, 9) ' ❌ erreur d'exécution
```

> 💡 `Array.Resize` ne s'applique **qu'aux tableaux à une dimension**. Pour le cas 1-D et la dernière
> dimension, `ReDim Preserve` reste l'outil ; au-delà, il faut recopier manuellement (`Array.Copy`)
> ou repenser la structure (souvent une collection, ou un tableau de tableaux).

---

## 🧱 `Erase` : sémantique modifiée

| | VB6 | VB.NET |
|---|---|---|
| `Erase` sur tableau **dynamique** | libère la mémoire (tableau non initialisé) | `a = Nothing` |
| `Erase` sur tableau **fixe** | **réinitialise** les éléments | `a = Nothing` *(tous les tableaux sont des références)* |

Pour **vider le contenu sans réallouer** (l'ancien `Erase` d'un tableau fixe), utiliser
`Array.Clear` :

```vb
Array.Clear(a, 0, a.Length)   ' remet tous les éléments à leur valeur par défaut
```

---

## 🔁 Boîte à outils idiomatique

| Besoin | VB6 | VB.NET |
|--------|-----|--------|
| Dimensionner | `ReDim a(n)` | `ReDim a(n)` ou `a = New T(n) {}` |
| Agrandir en conservant | `ReDim Preserve a(n)` | `ReDim Preserve a(n)` ou `Array.Resize(a, n + 1)` |
| **Croissance fréquente** | `ReDim Preserve` en boucle | **`List(Of T)`** (`.Add`) |
| Vider le contenu | `Erase` (tableau fixe) | `Array.Clear(a, 0, a.Length)` |
| Libérer | `Erase` | `Erase` (→ `Nothing`) |
| Copier réellement | `b = a` (copie) | `a.Clone()` / `Array.Copy` |

---

## ⚠️ Pièges silencieux — récapitulatif

- **`ReDim` alloue un nouveau tableau** : les autres références gardent l'**ancien** objet.
- **Affectation** : `b = a` **copiait** en VB6, **aliase** en VB.NET → modifications partagées.
- **`ReDim Preserve` en boucle = O(n²)** → préférer `List(Of T)`.
- **`Array.Resize(a, longueur)` ≠ `ReDim Preserve a(borne)`** : longueur vs borne haute (décalage de 1).
- **`Erase`** met le tableau à **`Nothing`** ; pour vider sans réallouer, `Array.Clear`.
- **0-based** s'applique à `ReDim` (cf. 8.1) ; bornes non nulles **interdites**.

---

## 🧭 Stratégie de migration

1. **Conserver `ReDim`/`ReDim Preserve`** pour une première migration fidèle (ils compilent).
2. **Repérer les `ReDim Preserve` en boucle** et les **remplacer par `List(Of T)`** lors de la
   fiabilisation (gain de performance et de lisibilité).
3. **Auditer les affectations de tableaux** (`b = a`) qui supposaient une **copie** → ajouter
   `Clone`/`Array.Copy` si l'indépendance est requise.
4. **Remplacer `Erase` … = réinitialisation** par `Array.Clear` quand on veut vider **sans**
   réallouer.
5. **Filet de sécurité** : le ***golden master*** (chapitre 5.5) révèle une copie devenue alias, ou
   un contenu perdu après un `ReDim` mal placé.

---

## ✅ À retenir

- En .NET, **`ReDim` réalloue** (taille de tableau immuable) ; `ReDim Preserve` ≈ `Array.Resize`
  (nouveau tableau + copie).
- **Affecter un tableau l'aliase** désormais (il le **copiait** en VB6) — utiliser `Clone`/`Copy`
  pour une vraie copie.
- **`ReDim Preserve` répété est quadratique** : le bon outil de croissance est **`List(Of T)`**.
- **`Erase` → `Nothing`** ; pour vider le contenu, **`Array.Clear`**.
- Tout reste **0-based** (cf. 8.1).

---

## 🔗 Pour aller plus loin

- **8.1 — Tableaux 0-based** : bornes et indexation qui s'appliquent aussi à `ReDim`.
- **8.4 — `Collection` → collections .NET** : `List(Of T)`, `Dictionary(Of K,V)` et leurs usages.
- **7.7 — `Nothing`** : tableaux non initialisés et accès protégés.
- **Chapitres 17.5 et 17.6** — refactoring idiomatique et revue de performance/mémoire.
- **Annexe B.4** (Tableaux 0-based et `Option Base`) ; **Annexe D** (correspondance des types).

➡️ Section suivante : **[8.3 — Tableaux de `Variant` / d'objets](03-tableaux-variant.md)**  
⬆️ Retour au **[sommaire du chapitre 8](README.md)**

⏭️ [Tableaux de `Variant` / d'objets](/08-tableaux-collections/03-tableaux-variant.md)
