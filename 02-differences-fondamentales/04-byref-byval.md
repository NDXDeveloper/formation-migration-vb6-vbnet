🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2.4 ByRef par défaut → ByVal par défaut ⭐ ⚠️

> **Chapitre 2 — Différences fondamentales VB6 ↔ VB.NET**
> Un piège qui ne tient qu'à **un mot non écrit** : la valeur par défaut d'un paramètre, qui s'inverse entre les deux langages.

---

## Le piège qui retourne une valeur par défaut

Voici l'archétype du **changement silencieux** : il ne porte pas sur un mot-clé que l'on remplace,  
mais sur un comportement **par défaut** — donc sur ce que vous **n'écrivez pas**. En VB6, un  
paramètre dont on ne précise rien est passé **par référence** (`ByRef`). En VB.NET, le même paramètre,
écrit **exactement pareil**, est passé **par valeur** (`ByVal`). La valeur par défaut **s'est
inversée**.

Conséquence : une procédure peut **cesser de modifier** la variable de l'appelant — sans qu'aucune  
ligne ne change, sans aucune erreur de compilation. C'est l'un des pièges qui causent le plus de  
régressions, et il rejoint celui des entiers redimensionnés (2.3) dans la catégorie « même  
apparence, comportement différent ».

> ⚠️ **La grille du chapitre, appliquée ici.** `ByRef` et `ByVal` existent **dans les deux
> langages**, avec **le même sens**. Ce qui change, c'est **celui qui s'applique quand on ne précise
> rien**. Le danger n'est donc pas dans le mot-clé — il est dans son **absence**.

---

## 🔁 Rappel : ByRef vs ByVal

Avant le piège, le mécanisme — identique des deux côtés :

- **`ByRef`** (par référence) : la procédure reçoit un **accès à la variable de l'appelant**. Si elle modifie le paramètre, **la variable d'origine change**.
- **`ByVal`** (par valeur) : la procédure reçoit une **copie**. Elle peut la modifier librement, mais **la variable d'origine reste intacte**.

Toute la section tient dans une question : *lequel des deux s'applique par défaut ?* — et la réponse
**n'est pas la même** en VB6 et en VB.NET.

---

## 🔵 Côté VB6 : ByRef par défaut

En VB6, **l'absence de mot-clé signifie `ByRef`**. Les trois écritures suivantes sont **strictement
équivalentes** :

```vb
' VB6 — ces trois déclarations sont IDENTIQUES
Sub Exemple(x As Integer)
Sub Exemple(ByRef x As Integer)
' (ByRef est le défaut implicite)
```

Cette valeur par défaut a façonné une **pratique omniprésente** en VB6 : « retourner » des résultats
**par les paramètres**. Comme tout est `ByRef` par défaut, il était naturel et fréquent d'écrire des
procédures qui **renseignent** les variables de l'appelant :

```vb
' VB6 — pattern TRÈS courant : "retour" via paramètre (ByRef implicite)
Sub CalculerTVA(montant As Double, tva As Double, totalTTC As Double)
    tva = montant * 0.2          ' modifie la variable de l'appelant…
    totalTTC = montant + tva     ' …et celle-ci aussi
End Sub

Dim t As Double, ttc As Double
CalculerTVA 100, t, ttc
' t vaut 20, ttc vaut 120 — les modifications sont remontées
```

Ce style est **partout** dans le code VB6 : sous-programmes qui remplissent plusieurs « sorties »,  
fonctions qui renvoient un code de retour **et** des valeurs par paramètres, etc. C'est précisément  
ce qui rend le piège **massif**.

---

## 🟢 Côté VB.NET : ByVal par défaut

En VB.NET, la valeur par défaut **s'inverse** : **l'absence de mot-clé signifie `ByVal`**. L'éditeur  
de Visual Studio insère d'ailleurs souvent `ByVal` **explicitement** quand vous tapez un paramètre —  
matérialisant le nouveau défaut.

```vb
' VB.NET — ces deux déclarations sont IDENTIQUES
Sub Exemple(x As Integer)
Sub Exemple(ByVal x As Integer)
' (ByVal est désormais le défaut implicite)
```

La même déclaration `Sub Exemple(x As Integer)` qui signifiait **`ByRef`** en VB6 signifie donc
**`ByVal`** en VB.NET. **Le mot non écrit a changé de sens.**

---

## ⚠️ Pourquoi c'est silencieux — la démonstration

Reprenons le sous-programme `CalculerTVA` et migrons-le **mécaniquement**, en gardant la déclaration  
telle quelle :

```vb
' VB.NET — MÊME déclaration que VB6, mais ByVal par défaut → silencieusement cassé
Sub CalculerTVA(montant As Double, tva As Double, totalTTC As Double)   ' tva, totalTTC : ByVal !
    tva = montant * 0.2          ' modifie une COPIE locale
    totalTTC = montant + tva     ' modifie une COPIE locale
End Sub

Dim t As Double, ttc As Double
CalculerTVA(100, t, ttc)
' t vaut 0, ttc vaut 0 — les modifications sont PERDUES
```

Faites le constat, point par point — c'est la signature exacte du changement silencieux :

- le code **compile** sans le moindre avertissement ;
- il **s'exécute** sans erreur ;
- il rend simplement un **résultat faux** : les variables de l'appelant **ne sont plus mises à
  jour**. Là où VB6 renvoyait `20` et `120`, VB.NET laisse `0` et `0`.

Aucun garde-fou ne se déclenche. Le bug ne se révèle que **par le comportement** — souvent loin du  
point de conversion, sous la forme de valeurs « mystérieusement » restées à zéro ou inchangées.

---

## 🧩 La subtilité essentielle : types valeur vs types référence

Voici le point qui **explique** pourquoi certains cas semblent « marcher quand même » alors que  
d'autres cassent — et qu'il faut absolument maîtriser pour ne pas se tromper de diagnostic.

**`ByVal` ne veut pas dire « rien ne peut changer ».** Tout dépend de la **nature du type** (rappel
du CTS, 2.1) :

- **Type valeur** (`Integer`, `Double`, `Boolean`, `Structure`…) passé `ByVal` : la procédure reçoit une **copie complète** de la donnée. **Aucune** modification ne remonte. *(C'est le cas de `CalculerTVA` ci-dessus → cassé.)*
- **Type référence** (un objet, une `List`, un formulaire…) passé `ByVal` : la procédure reçoit une **copie de la référence**, qui **pointe vers le même objet**. Conséquence en deux temps :
  - **muter les membres** de l'objet (modifier ses propriétés, ajouter à la liste…) **remonte** bien à l'appelant — les deux références désignent le même objet ;
  - mais **réaffecter** la variable à un **nouvel** objet (`o = New ...`) **ne remonte pas** — on n'a réaffecté qu'une **copie** de la référence.

```vb
' Type référence passé ByVal (défaut VB.NET)
Sub Modifier(o As List(Of String))     ' ByVal
    o.Add("ajouté")                     ' ✅ mute l'objet partagé → VISIBLE par l'appelant
    o = New List(Of String)             ' ❌ réaffecte une COPIE de la référence → INVISIBLE
End Sub
```

> 💡 **Pourquoi ce point est crucial en migration.** Beaucoup de procédures VB6 reçoivent des objets
> et se contentent d'en **modifier les membres**. Pour celles-là, le passage à `ByVal` est
> **inoffensif** (la mutation remonte quand même). Le piège se concentre sur **deux cas précis** :
> les **types valeur** modifiés (le pattern « sortie » classique), et la **réaffectation** d'une
> référence (le `Set obj = New ...` de VB6 dans une procédure). Savoir distinguer ces cas évite à la
> fois de **rater** un vrai bug et de **sur-corriger** un cas sain.

---

## 🗺️ Le tableau des quatre combinaisons

À garder en tête pour trancher chaque paramètre :

| Passage | **Type valeur** (`Integer`, `Structure`…) | **Type référence** (objet, `List`…) |
|---|---|---|
| **`ByVal`** | Copie totale → **aucune** remontée | Mutation des membres : **remonte** · Réaffectation : **ne remonte pas** |
| **`ByRef`** | **Remonte** (la variable d'origine change) | Mutation **et** réaffectation **remontent** |

La règle se lit ainsi : pour **préserver le comportement VB6**, il faut `ByRef` partout où la  
procédure d'origine **comptait** sur la remontée — c'est-à-dire pour **tout type valeur modifié** et  
pour **toute réaffectation de référence**.

---

## 🛠️ Ce que font les outils — et le piège du « nettoyage »

Comme pour les entiers redimensionnés (2.3), les bons outils de migration adoptent le choix
**sûr** : ils **préservent la sémantique** en insérant `ByRef` **explicitement** partout où VB6
appliquait le `ByRef` implicite.

```vb
' Conversion outillée FIDÈLE : le ByRef implicite de VB6 est rendu explicite
Sub CalculerTVA(ByVal montant As Double, ByRef tva As Double, ByRef totalTTC As Double)
    ' comportement identique à VB6 ✔
```

C'est **correct** — mais cela produit un code **truffé de `ByRef`**, y compris sur des paramètres qui  
n'étaient **jamais réellement modifiés**. D'où une tentation dangereuse :

> ⚠️ **Ne « nettoyez » JAMAIS les `ByRef` à l'aveugle.** Retirer un `ByRef` « qui semble inutile »
> peut **casser silencieusement** un appel, si le paramètre était en réalité utilisé comme sortie.
> Avant de remplacer un `ByRef` par `ByVal`, il faut **vérifier** que la procédure ne modifie pas ce
> paramètre **et** qu'aucun appelant ne dépend de cette modification. C'est un travail d'**audit**,
> pas de cosmétique.

Et surtout, **ce filet n'existe pas toujours** : migration **manuelle**, outil qui ne préserve pas  
la sémantique, ou code recopié à la main. Dans tous ces cas, c'est à vous de rétablir le bon  
mécanisme — paramètre par paramètre.

---

## ↔️ Cas connexes (renvois)

- **`ParamArray`** : en VB6, il est passé **`ByRef`** ; en VB.NET, c'est un **tableau passé `ByVal`**. Un changement de comportement à part entière, traité en **section 10.4**.
- **L'astuce des parenthèses en VB6** : en VB6, entourer un argument de parenthèses **supplémentaires** — `Exemple((x))` — force le passage **`ByVal`** même sur un paramètre `ByRef`. À connaître pour **lire** correctement du code VB6 (un `((x))` n'est pas une coquille) ; détaillé côté appels en **section 10.1**.
- **L'audit systématique des appels** — la méthode concrète pour traiter ce piège à l'échelle d'un projet — est l'objet de la **section 10.2**, et le piège est fiché en **annexe B.1**.

---

## 🧭 La règle de migration à retenir

Face à chaque paramètre **sans mot-clé** dans du code VB6, rappelez-vous qu'il était **`ByRef`**, et  
tranchez :

| Le paramètre VB6 (ByRef implicite)… | Cible .NET |
|---|---|
| …est **modifié** et l'appelant **compte** sur la remontée (type valeur, ou réaffectation d'objet) | **`ByRef` explicite** ✅ (préserve le comportement) |
| …n'est **jamais modifié** (simple lecture) | **`ByVal`** — le nouveau défaut, idiomatique |
| …est un **objet** dont on modifie seulement les **membres** | **`ByVal`** suffit (la mutation remonte quand même) |

> ⭐ **Le principe directeur.** En cas de **doute**, conservez **`ByRef`** : c'est le choix qui
> **préserve le comportement VB6**. Le passage à `ByVal` est une **optimisation/clarification** à ne
> faire **qu'après vérification**, jamais par réflexe. Préserver d'abord, simplifier ensuite.

---

## ✅ À retenir

- La **valeur par défaut s'inverse** : un paramètre sans mot-clé est **`ByRef` en VB6** mais **`ByVal` en VB.NET**. Le danger tient à ce que l'on **n'écrit pas**.
- C'est un piège **silencieux** : la même déclaration **compile des deux côtés**, mais la procédure **cesse de modifier** la variable de l'appelant → **résultats faux**, sans erreur.
- Le pattern VB6 « **retour via paramètre** » (très répandu) est la principale victime.
- **Subtilité clé** : `ByVal` ne bloque pas tout. Type **valeur** → aucune remontée ; type **référence** → la **mutation des membres remonte**, mais **pas la réaffectation**. Distinguer ces cas évite de rater un bug **ou** de sur-corriger.
- Les outils préservent la sémantique en ajoutant `ByRef` **explicite** (code **correct** mais **truffé de `ByRef`**) : **ne pas les retirer sans audit**.
- **Règle d'or** : dans le doute, **garder `ByRef`** (préserve le comportement) ; passer à `ByVal` seulement **après vérification**. Méthode d'audit en **10.2**, cas `ParamArray` en **10.4**, fiche en **annexe B.1**.

---

> 🧭 **Section suivante** → [2.5 Propriétés par défaut et mot-clé `Set` (disparu)](05-default-properties-set.md) ⚠️  
> 🔝 **Retour** → [Sommaire du chapitre 2](README.md)

⏭️ [Propriétés par défaut et mot-clé `Set` (disparu)](/02-differences-fondamentales/05-default-properties-set.md)
