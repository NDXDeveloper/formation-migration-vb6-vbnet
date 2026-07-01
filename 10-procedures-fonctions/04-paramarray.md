🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.4 — `ParamArray` (passé ByRef en VB6 → tableau ByVal en .NET) ⚠️

> **Le nombre variable d'arguments survit — mais change de *nature*.** En VB.NET, le `ParamArray` devient un **tableau typé**, passé **`ByVal`**. Le rare cas où VB6 permettait de **modifier les variables de l'appelant** à travers lui **casse en silence**.

📍 *Module 10 · § 10.4 · [↑ Introduction du chapitre](README.md) · [← § 10.3](03-optional-ismissing.md)*

---

## ⚡ En bref

- `ParamArray` permet de recevoir un **nombre variable d'arguments**, et reste le **dernier** paramètre — concept conservé en VB.NET.
- **Trois changements** : (1) le **type** des éléments passe de `Variant` (imposé) à un **tableau typé** (`Object()` au plus proche) ; (2) le **mode de passage** bascule de **`ByRef` à `ByVal`** ; (3) un appel **sans argument** se teste différemment (`Length = 0` au lieu de `UBound < LBound`).
- **Le piège ⚠️** est le n°2 : en VB6, on pouvait écrire dans les variables de l'appelant via le `ParamArray` ; en VB.NET, pour des arguments **isolés**, un **nouveau tableau** est créé et ces écritures sont **perdues**. Et l'on **ne peut pas** déclarer `ByRef ParamArray` pour rétablir le comportement.

---

## 1. Rappel : à quoi sert `ParamArray`

`ParamArray` autorise un appel avec un nombre **indéterminé** d'arguments — l'archétype étant une fonction de somme, de concaténation ou de journalisation. Dans les deux langages, c'est le **dernier** paramètre de la signature, et il ne peut y en avoir **qu'un seul**.

```vb
' VB6 — ParamArray : toujours un tableau de Variant
Function SommeTout(ParamArray valeurs()) As Double
    Dim i As Integer, total As Double
    For i = LBound(valeurs) To UBound(valeurs)
        total = total + valeurs(i)
    Next
    SommeTout = total
End Function
```

```vb
' VB.NET — ParamArray : tableau TYPÉ
Function SommeTout(ParamArray valeurs() As Double) As Double
    Dim total As Double = 0
    For Each v As Double In valeurs
        total += v
    Next
    Return total
End Function
```

---

## 2. Trois changements à connaître

| Aspect | VB6 | VB.NET (4.7.2) |
|---|---|---|
| Type des éléments | `Variant` (**imposé**) | **Tableau typé** (`Object()` au plus proche, ou type fort) |
| Mode de passage | **`ByRef`** | **`ByVal`** (non modifiable) |
| Modifier les variables de l'appelant | Possible (arguments isolés) | **Impossible** pour des arguments isolés ⚠️ |
| Aucun argument | `UBound < LBound` (UBound = -1) | `Length = 0` (jamais `Nothing` en appel normal) |
| Bornes du tableau | 0-based | 0-based (`LBound` toujours 0) |
| Position dans la signature | Dernier paramètre | Dernier paramètre |
| Combinaison avec `Optional` | Interdite | Interdite |

---

## 3. Le changement de mode : `ByRef` → `ByVal` (le piège)

C'est le point sensible de cette section. En **VB6**, les éléments d'un `ParamArray` sont passés **`ByRef`** : une procédure peut **écrire dans les variables de l'appelant** transmises comme arguments. Usage rare, mais bien réel.

```vb
' VB6 — ParamArray ByRef : écrit dans les variables de l'appelant
Sub RemettreAZero(ParamArray valeurs())
    Dim i As Integer
    For i = LBound(valeurs) To UBound(valeurs)
        valeurs(i) = 0
    Next
End Sub

Dim a As Long, b As Long
a = 10 : b = 20
RemettreAZero a, b
' a = 0, b = 0   ✅  les variables ont bien été modifiées
```

En **VB.NET**, le `ParamArray` est passé **`ByVal`**. Pour des arguments **isolés**, le compilateur **synthétise un nouveau tableau** à partir des valeurs : écrire dans ce tableau **ne touche pas** les variables d'origine.

```vb
' VB.NET — ParamArray ByVal : un NOUVEAU tableau est créé pour des arguments isolés
Sub RemettreAZero(ParamArray valeurs() As Integer)
    For i As Integer = 0 To valeurs.Length - 1
        valeurs(i) = 0          ' ⚠️ écrit dans le tableau synthétisé, PAS dans a/b
    Next
End Sub

Dim a As Integer = 10, b As Integer = 20
RemettreAZero(a, b)
' a = 10, b = 20   ⚠️  les modifications sont PERDUES
```

Comme pour le **[§ 10.2](02-byref-byval-audit.md)**, le code **compile sans avertissement** ; seul le comportement change. Et surtout : **on ne peut pas écrire `ByRef ParamArray`** en VB.NET (c'est une erreur de compilation). Impossible, donc, de « rétablir » le comportement par un simple modificateur — il faut **redessiner** (voir § 7).

> 🔎 **Nuance utile** : si l'appelant passe un **vrai tableau** (et non des arguments isolés), la procédure reçoit **ce tableau même**, et modifier ses éléments **propage** bien les changements (c'est le même objet) :
>
> ```vb
> Dim valeurs() As Integer = {10, 20}
> RemettreAZero(valeurs)
> ' valeurs(0) = 0, valeurs(1) = 0   ← ici, propagation (même objet tableau)
> ```
>
> Le comportement dépend donc de **la façon d'appeler** : arguments isolés (nouveau tableau, pas de propagation) vs tableau existant (même tableau, propagation des éléments). Cette asymétrie n'existait pas en VB6.

---

## 4. Le changement de type : `Variant()` → tableau typé

En VB6, un `ParamArray` est **forcément** un tableau de `Variant`. En VB.NET, il faut un **type explicite**. Deux options :

- **Transcription fidèle** : `ParamArray args() As Object` — `Object` est l'équivalent le plus proche du `Variant`, idéal quand les arguments sont **hétérogènes** (chaîne, nombre, date…) :

  ```vb
  ' VB.NET — proche du Variant : tableau d'Object
  Sub Journaliser(ByVal format As String, ParamArray args() As Object)
      ' args ≈ le ParamArray Variant de VB6
  End Sub
  ```

- **Type fort** : `ParamArray valeurs() As Double` — préférable quand tous les arguments sont **du même type**, pour gagner en sûreté et en performance.

> 🔗 Passer par `Object()` réintroduit les considérations habituelles du `Variant` → `Object` : *boxing*, conversions, *late binding*. Voir le **[§ 7.1](../07-types-variables/01-variant-object.md)**. Quand c'est possible, **préférez un type fort**.

---

## 5. Tableau vide et bornes

Tester « aucun argument fourni » change de forme.

```vb
' VB6 — tableau "vide" : UBound < LBound (UBound vaut -1)
If UBound(valeurs) < LBound(valeurs) Then
    ' aucun argument
End If
```

```vb
' VB.NET — tableau de longueur 0 (jamais Nothing en appel variadic normal)
If valeurs.Length = 0 Then
    ' aucun argument
End If
```

Points à noter :

- en VB.NET, un `ParamArray` non fourni donne un **tableau de longueur 0**, **pas** `Nothing` — inutile (et trompeur) d'y appliquer les vieilles astuces `UBound`/`LBound`.
- les tableaux VB.NET sont **toujours 0-based** : `LBound` vaut 0, `UBound` vaut `Length - 1`. C'est le sujet du **[§ 8.1](../08-tableaux-collections/01-tableaux-0-based.md)**.

> 💡 *Code défensif* : si une procédure peut être appelée en lui passant explicitement `Nothing` comme tableau, un test `If valeurs Is Nothing OrElse valeurs.Length = 0` couvre les deux cas. En appel variadic normal, `Length = 0` suffit.

---

## 6. Règles et combinaisons

- **Un seul `ParamArray`**, et **en dernier** : inchangé entre VB6 et VB.NET.
- **Pas de `ParamArray` avec `Optional`** dans la même signature (vrai dans les deux langages) → voir **[§ 10.3](03-optional-ismissing.md)**.
- **Pas de modificateur `ByVal`/`ByRef` explicite** sur un `ParamArray` : absent en VB6 comme en VB.NET. Le mode est implicite — et c'est précisément ce mode implicite qui bascule (`ByRef` → `ByVal`).

---

## 7. Migration : outils, travail manuel et refactoring

### Ce que font les outils

- L'**[*Upgrade Wizard*](../04-outils-migration/01-upgrade-wizard.md)** convertit `ParamArray args()` en `ParamArray args() As Object` et adapte les sites d'appel. Pour l'usage **courant** (lecture des arguments : somme, concaténation, journalisation), la conversion est **fidèle**.
- Le cas **rare** où VB6 écrivait dans les variables de l'appelant est **hors de portée** d'une conversion automatique : VB.NET interdisant `ByRef ParamArray`, aucun outil ne peut reproduire cette sémantique à l'identique.

### Refactoriser le cas « écriture dans l'appelant »

Quand l'audit révèle un `ParamArray` utilisé pour **modifier** des variables de l'appelant, il faut **changer de conception** :

- **renvoyer** les valeurs (par la valeur de retour, un tableau, ou une structure) plutôt que les écrire en place ;
- ou exposer un paramètre **tableau explicite passé `ByRef`** — au prix de changer la convention d'appel (l'appelant doit construire un tableau).

Ce motif étant à la fois **rare** et **silencieux**, c'est un cas d'école pour les **tests de non-régression / golden master** (voir **[§ 10.2](02-byref-byval-audit.md)** et **[§ 17.4](../17-valider-refactoriser/04-traquer-pieges.md)**).

> 🛠️ **Règle pratique** : pour la quasi-totalité des `ParamArray` (lecture seule), la migration est **mécanique**. Le seul vrai travail de jugement concerne les `ParamArray` **écrits**, qui doivent être repérés et redessinés.

---

## 📋 Tableau de synthèse

| Aspect | VB6 | VB.NET (4.7.2) | Nature du changement |
|---|---|---|---|
| Type des éléments | `Variant` imposé | Tableau typé (`Object()` ou type fort) | ⚠️ Typage (→ § 7.1) |
| Mode de passage | `ByRef` | `ByVal` (non modifiable) | ⚠️ Sémantique silencieuse |
| Écriture dans l'appelant (args isolés) | Possible | **Perdue** (nouveau tableau) | ⚠️ Élevé (mais rare) |
| Passage d'un **tableau existant** | — | Élément modifiables (même objet) | Asymétrie nouvelle |
| Aucun argument | `UBound < LBound` | `Length = 0` | Mécanique |
| Bornes | 0-based | 0-based | ✅ Stable |
| `ByRef ParamArray` explicite | Interdit (et inutile) | **Interdit** | — |

---

## ✅ À retenir

1. **Le concept survit** : nombre variable d'arguments, en dernier paramètre, un seul par signature.
2. **Le type change** : `Variant()` → tableau typé. `Object()` pour rester fidèle à l'hétérogène ; **type fort** dès que possible.
3. **Le mode bascule `ByRef` → `ByVal`** : pour des arguments **isolés**, VB.NET crée un **nouveau tableau**, et les écritures **ne remontent plus**. Et l'on **ne peut pas** déclarer `ByRef ParamArray`.
4. **Le test du tableau vide change** : `Length = 0` en VB.NET, pas les astuces `UBound`/`LBound`.
5. **Lecture seule = migration mécanique ; écriture dans l'appelant = refactoring** à repérer et valider par les tests.

---

*Section suivante : **[§ 10.5 — Surcharge et valeurs de retour (`Return` vs `NomFonction = valeur`)](05-surcharge-retour.md)***

**Cible** : .NET Framework 4.7.2 · **Indicateurs** : ⭐ point clé · ⚠️ piège · 🛠️ outils

⏭️ [Surcharge (nouveauté .NET) et valeurs de retour (`Return` vs `NomFonction = valeur`)](/10-procedures-fonctions/05-surcharge-retour.md)
