🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.3 — Le type `Date` : un `Double` en VB6 → `DateTime` en .NET ⚠️

> **Chapitre 7 — Types, variables et déclarations** · Partie 3 — Migrer le langage
>
> Le mot-clé `Date` n'a pas changé… mais **la nature de la valeur**, oui. En VB6, une date **est un
> nombre** (un `Double`) : on l'additionne, on la soustrait, on lit son numéro de série. En .NET,
> `DateTime` est une **structure** — pas un nombre. Résultat : les idiomes les plus courants de VB6
> (`d + 1`, `d2 - d1`) **ne fonctionnent plus**, et la valeur « date vide » **change de repère**.

---

## 🧭 VB6 : une date **est** un nombre

En VB6, le type `Date` est stocké sous forme de **`Double`** (8 octets) :

- la **partie entière** = nombre de **jours** depuis l'époque **30 décembre 1899** (date sérielle) ;
- la **partie fractionnaire** = **heure du jour** (`0,5` = midi, `0,25` = 6 h).

Donc `0` ⟶ `1899-12-30 00:00`, `1` ⟶ `1899-12-31`, `2,5` ⟶ `1900-01-01 12:00`. Comme c'est un
`Double`, l'**arithmétique directe** est partout :

```vb
' VB6 — une date EST un nombre
Dim d As Date
d = Now
d = d + 1            ' demain (ajoute 1 jour)
d = d + 0.5          ' + 12 heures
Dim ecart As Double
ecart = d2 - d1      ' écart en JOURS (la fraction = les heures)
Dim serie As Double
serie = CDbl(Now)    ' le numéro de série sous-jacent
```

---

## 🎯 .NET : `DateTime` est une **structure**, pas un nombre

En .NET, `System.DateTime` est un **type valeur** (une structure) qui stocke un compte de **« ticks »**
(intervalles de 100 ns) depuis l'époque **1ᵉʳ janvier de l'an 0001**, plus une étiquette
**`DateTimeKind`** (`Unspecified` / `Utc` / `Local`). Ce **n'est pas** un `Double` : on **ne peut
pas** écrire `d + 1` pour ajouter un jour.

```vb
' VB.NET — DateTime est une STRUCTURE
Dim d As DateTime = DateTime.Now
d = d.AddDays(1)                    ' demain
d = d.AddHours(12)                  ' + 12 heures
Dim ecart As TimeSpan = d2 - d1     ' une DURÉE (TimeSpan), pas un nombre
Dim jours As Double = ecart.TotalDays
Dim serie As Double = DateTime.Now.ToOADate()   ' numéro de série OA (compatible VB6)
```

### Correspondance conceptuelle

| | **VB6 `Date`** | **.NET `DateTime`** |
|---|---|---|
| Nature | `Double` (nombre) | **structure** (type valeur) |
| Stockage | jours depuis 1899-12-30 | **ticks** (100 ns) depuis 0001-01-01 |
| Résolution | limitée (fraction de jour) | **100 ns** (bien plus fine) |
| Arithmétique | `+` / `-` directs | méthodes `AddDays`/`Add…` ; `-` → `TimeSpan` |
| « Date vide » par défaut | `0` = **1899-12-30** | **`DateTime.MinValue`** = 0001-01-01 |
| Fuseau horaire | aucune notion | **`DateTimeKind`** (Local/Utc/Unspecified) |

---

## 💥 Piège n°1 — l'arithmétique directe disparaît

C'est **le** piège majeur. Deux idiomes VB6 omniprésents cessent de fonctionner :

```vb
' ❌ Migration naïve — erreur de compilation (Option Strict On) ou faux (Option Strict Off)
d = d + 1
If d2 - d1 > 7 Then ...          ' compare une TimeSpan à un entier 7 !
```

`DateTime + Integer` n'est **pas défini** (seul `DateTime + TimeSpan` l'est), et `d2 - d1` renvoie
une **`TimeSpan`**, pas un nombre de jours. Deux corrections possibles :

```vb
' ✅ Idiomatique .NET
d = d.AddDays(1)
If (d2 - d1).TotalDays > 7 Then ...

' ✅ En conservant la sémantique VB6 (fonctions encore disponibles)
d = DateAdd(DateInterval.Day, 1, d)
If DateDiff(DateInterval.Day, d1, d2) > 7 Then ...
```

> 💡 Les fonctions **`DateAdd`**, **`DateDiff`**, **`DatePart`** existent toujours dans
> `Microsoft.VisualBasic` et **préservent les résultats VB6**. C'est une excellente passerelle : on
> migre d'abord à l'identique, puis on **refactorise** vers `AddDays`/`TimeSpan` quand on le décide.

---

## 💥 Piège n°2 — la passerelle `Date` ⟷ `Double` n'est plus implicite

En VB6, on glissait librement d'une date à son numéro de série (`CDbl(Now)`, `dt = 45000#`). En
.NET, ces conversions implicites **disparaissent** : une `DateTime` n'est pas un nombre. Le pont
explicite est la paire **`ToOADate()` / `DateTime.FromOADate()`** (OA = *OLE Automation*, le **même**  
système sériel que VB6/COM).

```vb
' Donnée VB6 persistée sous forme de Double sériel (fichier, base, Variant)…
Dim serieStockee As Double = LireDepuisFichier()
Dim dt As DateTime = DateTime.FromOADate(serieStockee)   ' ✅ reconstruction correcte

' …et inversement, pour écrire une valeur compatible avec l'existant VB6 :
Dim serie As Double = dt.ToOADate()
```

> ⚠️ **À ne pas confondre :** un `Double` issu de VB6 est un **numéro de série OA** (époque
> 1899-12-30). Il ne faut **surtout pas** l'interpréter comme des **ticks .NET** (époque 0001-01-01).
> Toujours passer par `FromOADate`/`ToOADate`. C'est vital pour la **compatibilité des données**
> (fichiers binaires, colonnes de base, exports) — un sujet d'interop traité au **module 16**.

> 🧪 **Dates anciennes :** par héritage de Lotus 1-2-3, **Excel** compte à tort un **29 février
> 1900** (1900 n'est pas bissextile) ; le système sériel **OA** (VB6/.NET) s'**aligne** sur cet
> écosystème (époque au 30/12/1899), d'où des **divergences possibles autour de février 1900**. Pour
> des dates **antérieures à 1900** ou proches de février 1900, **valider** la conversion par des
> tests plutôt que de la supposer exacte.

---

## 💥 Piège n°3 — la valeur « date vide » change de repère

Conséquence directe du changement d'époque, et **piège typique** : la valeur par défaut d'une date  
non initialisée **n'est plus la même**.

```vb
' VB6 : une Date non initialisée vaut 0 = 1899-12-30
If dt = 0 Then ...                     ' convention « pas de date »

' VB.NET : une Date non initialisée vaut DateTime.MinValue = 0001-01-01 !
If dt = DateTime.MinValue Then ...     ' le sentinelle a changé
```

Tout code qui testait `dt = 0` (ou `dt = CDate(0)`) pour signifier « **non renseigné** » doit être  
revu. Et comme `DateTime` est un **type valeur**, `dt = Nothing` ne donne **pas** un « vide » : il  
assigne la valeur par défaut, c'est-à-dire **`DateTime.MinValue`**. Pour exprimer une **absence**  
réelle de valeur, utiliser un type **nullable** :

```vb
Dim dtOpt As DateTime? = Nothing       ' Nullable(Of DateTime)
If dtOpt.HasValue Then Traiter(dtOpt.Value)
```

> 🔗 La disparition des états (`Empty`/`Null`/`Missing`) et le rôle de `Nothing` sont approfondis en
> **7.7**.

---

## 💥 Piège n°4 — précision : un aller-retour par `Double` **perd** des décimales

`DateTime` (ticks à 100 ns) est **plus précis** que l'ancien `Double` (fraction de jour, précision
sous-seconde médiocre). Donc un aller-retour `DateTime → ToOADate → FromOADate` peut **perdre les  
millisecondes** :

```vb
Dim t1 As DateTime = DateTime.Now                  ' précis à 100 ns
Dim t2 As DateTime = DateTime.FromOADate(t1.ToOADate())
' t2 ≈ t1, mais la partie sous-seconde peut différer
```

Conséquence pratique : éviter de **comparer pour égalité stricte** des dates qui ont transité par un
`Double`, et ne persister en sériel que ce qui n'a pas besoin de précision fine.

---

## 💥 Piège n°5 — analyse et formatage **sensibles à la culture**

`CDate`, `Format`, `DateTime.Parse` et `ToString` dépendent de la **culture courante**
(`CultureInfo.CurrentCulture`). `"01/02/2026"` se lit **2 janvier** (en-US) ou **1ᵉʳ février**
(fr-FR) :

```vb
' Ambigu et fragile (dépend de la culture d'exécution)
Dim d1 As Date = CDate("01/02/2026")

' Robuste pour des données persistées / d'échange : format ET culture explicites
Dim d2 As Date = DateTime.ParseExact("2026-02-01", "yyyy-MM-dd",
                                     Globalization.CultureInfo.InvariantCulture)
Dim s As String = d2.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)
```

> ⚠️ Une application migrée tournant sous une **autre culture** que celle de développement peut lire
> et écrire les dates **différemment**. Pour tout ce qui est **stocké, échangé ou comparé**,
> préférer un **format explicite** (`ParseExact`/`ToString(format, culture)`) et la **culture
> invariante**. C'est la même famille de pièges que **Annexe B.10**. Préférer aussi les **années à
> 4 chiffres** : l'interprétation des années à 2 chiffres (fenêtre glissante) peut différer.

---

## 💥 Piège n°6 — `DateTimeKind` et fuseaux horaires (nouvel axe)

VB6 ignorait totalement les fuseaux. .NET ajoute un **`Kind`** à chaque `DateTime` :

- `DateTime.Now` → heure **locale** (`Kind = Local`) ;
- `DateTime.UtcNow` → **UTC** ; `DateTime.Today` → date locale à minuit.

Pour une logique purement **locale**, ignorer le `Kind` reste souvent sans conséquence. Mais dès que  
des dates **traversent des fuseaux**, sont **sérialisées** (ISO 8601 avec décalage) ou stockées en  
base, cela devient un vrai sujet — et l'on dispose alors de **`DateTimeOffset`** (date + décalage)  
pour lever l'ambiguïté. *(Sujet avancé, à traiter au besoin lors de la fiabilisation — module 17.)*

---

## 🔁 Les fonctions de date

Bonne nouvelle (à double tranchant) : **beaucoup de fonctions VB6 existent encore** dans
`Microsoft.VisualBasic`, donc une grande partie du code **compile sans modification**.

| VB6 | Équivalent encore disponible | Équivalent idiomatique .NET |
|-----|------------------------------|-----------------------------|
| `Now` | `Now` | `DateTime.Now` |
| `Date` (date du jour) | `Today` | `DateTime.Today` |
| `Time` | `TimeOfDay` | `DateTime.Now.TimeOfDay` (`TimeSpan`) |
| `DateSerial(a,m,j)` | `DateSerial(a,m,j)` | `New DateTime(a, m, j)` |
| `DateAdd` / `DateDiff` / `DatePart` | *(idem, sémantique VB6)* | `AddDays`… / `TimeSpan` / `.Year`,`.Month`… |
| `Year`/`Month`/`Day`/`Hour`… | *(idem)* | `dt.Year`, `dt.Month`, `dt.Day`, `dt.Hour`… |
| `Format(dt, f)` | `Format(dt, f)` | `dt.ToString(f, culture)` |
| `CDate` / `IsDate` | `CDate` / `IsDate` | `DateTime.Parse` / `ParseExact` / `TryParse` |

> ⚠️ **Le double tranchant :** comme ces fonctions compilent, on **croit** la migration des dates
> terminée. Or l'**arithmétique avec `+`/`-`** et la **valeur « vide »** ne suivent **pas** — ce
> sont précisément les pièges 1 et 3 ci-dessus.

---

## 🧭 Stratégie de migration

1. **Traquer l'arithmétique de dates** (`+`, `-`, comparaisons d'écarts). C'est la priorité absolue :
   remplacer par `AddDays`/`TimeSpan`, ou conserver `DateAdd`/`DateDiff` en transition.
2. **Repérer les `Date` ⟷ `Double`** (numéros de série persistés, dates dans des `Variant`, calculs
   sur la valeur brute) → `ToOADate`/`FromOADate` **explicites**.
3. **Auditer les sentinelles `dt = 0`** (« pas de date ») → `DateTime.MinValue`, ou mieux,
   `DateTime?` pour une vraie absence de valeur.
4. **Fixer culture et format** pour toute date **stockée, échangée ou comparée**
   (`ParseExact`/`InvariantCulture`, années à 4 chiffres).
5. **Décider du `Kind`** seulement là où les fuseaux comptent ; sinon, rester simple.
6. **Filet de sécurité** : le ***golden master*** (chapitre 5.5) compare les dates calculées avant /
   après et révèle un écart d'arithmétique, d'arrondi ou de sentinelle.

---

## ✅ À retenir

- `Date` **n'est plus un nombre** : `DateTime` est une **structure**. `d + 1` et `d2 - d1`
  (→ `TimeSpan`) **ne se migrent pas tels quels**.
- Le **pont vers le `Double` sériel** est explicite : **`ToOADate` / `FromOADate`** (compatible VB6),
  à ne pas confondre avec les **ticks** .NET.
- La **« date vide » a changé** : `1899-12-30` (VB6) → **`DateTime.MinValue` = 0001-01-01** (.NET).
  Pour « non renseigné », utiliser **`DateTime?`**.
- **Précision** plus fine en .NET : un aller-retour par `Double` **perd** le sous-seconde.
- **Analyse/formatage sensibles à la culture** : format **explicite** + **culture invariante** pour
  les données persistées.
- Beaucoup de **fonctions VB6 subsistent** (compile sans changement) — d'où le **faux sentiment**
  que les dates sont déjà migrées.

---

## 🔗 Pour aller plus loin

- **7.1 — `Variant` → `Object`** : conversions et `Option Strict` (les passerelles implicites
  supprimées).
- **7.7 — `Nothing`, `Empty`, `Null`, `Missing`** : la « date vide » et l'absence de valeur.
- **Chapitre 9** — opérateurs et fonctions de chaînes (comparaisons, formatage).
- **Module 16** — interopérabilité COM et persistance binaire : où `ToOADate`/`FromOADate` sont
  indispensables.
- **Annexe B.9** (dates et arrondis) et **B.10** (conversions sensibles à la culture) ; **Annexe D**
  (correspondance des types).

➡️ Section suivante : **[7.4 — Chaînes de longueur fixe](04-chaines-longueur-fixe.md)**  
⬆️ Retour au **[sommaire du chapitre 7](README.md)**

⏭️ [Chaînes de **longueur fixe** (`Dim s As String * 10`) : `VBFixedString` ou refactoring](/07-types-variables/04-chaines-longueur-fixe.md)
