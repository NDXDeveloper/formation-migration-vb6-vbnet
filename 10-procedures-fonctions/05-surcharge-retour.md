🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.5 — Surcharge (nouveauté .NET) et valeurs de retour (`Return` vs `NomFonction = valeur`)

> **Deux nouveautés/différences à exploiter.** La **surcharge** — absente de VB6 — permet de remplacer proprement les vieux contournements (`Optional`, `Variant`). Et côté retour, `Return` **n'est pas** un simple synonyme de `NomFonction = valeur` : il **sort immédiatement** de la fonction.

📍 *Module 10 · § 10.5 · [↑ Introduction du chapitre](README.md) · [← § 10.4](04-paramarray.md)*

---

## ⚡ En bref

- **Surcharge** : en VB.NET, plusieurs procédures peuvent porter **le même nom** si leurs **signatures diffèrent** (nombre et/ou types de paramètres). VB6 ne le permettait pas. C'est une **opportunité**, pas un piège : elle remplace avantageusement les patterns `Optional`/`IsMissing` (**[§ 10.3](03-optional-ismissing.md)**) et les `Variant` qui testaient le type à l'exécution (**[§ 7.1](../07-types-variables/01-variant-object.md)**).
- **Valeurs de retour** : VB6 n'avait qu'**une** façon de renvoyer une valeur — affecter le **nom de la fonction**. VB.NET en ajoute une seconde, `Return`, qui **affecte *et* sort**. Convertir naïvement `NomFonction = valeur` en `Return` peut provoquer une **sortie prématurée** ⚠️.

---

## Partie 1 — La surcharge

### 1. VB6 n'avait pas de surcharge

En VB6, **chaque nom de procédure devait être unique** dans sa portée. Pour gérer « plusieurs variantes d'une même opération », on utilisait des **contournements** :

- des **noms différents** (`AireCarre`, `AireRectangle`…) ;
- un paramètre **`Optional`** avec branchement interne ;
- un paramètre **`Variant`** dont on testait le **type à l'exécution** (`TypeName`, `VarType`).

### 2. VB.NET : même nom, signatures distinctes

VB.NET autorise la **surcharge** : plusieurs procédures partagent un nom, et le compilateur choisit la bonne selon les **arguments** fournis.

```vb
' VB.NET — deux surcharges du même nom
Function Aire(ByVal cote As Double) As Double
    Return cote * cote
End Function

Function Aire(ByVal largeur As Double, ByVal hauteur As Double) As Double
    Return largeur * hauteur
End Function

' Aire(5)        → première surcharge
' Aire(4, 3)     → seconde surcharge
```

### 3. Les règles à connaître

Les surcharges doivent **différer par leur liste de paramètres** — c'est-à-dire par le **nombre** et/ou les **types** des paramètres. Plusieurs distinctions **ne suffisent pas** :

| Différence entre deux surcharges | Permise ? |
|---|---|
| Nombre de paramètres | ✅ Oui |
| Types des paramètres | ✅ Oui |
| **Seul** le type de **retour** | ❌ **Non** |
| **Seul** le mode `ByVal` / `ByRef` | ❌ Non |
| **Seul** le **nom** des paramètres | ❌ Non |

```vb
' VB.NET — ❌ NON valide : surcharge par le seul type de retour
Function Lire(ByVal cle As String) As Integer
    ' ...
End Function
Function Lire(ByVal cle As String) As String     ' ❌ même liste de paramètres
    ' ...
End Function
```

> 💡 Le **mode de passage** (`ByVal`/`ByRef`) ne fait **pas** partie de la signature de surcharge — un rappel utile après le **[§ 10.2](02-byref-byval-audit.md)** : on ne peut pas distinguer deux surcharges par ce seul critère.

**Le mot-clé `Overloads`** est **optionnel** à l'intérieur d'une même classe (mais s'il est présent, il doit l'être sur **toutes** les surcharges du nom). Il devient utile, voire nécessaire, dès qu'on **surcharge un membre hérité** d'une classe de base — un sujet qui relève de l'héritage, traité au **[§ 12.5](../12-poo/05-interfaces-heritage.md)**.

### 4. Pourquoi c'est utile en migration

La surcharge est l'outil idéal pour **remplacer les contournements VB6** :

- au lieu d'un **`Optional` ambigu** (où l'on ne distingue plus « omis » de « fourni » — **[§ 10.3](03-optional-ismissing.md)**), deux surcharges **explicites** :

  ```vb
  Sub Configurer()
      Configurer(3)                 ' "sans argument" = comportement par défaut
  End Sub
  Sub Configurer(ByVal niveau As Integer)
      ' ...
  End Sub
  ```

- au lieu d'un **`Variant`** qui testait le type à l'exécution, des surcharges **fortement typées** (une par type attendu), ce qui supprime le *late binding* et fiabilise le code.

> ⚠️ **Ne mélangez pas `Optional` et surcharge** sur un même nom sans précaution : une surcharge sans le paramètre et une surcharge avec ce paramètre **optionnel** peuvent rendre certains appels **ambigus**. Choisissez **un** mécanisme — surcharge **ou** paramètre optionnel — pas les deux pour la même intention.

---

## Partie 2 — Les valeurs de retour

### 5. VB6 : une seule façon — affecter le nom de la fonction

En VB6, on renvoie une valeur en l'affectant au **nom de la fonction**. Il n'existe **pas** d'instruction `Return valeur`.

```vb
' VB6
Function Doubler(x As Integer) As Integer
    Doubler = x * 2               ' seule façon de définir la valeur de retour
End Function
```

> 🔗 **Faux-ami** : le mot-clé `Return` **existe** en VB6, mais uniquement pour revenir d'un `GoSub` (`GoSub … / … Return`) — un mécanisme **supprimé** en VB.NET (**[§ 9.5](../09-operateurs-controle/05-gosub-ongoto-supprimes.md)**). En VB.NET, `Return` est **réaffecté** à un tout autre rôle : renvoyer une valeur. Même mot, sens radicalement différent.

### 6. VB.NET : deux façons, qui ne sont pas équivalentes

VB.NET conserve l'affectation au nom **et** ajoute `Return`. La différence est **comportementale** :

| Forme | Définit la valeur ? | Sort de la fonction ? |
|---|---|---|
| `Return expr` | Oui | **Oui, immédiatement** |
| `NomFonction = expr` | Oui | **Non** — l'exécution **continue** |

Avec `NomFonction = expr`, la fonction ne se termine qu'en atteignant `End Function`, un `Exit Function` ou un `Return` ; **le code qui suit s'exécute encore** et peut même **réécrire** la valeur. Avec `Return expr`, tout ce qui suit sur ce chemin est **ignoré**.

### 7. ⚠️ Le piège de la conversion naïve

C'est le point de vigilance de cette section. Convertir mécaniquement `NomFonction = expr` en `Return expr` **change le comportement** dès qu'il reste du code après l'affectation.

```vb
' VB6 — on affecte le nom, PUIS on continue (plafonnement)
Function Calculer(x As Integer) As Integer
    Calculer = x * 2
    If x > 100 Then
        Calculer = 100            ' réaffectation : la valeur finale est plafonnée
    End If
End Function                      ' renvoie la DERNIÈRE valeur affectée
```

```vb
' VB.NET — ❌ conversion naïve en Return : SORTIE PRÉMATURÉE
Function Calculer(x As Integer) As Integer
    Return x * 2                  ' ⚠️ sort ICI — le plafonnement ne s'exécute jamais
    If x > 100 Then
        Return 100                ' code inatteignable
    End If
End Function
```

Deux corrections possibles :

```vb
' VB.NET — ✅ option 1 : conserver l'affectation au nom (comportement identique à VB6)
Function Calculer(x As Integer) As Integer
    Calculer = x * 2
    If x > 100 Then Calculer = 100
End Function                      ' la valeur affectée est renvoyée en fin de fonction
```

```vb
' VB.NET — ✅ option 2 : restructurer pour un seul Return en fin
Function Calculer(x As Integer) As Integer
    Dim resultat As Integer = x * 2
    If x > 100 Then resultat = 100
    Return resultat
End Function
```

> 💡 En VB.NET, **affecter le nom puis « tomber » sur `End Function`** renvoie bien la valeur affectée : un `Return` final n'est pas obligatoire. C'est ce qui rend l'**option 1** sûre et fidèle au comportement VB6.

### 8. Un filet de sécurité offert par VB.NET

Là où VB6 renvoyait **silencieusement** la valeur par défaut du type si aucune affectation n'avait eu lieu, VB.NET **avertit** lorsqu'une fonction peut se terminer **sans** avoir défini sa valeur sur **tous les chemins** (« la fonction ne renvoie pas de valeur sur tous les chemins du code »). C'est un changement **« bruyant »** bienvenu : il attire l'attention sur un oubli que VB6 masquait.

`Exit Function` (et `Exit Sub`) existent toujours en VB.NET, avec le même sens qu'en VB6 : ils sortent de la procédure en renvoyant, pour une fonction, la **dernière valeur affectée** au nom. `Return` en est l'équivalent **concis** qui définit *et* sort en une seule instruction.

### 9. Recommandation

- **Code neuf ou refactorisé** : préférez **`Return`** — plus lisible, points de sortie explicites.
- **Migration** : **conservez `NomFonction = valeur`** tant que la logique repose sur « affecter puis continuer / réaffecter ». Ne passez à `Return` qu'après avoir vérifié que l'affectation est bien la **dernière** action sur ce chemin — ou en **restructurant** délibérément. L'**[*Upgrade Wizard*](../04-outils-migration/01-upgrade-wizard.md)** conserve d'ailleurs l'affectation au nom, par sécurité.

---

## 📋 Tableau de synthèse

| Aspect | VB6 | VB.NET (4.7.2) | Nature |
|---|---|---|---|
| Surcharge de procédures | **Inexistante** | Disponible (signatures distinctes) | ✅ Opportunité |
| Surcharge par le seul type de retour / mode / nom | — | **Interdite** | Règle |
| Mot-clé `Overloads` | — | Optionnel (cohérent) ; utile avec l'héritage | → § 12.5 |
| Définir la valeur de retour | `NomFonction = valeur` (seule façon) | `NomFonction = valeur` **ou** `Return` | Ajout |
| `Return` | Retour d'un `GoSub` (→ supprimé) | Renvoie une valeur **et sort** | ⚠️ Faux-ami |
| `Return` vs affectation au nom | — | `Return` **sort** ; l'affectation **continue** | ⚠️ Sémantique |
| Chemin sans valeur de retour | Défaut renvoyé **en silence** | **Avertissement** du compilateur | ✅ Filet bruyant |

---

## ✅ À retenir

1. **La surcharge est une nouveauté .NET** : même nom, signatures différentes (nombre/types de paramètres) — **jamais** par le seul type de retour, mode de passage ou nom de paramètre.
2. **C'est l'outil de nettoyage idéal** des contournements VB6 : il remplace `Optional`/`IsMissing` et les `Variant` testés à l'exécution par des surcharges **fortement typées**.
3. **`Return` n'est pas `NomFonction = valeur`** : `Return` **sort immédiatement**, l'affectation **laisse continuer** (et autorise une réaffectation).
4. **Attention à la conversion naïve** en `Return` : elle peut provoquer une **sortie prématurée**. En migration, **gardez l'affectation au nom** quand du code suit.
5. **VB.NET vous avertit** des chemins sans valeur de retour — là où VB6 renvoyait le défaut en silence.

---

*Section suivante : **[§ 10.6 — `Property Get`/`Let`/`Set` → propriétés](06-property-get-let-set.md)** ⚠️*

**Cible** : .NET Framework 4.7.2 · **Indicateurs** : ⭐ point clé · ⚠️ piège · 🛠️ outils

⏭️ [`Property Get`/`Let`/`Set` → propriétés (`Let` **et** `Set` fusionnés en un unique accesseur `Set` — un faux-ami du `Set` VB6)](/10-procedures-fonctions/06-property-get-let-set.md)
