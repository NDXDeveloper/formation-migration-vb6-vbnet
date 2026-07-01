🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.2 — Convertir vers `Try`/`Catch`/`Finally` (et les filtres `When`)

> **Le modèle cible.** Un bloc protégé (`Try`), des `Catch` **typés** ordonnés du **plus précis au plus général**, un `Finally` **toujours exécuté**, et des **filtres `When`** pour discriminer par condition. Plus puissant que `On Error` — et plus sûr, à condition de respecter quelques règles.

📍 *Module 11 · § 11.2 · [↑ Introduction du chapitre](README.md) · [← § 11.1](01-on-error.md)*

---

## ⚡ En bref

- Un `Try` doit comporter **au moins** un `Catch` **ou** un `Finally` (ou les deux).
- Les `Catch` sont évalués **de haut en bas** : on les classe du **plus spécifique au plus général**, le `Catch … As Exception` en **dernier**.
- Le `Finally` s'exécute **toujours** — y compris en cas de `Return` ou de relance. C'est le bon endroit pour le **nettoyage** (et il **dédoublonne** le nettoyage que VB6 écrivait deux fois).
- Les **filtres `When`** discriminent par **condition** (un numéro, un état) — l'équivalent souple du `Select Case Err.Number`.
- Pour relancer : **`Throw` nu** (préserve la pile). **Jamais `Throw ex`** (qui la réinitialise).

---

## 1. La structure de base

```vb
Try
    ' code protégé
Catch ex As FileNotFoundException
    ' cas spécifique
Catch ex As IOException
    ' cas plus large
Catch ex As Exception
    ' tout le reste
Finally
    ' nettoyage systématique
End Try
```

Un `Try` **ne peut pas** être seul : il lui faut **au moins** un `Catch` **ou** un `Finally`. Les combinaisons valides sont donc : `Try`+`Catch`, `Try`+`Finally`, ou `Try`+`Catch`+`Finally`.

---

## 2. Les clauses `Catch`

### 2.1 Capturer par type

`Catch ex As IOException` attrape les exceptions de ce type **et de ses types dérivés**. La variable (`ex`) donne accès à l'objet exception (message, pile, cause…) — détaillé au **[§ 11.4](04-err-vs-exception.md)**.

### 2.2 Plusieurs `Catch` : du plus précis au plus général

Quand il y a plusieurs `Catch`, **le premier dont le type correspond l'emporte**. Il faut donc les ordonner du **plus spécifique au plus général** ; sinon, un `Catch` général **masque** les suivants.

```vb
' ✅ correct : du plus précis au plus général
Try
    Lire(chemin)
Catch ex As FileNotFoundException      ' dérive de IOException
    ' ...
Catch ex As IOException
    ' ...
End Try
```

```vb
' ❌ incorrect : le général masque le spécifique
Try
    Lire(chemin)
Catch ex As IOException
    ' ...
Catch ex As FileNotFoundException      ' ⚠️ jamais atteint — le compilateur avertit
    ' ...
End Try
```

### 2.3 Le `Catch` attrape-tout

```vb
Catch ex As Exception   ' attrape toute exception .NET (forme recommandée)
' ou
Catch                   ' attrape-tout, sans variable
```

> 💡 Un `Catch` attrape-tout en **dernière** position sert de filet. Mais attraper **tout** sans rien en faire revient au piège du **[§ 11.3](03-on-error-resume-next.md)** : un `Catch` doit **traiter**, **journaliser**, ou **relancer** — pas se taire.

---

## 3. Le bloc `Finally` : le nettoyage garanti

Le `Finally` s'exécute **dans tous les cas** : erreur ou non, exception attrapée ou non, et **même** si le `Try` ou le `Catch` contient un `Return`, un `Exit` ou une relance. C'est la garantie qu'attendait tout nettoyage de ressources.

```vb
Try
    Ouvrir()
    Return                 ' même avec un Return…
Finally
    Fermer()               ' …le Finally s'exécute AVANT de quitter la procédure
End Try
```

Un `Try`+`Finally` **sans `Catch`** est **valide** : il **ne gère pas** l'erreur (elle se propage) mais **garantit** le nettoyage.

```vb
Try
    Travailler()
Finally
    Liberer()              ' nettoyage garanti ; l'exception continue de remonter
End Try
```

### ⭐ Avantage de migration : ne plus dédoublonner le nettoyage

En VB6, le nettoyage devait souvent être écrit **deux fois** — sur le chemin normal **et** dans le gestionnaire :

```vb
' VB6 — nettoyage DUPLIQUÉ
Sub Traiter()
    On Error GoTo Gestion
    Ouvrir()
    Travailler()
    Fermer()                ' nettoyage (chemin normal)
    Exit Sub
Gestion:
    Fermer()                ' nettoyage (chemin erreur) — la même ligne, en double
End Sub
```

Le `Finally` le centralise **une seule fois** :

```vb
' VB.NET — nettoyage UNIQUE et garanti
Sub Traiter()
    Try
        Ouvrir()
        Travailler()
    Catch ex As Exception
        Journaliser(ex)
    Finally
        Fermer()            ' une seule fois, quel que soit le chemin
    End Try
End Sub
```

> ⚠️ **Ne placez ni `Return` ni `Throw` dans un `Finally`** : cela peut **avaler** l'exception en cours et masquer le vrai problème. Le `Finally` ne doit faire que du **nettoyage**.
>
> 🔗 Pour libérer un objet **`IDisposable`**, préférez **`Using`** (sucre syntaxique d'un `Try`/`Finally` qui appelle `Dispose`) → voir **[§ 12.3](../12-poo/03-idisposable-using.md)**.

---

## 4. Les filtres `When`

Un filtre `When` ajoute une **condition** à un `Catch` : la clause ne se déclenche que si **le type correspond ET la condition est vraie**. Sinon, la recherche continue vers le `Catch` suivant (ou l'exception se propage).

```vb
' discriminer par CONDITION : ici un numéro d'erreur du fournisseur
Try
    ExecuterRequete()
Catch ex As SqlException When ex.Number = 1205    ' interblocage (deadlock)
    Reessayer()
Catch ex As SqlException
    Journaliser(ex.Message)
End Try
```

C'est l'équivalent souple du `Select Case Err.Number` lorsque la distinction ne porte **pas** sur le **type** mais sur une **valeur** ou un **état** au sein d'un même type.

> 🔎 **Subtilité utile** : le filtre `When` est évalué **avant** que la pile ne se déroule. On peut donc **inspecter le contexte** (et le journaliser) **avant** tout nettoyage intermédiaire.

**Idiome avancé — journaliser sans attraper.** Un filtre qui produit un effet de bord (journalisation) puis renvoie `False` laisse l'exception **se propager** tout en l'ayant tracée :

```vb
' Journaliser(ex) renvoie False → l'exception N'EST PAS attrapée ici, elle continue
Try
    Operation()
Catch ex As Exception When Journaliser(ex)
    ' jamais exécuté (le filtre a renvoyé False)
End Try
```

---

## 5. Relancer correctement : `Throw` vs `Throw ex`

Relancer une exception est fréquent (journaliser puis laisser remonter). **Le choix de l'écriture a des conséquences** :

```vb
Catch ex As Exception
    Journaliser(ex)
    Throw                 ' ✅ relance en PRÉSERVANT la pile d'origine
End Try
```

```vb
Catch ex As Exception
    Throw ex              ' ⚠️ RÉINITIALISE la pile à cette ligne — l'origine est perdue
End Try
```

`Throw` **nu** relance l'exception courante en conservant sa **pile d'appels** (où elle s'est réellement produite). `Throw ex` repart de la ligne de relance et **efface** cette information — un **piège fréquent** qui complique tout diagnostic ultérieur.

Pour **ajouter du contexte** sans perdre la cause, **enveloppez** :

```vb
Catch ex As SqlException
    ' on enrichit, on conserve la cause comme InnerException
    Throw New DataException("Échec de la mise à jour du compte", ex)
End Try
```

---

## 6. Mettre le tout ensemble : du gestionnaire VB6 au `Try`/`Catch`

Un gestionnaire VB6 typique — armement, discrimination par `Err.Number`, nettoyage — se convertit en `Try` avec **`Catch` typés** (et/ou `When`) et un **`Finally`** unique :

```vb
' VB6
Sub MettreAJour()
    On Error GoTo Gestion
    Connexion.Ouvrir
    Connexion.Executer requete
    Connexion.Fermer
    Exit Sub
Gestion:
    Select Case Err.Number
        Case 3704:  Journaliser "Connexion non ouverte"
        Case Else:  Journaliser Err.Description
    End Select
    If Connexion.EnCours Then Connexion.Fermer   ' nettoyage dans le gestionnaire
End Sub
```

```vb
' VB.NET  (types d'exception illustratifs — correspondance détaillée au § 11.4)
Sub MettreAJour()
    Try
        Connexion.Ouvrir()
        Connexion.Executer(requete)
    Catch ex As InvalidOperationException
        Journaliser("Connexion non ouverte")
    Catch ex As Exception
        Journaliser(ex.Message)
    Finally
        If Connexion.EnCours Then Connexion.Fermer()   ' nettoyage unique et garanti
    End Try
End Sub
```

On note les trois gains : discrimination par **type** plutôt que par numéro, nettoyage **centralisé** dans `Finally`, et disparition de l'`Exit Sub`/de l'étiquette.

---

## 7. Discriminer : type ou `When` ?

| Le besoin | Mécanisme |
|---|---|
| Distinguer des erreurs de **nature différente** (E/S, conversion, accès refusé…) | Plusieurs **`Catch` typés** |
| Distinguer par un **numéro / une valeur** au sein d'un **même type** (codes d'un fournisseur) | Filtre **`When`** |
| Réagir selon une **condition applicative** (réessai possible, état, heure…) | Filtre **`When`** |
| **Journaliser sans attraper** | `When` à effet de bord renvoyant `False`, ou `Catch` + **`Throw`** nu |
| **Nettoyer** quoi qu'il arrive | **`Finally`** (ou `Using`) |

---

## 📋 Tableau de synthèse

| Élément | Rôle | Règle / piège |
|---|---|---|
| `Try` | Délimite le code protégé | Exige au moins un `Catch` **ou** un `Finally` |
| `Catch … As Type` | Attrape un type (et dérivés) | **Du plus précis au plus général** ; base en dernier |
| `Catch` / `Catch … As Exception` | Attrape-tout | À traiter / journaliser / relancer — pas à taire |
| `Finally` | Nettoyage **garanti** | Toujours exécuté ; **pas** de `Return`/`Throw` dedans |
| `When` | Filtre conditionnel d'un `Catch` | Évalué avant déroulement de pile ; remplace `Select Case Err.Number` |
| `Throw` (nu) | Relance | **Préserve** la pile d'origine |
| `Throw ex` | Relance | ⚠️ **Réinitialise** la pile — à éviter |
| `Throw New …(…, ex)` | Enveloppe | Conserve la cause en `InnerException` |

---

## ✅ À retenir

1. **Ordonnez les `Catch`** du **plus précis au plus général** ; le `Catch … As Exception` vient **en dernier**.
2. **`Finally` s'exécute toujours** : c'est le nettoyage garanti — et il **supprime la duplication** que VB6 imposait (chemin normal + gestionnaire). N'y mettez **ni `Return` ni `Throw`**.
3. **Les filtres `When`** discriminent par **condition** : l'équivalent souple du `Select Case Err.Number`, idéal pour un numéro/état au sein d'un même type.
4. **Relancez avec `Throw` nu** pour **préserver la pile** ; **`Throw ex` est un piège**. Enveloppez (`Throw New …(…, ex)`) pour ajouter du contexte.
5. **Un `Catch` attrape-tout n'est pas une excuse pour ne rien faire** : traiter, journaliser ou relancer — jamais se taire (→ § 11.3).

---

*Section suivante : **[§ 11.3 — Le cas épineux de `On Error Resume Next` : pourquoi c'est dangereux à migrer](03-on-error-resume-next.md)** ⭐ ⚠️*

**Cible** : .NET Framework 4.7.2 · **Indicateurs** : ⭐ point clé · ⚠️ piège · 🛠️ outils

⏭️ [Le cas épineux de `On Error Resume Next` : pourquoi c'est dangereux à migrer](/11-gestion-erreurs/03-on-error-resume-next.md)
