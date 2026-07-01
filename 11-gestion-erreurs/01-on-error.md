🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.1 — `On Error GoTo`, `Resume`, `Resume Next` : équivalents et logique de remplacement

> **Les mécanismes VB6 d'armement et de reprise — et comment les *re-exprimer*.** Il n'existe **pas** de traduction mot à mot : `Resume` (réessayer) et `Resume Next` (poursuivre) n'ont **aucun équivalent direct** en `Try`/`Catch`. Le travail consiste à reconstruire la **logique**, pas à transposer la syntaxe.

📍 *Module 11 · § 11.1 · [↑ Introduction du chapitre](README.md)*

---

## ⚡ En bref

- VB6 arme un gestionnaire d'erreur avec **`On Error GoTo label`** : à la moindre erreur, le flux **saute** vers une étiquette, dans la **même procédure**.
- La famille **`Resume`** (utilisée *dans* le gestionnaire) décide **où reprendre** ensuite : réessayer la ligne fautive (`Resume`), passer à la suivante (`Resume Next`), ou aller à une étiquette (`Resume label`).
- VB.NET n'offre **ni saut vers un gestionnaire séparé, ni retour dans le code protégé** : après un `Catch`, on continue **après** le bloc `Try`. Réessai et « poursuivre » se **reconstruisent** (boucle, ou `Try` resserré).

---

## ⚠️ D'abord, lever une ambiguïté : deux « Resume Next » différents

Le mot « Resume Next » apparaît dans **deux constructions distinctes** qu'il ne faut surtout pas confondre :

| Écriture | Catégorie | Effet | Traité |
|---|---|---|---|
| `On Error Resume Next` | **Mode** (à l'armement) | Ignore l'erreur et continue, **pour tout le code qui suit** | → **[§ 11.3](03-on-error-resume-next.md)** ⚠️ |
| `Resume Next` | **Instruction** (dans un gestionnaire) | Après traitement, reprend à la ligne **suivant** la ligne fautive | **ici** |
| `Resume` | Instruction (dans un gestionnaire) | Après traitement, **réessaie** la ligne fautive | **ici** |
| `Resume label` | Instruction (dans un gestionnaire) | Après traitement, reprend à `label` | **ici** |

> 💡 Retenez la distinction : **`On Error Resume Next`** est un **mode** d'ignorance globale (le cas dangereux du chapitre, traité au § 11.3). Les **instructions `Resume`** sont des **outils de reprise** employés *à l'intérieur* d'un gestionnaire `On Error GoTo`. Cette section traite ces dernières.

---

## 1. Le modèle VB6 : armer un gestionnaire et y brancher

### 1.1 `On Error GoTo label` — armer

`On Error GoTo Gestion` installe un gestionnaire : toute erreur survenant **après** cette ligne, dans la **même procédure**, provoque un saut vers l'étiquette `Gestion`. La structure canonique place le gestionnaire **en fin de procédure**, précédé d'un `Exit Sub`/`Exit Function` pour éviter d'y « tomber » en exécution normale :

```vb
' VB6 — structure canonique
Sub Traiter()
    On Error GoTo Gestion
    Etape1()
    Etape2()
    Exit Sub                         ' empêche d'entrer dans le gestionnaire sans erreur
Gestion:
    Journaliser Err.Number, Err.Description
End Sub
```

Points clés :

- la portée est **la procédure** : un seul gestionnaire couvre tout le code situé après le `On Error` ;
- une erreur **non gérée localement** se **propage à l'appelant** (à *son* gestionnaire) — comme les exceptions, mais sans typage ;
- on peut **ré-armer** en cours de route : un nouveau `On Error GoTo AutreGestion` **remplace** le précédent.

### 1.2 `On Error GoTo 0` — désarmer

`On Error GoTo 0` **désactive** le gestionnaire courant : au-delà, les erreurs redeviennent non gérées (et se propagent).

> 🔎 *Forme rare* : `On Error GoTo -1` **réinitialise l'erreur active** (pour pouvoir en piéger une nouvelle dans le même gestionnaire). Peu utilisée, mais on peut la croiser dans du code ancien.

### 1.3 Discriminer les erreurs

Dans le gestionnaire, VB6 distingue souvent les cas avec `Select Case Err.Number` :

```vb
Gestion:
    Select Case Err.Number
        Case 53:  ' fichier introuvable
        Case 70:  ' permission refusée
        Case Else
    End Select
```

> 🔗 Cette discrimination par **numéro** devient une discrimination par **type** d'exception (plusieurs `Catch`, filtres `When`) — voir **[§ 11.2](02-vers-try-catch.md)** et **[§ 11.4](04-err-vs-exception.md)**.

---

## 2. La famille `Resume` : où reprendre après le gestionnaire

Une fois l'erreur traitée, une instruction `Resume` décide **où** l'exécution repart **dans le code normal**.

| Instruction | Reprise |
|---|---|
| `Resume` | **Réexécute** l'instruction qui a échoué (réessai) |
| `Resume Next` | Continue à l'instruction **suivant** celle qui a échoué (saute la fautive) |
| `Resume label` | Continue à l'étiquette `label` |

```vb
' VB6 — réessai après correction de la condition
Sub Ouvrir()
    On Error GoTo Gestion
Tentative:
    OuvrirRessource()
    Exit Sub
Gestion:
    If CorrigerCondition() Then
        Resume Tentative            ' on réessaie
    End If
End Sub
```

C'est ce **retour dans le flux normal** — réessayer, ou sauter, ou aller ailleurs — qui n'a pas d'équivalent direct en `Try`/`Catch`.

---

## 3. Le changement de modèle (la clé de cette section)

| | VB6 | VB.NET |
|---|---|---|
| Le gestionnaire est… | une **région séparée** où l'on **saute** | un bloc **`Catch` inline** |
| Après traitement, on… | **décide où revenir** via `Resume` | continue **après** le `Try` (pas de retour dans le code protégé) |
| Portée | la **procédure** (depuis le `On Error`) | le **bloc** `Try` (granularité choisie) |
| Réessayer / sauter | `Resume` / `Resume Next` | à **reconstruire** (boucle / `Try` resserré) |

La bascule mentale est là : en VB6, on **sort** du flux pour gérer l'erreur, puis on **réinjecte** le flux où l'on veut. En VB.NET, le `Catch` est **sur place** ; une fois passé (avec l'éventuel `Finally`), l'exécution **reprend naturellement après le `Try`** — il n'existe pas de « rentrer à nouveau » dans le code protégé. Pour réessayer, on **boucle** ; pour ignorer une instruction précise, on **resserre** le `Try` autour d'elle.

> 💡 **Conséquence sur la granularité** : VB6 couvre toute la procédure avec un gestionnaire ; VB.NET vous laisse **choisir l'étendue** — tout le corps, ou seulement les instructions risquées. Plus fin = plus précis. C'est un gain, mais une **décision de périmètre** à prendre.

---

## 4. La logique de remplacement, cas par cas

> La **syntaxe** complète de `Try`/`Catch`/`Finally` est détaillée au **[§ 11.2](02-vers-try-catch.md)**. On présente ici la **correspondance de logique**.

### 4.1 `On Error GoTo` → `Try`/`Catch` (le cas direct)

Le code risqué passe dans `Try`, le gestionnaire dans `Catch`. L'`Exit Sub` et l'étiquette **disparaissent** : le `Catch` ne s'exécute **que** sur erreur.

```vb
' VB.NET
Sub Traiter()
    Try
        Etape1()
        Etape2()
    Catch ex As Exception
        Journaliser(ex)              ' plus d'Exit Sub, plus d'étiquette
    End Try
End Sub
```

### 4.2 `Resume` (réessai) → boucle de réessai autour du `Try`

Le réessai se reconstruit avec une **boucle**, en prévoyant une **garde** contre la boucle infinie.

```vb
' VB.NET — équivalent logique de Resume (réessayer l'instruction)
Dim tentatives As Integer = 0
Do
    Try
        OuvrirRessource()
        Exit Do                      ' succès → on sort de la boucle
    Catch ex As IOException
        tentatives += 1
        If tentatives >= 3 Then Throw ' garde : on abandonne après N essais
        ' éventuellement corriger la condition avant de reboucler
    End Try
Loop
```

### 4.3 `Resume Next` (instruction) → `Try` resserré autour de l'instruction

« Poursuivre malgré l'échec de cette instruction » se traduit en **isolant** l'instruction dans son propre `Try`, de sorte que la suivante s'exécute quoi qu'il arrive.

```vb
' VB.NET — équivalent logique de "Resume Next" (poursuivre après la fautive)
Try
    EtapeFacultative()
Catch ex As Exception
    ' décision EXPLICITE d'ignorer (idéalement : journaliser)
End Try
EtapeSuivante()                      ' s'exécute, que la précédente ait échoué ou non
```

> ⚠️ **Ne confondez pas avec `On Error Resume Next`.** Ici, l'ignorance est **resserrée sur une seule instruction** et **délibérée** : c'est légitime. Ignorer **en bloc** toutes les erreurs d'une procédure est le cas dangereux du **[§ 11.3](03-on-error-resume-next.md)**. La règle : un `Catch` qui ignore doit être **étroit**, **justifié**, et de préférence **journaliser**.

### 4.4 `Resume label` → restructuration du flux

Reprendre à une étiquette n'a pas d'équivalent : on **restructure**. Le plus souvent, le code visé par l'étiquette devient ce qui suit **après** le `Try` (puisque l'exécution y reprend naturellement), ou l'on **découpe** en plusieurs `Try` distincts selon les chemins voulus.

### 4.5 Gestionnaires multiples / ré-armement → `Try` imbriqués ou successifs

Un `On Error GoTo H1` suivi plus loin d'un `On Error GoTo H2` (chaque armement remplaçant le précédent) se traduit par **plusieurs blocs `Try`** — successifs ou imbriqués — chacun avec son `Catch` adapté à la portion de code concernée.

---

## 📋 Tableau de correspondance

| VB6 | Logique VB.NET | Remarque |
|---|---|---|
| `On Error GoTo label` + étiquette | `Try … Catch` | L'`Exit Sub` et l'étiquette disparaissent |
| `On Error GoTo 0` | Fin de portée d'un `Try` | Désarme |
| `Select Case Err.Number` (dans le gestionnaire) | Plusieurs `Catch` typés + `When` | → § 11.2 / § 11.4 |
| `Resume` | **Boucle de réessai** autour du `Try` (avec garde) | Pas d'équivalent direct |
| `Resume Next` (instruction) | `Try` **resserré** sur l'instruction, puis on continue | ⚠️ Étroit et délibéré (≠ § 11.3) |
| `Resume label` | **Restructuration** du flux (`Try` découpés) | Pas d'équivalent direct |
| Ré-armement (`On Error GoTo` multiple) | `Try` successifs / imbriqués | Granularité choisie |

---

## ✅ À retenir

1. **Deux « Resume Next »** : le **mode** `On Error Resume Next` (ignorer en bloc → § 11.3) n'a rien à voir avec l'**instruction** `Resume Next` (poursuivre après la fautive, dans un gestionnaire).
2. **Changement de modèle** : VB6 **saute** vers un gestionnaire séparé puis **réinjecte** le flux via `Resume` ; VB.NET traite **sur place** et reprend **après** le `Try` — **pas** de retour dans le code protégé.
3. **`On Error GoTo` → `Try`/`Catch`** est direct ; l'`Exit Sub` et l'étiquette **disparaissent**.
4. **`Resume` se reconstruit** : réessai → **boucle** (avec garde anti-boucle) ; « poursuivre » → **`Try` resserré** sur l'instruction.
5. **Choisissez la granularité** : VB.NET protège un **bloc**, pas toute la procédure — plus fin, plus précis, mais c'est une décision de périmètre.

---

*Section suivante : **[§ 11.2 — Convertir vers `Try`/`Catch`/`Finally` (et les filtres `When`)](02-vers-try-catch.md)***

**Cible** : .NET Framework 4.7.2 · **Indicateurs** : ⭐ point clé · ⚠️ piège · 🛠️ outils

⏭️ [Convertir vers `Try`/`Catch`/`Finally` (et les filtres `When`)](/11-gestion-erreurs/02-vers-try-catch.md)
