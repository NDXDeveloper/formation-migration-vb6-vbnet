🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.5 — Stratégie de migration progressive (cohabitation `On Error` / `Try` temporaire)

> **Pas besoin de tout réécrire d'un coup.** `On Error` **survit** en VB.NET : les deux modèles peuvent **cohabiter** au niveau de l'application — mais **pas dans une même procédure**. L'**unité de migration** est donc **la procédure** : on convertit une par une, en validant à chaque étape, jusqu'à converger vers du `Try`/`Catch` partout.

📍 *Module 11 · § 11.5 · [↑ Introduction du chapitre](README.md) · [← § 11.4](04-err-vs-exception.md)*

---

## ⚡ En bref

- VB.NET **conserve** `On Error`, `Err`, `Resume`, `Err.Raise` (compatibilité) : le code migré mécaniquement **compile et tourne** en style VB6.
- **Contrainte unique** : on **ne mélange pas** `On Error` et `Try`/`Catch` dans la **même procédure** (erreur de compilation). → l'**unité de migration est la procédure** (tout `On Error` **ou** tout `Try`).
- **Rassurant** : les deux modèles s'**entendent à la frontière**, car `On Error` repose en VB.NET **sur les exceptions**. Une exception (`Throw` ou `Err.Raise`) traverse correctement les deux mondes.
- **Démarche** : base mécanique → **filet** (golden master) → conversion **procédure par procédure** → **élimination** des résidus.

---

## 1. Ce qui rend la cohabitation possible

VB.NET prend en charge **les deux modèles** d'erreur. On peut donc avoir, dans la même application, des procédures encore en `On Error` et d'autres déjà en `Try`/`Catch`.

La **seule** règle structurante : **interdiction de mélanger les deux dans une même procédure**.

```vb
' VB.NET — ❌ ERREUR DE COMPILATION : les deux modèles dans une même procédure
Sub Traiter()
    On Error GoTo Gestion         ' modèle "On Error"…
    Try                           ' …ET modèle "Try" → refusé dans la même procédure
        Etape()
    Catch ex As Exception
    End Try
    Exit Sub
Gestion:
End Sub
```

```vb
' VB.NET — ✅ OK : deux procédures, deux modèles (cohabitation au niveau application)
Sub AnciennePartie()
    On Error GoTo Gestion          ' encore en style VB6 — compile et fonctionne
    Etape()
    Exit Sub
Gestion:
    Journaliser(Err.Description)
End Sub

Sub NouvellePartie()
    Try                            ' déjà migrée
        Etape()
    Catch ex As Exception
        Journaliser(ex.Message)
    End Try
End Sub
```

> 💡 **Conséquence directe** : l'**unité de migration est la procédure**. On la convertit **entièrement** (de `On Error` vers `Try`/`Catch`), puis on passe à la suivante — sans jamais devoir tout faire d'un coup.

---

## 2. Pourquoi les deux modèles s'entendent à la frontière

C'est le point qui rend la cohabitation **sûre**. En VB.NET, `On Error` est **bâti sur le même socle** que `Try`/`Catch` : les **exceptions**. Une erreur est *toujours*, au fond, une exception — qu'elle ait surgi via `Throw` ou via `Err.Raise`.

Conséquences aux **points de contact** entre une procédure « ancienne » et une « nouvelle » :

```vb
' Procédure encore en On Error qui appelle du code levant une exception
Sub Appelante()
    On Error GoTo Gestion
    Appelee()                     ' lève une exception (Throw) en interne
    Exit Sub
Gestion:
    ' le On Error GoTo ATTRAPE bien l'exception ; Err est renseigné à partir d'elle
    Journaliser(Err.Number & " - " & Err.Description)
End Sub
```

```vb
' Procédure en Try/Catch qui appelle du code VB6 faisant Err.Raise
Sub Appelante2()
    Try
        AppeleeVB6()              ' fait Err.Raise vbObjectError + N en interne
    Catch ex As Exception
        ' le Catch ATTRAPE bien : Err.Raise lève une vraie exception
        Journaliser(ex.Message)
    End Try
End Sub
```

Les erreurs **traversent la frontière dans les deux sens** : un `On Error GoTo` attrape une exception venue d'en bas ; un `Catch` attrape un `Err.Raise`. Aucune erreur ne se « perd » au passage.

> ⚠️ **Nuance** : la **discrimination fine** peut devenir **approximative** à la frontière. Un `Err.Raise vbObjectError + N` capté par un `Catch` ne porte pas forcément un **type** dédié ; une exception .NET captée par `On Error` ne donne pas toujours le **numéro** qu'aurait produit VB6. La passerelle **fonctionne**, mais distinguer précisément un cas particulier **à travers** les deux modèles est moins net — une raison de plus de **terminer** la migration plutôt que de s'appuyer durablement sur la cohabitation.

---

## 3. La démarche par phases

### Phase 0 — Base mécanique

La conversion automatique (l'**[*Upgrade Wizard*](../04-outils-migration/01-upgrade-wizard.md)**) amène le code en VB.NET en **conservant** `On Error`/`Err` : tout **compile** et **tourne**, mais rien n'est encore idiomatique. C'est la **ligne de départ**.

### Phase 1 — Établir le filet (avant de refactorer)

Avant de toucher à la gestion d'erreur, **capturez le comportement actuel** avec un **[harnais de tests de référence (golden master)](../05-preparer-code-vb6/05-golden-master.md)** et des **[tests de non-régression](../17-valider-refactoriser/03-non-regression.md)** — y compris le comportement des erreurs aujourd'hui **avalées** (cf. **[§ 11.3](03-on-error-resume-next.md)**). Sans ce filet, impossible de savoir si une conversion change quelque chose.

### Phase 2 — Convertir, procédure par procédure

Choisissez une procédure, convertissez **tout** son `On Error`/`Resume`/`Err` vers `Try`/`Catch`/`Finally`/`Exception` (mécanique des **[§ 11.1](01-on-error.md)** à **[§ 11.4](04-err-vs-exception.md)**), **validez** contre le golden master, puis passez à la suivante. Chaque procédure migrée est un incrément **testé**.

### Phase 3 — Éliminer les résidus

Retirez les **derniers** `On Error`, ainsi que les béquilles de compatibilité, pour **converger** vers du `Try`/`Catch` idiomatique partout. Ce nettoyage rejoint l'activation progressive d'**[`Option Strict`](../17-valider-refactoriser/01-option-strict.md)** et la **[suppression des dépendances `Compatibility`](../17-valider-refactoriser/02-supprimer-compatibility.md)**.

---

## 4. Par quoi commencer (priorisation)

| Priorité | Cibles | Pourquoi |
|---|---|---|
| **1** | Procédures à `On Error Resume Next` | La **dette** la plus dangereuse (**[§ 11.3](03-on-error-resume-next.md)**) — à traiter avec soin |
| **2** | Chemins à **fort enjeu** ou **fort risque** | Le gain de fiabilité y est maximal |
| **3** | Procédures **déjà ouvertes** pour d'autres raisons | Conversion **opportuniste**, à moindre coût |
| **4** | `On Error GoTo` **stables et simples** | Ils fonctionnent ; on les garde **pour la fin** |

> 💡 Inutile de tout convertir au même rythme : on **cible** d'abord ce qui est risqué ou fréquemment touché, et l'on laisse le code stable en `On Error` le temps qu'il faut.

---

## 5. Ce qu'il faut éviter

- **Mélanger** `On Error` et `Try` dans une même procédure — de toute façon **impossible** (le compilateur l'interdit).
- **Laisser `On Error Resume Next`** « parce que ça compile » : c'est figer la dette du **[§ 11.3](03-on-error-resume-next.md)**.
- **Laisser le temporaire devenir permanent** : la cohabitation est une **passerelle**, pas une destination. Fixez un objectif de **`Try`/`Catch` partout**.

---

## 6. Liens avec la stratégie d'ensemble

Cette migration progressive de la gestion d'erreur est une **déclinaison** de la stratégie incrémentale de la formation :

- **[§ 3.4 — Stratégies : *big-bang* vs incrémentale](../03-evaluer-strategie/04-strategies.md)** : le cadre général dont ce chapitre est un cas d'application.
- **[§ 17.1 / 17.2 / 17.3](../17-valider-refactoriser/01-option-strict.md)** : la phase de **convergence et de nettoyage** (Option Strict, suppression des béquilles, non-régression).
- **[§ 5.5 — Golden master](../05-preparer-code-vb6/05-golden-master.md)** : le **filet** indispensable à toute conversion sûre.

---

## 📋 Tableau de synthèse

| Phase | Objectif | Appui |
|---|---|---|
| 0 — Base mécanique | Compiler en VB.NET, `On Error` **intact** | Upgrade Wizard (§ 4.1) |
| 1 — Filet | **Capturer** le comportement actuel | Golden master (§ 5.5), non-régression (§ 17.3) |
| 2 — Conversion | **Procédure par procédure** → `Try`/`Catch` | § 11.1 à § 11.4 ; valider à chaque étape |
| 3 — Convergence | **Supprimer** les résidus `On Error` | § 17.1 (Option Strict), § 17.2 (Compatibility) |

| Règle | Détail |
|---|---|
| Cohabitation | Possible **entre** procédures, **pas dans** une même procédure |
| Frontière | Sûre (socle commun : exceptions) ; discrimination fine **approximative** |
| Unité de migration | **La procédure** (tout `On Error` **ou** tout `Try`) |
| Durée | **Temporaire** — viser `Try`/`Catch` partout |

---

## ✅ À retenir

1. **`On Error` survit en VB.NET** : on peut migrer **progressivement**, sans big-bang.
2. **On ne mélange pas les deux modèles dans une même procédure** → l'**unité de migration est la procédure**.
3. **La frontière est sûre** : `On Error` repose sur les exceptions, donc `Throw`/`Err.Raise` traversent les deux mondes — mais la **discrimination fine** y est approximative.
4. **Procédez par phases** : base mécanique → **golden master** → conversion **procédure par procédure** (validée) → élimination des résidus.
5. **Priorisez `On Error Resume Next`** et les chemins à risque ; ne laissez **pas** la cohabitation s'éterniser.

---

## 🏁 Fin du module 11

Vous avez parcouru toute la **gestion des erreurs** : les mécanismes VB6 et leur logique de remplacement (§ 11.1), le modèle cible `Try`/`Catch`/`Finally` et les filtres `When` (§ 11.2), le cas critique de `On Error Resume Next` (§ 11.3), la correspondance `Err` ↔ `Exception` et `Err.Raise` → `Throw` (§ 11.4), et la **stratégie progressive** (§ 11.5).

*Chapitre suivant : **[12 — Programmation orientée objet](../12-poo/README.md)** ⭐ ⚠️* — où l'on retrouve, entre autres, la **finalisation déterministe** (`Class_Terminate` → `IDisposable`/`Using`), prolongement naturel du `Finally` vu ici.

**Cible** : .NET Framework 4.7.2 · **Indicateurs** : ⭐ point clé · ⚠️ piège · 🛠️ outils · 🔗 interop

⏭️ [Programmation orientée objet](/12-poo/README.md)
