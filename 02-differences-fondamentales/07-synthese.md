🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2.7 Tableau de synthèse des incompatibilités majeures

> **Chapitre 2 — Différences fondamentales VB6 ↔ VB.NET**
> La page récapitulative du chapitre : **l'aide-mémoire** à garder ouvert pendant toute la migration.

---

## Comment utiliser cette page

Cette section ne présente **pas** de nouvelle notion : elle **rassemble** en tableaux ce que le  
chapitre 2 a expliqué, et **renvoie** chaque sujet vers son traitement détaillé (parties 3 à 5 et  
annexes). C'est la page à **garder sous la main** : un coup d'œil pour savoir *ce qui change*, *à  
quel point c'est dangereux*, et *où trouver le détail*.

Pour rester complète en tant que synthèse des incompatibilités **majeures**, elle déborde  
légèrement du chapitre 2 : quelques différences traitées plus loin (tableaux 0-based, structures de  
contrôle supprimées…) y figurent **avec leur renvoi**, afin d'offrir une vue d'ensemble d'un seul  
tenant.

### La clé de lecture : 🟢 visible vs 🔴 silencieux

Tout le chapitre 2 repose sur cette distinction — elle structure aussi ces tableaux. Dans la colonne
**Détection** :

- 🟢 **Visible** — le **compilateur** vous arrête (mot-clé disparu, type incompatible…). Pénible, mais **sûr** : impossible de passer à côté.
- 🔴 **Silencieux** — **rien** ne vous alerte : le code **compile** et **s'exécute**, mais le **comportement change**. C'est **là** qu'est le danger, et c'est ce que cette page sert avant tout à rendre visible.

> ⚠️ **Lecture stratégique.** Parcourez d'abord toutes les lignes **🔴** : ce sont vos priorités de
> validation et de test. Les lignes 🟢, le compilateur s'en charge.

---

## 🏛️ Runtime et mémoire (les ruptures fondatrices)

| Sujet | VB6 | VB.NET | Détection | Détail |
|---|---|---|---|---|
| **Moteur d'exécution** | VBVM (`msvbvm60.dll`, COM) | CLR (managé, IL + JIT) | 🔴 invisible en source | 2.1 |
| **Gestion mémoire** | Comptage de références | **Ramasse-miettes (GC)** | 🔴 | 2.1 |
| **Finalisation** | `Class_Terminate` **déterministe** | `Finalize` non déterministe → **`IDisposable`/`Using`** | 🔴 **piège n°1** | 2.2 · 12.3 · **B.3** |
| **Bits / threads** | 32 bits, STA mono-thread | 32/64 bits, multithread | 🔴 (interop, binaire) | 2.1 · 16.5 |

---

## 🔢 Types et conversions

| Type VB6 | Devient | Point de vigilance | Détection | Détail |
|---|---|---|---|---|
| **`Variant`** | `Object` | Boxing, coercitions, états perdus | 🟢 mot-clé invalide | 2.3 · 7.1 |
| **`Integer`** (16 b) | `Short` *(pour garder 16 b)* | API, binaire, débordements, bits | 🔴 **même mot-clé !** | 2.3 · 7.2 · **B.2** |
| **`Long`** (32 b) | `Integer` *(pour garder 32 b)* | API (`Declare`), binaire, débordements | 🔴 **même mot-clé !** | 2.3 · 7.2 · **B.2** |
| **`Currency`** | `Decimal` *(jamais `Single`/`Double`)* | Arrondis monétaires | 🟢 mot-clé invalide | 2.3 · 7.2 · **B.9** |
| **`Date`** | `DateTime` (était un `Double`) | Arithmétique de dates, arrondis | 🟢 + 🔴 | 7.3 · **B.9** |
| **`Boolean`** | `Boolean` (`True = -1` en VB6) | Conversions booléen ↔ entier | 🔴 | 7.6 · **B.8** |
| **`String * n`** (longueur fixe) | `VBFixedString` ou refactoring | Sémantique de remplissage | 🟢 | 7.4 |
| **`Empty` / `Null` / `Missing`** | États du `Variant` **disparus** | Tests d'état à réécrire | 🔴 | 7.7 · **B.5** |
| Conversions `CDbl`/`CDate`/`Format` | Idem, mais **sensibles à la culture** | Séparateur décimal, format de date régionaux | 🔴 | **B.10** |

---

## 🧩 Procédures, paramètres et appels

| Sujet | VB6 | VB.NET | Détection | Détail |
|---|---|---|---|---|
| **Passage par défaut** | **`ByRef`** | **`ByVal`** | 🔴 **le mot non écrit** | 2.4 · 10.2 · **B.1** |
| **`ParamArray`** | passé `ByRef` | tableau passé `ByVal` | 🔴 | 10.4 |
| **`Optional`** | `IsMissing` détecte l'absence | **valeur par défaut obligatoire** | 🟢 + 🔴 | 10.3 |
| **Parenthèses d'argument** `((x))` | force `ByVal` localement | sémantique différente | 🔴 (lecture du code VB6) | 10.1 |
| **Valeur de retour** | `NomFonction = valeur` | `Return` (ou les deux) | 🟢 | 10.5 |

---

## 🧱 Programmation orientée objet et cycle de vie

| Sujet | VB6 | VB.NET | Détection | Détail |
|---|---|---|---|---|
| **`Set` (affectation)** | `Set a = obj` | **supprimé** → `a = obj` | 🟢 erreur de compil. | 2.5 · 10.6 |
| **Propriétés par défaut** | implicites (`txt = "x"`) | **nommées** (`txt.Text`) — sauf paramétrées | 🔴 en **lecture** sous `Strict Off` | 2.5 · **B.7** |
| **`Property Let` + `Set`** | deux accesseurs distincts | **fusionnés** en un `Set` (faux-ami) | 🟢 | 10.6 |
| **`Class_Initialize`** | événement | **constructeur** `New()` | 🟢 | 12.2 |
| **Héritage de classe** | absent (`Implements` seul) | **présent** (nouveauté) | nouveauté | 12.5 |
| **`Type…End Type`** | type structuré | **`Structure`** (sémantique de valeur) | 🔴 comportement | 12.7 |
| **`Friend`** | portée **projet** | portée **assembly** | 🔴 si multi-assembly | 12.6 |
| **Événements** | `Event`/`RaiseEvent`/`WithEvents` | + `AddHandler`/`RemoveHandler` | 🟢 + 🔴 | 12.8 |

---

## 📦 Tableaux et structures de contrôle

| Sujet | VB6 | VB.NET | Détection | Détail |
|---|---|---|---|---|
| **Bornes de tableaux** | `Option Base`, bornes non nulles | **toujours 0-based** | 🔴 | 8.1 · **B.4** |
| **`Collection`** | objet `Collection` VB6 | `List(Of T)` / `Dictionary(Of K,V)` | 🟢 | 8.4 |
| **`GoSub` / `Return`, `On…GoTo`** | présents | **supprimés** → refactoring | 🟢 erreur de compil. | 9.5 |
| **`While…Wend`** | `While…Wend` | `While…End While` | 🟢 | 9.4 |
| **`&` vs `+`** (concaténation) | `+` concatène parfois | préférer `&`, `+` = addition | 🔴 (chaînes/nombres) | 9.1 |

---

## ⚠️ Gestion d'erreurs

| Sujet | VB6 | VB.NET | Détection | Détail |
|---|---|---|---|---|
| **Modèle** | `On Error` / `Err` / `Resume` | `Try`/`Catch`/`Finally` / `Exception` | 🟢 **`On Error` survit** | 2.6 · 11 |
| **`On Error Resume Next`** | ignore les erreurs | **dangereux à migrer** | 🔴 | 11.3 · **B.6** |
| **Reprise** (`Resume`/`Resume Next`) | reprend l'exécution | **pas d'équivalent** structuré | 🔴 | 11.1 |
| **Déclenchement** | `Err.Raise` | `Throw` | 🟢 | 11.4 |

---

## 🔴 Le palmarès des pièges silencieux

Si vous ne deviez retenir qu'une chose de ce chapitre, ce serait **cette liste** : les changements  
qui **compilent sans broncher** mais **modifient le comportement**. Chacun renvoie à sa fiche
**symptôme → cause → correction** dans l'**annexe B** (le catalogue des pièges, l'annexe la plus
importante de la formation).

| # | Le piège silencieux | Conséquence typique | Fiche |
|---|---|---|---|
| 1 | **Finalisation perdue** (`Class_Terminate` → GC) | Connexions épuisées, fichiers verrouillés, fuites | **B.3** |
| 2 | **ByRef → ByVal** par défaut | Valeurs « de retour » perdues, variables inchangées | **B.1** |
| 3 | **Entiers redimensionnés** (`Integer`/`Long`) | Appels d'API faussés, formats binaires décalés, débordements | **B.2** |
| 4 | **Tableaux 0-based** | Décalage d'indices, élément manquant ou en trop | **B.4** |
| 5 | **`Empty`/`Null`/`Missing`** perdus | Tests d'état du `Variant` qui changent de résultat | **B.5** |
| 6 | **`On Error Resume Next`** | Erreurs masquées, bugs avalés silencieusement | **B.6** |
| 7 | **Propriétés par défaut** (lecture) | Comportement différent sous `Option Strict Off` | **B.7** |
| 8 | **`True = -1`** / conversions booléennes | Valeurs entières inattendues | **B.8** |
| 9 | **Dates et arrondis monétaires** | Écarts au centime, comparaisons de dates faussées | **B.9** |
| 10 | **Conversions sensibles à la culture** | Séparateur décimal / format de date selon la machine | **B.10** |

> ⭐ **Le réflexe à garder.** Le compilateur attrape les lignes 🟢. **Vous** (aidé des tests de
> non-régression et de l'annexe B) êtes le seul filet pour les lignes 🔴. C'est exactement la raison
> d'être du **golden master** (5.5) : capturer le comportement VB6 **avant**, le comparer **après**,
> et faire ressortir ce qu'aucun compilateur ne signale.

---

## 🧭 Où trouver le détail (carte des renvois)

| Famille d'incompatibilités | Traitement détaillé |
|---|---|
| Runtime, GC, finalisation | Sections **2.1–2.2**, chapitre **12** (12.2–12.3) |
| Types et conversions | Chapitre **7**, annexe **D** (tailles et plages), annexe **B** |
| Tableaux et collections | Chapitre **8** |
| Opérateurs et structures de contrôle | Chapitre **9** |
| Procédures et paramètres | Chapitre **10** |
| Gestion d'erreurs | Chapitre **11** |
| POO, classes, événements | Chapitre **12** |
| **Tous les pièges silencieux** | **Annexe B** — *symptôme · cause · correction* |
| Correspondance mot-clé à mot-clé | **Annexe A** |

---

## ✅ À retenir

- Cette page est l'**aide-mémoire** du chapitre 2 : elle **synthétise** les incompatibilités majeures et **pointe** vers leur traitement détaillé. À garder ouverte pendant la migration.
- La clé de lecture est **🟢 visible / 🔴 silencieux** : le compilateur gère le visible ; **le silencieux est votre priorité** de validation.
- Les ruptures **fondatrices** sont dans le **runtime et la mémoire** (2.1–2.2) ; les pièges les plus mordants concernent les **types** (entiers, `Currency`, dates), les **appels** (ByRef/ByVal) et la **finalisation**.
- Le **palmarès des 10 pièges silencieux** condense le risque ; chacun a sa fiche en **annexe B**.
- Le compilateur ne voit pas le silencieux : c'est le rôle du **golden master** (5.5) et des **tests de non-régression** (17.3) de le révéler.

---

## 🎓 Fin du chapitre 2 — ce que vous savez maintenant

Vous avez posé le **socle conceptuel** de toute la migration. À ce stade, vous êtes en mesure de :

- **distinguer les deux runtimes** (VBVM/CLR) et comprendre que ce changement de moteur **invisible** conditionne tout le reste (2.1) ;
- **maîtriser le piège n°1** — la finalisation déterministe perdue — et sa parade `IDisposable`/`Using` (2.2) ;
- **repérer les pièges de typage** (entiers redimensionnés, `Variant`, `Currency`) avant qu'ils ne mordent (2.3) ;
- **auditer les appels** à la lumière du basculement **ByRef → ByVal** (2.4) ;
- **anticiper** les propriétés par défaut, le `Set` disparu et le changement de modèle d'erreurs (2.5–2.6) ;
- **disposer d'une vue de synthèse** des incompatibilités, et savoir **où** chercher le détail (2.7).

Surtout, vous avez intériorisé le **fil conducteur** : le danger n'est pas le code qui ne compile  
pas, c'est le code qui **compile mais ne se comporte plus pareil**. Tout le reste de la formation  
applique ce principe.

La **Partie 1** n'est pas tout à fait terminée : avant de préparer le terrain (Partie 2), il faut  
encore **évaluer** ce que l'on a et **choisir une stratégie**. C'est l'objet du **chapitre 3** —  
inventaire, cartographie des dépendances, estimation de l'effort, et l'arbitrage *big-bang* vs  
incrémentale.

---

> 🧭 **Chapitre suivant** → [3. Évaluer l'existant et choisir une stratégie](../03-evaluer-strategie/README.md) ⭐  
> 🔝 **Retour** → [Sommaire du chapitre 2](README.md)

⏭️ [Évaluer l'existant et choisir une stratégie](/03-evaluer-strategie/README.md)
