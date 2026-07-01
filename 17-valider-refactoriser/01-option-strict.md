🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.1 — Activer `Option Strict` progressivement (éliminer le *late binding* résiduel) ⭐

**Faire passer tout le code migré en mode strict — non pas d'un coup, mais fichier par fichier — pour éliminer le liage tardif résiduel et les conversions implicites, sans jamais casser la compilation ni perdre le fil.**

> ⭐ Le paiement de deux promesses : celle du module 6.4 (« démarrer en `Option Strict Off`, puis
> l'activer **progressivement** ») et celle du module 16.3 (éliminer le *late binding* résiduel, piloté
> **globalement ici**). Le module 16.3 a posé le **concept** (liage anticipé vs tardif) et le **patron
> d'isolation** ; cette section traite la **méthode d'activation à l'échelle de tout le projet**.

---

## Ce qu'`Option Strict On` impose (bref rappel)

Activer `Option Strict On` interdit **deux** familles de constructions tolérées par le mode `Off`
(détail complet au module 16.3) :

1. **Le liage tardif** — tout appel de membre sur un `Object` (résolu à l'exécution) devient une
   **erreur de compilation**.
2. **Les conversions implicites avec perte** — un `Long` glissé dans un `Integer`, un `Double` dans un
   `Single`… exigent désormais une conversion **explicite**.

> 🧨 **C'est ce double effet qui fait toute la valeur de l'opération** : on n'élimine pas seulement le
> liage tardif, on force aussi à **regarder en face** chaque conversion douteuse — exactement le
> terrain des entiers redimensionnés du module 7.2. Activer le mode strict, c'est faire **remonter à la
> compilation** une foule de défauts qui, sinon, resteraient silencieux jusqu'en production.

---

## Pourquoi *progressivement*, et pas d'un coup

Un code fraîchement porté en `Option Strict Off` peut receler **des centaines, voire des milliers** de  
sites de liage tardif et de conversions implicites. Basculer le réglage projet sur `On` **d'un seul  
geste** produit un **mur d'erreurs** : impossible à trier, démoralisant, et — surtout — l'application
**ne compile plus** tant que tout n'est pas corrigé. On perd la capacité de livrer entre-temps.

L'activation **progressive** offre l'inverse :

- le code **compile en permanence** ;
- on traite **une unité à la fois** ;
- on **revalide** après chaque lot contre le *golden master* (le filet de sécurité du module 17.3) ;
- on mesure une progression **tangible**.

C'est la déclinaison, au niveau de la fiabilisation, de la philosophie **incrémentale** de toute la  
formation (module 3.4).

---

## Le levier : niveau fichier vs niveau projet

La clé technique : **`Option Strict On` placé en tête d'un fichier prime sur le réglage du projet**
(rappel du module 16.3). C'est ce qui rend l'activation fichier par fichier possible.

La stratégie type :

1. Laisser le **réglage projet sur `Off`** (base tolérante) pendant la transition.
2. Ajouter **`Option Strict On`** en tête des fichiers, **un par un**, à mesure qu'on les nettoie.
   Chaque fichier terminé est ainsi **verrouillé** en mode strict.
3. Quand **tous** les fichiers portent la directive, basculer le **réglage projet sur `On`** et
   **retirer** les directives par fichier devenues redondantes — ce qui impose aussi le mode strict à
   **tout code futur**.

---

## Le gradient fin : les avertissements comme **liste de travail**

Complément précieux du levier fichier par fichier : VB.NET permet de configurer **individuellement** les  
conditions d'`Option Strict` comme **avertissements** plutôt qu'erreurs (propriétés du projet →
*Compilation* → *Configurations d'avertissement* ; l'état résultant est `Option Strict Custom`).

L'astuce : en gardant le projet sur `Off`, passez les conditions **« Liaison tardive »** et
**« Conversion implicite »** sur **Avertissement**. Vous obtenez alors, dès la compilation, la **liste
complète** des sites à traiter — **sans casser le build**. Le **décompte des avertissements** devient un  
backlog mesurable, que l'on fait fondre vers zéro avant de promouvoir ces conditions en **erreurs**.

> 💡 **Deux leviers complémentaires** : le **niveau fichier** (`On`/`Off`) verrouille un fichier propre ;
> le **niveau projet** (avertissements) **révèle l'ampleur** du chantier et sert de fil conducteur.

---

## La combinaison recommandée : `Strict On` + `Infer On` + `Explicit On`

Activez `Option Strict On` **avec** `Option Infer On`, sinon une déclaration comme `Dim x = Expr()`  
risque de retomber en `Object` (donc en liage tardif). Avec l'inférence active, `x` prend le **type  
statique** de l'expression. Le trio idiomatique en tête de fichier :

```vbnet
Option Strict On
Option Explicit On
Option Infer On
' Le fichier est désormais en mode strict, quel que soit le réglage du projet
```

---

## Méthode d'activation, étape par étape

Une boucle simple, à répéter unité après unité (ce n'est pas un exercice, mais le **mode opératoire**) :

1. **Établir la base** : projet sur `Off`, conditions « liaison tardive » et « conversion implicite »
   sur **Avertissement**. Noter le **nombre** d'avertissements (le backlog de départ).
2. **Choisir une unité** — un fichier, une classe, un petit projet —, de préférence **peu couplée et de
   petite taille** pour bâtir une base saine et prendre de l'élan.
3. **Ajouter `Option Strict On`** en tête, et **corriger** chaque erreur (voir le tri ci-dessous) :
   privilégier le **vrai type** (liage anticipé) au cast ; **expliciter ET vérifier** chaque conversion ;
   **isoler** tout liage tardif inévitable (patron du module 16.3).
4. **Revalider** : compiler, puis rejouer la comparaison **golden master** (module 17.3) sur la zone
   concernée. Résoudre tout écart.
5. **Committer** sur la branche de migration (Git, module 6.5), puis passer à l'unité suivante.
6. **Clore** : quand tous les fichiers sont stricts, basculer le **projet sur `On`**, retirer les
   directives redondantes, et promouvoir les avertissements restants en erreurs.

---

## Trier et corriger : les catégories d'erreurs

### Liage tardif → donner le vrai type (ou isoler)

> *« `Option Strict On` interdit les liaisons tardives. »*

C'est le **« late binding résiduel »** du titre. La correction préférée : **déclarer le type réel**
(ajouter la référence si besoin) pour repasser en liage anticipé.

```vbnet
' ❌ Liage tardif hérité du style « tout Variant »
Dim client As Object = ChargerClient(id)
Dim nom As String = client.Nom              ' résolu à l'exécution

' ✅ Type réel → liage anticipé, vérifié à la compilation
Dim client As Client = ChargerClient(id)
Dim nom As String = client.Nom
```

Si le liage tardif est **réellement nécessaire** (pas de bibliothèque de types, automatisation  
indépendante de la version…), ne relâchez pas le fichier entier : **isolez-le** derrière une interface  
stricte, dans une cellule dédiée (patron d'isolation du module 16.3).

### Conversion implicite → expliciter **ET vérifier** ⚠️

> *« `Option Strict On` interdit les conversions implicites de 'Long' en 'Integer'. »*

Ici se joue le **piège n°1** de toute la démarche. La tentation est de **faire taire le compilateur** en  
emballant tout dans un `CType`/`CInt` — ce qui **réintroduit** précisément les bugs silencieux que l'on  
cherche à éliminer (module 7.2) :

```vbnet
Dim total As Long = CalculerTotal()

' ❌ MAUVAIS : silencer le compilateur sans réfléchir → débordement silencieux possible
Dim compteur As Integer = CInt(total)       ' et si total > 2 147 483 647 ?

' ✅ BON : se demander si la conversion est correcte. Ici, garder Long EST la vraie réponse.
Dim compteur As Long = total
```

> ⚠️ **`Option Strict On` n'est pas là pour qu'on le contourne — il est là pour qu'on *réfléchisse* à
> chaque conversion.** Un cast mécanique transforme une erreur visible en bug invisible. Chaque
> conversion explicitée doit l'être en **connaissance de cause**.

### `DirectCast` / `CType` / `TryCast` : choisir le bon

Quand une conversion est légitime, le bon outil dépend de l'intention :

| Outil | Usage | Comportement en cas d'échec |
|-------|-------|-----------------------------|
| `DirectCast` | Le type **exact** est connu (transtypage de référence) | `InvalidCastException` (rapide, strict) |
| `CType` | **Vraie conversion** (avec opérateurs/conversions VB) | `InvalidCastException` (plus permissif) |
| `TryCast` | « **C'est peut-être** ce type » (références uniquement) | renvoie `Nothing` |

> ⚠️ Choisir le mauvais outil peut **masquer** ou **modifier** un comportement (un `CType` permissif là
> où un `DirectCast` strict était attendu, par exemple).

---

## Où se concentre le travail

Dans un code VB6 migré, le liage tardif et les conversions implicites se **regroupent** à des endroits  
prévisibles. Attendez-vous au gros du chantier sur :

- le code **« tout `Variant` »** (module 7.1) → quantité de variables `Object` ;
- l'**accès aux données** ADO : `Recordset.Fields(...).Value` renvoie `Object` (module 15.3), avec en
  prime la gestion de `DBNull` ;
- les cas d'**interop** légitimement en liage tardif (module 16.3), à **isoler** ;
- l'accès aux **contrôles par nom** (`Controls("…")` renvoyant un `Control` générique) ;
- les idiomes de **dispatch dynamique** (`CallByName`…).

Traitez les zones simples et peu couplées en premier ; abordez ces **points chauds** en dernier, avec  
soin et sous la protection du *golden master*.

---

## Le filet de sécurité à chaque lot

Rappel du **principe cardinal** du module (README du 17) : ces corrections **modifient le code**.  
Expliciter une conversion en **préserve** généralement le comportement — mais **corriger** une  
conversion erronée (un vrai bug) le **change**. Seule la comparaison **golden master** (module 17.3)  
permet de distinguer un correctif voulu d'une régression introduite. **Revalidez après chaque lot.**

---

## Mesurer l'avancement

Le **décompte des avertissements** (ou des fichiers encore en `Option Strict Off`) constitue un  
excellent **indicateur de progression** — une courbe qui descend vers zéro, à rapprocher des critères de  
réussite définis au module 3.6. Chaque fichier non encore strict est une **dette** identifiée et  
suivie, pas un oubli.

---

## ⚠️ Pièges à connaître

- **L'anti-pattern du cast mécanique** : emballer chaque conversion signalée dans un `CType`/`CInt`
  juste pour compiler **réintroduit** les bugs du module 7.2. À proscrire — réfléchir à chaque
  conversion. ⚠️
- **Oublier `Option Infer On`** : sans elle, des `Dim x = …` retombent en `Object` → encore du liage
  tardif.
- **Mauvais choix `DirectCast`/`CType`/`TryCast`** : peut masquer ou altérer un comportement. ⚠️
- **Basculer tout le projet d'un coup** : mur d'erreurs ingérable, perte de la capacité à livrer.
- **Ne pas revalider les conversions** contre le *golden master* : une conversion « corrigée » qui est
  en réalité un changement de comportement passe inaperçue.
- **Liage tardif nécessaire mais non isolé** : s'il est inévitable, **isolez-le** (module 16.3) — ne
  laissez pas le fichier entier en `Off`.
- **Fichiers `Off` oubliés** : chaque fichier sans `Option Strict On` est une dette à **tracer**.

---

## En résumé

- `Option Strict On` interdit **le liage tardif** *et* **les conversions implicites** : son activation
  fait remonter à la compilation une classe entière de défauts silencieux (le double bénéfice du module
  16.3, le second rejoignant le 7.2).
- On l'active **progressivement** : le code compile en permanence, on traite une unité à la fois, on
  revalide après chaque lot.
- **Deux leviers** : `Option Strict On` **au niveau fichier** (verrouiller un fichier propre) et les
  **avertissements au niveau projet** (`Option Strict Custom`, pour révéler tout le backlog). Combinaison
  idiomatique : **`Strict On` + `Infer On` + `Explicit On`**.
- **Tri des erreurs** : liage tardif → **donner le vrai type** (ou isoler) ; conversion implicite →
  **expliciter ET vérifier** (jamais un cast mécanique) ; choisir entre `DirectCast`/`CType`/`TryCast`
  selon l'intention.
- **Filet de sécurité** : revalider chaque lot contre le *golden master* (17.3). **Mesurer** la
  progression (décompte des avertissements → 0) comme indicateur (module 3.6).

> ➡️ **Autre dette de migration à solder : les appels à l'espace `Microsoft.VisualBasic.Compatibility`,
> la « béquille » du module 4.2, qu'il faut maintenant retirer pour ne pas figer la migration sur un
> échafaudage.** (Module 17.2)

---

🏷️ **Indicateurs** : ⭐ Point clé (fiabilisation) · lié aux modules 6.4 (options), 16.3 (liage), 7.2 (conversions), 17.3 (golden master), 3.6 (indicateurs)
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Supprimer les dépendances à `Microsoft.VisualBasic.Compatibility`](/17-valider-refactoriser/02-supprimer-compatibility.md)
