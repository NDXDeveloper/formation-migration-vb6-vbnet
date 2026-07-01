🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 🧱 Chapitre 12 — Programmation orientée objet (POO)
**Du modèle objet COM de VB6 au modèle objet CLR de .NET — migrer les classes sans casser la durée de vie des objets**

> **Partie 3 — Migrer le langage (le cœur de la formation)**
> Indicateurs : ⭐ chapitre central · ⚠️ pièges fréquents

---

## 🧭 De quoi parle ce chapitre

Souvenez-vous des **trois changements simultanés** de toute migration VB6 → VB.NET : le **langage**, le **runtime** et le **modèle objet**. Jusqu'ici (types, opérateurs, procédures, erreurs), c'est surtout le *langage* qui a changé. **Dans ce chapitre, les trois convergent** — et c'est le *modèle objet*, porté par le passage du **compteur de références** de VB6 au **garbage collector** de .NET, qui devient le sujet dominant.

C'est probablement le chapitre **le plus structurant** de la formation, et l'un des plus risqués. Pas parce que la syntaxe des classes serait difficile à traduire — elle ne l'est pas — mais parce que la **durée de vie des objets** n'obéit plus aux mêmes règles.

> ⚠️ Rappel du fil rouge : *le vrai danger n'est pas le code qui ne compile pas, c'est le code qui compile mais ne se comporte plus pareil.* La POO en est l'illustration parfaite : une classe migrée **compile et s'exécute**, mais si son `Class_Terminate` ne se déclenche plus au bon moment, les fichiers restent ouverts, les connexions traînent et les *handles* fuient — **silencieusement**.

---

## 🔀 Le grand écart : deux modèles objet

VB6 repose sur un modèle objet hérité de **COM** : simple, mais limité. VB.NET expose le modèle objet complet du **CLR** : riche, mais avec des règles de vie d'objet radicalement différentes.

| Aspect | VB6 (modèle COM) | VB.NET (modèle CLR) |
|--------|------------------|---------------------|
| Fichier de classe | `.cls` (une classe par fichier) | `.vb` (plusieurs types possibles par fichier) |
| Création | `Class_Initialize` | constructeur `Sub New()` |
| Destruction | `Class_Terminate`, **déterministe** (compteur de réf. → 0) | `Finalize`, **non déterministe** (le GC décide *quand*) |
| Nettoyage maîtrisé | implicite, dans `Class_Terminate` | explicite : `IDisposable` + `Using` |
| Héritage | **absent** (seulement `Implements`) | **complet** (`Inherits`, redéfinition, classes abstraites) |
| Affectation d'objet | `Set obj = New Classe` | `obj = New Classe` (`Set` **disparu**) |
| Propriétés | `Get` / `Let` (valeur) / `Set` (objet) | propriété unifiée ; `Let` **et** `Set` fusionnés |
| Types utilisateur | `Type…End Type` | `Structure…End Structure` |
| Code procédural global | module `.bas` | `Module` |
| Portée intermédiaire | `Friend` = **projet** | `Friend` = **assembly** |

À retenir : presque chaque ligne de ce tableau est soit une **simplification bienvenue**, soit un **piège silencieux**. Tout le travail de ce chapitre consiste à savoir reconnaître lesquelles.

---

## 🎯 Le piège n°1 de toute la migration : la finalisation

Si vous ne deviez retenir qu'une seule chose de cette formation entière, ce serait celle-ci (déjà annoncée au **chapitre 2.2**, et reprise en profondeur ici).

En **VB6**, un objet est détruit **immédiatement** quand sa dernière référence disparaît : le compteur de références tombe à zéro et `Class_Terminate` s'exécute **tout de suite**, de façon **prévisible**. D'innombrables programmes VB6 s'appuient sur ce comportement pour libérer leurs ressources — fermer un fichier, valider une transaction, relâcher un objet COM — pile au bon moment.

En **.NET**, il n'existe **aucune destruction déterministe**. Le *garbage collector* récupère la mémoire **quand il le décide**, parfois longtemps après que l'objet est devenu inutile. Transposer naïvement `Class_Terminate` vers `Finalize` revient donc à **déplacer le nettoyage dans le temps** — un changement de comportement invisible à la compilation et redoutable à l'exécution.

La réponse idiomatique de .NET est le **patron `IDisposable` / `Using`** : rendre le nettoyage **explicite et déterministe**, à la demande. Les sections **12.2** et **12.3** y sont entièrement consacrées.

> ⭐ Règle d'or de ce chapitre : **tout ce qui était nettoyé dans `Class_Terminate` doit être audité**, puis migré vers `Dispose` ou un bloc `Using` — **et non** vers `Finalize` seul.

---

## 🧩 Les autres bascules à connaître

Au-delà de la finalisation, ce chapitre traite plusieurs changements, dont certains sont de **vrais faux-amis** :

- **Propriétés** — VB6 distingue `Property Let` (affecter une valeur) et `Property Set` (affecter un objet). VB.NET les **fusionne** en un unique accesseur `Set`. Le mot `Set` existe donc dans les deux langages… mais ne veut **plus dire la même chose** (à ne pas confondre avec le `Set` d'affectation, lui, purement et simplement supprimé). Voir aussi le chapitre **10.6**.
- **Portée `Friend`** — accessible dans tout le **projet** en VB6, dans tout l'**assembly** en .NET. La correspondance est souvent directe, mais la sémantique diffère dès qu'un découpage en plusieurs projets entre en jeu. ⚠️
- **`Type` → `Structure`** — la **sémantique de valeur** est conservée, mais une `Structure` .NET se comporte différemment lors des affectations, des passages de paramètres et du stockage en collection. À migrer avec attention. ⚠️
- **Modules** — les modules standards `.bas` deviennent des `Module`. Le concept survit, mais la **portée des variables globales** et leur initialisation méritent un regard neuf.
- **Événements** — `Event` / `RaiseEvent` / `WithEvents` survivent, complétés par `Handles`, `AddHandler` et `RemoveHandler`. La migration est plutôt douce, mais un abonnement à un événement peut **maintenir un objet en vie** — un point qui rejoint, encore une fois, la question du GC.

---

## ⚖️ Une tentation à garder sous contrôle : l'héritage

VB6 **n'a pas d'héritage d'implémentation** : on ne peut que *contractualiser* un comportement via `Implements`. VB.NET, lui, offre l'**héritage complet**. C'est une vraie richesse… et un **piège de projet**.

> ⚠️ Migration ≠ réécriture. La disponibilité soudaine de l'héritage donne envie de **redessiner** les hiérarchies de classes. Résistez. L'objectif de cette phase est de **préserver le comportement existant**. La modélisation idiomatique (héritage, interfaces, génériques) viendra **après**, lors du refactoring (**chapitre 17**), une fois la non-régression assurée.

---

## 🗺️ Plan du chapitre

Le chapitre suit la logique « **d'abord la structure, puis la durée de vie, puis l'organisation, enfin les interactions** » :

1. **12.1 — [Classes (`.cls`) → classes `.vb`](01-classes.md)** : champs, propriétés et méthodes — la traduction de base d'une classe.
2. **12.2 — [`Class_Initialize` → constructeur ; `Class_Terminate` → le problème de la finalisation](02-initialize-terminate.md)** ⭐ ⚠️ : le cœur du risque.
3. **12.3 — [Finalisation déterministe perdue → `IDisposable` / `Using` / `Finalize`](03-idisposable-using.md)** ⭐ : le patron *Dispose*, la bonne réponse.
4. **12.4 — [Modules standards (`.bas`) → `Module`](04-modules-globales.md)** : variables globales et portée.
5. **12.5 — [Interfaces : `Implements` et héritage](05-interfaces-heritage.md)** : ce qui change pour `Implements`, et la nouveauté de l'héritage.
6. **12.6 — [`Public` / `Private` / `Friend`](06-portee.md)** ⚠️ : `Friend` passe à la portée **assembly**.
7. **12.7 — [`Type…End Type` → `Structure`](07-type-vers-structure.md)** ⚠️ : sémantique de valeur et pièges de comportement.
8. **12.8 — [Événements](08-evenements.md)** : `Event` / `RaiseEvent` / `WithEvents` / `Handles`, et `AddHandler` / `RemoveHandler`.

---

## ✅ Messages clés à emporter

- Le passage du **compteur de références** (COM) au **garbage collector** (CLR) est **le** changement central de ce chapitre.
- La **finalisation n'est plus déterministe** : `Class_Terminate` ne se migre **pas** tel quel — il s'audite et se transforme en `Dispose` / `Using`.
- Plusieurs mots-clés sont des **faux-amis** (`Set`, `Friend`) : même orthographe, sémantique différente.
- `Type` devient `Structure` : la valeur reste, **le comportement peut surprendre**.
- L'**héritage** est un cadeau de .NET — à **n'utiliser qu'après** la migration, pas pendant.

---

## 🔗 À relire en parallèle

- **Chapitre 2.2** — *La finalisation déterministe perdue* : la première présentation du piège, en version « cadrage ».
- **Chapitre 10.6** — *`Property Get`/`Let`/`Set`* : le détail du faux-ami `Set`.
- **Annexe B.3** — *`Class_Terminate` qui ne se déclenche plus* : symptôme, cause, correction.
- **Chapitre 17** — *Valider et refactoriser* : où l'on s'autorise (enfin) la modélisation idiomatique.

---

> Étape suivante, une fois les sections 12.1 à 12.8 parcourues : **[13. Des formulaires VB6 à Windows Forms](../13-formulaires-winforms/README.md)** — la migration de l'interface utilisateur.

⏭️ [Classes (`.cls`) → classes `.vb` : champs, propriétés, méthodes](/12-poo/01-classes.md)
