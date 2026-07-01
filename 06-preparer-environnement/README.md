🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6. Préparer l'environnement .NET

> *Partie 2 — Préparer le terrain (avant de migrer)*

**Avant de déplacer la moindre ligne de code, on aménage la zone d'atterrissage.** Le module 5 a préparé la **source** — le code VB6. Ce module prépare la **destination** — l'environnement .NET dans lequel ce code va venir vivre. Les deux forment le « terrain » sur lequel la migration se déroulera.

On ne migre pas dans un projet vide ouvert à la hâte : on installe un **environnement délibérément configuré**, où les décisions de fond (framework cible, structure des projets, interop COM, règles de compilation, gestion de versions) sont **prises à l'avance** plutôt qu'improvisées au fil de l'eau.

---

## Où en sommes-nous ?

```
Partie 2 — Préparer le terrain
   ┌───────────────────────────┬───────────────────────────────────┐
   │  Module 5                 │  Module 6   ◄── vous êtes ici     │
   │  Préparer la SOURCE       │  Préparer la DESTINATION          │
   │  (le code VB6)            │  (l'environnement .NET)           │
   └───────────────────────────┴───────────────────────────────────┘
                              │
                              ▼
   Partie 3 — Migrer le langage (le cœur de la formation)
```

À la fin du module 5, le code VB6 est nettoyé, ses comportements critiques sont documentés, le *golden master* est capturé et le périmètre est gelé. **Tout est prêt côté source.** Reste à préparer le côté cible — c'est l'objet de ce module — avant d'attaquer la migration du langage proprement dite en Partie 3.

---

## Pourquoi préparer — plutôt que foncer ?

Parce que les décisions prises **ici** conditionnent **toute** la migration :

- **Le framework cible** (.NET Framework 4.7.2, le « pont ») détermine la surface d'API disponible et le nombre de changements à absorber d'un coup.
- **La structure des projets** décide comment des centaines de `.frm`, `.bas` et `.cls` vont s'organiser en fichiers `.vb` — un choix coûteux à reprendre une fois la migration lancée.
- **Les options de compilation** (`Option Strict` en tête) fixent la sévérité du compilateur : trop strict trop tôt, et le code hérité ne compile plus ; pas assez, et on perd un filet précieux.
- **La gestion de versions** offre le point de retour en arrière et le lieu où vivent la référence *golden master* et la baseline gelée.

> ⚠️ Mal configurer l'environnement au départ, c'est se condamner à des reprises pénibles **au pire moment** : en pleine migration, quand on a déjà mille autres choses à gérer. Une heure de préparation ici en économise dix plus tard.

L'environnement n'est donc pas un simple décor : c'est, lui aussi, **une partie du filet de sécurité** (versionnement, jalons, compilation explicite), complémentaire du golden master du § 5.5.

---

## Le fil conducteur du module

Un principe traverse les cinq sections : **un atterrissage pragmatique et progressif.**

- **Pragmatique** : on configure l'environnement pour accueillir l'existant **tel qu'il est**, pas l'application idéale qu'on aimerait avoir. Exemple emblématique : on démarre en `Option Strict Off` pour tolérer le *late binding* hérité, quitte à le resserrer ensuite.
- **Progressif** : la cible 4.7.2 minimise les changements simultanés ; la structure reflète d'abord le VB6 ; l'idiomatique et la rigueur maximale viennent **après** la migration (module 17), pas pendant.

C'est la même philosophie que pour le reste de la formation : **on ne cumule pas les sauts.** L'environnement est conçu pour faire **un seul** changement à la fois.

---

## Ce que vous aurez mis en place à la fin du module

- ✅ Une **solution Visual Studio** ciblant **.NET Framework 4.7.2** (le « pont »).
- ✅ Une **structure de projets** qui reflète l'organisation du code VB6 (`.frm`/`.bas`/`.cls` → fichiers `.vb`).
- ✅ Les **références aux composants COM/ActiveX** nécessaires, en place et fonctionnelles.
- ✅ Les **options de compilation** décidées et appliquées — dont `Option Strict Off` au départ, à resserrer progressivement.
- ✅ Un dépôt **Git** avec une **branche de migration** et des **jalons** — prêt à héberger la référence *golden master* et la baseline gelée.

Autrement dit : une **zone d'atterrissage propre et outillée**, dans laquelle la migration du langage (Partie 3) pourra commencer sans friction d'environnement.

---

## Au programme

| § | Section | En une phrase |
|---|---------|---------------|
| **6.1** | Visual Studio et le ciblage **.NET Framework 4.7.2** | Installer et configurer l'IDE, et choisir le bon framework cible — celui qui minimise les changements simultanés. |
| **6.2** | Structure de solution et de projets | Décider comment les `.frm`/`.bas`/`.cls` deviennent des fichiers `.vb`, et comment organiser solution et projets. |
| **6.3** | Référencer les composants COM/ActiveX 🔗 | Mettre en place les *wrappers* d'interopérabilité (RCW, *Primary Interop Assemblies*) pour réutiliser l'existant COM. |
| **6.4** | Les options de compilation | Régler `Option Strict` / `Explicit` / `Infer` / `Compare` — avec le choix clé de démarrer `Strict Off` puis de resserrer (module 17). |
| **6.5** | Stratégie de gestion de versions | Mettre en place Git, une branche de migration et des jalons, pour avancer sûrement et pouvoir revenir en arrière. |

---

## Prérequis

- **Le module 5 réalisé** : code VB6 nettoyé (§ 5.1), comportements critiques documentés (§ 5.4), *golden master* capturé (§ 5.5), périmètre gelé (§ 5.6).
- **L'inventaire (§ 3.1) et la cartographie des dépendances (§ 3.2)** : on a besoin de la liste des OCX, DLL et références COM pour configurer correctement l'environnement (notamment en 6.3).
- **La stratégie choisie (§ 3.4)** : *big-bang* ou incrémentale — elle oriente la structure des projets (6.2) et la cohabitation éventuelle VB6/.NET.
- **Visual Studio installé** (le choix de version et le ciblage sont traités en 6.1).

---

## Points de vigilance

> ⚠️ **Cibler 4.7.2, pas .NET moderne.** Le réflexe « tant qu'à migrer, autant viser le dernier .NET » multiplie les changements simultanés. Le saut vers .NET 8/10 est une migration **distincte**, à faire **après** et seulement si nécessaire (rappel du § 1.2 ; passerelle au module 20).

> ⚠️ **`Option Strict Off` au départ est un compromis temporaire et assumé, pas un renoncement.** C'est ce qui permet au code hérité (riche en *late binding*) de compiler et de tourner pendant qu'on migre. On le resserre ensuite, module par module (§ 17.1).

> 💡 **La structure reflète d'abord le VB6 ; l'idiomatique vient après.** Réorganiser l'architecture *en même temps* qu'on migre, c'est rouvrir le périmètre qu'on vient de geler. On calque d'abord, on modernise plus tard (§ 17.5).

---

## À retenir

> - Ce module prépare la **destination** (.NET), comme le module 5 préparait la **source** (VB6) : ensemble, ils constituent la Partie 2 « préparer le terrain ».
> - Les décisions prises ici — **framework cible, structure, interop COM, options de compilation, gestion de versions** — façonnent toute la migration ; les improviser coûte cher plus tard.
> - Le mot d'ordre est **pragmatique et progressif** : accueillir l'existant tel quel, faire **un seul** changement à la fois, resserrer la rigueur ensuite.
> - On en sort avec une **zone d'atterrissage propre et outillée**, prête pour la migration du langage (Partie 3).

---

## Pour situer ce module

- **§ 1.2** — Pourquoi cibler .NET Framework 4.7.2 : le « pont » naturel *(le fondement du ciblage)*
- **§ 3.1 / 3.2** — Inventaire et cartographie des dépendances *(les entrées nécessaires à la configuration)*
- **§ 3.4** — Stratégies : *big-bang* vs incrémentale *(oriente structure et cohabitation)*
- **§ 5.5 / 5.6** — *Golden master* et gel du périmètre *(ce que la branche de migration accueille)*
- **Module 17** — Valider, fiabiliser et refactoriser *(la suite de 6.4 : activer `Option Strict` progressivement)*
- **Module 20** — De 4.7.2 vers .NET moderne *(le second saut, optionnel et séparé)*

⏭️ [Visual Studio et le ciblage **.NET Framework 4.7.2**](/06-preparer-environnement/01-vs-ciblage-472.md)
