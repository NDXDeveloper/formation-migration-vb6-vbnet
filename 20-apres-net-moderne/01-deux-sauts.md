🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.1 Pourquoi 4.7.2 d'abord, .NET moderne ensuite : deux sauts séparés ⭐

> 🧭 *Module 20 — Et après ? De 4.7.2 vers .NET moderne · Section 1 / 5*

L'introduction du module a posé le principe : **deux sauts, pas un**. Cette section en développe le **pourquoi**. La réponse tient à un **seul principe d'ingénierie**, appliqué deux fois : *minimiser le nombre de choses qui changent en même temps*, parce que le vrai danger de cette migration n'est pas le code qui ne compile pas — c'est le code qui **compile mais ne se comporte plus pareil**, et que des changements **groupés** rendent ces régressions à la fois **plus nombreuses** et **impossibles à diagnostiquer**.

---

## Le point de départ : VB6 → VB.NET change déjà trois choses à la fois

Migrer de VB6 vers VB.NET n'est pas une conversion de syntaxe : c'est un changement de **trois choses simultanément** ([rappel 1.3](../01-cadrage-migration/03-ce-qui-change.md), [module 2](../02-differences-fondamentales/README.md)) :

1. **le langage** (mots-clés, structures, types) ;
2. **le runtime** ([VBVM → CLR](../02-differences-fondamentales/01-vbvm-vs-clr.md), avec un *garbage collector*) ;
3. **le modèle objet** ([finalisation](../02-differences-fondamentales/02-finalisation-deterministe.md), durée de vie, événements).

C'est **déjà beaucoup** de changement simultané. Toute la question du « et après ? » est de savoir s'il faut en **empiler davantage** d'un coup — ou non.

---

## Le principe : minimiser les changements simultanés

Pourquoi est-ce **la** règle ? Parce que les régressions de cette migration sont **silencieuses**, et que le nombre de changements groupés a trois effets délétères :

- **Chaque changement simultané ajoute une source indépendante** de régression silencieuse.
- **On ne peut plus attribuer une régression à sa cause.** Si dix choses ont changé d'un bloc et qu'un comportement dévie, lequel des dix est responsable ? Sans **isolation des variables**, le diagnostic devient ingérable.
- **Le risque ne s'additionne pas, il se multiplie.** Les changements **interagissent** entre eux, créant des défaillances qu'aucun des changements n'aurait produites seul.

> ⭐ **La règle d'or :** un saut = **un lot de changements isolé**, **validé** au *golden master*, **avant** d'envisager le suivant. C'est cette discipline qui transforme une migration terrifiante en étapes maîtrisables.

---

## Pourquoi 4.7.2 est le palier naturel (le « pont »)

Le génie de 4.7.2 est de **garder la plateforme au plus près de VB6** tout en changeant le langage, le runtime et le modèle objet ([rappel 1.2](../01-cadrage-migration/02-pourquoi-472.md)) :

- **surface d'API la plus proche de VB6** (Windows, COM, GDI+, impression) ;
- **meilleure compatibilité COM / ActiveX / OCX** et **Crystal Reports** ;
- **parité du concepteur Windows Forms** et des contrôles intrinsèques ;
- **runtime supporté**, adapté à du *legacy* d'entreprise.

Autrement dit, 4.7.2 est le **plus petit premier saut viable** : on change **ce qu'on doit** (le langage et son moteur), et on **garde stable** tout le reste — plateforme, interop, UI. C'est l'application directe du principe ci-dessus.

---

## Le second saut est un *autre* lot de changements

Passer de 4.7.2 au **.NET moderne (8/10)** n'est pas « une mise à jour » : c'est un **nouveau front** de changements, distinct du premier (détail en [20.2](02-evaluer-second-saut.md) et [20.3](03-ce-qui-change-encore.md)). De façon synthétique :

| | **Saut 1 — VB6 → 4.7.2** | **Saut 2 — 4.7.2 → .NET moderne** |
|---|---|---|
| **Langage** | VB6 → VB.NET | VB.NET (stable) |
| **Runtime** | VBVM → CLR (.NET Framework) | runtime .NET moderne **unifié / multiplateforme** |
| **Modèle objet** | comptage de références → GC, `IDisposable` | stable (ajustements) |
| **API / plateforme** | surface **proche de VB6** | **API retirées** ou modifiées |
| **Windows Forms** | concepteur **en parité** | **réimplémenté** sur .NET moderne |
| **Interop COM** | au plus près de VB6 | **différences** notables |
| **Espace `My`** | complet | **partiel** |
| **Outillage** | assistants de migration, béquilles disponibles | autre outillage |

Empiler le saut 1 **et** le saut 2 revient à changer **presque tout en même temps** : c'est exactement la **multiplication du risque** rendue concrète.

---

## Pourquoi l'*ordre* compte (4.7.2 d'abord)

Trois raisons rendent l'ordre **non interchangeable** :

1. **L'outillage et les béquilles de compatibilité sont au bon endroit.** Les assistants de migration, la **parité du concepteur WinForms**, l'espace [`Microsoft.VisualBasic.Compatibility`](../04-outils-migration/02-compatibility-namespace.md) comme **béquille temporaire**, et la mécanique des *Primary Interop Assemblies* COM vivent le mieux sur **.NET Framework**. Aller **directement** au .NET moderne **retire** ces appuis **précisément** quand on en a le plus besoin. Passer par 4.7.2 permet de **s'appuyer** dessus, puis de les **retirer** proprement (module 17).

2. **Une base validée sert de référence.** Atterrir sur 4.7.2 avec un *golden master* **au vert** ([module 5.5](../05-preparer-code-vb6/05-golden-master.md), [Annexe F](../annexes/golden-master/README.md)) fournit une **base connue-bonne**. Le second saut se mesure alors contre **cette référence stable**, et non contre la cible **mouvante** qu'était VB6. Sans ce **palier intermédiaire validé**, deux classes de bugs se mélangent **sans moyen de les séparer**.

3. **Deux migrations traitables au lieu d'une ingérable.** Le séquençage convertit **une** opération impossible à diagnostiquer en **deux** étapes maîtrisables, chacune avec **son propre point de contrôle** au *golden master*.

---

## Une image pour fixer l'idée

4.7.2 est le **pont** qui franchit la rivière, de VB6 au monde **managé** de .NET. **Traverser ce pont est un trajet complet et utile en soi.** Décider ensuite de gravir la **montagne d'en face** — le .NET moderne — est une **expédition séparée**, qu'on choisit une fois **arrivé sain et sauf** sur l'autre rive… et **jamais au milieu du gué**. Changer le moteur, le carburant **et** le châssis d'un seul coup, c'est se condamner à ne plus savoir **ce qui a cassé**.

---

## Le corollaire : « ensuite » peut vouloir dire « jamais »

« Deux sauts séparés » ne signifie **pas** « deux sauts obligatoires ». Une fois sur 4.7.2, le second saut est une **décision** fondée sur la **valeur** ([20.2](02-evaluer-second-saut.md)), et **rester** est un choix parfaitement **légitime** ([20.4](04-rester-sur-472.md)). La logique de **séparation** tient, que l'on prenne un jour le second saut **ou non** : elle protège la migration **dans tous les cas**.

---

## À retenir

**Un principe, appliqué deux fois : minimiser le changement simultané.** 4.7.2 **d'abord**, parce que c'est le **plus petit premier saut viable** — la plateforme reste **proche de VB6** — et qu'il livre une **base validée**. Le .NET moderne **ensuite** (et seulement si nécessaire), parce que c'est un **lot de changements distinct et séparable**. Le séquençage **isole les variables**, rend les régressions **diagnosticables**, et transforme une migration **terrifiante** en **deux** étapes **maîtrisées et validées séparément**.

La section suivante passe à la décision concrète : **comment évaluer** s'il faut, ou non, entreprendre ce second saut.

---

> 🧭 **Navigation**  
> ⬅️ [20. Et après ? — introduction](README.md)  
> ⬆️ [Sommaire de la formation](../SOMMAIRE.md)  
> ➡️ Suivant : [20.2 — Évaluer un second saut vers .NET 8/10](02-evaluer-second-saut.md)

---

**Juin 2026** · Cible actuelle : .NET Framework 4.7.2 (le « pont » depuis VB6) · Suite optionnelle : .NET 10 LTS · Licence : Creative Commons BY-NC-SA 4.0

⏭️ [Évaluer un second saut vers .NET 8/10 (intérêt, coût, risques)](/20-apres-net-moderne/02-evaluer-second-saut.md)
