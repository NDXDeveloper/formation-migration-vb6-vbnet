🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.3 Détecter les pièges avec l'IA (ByRef, `Variant`, finalisation, entiers redimensionnés) ⭐ ⚠️

> 🧭 *Module 19 — Migrer avec l'aide de l'IA · Section 3 / 5*

Voici le paradoxe central de ce module. La section 19.1 l'a posé sans détour : l'IA est **la plus faible là où cette migration est la plus dangereuse**, c'est-à-dire sur les **pièges silencieux** — au point qu'elle en **produit** elle-même. Comment, dès lors, en faire un outil de **détection** ?

La réponse tient en un mot : **aiguillage**. L'IA ne signale **pas spontanément** ces pièges, mais si vous la **dirigez** vers un motif précis et nommé — « cherche les passages `ByRef` devenus `ByVal` », « repère les pertes de `Variant` » — elle devient un **scanner de premier passage** très utile. C'est une lampe torche qu'on **pointe**, pas un gardien autonome.

> ⚠️ **Deux règles non négociables.**
> 1. L'IA **complète** le *golden master* et l'[Annexe B](../annexes/pieges-silencieux/README.md), elle ne les **remplace jamais**.
> 2. Un piège n'est **jamais** déclaré « absent » au seul motif que l'IA ne l'a pas trouvé. **L'absence de signalement n'est pas une preuve d'absence de bug.**

---

## Le bon réflexe : une IA qu'on aiguille, pas un garde autonome

Demander « ce code a-t-il des problèmes ? » donne des réponses vagues et trompeusement rassurantes. La détection efficace fait **l'inverse de la confiance** : on **présuppose** que chaque piège connu *peut* être présent, et on demande à l'IA de le **chercher activement**, un par un. On exploite ainsi sa vraie force (reconnaître des **motifs syntaxiques** nommés) tout en contournant sa faiblesse (juger une **sémantique d'exécution** qu'elle ne « voit » pas).

---

## La méthode : un piège à la fois, nommé explicitement

Reprenez le **bloc de contexte** de la section 19.2 (source VB6, cible .NET Framework 4.7.2, contraintes), puis appliquez trois principes :

- **Une famille de piège par passe.** Mélanger les recherches dilue l'attention de l'IA et la vôtre.
- **Nommer le motif précisément** — pas « des bugs », mais « les paramètres dont le mode `ByRef`/`ByVal` a changé ».
- **Exiger une sortie au format de l'Annexe B** : pour chaque occurrence, **emplacement | symptôme | cause | correction proposée | comment le tester**. Les constats de l'IA s'alignent alors directement sur votre catalogue de pièges.

**Modèle de *prompt* d'audit (à copier et spécialiser) :**

```text
[Bloc de contexte : source VB6 → cible VB.NET .NET Framework 4.7.2]

Voici (a) le code VB6 source et (b) sa conversion VB.NET :
--- VB6 ---
<<< coller ici >>>
--- VB.NET ---
<<< coller ici >>>

Demande : agis comme un AUDITEUR de migration.
Cherche UNIQUEMENT le piège suivant : <<< NOMMER LE PIÈGE >>>.
Pour chaque occurrence trouvée, donne un tableau :
  emplacement | symptôme | cause | correction proposée | comment le tester.
Indique clairement ce dont tu n'es pas sûr.
Ne signale aucun autre type de problème.
```

---

## Les quatre pièges à faire chercher

### 1. `ByRef` devenu `ByVal` — [B.1](../annexes/pieges-silencieux/README.md)
**Le piège.** En VB6, `ByRef` est le mode **implicite** ; en VB.NET, c'est `ByVal` ([rappel 2.4](../02-differences-fondamentales/04-byref-byval.md)). Une conversion littérale qui laisse tomber le mot-clé **inverse silencieusement** la sémantique : une procédure qui modifiait son argument ne le modifie plus chez l'appelant.

**Ce qu'on demande à l'IA.** Repérer chaque paramètre dont le mode a changé entre source et conversion, **et** chaque paramètre **affecté dans le corps** de la procédure mais désormais passé `ByVal`.

**Ce qu'elle repère bien.** Les **affectations explicites** au paramètre (motif syntaxique net). **Ce qu'elle manque.** Les mutations **indirectes** et surtout la question décisive : *l'appelant dépend-il de la modification ?* — réponse qui exige les **sites d'appel**, souvent hors du fragment fourni.

**Vérification humaine.** *Golden master* sur les routines concernées.

### 2. `Variant` : `Empty` / `Null` / `Missing` perdus — [B.5](../annexes/pieges-silencieux/README.md)
**Le piège.** Le `Variant` de VB6 distingue `Empty`, `Null` et `Missing` — distinctions sans équivalent propre en VB.NET. La conversion vers `Object` (ou un type concret) **efface** ces états, et la logique `IsEmpty`/`IsNull`/`IsMissing` se met à brancher autrement.

**Ce qu'on demande à l'IA.** Localiser tous les usages de `Variant` et des fonctions `IsEmpty`/`IsNull`/`IsMissing`/`VarType`, et expliquer comment **chaque** cas `Empty`/`Null`/`Missing` est traité après conversion ; signaler toute distinction perdue.

**Ce qu'elle repère bien.** Le mot-clé `Variant` et les fonctions intrinsèques. **Ce qu'elle manque.** Les `Variant` **implicites** (variables sans clause `As`) et les états **réels** des données qui transitent — qu'elle ne voit pas.

**Vérification humaine.** Contrôle des valeurs effectives + *golden master* sur les branches concernées.

### 3. Finalisation : `Class_Terminate` qui ne se déclenche plus — [B.3](../annexes/pieges-silencieux/README.md)
**Le piège.** En VB6, `Class_Terminate` se déclenche **de façon déterministe** quand le compteur de références tombe à zéro ; en VB.NET, le *garbage collector* décide **quand** ([rappel 2.2](../02-differences-fondamentales/02-finalisation-deterministe.md)). Sans bascule vers `IDisposable`/`Using`, les ressources (handles de fichiers, objets COM, connexions) sont libérées **au mauvais moment**, voire tardivement.

**Ce qu'on demande à l'IA.** Repérer chaque classe avec `Class_Terminate` (ou qui détient des ressources) et proposer un équivalent fondé sur `IDisposable`/`Using` ; signaler tout nettoyage déterministe que le GC va désormais **différer**.

**Ce qu'elle repère bien.** Le gestionnaire `Class_Terminate` et le patron `Dispose`. **Ce qu'elle manque.** La **dépendance de *timing*** : qu'un autre code ait compté sur un nettoyage survenant **exactement** au moment où l'objet sortait de portée — un invariant qu'elle ne peut pas inférer sans la vue d'ensemble.

**Vérification humaine.** Durées de vie des ressources et fuites de *handles* ([module 17.6](../17-valider-refactoriser/06-performance-memoire.md)).

### 4. Entiers redimensionnés et débordements — [B.2](../annexes/pieges-silencieux/README.md)
**Le piège.** Les types changent de taille ([rappel 2.3](../02-differences-fondamentales/03-typage.md)) : VB6 `Integer` (16 bits) → VB.NET `Integer` (**32 bits**), `Long` (32 bits) → `Long` (**64 bits**), et `Currency` → `Decimal`. Garder le même mot-clé **change la sémantique** : les points de débordement se déplacent, l'interop à largeur fixe se désaligne, et `Currency → Decimal` modifie le comportement d'arrondi.

**Ce qu'on demande à l'IA.** Établir le mapping de **chaque** type numérique avec sa **largeur exacte**, signaler tout endroit où la largeur a changé, et identifier le code reposant sur le **débordement**, sur l'**interop à largeur fixe** (`Declare`/structures d'API) ou sur l'**arrondi `Currency`**.

**Ce qu'elle repère bien.** Les déclarations de types, les écarts de taille évidents et la bascule `Currency → Decimal`. **Ce qu'elle manque.** La **dépendance subtile au débordement** et la **disposition exacte** des structures d'interop, sauf si elle voit les signatures d'API concernées.

**Vérification humaine.** [Annexe D](../annexes/correspondance-types/README.md) (types) + *golden master* ; dispositions `P/Invoke` ([module 16.5](../16-interop-com-api/05-marshaling.md)).

---

## Les limites de la détection par IA

- **Faux négatifs.** L'IA ne **garantit pas l'exhaustivité** : elle rate des occurrences, surtout celles qui dépendent du contexte global.
- **Faux positifs.** Elle signale des « problèmes » qui n'en sont pas ; chaque alerte doit être **vérifiée**, jamais corrigée à l'aveugle.
- **Ni *runtime*, ni vue d'ensemble.** Pas de modèle d'exécution, pas d'accès aux sites d'appel ni aux définitions de structures situées ailleurs.
- **Non-déterminisme.** Mêmes entrées, **résultats variables** : relancez l'audit plusieurs fois — mais cela **réduit** le risque sans l'**annuler**.

> ⚠️ **À graver :** l'audit IA est une **couche de triage**, le *golden master* est le **verdict** ([module 5.5](../05-preparer-code-vb6/05-golden-master.md), [Annexe F](../annexes/golden-master/README.md)). On utilise l'IA pour **diriger l'attention**, pas pour **certifier** l'absence de piège.

---

## Un déroulé pratique

1. **Conversion** (section 19.2), code source et code converti côte à côte.
2. **Audit dirigé**, *une famille de piège à la fois*, avec le modèle de *prompt* ci-dessus.
3. **Revue humaine** de chaque signalement : confirmer, écarter le faux positif, corriger.
4. **Confirmation au *golden master*** : c'est lui qui tranche le comportement.
5. **Bouclage avec l'Annexe B** : cocher chaque piège **examiné** (et non « supposé absent »).

La liste des motifs à faire chercher se lit directement dans l'[Annexe B](../annexes/pieges-silencieux/README.md) — au-delà des quatre familles ci-dessus, pensez aussi à `On Error Resume Next`, aux tableaux **0-based**, aux propriétés par défaut, à `True = -1` et aux conversions **sensibles à la culture**.

---

## Tableau de synthèse

| Piège | Ce qu'on demande à l'IA de chercher | Ce qu'elle repère bien | Ce qui reste humain |
|-------|-------------------------------------|------------------------|---------------------|
| `ByRef`→`ByVal` (B.1) | params dont le mode a changé + affectés dans le corps | affectations explicites au paramètre | dépendance de l'appelant (sites d'appel), *golden master* |
| `Variant` (B.5) | `Variant` + `IsEmpty`/`IsNull`/`IsMissing`/`VarType` | mots-clés et fonctions `Is*` | `Variant` implicites, états réels des données |
| Finalisation (B.3) | `Class_Terminate` + ressources détenues | le gestionnaire, le patron `IDisposable`/`Using` | dépendances de **timing** du nettoyage |
| Entiers redim. (B.2) | mapping de types + largeurs + overflow/interop | déclarations, `Currency → Decimal` | reliance sur l'overflow, disposition `P/Invoke` exacte |

---

## À retenir

L'IA est un **détecteur dirigé**, pas un gardien : elle ne trouve bien que ce qu'on lui demande **explicitement** de chercher, motif par motif, et elle ne **certifie** jamais l'absence d'un piège. Bien employée, elle **accélère le triage** des régressions silencieuses ; mal employée, elle **endort la vigilance** par des rapports incomplets. Le *golden master* reste l'**arbitre**, l'Annexe B reste la **liste de contrôle**.

La section suivante prolonge cette logique de filet de sécurité : faire **générer** par l'IA les **tests de non-régression** et la **documentation XML** qui rendent ce contrôle reproductible.

---

> 🧭 **Navigation**  
> ⬅️ [19.2 — Prompting : convertir, expliquer](02-prompting.md)  
> ⬆️ [Sommaire de la formation](../SOMMAIRE.md)  
> ➡️ Suivant : [19.4 — Générer tests de non-régression et documentation XML](04-tests-documentation.md)

---

**Juin 2026** · Cible : .NET Framework 4.7.2 · Licence : Creative Commons BY-NC-SA 4.0

⏭️ [Générer tests de non-régression et documentation XML](/19-migration-ia/04-tests-documentation.md)
