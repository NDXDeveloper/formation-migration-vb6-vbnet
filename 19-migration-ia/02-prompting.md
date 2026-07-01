🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.2 Prompting : convertir un module/une classe, expliquer un comportement VB6 obscur ⚠️

> 🧭 *Module 19 — Migrer avec l'aide de l'IA · Section 2 / 5*

La section 19.1 a dressé la carte : où l'IA aide, où elle échoue. Cette section porte sur le **levier** qui décide de quel côté vous tombez — le **prompt**. Un modèle ne vaut que par la façon dont on le pilote : le même assistant produit du déchet ou de l'or selon le cadrage qu'on lui donne. On y traite les deux usages les plus rentables identifiés au point précédent — **convertir** un module/une classe et **expliquer** un comportement VB6 obscur — et la **règle qui gouverne tout le reste**.

> ⚠️ **La règle qui prime sur toutes les autres.**
> Dans **chaque** *prompt*, précisez explicitement **« VB6 »** (le code source) vs **« VB.NET »** (la cible), **et** la cible **« .NET Framework 4.7.2 »** (pas .NET 8/10). C'est la **première cause** de réponses hors-sujet : sans cette précision, l'IA mélange les deux dialectes et dérive vers le .NET moderne. Cette habitude est le **geste le plus rentable** de tout le module.

---

## Le bloc de contexte : l'en-tête de tout bon prompt

Avant la moindre demande, donnez à l'IA le **contexte qu'elle ne peut pas deviner** (rappel de la limite n°4 de la section 19.1 : elle ne voit que ce que vous lui fournissez). Le plus simple est de réutiliser un **bloc de contexte** standard :

```text
Contexte :
- Source : Visual Basic 6 (VB6, runtime VBVM).
- Cible : VB.NET sur .NET Framework 4.7.2 (Visual Studio 2026) — surtout PAS .NET 8/10.
- Contraintes :
  • conserver à l'identique le comportement d'exécution ;
  • viser un code compatible « Option Strict On » (signale ce qui l'empêche) ;
  • ne PAS utiliser l'espace de noms Microsoft.VisualBasic.Compatibility ;
  • préserver les passages ByRef, les tailles de types et les règles d'arrondi.
```

Ce préambule à lui seul élimine la majorité des malentendus. Adaptez-le à votre projet (références COM, postures `Option Strict`/`Explicit`, conventions internes).

---

## Les principes d'un *prompt* efficace

- **Désambiguïser le dialecte et la cible** — la règle ci-dessus, dans chaque échange.
- **Fournir le contexte invisible** — dépendances, rôle du code, références COM/ActiveX que l'IA ne peut pas inférer.
- **Une tâche à la fois, périmètre restreint** — un module ou une classe, pas un projet entier (rappel : « laisse courte »).
- **Demander le raisonnement, pas seulement le résultat** — « explique tes choix », « signale les changements de comportement potentiels ». C'est ce qui fait **remonter les pièges silencieux** au lieu de les enterrer.
- **Imposer des contraintes explicites** — `Option Strict`, pas de `Compatibility`, conservation des `ByRef`, des types et des arrondis.
- **Demander à l'IA de signaler ses incertitudes** — « indique ce dont tu n'es pas sûr ». On transforme ainsi l'assistant en **détecteur** (objet de la section 19.3).
- **Itérer** — raffiner par étapes vaut mieux qu'accepter la première réponse.

---

## Cas 1 — Convertir un module ou une classe

Fournissez l'**unité complète** (la classe ou le module entier), jamais des fragments isolés : l'IA a besoin du contexte **local** pour ne pas casser un invariant interne. Demandez ensuite **trois choses en plus de la conversion** : la liste des risques de comportement, un mapping de types explicite, et un tableau de points à vérifier.

**Modèle de *prompt* (à copier et compléter) :**

```text
[Bloc de contexte ci-dessus]

Voici la classe VB6 à convertir :
<<< coller ici le .cls complet >>>

Demande :
1. Convertis cette classe en VB.NET ciblant .NET Framework 4.7.2.
2. Liste TOUS les changements de comportement POTENTIELS : ByRef→ByVal,
   types redimensionnés (Integer/Long/Currency), Variant (Empty/Null/Missing),
   finalisation (Class_Terminate → IDisposable/Using), tableaux 0-based,
   conversions sensibles à la culture (CDbl/CDate/Format).
3. Donne un tableau « à vérifier » : élément | risque | comment le tester.
4. Explique tes choix non triviaux, et indique ce dont tu n'es pas sûr.
```

**Pourquoi le cadrage change tout.** Comparez :

> ❌ *Prompt* faible : « Convertis ce code : … »
> Aucun dialecte (VB6 ou VB.NET ?), aucune version cible (dérive vers .NET 8/10), aucune contrainte de préservation, aucune demande de signaler les pièges. Résultat probable : du code qui **compile mais ne se comporte plus pareil**, sans le moindre avertissement.

> ✅ *Prompt* fort : le modèle ci-dessus.
> Dialecte et cible explicites, contraintes posées, et surtout l'IA est **forcée d'exposer** les risques et son raisonnement — exactement là où se cachent les régressions.

La conversion obtenue reste une **proposition** : elle passe ensuite au *golden master* et à la relecture via l'[Annexe B](../annexes/pieges-silencieux/README.md). Pour les correspondances de mots-clés et de fonctions, recoupez avec l'[Annexe A](../annexes/correspondance-vb6-vbnet/README.md).

---

## Cas 2 — Expliquer un comportement VB6 obscur

C'est l'atout sous-estimé de la section 19.1. Sur du *legacy* peu documenté, demandez d'abord à l'IA de **comprendre**, pas de convertir : faites-lui **rendre explicites** les comportements **implicites** de VB6 (propriétés par défaut, `ByRef` implicite, ordre des événements, sémantique de `Variant`), afin de pouvoir les **préserver**. C'est aussi la matière idéale pour [documenter les comportements critiques](../05-preparer-code-vb6/04-documenter-comportements.md) avant la migration.

**Modèle de *prompt* (explication, sans conversion) :**

```text
Contexte : le code ci-dessous est du code SOURCE en VB6 (Visual Basic 6).
Je veux COMPRENDRE son comportement avant de le migrer vers VB.NET
(.NET Framework 4.7.2). Ne convertis rien pour l'instant.

Voici le code VB6 :
<<< coller ici >>>

Demande :
1. Explique ce que fait ce code et POURQUOI (ligne par ligne si nécessaire).
2. Rends explicites les comportements IMPLICITES de VB6 : propriétés par défaut,
   ByRef implicite, ordre des événements, sémantique de Variant (Empty/Null/Missing),
   arrondis monétaires.
3. Signale les cas limites et les hypothèses cachées.
4. Indique ce qu'il faudra impérativement PRÉSERVER lors de la migration.
```

Cette explication n'est pas parole d'évangile non plus : c'est une **lecture assistée** qui accélère votre compréhension, à confirmer sur le code réel. Sa valeur est de **mettre des mots** sur ce qui était tacite — première étape pour ne rien casser.

---

## Désambiguïsation : pourquoi c'est vital (et comment faire)

Sans précision du dialecte et de la version, l'IA **confond les deux mondes** : elle « corrige » du VB6 avec de la syntaxe VB.NET, convertit vers du .NET moderne, ou propose des tournures qui n'existent que dans l'un des deux langages. Quelques cas typiques :

| Ambiguïté si non précisé | Risque concret | La précision qui l'évite |
|--------------------------|----------------|--------------------------|
| « convertis ce code » sans dialecte | l'IA mélange VB6 et VB.NET (syntaxe « corrigée » à tort) | dire « source **VB6** » / « cible **VB.NET** » |
| version .NET non indiquée | dérive vers **.NET 8/10** : API et idiomes **absents de 4.7.2** (relève du [module 20](../20-apres-net-moderne/README.md)) | écrire **« .NET Framework 4.7.2 »** explicitement |
| mot-clé `Set`, propriétés par défaut | confusion sur l'instanciation et les accès (VB6 a `Set`, pas VB.NET) | rappeler qu'on **part de VB6** |
| tailles de types (`Integer`/`Long`) | mapping erroné, débordements silencieux | préciser **source VB6** + demander un **mapping explicite** |

> 💡 Le réflexe à automatiser : **« source VB6 → cible VB.NET sur .NET Framework 4.7.2 »** dans chaque *prompt*, sans exception.

---

## Anti-modèles de *prompt* (à éviter)

- **« Convertis ce code »** — sans dialecte ni cible : la porte ouverte aux confusions.
- **Coller un fragment hors contexte** — l'IA ne peut pas raisonner sur ce qu'elle ne voit pas.
- **Demander un projet entier d'un coup** — impossible à auditer, incohérences garanties.
- **Accepter la première réponse** — sans exiger la liste des risques ni le raisonnement.
- **Omettre la version .NET** — c'est la dérive vers le .NET moderne assurée.
- **Faire confiance aux « équivalents »** proposés sans les vérifier (voir 19.5).

---

## À retenir

Bien *prompter*, c'est **donner à l'IA le contexte qui lui manque** (ses angles morts de 19.1) et la **forcer à exposer son raisonnement et ses risques**. La règle de désambiguïsation — **« VB6 » vs « VB.NET » + cible « .NET Framework 4.7.2 »** — est l'habitude au plus fort rendement de tout le module. Les modèles de *prompts* ci-dessus sont des points de départ : adaptez-les à votre projet, mais gardez intacts le **cadrage** et la **demande de signaler les pièges**.

La section suivante exploite directement ces *prompts* pour un objectif précis : faire de l'IA un **détecteur** des pièges silencieux, plutôt que leur producteur.

---

> 🧭 **Navigation**  
> ⬅️ [19.1 — Pourquoi l'IA aide (et ses limites)](01-pourquoi-ia.md)  
> ⬆️ [Sommaire de la formation](../SOMMAIRE.md)  
> ➡️ Suivant : [19.3 — Détecter les pièges avec l'IA](03-detecter-pieges.md)

---

**Juin 2026** · Cible : .NET Framework 4.7.2 · Licence : Creative Commons BY-NC-SA 4.0

⏭️ [Détecter les pièges avec l'IA (ByRef, `Variant`, finalisation, entiers redimensionnés)](/19-migration-ia/03-detecter-pieges.md)
