🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.5 Pièges de l'IA : hallucinations, faux équivalents VB6/VB.NET, validation systématique ⚠️

> 🧭 *Module 19 — Migrer avec l'aide de l'IA · Section 5 / 5*

Les sections précédentes ont montré **où** l'IA aide et **comment** l'employer. Cette dernière section regarde l'IA **en face** : non plus les pièges de la migration (la sémantique VB6 → VB.NET), mais les **modes de défaillance propres à l'IA**. Il y en a deux principaux — les **hallucinations** et les **faux équivalents** — et ils partagent une signature redoutable : *l'IA est la plus dangereuse quand elle est la plus assurée.*

> ⚠️ **Le constat qui commande tout le reste.** Le problème n'est pas que l'IA se trompe — c'est qu'elle se trompe **de façon convaincante** : un texte fluide, plausible et sans la moindre hésitation. On ne peut donc **pas** détecter ses erreurs « au feeling ». La seule parade est un **protocole de validation systématique**, appliqué à **chaque** sortie.

---

## 1. Les hallucinations

**Ce que c'est.** L'IA **invente** des éléments qui n'existent pas mais paraissent réels : méthodes, surcharges, propriétés, espaces de noms, signatures — voire des **comportements VB6** qui n'ont jamais existé. Un modèle de langage génère du texte **plausible**, pas des faits **vérifiés** ; « plausible » et « vrai » divergent surtout sur les API rares ou très spécifiques.

**Les formes propres à cette migration :**

- **API .NET inventée** : méthode, surcharge ou espace de noms qui « devrait » exister mais n'existe pas.
- **Hallucination de version** : méthode bien réelle dans le **.NET moderne (8/10)** mais **absente de 4.7.2** — fausse relativement à *votre* cible (rappel de 19.1 et de la règle de cible de 19.2).
- **Comportement VB6 mal restitué** : « VB6 faisait X » alors que non.
- **Béquille dépréciée présentée comme bonne pratique** : recommander `Microsoft.VisualBasic.Compatibility` comme si c'était recommandé — précisément ce qu'il faut [éviter](../04-outils-migration/02-compatibility-namespace.md).
- **Déclaration d'API plausible mais fausse** : signature `P/Invoke` ou disposition de structure inexacte.

**Comment les rattraper.** La **compilation** filtre une grande partie des hallucinations (ce qui n'existe pas ne compile pas) — à condition de **cibler 4.7.2**, pas une version plus récente. Ce qui **compile quand même** relève de l'autre catégorie.

---

## 2. Les faux équivalents VB6/VB.NET

**Ce que c'est.** Le mode de défaillance **le plus insidieux**, parce qu'il **compile** souvent. L'IA propose une construction VB.NET comme « équivalente » à une construction VB6, mais la **sémantique d'exécution diffère**. C'est, vu du côté de l'IA, **tout le catalogue des pièges silencieux** de l'[Annexe B](../annexes/pieges-silencieux/README.md) : l'IA est la **source** du faux équivalent qui introduit le changement silencieux.

**Exemples (déjà croisés) :**

- `ByVal` proposé comme « équivalent » d'une routine VB6 qui dépendait de `ByRef`.
- Une conversion/un arrondi .NET présenté comme équivalent à `Currency`/`CDbl`/`Format`, mais au comportement **différent** aux cas limites (culture, séparateur décimal).
- `Object` présenté comme équivalent de `Variant`, perdant `Empty`/`Null`/`Missing`.
- `Dispose` proposé comme équivalent de `Class_Terminate`, mais **sans le *timing* déterministe**.
- Type homonyme conservé (`Integer` → `Integer`) comme « équivalent », alors que la **largeur** a changé.
- `On Error Resume Next` « traduit » en `Try/Catch` vide qui **avale** les exceptions autrement.

**Pourquoi c'est pire qu'une hallucination.** L'hallucination est souvent **bruyante** (ne compile pas) ; le faux équivalent est **silencieux** (compile et paraît juste). Il exploite exactement l'écart entre *« ça compile »* et *« ça se comporte pareil »* — le **danger central** de toute cette formation.

La distinction, sur l'exemple fil rouge :

```text
Hallucination (bruyante) :
  total.AjouterMontant(montant)        ' méthode inexistante -> NE COMPILE PAS, erreur immédiate

Faux équivalent (silencieux) :
  Sub Ajouter(montant As Decimal, ByVal total As Decimal)
                                        ' « équivalent » du ByRef VB6 -> COMPILE,
                                        ' mais « total » n'est plus mis à jour. Aucun signal.
```

**Comment les rattraper.** **Pas** en relisant le code (il a l'air correct). **Uniquement** par le *golden master*, les **tests de non-régression** et la **relecture via l'Annexe B**. C'est toute la raison d'être de la discipline d'oracle de la section 19.4.

---

## 3. La signature commune : une erreur convaincante

Les deux modes de défaillance partagent une propriété : l'IA est **fluide et assurée** que sa réponse soit juste **ou** fausse. Il faut en tirer une conséquence stricte.

> ⚠️ **Confiance ≠ exactitude.** L'assurance du ton n'est **aucun** indice de justesse ; l'absence d'hésitation ne **prouve rien**. Ne déduisez **jamais** la correction d'une réponse de la façon dont elle est formulée.

---

## 4. La validation systématique : la seule défense

La parade n'est pas « rester prudent » (un état d'esprit, donc faillible), mais un **protocole** appliqué à **chaque** sortie de l'IA, **sans exception** :

1. **Compiler** — nécessaire, jamais suffisant. Filtre les hallucinations inexistantes ; ne dit **rien** du comportement. Compiler avec la cible **.NET Framework 4.7.2**, pas une version plus récente.
2. **Confronter au *golden master*** et aux **tests de non-régression** — le **verdict de comportement**, et la **seule** chose qui attrape les faux équivalents ([module 5.5](../05-preparer-code-vb6/05-golden-master.md), [Annexe F](../annexes/golden-master/README.md)).
3. **Relire à la lumière de l'[Annexe B](../annexes/pieges-silencieux/README.md)** — la liste de contrôle humaine des pièges nommés ([module 17.4](../17-valider-refactoriser/04-traquer-pieges.md)).
4. **Vérifier l'existence réelle des API** dans la documentation — et leur présence **dans 4.7.2** précisément.
5. **Traiter chaque « équivalent » comme une hypothèse** à prouver, pas une affirmation à croire.
6. **Ne jamais déduire la justesse de l'assurance** — point 3 ci-dessus.
7. **Garder l'humain responsable** — l'IA n'est jamais l'autorité ; c'est **vous** qui validez et signez.

> ✅ **La règle qui résume le protocole :** rien de ce que produit l'IA n'entre dans la base de code sans avoir été **compilé**, **confronté au comportement de référence** et **relu**. Le reste — l'assurance, la fluidité, la vraisemblance — ne compte pas.

---

## Tableau de synthèse

| Mode de défaillance | Signe distinctif | Compile ? | Ce qui le rattrape |
|---------------------|------------------|-----------|--------------------|
| **Hallucination** | API ou comportement **inventé** | souvent **non** (bruyant) | compilation + documentation (cible **4.7.2**) |
| **Faux équivalent** | sémantique **différente** sous une forme « équivalente » | **oui** (silencieux) | *golden master* + tests + **Annexe B** |

---

## Pour clore le module

Ce module se ramène à une seule doctrine, déclinée en cinq temps : **connaître les limites** de l'IA ([19.1](01-pourquoi-ia.md)), **désambiguïser** dans chaque *prompt* ([19.2](02-prompting.md)), s'en servir comme **détecteur dirigé** ([19.3](03-detecter-pieges.md)), **bâtir un filet de sécurité validé** ([19.4](04-tests-documentation.md)), et **se défendre** de ses modes de défaillance (cette section). La constante des cinq : **le *golden master* est l'arbitre, l'humain est responsable.**

Bien employée, l'IA rend cette migration **plus rapide** et **moins ingrate**. Mais elle ne **change pas la discipline** — elle change la **vitesse** à laquelle on l'applique. L'assistant propose ; **vous validez**. Toujours.

---

> 🧭 **Navigation**  
> ⬅️ [19.4 — Générer tests de non-régression et documentation XML](04-tests-documentation.md)  
> ⬆️ [Sommaire de la formation](../SOMMAIRE.md) · [Introduction du module 19](README.md)  
> ➡️ Suivant : [20. Et après ? De 4.7.2 vers .NET moderne](../20-apres-net-moderne/README.md)

---

**Juin 2026** · Cible : .NET Framework 4.7.2 · Licence : Creative Commons BY-NC-SA 4.0

⏭️ [Et après ? De 4.7.2 vers .NET moderne (optionnel)](/20-apres-net-moderne/README.md)
