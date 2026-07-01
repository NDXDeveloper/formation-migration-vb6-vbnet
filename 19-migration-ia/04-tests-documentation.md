🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.4 Générer tests de non-régression et documentation XML ⭐

> 🧭 *Module 19 — Migrer avec l'aide de l'IA · Section 4 / 5*

La section 19.3 a fait de l'IA un **détecteur dirigé** ; mais la détection réclame un **verdict**, et ce verdict, c'est le *golden master*. Cette section porte sur l'usage de l'IA pour **construire** ce filet de sécurité — les **tests de non-régression** — et pour **documenter** le résultat — la **documentation XML**. Deux livrables, deux profils de risque bien différents.

> ⚠️ **Le piège fondateur, à comprendre avant tout le reste.**
> Un test ne vaut que par son **oracle** — la « bonne réponse » à laquelle il compare. Si l'IA écrit des tests à partir du **code migré**, elle teste *ce que le code fait maintenant*, pas *ce qu'il devrait faire* : un bug silencieux introduit par la migration devient alors un **test au vert**. Or un test vert sur du code faux est **pire que pas de test** — il donne une fausse confiance. **L'oracle, c'est le comportement VB6 de référence**, pas le code converti.

---

## Partie A — Générer des tests de non-régression

### Le principe : le *golden master* fournit la vérité, l'IA l'échafaudage

Le [*golden master*](../05-preparer-code-vb6/05-golden-master.md) ([Annexe F](../annexes/golden-master/README.md)) capture le **comportement de référence** de l'application VB6 : des entrées et leurs sorties **attendues**, saisies depuis l'original. C'est lui la **source de vérité**. Le rôle **légitime** de l'IA se limite à l'**échafaudage** : monter le harnais, proposer des cas, câbler la comparaison, mettre le code de test en forme. Son rôle **interdit** : **être** l'oracle, c'est-à-dire inventer ou recopier depuis le code migré les valeurs attendues.

> ⚠️ Règle d'or des tests : **les sorties attendues tracent vers le VB6 de référence**, jamais vers le code converti ni vers une « estimation » de l'IA.

### Où l'IA aide concrètement

- **Monter la structure du harnais** dans un *framework* .NET (MSTest, NUnit ou xUnit, tous compatibles .NET Framework 4.7.2), selon le schéma *arrange / act / assert*.
- **Proposer des jeux de cas**, en priorité les **cas limites liés aux pièges** (rappel de 19.3) : valeurs autour des **bornes de types** (débordements), `Empty`/`Null`/`Missing` pour `Variant`, **dates** et **arrondis monétaires**, **séparateur décimal** sensible à la culture, **bornes de tableaux** 0-based. L'IA est efficace pour **énumérer** ces cas dès qu'on lui **nomme** les familles de pièges.
- **Écrire les assertions et le code de comparaison** entre la sortie migrée et la valeur de référence.
- **Transformer un *golden master* « fichier »** (entrées/sorties enregistrées) en **cas de tests automatisés**.

### Le bon montage : d'où vient l'oracle

Le montage correct **sépare** ce que produit la référence (les sorties attendues) de ce que produit l'IA (le code de test). On **fournit** à l'IA les sorties de référence ; on ne lui demande **jamais** de les deviner.

**Gabarit de *prompt* (génération de tests, oracle imposé) :**

```text
[Bloc de contexte : code VB.NET migré sur .NET Framework 4.7.2]

Fonction migrée (signature / contrat) :
<<< coller la signature VB.NET >>>

Jeu d'entrées et SORTIES DE RÉFÉRENCE (issues du comportement VB6 / golden master) :
<<< entrée1 -> sortie_attendue1 ; entrée2 -> sortie_attendue2 ; ... >>>

Demande :
1. Génère des tests de non-régression (framework : MSTest / NUnit / xUnit — précise lequel)
   qui comparent la sortie de la fonction migrée aux SORTIES DE RÉFÉRENCE ci-dessus.
   N'INVENTE AUCUNE sortie attendue : utilise EXCLUSIVEMENT celles fournies.
2. Ajoute des cas limites pour les pièges suivants : bornes de types (overflow),
   Empty/Null/Missing (Variant), dates et arrondis monétaires, séparateur décimal
   sensible à la culture, bornes de tableaux. Pour ces cas, LAISSE la sortie attendue
   À COMPLÉTER par moi (depuis la référence) — ne la devine pas.
3. Structure arrange / act / assert, un test par cas. Signale ce dont tu n'es pas sûr.
```

### Vérifier les tests générés

- **Tests verts pour de mauvaises raisons** : tautologies, assertions triviales, ou oracle tiré du **code migré**. À traquer en priorité.
- **Cas limites manqués** — surtout autour des pièges silencieux.
- **Faux équivalents dans les assertions** : par exemple une comparaison d'arrondi **trop tolérante** qui masque une dérive réelle (voir 19.5).
- **Le test doit échouer quand il doit échouer** : un test qui ne casse jamais ne protège de rien.

Ces tests deviennent ensuite le socle de la phase de validation ([module 17.3](../17-valider-refactoriser/03-non-regression.md)) et de la traque des pièges ([module 17.4](../17-valider-refactoriser/04-traquer-pieges.md)).

---

## Partie B — Générer la documentation XML

### Ce que c'est

La **documentation XML** de VB.NET (`''' <summary>`, `<param>`, `<returns>`, `<remarks>`, `<exception>`) alimente IntelliSense et les générateurs de documentation. Lors d'une migration, elle a une valeur particulière : **consigner les décisions de migration** et les **comportements préservés**.

```vb
''' <summary>
''' Ajoute un montant à un total cumulé.
''' </summary>
''' <param name="montant">Montant à ajouter.</param>
''' <param name="total">Total cumulé, mis à jour par référence (ByRef).</param>
''' <remarks>
''' Migration : ByRef préservé volontairement (comportement VB6 d'origine, cf. B.1).
''' Currency → Decimal : arrondi vérifié au golden master.
''' </remarks>
```

### Où l'IA aide

- **Rédiger** les `<summary>`/`<param>`/`<returns>` à partir du code **et** de l'explication produite en [19.2](02-prompting.md).
- **Capturer l'intention de migration** dans `<remarks>` : transformer les explications de 19.2 et les [comportements documentés en amont](../05-preparer-code-vb6/04-documenter-comportements.md) en notes durables (« `ByRef` préservé », « `Class_Terminate` → `Dispose` », « arrondi vérifié »).
- **Uniformiser le style** de documentation sur l'ensemble de la base de code.

**Gabarit de *prompt* (documentation XML, sans invention) :**

```text
[Bloc de contexte : code VB.NET migré sur .NET Framework 4.7.2]

Code à documenter :
<<< coller ici >>>

Demande :
1. Génère la documentation XML VB.NET (''' <summary>, <param>, <returns>,
   <exception> le cas échéant).
2. Dans <remarks>, consigne les DÉCISIONS DE MIGRATION et COMPORTEMENTS PRÉSERVÉS
   que je te fournis — n'invente AUCUN comportement.
   Notes de migration : <<< ByRef préservé / Currency→Decimal vérifié / Class_Terminate→Dispose / ... >>>
3. Ne documente que ce qui est VÉRIFIABLE dans le code ; signale ce dont tu n'es pas sûr.
```

### Vérifier la documentation générée

- **Documentation qui ment** : `<summary>` ou `<remarks>` affirmant un comportement que le code **n'a pas**. Une documentation fausse est **pire** qu'absente.
- **Comportements supposés, non vérifiés** : la doc doit refléter le **réel** (et le *golden master*), pas une **hypothèse** de l'IA.
- **Fausse « équivalence VB6 »** : ne laissez pas une note affirmer une parité de comportement qui n'a **pas** été contrôlée.

La documentation ainsi fiabilisée prépare directement le **transfert de compétences** à l'équipe ([module 18.4](../18-deploiement-bascule/04-documentation-transfert.md)).

---

## Tests et documentation : la mémoire de la migration

Les deux livrables se nourrissent de la **même** source — l'explication du comportement obtenue en 19.2 — et se complètent : les **tests verrouillent** le comportement, la **documentation le décrit**. Ensemble, ils constituent la **mémoire** de la migration : ce qu'on a préservé, pourquoi, et comment on le vérifie. Dans les deux cas, l'IA **rédige**, mais ne **certifie** rien : la vérité reste le comportement **VB6 de référence** et le code **réel**.

---

## Tableau de synthèse

| Livrable | Rôle de l'IA | Source de vérité (oracle) | Vérification humaine |
|----------|--------------|---------------------------|----------------------|
| **Tests de non-régression** | échafaudage, cas limites, assertions | comportement **VB6 de référence** (*golden master*) | oracle **non** issu du code migré ; tests qui **échouent quand ils doivent** |
| **Documentation XML** | rédiger `summary`/`param`/`returns`/`remarks` | comportement **vérifié** du code | **aucune invention** ; refléter le réel, pas une hypothèse |

---

## À retenir

L'IA **construit le filet** et **rédige les étiquettes** ; elle ne doit **jamais** être la source de vérité. Pour les tests, l'oracle est le **comportement VB6 de référence**, pas le code migré — sans quoi on **fige les bugs** au vert. Pour la documentation, on ne consigne que du **vérifié**, jamais du supposé. Ce filet de sécurité est précisément ce qui rend le reste de la migration **auditable**.

La section suivante regarde l'IA en face : ses **propres modes de défaillance** — hallucinations, **faux équivalents** VB6/VB.NET — et la discipline de **validation systématique** qui en découle.

---

> 🧭 **Navigation**  
> ⬅️ [19.3 — Détecter les pièges avec l'IA](03-detecter-pieges.md)  
> ⬆️ [Sommaire de la formation](../SOMMAIRE.md)  
> ➡️ Suivant : [19.5 — Pièges de l'IA : hallucinations, faux équivalents](05-limites-pieges.md)

---

**Juin 2026** · Cible : .NET Framework 4.7.2 · Licence : Creative Commons BY-NC-SA 4.0

⏭️ [Pièges de l'IA : hallucinations, **faux équivalents** VB6/VB.NET, validation systématique](/19-migration-ia/05-limites-pieges.md)
