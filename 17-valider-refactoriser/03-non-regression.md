🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.3 — Tests de non-régression : comparer avec le *golden master* ⭐

**Le filet de sécurité de toute la migration : rejouer, sur l'application migrée, les scénarios capturés sur l'application VB6 d'origine, et comparer automatiquement les sorties — toute différence devenant un signal à examiner.**

> ⭐ La **pièce maîtresse** du module, invoquée partout ailleurs (17.1, 17.2, 17.4, 17.5…). Le module
> 5.5 a **monté le harnais** et **capturé la référence** sur le VB6, *avant* la migration ; cette
> section traite la **phase de comparaison** et la méthode de non-régression. C'est ce mécanisme qui
> permet de durcir et de refactoriser **en sécurité** — et de débusquer les changements silencieux que
> l'inspection à l'œil ne révèle pas.

---

## Le principe : prouver l'**équivalence**, pas la correction

Le test *golden master* (ou test de **caractérisation**) consiste à **capturer le comportement réel**  
d'un système pour un large jeu d'entrées, à figer ces sorties comme **référence**, puis — après  
modification — à **rejouer et comparer**. Toute différence est signalée pour **jugement humain**.

Pour une migration, c'est l'outil idéal, car la question n'est pas « le résultat est-il *correct* dans  
l'absolu ? » mais « l'application migrée se comporte-t-elle **à l'identique** de l'original ? ». **Le  
comportement VB6 fait office de spécification — bugs compris.** On sépare ainsi nettement deux  
questions : *rendre équivalent* (l'objet de la migration) et *corriger des bugs* (une décision
**ultérieure**, prise sciemment — discipline du gel de périmètre, module 5.6).

> ⏱️ **La séparation temporelle est cruciale.** La référence **doit** provenir de l'application **VB6
> d'origine**, capturée **avant** toute modification (module 5.5). Capturer la référence sur
> l'application **déjà migrée** reviendrait à la comparer **à elle-même** — sans aucune valeur.

---

## Que capturer, que comparer ?

La difficulté d'une application de bureau : la « sortie » dépasse la simple valeur de retour d'une  
fonction. Plusieurs surfaces, par ordre de valeur et de facilité :

- **Le cœur calculatoire** ⭐ — résultats de la logique métier : factures, totaux, états, valeurs
  calculées. Le plus **rentable** et le plus facile à rendre déterministe — et c'est **précisément là**
  que les pièges silencieux (arrondis, dates, débordements, ByRef) se traduisent par des **nombres
  faux**.
- **Les fichiers générés** — exports, états, fichiers de données : comparer leur contenu exact.
- **Les effets en base** — lignes écrites/modifiées : comparer l'état résultant ou les opérations.
- **L'impression** — flux d'impression / rendus : plus délicat, on compare le contenu rendu.
- **L'état de l'UI** — valeurs affichées après une séquence d'actions : le plus difficile à automatiser.

> 🎯 **Stratégie** : concentrer le *golden master* sur le **cœur métier** d'abord (valeur maximale,
> déterminisme atteignable), et valider l'**UI** par des tests ciblés ou manuels. C'est le **découplage
> UI / logique** du module 5.3 qui rend ce cœur testable de façon fiable.

---

## Rendre les sorties déterministes

Un *golden master* exige des sorties **reproductibles** : la seule variable doit être « code VB6 vs code  
VB.NET », rien d'autre. Maîtrisez donc les sources de non-déterminisme :

- **Date et heure courantes** → injecter une **horloge fixe** (ne pas appeler `Now`/`Today` en dur).
- **Aléatoire** → **graine** fixe ; **GUID** → valeurs contrôlées.
- **Ordre** d'énumération (qui peut différer VB6 → .NET, module 8.5) → ordre stabilisé.
- **État initial de la base** → jeu de données **connu**, réinitialisé avant chaque exécution.
- **Culture / paramètres régionaux** ⚠️ → **épingler** la culture, car le piège des conversions
  sensibles à la culture (annexe B.10) fait varier les sorties selon la locale. Épinglez-la pour la
  reproductibilité — et testez **séparément** et **délibérément** les locales que vous livrez.

---

## Le corpus d'entrées

Le *golden master* ne détecte que les bugs des chemins que vos entrées **exercent**. La **couverture**  
est donc déterminante. Un bon corpus combine :

- des **données réelles** représentatives (anonymisées, issues de la production) ;
- des **cas limites** construits exprès pour viser les pièges du **catalogue de l'annexe B** : valeurs
  aux **bornes** (débordements d'entiers, bornes de tableaux), **vide / `Null` / `Empty`** (états du
  `Variant`), grandes valeurs, **caractères spéciaux** (ANSI/Unicode, module 16.5), valeurs
  **sensibles à la locale** (séparateurs, dates — B.10), **années bissextiles** et bornes de mois.

> 💡 **Les cas limites sont là où vivent les bugs silencieux.** Ne vous contentez pas du « chemin
> heureux » : ciblez explicitement les pièges connus.

---

## ⚠️ Normaliser le bruit — **sans masquer les bugs**

C'est le point le plus délicat de la section. Certaines différences sont **attendues et anodines** et  
doivent être filtrées, sinon chaque exécution échoue pour de mauvaises raisons. **Mais** une  
normalisation trop agressive **détruit l'intérêt même** de la démarche : elle **cache** les bugs  
silencieux que l'on cherche.

| Différence | Nature | Action |
|------------|--------|--------|
| Horodatage « généré le … » | **Bruit** (non pertinent) | ✅ Normaliser (masquer) |
| Chemin / machine / utilisateur | **Bruit** | ✅ Normaliser |
| Un total qui change (1000,00 → 999,99) | **Valeur** | ❌ **NE PAS** normaliser → bug potentiel (arrondi, 7.2 / B.9) |
| Ordre d'une liste | **Peut-être significatif** | ⚠️ Trier **seulement** si l'ordre est prouvé non pertinent (8.5) |
| Séparateur décimal (1,000.00 vs 1 000,00) | **Piège de culture** | ⚠️ À **examiner** (B.10), pas à masquer |

> ⚠️ **Le principe** : **normaliser uniquement ce qui est *prouvé* non pertinent ; traiter toute autre
> différence comme un bug potentiel jusqu'à preuve du contraire.** N'« arrondissez » jamais un écart de
> **valeur** pour faire passer le test — ce serait transformer le détecteur de bugs en générateur de
> bugs.

---

## Le filet en action : l'intégrer au workflow

La comparaison *golden master* se lance **après chaque lot de modifications** — corrections
`Option Strict` (17.1), retrait de `Compatibility` (17.2), correction de pièges (17.4), refactoring
(17.5) :

```vbnet
' Forme d'un test de non-régression : le harnais rejoue, puis compare
For Each scenario In CorpusDeScenarios()
    Dim resultat As String = Normaliser(ExecuterLogiqueMigree(scenario.Entree))
    Dim reference As String = Normaliser(LireGoldenMaster(scenario.Nom))   ' capturé sur le VB6 (5.5)

    If resultat <> reference Then
        EnregistrerDiff(scenario.Nom, reference, resultat)                 ' pour jugement humain
    End If
Next
```

- **Vert** (tout concorde, ou tout écart est compris et justifié) → on peut **avancer**.
- **Rouge** (différence inattendue) → on **s'arrête**, on diagnostique, on corrige.
- **Automatisez-le** : un seul script / une étape d'**intégration continue** (module 6.5), exécutée
  **fréquemment**. Stockez les fichiers de référence en **gestion de versions**, aux côtés du code, pour
  que les écarts soient **visibles en revue**.

---

## Vert ou rouge : classer **chaque** écart

Rappel du **principe cardinal** du module : tout écart appartient à l'une de **deux** catégories — jamais  
une troisième « c'est sûrement bénin, on ignore » :

1. **Bug de migration** → corriger le code, rejouer, l'écart disparaît.
2. **Différence attendue et justifiée** (vous corrigez *sciemment* un bug VB6, ou un écart est
   réellement acceptable) → **documenter** la raison, puis **re-baseliner** (entériner la nouvelle
   sortie comme référence pour ce scénario).

> ⚠️ **Le re-baselining doit être délibéré et revu.** Jamais un « tout accepter » réflexe pour faire
> passer le rouge au vert : ce serait **bénir silencieusement** des régressions. On n'entérine une
> nouvelle référence qu'**en connaissance de cause**, avec la justification tracée.

---

## *Golden master* et chasse aux pièges : un duo

Le *golden master* et le module **17.4** forment un couple :

- le **17.3 détecte** : « le scénario 47 a divergé » ;
- le **17.4 diagnostique** : le catalogue de l'annexe B aide à identifier *pourquoi* — un `ByRef` devenu
  `ByVal`, un arrondi monétaire, une date, une borne de tableau, une finalisation.

L'un sans l'autre est incomplet : détecter sans savoir interpréter, ou chercher des pièges sans filet  
pour confirmer qu'on les a bien neutralisés.

---

## Limites honnêtes

- **Couverture** : le *golden master* ne voit que ce que le corpus exerce. Les chemins non testés
  restent des angles morts — d'où l'importance de viser les cas limites.
- **Équivalence ≠ correction** : si le VB6 avait un bug, la référence le **fige**, et une migration
  fidèle le **reproduit**. C'est voulu pendant la migration (séparer « rendre équivalent » de « corriger
  », module 5.6) ; les corrections viendront après, **re-baselinées** sciemment.
- **UI et interactif** : difficiles à capturer ; privilégier la logique, et des tests ciblés/manuels
  pour l'interface.
- **Non-déterminisme** : tout ce qui dépend du temps, de la culture ou de l'ordre doit être **maîtrisé**
  avant de pouvoir comparer.

> 🛠️ Des **modèles** de harnais et de comparaison sont fournis à l'**annexe F** ; la **checklist** de
> l'annexe E aide à ne rien oublier. La construction peut rester **simple** (rejouer des scénarios,
> écrire les sorties dans des fichiers, comparer aux fichiers de référence versionnés) — pas besoin
> d'outil sophistiqué.

---

## En résumé

- Le test *golden master* prouve l'**équivalence** (pas la correction) : on rejoue sur le code migré les
  scénarios **capturés sur le VB6 d'origine** (5.5) et on **compare**. Le comportement VB6 est la
  référence, bugs compris.
- **Capturer la référence AVANT** la migration, sur l'original — jamais sur le code déjà migré.
- **Capturer le cœur calculatoire** en priorité (là où les pièges deviennent des nombres faux), rendre
  les sorties **déterministes** (horloge, aléatoire, ordre, base, **culture**), et bâtir un **corpus**
  qui vise les cas limites de l'**annexe B**.
- **Normaliser uniquement le bruit prouvé non pertinent** ; ne **jamais** masquer un écart de **valeur**
  (ce serait cacher le bug recherché). ⚠️
- **Filet en action** : exécuter après **chaque** lot (17.1/17.2/17.4/17.5), **automatiser** (CI, 6.5),
  **classer chaque écart** (bug → corriger ; différence voulue → documenter et **re-baseliner
  sciemment**).
- Le *golden master* **détecte**, le module 17.4 **diagnostique** : un duo indissociable.

> ➡️ **Quand un scénario diverge, encore faut-il savoir *pourquoi*. Cap sur l'audit systématique des
> pièges silencieux — ByRef, arrondis, dates, bornes de tableaux, finalisation.** (Module 17.4)

---

🏷️ **Indicateurs** : ⭐ Point clé (le filet de sécurité) · lié aux modules 5.3 (découplage), 5.5 (capture), 5.6 (gel du périmètre), 6.5 (CI), 8.5 (ordre), 16.5 (Unicode), 17.4 (pièges) · annexes B, E, F
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Traquer les **pièges silencieux** (ByRef, arrondis, dates, bornes de tableaux, finalisation)](/17-valider-refactoriser/04-traquer-pieges.md)
