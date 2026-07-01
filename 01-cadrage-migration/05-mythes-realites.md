🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 1.5 Mythes et réalités ⚠️

> **Chapitre 1 — Introduction : pourquoi et comment migrer**
> Dernière étape du cadrage : nettoyer les idées reçues avant de commencer.

---

## Pourquoi une section entière sur des « mythes » ?

Parce que les croyances erronées sur la migration VB6 → VB.NET **coûtent plus cher que n'importe  
quelle difficulté technique**. Une difficulté technique, on la rencontre, on la traite. Un **mythe**,  
lui, fausse les décisions **en amont** — il gonfle les promesses, sous-dimensionne les plannings,  
laisse de côté les tests — et la facture tombe des mois plus tard, sous forme de régressions, de  
retards et de budgets explosés.

Cette section reprend les idées reçues les plus répandues et les confronte à la réalité. L'objectif  
n'est pas de **décourager** — migrer est tout à fait faisable — mais de partir avec des **attentes  
justes**. C'est la meilleure protection contre les mauvaises surprises.

> 💡 **Le test du bon état d'esprit.** Si, à la fin de cette section, la migration vous paraît
> *un peu plus exigeante* que vous ne le pensiez mais *tout à fait atteignable avec méthode*, alors
> vous abordez le projet dans les bonnes dispositions.

---

## ❌ Mythe n°1 — « L'assistant de migration fait (presque) tout »

C'est le mythe le plus coûteux, parce qu'il **dicte le planning** : on prévoit « quelques jours  
d'outil », et on découvre des mois de travail manuel.

**La réalité.** Aucun outil ne migre **tout** automatiquement. Les assistants — l'*Upgrade Wizard*
historique de Visual Studio (intégré jusqu'à VS 2008, retiré depuis) et, aujourd'hui, les outils  
commerciaux comme **VBUC** ou **VB Migration Partner** — abattent **60 à 90 %** du travail
**mécanique** (la conversion de syntaxe répétitive, structurelle). Mais **le reste — souvent le plus
difficile — relève du jugement humain** :

- les **changements silencieux** (ByRef, entiers redimensionnés, finalisation) qu'aucun outil ne peut deviner à votre place, parce qu'ils dépendent de l'**intention** du code ;
- l'**interface** (les *control arrays* disparus, la géométrie twips→pixels, le redécoupage des formulaires) ;
- les **choix de conception** (un `Recordset` devient-il un `DataReader` ou un `DataSet` ? faut-il un `Using` ici ?) ;
- l'**élimination des béquilles** que les outils laissent derrière eux (voir mythe n°5).

> ⚠️ **Le piège dans le piège.** Un outil produit du code **qui compile** — ce qui donne une
> impression trompeuse de « presque fini ». Or compiler n'est **pas** se comporter pareil (rappel
> du 1.3). Le code converti automatiquement est un **point de départ à valider**, pas un produit
> fini. *Le plus dur commence souvent là où l'outil s'arrête.*

👉 **À creuser** : chapitre 4 (outils de migration et leurs **limites**), et chapitre 19 (l'IA, qui
n'échappe pas à cette règle).

---

## ❌ Mythe n°2 — « Ce n'est que de la syntaxe »

L'idée que VB6 et VB.NET, c'est « le même langage en plus moderne », et qu'il suffit donc de
« traduire » mot à mot.

**La réalité.** C'est le thème central du **1.3** : migrer, c'est changer **trois choses à la fois**
— le **langage**, le **runtime** (VBVM → CLR) et le **modèle objet** (dont la finalisation
déterministe perdue). La syntaxe n'est que la **partie visible** ; en dessous, le **comportement à  
l'exécution** change. Réduire la migration à de la syntaxe, c'est ignorer pile la catégorie de  
changements la **plus dangereuse** : les **silencieux**.

> ⚠️ Un `ByRef` devenu `ByVal`, un `Integer` qui passe de 16 à 32 bits, un `Class_Terminate` qui ne
> se déclenche plus au bon moment : **rien de tout cela n'est une affaire de syntaxe**, et **rien**
> ne sera signalé par le compilateur.

---

## ❌ Mythe n°3 — « Le code compile, donc la migration est finie »

Le corollaire des deux précédents, et **le plus dangereux de tous** : prendre la **compilation  
réussie** pour la **ligne d'arrivée**.

**La réalité.** « Ça compile » signifie seulement que le code est **syntaxiquement valide** — pas
qu'il **fait la même chose** qu'avant. Entre « compile » et « se comporte à l'identique », il y a  
tout le travail de **non-régression** : retrouver et corriger les changements silencieux, puis
**prouver** par des tests que le comportement est préservé.

> ⭐ **La vraie ligne d'arrivée**, ce n'est pas « le projet compile ». C'est « l'application **se
> comporte comme avant**, et on l'a **démontré** » — d'où l'importance capitale du *golden master*
> (5.5) et des tests de non-régression (17.3).

---

## ❌ Mythe n°4 — « Tant qu'à faire, autant tout réécrire »

À l'inverse du mythe « ce n'est que de la syntaxe », celui qui **surestime** l'ampleur du bon
chemin : puisqu'on touche au code, autant repartir d'une feuille blanche.

**La réalité.** Réécrire *from scratch* et **migrer** sont deux projets **différents**, aux risques
très différents. La réécriture jette **le comportement éprouvé** accumulé pendant quinze ou vingt  
ans — y compris les milliers de **règles métier implicites** et de **corrections de bugs** que plus  
personne ne sait reconstituer. C'est un risque majeur, et un classique des projets qui s'enlisent.

Cette formation fait le choix assumé de **migrer** un existant qui doit **continuer à tourner**, pas  
de le **redessiner**. Migrer **préserve le comportement** (et permet de le prouver) ; la réécriture  
recommence à zéro, sans filet.

> 💡 Migrer d'abord, **moderniser ensuite**. Une fois le code en VB.NET et sécurisé par des tests,
> le refactoring idiomatique (LINQ, génériques, espace `My`…) se fait **progressivement et sans
> risque** (chapitre 17.5) — pas en même temps que la migration.

---

## ❌ Mythe n°5 — « L'espace `Microsoft.VisualBasic.Compatibility` règle le problème »

La tentation de s'appuyer sur la couche de compatibilité (que les outils injectent volontiers) pour
« faire taire » les différences et avancer plus vite.

**La réalité.** Cet espace de noms est une **béquille de transition**, **pas une destination**. Il
permet à du code converti de fonctionner à court terme, mais il **fige** l'application dans un état  
hybride : dépendances héritées maintenues, code non idiomatique, et une dette qui **réapparaîtra**
(notamment lors d'un éventuel saut vers .NET moderne, où cette couche est en grande partie absente).

> ⚠️ S'appuyer durablement sur `Compatibility`, c'est **migrer à moitié** : on a changé de
> plateforme sans récolter les bénéfices, tout en conservant les contraintes. L'objectif est de
> **s'en débarrasser** (chapitre 17.2), pas de s'y installer.

👉 **À creuser** : section 4.2 (la béquille à éviter) et 17.2 (l'éliminer).

---

## ❌ Mythe n°6 — « L'IA va tout convertir toute seule »

La version 2026 du mythe n°1 : puisque l'IA écrit du code, elle devrait **migrer l'application** en  
autonomie.

**La réalité.** L'IA est une **aide réelle et précieuse** pour VB6 → VB.NET — convertir un module,  
**expliquer** un comportement VB6 obscur, suggérer des tests, repérer des pièges. Mais elle partage
la limite des outils automatiques, **plus une qui lui est propre** : elle peut produire des **faux
équivalents** VB6/VB.NET plausibles mais faux, ou **halluciner** des API. Et le risque est sournois
parce que sa sortie est **convaincante**.

> ⚠️ Toute production de l'IA doit être **validée systématiquement** (et il faut toujours lui
> préciser **« VB6 »** vs **« VB.NET »** et la cible **.NET Framework 4.7.2**, sous peine de
> confusions). L'IA **accélère** un migrateur compétent ; elle ne le **remplace** pas.

👉 **À creuser** : chapitre 19 entier, en particulier 19.5 (les pièges de l'IA).

---

## ✅ Mythes vs réalités — la synthèse

| Le mythe | La réalité |
|---|---|
| « L'assistant fait tout » | 60-90 % du **mécanique** ; le reste (silencieux, UI, conception) est **humain** |
| « Ce n'est que de la syntaxe » | **Trois** changements : langage **+** runtime **+** modèle objet |
| « Ça compile, c'est fini » | Compiler ≠ se comporter pareil ; la fin, c'est la **non-régression prouvée** |
| « Autant tout réécrire » | Réécrire jette le **comportement éprouvé** ; on **migre**, on modernise **ensuite** |
| « `Compatibility` règle tout » | **Béquille** de transition à **éliminer**, pas une destination |
| « L'IA convertit toute seule » | Aide réelle, mais **faux équivalents** et hallucinations → **validation systématique** |

Un fil rouge unique traverse ces six mythes : ils **sous-estiment les changements silencieux** et
**surestiment l'automatisation** (mythes 1, 2, 3, 5, 6), ou bien ils **confondent migrer et
réécrire** (mythe 4). Garder cette grille en tête, c'est désamorcer à l'avance la plupart des  
mauvaises surprises.

---

## ✅ À retenir

- Les **mythes coûtent plus cher que les difficultés techniques**, parce qu'ils faussent les décisions **en amont** (promesses, plannings, tests).
- **Aucun outil — ni l'IA — ne migre tout** : ils traitent le mécanique (60-90 %) ; le **jugement humain** traite le reste, à commencer par les **changements silencieux**.
- **« Ce n'est que de la syntaxe » est faux** : c'est langage **+** runtime **+** modèle objet.
- **« Ça compile » n'est pas « c'est fini »** : la vraie ligne d'arrivée est la **non-régression prouvée**.
- On **migre**, on ne **réécrit** pas : préserver le comportement éprouvé d'abord, **moderniser ensuite**.
- La couche **`Compatibility`** et l'**IA** sont des **aides de transition à encadrer**, pas des solutions clés en main.

---

## 🎓 Fin du chapitre 1 — ce que vous savez maintenant

Vous avez terminé le cadrage. À ce stade, vous êtes en mesure de :

- **justifier** la migration (1.1 : IDE hors support, runtime gelé, dette, recrutement, sécurité) ;
- **choisir la cible** et défendre **4.7.2 comme « pont »** qui minimise les changements simultanés (1.2) ;
- **comprendre ce qui change vraiment** — langage, runtime, GC, modèle objet — et distinguer le **visible** du **silencieux** (1.3) ;
- **cartographier les cinq fronts** d'une migration : langage, formulaires, données, COM, API (1.4) ;
- **aborder le projet avec des attentes justes**, à l'abri des mythes les plus coûteux (1.5).

La suite logique : passer du **pourquoi** au **quoi en détail**. Le **chapitre 2** plonge dans les
**différences fondamentales** VB6 ↔ VB.NET — à commencer par les deux runtimes et **le piège n°1**,
la finalisation déterministe perdue.

---

> 🧭 **Chapitre suivant** → [2. Différences fondamentales VB6 ↔ VB.NET](../02-differences-fondamentales/README.md)  
> 🔝 **Retour** → [Sommaire du chapitre 1](README.md)

⏭️ [Différences fondamentales VB6 ↔ VB.NET](/02-differences-fondamentales/README.md)
