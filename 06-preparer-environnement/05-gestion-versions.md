🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.5 Stratégie de gestion de versions (Git, branche de migration, jalons)

> *Partie 2 — Préparer le terrain · Module 6 — Préparer l'environnement .NET*

---

## Objectif de cette section

La gestion de versions est **l'infrastructure de sécurité** de la migration. C'est elle qui héberge la **ligne de base VB6 gelée** et la **référence *golden master*** (§ 5.5), qui rend possible le **retour en arrière**, qui permet de **comparer** précisément ce qui a changé, et qui **mesure l'avancement** d'un effort long via des jalons.

C'est la **dernière pierre** de la préparation de l'environnement : une fois posée, la migration du langage (Partie 3) peut commencer sur des fondations sûres.

> 💡 On parle ici de **Git**, devenu le standard. Si votre *legacy* est sous **Visual SourceSafe**, sur un partage réseau, ou sans gestion de versions du tout, la **première étape** est précisément d'en sortir (voir ci-dessous).

---

## D'abord : mettre le VB6 sous Git (la ligne de base)

Avant de migrer la moindre ligne, on place le **code VB6 existant** sous Git et on en fait la **ligne de base** — le point de référence figé.

C'est sur cette base, et **uniquement sur elle**, que :

- le **golden master** est capturé, sur l'application VB6 **intacte** (§ 5.5) ;
- le **périmètre est gelé** (§ 5.6) : `main` devient le VB6 figé, point de départ immuable.

> ⚠️ **Ordre impératif :** dépôt initialisé avec le VB6 → **jalon « ligne de base »** (un *tag*) → capture du golden master → puis seulement, démarrage de la migration. Migrer d'abord et versionner ensuite, c'est se priver du point de comparaison.

---

## Spécificités VB6 sous Git

Le code VB6 a quelques particularités à régler pour éviter un dépôt « bruité » :

- **Fins de ligne (CRLF).** Le code VB est en CRLF. On normalise via un fichier **`.gitattributes`** pour éviter que chaque outil ne réécrive les fins de ligne et ne pollue les *diffs*.
- **Fichiers binaires.** `.frm` est du texte, mais ses compagnons `.frx`/`.ctx` (ressources binaires), ainsi que `.res` et les OCX/DLL, sont **binaires** : pas de fusion ni de *diff* texte utile.
- **Fichiers à ignorer.** Positions de fenêtres (`*.vbw`, propres à l'utilisateur), cache du concepteur (`*.dca`), liaisons VSS (`*.scc`), journaux, et **sorties de compilation**.

```gitattributes
# .gitattributes — normaliser le texte, marquer le binaire
*           text=auto
*.frm       text
*.bas       text
*.cls       text
*.ctl       text
*.vb        text

*.frx       binary
*.ctx       binary
*.res       binary
*.ocx       binary
*.dll       binary
```

```gitignore
# VB6 — fichiers utilisateur / cache / liaisons
*.vbw
*.dca
*.scc
*.log
```

> ⚠️ **Sorties de compilation vs binaires tiers.** On ignore les `*.exe`/`*.dll`/`*.ocx` **que vous produisez**. En revanche, les **OCX/DLL tiers** dont dépend l'application ne sont pas des sorties de build : à **conserver délibérément** (dans un dossier dédié, ou via un mécanisme de dépendances documenté — voir § 6.3), pas à effacer d'un coup de `*.dll`.

---

## Spécificités .NET / Visual Studio sous Git

Côté solution .NET, on **versionne les sources** et on **ignore les artefacts générés** :

```gitignore
# Visual Studio / .NET
bin/
obj/
.vs/
*.user          # *.vbproj.user — réglages locaux
*.suo
packages/       # paquets NuGet restaurés
```

On **garde** : `.sln`/`.slnx`, `.vbproj`, fichiers `.vb` (dont les `*.Designer.vb`), `.resx`.

> 🔗 **Assemblies d'interop générés** (§ 6.3) : les commiter ou non est un **choix**. Les **commiter** rend le *build* reproductible **sans** dépendre de l'enregistrement COM sur la machine de build ; les **régénérer** garde le dépôt léger mais impose des composants COM enregistrés partout. À trancher en équipe, et à documenter.

---

## La branche de migration

C'est le cœur de la stratégie, et le terme figure dans le titre de la section.

L'idée : **isoler le travail de migration** sur une branche dédiée, pendant que `main` conserve le **VB6 gelé** — disponible pour les **exceptions** inévitables du gel (correctifs **légaux**, **sécurité**, **bugs P1** — § 5.6).

```
main (master)
  o──o──o───────────────────────────o     ← VB6 GELÉ (exceptions légales/P1 seulement)
   \                                /
    o──o──o──o──o──o──o──o──o──o──o─       ← branche « migration » (le travail .NET)
    │        │           │          │
  jalon    jalon       jalon      jalon
 baseline  compile    golden-OK   strict-on
```

Conséquences pratiques, en lien direct avec le **gel du périmètre (§ 5.6)** :

- **Une exception appliquée à `main`** (un changement légal sur le VB6) doit ensuite être **reportée** dans la branche de migration (cherry-pick / merge). C'est **l'impôt de la double maintenance** — d'où l'intérêt d'une migration **courte**.
- **La stratégie (§ 3.4) façonne le découpage :**
  - **Big-bang** : **une** branche de migration, fusionnée/basculée à la fin.
  - **Incrémentale** : souvent **une branche par module** fusionnée dans une branche d'intégration, au fil de l'avancement.
- **Branche ou dépôt séparé ?** Un **dépôt .NET distinct** est parfois préféré pour un big-bang (séparation nette) ; une **branche dans le même dépôt** facilite la cohabitation VB6/.NET en incrémental. Les deux sont valides — à choisir selon la stratégie.

---

## Les jalons (*milestones*)

Sur un effort long, on balise les **états significatifs** par des *tags*, qui servent à la fois de **points de retour** et de **structure de plan** :

| Jalon (*tag*) | Signification |
|---|---|
| `baseline-vb6` | VB6 **gelé**, golden master capturé (§ 5.5/5.6) |
| `env-ready` | Environnement .NET configuré : la solution existe et compile (module 6) |
| `migration-compiles` | Le code migré **compile** en VB.NET (encore en `Option Strict Off`) |
| `golden-master-ok` | **Parité comportementale** validée contre la référence (§ 17.3) — le grand jalon |
| `strict-on-complete` | `Option Strict On` **partout**, liaison tardive résiduelle éliminée (§ 17.1) |

En migration **incrémentale**, on ajoute des jalons **par module** (« module *Facturation* migré et validé »). Les jalons s'**alignent** sur la stratégie (module 3) et sur les **portes de validation** (golden master, module 17).

> 💡 « Compile » et « se comporte pareil » sont **deux jalons distincts**, et c'est volontaire : rappelons que *le vrai danger n'est pas le code qui ne compile pas, c'est le code qui compile mais ne se comporte plus pareil*. Le jalon `golden-master-ok` est le plus important des deux.

---

## Discipline de commits

La gestion de versions ne protège que si les commits sont **lisibles** :

- **Petits commits ciblés**, à message clair. Indispensable pour garder **séparables** les changements **mécaniques** de migration et les (rares) **changements de comportement** assumés.
- **Ne pas mélanger**, dans un même commit, une conversion mécanique et une correction fonctionnelle exceptionnelle.
- **Un commit qui met à jour la référence *golden master*** doit être **explicitement libellé et justifié** (le « pourquoi » dans le message) — c'est la mécanique de « bénédiction » d'un écart assumé (§ 5.5) et le **processus d'exception** du gel (§ 5.6).

> 💡 Comme les réglages du `.vbproj` sont versionnés (cible 4.7.2 du § 6.1, options du § 6.4), l'historique Git **montre noir sur blanc** le moment où `Option Strict` bascule de `Off` à `On`, projet ou fichier par fichier.

---

## Gérer les deux mondes (migration incrémentale)

En cohabitation VB6/.NET via interop COM (§ 3.5, § 18.2), les deux bases vivent **en parallèle** pendant la transition. La gestion de versions doit alors :

- **organiser** ce qui relève du VB6 et ce qui relève du .NET, et comment ils se **construisent ensemble** ;
- **gérer l'impôt de double maintenance** : un correctif du gel qui doit atterrir **des deux côtés** se propage par *cherry-pick* / *merge* maîtrisé.

Plus la fenêtre de cohabitation est **courte**, moins cet impôt coûte — la discipline de versions et la **rapidité** de migration se renforcent (§ 5.6).

---

## Vers un *build* reproductible

La gestion de versions est la **fondation** d'un *build* reproductible — lui-même nécessaire pour **exécuter automatiquement** le golden master (§ 17.3) et comparer à chaque étape.

Deux contraintes héritées des sections précédentes s'invitent ici :

- la **bitness** (cible **x86** si l'on consomme du COM 32 bits — § 6.3), à reproduire sur les agents de build ;
- l'**enregistrement COM** des composants sur la machine de build (sauf assemblies d'interop commités, voir plus haut).

> 🔗 L'automatisation complète (intégration continue, packaging, déploiement) relève des modules de **finalisation** (§ 18). À ce stade, retenez simplement que **versionnement + build reproductible** sont le socle sur lequel tout le reste s'appuie.

---

## Points de vigilance

> ⚠️ **Versionner le VB6 d'abord, baliser la ligne de base, puis migrer.** Sans baseline figée, pas de point de comparaison ni de golden master valide.

> ⚠️ **`main` = VB6 gelé ; le travail vit sur la branche de migration.** Les exceptions du gel (légal/P1) appliquées à `main` se **reportent** sur la branche — c'est la double maintenance, à minimiser par une migration courte.

> 💡 **Commits petits et lisibles ; référence golden master mise à jour = commit justifié.** Garder séparables migration mécanique et changements de comportement assumés.

> 💡 **Deux jalons distincts : « compile » et « se comporte pareil ».** Le second (`golden-master-ok`) est le vrai objectif.

---

## À retenir

> - La gestion de versions est **l'infrastructure de sécurité** de la migration : elle héberge la **ligne de base VB6 gelée** et la **référence golden master**, et permet retour arrière, comparaison et suivi.
> - **Étape zéro** : mettre le **VB6 sous Git**, **baliser** la ligne de base, **puis** capturer le golden master et migrer.
> - **Configurer le dépôt** : `.gitattributes` (CRLF, binaires `.frx`/`.ctx`/`.res`), `.gitignore` (VB6 : `*.vbw`/`*.dca`/`*.scc` ; .NET : `bin/`/`obj/`/`.vs/`/`*.user`).
> - **Branche de migration** : `main` garde le **VB6 gelé** (exceptions seulement) ; le travail .NET se fait à part. La **stratégie** (big-bang/incrémentale, module 3) dicte le découpage.
> - **Jalons** : `baseline-vb6` → `env-ready` → `migration-compiles` → **`golden-master-ok`** → `strict-on-complete` ; par module en incrémental. « Compile » et « se comporte pareil » sont **deux** jalons.
> - **Discipline de commits** : petits, lisibles, et tout **changement de comportement** (mise à jour de la référence) clairement justifié (§ 5.5/5.6).

---

## Liens utiles

- **§ 5.5** — *Golden master* *(la référence hébergée et figée dans le dépôt)*
- **§ 5.6** — Geler le périmètre *(`main` gelé, processus d'exception, double maintenance)*
- **§ 6.1 / 6.4** — Ciblage et options de compilation *(réglages `.vbproj` visibles dans l'historique)*
- **§ 6.3** — Référencer les composants COM/ActiveX *(commiter ou régénérer les assemblies d'interop)*
- **§ 3.4** — Stratégies : *big-bang* vs incrémentale *(façonne le branchement)*
- **§ 3.5 / 18.2** — Cohabitation VB6 / .NET *(gérer les deux mondes en versionnement)*
- **§ 17.1 / 17.3** — `Option Strict` progressif et non-régression *(jalons `strict-on` et `golden-master-ok`)*
- **§ 18** — Déploiement et bascule *(intégration continue, packaging — la suite)*

⏭️ [Types, variables et déclarations](/07-types-variables/README.md)
