🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.3 Ce qui devra *encore* changer (API retirées, WinForms .NET, espace `My` partiel…)

> 🧭 *Module 20 — Et après ? De 4.7.2 vers .NET moderne · Section 3 / 5*

La section 20.2 a pesé la **décision** ; celle-ci en détaille la **surface de changement**, pour que le **coût** et le **risque** deviennent concrets. Le message honnête : **l'essentiel est du re-ciblage mécanique** (projet, configuration, déploiement) que l'outillage prend en charge ; mais **quelques zones sont vraiment dures** (API retirées, interop, composants tiers), et **un vieil ennemi revient sous un nouveau visage** (la sensibilité à la culture).

> ⚠️ **Cadre.** Ceci est la surface du **second saut** — **distincte** de VB6 → VB.NET, et en général **plus petite**, mais réelle. La **liste exacte dépend de votre version cible et de votre application** : faites-la **énumérer** par les analyseurs officiels (*.NET Portability Analyzer*, *.NET Upgrade Assistant*) plutôt que de l'estimer de mémoire.

---

## 1. Le système de projet et le *build* — *mécanique*

Le projet passe du format **.NET Framework** (`.vbproj` ancien, `packages.config`) au **format SDK** moderne (`PackageReference`, fichier projet allégé), avec un **`TargetFramework`** du type `net8.0-windows` / `net10.0-windows`. C'est **largement mécanique** et **outillé** (*try-convert*, *Upgrade Assistant*), mais c'est le **point d'entrée obligé** de tout le reste.

---

## 2. Les API retirées ou non portées — *point dur*

Certaines API de .NET Framework **n'existent plus** ou n'ont **pas d'équivalent direct** sur .NET moderne. Les cas notables : **.NET Remoting** (retiré), **Code Access Security / CAS** (retiré), l'**isolation par `AppDomain`** (très limitée), une partie du modèle **`System.Configuration`**, et divers scénarios **`System.Drawing`** / **`System.Web`** (Web Forms) propres au monde Framework.

Il faut **distinguer deux situations** :

- **Restauré via un paquet de compatibilité** — beaucoup d'API Windows reviennent par le **Windows Compatibility Pack** (`Microsoft.Windows.Compatibility`) : c'est alors une simple **référence à ajouter**.
- **Réellement disparu** (Remoting, CAS, isolation `AppDomain`…) — là, il faut **repenser** la portion de code concernée.

Les analyseurs officiels **signalent** précisément ce qui, dans *votre* code, tombe dans l'une ou l'autre catégorie.

---

## 3. WinForms sur .NET moderne — *re-ciblage + vigilance UI*

Bonne nouvelle : **WinForms est supporté** sur .NET moderne (Windows-only), avec son **concepteur**. Mais c'est une **réimplémentation** : attendez-vous à des **différences subtiles** de rendu et de comportement, à un **modèle haute résolution (DPI)** différent par défaut, et au fait que **certains contrôles anciens ou tiers** peuvent ne pas se charger dans le concepteur ou à l'exécution. Le projet active l'UI via l'indicateur `<UseWindowsForms>`.

Conséquence pratique : la **même vigilance** que pour la migration des formulaires (modules 13-14, [Annexe C](../annexes/correspondance-controles/README.md)) s'applique — il faut **re-valider l'interface**, pas seulement recompiler.

---

## 4. L'espace `My` *partiel* — *audit ciblé*

L'espace **`My`** propre à VB (`My.Computer`, `My.Settings`, `My.Application`, `My.Resources`…) n'est que **partiellement** porté sur .NET moderne. Une **bonne partie fonctionne** (notamment pour les applications WinForms), mais **sans parité complète** : certains membres sont **absents** ou se **comportent différemment**. Tout usage de `My.*` doit donc être **audité** au cas par cas — ne présumez pas que ce qui marchait sur 4.7.2 marche à l'identique.

---

## 5. La configuration — *transition*

Le modèle **`app.config` / `System.Configuration`** cède la place, de façon idiomatique, à **`appsettings.json`** avec `Microsoft.Extensions.Configuration`. Une **continuité partielle** reste possible (`ConfigurationManager` est disponible via un paquet dédié) pour adoucir la transition, mais les **schémas d'accès** aux `connectionStrings` et `appSettings` **changent**. Mécanique pour l'essentiel, à condition de **recâbler** proprement la lecture de configuration.

---

## 6. Le déploiement — *modèle différent*

Le modèle de déploiement **diffère** : choix entre **dépendant du framework** et **autonome** (*self-contained*), **publication mono-fichier**, et **plus d'installation du runtime à l'échelle de la machine**. **ClickOnce** est de nouveau disponible sur .NET moderne. À la clé, un **re-packaging** de l'installeur et des prérequis (à rapprocher de la logique de [packaging du module 18](../18-deploiement-bascule/README.md)).

---

## 7. L'interop COM et ActiveX — *point dur, bloquant possible*

COM **fonctionne** toujours sur Windows en .NET moderne, mais **avec des différences** par rapport au monde Framework ([rappel module 16](../16-interop-com-api/README.md)) : plus de *Primary Interop Assemblies* dans le GAC de la même manière, **génération des assemblages d'interop** différente, indicateur **`<EnableComHosting>`** pour **exposer** du .NET en COM, et nuances de **COM sans inscription**. L'**hébergement d'ActiveX/OCX** dans WinForms passe en général, mais avec des **réserves**.

C'est, **avec les composants tiers**, l'endroit où se logent les **bloquants** les plus probables — d'où la consigne de la section 20.2 : **inventorier tôt** ([dépendances 3.2](../03-evaluer-strategie/02-dependances.md)).

---

## 8. Le retour d'un vieil ennemi : la sensibilité à la culture (ICU vs NLS)

Voici le **rappel** le plus important de cette section. Sur .NET moderne, la **globalisation** repose par défaut sur **ICU** (et non plus **NLS** comme sur .NET Framework). Conséquence : le **tri**, la **casse**, la **comparaison** de chaînes et le **formatage** peuvent **différer** de ce que faisait 4.7.2. Autrement dit, la **même classe de régression** que la sensibilité à la culture rencontrée au premier saut ([Annexe B.10](../annexes/pieges-silencieux/README.md)) **réapparaît**, sous une **forme nouvelle**.

À surveiller également : **`BinaryFormatter`** est **obsolète/retiré** (sérialisation, pour raisons de sécurité) — un point sensible si l'application s'en sert.

> ⚠️ **Le piège silencieux revient.** Comportements identiques à la compilation, **différents à l'exécution** : exactement le danger central du cours, dans un nouveau costume. La parade ne change pas — **le *golden master* rejugé** ([module 5.5](../05-preparer-code-vb6/05-golden-master.md), [Annexe F](../annexes/golden-master/README.md)) et la **relecture via l'Annexe B**, appliqués **au second saut** comme au premier.

---

## Checklist de surface

| Domaine | Nature du changement | Difficulté | Repère |
|---------|----------------------|------------|--------|
| **Projet / build** | format SDK, `TargetFramework`, `PackageReference` | mécanique | *Upgrade Assistant* |
| **API retirées** | Remoting/CAS/`AppDomain`… ; *compat pack* pour le reste | **dure** (selon usage) | *Portability Analyzer* |
| **WinForms** | réimplémenté, DPI, contrôles | moyenne + **vigilance UI** | re-validation UI (Annexe C) |
| **Espace `My`** | partiel | moyenne (**audit** de chaque `My.*`) | audit ciblé |
| **Configuration** | `app.config` → `appsettings.json` | mécanique/moyenne | `ConfigurationManager` (transition) |
| **Déploiement** | autonome / mono-fichier / ClickOnce | moyenne | re-packaging |
| **Interop COM / ActiveX** | différences, hébergement | **dure** (bloquant possible) | inventaire tôt (3.2) |
| **Tiers / NuGet** | cible .NET Standard / moderne requise | **dure** (selon dépendances) | coût (20.2) |
| **Culture (ICU/NLS), `BinaryFormatter`** | comportement *runtime* | **sournoise** | *golden master*, Annexe B |

---

## À retenir

L'**essentiel** de la surface est du **re-ciblage mécanique** — projet, configuration, déploiement — pris en charge par l'**outillage**. Les parties **réellement dures** sont les **API disparues**, l'**interop COM/ActiveX** et la **disponibilité des composants tiers** : exactement les **bloquants** que 20.2 demande d'**inventorier tôt**. Et le **piège de la culture revient** (ICU vs NLS) : la discipline **\*golden master\* + Annexe B** s'applique **aussi** au second saut. La **liste précise** dépend de votre **version cible** et de votre **application** — laissez les **analyseurs officiels** l'établir.

La section suivante prend l'autre versant de la décision : **quand rester sur 4.7.2** est le bon choix — et parfaitement légitime.

---

> 🧭 **Navigation**  
> ⬅️ [20.2 — Évaluer un second saut (intérêt, coût, risques)](02-evaluer-second-saut.md)  
> ⬆️ [Sommaire de la formation](../SOMMAIRE.md)  
> ➡️ Suivant : [20.4 — Quand rester sur 4.7.2 (et c'est légitime)](04-rester-sur-472.md)

---

**Juin 2026** · Cible actuelle : .NET Framework 4.7.2 (le « pont » depuis VB6) · Suite optionnelle : .NET 10 LTS · Licence : Creative Commons BY-NC-SA 4.0

⏭️ [Quand **rester** sur 4.7.2 (et c'est un choix légitime)](/20-apres-net-moderne/04-rester-sur-472.md)
