🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16. Interopérabilité COM et appels d'API Windows 🔗 ⭐

**Faire cohabiter l'ancien et le nouveau : appeler du COM depuis .NET, exposer du .NET à VB6, et dialoguer directement avec Windows — sans réécrire ce qui n'a pas (encore) besoin de l'être.**

> 🔑 **Ce chapitre est la clé de voûte d'une migration incrémentale.** Tant que l'interopérabilité
> reste théorique, la seule stratégie possible est le *big-bang* : tout migrer d'un coup, et croiser
> les doigts. Dès qu'on maîtrise les ponts présentés ici, on peut migrer **morceau par morceau**, en
> laissant cohabiter du code VB6 et du code .NET dans une application qui continue de tourner en
> production. C'est ce que promettaient les modules 3.5 (cohabitation COM) et 18.2 (déploiement
> progressif) : voici comment ça marche concrètement.

---

## Où nous en sommes

Nous avons migré le **langage** (partie 3), l'**interface** (partie 4) et l'**accès aux données**
(module 15). Restent deux frontières que beaucoup d'applications VB6 franchissent en permanence,
souvent sans qu'on s'en rende compte :

1. **La frontière COM** — vos composants tiers (OCX, ActiveX, DLL ActiveX), vos propres bibliothèques
   VB6 compilées, le `Scripting.FileSystemObject`, des moteurs comme Crystal Reports… tout cela, c'est
   du **COM**. Une partie de votre code migré devra continuer à s'en servir, et — dans l'autre sens —
   votre code VB6 *pas encore migré* devra pouvoir appeler vos nouveaux composants .NET.
2. **La frontière du système** — chaque `Declare Function … Lib "user32"` est un appel direct à
   l'**API Win32**. Ces appels ne disparaissent pas : ils changent de forme.

Ce module traite ces deux frontières ensemble, parce qu'elles reposent sur le même socle technique
(le marshaling, voir plus bas) et soulèvent les mêmes pièges silencieux.

---

## Deux mondes réunis par un socle commun : COM

Rappel des modules 1 et 2 : VB6 s'exécute sur son propre moteur (la *VBVM*), .NET sur le **CLR**. Ce  
sont bien deux mondes — **mais ils partagent un terrain d'entente au niveau du système : COM**. VB6
*était* fondamentalement une technologie COM ; et .NET a été conçu, dès l'origine, avec une couche
d'interopérabilité COM mature et entièrement supportée, regroupée dans l'espace de noms
**`System.Runtime.InteropServices`** (`DllImport`, `Marshal`, `StructLayout`, `ComVisible`, etc.).

> ✅ **Ce n'est pas un bricolage ni une rustine.** Contrairement à l'espace
> `Microsoft.VisualBasic.Compatibility` (la béquille à éviter, vue au module 4.2),
> l'interopérabilité COM et le P/Invoke sont des mécanismes **officiels, documentés et pérennes** de
> la plateforme. On peut s'appuyer dessus sans honte — à condition de comprendre ce qu'ils coûtent.

Le CLR sait jouer les deux rôles : présenter un objet COM à du code .NET, **et** présenter un objet
.NET à du code COM. C'est ce double sens qui rend la migration progressive possible.

---

## Les deux sens de l'interop COM : RCW et CCW

Tout le chapitre repose sur cette distinction. Retenez-la avant tout le reste :

```
   VOTRE CODE .NET (CLR)                     CODE LEGACY COM (VB6, OCX, ActiveX)
   ─────────────────────                     ───────────────────────────────────

      .NET  ──►  [ RCW ]  ──►  COM      « J'appelle un composant COM existant
              Runtime Callable Wrapper       depuis mon nouveau code .NET »

      COM   ──►  [ CCW ]  ──►  .NET     « Mon vieux code VB6 appelle un nouveau
              COM Callable Wrapper           composant .NET, comme si c'était
                                             une DLL ActiveX classique »
```

- **RCW — *Runtime Callable Wrapper*** (module 16.1). Quand votre code .NET appelle un composant COM,
  le CLR fabrique un **proxy** (le RCW) qui traduit les appels .NET en appels COM. C'est ce qui vous
  permet de **réutiliser l'existant** : un OCX tiers, un contrôle ActiveX, une DLL VB6 que vous n'avez
  pas encore migrée. Concrètement, Visual Studio génère pour vous un *assembly d'interopérabilité*
  (les fameuses *Primary Interop Assemblies*, déjà évoquées au module 6.3).

- **CCW — *COM Callable Wrapper*** (module 16.2). C'est le sens **inverse**, et c'est **le pilier de
  la migration incrémentale** : vous réécrivez une classe en .NET, vous la rendez visible au COM
  (`<ComVisible(True)>`, enregistrement via `regasm`), et votre code VB6 *encore en place* l'appelle
  comme n'importe quel objet COM. Vous pouvez ainsi vider votre application VB6 de sa substance,
  classe après classe, sans jamais tout casser d'un coup.

| | **RCW** | **CCW** |
|---|---|---|
| **Sens de l'appel** | .NET → COM | COM → .NET |
| **Qui appelle qui** | Du neuf consomme du vieux | Du vieux consomme du neuf |
| **À quoi ça sert** | Réutiliser OCX / ActiveX / DLL VB6 | Migrer **par étapes** sans tout réécrire |
| **Outil clé** | `tlbimp` / référence COM dans VS | `regasm`, `<ComVisible>`, `tlbexp` |
| **Module** | 16.1 | 16.2 🔗 |

Une fois le pont établi dans les deux sens, reste la question du **liage** : faut-il connaître le type  
exact à la compilation (*early binding*) ou résoudre les appels à l'exécution (*late binding*, qui  
impose `Option Strict Off`) ? Le module **16.3** précise où le *late binding* reste nécessaire pour  
l'interop… et où il faut l'éliminer pour retrouver la sécurité de type promise par .NET (un objectif  
repris au module 17.1).

---

## Le second pilier : parler directement à Windows

Au-delà des composants COM, beaucoup de code VB6 appelle **l'API Win32** via l'instruction `Declare` :

- En VB6, on écrit `Declare Function … Lib "kernel32" Alias … (…) As Long`.
- En VB.NET, l'instruction `Declare` existe **toujours** (avec ses variantes `Ansi`/`Unicode`/`Auto`),
  mais la forme **idiomatique et recommandée** est l'attribut **`<DllImport>`** (P/Invoke), couvert au
  module **16.4**.

La traduction n'est presque jamais une simple recopie. Deux écueils majeurs attendent :

- Les **types ont changé de taille** (le piège du module 7.2 ressurgit ici) : un `Long` VB6 vaut
  32 bits, donc un paramètre d'API `Long` devient `Integer` en .NET, pas `Long`. Se tromper, c'est
  corrompre la pile d'appel.
- Les **chaînes et les structures** doivent être *marshalées* correctement — c'est le sujet du module
  **16.5**, et le fil rouge de tout le chapitre.

---

## Le fil rouge : le *marshaling* (et le piège ANSI / Unicode)

Le **marshaling**, c'est la conversion automatique des données entre le monde .NET (managé, Unicode,  
ramassé par le *garbage collector*) et le monde natif/COM (non managé, souvent ANSI, à durée de vie  
gérée à la main). Le CLR fait beaucoup de ce travail pour vous, mais **pas tout**, et ses valeurs par  
défaut ne correspondent pas toujours à ce que VB6 supposait.

Le cas d'école : **ANSI vs Unicode**. VB6 passait historiquement des chaînes **ANSI** aux API ; les  
chaînes .NET sont **Unicode (UTF-16)**. Si vous ne précisez pas le bon `CharSet` sur votre `DllImport`,  
vous obtenez des chaînes tronquées, vides ou illisibles — un grand classique des migrations ratées.

De même, les **structures** : un `Type … End Type` VB6 devient un `Structure … End Structure`, mais  
pour le passer à une API il faut généralement annoter sa disposition mémoire
(`<StructLayout(LayoutKind.Sequential)>`, parfois un `Pack` explicite, et des `<MarshalAs>` pour les
chaînes de longueur fixe ou les tableaux). Tout cela est détaillé au module **16.5**.

---

## Panorama des pièges de l'interopérabilité ⚠️

L'interop est puissante, mais c'est l'un des terrains les plus minés de toute la migration. Gardez ce  
panorama en tête ; chaque point est approfondi dans les sous-sections (et au catalogue de l'annexe B) :

- **Durée de vie des objets COM ≠ ramasse-miettes.** Le COM compte les références ; le CLR, lui,
  ramasse quand il veut (le problème de finalisation déterministe des modules 2.2 et 12.3). Un objet
  COM tenu par un RCW n'est libéré que lorsque le RCW est collecté — donc **pas de manière
  déterministe**. Pour des ressources sensibles (connexions, fichiers verrouillés, gros objets
  natifs), il faut souvent appeler `Marshal.ReleaseComObject` / `FinalReleaseComObject`, ou encadrer
  l'usage avec soin. Le réflexe « je ferme quand je veux » de VB6 ne tient plus.
- **Le piège ANSI / Unicode** sur les chaînes d'API (voir ci-dessus). ⚠️
- **ByRef / ByVal à la frontière.** Le changement de défaut (modules 2.4 et 10.2) est encore plus
  dangereux côté API : un paramètre passé dans le mauvais sens corrompt silencieusement la mémoire au
  lieu de lever une erreur propre.
- **Disposition et alignement des structures.** Mauvais `LayoutKind`, `Pack` oublié, tableau ou chaîne
  fixe mal marshalé → données décalées, plantages aléatoires.
- **Modèle de cloisonnement (*apartment* STA / MTA).** Beaucoup de composants COM (et toute l'UI
  Windows Forms) exigent un *thread* **STA** : le `[STAThread]` sur le point d'entrée n'est pas
  décoratif. Appeler du COM depuis le mauvais *thread* provoque des blocages ou des exceptions.
- **« Bitness » 32 / 64 bits.** Un composant COM ou un OCX **32 bits** ne peut être chargé que dans un
  processus 32 bits. Si vous dépendez d'un tel composant, vous devrez cibler **x86** (et non
  *AnyCPU*), avec toutes les conséquences que cela implique. C'est l'une des contraintes pratiques les
  plus fréquemment sous-estimées.
- **Enregistrement et déploiement.** RCW et CCW supposent un **enregistrement** (regsvr32, regasm,
  *manifests*…) : ce qui « marche sur ma machine » peut échouer en production. À relier directement au
  packaging du module 18.1.

---

## Une béquille **légitime** — mais une béquille

> ⚖️ **L'interop est un pont, pas une destination.** Chaque frontière COM/natif est un coût (perf,
> complexité, déploiement) et un risque (les pièges ci-dessus). C'est un outil **stratégique** —
> idéal pour *gagner du temps*, différer les réécritures risquées et garder l'application vivante
> pendant la transition — mais l'objectif à terme reste de **réduire** ces frontières, pas de les
> multiplier.

Autrement dit : on assume l'interop pour migrer sereinement, puis on en retire progressivement les  
dépendances quand le composant cible est lui-même migré (par exemple en remplaçant un OCX par un  
contrôle .NET natif, ou en éliminant un `Option Strict Off` devenu inutile). Cette logique de
**dette temporaire et assumée** est exactement celle de l'activation progressive d'`Option Strict`
(module 17.1) et de la suppression des dépendances de compatibilité (module 17.2).

---

## Feuille de route du chapitre (16.1 → 16.6)

- **16.1 — Réutiliser des composants COM/DLL existants depuis .NET (RCW).** Référencer un composant
  COM, comprendre l'assembly d'interop généré, appeler l'objet, et gérer proprement sa libération.
- **16.2 — Exposer du .NET à du VB6 (CCW). 🔗** Rendre une classe .NET visible au COM, l'enregistrer,
  et l'appeler depuis du code VB6 — la mécanique concrète de la migration **incrémentale**.
- **16.3 — *Early binding* vs *late binding*.** Où le liage tardif (`Option Strict Off`) reste
  indispensable pour l'interop, et où il faut au contraire l'éliminer.
- **16.4 — Déclarations d'API Win32 : `Declare` → P/Invoke (`DllImport`). ⚠️** Traduire les appels
  système, avec la bonne correspondance de types (rappel du module 7.2).
- **16.5 — Marshaling, ANSI/Unicode, types et structures (`Type` → `Structure` + `StructLayout`). ⚠️**
  Le cœur technique : faire traverser correctement chaînes, structures et tableaux.
- **16.6 — L'objet `App`, le registre et `FileSystemObject` → `System.IO` / espace `My`.** Remplacer
  les services « runtime » de VB6 par leurs équivalents .NET (souvent **sans** interop du tout).

---

## Prérequis et liens utiles

Avant d'attaquer ce module, il est utile d'avoir vu :

- **Module 6.3** — référencer les composants COM/ActiveX (RCW, *Primary Interop Assemblies*) : la mise
  en place côté projet.
- **Module 7.2** — les entiers redimensionnés : indispensable pour traduire correctement les
  signatures d'API.
- **Modules 2.2 / 12.3** — la finalisation déterministe perdue et le *pattern* `IDisposable`/`Using` :
  on en a besoin pour libérer proprement les objets COM.
- **Modules 2.4 / 10.2** — ByRef → ByVal : critique à la frontière native.

Et en parallèle de ce chapitre, gardez ouverts :

- **Module 3.5** — l'approche par interopérabilité COM (la **stratégie** dont ce module est la mise en
  œuvre).
- **Module 14.5** — contrôles ActiveX/OCX : un cas d'interop spécifique à l'UI.
- **Module 18.2** — cohabitation VB6/.NET pendant le déploiement.
- **Annexe A** (correspondances de mots-clés), **Annexe D** (correspondance des types) et **Annexe G**
  (glossaire : COM, CLR, *marshaling*, RCW/CCW…).

---

## En résumé

- L'interopérabilité repose sur un socle commun aux deux mondes : **COM**, via l'espace
  `System.Runtime.InteropServices` — un mécanisme **officiel et pérenne**, pas une rustine.
- **Deux sens à ne jamais confondre** : **RCW** (.NET appelle du COM, pour *réutiliser* l'existant) et
  **CCW** (du COM appelle du .NET, la **clé de la migration incrémentale**).
- Un **second pilier** : les appels d'API Windows, où `Declare` cède la place à **P/Invoke
  (`DllImport`)**.
- Le **marshaling** est le fil rouge, avec le piège récurrent **ANSI / Unicode** et la disposition des
  structures.
- L'interop concentre des **pièges silencieux** redoutables : durée de vie des objets COM, ByRef,
  cloisonnement STA/MTA, **bitness 32/64**, enregistrement et déploiement.
- C'est une **béquille légitime mais temporaire** : on s'en sert pour migrer en sécurité, puis on en
  réduit progressivement les dépendances.

> ➡️ **Commençons par le sens le plus courant : appeler un composant COM existant depuis .NET grâce au
> RCW.** (Module 16.1)

---

🏷️ **Indicateurs du module** : 🔗 Interop · ⭐ Point clé · ⚠️ Pièges (16.4, 16.5)
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Réutiliser des composants COM/DLL existants depuis .NET (**RCW**)](/16-interop-com-api/01-rcw.md)
