🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.5 L'approche par interopérabilité COM : faire cohabiter VB6 et .NET 🔗

> **Module 3 — Évaluer l'existant et choisir une stratégie**
> La section 3.4 a présenté la migration **incrémentale** et désigné l'interopérabilité COM
> comme son **moteur technique**. Cette section explique *pourquoi* et *comment* VB6 et .NET
> peuvent coexister pendant la transition — au niveau **stratégique**. La mécanique détaillée
> (RCW, CCW, *binding*, *marshaling*) relève du **module 16**.

---

## 🧭 Pourquoi la cohabitation est la clé de l'incrémental

Reprenons le problème posé en 3.4. La migration incrémentale consiste à migrer l'application
**morceau par morceau**, en gardant à chaque étape une application qui **fonctionne**. Mais cela
soulève une question immédiate : comment une application peut-elle tourner alors qu'une **moitié**  
est en VB6 et l'**autre** en .NET ? Une application n'est pas une juxtaposition de morceaux  
indépendants : ses parties **s'appellent** mutuellement. Un formulaire migré en .NET a besoin  
d'appeler une fonction métier restée en VB6 — et inversement.

La réponse tient en un mot : **l'interopérabilité COM**. Parce que VB6 et .NET savent tous deux  
parler le langage de **COM** (*Component Object Model*), ils peuvent **s'appeler dans les deux  
sens** par-dessus la frontière. C'est ce pont qui rend la coexistence viable — et donc l'approche  
incrémentale possible.

> 💡 **Sans interopérabilité, l'incrémental n'existerait pas.** On ne pourrait pas faire tourner
> une moitié d'application : il faudrait tout migrer d'un coup (big-bang) ou ne rien migrer. COM
> est précisément ce qui autorise l'**état intermédiaire** — une application mi-VB6, mi-.NET, qui
> fonctionne le temps de la transition. C'est l'échafaudage qui permet de remplacer le bâtiment
> pièce par pièce sans l'effondrer.

Cette cohérence ne doit rien au hasard : c'est l'une des raisons du choix de **.NET Framework
4.7.2** comme cible (module 1.2). Sa **compatibilité COM** est la meilleure du paysage .NET — un
atout décisif pour une migration progressive.

---

## 🔁 Les deux directions de la frontière

L'interopérabilité COM fonctionne **dans les deux sens**, et chaque sens porte un nom. Comprendre  
cette symétrie est essentiel pour raisonner sur la cohabitation.

### Sens 1 — Du .NET appelle du COM (RCW)

Le code .NET nouvellement migré a besoin d'appeler du code **existant** : une DLL ActiveX maison  
restée en VB6, un contrôle OCX, une bibliothèque COM. Il le fait via un **RCW** (*Runtime Callable  
Wrapper*) : un objet .NET généré automatiquement qui **enveloppe** le composant COM et le présente  
au monde .NET comme s'il était natif.

C'est le sens « **réutiliser l'existant** » : il permet à la partie déjà migrée de continuer à  
s'appuyer sur des composants qu'on n'a pas (encore) migrés, ou qu'on ne migrera **jamais** (un  
composant tiers que l'on garde en COM — décision « réutiliser » de la cartographie 3.2). Le module
16.1 en détaille la mécanique.

### Sens 2 — Du VB6 appelle du .NET (CCW) — **la clé de l'incrémental**

C'est l'autre sens, et le plus important pour la migration progressive. Le code **VB6 encore en  
place** a besoin d'appeler du code **déjà migré** en .NET. Il le fait via un **CCW** (*COM Callable  
Wrapper*) : on **expose** un composant .NET *comme s'il était un composant COM*, de sorte que le  
VB6 puisse l'appeler exactement comme il appelait ses anciennes DLL ActiveX.

> 💡 **Le CCW est le mécanisme qui rend l'incrémental réellement possible.** Sans lui, on pourrait
> migrer le code qui *consomme* des composants, mais pas le remplacer **par-dessous** : dès qu'un
> morceau est migré en .NET, le VB6 restant doit pouvoir continuer à l'utiliser. Exposer du .NET au
> VB6, c'est permettre de **remplacer progressivement le cœur** d'une application tout en gardant
> l'enveloppe VB6 opérationnelle. Le module 16.2 lui est entièrement consacré.

### La symétrie en un coup d'œil

| Direction | Mécanisme | Rôle dans la migration | Module |
|-----------|-----------|------------------------|--------|
| **.NET → COM/VB6** | **RCW** (*Runtime Callable Wrapper*) | Réutiliser l'existant depuis le code migré | 16.1 |
| **VB6 → .NET** | **CCW** (*COM Callable Wrapper*) | Faire consommer le code migré par le VB6 restant — **clé de l'incrémental** | 16.2 |

C'est cette **bidirectionnalité** qui fait toute la richesse de la cohabitation : selon ce que l'on  
migre en premier (le cœur ou l'enveloppe), on s'appuiera tantôt sur l'un, tantôt sur l'autre — et  
souvent sur les deux à la fois.

---

## 🏗️ À quoi ressemble une application en cohabitation

Concrètement, pendant la transition, l'application n'est plus « une appli VB6 » ni « une appli
.NET » : c'est un **assemblage** des deux, articulé autour de la frontière COM. Plusieurs
configurations typiques :

- **Cœur migré, UI conservée** : la logique métier (souvent du `.bas` plus simple à migrer, module
  5.3) passe en .NET la première et est **exposée au VB6** (CCW) ; les formulaires VB6 restent en
  place et appellent ce nouveau cœur. On traite l'UI ensuite — souvent la partie la plus coûteuse
  (modules 13-14).
- **UI migrée, cœur conservé** : à l'inverse, on migre des formulaires en Windows Forms qui
  **appellent** la logique VB6 restée en COM (RCW), puis on migre le cœur progressivement.
- **Migration par sous-système** : un sous-ensemble fonctionnel entier (par exemple un module de
  facturation) passe en .NET et dialogue avec le reste de l'application VB6 via COM, dans un sens
  ou dans l'autre.

> 💡 L'ordre — cœur d'abord ou UI d'abord — n'est pas neutre, et rejoint les **approches hybrides**
> évoquées en 3.4. Migrer le cœur en premier a souvent du sens : la logique métier est plus simple
> à porter, se prête mieux aux tests de non-régression, et une fois exposée en CCW elle bénéficie à
> tout ce qui l'appelle. Mais le bon ordre dépend du découplage réel de l'application — d'où,
> encore, l'intérêt de **découpler l'UI de la logique métier en amont** (module 5.3).

---

## ✅ Ce que la cohabitation apporte

- **Elle rend l'incrémental viable.** C'est sa fonction première : permettre une application
  mi-VB6, mi-.NET qui **fonctionne** à chaque étape (3.4).
- **Elle fractionne le risque dans le temps.** On bascule par petits morceaux réversibles plutôt
  qu'en un saut unique — l'esprit même de la migration progressive et du déploiement par étapes
  (module 18.2).
- **Elle permet de garder durablement certains composants en COM.** Un composant tiers maintenu et
  sans équivalent .NET évident peut **rester** en interop — la décision « réutiliser » de la
  cartographie (3.2). Tout n'a pas vocation à être migré.
- **Elle découple le calendrier des dépendances.** Un point dur (3.3) peut être isolé derrière la
  frontière COM et traité **plus tard**, sans bloquer la migration du reste.
- **Elle autorise un apprentissage progressif.** L'équipe migre, valide et déploie des incréments,
  et **calibre** les suivants sur l'expérience des premiers.

---

## ⚠️ Ce qu'elle coûte — et ses limites

La cohabitation n'est **pas gratuite**. C'est un échafaudage **réel**, à construire et à maintenir,  
et il faut le mettre en regard du risque qu'il réduit.

- **Une complexité de construction et de maintenance.** Faire coexister deux mondes ajoute une
  couche à concevoir, déployer et faire évoluer. Tant que la frontière existe, elle doit être
  **entretenue**.
- **Un coût de déploiement accru.** Une application hybride embarque **deux runtimes** (le runtime
  VB6 et le CLR .NET) et impose l'**enregistrement** des composants COM des deux côtés (module
  18.2). Le packaging est plus lourd qu'une cible unique.
- **Un surcoût à chaque passage de frontière.** Tout appel qui traverse la limite VB6/.NET paie le
  prix du *marshaling* (conversion des types et des données d'un monde à l'autre, module 16.5). En
  général négligeable, mais à surveiller sur des frontières très sollicitées (boucles serrées,
  appels massifs).
- **Des pièges de *marshaling* à la frontière.** Le passage des types, des chaînes (ANSI/Unicode),
  des structures et des références d'un monde à l'autre est une **source classique de pièges
  silencieux** (module 16.5, Annexe B). La frontière est un endroit à tester avec soin.
- **La contrainte du 32 bits.** Beaucoup de composants COM/OCX hérités sont **32 bits** uniquement
  (signalé sur la carte 3.2) : la cohabitation force alors une compilation en `x86`, ce qui pèse
  sur l'architecture cible.

> ⚠️ **La cohabitation est un état de transition, pas une destination.** C'est le piège central de
> l'incrémental, déjà signalé en 3.4 : un échafaudage qu'on oublie de démonter devient une
> **architecture hybride permanente** — on cumule alors les coûts des deux mondes sans jamais
> récolter les bénéfices d'une migration achevée. La cohabitation se justifie **le temps de la
> transition** ; la mener à son **terme** (retirer la frontière, supprimer l'interop devenue
> inutile) fait partie intégrante du plan.

---

## 🔗 Où cette approche est traitée dans la formation

Cette section reste au niveau **stratégique** : elle explique *pourquoi* la cohabitation est le  
moteur de l'incrémental et *à quoi* elle ressemble. La suite de la formation en détaille la
**mécanique** et la **mise en œuvre** :

- **Module 16** 🔗 — Le cœur technique : **RCW** (16.1), **CCW** (16.2), *early/late binding*
  (16.3), et le *marshaling* à la frontière (16.5).
- **Module 6.3** — Référencer concrètement les composants COM/ActiveX nécessaires (RCW, *Primary
  Interop Assemblies*) dans l'environnement .NET.
- **Module 18.2** — La **cohabitation au déploiement** : faire tourner les deux runtimes côte à
  côte, enregistrement des composants, déploiement progressif.
- **Module 5.3** — En amont, **découpler l'UI de la logique métier** pour faciliter le découpage en
  morceaux migrables — un préalable qui conditionne la faisabilité de l'incrémental.

---

## ✅ En résumé

- L'**interopérabilité COM** est ce qui rend la migration **incrémentale** possible : elle permet
  à VB6 et .NET de **coexister** et de **s'appeler dans les deux sens** pendant la transition.
- La frontière fonctionne dans **deux directions** : le **RCW** (.NET appelle du COM/VB6 — réutiliser
  l'existant) et le **CCW** (VB6 appelle du .NET — **la clé de l'incrémental**, qui permet de
  remplacer le code par-dessous tout en gardant le VB6 opérationnel).
- Pendant la transition, l'application est un **assemblage** VB6 + .NET articulé autour de COM ;
  l'**ordre** de migration (cœur d'abord ou UI d'abord) rejoint les approches hybrides de 3.4 et
  dépend du découplage de l'application (module 5.3).
- La cohabitation **fractionne le risque**, **découple le calendrier des dépendances** et permet de
  **garder durablement** certains composants en COM (décision « réutiliser » de 3.2).
- Elle a un **coût réel** : complexité, déploiement de **deux runtimes**, *marshaling* à la
  frontière (et ses pièges silencieux), contrainte **32 bits**.
- C'est un **état de transition, pas une destination** : l'échafaudage doit être **démonté** une
  fois la migration aboutie, sous peine d'hybride permanent.
- La **mécanique détaillée** est traitée au **module 16** ; cette section en pose les **enjeux
  stratégiques**.

---

⬅️ Section 3.4 — [Stratégies : *big-bang* vs incrémentale](04-strategies.md)  
➡️ Section 3.6 — [Définir le périmètre, les critères de réussite et les indicateurs](06-perimetre-criteres.md)

---

**Juin 2026** · .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Définir le périmètre, les critères de réussite et les indicateurs](/03-evaluer-strategie/06-perimetre-criteres.md)
