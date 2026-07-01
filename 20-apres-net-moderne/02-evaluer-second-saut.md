🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.2 Évaluer un second saut vers .NET 8/10 (intérêt, coût, risques)

> 🧭 *Module 20 — Et après ? De 4.7.2 vers .NET moderne · Section 2 / 5*

La section 20.1 a établi que le second saut est **séparable** et **optionnel**. Cette section porte sur la **décision** : faut-il, ou non, l'entreprendre ? Ce n'est **pas** un « oui par défaut » — c'est un **jugement de valeur** qui pèse ce qu'on **gagne** (l'intérêt), ce que cela **coûte**, et ce que cela **risque**. Le cadre honnête : le second saut doit **mériter sa place**.

> ⚠️ **La posture de décision.** « Plus récent » **n'est pas un bénéfice en soi**. On ne migre **pas** vers le .NET moderne parce qu'il est moderne, mais **si et seulement si** ses avantages, **pour votre application**, l'emportent sur le coût et le risque. Le défaut raisonnable est **rester** (section 20.4), sauf **moteur concret** justifiant le mouvement.

---

## Axe 1 — L'intérêt : qu'est-ce qu'on gagne ?

Le .NET moderne (8/10) offre des avantages **réels**, mais **conditionnels** : chacun ne compte que s'il **s'applique à votre application**.

- **Performance** — gains réels de débit, de démarrage et de mémoire ; **décisifs** pour certaines charges, **indifférents** pour beaucoup d'applications de gestion sur poste.
- **Multiplateforme et conteneurs** — déploiement Linux/conteneurs ; utile **uniquement** si vous en avez **besoin** : nombre d'applications WinForms d'entreprise resteront **Windows-only** indéfiniment.
- **Trajectoire de support et de sécurité** — le .NET moderne propose des versions **LTS** (p. ex. .NET 10 LTS) avec support continu. 4.7.2 reste **supporté** (lié au cycle de vie de Windows) mais en mode **stable, sans nouveautés** ; l'horizon de support peut être un moteur **légitime**.
- **Écosystème et bibliothèques** — de plus en plus de paquets NuGet et de bibliothèques **récentes** ne ciblent que le .NET moderne ; certaines capacités **n'existent** que là.
- **Langage et runtime** — nouvelles fonctionnalités, améliorations (async, etc.) — à ne retenir que si vous les **utiliserez**.
- **Recrutement et compétences** — c'est la direction où vont l'écosystème et les nouveaux développeurs.

> 💡 La bonne question n'est pas « ces avantages existent-ils ? » (oui) mais **« lesquels s'appliquent vraiment à *cette* application ? »**. Un gain qui ne sert pas votre cas **n'est pas un gain**.

---

## Axe 2 — Le coût : qu'est-ce que ça demande ?

Le coût du second saut est **dominé par ce qui ne se porte pas proprement** — pas par le langage, qui reste **stable** d'un saut à l'autre.

- **Effort de portage** — **API retirées** ou modifiées à remplacer, **WinForms** à re-cibler, espace [`My`](../16-interop-com-api/06-app-registre-fso.md) **partiel** à contourner, conversion du **système de projet/SDK**, modèle de **configuration** différent (détail en [20.3](03-ce-qui-change-encore.md)).
- **Reprise de l'interop COM** — COM fonctionne toujours sur Windows en .NET moderne, mais **avec des différences** ; certains composants **ActiveX/OCX** ou tiers peuvent **ne pas avoir d'équivalent** moderne — un **point dur** potentiellement bloquant.
- **Composants tiers et dépendances** — Crystal Reports, OCX, contrôles commerciaux : disponibilité **variable** sur .NET moderne, coût de **remplacement** ou de **licence**.
- **Re-validation complète** — un **nouveau** *golden master* : rejouer toute la suite de non-régression, re-tester, re-déployer. Le second saut est **une migration à part entière**, avec sa **charge de validation**.
- **Gel et coût d'opportunité** — une **nouvelle** période de [gel du périmètre](../05-preparer-code-vb6/06-geler-perimetre.md) et une équipe détournée des évolutions fonctionnelles.
- **Formation** — montée en compétence de l'équipe sur le .NET moderne ([passerelle 20.5](05-passerelle-net10.md)).

> 💡 **Où se cache le coût :** dans l'**interop**, les **composants tiers** et les **API retirées** — c'est-à-dire dans ce qui **n'avait pas changé** au premier saut. C'est exactement là qu'il faut **chiffrer**.

---

## Axe 3 — Les risques : qu'est-ce qui peut mal tourner ?

- **Nouvelles régressions silencieuses** — un second lot de changements, donc une **seconde vague** de code qui « compile mais ne se comporte plus pareil » ; **plus petite et mieux isolée** que la première grâce au séquençage, mais bien réelle.
- **Dépendances bloquantes** — un composant **COM/ActiveX/tiers critique sans équivalent** moderne peut **bloquer** tout le saut. À **découvrir tôt** (inventaire, comme au [module 3](../03-evaluer-strategie/01-inventaire.md)).
- **Régressions d'interface** — WinForms réimplémenté : **différences subtiles** de rendu et de comportement des contrôles et du concepteur.
- **Risque projet** — dérive de coût/délai, et surtout la tentation de **« moderniser/réécrire tant qu'on y est »**. À **refuser** : on **migre**, on ne **redessine pas** — la règle de toute la formation.
- **Le risque de *ne pas* sauter** — rester comporte aussi ses risques (dérive de l'écosystème, questions de support à terme), **plus lents et plus petits**. La vraie question n'est pas « risque ou pas », mais **quel profil de risque** convient.

---

## Comment décider : une approche

La décision se construit, elle ne s'improvise pas. Une démarche en six temps :

1. **Identifier les moteurs concrets.** Existe-t-il une raison **réelle et nommée** ? (besoin Linux/conteneurs, bibliothèque/API requise **uniquement** sur .NET moderne, **échéance** de support/sécurité, exigence de **performance** non atteignable sur Framework.) **Aucun moteur concret → fort signal de rester.**
2. **Inventorier les bloquants tôt.** COM/ActiveX/tiers **sans chemin** moderne ([dépendances 3.2](../03-evaluer-strategie/02-dependances.md)). **Un seul** bloquant dur peut **trancher** la question.
3. **Chiffrer le coût sur les points durs** (interop, tiers, API retirées — [20.3](03-ce-qui-change-encore.md)), pas sur le langage ([estimer l'effort 3.3](../03-evaluer-strategie/03-estimer-effort.md)).
4. **Peser le bénéfice net** contre coût + risque, **pour cette application** — jamais dans l'abstrait.
5. **Décider par application / par lot**, pas en bloc : certaines avancent, d'autres restent — une logique **incrémentale**, comme les [stratégies 3.4](../03-evaluer-strategie/04-strategies.md).
6. **Si « go » : traiter comme une migration à part entière** — son [*golden master*](../05-preparer-code-vb6/05-golden-master.md), son périmètre **gelé**, toute la discipline du cours **réappliquée** au second saut.

**Grille de signaux :**

| Signaux qui plaident pour le second saut | Signaux qui plaident pour rester ([20.4](04-rester-sur-472.md)) |
|------------------------------------------|--------------------------------------|
| besoin réel **multiplateforme / conteneurs** | application **Windows-only**, stable |
| bibliothèque/API requise **uniquement** sur .NET moderne | **aucune** dépendance moderne requise |
| exigence de **performance** non atteignable sur Framework | performance **déjà suffisante** |
| **échéance** de support/sécurité concrète | runtime **supporté**, pas d'échéance proche |
| **pas de bloquant** COM/tiers (ou alternatives identifiées) | dépendance **COM/ActiveX/OCX sans équivalent** moderne |
| équipe **prête / formée** (ou plan de montée en compétence) | coût/risque de re-validation **non justifié** |

---

## À retenir

Le second saut doit **mériter sa place** par des **moteurs concrets**, pas par le réflexe « plus récent = mieux ». Le **coût** est dominé par ce qui **ne se porte pas** — interop, composants tiers, API retirées — et non par le langage. Le **risque** existe **des deux côtés** : il s'agit de choisir le **profil** qui convient. On décide **par application**, sur la **valeur nette** pour *votre* cas, et si l'on avance, on traite le saut comme **une migration complète** avec son propre **point de contrôle**.

La section suivante détaille, concrètement, **ce qui devra encore changer** si vous décidez de franchir ce second saut.

---

> 🧭 **Navigation**  
> ⬅️ [20.1 — Pourquoi 4.7.2 d'abord, deux sauts séparés](01-deux-sauts.md)  
> ⬆️ [Sommaire de la formation](../SOMMAIRE.md)  
> ➡️ Suivant : [20.3 — Ce qui devra encore changer](03-ce-qui-change-encore.md)

---

**Juin 2026** · Cible actuelle : .NET Framework 4.7.2 (le « pont » depuis VB6) · Suite optionnelle : .NET 10 LTS · Licence : Creative Commons BY-NC-SA 4.0

⏭️ [Ce qui devra **encore** changer (APIs retirées, WinForms .NET, espace `My` partiel…)](/20-apres-net-moderne/03-ce-qui-change-encore.md)
