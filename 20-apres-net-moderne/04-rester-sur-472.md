🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.4 Quand *rester* sur 4.7.2 (et c'est un choix légitime)

> 🧭 *Module 20 — Et après ? De 4.7.2 vers .NET moderne · Section 4 / 5*

La section 20.2 a pesé s'il fallait **sauter** ; la 20.3 a détaillé ce que sauter **coûte**. Cette section prend l'autre réponse **au sérieux** : **rester** sur 4.7.2 comme **décision délibérée et légitime** — ni un échec, ni une dette technique par défaut, ni un travail « inachevé ». Atteindre 4.7.2 est une **destination** : l'objectif de la migration (quitter VB6, rejoindre le monde **managé** de .NET, **comportement préservé**) est **déjà atteint**.

> ⚠️ **Le renversement de la charge de la preuve.** « **Aucun moteur concret** » est une **réponse complète**. On n'a pas à **justifier de rester** ; on a à **justifier de bouger** (rappel de la posture de 20.2). Sans bénéfice concret pour *cette* application, **rester est le défaut raisonnable**.

---

## Pourquoi rester n'est pas un échec

4.7.2 est une **vraie cible**, pas une salle d'attente ([rappel 1.2](../01-cadrage-migration/02-pourquoi-472.md)) : **runtime supporté** (lié au cycle de vie de Windows), **surface d'API la plus proche de VB6**, **meilleure compatibilité COM/ActiveX**, **parité du concepteur Windows Forms**. Le but de la migration est **rempli** dès ce palier.

Le mot « **dette** » mérite d'être pesé : un code n'est en dette que s'il **bloque un besoin** non satisfait. Une application **stable**, **supportée** et qui **remplit ses exigences** n'est pas en dette — elle est **terminée**. Confondre « pas la dernière version » avec « dette technique » est une **erreur de cadrage** : la nouveauté n'est pas une exigence, c'est une **option**.

---

## Les bonnes raisons de rester

Plusieurs situations font de « rester » le **bon** choix :

1. **Aucun moteur concret.** Pas de besoin multiplateforme, pas de dépendance *moderne-only*, pas d'écart de performance, pas d'échéance de support → il n'y a tout simplement **rien à gagner**. *C'est le cas le plus fréquent.*
2. **Application Windows-only et stable.** Logiciel de gestion sur poste, **Windows pour toujours**, mature, à faible taux d'évolution : les bénéfices du .NET moderne (conteneurs, Linux) **ne s'appliquent pas**.
3. **Dépendance bloquante sans équivalent moderne.** Un composant **COM/ActiveX/OCX** ou **tiers** critique sans chemin moderne ([dépendances 3.2](../03-evaluer-strategie/02-dependances.md)) : le bloquant qui rend le saut **impraticable** rend, du même coup, le maintien **rationnel**.
4. **Valeur nette négative.** Quand l'analyse coût/risque de la section 20.2 ressort **défavorable** pour cette application.
5. **Application en fin de vie ou en pur maintien.** Logiciel destiné au **remplacement**, à la **retraite**, ou en **maintenance seule** : un second chantier de migration serait un **gaspillage**.
6. **Ressources limitées, priorités ailleurs.** L'effort de l'équipe rend **plus** ailleurs (fonctionnalités, autres systèmes) : le second saut est un **coût d'opportunité**.
7. **Contraintes de certification / réglementaires.** Certains environnements **certifient** un runtime précis ; le coût de **re-certification** est élevé et la **stabilité** prime sur la **modernité**.

---

## Rester ≠ ne rien faire : rester *de façon responsable*

Voici la **nuance décisive**. Rester **légitimement** n'est pas **négliger**. Un maintien responsable suppose :

- **Surveiller l'horizon de support.** 4.7.2 est **lié au cycle de vie de Windows** : suivez ce cycle pour **ne pas être surpris**. « Rester » est une décision à **réexaminer**, pas un « oublier pour toujours ».
- **Maintenir la sécurité.** Continuez d'appliquer les mises à jour de sécurité de l'OS et du runtime ; 4.7.2 reste **servi en sécurité** via Windows.
- **Garder l'option ouverte.** Le travail déjà fait — *golden master*, [découplage UI/métier](../05-preparer-code-vb6/03-decoupler-ui-metier.md), [`Option Strict On`](../17-valider-refactoriser/01-option-strict.md), suppression des béquilles de compatibilité ([module 17](../17-valider-refactoriser/README.md)) — rend un **éventuel** second saut **moins cher**, *s'il devient un jour nécessaire*. **Bien rester, c'est rester prêt.**
- **Documenter la décision.** Consignez **pourquoi** vous restez **et** les **conditions** qui déclencheraient un réexamen. Une décision **datée**, assortie de **déclencheurs**, vaut mille « on verra ».

Ainsi compris, « rester » est un choix **actif, surveillé et documenté** — pas un **abandon**.

---

## Quand *ré-ouvrir* la question

La décision de rester est **provisoire** et **conditionnelle**. Quelques **déclencheurs** doivent la remettre sur la table :

- un **moteur concret apparaît** : une bibliothèque/API requise devient *moderne-only*, une exigence de **performance** émerge, une **échéance** de support/sécurité approche, un besoin **multiplateforme** surgit ;
- l'**horizon de support** de Windows/4.7.2 se **raccourcit** sensiblement ;
- une **dépendance bloquante** obtient enfin un **équivalent moderne** (le verrou saute) ;
- l'**équilibre équipe/écosystème** bascule (difficulté de recrutement sur Framework).

Tant qu'aucun de ces déclencheurs ne se présente, **rester reste la bonne réponse**.

---

## Tableau de synthèse

| Raison de rester | Ce que cela implique | Ce qui ré-ouvrirait la question |
|------------------|----------------------|---------------------------------|
| **aucun moteur concret** | rien à gagner aujourd'hui | un moteur concret apparaît |
| **Windows-only et stable** | bénéfices modernes sans objet | besoin multiplateforme / conteneurs |
| **bloquant COM / tiers** | saut impraticable | un équivalent moderne devient disponible |
| **valeur nette négative** | coût/risque non justifié | le rapport bénéfice/coût s'inverse |
| **fin de vie / faible évolution** | second chantier = gaspillage | prolongation imprévue de la durée de vie |
| **horizon de support OK** | surveiller le cycle Windows | l'horizon se raccourcit nettement |

---

## À retenir

**Rester sur 4.7.2 est un choix légitime, et souvent optimal.** L'objectif de la migration est **déjà atteint**, et « **aucun moteur concret** » suffit à **justifier le maintien** — c'est **bouger** qui doit se prouver. Mais restez **de façon responsable** : **surveillez** l'horizon de support, **maintenez** la sécurité, **restez prêt** (l'hygiène du module 17 garde l'option **bon marché**) et **documentez** la décision avec ses **déclencheurs** de réexamen. « Rester » est un choix **provisoire et surveillé**, **jamais** un abandon.

La section suivante referme le module et la formation : **si**, un jour, vous décidez d'aller plus loin, la **passerelle** vers la formation dédiée « VB.NET avec .NET 10 LTS ».

---

> 🧭 **Navigation**  
> ⬅️ [20.3 — Ce qui devra encore changer](03-ce-qui-change-encore.md)  
> ⬆️ [Sommaire de la formation](../SOMMAIRE.md)  
> ➡️ Suivant : [20.5 — Passerelle vers « VB.NET avec .NET 10 LTS »](05-passerelle-net10.md)

---

**Juin 2026** · Cible actuelle : .NET Framework 4.7.2 (le « pont » depuis VB6) · Suite optionnelle : .NET 10 LTS · Licence : Creative Commons BY-NC-SA 4.0

⏭️ [Passerelle vers la formation « **VB.NET avec .NET 10 LTS** »](/20-apres-net-moderne/05-passerelle-net10.md)
