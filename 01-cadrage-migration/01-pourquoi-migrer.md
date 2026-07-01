🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 1.1 Pourquoi migrer VB6 en 2026 ⭐

> **Chapitre 1 — Introduction : pourquoi et comment migrer**
> Première question de cadrage : *pourquoi* ne plus rester sur VB6 ?

---

## « Mon application tourne. Pourquoi y toucher ? »

C'est la question — parfaitement légitime — que pose toute équipe assise sur une application  
Visual Basic 6 qui fonctionne depuis quinze ou vingt ans. Et la réponse honnête commence par un  
constat tout aussi honnête : **oui, votre application VB6 tourne encore en 2026, et elle continuera  
sans doute de tourner demain matin.** Migrer n'est pas une réaction de panique face à un logiciel  
qui s'arrête.

Le vrai sujet n'est donc pas « ça marche ou pas ». C'est l'évolution du **risque dans le temps**.  
Rester sur VB6, c'est faire reposer une application critique sur un socle **figé**, dont le destin  
ne dépend plus de vous, et dont le coût de sortie **augmente chaque année**. Ce sont ces cinq  
pressions — l'IDE, le runtime, la dette technique, le recrutement et la sécurité — qui transforment  
peu à peu un confort apparent en risque réel.

---

## 🛑 1. L'environnement de développement n'est plus supporté

C'est le point le moins discutable : **l'IDE de Visual Basic 6.0 n'est plus pris en charge par  
Microsoft depuis le 8 avril 2008.** Depuis cette date, plus aucun correctif, plus aucune mise à  
jour, plus aucune garantie de compatibilité pour **l'outil avec lequel on écrit et maintient** le  
code.

Concrètement, cela signifie :

- **Aucun support officiel pour développer.** Il n'existe tout simplement plus de méthode supportée pour créer ou maintenir des applications VB6 — un point sur lequel Microsoft est on ne peut plus **explicite** (voir l'encadré ci-dessous).
- **Une installation fragile sur Windows récent.** Faire tourner l'IDE sur Windows 10 ou 11 relève de la débrouille (modes de compatibilité, manipulations non garanties) : cela fonctionne parfois, mais sans aucun filet. Si demain l'IDE refuse de s'installer après une mise à jour de Windows, personne ne corrigera quoi que ce soit.
- **Un outillage de l'an 2000.** Pas d'intégration native avec les gestionnaires de sources modernes, pas de frameworks de test actuels, pas de gestion de paquets, pas de chaîne d'intégration continue. L'écosystème autour de l'IDE est mort.

> 📄 **La position officielle de Microsoft**, dans sa déclaration de support de Visual Basic 6.0 :
> « Comme il n'existe aucune méthode supportée pour créer ou maintenir des applications Visual
> Basic 6, **Microsoft recommande fortement de remplacer vos applications par une technologie
> moderne.** » *(texte original : « Because there is no supported method to create or maintain
> Visual Basic 6 applications, Microsoft strongly recommends that you replace your applications with
> modern technology. »)*

> ⚠️ **À ne pas confondre.** « L'IDE n'est plus supporté » ne veut **pas** dire « les applications
> ne tournent plus ». Ce sont deux choses distinctes — et c'est précisément la nuance du point
> suivant. La distinction *développer* / *exécuter* est au cœur de la compréhension du sujet.

---

## ⚙️ 2. Le runtime : « ça tourne »… mais sous conditions

Voici la source de la plupart des malentendus. Beaucoup d'équipes concluent, à juste titre, que
« le runtime VB6 est supporté » — et en déduisent, à tort, qu'il n'y a aucune urgence. Regardons
ce que ce support recouvre **vraiment**.

Le **runtime** (le moteur d'exécution qui fait tourner les binaires VB6 compilés, `msvbvm60.dll` et  
ses fichiers de base) est **livré avec Windows** et **pris en charge pour la durée de vie des  
versions de Windows dans lesquelles il est fourni**. C'est la politique dite *« It Just Works »*
(IJW) : une application VB6 existante doit « simplement fonctionner » sur les versions de Windows
supportées, sans installation supplémentaire.

C'est une bonne nouvelle réelle. Mais elle est **encadrée par quatre limites importantes** :

- **Un support « gelé », pas un support « vivant ».** La prise en charge se limite aux **régressions graves** et aux **problèmes de sécurité critiques** des applications existantes. Aucune évolution, aucune nouvelle fonctionnalité, aucune amélioration : le runtime est figé en l'état.
- **Un destin lié à Windows, pas à vous.** Le runtime est supporté **tant que** la version de Windows qui l'embarque l'est. Ce n'est pas une garantie perpétuelle et indépendante : c'est une dépendance **dont la durée est décidée ailleurs**. Une évolution de la politique de Microsoft ou une future version de Windows peut déplacer le sol sous vos pieds — sans que vous ayez voix au chapitre.
- **« Runtime » ≠ produit complet.** Le moteur de base est couvert ; **l'IDE, lui, ne l'est pas** (voir point 1). Et entre les deux, il y a les **« fichiers étendus »** (*Runtime Extended Files*) — une sélection de contrôles ActiveX/OCX et d'outils : certains sont supportés, d'autres sont **sortis du support** ou n'ont **jamais** fait partie du redistribuable (notamment les contrôles tiers et les outils hérités de VB4/VB5). Beaucoup d'applications VB6 réelles dépendent justement de ces contrôles-là.
- **32 bits uniquement.** Les fichiers du runtime VB6 restent **exclusivement 32 bits**. Sur un Windows 64 bits, une application VB6 ne s'exécute que via la **couche de compatibilité WOW64** (*Windows 32-bit on Windows 64-bit*). Pas de version 64 bits native — ce qui ferme la porte à certains scénarios et intégrations qui exigent du 64 bits (chargement dans un processus 64 bits, DLL 64 bits, certains pilotes).

| | **Développer (IDE)** | **Exécuter (runtime de base)** |
|---|---|---|
| **Statut** | Hors support depuis avril 2008 | Pris en charge tant que le Windows hôte l'est |
| **Périmètre** | Aucun | Régressions graves + sécurité critique seulement |
| **Évolution** | Aucune | Aucune (figé) |
| **Qui décide de la fin ?** | Déjà terminé | Microsoft, via le cycle de vie de Windows |

> 💡 **Le bon résumé** : le runtime VB6 vous offre du **temps**, pas une **destination**. Il vous
> permet de migrer **sereinement plutôt que dans l'urgence** — mais il ne supprime pas le besoin de
> migrer.

---

## 🏗️ 3. La dette technique qui s'accumule

Même en mettant de côté le support, rester sur VB6 fige l'application sur un socle qui **s'éloigne  
un peu plus chaque année** du reste du monde logiciel.

- **Un langage arrêté en 1998.** Pas de généricité, pas de LINQ, pas d'héritage de classe complet, pas d'`async`, pas de gestion d'erreurs structurée native. Tout ce que l'écosystème .NET tient pour acquis depuis vingt ans manque à l'appel.
- **Une intégration de plus en plus laborieuse.** Services web REST/JSON, authentification moderne, cloud, API tierces récentes : chaque connexion au monde actuel demande des contournements coûteux depuis VB6.
- **Un effet « boule de neige ».** Plus on attend, plus la base de code grossit, plus l'écart avec la plateforme cible se creuse — et **plus la migration finale sera lourde**. La dette ne se stabilise pas toute seule : elle croît.

> ⚠️ Le paradoxe de l'attente : repousser la migration ne réduit pas son coût, **il l'augmente**.
> Le moment le moins cher pour migrer, c'est toujours « le plus tôt possible » — et le second
> meilleur moment, c'est maintenant.

---

## 👥 4. Le recrutement et la perte de compétences

C'est souvent le risque le plus **immédiat** et le plus sous-estimé — un risque humain avant d'être  
technique.

- **Une population de développeurs vieillissante.** Les développeurs maîtrisant VB6 partent progressivement en retraite. La relève n'existe pas : VB6 ne s'enseigne plus, et aucun nouvel arrivant ne l'apprend spontanément.
- **Un marché de recrutement à sec.** Trouver — et garder — quelqu'un capable de maintenir du VB6 devient difficile et coûteux. Les compétences se raréfient à mesure que la demande résiduelle persiste.
- **Le facteur « autobus ».** Dans beaucoup d'organisations, **une ou deux personnes seulement** comprennent réellement l'application historique. Leur départ (retraite, démission, imprévu) transforme une application qui « tournait très bien » en boîte noire que plus personne ne sait faire évoluer ni dépanner. La connaissance concentrée est un point de défaillance unique.

> 💡 Migrer, ce n'est pas seulement changer de technologie : c'est **rouvrir le vivier** des
> personnes capables de maintenir et faire évoluer l'application. Un code VB.NET sur .NET trouve
> des mainteneurs ; un code VB6, de moins en moins.

---

## 🔒 5. La sécurité et la conformité

Sur un système figé et relié au monde extérieur, l'écart de sécurité se creuse silencieusement.

- **Des pratiques d'une autre époque.** Les applications VB6 reposent fréquemment sur des schémas datés : algorithmes de chiffrement faibles, protocoles réseau anciens, identifiants en dur, accès aux données vulnérables aux injections. Les standards de sécurité actuels sont difficiles, voire impossibles, à appliquer proprement depuis VB6.
- **L'angle ActiveX/OCX.** Les contrôles ActiveX dont dépendent beaucoup d'applications VB6 constituent une **surface d'attaque** connue. Certains de ces composants ne sont plus maintenus du tout (voir point 2), et leurs vulnérabilités éventuelles ne seront jamais corrigées.
- **Un enjeu de conformité et d'audit.** Faire reposer une activité sur des **outils de développement non supportés** échoue de plus en plus souvent aux audits de sécurité et aux exigences réglementaires (RGPD, normes sectorielles, certifications). « Ça marche » ne suffit pas quand un auditeur demande qui corrigera la prochaine faille critique.

---

## ⚖️ Faut-il migrer dans l'urgence ?

**Non — et c'est important de le dire clairement.** Grâce au support du runtime, votre application
ne va pas s'éteindre du jour au lendemain. Céder à la panique conduit aux pires décisions
(réécritures précipitées, *big-bang* mal préparé). Ce n'est pas le message de cette formation.

Le message, c'est que **le risque est asymétrique et croissant** :

- ce qui rend la migration **plus chère** ne fait qu'augmenter (la base de code grossit, la dette s'accumule) ;
- ce qui la rend **possible** ne fait que diminuer (les experts partent, les composants sortent du support) ;
- et certains déclencheurs **ne dépendent pas de vous** (une décision de Microsoft, une nouvelle version de Windows, un incident de sécurité, le départ de la dernière personne qui connaît le code).

> ⭐ **L'idée à retenir.** Le bon moment pour migrer, c'est **tant que l'application tourne encore
> et que les compétences existent encore** — c'est-à-dire **depuis une position de force**, et non
> dans la crise une fois le déclencheur subi. Migrer par choix coûte bien moins cher, et fait
> courir bien moins de risques, que migrer sous contrainte.

C'est exactement ce que permet la stratégie défendue par cette formation : utiliser le **temps**  
offert par le runtime pour migrer **progressivement**, vers la cible qui **minimise les risques** —  
le sujet de la section suivante.

---

## ✅ À retenir

- **L'IDE VB6 est hors support depuis avril 2008** : développer et maintenir du VB6 ne bénéficie d'aucune garantie.
- **Le runtime, lui, est encore supporté** — mais c'est un support **gelé** (régressions graves et sécurité critique seulement), **lié à la durée de vie de Windows**, **32 bits uniquement**, et qui **ne couvre ni l'IDE ni tous les contrôles**. Il offre du temps, pas une destination.
- **La dette technique** s'accumule : plus on attend, plus la migration finale est lourde.
- **Le risque humain** (départ des experts VB6, recrutement impossible) est souvent le plus pressant.
- **La sécurité et la conformité** se dégradent sur un socle figé et relié au monde extérieur.
- Migrer n'est **pas une urgence**, mais un **risque croissant** : le bon moment est **tant que tout fonctionne encore**, depuis une position de force.

---

> 🧭 **Section suivante** → [1.2 Pourquoi cibler .NET Framework 4.7.2 : le « pont » naturel](02-pourquoi-472.md)

⏭️ [Pourquoi cibler .NET Framework 4.7.2 : le « pont » naturel](/01-cadrage-migration/02-pourquoi-472.md)
