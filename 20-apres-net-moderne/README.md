🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20. Et après ? De 4.7.2 vers .NET moderne (optionnel) ⭐

**Vous avez atteint .NET Framework 4.7.2. Et maintenant ? Pourquoi le saut vers le .NET moderne est un projet *séparé* et *optionnel* — et pourquoi *rester* est un choix légitime.**

> 🧭 *Partie 7 — IA, suite de parcours et ressources*
> Module de **cadrage et de décision**, destiné en priorité aux **décideurs** et **chefs de projet**.

---

## De quoi parle ce module

Vous avez fait le plus difficile : migrer de VB6 vers VB.NET sur **.NET Framework 4.7.2**. L'application **compile**, se **comporte comme l'originale** (le *golden master* le confirme) et tourne sur un **runtime supporté**. Vient alors la question inévitable : « **faut-il maintenant passer au .NET moderne (8/10) ?** »

La première réponse de ce module — et la plus importante — est un **recadrage** : atteindre 4.7.2 est une **destination légitime à part entière**, pas une demi-migration qu'on serait *obligé* de terminer. Le passage au .NET moderne est un **second projet, distinct et optionnel**, à entreprendre **plus tard** et **seulement s'il apporte une valeur concrète**.

C'est donc un module de **stratégie**, pas un module *pas à pas* : il aide à **évaluer, peser et décider**, pas à convertir du code.

> ⭐ **Le principe directeur : deux sauts, pas un.**
> VB6 → 4.7.2 (fait) et 4.7.2 → .NET moderne (à décider) sont **deux migrations distinctes**. Les enchaîner d'un bloc **multiplie les risques** — c'est exactement la logique qui a fait de 4.7.2 le **« pont »**. *(Le raisonnement détaillé : section 20.1.)*

---

## Pourquoi un module dédié

La **discipline de gestion des risques** qui a fait de 4.7.2 le pont — *minimiser le nombre de changements simultanés* — continue de gouverner ce qui vient **après**. Considérer le « .NET moderne » comme l'étape suivante **automatique** reviendrait à **défaire** cette discipline. Ce module existe pour que le second saut soit une **décision** prise en connaissance de cause, et non un **réflexe** — et pour **légitimer** l'option de s'arrêter à 4.7.2.

---

## Ce que vous allez apprendre

À l'issue de ce module, vous saurez :

- pourquoi **4.7.2 d'abord, .NET moderne ensuite** sont **deux sauts séparés**, et pourquoi l'**ordre** compte (**20.1**) ;
- **évaluer un second saut** vers .NET 8/10 selon l'**intérêt**, le **coût** et les **risques** (**20.2**) ;
- anticiper ce qui devra **encore** changer — **API retirées**, **WinForms** sur .NET moderne, espace **`My`** partiel… (**20.3**) ;
- reconnaître quand **rester sur 4.7.2** est le **bon choix**, et **légitime** (**20.4**) ;
- faire la **passerelle** vers la formation dédiée « **VB.NET avec .NET 10 LTS** » si vous décidez d'aller plus loin (**20.5**).

---

## La place de ce module dans la formation

Ce module **clôt la formation principale**. Il suppose la migration vers 4.7.2 **réalisée** (ou planifiée) et renvoie à deux fondations posées au début : [pourquoi 4.7.2 est le « pont »](../01-cadrage-migration/02-pourquoi-472.md) (section 1.2) et les [choix de stratégie](../03-evaluer-strategie/04-strategies.md) *big-bang* vs incrémentale (section 3.4). Pour les **décideurs** et **chefs de projet**, il prolonge naturellement l'épine dorsale de **cadrage** des [modules 1 à 4](../01-cadrage-migration/README.md).

---

## Plan du module

| § | Sujet | Repère |
|---|-------|--------|
| [20.1](01-deux-sauts.md) | **Pourquoi 4.7.2 d'abord, .NET moderne ensuite** — deux sauts séparés, et l'importance de l'ordre. | ⭐ |
| [20.2](02-evaluer-second-saut.md) | **Évaluer un second saut** vers .NET 8/10 — intérêt, coût, risques. | |
| [20.3](03-ce-qui-change-encore.md) | **Ce qui devra encore changer** — API retirées, WinForms .NET, espace `My` partiel… | |
| [20.4](04-rester-sur-472.md) | **Quand rester sur 4.7.2** — et pourquoi c'est un choix légitime. | |
| [20.5](05-passerelle-net10.md) | **Passerelle** vers la formation « VB.NET avec .NET 10 LTS ». | 🔗 |

---

## « Optionnel » n'est pas « inachevé »

Il faut le dire clairement : **rester sur 4.7.2 n'est pas de la dette technique par défaut**. C'est une cible **supportée**, à la **meilleure compatibilité COM/ActiveX**, à **parité du concepteur Windows Forms**, adaptée à du *legacy* d'entreprise. Pour beaucoup d'applications, c'est une **destination défendable et durable**. Ce module traite donc le choix de **rester** comme une **option de premier rang** (section 20.4), au même titre que celui d'avancer.

---

## À retenir avant de commencer

Atteindre 4.7.2 est une **destination**, pas une étape obligée vers autre chose. Le saut vers le .NET moderne est un **projet séparé et optionnel** — à **décider** sur l'intérêt, le coût et le risque, non à **présumer**. Ce module vous donne le **cadre** pour trancher **lucidement** : avancer, ou rester, en toute connaissance de cause.

---

> 🧭 **Navigation**  
> ⬅️ Précédent : [19. Migrer avec l'aide de l'IA](../19-migration-ia/README.md)  
> ⬆️ [Sommaire de la formation](../SOMMAIRE.md)  
> ➡️ Suivant : [20.1 — Pourquoi 4.7.2 d'abord, .NET moderne ensuite : deux sauts séparés](01-deux-sauts.md)

---

**Juin 2026** · Cible actuelle : .NET Framework 4.7.2 (le « pont » depuis VB6) · Suite optionnelle : .NET 10 LTS · Licence : Creative Commons BY-NC-SA 4.0

⏭️ [Pourquoi 4.7.2 d'abord, .NET moderne ensuite : **deux sauts séparés**](/20-apres-net-moderne/01-deux-sauts.md)
