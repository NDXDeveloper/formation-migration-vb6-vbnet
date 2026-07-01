🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.5 Passerelle vers la formation « VB.NET avec .NET 10 LTS » 🔗

> 🧭 *Module 20 — Et après ? De 4.7.2 vers .NET moderne · Section 5 / 5*

Voici le **hand-off**. Si vous avez décidé, via la [section 20.2](02-evaluer-second-saut.md), d'entreprendre le **second saut**, cette section vous oriente vers la **formation dédiée** qui le prend en charge — et clarifie la **frontière** entre les deux parcours. C'est aussi la **dernière section** de la présente formation : elle la **referme**.

---

## Où vous en êtes : deux formations, une trajectoire

Le séquençage en **deux sauts** ([20.1](01-deux-sauts.md)) se reflète dans **deux formations** complémentaires :

```text
VB6
  └─►  [ Cette formation : langage + runtime + modèle objet ]
        └─►  VB.NET sur .NET Framework 4.7.2     ◄── vous êtes ici
              └─►  Décision  (20.2 / 20.4)
                    ├─►  Rester sur 4.7.2        (choix légitime — 20.4)
                    └─►  Second saut  ─►  Formation « VB.NET avec .NET 10 LTS »
                                            └─►  VB.NET sur .NET 10 LTS
```

Vous avez fait le **plus difficile** : sortir de VB6 et atteindre le monde **managé** de .NET, **comportement préservé**, sur la cible la **plus proche de VB6**. La suite — *si* vous la prenez — est un **autre** voyage, **balisé** par la formation dédiée.

---

## Pourquoi une formation *distincte* (et pas un module de plus)

Tout le module 20 repose sur une thèse : **deux sauts séparés**, parce que grouper les changements **multiplie les risques**. Il aurait été **contradictoire** de traiter ici, en profondeur, le second saut. Celui-ci mobilise un **corps de connaissances différent** — runtime .NET moderne, projets SDK, globalisation **ICU**, WinForms réimplémenté, **API retirées**, modèle de déploiement, différences d'interop COM ([rappel 20.3](03-ce-qui-change-encore.md)). Il **mérite** donc sa **propre formation**, avec sa **profondeur**, son **outillage** et son **point de contrôle** au *golden master*.

---

## Ce que la suite « VB.NET avec .NET 10 LTS » aborde

La formation dédiée **reprend exactement là où celle-ci s'arrête** : sur une base VB.NET **stable et validée** en 4.7.2. Son périmètre **naturel** suit la surface de changement détaillée en [20.3](03-ce-qui-change-encore.md) :

- le **runtime .NET moderne** et le **format de projet SDK** ;
- les **API retirées** et le **Windows Compatibility Pack**, en profondeur ;
- **WinForms sur .NET moderne** (DPI, contrôles, concepteur) et la **re-validation de l'interface** ;
- l'**espace `My` partiel** : **audit** et contournements ;
- la **configuration** (`appsettings.json`) et le **déploiement moderne** (autonome, mono-fichier, ClickOnce) ;
- l'**interop COM/ActiveX** sur .NET moderne ;
- la **sensibilité à la culture (ICU vs NLS)** et les autres **comportements *runtime*** ;
- l'**outillage** du saut (*.NET Upgrade Assistant*, *.NET Portability Analyzer*).

> 🔗 Cette section est un **pointeur**, pas un substitut : elle situe la suite et prépare le passage de relais. Le **détail** appartient à la formation dédiée.

---

## Ce que vous emportez de cette formation

C'est le point le plus important du hand-off : **vous ne repartez pas de zéro**. La discipline et les actifs construits ici **se transfèrent directement** et rendent le second saut **plus sûr** et **moins cher** :

- **Le *golden master*** ([module 5.5](../05-preparer-code-vb6/05-golden-master.md)) — votre **référence de comportement** et votre **point de contrôle**, **réutilisé tel quel** pour la non-régression du second saut.
- **L'architecture découplée** ([UI/métier](../05-preparer-code-vb6/03-decoupler-ui-metier.md)) — facilite le re-ciblage WinForms et l'**isolation** des changements.
- **`Option Strict On` et la suppression des béquilles de compatibilité** ([module 17](../17-valider-refactoriser/README.md), [17.1](../17-valider-refactoriser/01-option-strict.md)) — un code VB.NET **propre** et **prêt** pour le moderne.
- **L'inventaire des dépendances** ([module 3.2](../03-evaluer-strategie/02-dependances.md)) — votre **carte** COM/ActiveX/tiers, qui alimente directement l'analyse des **bloquants** du second saut.
- **L'état d'esprit « pièges silencieux »** ([Annexe B](../annexes/pieges-silencieux/README.md)) — la **même vigilance** s'applique à ICU/NLS et au reste.
- **La méthode IA encadrée** ([module 19](../19-migration-ia/README.md)) — le **même** trio *prompting / détection / validation* se transpose au saut moderne.

**Bien rester, c'était rester prêt** ([20.4](04-rester-sur-472.md)) — et c'est exactement ce qui rend la **porte** de la formation suivante **bon marché à ouvrir**.

---

## Si vous restez : la passerelle attend

Ce relais **n'est pas une obligation**. Pour qui a choisi de **rester** sur 4.7.2 ([20.4](04-rester-sur-472.md)), cette passerelle reste **disponible** — elle s'active **le jour** où un **moteur concret** apparaît. **Garder l'option ouverte** ne coûte rien d'autre que la **veille** et l'**hygiène** déjà en place. La porte est là ; vous l'ouvrirez **si**, et **quand**, ce sera utile.

---

## Clôture de la formation

Vous arrivez au bout du **parcours principal**. Vous avez migré une **application vivante** — qui doit continuer à tourner — de VB6 vers VB.NET, en **préservant son comportement** : le **vrai** but, atteint.

Le **fil rouge** ne vous quittera plus, que vous restiez sur 4.7.2 ou que vous poursuiviez vers .NET 10 LTS :

- le **danger** n'est pas le code qui ne compile pas, c'est le code qui **compile mais ne se comporte plus pareil** ;
- le ***golden master*** est l'**arbitre** ;
- l'**outillage** et l'**IA** **accélèrent**, mais l'**humain** reste **responsable** ;
- 4.7.2 est le **pont**, et le saut moderne une **décision séparée et optionnelle**.

Quelle que soit la suite, vous l'aborderez **de la même façon** : **progressivement**, **validé**, **lucidement**. C'est tout ce que cette formation aura voulu transmettre.

> Pour la **référence** au quotidien, gardez ouvertes les [**annexes**](../SOMMAIRE.md) — en particulier l'**Annexe B** (pièges silencieux) et l'**Annexe F** ([*golden master*](../annexes/golden-master/README.md)).

---

> 🧭 **Navigation**  
> ⬅️ [20.4 — Quand rester sur 4.7.2 (et c'est légitime)](04-rester-sur-472.md)  
> ⬆️ [Sommaire de la formation](../SOMMAIRE.md)  
> 🏁 **Fin du parcours principal** · Suite optionnelle : formation dédiée **« VB.NET avec .NET 10 LTS »**

---

**Juin 2026** · Cible actuelle : .NET Framework 4.7.2 (le « pont » depuis VB6) · Suite optionnelle : .NET 10 LTS · Licence : Creative Commons BY-NC-SA 4.0

⏭️ [Tableau de correspondance VB6 → VB.NET](/annexes/correspondance-vb6-vbnet/README.md)
