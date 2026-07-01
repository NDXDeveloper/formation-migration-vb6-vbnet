🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.3 Estimer l'effort ⭐

> **Module 3 — Évaluer l'existant et choisir une stratégie**
> L'inventaire (3.1) a **recensé**, la cartographie (3.2) a **relié et décidé**. Cette section
> **chiffre** : combien de travail, et où se concentre-t-il ? L'objectif n'est pas de produire
> un nombre exact — c'est impossible — mais une **fourchette défendable**, fondée sur des
> mesures plutôt que sur une intuition.

---

## 🧭 Pourquoi l'intuition se trompe (presque toujours)

Demandez à une équipe d'estimer « à la louche » une migration VB6, et vous obtiendrez en général  
un chiffre **fondé sur la taille brute** : « c'est une grosse appli, comptez six mois ». Cette  
intuition est presque toujours fausse, et dans le même sens : elle **sous-estime**.

La raison est simple : l'effort de migration n'est **pas proportionnel au volume de code**. Deux  
applications de 50 000 lignes peuvent demander des efforts dans un rapport de un à dix selon ce  
qu'elles contiennent. Le code mécanique — celui que les outils convertissent à 60-90 % (module
4) — coûte peu. Ce qui coûte, c'est la **fraction résiduelle** : les pièges silencieux, les
dépendances sans équivalent, l'UI à repenser. Et cette fraction ne se répartit pas uniformément :  
elle se **concentre** dans des zones précises.

> ⚠️ **Le piège du chiffrage au volume.** « 142 formulaires × X heures » est une formule
> rassurante et trompeuse. Elle traite de la même façon une boîte de dialogue « OK/Annuler » et
> un formulaire de 40 contrôles avec dessin GDI, impression et *control arrays*. L'estimation
> sérieuse **pondère** le volume par la complexité et par la densité de points durs — exactement
> ce que les modules 3.1 et 3.2 ont permis de mesurer.

Estimer, c'est donc combiner **trois axes** — ceux du titre de cette section : la **taille**
(le volume), la **complexité** (la difficulté intrinsèque du code), et les **points durs** (les
multiplicateurs de risque). Aucun des trois ne suffit seul.

---

## 📏 Axe 1 — La taille (nécessaire, mais insuffisante)

La taille est le point de départ naturel : elle donne l'**ordre de grandeur** du chantier. Elle  
se mesure de plusieurs façons, déjà accessibles depuis l'inventaire (3.1) :

| Mesure de taille | Ce qu'elle indique | Limite |
|------------------|--------------------|--------|
| **Lignes de code (LOC)** | Volume brut à parcourir | Ne dit rien de la difficulté ; gonflée par le code mort |
| **Nombre de fichiers par type** | Profil du projet (UI / logique / objet) | Un `.frm` n'égale pas un autre `.frm` |
| **Nombre de formulaires et de contrôles par formulaire** | Charge d'interface (la plus coûteuse) | Compte mal le dessin, l'impression, les *control arrays* |
| **Nombre de procédures/fonctions** | Granularité de la logique | Ne distingue pas une fonction triviale d'un monstre de 500 lignes |

> 💡 **La taille est un dénominateur, pas un verdict.** Elle sert à rapporter les autres mesures
> à une échelle (« combien de `Declare` *pour 1000 lignes* ? », « combien de contrôles *par
> formulaire* ? »). C'est en croisant la taille avec la complexité et les points durs qu'elle
> devient utile.

Une précaution s'impose dès cet axe : **retrancher le code mort**. Une partie du volume inventorié  
n'est peut-être plus utilisée. Le chiffrer comme s'il devait être migré (et testé) gonfle  
artificiellement l'estimation — c'est précisément ce que le **nettoyage en amont** (module 5.1)  
viendra corriger.

---

## 🔀 Axe 2 — La complexité cyclomatique

La **complexité cyclomatique** mesure le nombre de **chemins d'exécution indépendants** dans un  
morceau de code. Concrètement, chaque point de décision — un `If`, un `ElseIf`, un `Case`, une  
boucle, un opérateur logique dans une condition — ajoute un chemin possible. Plus ce nombre est
élevé, plus le code est ramifié, et plus il est **difficile à comprendre, à migrer et à tester**.

Pourquoi cette mesure est-elle décisive pour une migration ?

- **Migrer du code complexe est plus lent et plus risqué.** Chaque chemin est une occasion
  d'introduire un changement de comportement involontaire. La complexité est donc directement
  corrélée à la probabilité de **pièges silencieux** (Annexe B).
- **Tester du code complexe coûte plus cher.** Le nombre de chemins est, en première approche, un
  **plancher du nombre de cas de test** nécessaires pour couvrir le comportement. Or la migration
  s'appuie entièrement sur les tests de non-régression (*golden master*, modules 5.5 et 17.3) :
  plus le code est complexe, plus le harnais de validation est lourd à construire.
- **La complexité localise l'effort.** Une poignée de procédures à très forte complexité peut
  concentrer l'essentiel du risque d'un module entier. Les repérer, c'est savoir où porter
  l'attention — et les revues.

> ⚠️ **Certaines constructions VB6 font exploser la complexité — ou la rendent illisible.**
> Le `On Error Resume Next` introduit un branchement **implicite** après *chaque* instruction
> (toute ligne peut échouer et l'exécution continue) : la complexité réelle dépasse alors de loin
> ce que le code laisse voir (module 11.3). Le `GoSub`/`Return` et le `On…GoTo` (module 9.5)
> produisent un flot de contrôle « spaghetti » que les outils refusent de convertir et qu'il faut
> **refactoriser à la main**. Ces zones sont à la fois les plus complexes *et* les plus pénibles
> à migrer : un double signal d'alerte.

En pratique, la complexité cyclomatique se mesure avec des outils d'analyse de code (certains  
outils commerciaux de migration la fournissent dans leur rapport d'évaluation — module 4.3). Le  
chiffre exact importe moins que la **distribution** : combien de procédures dépassent un seuil  
raisonnable ? Sont-elles concentrées dans quelques modules, ou disséminées partout ?

À titre de repère, les **seuils usuels** (d'après McCabe) hiérarchisent le risque :

| Complexité cyclomatique | Niveau | Lecture pour la migration |
|---|---|---|
| **1–10** | simple | Faible risque ; conversion outillée généralement fiable |
| **11–20** | modéré | À **relire** attentivement après conversion |
| **21–50** | élevé | **Zone rouge** : tests renforcés, revue humaine |
| **> 50** | très élevé | Quasi **intestable** en l'état ; refactoring souvent indispensable |

Ce sont des **repères, pas des couperets** — McCabe présentait lui-même la limite de 10 comme
« raisonnable, mais non magique ». Croisés avec le coût de test déjà évoqué (une procédure à 30
réclame *au moins* 30 cas pour couvrir tous ses chemins), ils indiquent **où** concentrer relecture  
et validation.

---

## 🧱 Axe 3 — Les points durs (les multiplicateurs)

Les **points durs** sont les éléments qui ne se convertissent pas mécaniquement et exigent du
**jugement humain** (ou de l'IA bien encadrée). Ce sont eux qui font la différence entre un
chiffrage qui tient et un chiffrage qui dérape. La bonne nouvelle : ils ont déjà été **localisés**  
par le balayage de code (3.1) et la cartographie des dépendances (3.2). Il ne reste qu'à les
**pondérer**.

| Point dur | Repéré en | Pourquoi c'est coûteux |
|-----------|-----------|------------------------|
| **OCX / composant tiers abandonné** | 3.2 | Remplacement (portage), pas conversion — coût élevé et incertain ⚠️ |
| ***Control arrays*** (`Index` sur contrôles) | 3.1 | Construction **disparue** en .NET → recréation manuelle (module 13.6) ⚠️ |
| **API Win32** (`Declare`) | 3.1 | Portage vers P/Invoke + marshaling ANSI/Unicode (module 16.4-16.5) ⚠️ |
| **`Class_Terminate`** (finalisation) | 3.1 | Le **piège n°1** : repenser la libération en `IDisposable`/`Using` (module 12.2-12.3) ⚠️ |
| **`On Error Resume Next`** | 3.1 | Masque des bugs ; migration délicate (module 11.3) ⚠️ |
| **Modules saturés de `Variant`** | 3.1 | Typage à clarifier au cas par cas (module 7.1) |
| **`Recordset` / `DataEnvironment`** | 3.2 | Changement de **paradigme** d'accès aux données (module 15.3) ⚠️ |
| **Crystal Reports / reporting** | 3.2 | Dépendance lourde, souvent sous-estimée (module 15.7) ⚠️ |
| **Dessin GDI / impression** sur formulaire | 3.1 | Réécriture vers GDI+ / `PrintDocument` (module 14.1, 14.4) ⚠️ |
| **`GoSub` / `On…GoTo`** | 3.1 | Constructions **supprimées** → refactoring manuel (module 9.5) ⚠️ |

> 💡 **Les points durs sont des multiplicateurs, pas des additions.** Un composant à **forte
> empreinte** (la colonne « utilisé par » de la carte 3.2) propage son coût à tout ce qui en
> dépend : remplacer une grille tierce câblée dans quarante écrans ne coûte pas « une fois » mais
> « quarante fois », plus la réécriture du composant lui-même. La densité et l'empreinte des
> points durs pèsent souvent plus lourd que le volume total de code.

---

## 🧮 Des trois axes à l'estimation

Aucun axe ne donne le chiffre à lui seul. L'estimation naît de leur **croisement** : un volume,  
pondéré par la complexité, majoré par la densité et l'empreinte des points durs. Une façon utile  
de raisonner consiste à **trier le code en strates de difficulté**.

| Strate | Nature | Effort relatif | Qui fait le travail |
|--------|--------|----------------|---------------------|
| **Vert** | Code simple, peu ramifié, sans dépendance externe (logique `.bas` pure) | Faible | Outils + relecture |
| **Orange** | Complexité moyenne, quelques pièges connus, dépendances réutilisables | Moyen | Outils + correction humaine |
| **Rouge** | Forte complexité, points durs, dépendances abandonnées, UI lourde | Élevé et **incertain** | Jugement humain (et IA encadrée) |

L'effort total n'est pas la somme uniforme des fichiers : c'est la **somme pondérée des strates**.  
Et l'expérience montre que la strate **rouge**, même minoritaire en volume, concentre souvent la
**majorité** de l'effort réel et **la quasi-totalité du risque**.

> 💡 **Le rappel honnête de la formation, chiffré.** Les outils traitent 60-90 % du travail
> *mécanique* — c'est essentiellement la strate verte et une partie de l'orange. Mais ce
> pourcentage porte sur le **volume**, pas sur l'**effort** : les 10-40 % restants (le rouge)
> peuvent représenter la majeure partie du temps et de la difficulté. Confondre « 80 % du code
> converti » avec « 80 % du travail fait » est l'une des illusions dénoncées au module 1.5.

---

## 📊 Estimer en fourchette, pas en point

Une estimation de migration honnête n'est **jamais un nombre unique**. La part d'incertitude — la  
strate rouge, les zones grises de la cartographie, les pièges qu'on ne découvrira qu'à l'exécution
— interdit la fausse précision. Plusieurs principes pour un chiffrage défendable :

- **Donner une fourchette** (basse / haute) plutôt qu'un point, et **assumer l'écart** : un écart
  large est plus honnête qu'un faux chiffre précis.
- **Attacher un niveau de confiance** à chaque zone : la strate verte se chiffre bien ; le rouge se
  chiffre mal, et c'est normal de le dire.
- **Réserver une marge pour les inconnues** : la cartographie (3.2) a laissé des zones grises (DLL
  sans documentation, composant non identifié). Tant qu'elles ne sont pas éclaircies, elles
  pèsent sur la borne haute.
- **Prévoir le coût des tests**, pas seulement de la conversion : construire le *golden master* et
  exécuter la non-régression (modules 5.5, 17.3) est un poste à part entière, proportionnel à la
  complexité mesurée à l'axe 2.
- **Ne pas oublier les postes annexes** : montée en compétences de l'équipe, mise en place de
  l'environnement (module 6), packaging et bascule (module 18), documentation (module 18.4).

> 💡 **Un outil commercial accélère le chiffrage initial.** VBUC (Mobilize.Net) ou VB Migration
> Partner (Code Architects) proposent une **phase d'évaluation** qui produit métriques de taille,
> de complexité et inventaire des éléments problématiques (module 4.3). C'est un excellent point
> de départ pour la fourchette — à **valider** par une lecture humaine, jamais à prendre pour un
> verdict.

---

## ⚠️ Les pièges de l'estimation

- **Chiffrer au volume brut.** Le piège central : traiter toutes les lignes comme égales. Le
  volume est un dénominateur, pas une estimation.
- **Croire le pourcentage d'automatisation.** « L'outil convertit 85 % » parle du code, pas de
  l'effort. Le reste peut être l'essentiel du travail (module 1.5).
- **Oublier la strate rouge dans l'estimation.** Sous-estimer ou ignorer les points durs est la
  première cause de dépassement. Ils sont peu volumineux mais déterminants.
- **Négliger le coût des tests.** Une migration sans non-régression n'est pas finie ; le harnais
  de tests est un poste majeur, croissant avec la complexité.
- **Annoncer un chiffre unique.** La fausse précision se retourne toujours contre celui qui l'a
  promise. La fourchette protège — et reflète la réalité.
- **Estimer une fois pour toutes.** L'estimation se **réactualise** à mesure que les zones grises
  s'éclaircissent et que la migration progresse. Le premier chiffrage est un cadrage, pas un
  contrat gravé.

---

## ✅ En résumé

- L'effort de migration **n'est pas proportionnel au volume de code** : l'intuition « au volume »
  sous-estime presque toujours.
- L'estimation croise **trois axes** : la **taille** (ordre de grandeur), la **complexité
  cyclomatique** (difficulté et coût de test), et les **points durs** (multiplicateurs de risque).
- La **taille** est un dénominateur utile mais insuffisant ; penser à en **retrancher le code
  mort**.
- La **complexité cyclomatique** mesure les chemins d'exécution : elle localise le risque de
  pièges silencieux et dimensionne le harnais de tests. Certaines constructions VB6 (`On Error
  Resume Next`, `GoSub`, `On…GoTo`) l'aggravent ou la masquent.
- Les **points durs**, déjà localisés en 3.1 et 3.2, sont des **multiplicateurs** : leur densité
  et leur **empreinte** pèsent souvent plus que le volume total.
- On chiffre en **strates de difficulté** (vert / orange / rouge) : la strate rouge, minoritaire
  en volume, concentre l'essentiel de l'effort et du risque.
- On exprime le résultat en **fourchette** assortie d'un niveau de confiance, on réserve une
  **marge** pour les zones grises, et on **réactualise** à mesure que l'on en sait plus.

---

⬅️ Section 3.2 — [Cartographie des dépendances](02-dependances.md)  
➡️ Section 3.4 — [Stratégies : *big-bang* vs incrémentale](04-strategies.md)

---

**Juin 2026** · .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Stratégies : *big-bang* vs incrémentale](/03-evaluer-strategie/04-strategies.md)
