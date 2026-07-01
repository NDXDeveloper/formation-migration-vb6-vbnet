🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.4 Migration manuelle assistée : quand et pourquoi

> **Module 4 — Outils de migration**
> Après l'assistant disparu (4.1) et les outils commerciaux (4.3), voici la dernière voie : la
> migration **manuelle**. Mais le mot est trompeur, et le titre le corrige d'emblée — il ne s'agit
> pas de retaper du code dans le Bloc-notes, mais d'une migration **assistée** par l'outillage
> moderne, où c'est le **jugement humain** qui mène. C'est la voie qui traite précisément ce que
> l'automatisation **ne sait pas** faire.

---

## 🧭 Lever d'emblée deux malentendus

Le terme « migration manuelle » prête à deux contresens qu'il faut dissiper avant tout.

**Ce n'est pas une migration sans outils.** Le titre dit « manuelle **assistée** » : le développeur
garde la main sur les décisions, mais s'appuie sur tout ce que l'environnement moderne met à sa  
disposition — l'IDE et ses outils de refactoring, l'**IA** (module 19), l'usage **sélectif** des  
outils commerciaux (4.3), les annexes de correspondance, et surtout le **harnais de tests de  
référence** (module 5.5). « Manuel » qualifie le **pilotage**, pas l'absence d'assistance.

**Ce n'est pas une réécriture *from scratch*.** C'est un point de principe de toute la formation :
on **migre un existant**, on ne le **redessine** pas. La migration manuelle consiste à **porter**  
le code VB6 avec discernement — transposer chaque construction vers son **vrai équivalent** .NET en  
préservant le comportement —, pas à repartir d'une page blanche. La nuance compte : redessiner, ce  
serait abandonner le comportement éprouvé du code existant et rouvrir tous les bugs déjà corrigés.

> 💡 Il existe une exception locale : certaines portions n'ont **pas** d'équivalent à transposer et
> doivent être **réécrites** — un *control array* disparu (module 13.6), un `GoSub` supprimé
> (module 9.5), un OCX tiers abandonné à remplacer (3.2). Mais c'est une réécriture **ciblée**,
> imposée par l'absence d'équivalent, pas un redesign de l'application. La règle reste : porter à
> l'identique partout où c'est possible.

---

## 🎯 Pourquoi : ce que seul le jugement humain sait faire

La raison d'être du manuel tient en une phrase : **il traite la strate « rouge »** (module 3.3) —  
cette fraction du code, minoritaire en volume mais majoritaire en effort, qui résiste à  
l'automatisation. Concrètement, voici ce que l'humain apporte là où l'outil échoue.

**Comprendre l'*intention*, pas seulement la syntaxe.** C'est le cœur du problème des **pièges
silencieux** (Annexe B). Un `ByRef` devenu `ByVal` n'est dangereux que si le passage par référence
était **voulu** — et seul quelqu'un qui connaît la **règle métier** peut le dire. Un outil voit du
code ; un humain voit une **intention**. Toute la traque des pièges (module 17.4) repose sur cette  
capacité à juger ce que le code était *censé* faire.

**Refactoriser ce qui a été *supprimé*.** Les constructions VB6 sans équivalent .NET — `GoSub`/`Return`,
`On…GoTo` (module 9.5), *control arrays* (13.6) — ne se **convertissent** pas : elles se
**repensent**. Transformer un flot « spaghetti » en code structuré demande une compréhension de la
logique, pas une règle de réécriture mécanique.

**Adopter le code *idiomatique*.** Un outil transcrit ; un humain **modernise**. Remplacer un
`Class_Terminate` par un vrai patron `IDisposable`/`Using` (module 12.3), exploiter LINQ, les
génériques, l'espace `My` (module 17.5) : ce passage du « VB6 déguisé en .NET » au « vrai .NET »  
relève du jugement.

**Traiter l'interface.** L'UI concentre souvent le travail le plus fin (twips → pixels, ancrage,
disposition, équivalences de contrôles — modules 13-14). L'œil humain y est irremplaçable : une  
conversion logique parfaite peut produire un formulaire visuellement décalé.

**Décider de ce qui relève du métier.** Qu'est-ce que du **code mort** (à ne pas migrer) ? Quel  
**comportement est critique** (à documenter et préserver — module 5.4) ? Que **garder**, que  
**laisser tomber** ? Ces décisions de **périmètre** (3.6) supposent une connaissance de
l'application qu'aucun outil ne possède.

---

## 📍 Quand : les situations qui appellent le manuel

La migration manuelle n'est pas un choix par défaut : elle s'impose (ou se justifie) dans des  
contextes précis.

- **Les petites applications.** Sur un projet modeste — typiquement le bon candidat au *big-bang*
  (3.4) —, mettre en place et apprendre un outil commercial peut coûter plus que migrer à la main.
  Le manuel est alors le choix **pragmatique**.
- **Les points durs, dans *tout* projet.** Même au sein d'une migration massivement outillée, les
  zones rouges identifiées en 3.3 (OCX abandonnés, *control arrays*, dessin GDI, `On Error Resume
  Next`…) **basculent en manuel**. Ce n'est pas « manuel **ou** outil » : c'est « outil pour le
  reste, manuel **pour les points durs** ».
- **Après une passe automatique.** Le travail résiduel d'un outil — chaque `UPGRADE_TODO`/`WARNING`
  de l'assistant (4.1), chaque `VB6.*` à supprimer (module 17.2), chaque avertissement à lever — est
  par nature **manuel**. Toute migration outillée **se termine** à la main.
- **L'exigence de qualité maximale.** Quand on vise un code **propre et idiomatique** dès la
  première écriture, le contrôle manuel garantit un résultat qu'aucune règle automatique ne donne
  exactement.
- **Le code trop particulier.** Une application gorgée de *late binding*, de constructions exotiques
  ou de patterns inhabituels peut **dérouter** les outils ; le manuel reprend alors la main là où
  l'automatisation patine.
- **La contrainte budgétaire.** Si l'achat d'un outil commercial n'est pas possible mais que
  l'**expertise** est disponible en interne, le manuel assisté (par l'IDE et l'IA) devient la voie
  raisonnable.

---

## 🧰 « Assistée » : par quoi, exactement ?

Tout l'enjeu est de ne **pas** confondre « manuel » et « à mains nues ». Le développeur dispose  
d'une batterie d'assistances qui rendent le manuel viable :

- **L'IDE moderne** (Visual Studio 2026) : refactoring, IntelliSense, navigation — et surtout les
  **avertissements d'`Option Strict`** (modules 6.4, 17.1), qui transforment le compilateur en
  **guide** signalant le *late binding* et les conversions hasardeuses à reprendre.
- **L'IA** (module 19) : pour **expliquer** un code VB6 obscur, **convertir** un module ou une
  classe, **détecter** des pièges (ByRef, `Variant`, finalisation), **générer** tests et
  documentation. Avec ses garde-fous : hallucinations et **faux équivalents** VB6/VB.NET à valider
  systématiquement (module 19.5).
- **Les outils commerciaux, en usage *sélectif*** : rien n'interdit de passer un outil (4.3) sur les
  80 % faciles, puis de traiter les 20 % durs à la main. C'est l'**approche hybride**, justement
  l'objet de la grille de décision suivante (4.5).
- **Les annexes de la formation** : la correspondance VB6 → VB.NET (Annexe A), le **catalogue des
  pièges** (Annexe B), la correspondance des types (Annexe D) — la documentation à garder ouverte.
- **Le *golden master*** (module 5.5) : le **filet de sécurité** sans lequel rien de ce qui précède
  ne tient. C'est lui qui rend chaque modification manuelle **vérifiable** par la non-régression
  (module 17.3).

> ⚠️ **Le manuel sans *golden master* est un pari aveugle.** C'est précisément en migration
> manuelle que le risque de **régression silencieuse** est le plus élevé : chaque ligne réécrite à
> la main peut introduire un écart de comportement invisible. Le harnais de tests de référence n'est
> pas une option ici — c'est la **condition** qui distingue une migration manuelle maîtrisée d'un
> bricolage dangereux.

---

## ⚖️ Manuel *vs* automatique : un faux dilemme

Il serait tentant de présenter ce module comme un choix entre deux camps — les outils d'un côté
(4.1, 4.3), le manuel de l'autre (4.4). **Ce n'est pas la bonne façon de voir.** La réalité d'un
projet est presque toujours un **dosage**, pas une alternative :

- même une migration **massivement outillée** se **termine** par du travail manuel (le résidu, les
  points durs, l'UI) ;
- même une migration dite **« manuelle »** s'appuie sur des **outils** (IDE, IA, voire un outil
  commercial sur les parties faciles).

La vraie question n'est donc pas « **outil ou manuel ?** » mais « **quel dosage ?** » — quelle part  
automatiser, quelle part traiter à la main, et **où placer la frontière**. C'est exactement ce que  
la **grille de décision** (section 4.5) aide à régler.

---

## ⚠️ Les exigences — et les limites — du manuel

Par lucidité (l'esprit du *repère honnête*), il faut nommer ce que le manuel coûte :

- **Une double expertise.** Migrer à la main suppose de maîtriser **VB6** (le code source) **et**
  **.NET** (la cible) — un profil rare et coûteux. C'est sa principale barrière.
- **De la discipline.** Sans *golden master* (5.5) et sans méthode, le manuel **introduit
  silencieusement** des régressions. Il exige une rigueur de validation que l'outil, lui, impose en
  partie par construction.
- **Du temps.** Le manuel est **lent** sur de gros volumes — d'où l'intérêt de réserver
  l'automatisation au gros œuvre mécanique et de **concentrer** l'effort manuel là où il apporte le
  plus (les points durs).
- **De la retenue.** C'est en migration manuelle que la tentation d'**« améliorer au passage »** est
  la plus forte — et c'est le **piège** du cadrage (3.6). Migrer à la main ne doit **pas** devenir
  l'occasion de refactoriser, corriger ou enrichir : d'abord migrer à l'identique (prouvable par la
  non-régression), **ensuite** améliorer (module 17.5).

---

## ✅ En résumé

- « Migration manuelle **assistée** » ne signifie ni **sans outils** (le pilotage est manuel, pas
  l'exécution) ni **réécriture *from scratch*** (on **porte** l'existant, on ne le redessine pas).
- **Pourquoi** : seul le jugement humain **comprend l'intention** (clé des pièges silencieux),
  **repense** les constructions supprimées (`GoSub`, *control arrays*), produit du code
  **idiomatique**, traite l'**UI**, et tranche les décisions **métier** (code mort, comportements
  critiques). C'est la voie de la **strate rouge** (3.3).
- **Quand** : petites applications, **points durs** de tout projet, **résidu** d'une passe
  automatique, exigence de qualité maximale, code trop particulier pour les outils, ou contrainte
  budgétaire avec expertise interne.
- **Assistée par** : l'IDE moderne (et les avertissements d'`Option Strict`), l'**IA** (module 19,
  avec ses garde-fous), l'usage **sélectif** des outils commerciaux, les **annexes**, et surtout le
  ***golden master*** (5.5) — sans lequel le manuel est un pari aveugle.
- **Faux dilemme** : ce n'est jamais « outil **ou** manuel » mais un **dosage** ; toute migration
  outillée finit à la main, toute migration manuelle s'outille (→ grille de décision, 4.5).
- **Coût et limites** : double expertise (VB6 **et** .NET), discipline, temps, et la **retenue** de
  ne pas améliorer en migrant (piège du cadrage, 3.6).

---

⬅️ Section 4.3 🛠️ — [Outils commerciaux : VBUC et VB Migration Partner](03-outils-commerciaux.md)  
➡️ Section 4.5 — [Grille de décision : outil automatique vs manuel vs hybride](05-grille-decision.md)

---

**Juin 2026** · .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Grille de décision : outil automatique vs manuel vs hybride](/04-outils-migration/05-grille-decision.md)
