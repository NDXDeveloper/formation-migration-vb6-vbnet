🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.2 L'espace de noms `Microsoft.VisualBasic.Compatibility` : la béquille à éviter ⚠️

> **Module 4 — Outils de migration**
> La section 4.1 l'a désigné sans le détailler : c'est sur cet espace de noms que l'*Upgrade
> Wizard* (et, à des degrés divers, les autres outils) **rabattait** tout ce qu'il ne savait pas
> convertir proprement. Cette section explique **ce qu'il est**, **pourquoi il est si tentant**,
> et surtout **pourquoi il faut l'éviter** — ou, s'il est déjà là, le **rembourser** (module 17.2).

---

## 🧭 De quoi parle-t-on exactement

`Microsoft.VisualBasic.Compatibility` est un **assemblage** (`Microsoft.VisualBasic.Compatibility.dll`)
livré avec .NET Framework, qui regroupe les espaces de noms `Microsoft.VisualBasic.Compatibility`  
et `Microsoft.VisualBasic.Compatibility.VB6`. Sa raison d'être tient en une phrase : **émuler le  
comportement de VB6 à l'intérieur de .NET**.

Concrètement, il fournit une collection de classes et de fonctions d'aide — les fameux appels en
`VB6.*` — qui **reproduisent** des mécanismes VB6 n'ayant pas d'équivalent .NET *direct* ou
*propre*. Plutôt que de porter le code vers une API native (ce qui exige du jugement), un outil
peut le **brancher** sur ces béquilles, qui imitent l'ancien comportement. Quelques exemples de ce  
que couvre cet espace : la prise en charge des ***control arrays*** (module 13.6), certaines  
opérations de bas niveau sur les formulaires et contrôles (récupération du contrôle ayant le focus,  
manipulation de l'ordre de superposition `ZOrder`, obtention du *handle* d'instance…), et divers  
comportements hérités des contrôles intrinsèques.

L'écart saute aux yeux sur un cas simple — l'ordre de superposition (`ZOrder`) :

```vb
' VB6 — méthode intrinsèque du contrôle
Command1.ZOrder 0

' VB.NET issu de l'outil — rabattu sur la béquille (obsolète, 32 bits)
VB6.Support.ZOrder(Command1, 0)      ' Microsoft.VisualBasic.Compatibility.VB6

' VB.NET natif — la vraie cible, sans aucune dépendance de compatibilité
Command1.BringToFront()
```

Les deux dernières lignes produisent le **même effet visuel** ; mais celle du milieu *compile et  
tourne* en traînant la béquille, quand seule la dernière est du .NET propre. Toute la section tient  
dans cet écart.

> 💡 L'idée fondatrice était louable : **faire compiler et tourner** rapidement un projet converti,
> en attendant que le développeur remplace, *plus tard*, chaque béquille par du vrai code .NET. Le
> problème — tout l'objet de cette section — est que ce « plus tard » n'arrive **jamais** si l'on
> n'y prend pas garde, et que la béquille a un coût bien plus lourd qu'il n'y paraît.

---

## 🩼 Pourquoi c'est si tentant (et pourquoi les outils s'en servent)

La béquille de compatibilité est le **chemin de moindre résistance**. Pour un outil de conversion  
automatique, elle résout d'un coup un problème difficile : que faire des constructions VB6 sans
équivalent .NET évident ? La réponse facile est : « les mapper sur un *shim* qui reproduit l'ancien
comportement ». Le code résultant **compile** et **s'exécute** immédiatement — ce qui donne  
l'agréable (et trompeuse) impression que la migration est presque finie.

C'est exactement ce que faisait l'*Upgrade Wizard* (4.1) : face à l'inconvertible, il **rabattait**  
sur `VB6.*` plutôt que de laisser un trou. D'où une caractéristique du code issu de cet assistant :  
il est **truffé d'appels `VB6.*`**.

> ⚠️ **Le pourcentage rassurant, encore une fois.** Un projet « entièrement converti et qui tourne »
> grâce à la béquille peut donner le sentiment que **100 % du travail** est fait. En réalité, une
> grande partie du travail a seulement été **reportée** : chaque `VB6.*` est une dette de migration
> qui reste à honorer. C'est l'illusion du module 1.5, dans sa forme la plus sournoise — non plus
> « 80 % du code converti » mais « 100 % du code converti… avec des béquilles partout ».

---

## ⚠️ Pourquoi c'est une *béquille* — et pas une solution

La métaphore est exacte, et il vaut la peine de la prendre au sérieux. Une béquille **vous tient  
debout** quand vous ne pouvez pas marcher seul — c'est utile, parfois indispensable, le temps de  
guérir. Mais une béquille qu'on **ne pose jamais** vous empêche de réapprendre à marcher : l'aide  
provisoire devient un **handicap permanent**.

Appliqué au code, cela donne :

- **Ce n'est pas du .NET idiomatique** : c'est de la **sémantique VB6 déguisée** en .NET. Le code
  *ressemble* à du VB.NET, mais *pense* encore en VB6.
- **Cela masque le travail réel** : tant que la béquille tient, on ne **voit** plus qu'il restait
  une vraie migration à faire à cet endroit. Le problème n'est pas résolu, il est **caché**.
- **Cela perpétue l'ancien modèle** : on ne tire aucun bénéfice de .NET (LINQ, génériques, espace
  `My`, types natifs) là où la béquille reproduit l'ancien monde.

> 💡 C'est le même piège, sous une autre forme, que ceux déjà rencontrés dans la formation : la
> **cohabitation COM** qui se fige en « hybride permanent » (3.5), la migration **incrémentale**
> qui « ne finit jamais » (3.4). À chaque fois, un **état de transition** légitime se transforme en
> **destination** par défaut d'avoir été démonté à temps. La béquille de compatibilité est l'avatar
> de ce piège au niveau du **code**.

---

## 🚫 Pourquoi l'éviter : les quatre raisons dures

Au-delà de la qualité du code, quatre raisons **concrètes** font de cet espace de noms une  
dépendance à fuir.

### 1. Il est officiellement obsolète

Les membres de ces espaces de noms ont été **marqués comme obsolètes** dès **.NET Framework 4**.  
Les utiliser génère des **avertissements de compilation** : Microsoft signale lui-même qu'il s'agit  
d'une **voie sans avenir**. Construire une migration neuve sur des API obsolètes, c'est bâtir sur un  
terrain dont on sait déjà qu'il est condamné.

### 2. Il ne fonctionne qu'en 32 bits

Les membres de `Microsoft.VisualBasic.Compatibility.VB6` ne sont pris en charge **que dans des  
processus 32 bits**. En dépendre **force** donc une compilation en `x86` et **interdit le 64 bits**.  
C'est la **contrainte 32 bits** déjà croisée à propos des composants COM/OCX hérités (3.2, 3.5) et  
du code de l'*Upgrade Wizard* (4.1) — mais ici, elle est imposée par la **béquille elle-même**, pas  
par une dépendance externe.

### 3. Il est **absent de .NET moderne** — le mur devant la suite

C'est la raison la plus décisive, et elle touche au cœur de la philosophie « pont » de cette  
formation. L'assemblage de compatibilité **n'a jamais été porté** vers .NET Core / .NET 5+ : il
**n'existe pas** dans .NET moderne. Toute fonction `VB6.*` (`ZOrder`, `GetHInstance`,
`GetActiveControl`, `ValidateControl`, `WhatsThisMode`…) y est tout simplement **indisponible**.

Les conséquences sont radicales :

- un code qui dépend de cette béquille est **verrouillé sur .NET Framework**, **définitivement**,
  tant que ces dépendances subsistent ;
- le **second saut** envisagé au module 20 (de 4.7.2 vers .NET 8/10) devient **impossible** sans
  avoir d'abord **éliminé toutes** ces dépendances.

> ⚠️ **La béquille ne fait pas que reporter le travail : elle dresse un mur.** On migre vers 4.7.2
> en s'appuyant sur `VB6.*`, on croit avoir gagné du temps… et l'on découvre, le jour où l'on veut
> moderniser (module 20), qu'il faut **tout** déboucler d'abord. La dette de compatibilité n'est pas
> seulement coûteuse : elle **bloque l'avenir** de l'application. C'est l'argument numéro un pour ne
> pas la laisser s'installer.

### 4. L'assemblage lui-même est sur un terrain mouvant

Microsoft a explicitement averti qu'il était **possible que l'assemblage** contenant cet espace de  
noms **soit retiré** d'une future version du redistribuable .NET Framework — ce qui obligerait alors
à l'**empaqueter séparément** et à le redistribuer avec l'application. Une dépendance dont on vous
prévient qu'elle pourrait disparaître n'est pas un socle sur lequel bâtir.

---

## 💡 La nuance honnête : tolérance transitoire, pas interdiction absolue

Faut-il pour autant **bannir** la béquille en toute circonstance ? Non — et il serait malhonnête de  
le prétendre. Comme `Option Strict Off` (module 6.4), elle peut être **tolérée transitoirement**  
pour obtenir une **première compilation** rapide d'un projet fraîchement converti, à condition de la  
traiter pour ce qu'elle est : un **échafaudage**, pas une fondation.

La règle est donc :

1. **Accepter** éventuellement les `VB6.*` comme point de départ, pour avoir un code qui tourne ;
2. les **inventorier** comme une dette explicite (combien ? où ?) ;
3. les **remplacer progressivement** par du code .NET natif — c'est précisément l'objet du module
   **17.2 (« Supprimer les dépendances à `Microsoft.VisualBasic.Compatibility` »)** ;
4. viser **zéro** `VB6.*` résiduel à la fin, comme indicateur de qualité (3.6).

> ⚠️ Le danger n'est pas d'utiliser la béquille un moment : c'est de **ne jamais la poser**. Un
> projet qui livre en production avec ses `VB6.*` partout n'a pas terminé sa migration — il l'a
> **gelée à mi-chemin**, en se condamnant au 32 bits et en se fermant la porte de .NET moderne.

---

## ✅ Que faire à la place

L'alternative à la béquille, c'est le **vrai équivalent .NET** — exactement ce que toute la
**partie 3** (migration du langage) s'attache à enseigner, construction par construction :

- les ***control arrays*** se **recréent** proprement (tableau géré à la main, partage de
  gestionnaire d'événements — module 13.6) plutôt que de s'appuyer sur la classe d'aide de
  compatibilité ;
- les opérations de formulaire/contrôle passent par les **API WinForms natives** (`BringToFront`,
  `SendToBack`, `ActiveControl`… — partie 4) ;
- `FileSystemObject` et consorts cèdent la place à **`System.IO`** et à l'espace **`My`** (module
  16.6) ;
- les fonctions intrinsèques trouvent leur équivalent dans les **types natifs** ou dans le runtime
  VB.NET supporté (voir la distinction ci-dessous).

> 💡 **Distinction cruciale à ne pas confondre :**
> `Microsoft.VisualBasic` **n'est pas** `Microsoft.VisualBasic.Compatibility`.
> - **`Microsoft.VisualBasic`** est le **runtime VB.NET supporté** : il contient les fonctions
>   familières (`Left`, `Right`, `Mid`, `MsgBox`, `IIf`, `Format`…), il est **partiellement porté**
>   vers .NET moderne, et son usage est **légitime** (pour du code idiomatique, on lui préfère
>   parfois les équivalents `System.*` — c'est l'objet du module 9.2, mais ce n'est **pas** une
>   béquille).
> - **`Microsoft.VisualBasic.Compatibility`** (et `.VB6`) est la **béquille obsolète** décrite ici :
>   32 bits, absente de .NET moderne, **à éviter et à supprimer**.
>
> Confondre les deux conduirait à se priver à tort d'un runtime parfaitement valide — ou, à
> l'inverse, à croire « propre » du code qui ne l'est pas. La frontière passe par le mot
> **`Compatibility`** dans le nom de l'espace.

---

## ✅ En résumé

- `Microsoft.VisualBasic.Compatibility` (et `.VB6`) est un assemblage qui **émule le comportement
  de VB6** dans .NET via des appels **`VB6.*`** ; c'est la dépendance sur laquelle l'*Upgrade
  Wizard* (4.1) rabattait l'inconvertible.
- Il est **tentant** parce qu'il fait **compiler et tourner** le code immédiatement — mais il
  **reporte** le travail au lieu de le **faire**, et nourrit l'illusion d'une migration achevée.
- C'est une **béquille, pas une solution** : code **non idiomatique** (sémantique VB6 déguisée),
  travail réel **masqué**, ancien modèle **perpétué**.
- Quatre raisons dures de l'éviter : **obsolète** (depuis .NET Framework 4), **32 bits uniquement**,
  **absent de .NET moderne** (il **verrouille** sur .NET Framework et **bloque** le second saut du
  module 20), et reposant sur un **assemblage** que Microsoft pourrait retirer.
- Nuance honnête : **tolérance transitoire** acceptable (comme `Option Strict Off`, module 6.4),
  à condition de **rembourser la dette** — inventorier puis **supprimer progressivement** les
  `VB6.*` (module 17.2), pour viser **zéro** résiduel.
- À la place : les **vrais équivalents .NET** (recréer les *control arrays* en 13.6, API WinForms
  natives, `System.IO`/`My` en 16.6…).
- **Ne pas confondre** : `Microsoft.VisualBasic` (runtime **supporté**, légitime) ≠
  `Microsoft.VisualBasic.Compatibility` (béquille **obsolète**, à éviter).

---

⬅️ Section 4.1 ⚠️ — [L'assistant de mise à niveau Visual Studio (*Upgrade Wizard*)](01-upgrade-wizard.md)  
➡️ Section 4.3 — [Outils commerciaux : VBUC et VB Migration Partner](03-outils-commerciaux.md)

---

**Juin 2026** · .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Outils commerciaux : **VBUC** (Mobilize.Net) et **VB Migration Partner** (Code Architects)](/04-outils-migration/03-outils-commerciaux.md)
