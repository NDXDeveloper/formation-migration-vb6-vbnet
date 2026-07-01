🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.1 L'assistant de mise à niveau Visual Studio (*Upgrade Wizard*) ⚠️

> **Module 4 — Outils de migration**
> C'est, historiquement, le tout premier outil que rencontre une migration VB6 → VB.NET :
> l'assistant **intégré à Visual Studio**, gratuit et automatique. Mais c'est aussi l'outil dont
> les **limites** sont les plus structurelles — au point qu'il n'existe plus dans les versions
> modernes de Visual Studio. Comprendre ce qu'il faisait reste pourtant essentiel : c'est le
> meilleur moyen de saisir ce qu'une conversion automatique sait — et ne sait pas — faire.

---

## ⚠️ L'avertissement à poser d'emblée

Commençons par la limite la plus importante, celle qui conditionne toute la lecture de cette  
section :

> ⚠️ **L'*Upgrade Wizard* n'existe plus dans Visual Studio moderne.** Il était intégré aux
> versions de Visual Studio **2002 (VS.NET) à 2008**. Microsoft l'a **retiré à partir de Visual
> Studio 2010**, et les membres de l'espace de compatibilité sur lesquels il reposait ont été
> marqués comme **obsolètes** dans .NET Framework 4. Autrement dit : dans **Visual Studio 2026**,
> cet assistant **n'est pas disponible**. Pour l'utiliser tel quel, il faudrait ressortir une
> version de Visual Studio vieille de plus de quinze ans.

Pourquoi, alors, lui consacrer une section entière ? Parce que :

- comprendre **comment il fonctionnait** donne le bon **modèle mental** de ce qu'est (et n'est
  pas) une conversion automatique — un acquis qui vaut pour *tous* les outils ;
- il illustre parfaitement le problème de la **béquille de compatibilité** (section 4.2 : c'est sa
  dépendance centrale) ;
- son **moteur de conversion** n'a pas disparu : il survit dans les **outils commerciaux** (section
  4.3), qui en sont les descendants directs ;
- son approche **« signaler ce que je ne sais pas faire »** est restée un modèle pour les outils
  qui lui ont succédé.

> 💡 Cette section est donc autant **conceptuelle** qu'historique. On n'apprend pas ici à lancer
> un outil qu'on ne peut plus lancer ; on apprend ce qu'une machine peut automatiser dans une
> migration VB6 → VB.NET, et où s'arrête, inévitablement, l'automatisation.

---

## 🧭 Ce que faisait l'assistant, et comment

Le principe de l'*Upgrade Wizard* était simple et plutôt sain. Lorsqu'on ouvrait un projet VB6
(`.vbp`) dans Visual Studio .NET, l'assistant se déclenchait automatiquement et **ne touchait pas
au projet d'origine** : il en créait une **copie convertie** dans un nouveau projet VB.NET, à un  
emplacement séparé. Le code VB6 de départ restait intact — une précaution précieuse, qui permettait  
de relancer la conversion autant de fois que nécessaire.

Sur cette copie, l'assistant effectuait la **conversion mécanique** :

- **les déclarations et le langage** : les variables étaient converties vers leurs types .NET
  équivalents, et les structures de langage (procédures, propriétés, gestionnaires d'événements,
  structures de contrôle, modules, classes, énumérations, `Type`…) transposées vers leur forme
  VB.NET ;
- **les fonctions intrinsèques** : les appels aux fonctions cœur de VB6 étaient remplacés par leurs
  équivalents .NET *quand un équivalent existait* — sinon, l'assistant **laissait une note** dans
  le code ;
- **les formulaires** : les `.frm` et leurs contrôles étaient convertis vers leurs équivalents
  Windows Forms ;
- **l'accès aux données** : les objets de données et de connexion étaient enveloppés en **RCW**
  (interop COM, module 16.1) plutôt que portés vers ADO.NET.

> 💡 Le **moteur** de cet assistant n'a pas été écrit par Microsoft : il provenait de la société
> **ArtinSoft** (devenue **Mobilize.Net**). C'est un point qui éclaire la suite du module : l'outil
> commercial **VBUC** (section 4.3) est issu de la **même lignée technologique** — la version
> commerciale et entretenue du moteur que Microsoft avait intégré gratuitement. Comprendre
> l'assistant, c'est donc déjà comprendre une partie de ce que font les outils commerciaux.

---

## 📋 Le rapport de migration : sa fonctionnalité la plus intelligente

Le véritable apport de l'*Upgrade Wizard* n'était pas tant le code qu'il produisait que sa
**lucidité sur ses propres limites**. À l'issue de la conversion, il générait deux choses :

1. **Un rapport de migration** (`_UpgradeReport.htm`), au format HTML, consultable dans Visual
   Studio ou dans un navigateur. Il listait les **problèmes globaux**, puis, fichier par fichier,
   le détail de ce qui n'avait pas pu être converti — avec la **localisation** exacte du code
   concerné et une description.
2. **Des commentaires insérés directement dans le code converti**, qui apparaissaient aussi comme
   des tâches dans la *Task List* de Visual Studio. Un double-clic menait droit à la ligne
   concernée.

Ces commentaires suivaient une convention en quatre niveaux, qu'il est utile de connaître car on  
les retrouve dans tout code issu de cet outil (et, sous des formes voisines, chez ses descendants) :

| Marqueur | Sens | Action attendue |
|----------|------|-----------------|
| `'UPGRADE_NOTE:` | Information : une transformation **automatique** a été appliquée | Vérifier que le comportement reste correct |
| `'UPGRADE_WARNING:` | Le code compile, mais son **comportement a pu changer** | **Examiner** — c'est le marqueur le plus dangereux ⚠️ |
| `'UPGRADE_TODO:` | Une intervention **manuelle** est nécessaire pour que ça fonctionne | Compléter le code à la main |
| `'UPGRADE_ISSUE:` | Quelque chose **n'a pas pu être converti** (souvent une erreur de compilation) | Corriger impérativement |

> 💡 **L'approche « signaler ce que je ne sais pas faire » est le meilleur héritage de cet
> assistant.** Plutôt que de prétendre tout convertir, il **balisait honnêtement** sa propre
> ignorance — exactement l'esprit du *repère honnête* de cette formation (« aucun outil ne migre
> *tout* »). Un code converti truffé de `UPGRADE_WARNING` et de `UPGRADE_TODO` n'est pas un échec
> de l'outil : c'est l'outil qui fait son travail, en désignant le travail qui reste à l'humain.

---

## ✅ Ses capacités réelles

Replacé dans son contexte, l'*Upgrade Wizard* faisait correctement ce pour quoi il était conçu :  
le **gros œuvre mécanique**. Il prenait en charge l'essentiel de la transposition syntaxique et  
structurelle — la part « facile » de la migration, celle qui est fastidieuse mais sans jugement.  
C'est exactement la fraction **60-90 %** évoquée dans l'introduction du module : le volume de code
« vert » (module 3.3) qui se convertit sans difficulté.

Ses points forts :

- **Gratuité et intégration** (à l'époque) : disponible dans l'IDE, sans installation ni achat ;
- **Non-destructif** : le projet VB6 d'origine restait intact ;
- **Couverture large du langage** : déclarations, structures, classes, énumérations, `Type`,
  formulaires de base étaient transposés de façon fiable ;
- **Traçabilité** : le rapport et les commentaires donnaient une **carte du travail restant**,
  localisée et documentée.

---

## ⚠️ Ses limites — le cœur de la section

Les limites de l'assistant ne sont pas des défauts d'implémentation qu'une mise à jour aurait  
corrigés : elles sont **structurelles**. Elles tiennent à ce qu'une conversion automatique *peut*  
faire et à la façon dont cet outil-là a choisi de combler ce qu'elle ne peut pas.

### 1. Il n'existe plus dans Visual Studio moderne

C'est la limite première, déjà posée plus haut : **retiré depuis Visual Studio 2010**, il est
**absent de Visual Studio 2026**. Cette seule réalité le disqualifie comme outil de migration
*pratique* aujourd'hui. Les démarches modernes passent par les **outils commerciaux** (4.3), la
**migration manuelle** (4.4) ou l'**IA** (module 19) — éventuellement combinés (4.5).

### 2. Il s'appuie entièrement sur la béquille de compatibilité

Lorsqu'il ne savait pas convertir vers une API .NET native, l'assistant ne refactorisait pas : il
**rabattait** le code sur l'espace de noms `Microsoft.VisualBasic.Compatibility` (les fameux appels
`VB6.*`). Or cet espace est **obsolète** et **fonctionne uniquement en 32 bits** — ce qui contraint
la cible (compilation `x86`, contrainte déjà rencontrée en 3.2 et 3.5). Le code produit
« fonctionne », mais traîne une **dette** qu'il faudra rembourser (section 4.2, et suppression au
module 17.2). Cette béquille fait l'objet de la **section suivante**.

### 3. Il ne voit pas les pièges silencieux

C'est la limite la plus dangereuse, et elle rejoint le fil rouge de toute la formation. L'assistant
**signalait** certains points sensibles via ses `UPGRADE_WARNING` (tableau déclaré avec `New`,
tableau dont la borne inférieure n'est pas nulle, usage de `Null`/`IsNull` que VB.NET ne connaît  
pas…), mais il **ne pouvait pas garantir l'équivalence de comportement**. Pire : il existait des  
cas où **aucune erreur n'apparaissait dans le rapport**, où le code compilait, **et où  
l'application ne se comportait plus pareil** à l'exécution. C'est la définition même du **piège  
silencieux** (Annexe B) : le `ByRef` devenu `ByVal`, l'arrondi qui change, la finalisation perdue.

> ⚠️ **Aucun outil — pas même celui-ci — ne dispense de la non-régression.** Le fait que
> l'assistant « ne signale rien » sur une portion de code ne prouve **pas** qu'elle se comporte
> comme avant. Seule la comparaison méthodique avec le **golden master** (modules 5.5 et 17.3)
> débusque ces écarts. C'est précisément pour cela que le critère de réussite (3.6) impose ce
> harnais de tests.

### 4. Une qualité de sortie peu idiomatique

Le code produit était **fonctionnel mais pas idiomatique** : truffé d'appels `VB6.*`, calquant les  
structures VB6 plutôt que d'adopter les tournures .NET modernes (LINQ, génériques, espace `My`…).  
Les ***control arrays*** (tableaux de contrôles, module 13.6) étaient gérés via une **classe  
d'aide** de compatibilité, et non recréés proprement. Le résultat exigeait donc une **reprise  
substantielle** pour atteindre un code VB.NET sain — tout l'objet de la phase de fiabilisation
(module 17).

### 5. Une robustesse fragile sur les dépendances tierces

En pratique, l'assistant pouvait **se bloquer** sur certains contrôles **ActiveX/OCX tiers** (les  
points durs identifiés en 3.2), parfois sans message d'erreur exploitable. La parade consistait à
**désactiver les références une à une** dans le `.vbp` pour isoler le composant fautif — une
démarche manuelle, fastidieuse, qui renvoyait au problème de fond : les composants sans équivalent  
ne se convertissent pas, ils se **remplacent** (modules 13.6, 14.6).

### 6. Une cible .NET ancienne

L'assistant produisait du code ciblant les **premières versions** de .NET (1.x / 2.0). Même en le  
ressortant aujourd'hui, on obtiendrait donc un projet d'une autre époque, qu'il faudrait ensuite
**reciblé vers .NET Framework 4.7.2** (module 6.1) — un détour de plus, contraire à la logique
« pont » directe que vise cette formation.

---

## 🧰 Ce qu'il faut en retenir pour aujourd'hui

L'*Upgrade Wizard* appartient au passé en tant qu'**outil**, mais reste vivant en tant que
**leçon**. Ce qu'il enseigne se transpose intégralement aux outils actuels :

- une conversion automatique prend en charge le **gros œuvre mécanique**, pas le **jugement** ;
- combler l'inconvertible par une **béquille** (compatibilité) crée une **dette**, pas une
  solution ;
- « pas d'erreur signalée » **n'est pas** « comportement identique » — les **pièges silencieux**
  passent sous les radars ;
- la sortie d'un outil est un **point de départ** à valider et fiabiliser (module 17), jamais une
  livraison.

> 💡 Le moteur de cet assistant, repris et entretenu par **Mobilize.Net**, est précisément ce que
> l'on retrouve — modernisé, étendu, capable de supprimer les dépendances de compatibilité et de
> mieux gérer les *control arrays* — dans **VBUC** (section 4.3). La suite logique de cette section
> est donc double : comprendre la **béquille** qu'il employait (4.2), puis découvrir les outils qui
> en ont pris la relève (4.3).

---

## ✅ En résumé

- L'*Upgrade Wizard* était l'assistant de migration VB6 → VB.NET **intégré à Visual Studio**, des
  versions **2002 à 2008**. Il a été **retiré à partir de Visual Studio 2010** et **n'existe pas
  dans Visual Studio 2026** : c'est sa limite la plus structurelle.
- Il fonctionnait de façon **non-destructive** (projet VB6 intact, copie convertie à part) et
  prenait en charge la **conversion mécanique** : déclarations, langage, formulaires, et accès aux
  données enveloppé en RCW.
- Sa fonctionnalité la plus précieuse était le **rapport de migration** et les **commentaires
  insérés** (`UPGRADE_NOTE`/`WARNING`/`TODO`/`ISSUE`) : il **balisait honnêtement** ce qu'il ne
  savait pas faire.
- Ses **capacités** couvraient le travail « facile » (la strate verte de 3.3), soit les fameux
  60-90 % de **volume** mécanique.
- Ses **limites** sont structurelles : indisponible aujourd'hui, **dépendance totale à la béquille
  de compatibilité** (`VB6.*`, obsolète et 32 bits — section 4.2), **incapacité à détecter les
  pièges silencieux** (Annexe B), sortie **peu idiomatique**, **fragilité** sur l'ActiveX tiers, et
  **cible .NET ancienne** à recibler ensuite.
- Son **moteur** (ArtinSoft → Mobilize.Net) survit dans l'outil commercial **VBUC** (section 4.3) :
  comprendre l'assistant, c'est déjà comprendre ses descendants.
- La leçon transposable à tout outil : l'automatisation traite le mécanique, **jamais** le
  jugement ; et « aucune erreur signalée » ne vaut **pas** preuve de non-régression.

---

⬅️ Module 4 — [Outils de migration](README.md)  
➡️ Section 4.2 ⚠️ — [L'espace de noms `Microsoft.VisualBasic.Compatibility` : la béquille à éviter](02-compatibility-namespace.md)

---

**Juin 2026** · .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [L'espace de noms `Microsoft.VisualBasic.Compatibility` : la **béquille à éviter** (et pourquoi)](/04-outils-migration/02-compatibility-namespace.md)
