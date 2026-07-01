🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.2 Cartographie des dépendances 🔗

> **Module 3 — Évaluer l'existant et choisir une stratégie**
> L'inventaire (3.1) a **recensé** ce que contient le projet. Cette section va plus loin :
> elle **relie** les éléments entre eux et, surtout, identifie ce dont l'application dépend
> **sans le contenir**. C'est là que se loge la part la plus imprévisible du risque.

---

## 🧭 Inventorier n'est pas cartographier

L'inventaire produit une **liste** : tant de formulaires, tant de classes, tant de références.  
La cartographie produit une **carte** : qui dépend de quoi, quel composant est utilisé par  
combien de modules, et — question décisive — **que se passe-t-il si tel composant n'a pas  
d'équivalent en .NET ?**

La différence est fondamentale. Le code que vous **possédez** (vos `.frm`, `.bas`, `.cls`), vous  
le migrez : c'est du travail, mais il est **sous votre contrôle**. Les dépendances que vous ne  
possédez pas — un contrôle ActiveX d'un éditeur tiers, une bibliothèque COM, un pilote de base de  
données — échappent à ce contrôle. Vous ne pouvez pas les « migrer » : vous pouvez seulement les
**réutiliser** (via l'interopérabilité), les **remplacer**, ou les **éliminer**.

> ⚠️ **La dépendance externe est le risque n°1 d'un chiffrage qui dérape.** On sait estimer le
> coût de conversion de son propre code. On découvre, parfois trop tard, qu'un composant tiers
> indispensable n'existe plus, n'a pas d'équivalent .NET, ou ne fonctionne que sous interop avec
> des limitations. Cartographier les dépendances **dès l'évaluation**, c'est transformer ces
> inconnues en décisions documentées.

---

## 🗺️ Les quatre familles de dépendances

Le titre de cette section les annonce : **COM**, **ActiveX**, **bases de données**, **composants  
tiers**. Ces catégories se recoupent partiellement (un OCX tiers *est* un composant ActiveX *et*  
un composant tiers), mais chacune appelle une analyse propre.

### 1. Les composants COM (non visuels) 🔗

Ce sont les bibliothèques appelées **par code**, sans présence visuelle sur un formulaire :  
référencées par *early binding* (lignes `Reference=` du `.vbp`) ou instanciées par *late binding*
(`CreateObject`). Exemples typiques :

- **ADO** (`Microsoft ActiveX Data Objects`) pour l'accès aux données — voir famille « bases de
  données » ci-dessous ;
- la **bibliothèque Scripting** (`FileSystemObject`, `Dictionary`) ;
- des **DLL ActiveX maison** (souvent issues d'un `.vbg`, c'est-à-dire d'autres projets VB6 de
  l'organisation) ;
- des automations Office (Word, Excel) pilotées depuis VB6.

Pour chacun, la cartographie pose la question : **le réutilise-t-on via interop (RCW), ou  
existe-t-il une voie native .NET ?** Une DLL ActiveX maison peut souvent rester en COM le temps de  
la transition (cohabitation, 3.5) ; `FileSystemObject` a un équivalent natif clair (`System.IO`,  
espace `My`, module 16.6) et n'a pas vocation à survivre.

### 2. Les contrôles ActiveX / OCX (visuels) ⚠️

Ce sont les composants **déposés sur les formulaires** (lignes `Object=` du `.vbp`) : grilles,  
calendriers, barres d'outils, contrôles communs Windows. Ils posent un problème particulier car  
ils sont à la fois une **dépendance** et un **élément d'interface**.

| Type d'OCX | Exemple | Stratégie typique |
|------------|---------|-------------------|
| **Contrôles communs Microsoft** | `MSComctlLib` (TreeView, ListView, ImageList…), `MSFlexGrid` | Réutilisables via interop, mais équivalents WinForms natifs souvent préférables |
| **Contrôles tiers maintenus** | Grilles commerciales encore éditées | Vérifier l'existence d'une **version .NET** chez l'éditeur |
| **Contrôles tiers abandonnés** | Composant d'un éditeur disparu | **Point dur** : remplacement quasi obligatoire ⚠️ |
| **UserControls maison** (`.ctl`) | Contrôle créé en interne | À recréer en WinForms (vous en avez le code) |

> ⚠️ **L'OCX tiers abandonné est le cauchemar classique de la migration.** Pas de version .NET,
> plus d'éditeur, parfois plus de licence : il faut lui trouver un **remplaçant** (composant
> WinForms équivalent) et **réécrire** le code d'interaction. C'est rarement une conversion
> ligne à ligne — voir modules 13.6 (*control arrays* et contrôles) et 14.6 (contrôles tiers).

### 3. Les bases de données et l'accès aux données

L'accès aux données est presque toujours une dépendance majeure d'une application VB6 de gestion.  
La cartographie doit identifier **deux choses distinctes** :

- **La technologie d'accès** utilisée dans le code : **DAO**, **RDO** ou **ADO** ? (parfois les
  trois cohabitent dans un même projet historique). Toutes devront converger vers **ADO.NET**
  (module 15) — mais l'effort diffère selon le point de départ.
- **Le ou les SGBD ciblés** : base **Access** (`.mdb`/`.accdb`), **SQL Server**, autre moteur via
  ODBC/OLE DB ? Cela conditionne les **chaînes de connexion**, les pilotes et certaines
  adaptations (module 15.5).

Points d'attention pour la carte :

- Le contrôle **`Data`** et la liaison de données visuelle (*data binding* VB6) n'ont pas
  d'équivalent direct et migrent vers `BindingSource` (module 15.4) ⚠️ ;
- Les **`Recordset`** (curseurs ADO/DAO) basculent vers un modèle différent — `DataReader`
  (connecté) ou `DataSet`/`DataTable` (déconnecté) — avec un **changement de paradigme** à
  anticiper (module 15.3) ⚠️ ;
- Les composants de **reporting**, **Crystal Reports** en tête, sont une dépendance à part
  entière, souvent sous-estimée (module 15.7) ⚠️ ;
- Les **`DataEnvironment`** (fichiers `.dsr`) repérés à l'inventaire relèvent de cette famille.

### 4. Les composants tiers (au sens large)

Catégorie « fourre-tout » mais essentielle : tout ce qui provient de l'**extérieur** et n'entre  
pas proprement dans les trois cases précédentes — bibliothèques de calcul, moteurs PDF, SDK  
matériels (lecteurs de codes-barres, balances, terminaux), composants de communication, etc.

Pour chacun, trois statuts possibles, à **trancher explicitement** :

1. **Encore disponible et maintenu** → vérifier s'il existe une version .NET native ; sinon,
   interop COM (s'il est exposé en COM) ;
2. **Disponible mais figé** (plus de support) → décision au cas par cas selon sa criticité ;
3. **Disparu** → **point dur** : recherche d'un remplaçant, voire réimplémentation.

---

## 🔍 Que noter pour chaque dépendance

La cartographie n'a de valeur que si elle est **systématique**. Pour chaque dépendance recensée  
en 3.1, la carte renseigne les mêmes colonnes :

| Information | Pourquoi |
|-------------|----------|
| **Nom et éditeur** | Identifier l'origine (Microsoft / tiers / interne) |
| **Version et GUID** | Retrouver le composant exact, vérifier sa disponibilité |
| **Type** | COM non visuel / OCX visuel / accès données / autre |
| **Statut de maintenance** | Maintenu / figé / abandonné — **le critère de risque** ⚠️ |
| **Existence d'un équivalent .NET** | Natif disponible ? chez quel éditeur ? |
| **Utilisé par** | Quels modules/formulaires en dépendent (l'**empreinte**) |
| **Décision** | **Réutiliser (interop) / Remplacer / Éliminer** |

La colonne **« Utilisé par »** est ce qui distingue vraiment une carte d'une liste : elle mesure  
l'**empreinte** de chaque dépendance. Un composant utilisé dans un seul formulaire isolé se traite  
différemment du même composant câblé dans quarante écrans. L'empreinte conditionne à la fois le
**coût** du remplacement et le **risque** de régression.

La colonne **« Décision »** est l'aboutissement : c'est elle qui alimente directement le chiffrage
(3.3) et le choix de stratégie (3.4).

---

## ⚖️ Les trois issues possibles d'une dépendance

Toute dépendance externe finit dans l'une de ces trois catégories. Les nommer explicitement, pour  
chaque composant, est l'objectif de la cartographie.

### Réutiliser via interopérabilité 🔗

Le composant **reste en COM** et .NET l'appelle à travers un **RCW** (*Runtime Callable Wrapper*,  
module 16.1). C'est la voie qui **minimise le changement immédiat** — cohérente avec la  
philosophie « pont » de la formation — et qui rend possible la **cohabitation** (3.5) et la  
migration **incrémentale**.

> 💡 Réutiliser via interop est souvent le **bon premier choix** pour un composant maintenu et
> sans équivalent natif évident : on migre le code, on garde le composant, on évalue son
> remplacement *plus tard* — ou jamais (rester sur l'interop est un choix légitime).

### Remplacer

Le composant **n'a pas d'avenir** (abandonné, pas d'équivalent COM réutilisable, ou son maintien  
en interop est plus coûteux que sa substitution). On lui trouve un **équivalent .NET** (contrôle  
WinForms, bibliothèque native) et on **réécrit** le code d'interaction. C'est le cas le plus  
coûteux : il ne s'agit pas d'une conversion mais d'un **portage**.

### Éliminer

Le composant **n'a plus de raison d'être** : il existe un équivalent **dans le framework lui-même**.
`FileSystemObject` → `System.IO` ; la bibliothèque Scripting `Dictionary` → `Dictionary(Of K,V)` ;
l'objet `App` → l'espace `My` (module 16.6). On supprime la dépendance et on adopte l'API native.

---

## 🎯 Faire émerger les points durs

La finalité concrète de la cartographie est de **faire remonter les points durs** — ces  
dépendances qui pèseront le plus lourd dans le chiffrage et le risque. Quelques signaux à  
surveiller en bâtissant la carte :

- **OCX ou composant tiers abandonné** sans équivalent identifié → remplacement obligatoire, coût
  élevé et incertain ;
- **Composant à forte empreinte** (utilisé partout) → tout changement le concernant se propage à
  l'ensemble de l'application ;
- **Crystal Reports** ou autre moteur de reporting → dépendance lourde, à traiter à part (15.7) ;
- **Accès données par `Recordset`/`DataEnvironment`** → changement de paradigme, pas seulement de
  syntaxe (15.3) ;
- **DLL maison sans code source ni documentation** → zone grise dangereuse, à éclaircir en
  priorité ;
- **Dépendances chargées dynamiquement** (`CreateObject`, late binding) → invisibles dans le
  `.vbp`, donc faciles à oublier (d'où l'importance du balayage de code en 3.1).

> 💡 **Un outil commercial de migration** (VBUC, VB Migration Partner — module 4.3) produit
> généralement, dans sa phase d'évaluation, un **rapport de dépendances** qui accélère
> considérablement ce travail : il liste les composants, signale ceux qu'il sait traiter et ceux
> qui demanderont une intervention. C'est un point de départ précieux — à **valider**, jamais à
> prendre pour argent comptant.

---

## ⚠️ Les pièges de la cartographie

- **Croire que tout ce qui est en COM se réutilise sans douleur.** L'interop fonctionne, mais
  elle a un coût (déploiement des composants, *marshaling*, parfois des limitations). Réutiliser
  n'est pas « gratuit ».
- **Oublier la chaîne de dépendances transitives.** Un composant tiers peut lui-même dépendre
  d'autres DLL ou d'une version précise d'un runtime. La carte doit suivre la chaîne, pas
  seulement le premier maillon.
- **Confondre « ça compile » et « ça fonctionne ».** Une dépendance peut se référencer sans
  erreur et se comporter différemment à l'exécution — c'est l'esprit même des **pièges silencieux**
  (Annexe B).
- **Sous-estimer les dépendances 32 bits.** De nombreux composants COM/OCX hérités sont **32
  bits** uniquement ; cela contraint la cible (compilation en `x86`) et mérite d'être noté sur la
  carte.
- **Négliger la dimension licence.** Un composant tiers « disponible » peut nécessiter une
  **licence** que l'organisation ne possède plus, ou ne couvre pas le nouveau contexte. Statut
  technique et statut juridique sont **deux choses différentes**.

---

## ✅ En résumé

- La cartographie **relie** ce que l'inventaire a listé, et se concentre sur les dépendances
  **externes** — ce que l'on ne possède pas et que l'on ne peut donc pas simplement « migrer ».
- Quatre familles à analyser : **composants COM** (non visuels), **contrôles ActiveX/OCX**
  (visuels), **bases de données** (DAO/RDO/ADO et SGBD), **composants tiers** (au sens large).
- Pour chaque dépendance, on note nom, version, type, **statut de maintenance**, équivalent .NET
  éventuel, **empreinte** (« utilisé par ») et, surtout, une **décision**.
- Trois issues possibles, à trancher explicitement : **réutiliser** via interop (RCW),
  **remplacer**, ou **éliminer** au profit d'une API native.
- L'objectif final est de **faire émerger les points durs** — OCX abandonnés, composants à forte
  empreinte, reporting, accès données — qui pèseront sur le chiffrage (3.3) et la stratégie (3.4).
- Attention aux pièges : dépendances transitives, composants 32 bits, dépendances chargées
  dynamiquement, et questions de **licence** distinctes des questions techniques.

---

⬅️ Section 3.1 — [Inventaire](01-inventaire.md)  
➡️ Section 3.3 — [Estimer l'effort](03-estimer-effort.md)

---

**Juin 2026** · .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Estimer l'effort (taille, complexité cyclomatique, points durs)](/03-evaluer-strategie/03-estimer-effort.md)
