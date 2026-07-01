🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 18.1 — Packaging .NET Framework (installeur, prérequis 4.7.2, redistribuables)

**Produire un artefact déployable, fiable et reproductible : un installeur qui pose l'application, garantit la présence du runtime 4.7.2, et enregistre les composants d'interop — là où les questions de *bitness* et d'enregistrement du module 16 arrivent à échéance.**

> 📦 On déploie ici une application **desktop Windows** (Windows Forms) ciblant **.NET Framework
> 4.7.2** — selon le modèle **classique** (installeur, runtime machine, enregistrement COM), et **non**
> le déploiement autonome du .NET moderne (ce serait une affaire du module 20). C'est aussi le point où
> les sujets d'**enregistrement** et de **bitness** des modules 16.1, 16.2 et 16.6 deviennent concrets.

---

## Le modèle de déploiement : un composant Windows, à l'échelle de la machine

Contrairement au .NET moderne (qui peut **embarquer** son runtime avec l'application), .NET Framework  
est un **composant de Windows**, installé **à l'échelle de la machine**. Votre application 4.7.2 a donc  
besoin que le **runtime** soit **présent** sur le poste cible. Le packaging consiste à livrer, de façon  
reproductible : l'**application** + ses **dépendances** + ce qu'il faut **enregistrer**.

---

## Le prérequis : le runtime 4.7.2 (ou supérieur 4.x)

- **Rétrocompatibilité 4.x.** Les versions de .NET Framework 4.x sont des mises à jour **« en place »** :
  **4.8** (et 4.8.1) **remplacent** 4.7.2 sur une même machine, et une application ciblant 4.7.2
  **s'exécute sur 4.8/4.8.1**. Le prérequis n'est donc pas « exactement 4.7.2 », mais **« 4.7.2 ou
  supérieur »**.
- **Présence sur Windows récent.** Windows 10 (à partir d'une certaine version) et Windows 11 livrent
  **4.8** d'origine ; le runtime est donc **souvent déjà là**. Les cibles plus anciennes (Windows 7 SP1,
  8.1, Server anciens) peuvent nécessiter l'installation du **redistribuable** 4.7.2 (ou 4.8).
- **Web vs hors-ligne.** Microsoft fournit un installeur **web** (petit, télécharge pendant
  l'installation — internet requis) et un installeur **hors-ligne/autonome** (complet, sans internet —
  **préférable** en environnement d'entreprise maîtrisé).
- **Détection de présence.** La version installée est inscrite au registre sous
  `HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full`, valeur **`Release`** (DWORD). Le test
  « 4.7.2 ou supérieur » consiste à vérifier **`Release >= 461808`**.
- **Chaînage automatique.** En pratique, on **ne code pas** cette détection à la main : un installeur
  correct **chaîne** le prérequis — il vérifie la présence du runtime et, à défaut, **installe le
  redistribuable avant** l'application. Les outils d'installation (ci-dessous) disposent d'un mécanisme
  de **prérequis** pour cela.

---

## Choisir un mécanisme de packaging

| Mécanisme | Atouts | Limites / quand l'éviter |
|-----------|--------|--------------------------|
| **Installeur MSI** (WiX, VS Installer Projects, InstallShield, Inno Setup…) | Contrôle complet ; **enregistrement COM** ; installation **machine** ; chaînage des prérequis | Plus de mise en place |
| **ClickOnce** | Publication simple, **mise à jour automatique**, installation **par utilisateur** (sans droits admin) | ⚠️ **Mal adapté à l'enregistrement COM machine** → conflit avec une app d'interop |
| **MSIX** | Format **moderne** (installation propre, **isolation**, mise à jour intégrée) | ⚠️ COM **confiné au paquet** (*Packaged COM* en emplacement **privé**) → **invisible** d'un hôte VB6 **externe** — même limite que ClickOnce pour l'interop |
| **xcopy / manuel** | Trivial pour un outil interne simple | Pas d'enregistrement COM, ne passe pas à l'échelle |

> 🎯 **Recommandation pour une application migrée** : un **installeur** (WiX — le standard de fait pour
> MSI, scriptable et adapté à l'intégration continue ; ou *Visual Studio Installer Projects* pour un cas
> simple). En effet, une migration incrémentale s'appuie sur de l'**interop COM** (modules 16.1/16.2),
> qui exige un **enregistrement machine** — précisément ce que **ClickOnce** (par utilisateur, sans
> élévation) gère mal. ClickOnce reste séduisant pour sa **mise à jour automatique** quand il n'y a
> **pas** de composant COM à enregistrer.

---

## Ce qui entre dans le paquet

- L'**exécutable** et vos **assemblies** (DLL de l'application).
- Les **bibliothèques tierces** non incluses dans le framework, et vos **composants CCW** (module 16.2).
- Les **assemblies d'interop** (`Interop.*.dll`, RCW du module 16.1) — **sauf** si l'option *« Incorporer
  les types d'interop »* a été utilisée (ils sont alors **embarqués**, rien à livrer). Idem pour les
  **PIA** non embarqués.
- Les **composants COM/ActiveX legacy** (OCX/DLL) dont l'application dépend, **à déployer et
  enregistrer** (module 16.1).
- Les fichiers de **configuration** (`app.config`), **ressources**, fichiers de **données**.
- D'éventuels **runtimes natifs** (par ex. un redistribuable Visual C++) requis par une dépendance.

---

## Enregistrer les composants COM à l'installation

C'est ici que l'**enregistrement** du module 16 se concrétise — et c'est une source fréquente d'échecs  
silencieux en production (« Classe non enregistrée »).

- **Côté RCW (16.1)** — un composant COM/OCX legacy que l'application consomme doit être **enregistré** :
  ```text
  regsvr32 MonComposant.ocx
  ```
- **Côté CCW (16.2)** — vos composants .NET exposés à du VB6 doivent l'être pour COM, via **`regasm`**
  (avec `/codebase` hors GAC, ou installation au **GAC** + `regasm` sans `/codebase`) :
  ```text
  regasm MaSociete.Composant.dll /tlb:MaSociete.Composant.tlb /codebase
  ```

> ⚠️ **Le piège de la *bitness* à l'enregistrement.** Un hôte VB6 est **32 bits**. L'enregistrement doit
> donc atterrir dans la **vue de registre 32 bits** (sous `WOW6432Node` sur un Windows 64 bits) :
> - utilisez le **`regsvr32` 32 bits** (`%SystemRoot%\SysWOW64\regsvr32.exe`) ;
> - utilisez le **`regasm` 32 bits** (dossier **`Framework`**, **pas** `Framework64`).
> Enregistrer **uniquement** côté 64 bits = composant **introuvable** par le processus VB6 (modules 16.2
> / 16.6).

> ⚠️ **Droits administrateur.** L'enregistrement COM machine (HKLM) exige une **élévation**. L'installeur
> d'une application d'interop requiert donc des droits **admin** — ce qui **confirme** l'inadéquation de
> ClickOnce (par utilisateur) pour ce cas.

> 💡 **Bonnes pratiques MSI** : capturer l'enregistrement COM dans les **tables de registre** du MSI est
> préférable à l'appel direct de `regsvr32`/*SelfReg* (déconseillé en MSI) ; en pratique, beaucoup
> appellent tout de même les outils. À l'autre extrémité, le **COM sans enregistrement** (*reg-free
> COM*, via manifestes côte-à-côte) **évite** l'enregistrement machine et ses soucis de *bitness* — mais
> sa mise en place est plus complexe.

---

## Configuration et build de *release*

- **Configuration *Release*** : optimisations activées ; ne pas livrer de symboles de débogage dans les
  binaires (ou livrer les **PDB** **à part** pour le diagnostic).
- **Cible de plateforme** : **AnyCPU** par défaut, mais **x86** si l'application charge du **COM 32 bits
  en processus** (la règle de *bitness* du module 16.1).
- **Signature** : **signer Authenticode** l'installeur et l'exécutable (certificat de **signature de
  code**) pour éviter les avertissements SmartScreen/UAC et établir la confiance éditeur ; **nom fort**
  (*strong name*) si installation au **GAC** (CCW).
- **`app.config`** : la cible y est déclarée, et des **redirections de liaison** peuvent être nécessaires
  si des dépendances référencent des versions différentes :
  ```xml
  <startup>
    <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.7.2" />
  </startup>
  ```
- Pour les **préférences utilisateur**, `My.Settings` écrit dans `user.config` (sous *AppData*) :
  **aucune** action d'installation requise (module 16.6).

---

## Tester le paquet sur une machine **propre**

> ⚠️ **Le piège « ça marche chez moi ».** Votre poste de développement a **déjà** le runtime, les
> dépendances et les composants COM **enregistrés** — il **masque** tous les problèmes de déploiement.
> (C'est l'avertissement du module 16 : *« ce qui marche sur ma machine peut échouer en production »*.)

Testez systématiquement l'installeur sur une **machine propre** (idéalement une **VM** réinitialisable)
**dépourvue** du runtime et de vos dépendances, pour vérifier réellement :

- le **chaînage du prérequis** 4.7.2 (installé si absent) ;
- l'**enregistrement COM** (RCW/CCW) et la **bonne *bitness*** ;
- les **versions d'OS cibles** et la **bitness** visées ;
- la **désinstallation** (suppression propre, **désenregistrement** COM) ;
- la **mise à jour** (installer la v2 par-dessus la v1).

---

## ⚠️ Pièges à connaître

- **Présumer le runtime présent** : le **vérifier/chaîner**, et **tester sur machine propre**. ⚠️
- **Oublier l'enregistrement COM** : les composants d'interop échouent silencieusement (« Classe non
  enregistrée ») sur la cible. ⚠️
- **Mauvaise *bitness* d'enregistrement** : enregistrer en **32 bits** (`WOW6432Node`, `SysWOW64`,
  `Framework`) pour les hôtes VB6/32 bits. ⚠️
- **ClickOnce + COM** : ClickOnce ne sait pas faire l'enregistrement COM machine → inadapté aux apps
  d'interop.
- **Assemblies d'interop / redirections de liaison manquantes** : échecs de chargement à l'exécution.
- **Absence de signature** : avertissements SmartScreen/UAC, blocages en entreprise.
- **`regsvr32`/*SelfReg* dans un MSI** : déconseillé (préférer les tables de registre du MSI).

---

## En résumé

- On déploie une application desktop **4.7.2** selon le modèle **classique** (installeur + runtime
  machine + enregistrement COM), pas en autonome.
- **Prérequis** : le **runtime 4.7.2 ou supérieur** (4.8 exécute les apps 4.7.2) ; souvent présent sur
  Windows récent, à **chaîner** via un redistribuable (web ou hors-ligne) sinon ; détection par
  `Release >= 461808`.
- **Mécanisme** : un **installeur** (WiX / VS Installer Projects) pour une app migrée — car l'interop
  exige un **enregistrement machine** que **ClickOnce** gère mal (mais qui offre la mise à jour
  automatique sans COM).
- **Enregistrement COM** : `regsvr32` (RCW, 16.1) et `regasm` (CCW, 16.2) — **impérativement en 32 bits**
  pour les hôtes VB6, avec **droits admin**.
- **Build de *release*** : *Release*, cible **x86** si COM 32 bits, **signature** (Authenticode + nom
  fort pour le GAC), `app.config` (cible + redirections).
- **Tester sur une machine propre** : c'est la seule façon de débusquer les problèmes de runtime,
  d'enregistrement et de *bitness* que le poste de développement **cache**.

> ➡️ **Le paquet est prêt — mais on ne déploie pas tout d'un coup. Place à la cohabitation VB6/.NET en
> production : faire coexister l'ancien et le nouveau pour un déploiement par étapes.** (Module 18.2)

---

🏷️ **Indicateurs** : packaging / déploiement · concrétise l'**enregistrement** et la **bitness** des modules 16.1/16.2/16.6 ; lié à 16.6 (`My.Settings`)
**Cible** : .NET Framework 4.7.2 (ou supérieur 4.x) · Visual Studio 2026 · WiX / VS Installer Projects · `regasm` / `regsvr32`

⏭️ [Cohabitation VB6 / .NET pendant la transition (déploiement progressif)](/18-deploiement-bascule/02-cohabitation.md)
