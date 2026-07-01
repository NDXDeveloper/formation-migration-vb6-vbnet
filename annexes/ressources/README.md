🔝 Retour au [Sommaire](/SOMMAIRE.md)

# H. Ressources et outils

Documentation officielle, **support du runtime VB6 sur Windows**, outils de migration, communauté.

> ⏳ **Avertissement de péremption.** Les liens, versions et offres commerciales **évoluent**. Les
> informations ci-dessous ont été vérifiées **à la date de cette formation (juin 2026)** ; **vérifier
> l'actualité** avant de s'y fier pour une décision. En cas de doute, partir des **domaines
> officiels** (`learn.microsoft.com`, `dotnet.microsoft.com`) et rechercher le sujet précis.

---

## 1. Documentation officielle Microsoft

### VB6 et migration
- **Énoncé de support de Visual Basic 6.0 sur Windows** — la référence sur ce qui est (et n'est pas)
  supporté : <https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/visual-basic-6-support-policy>
- **Annonce de support VB6 (Microsoft Lifecycle)** (FR) — l'IDE n'est plus supporté ; recommandation
  de migration : <https://learn.microsoft.com/fr-fr/lifecycle/announcements/visual-basic-6-support-announcement>
- **Partenaires de migration VB6 (Microsoft Learn)** — pages décrivant les outils commerciaux
  reconnus (ex. Mobilize.Net) :
  <https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/vb6-partners-mobilize-net>

### .NET, VB.NET, Windows Forms, interop
- **Documentation .NET / Visual Basic** — référence du langage, du runtime et des API :
  <https://learn.microsoft.com> (rechercher « Visual Basic language reference », « Windows Forms »,
  « P/Invoke », « COM interop », « Marshal class »).
- **Cycle de vie de .NET Framework** — politique officielle (composant de Windows) :
  <https://dotnet.microsoft.com/en-us/platform/support/policy/dotnet-framework>
- **Cycle de vie de .NET (Core et moderne)** — règles LTS/STS :
  <https://dotnet.microsoft.com/en-us/platform/support/policy/dotnet-core>
- **Microsoft .NET Framework — Lifecycle** (dates de retrait par version) :
  <https://learn.microsoft.com/en-us/lifecycle/products/microsoft-net-framework>

> 💡 Pour l'interop et le marshaling (module 16), chercher sur `learn.microsoft.com` :
> **`DllImportAttribute`**, **`StructLayoutAttribute`**, **`MarshalAsAttribute`**, **`Marshal`**,
> **« Runtime Callable Wrapper »**, **« COM Callable Wrapper »**.

---

## 2. Support du runtime VB6 sur Windows *(à connaître)*

Synthèse de la **politique officielle** (voir §1) — un point d'ancrage important du cadrage
(module 1.1) :

- **Le runtime VB6 est livré avec Windows** et **supporté pour la durée de vie** de la version de
  Windows dans laquelle il est fourni. En pratique : **aucune installation séparée** n'est
  nécessaire pour *exécuter* une application VB6 sur Windows 10/11.
- Le **niveau de support** se limite aux **régressions graves** et aux **failles de sécurité
  critiques** des applications **existantes** — il ne s'agit **pas** d'un développement actif.
- **L'IDE VB6 (et VB6 development) n'est plus supporté** depuis le **8 avril 2008**. Il n'existe pas
  de moyen supporté de **créer ou maintenir** des applications VB6.
- Le runtime est **32 bits uniquement** : sur Windows 64 bits, les applications VB6 s'exécutent dans
  l'**émulation WOW64**.
- ⚠️ Les **contrôles tiers OCX/ActiveX** ne sont **pas** couverts par le support Microsoft : pour
  ceux-ci, **s'adresser à l'éditeur d'origine**. La plupart des problèmes de portage tiennent
  d'ailleurs à des **OCX manquants**, pas au runtime de base.

> 🧭 **Conséquence pour la formation.** Ce support « *It Just Works* » du runtime explique pourquoi
> une migration peut être **planifiée sans urgence absolue** côté exécution — mais l'**IDE hors
> support**, la **pénurie de compétences** et la **dette technique** restent les vrais moteurs
> (module 1.1). C'est aussi pourquoi cibler **.NET Framework 4.7.2** (un runtime, lui, **supporté et
> proche de VB6**) constitue un **« pont »** à faible risque (module 1.2).

---

## 3. Outils de migration

> Panorama au module 4 (grille de décision : 4.5). Rappel : **aucun outil ne migre tout
> automatiquement** ; les assistants traitent 60-95 % du mécanique, le **jugement humain** (et l'IA
> encadrée, module 19) traite le reste — surtout les **pièges silencieux** (Annexe B) et l'UI.

### Outils Microsoft
- **Assistant de mise à niveau Visual Studio (*Upgrade Wizard*)** — l'historique convertisseur
  intégré, **présent jusqu'à Visual Studio 2008** puis **retiré** des versions suivantes. Utile
  surtout pour comprendre l'historique ; capacités **et limites** au module 4.1.
- **.NET Upgrade Assistant** — outil **gratuit** de Microsoft, orienté modernisation .NET, qui
  **assiste** aussi des scénarios VB6 → .NET. À chercher sur `learn.microsoft.com`
  (« .NET Upgrade Assistant »).

### Outils commerciaux *(les deux références du marché)*
- **VBUC — Visual Basic Upgrade Companion**, par **Mobilize.Net** (désormais rattaché à
  **GAPVelocity AI**). Convertit VB6 → **VB.NET ou C#**, transforme **ADO/DAO/RDO → ADO.NET**,
  **résout les types en liaison tardive**, gère `GoSub`/`On Error`, vise un haut niveau
  d'automatisation. Produit **du code .NET natif** (sans dépendance à un runtime de compatibilité).
  - Produit : <https://www.mobilize.net/products/app-migrations/vbuc> · essai :
    <https://www.mobilize.net/products/app-migrations/vbuc/free-trial>
  - Offre VB6 (GAPVelocity) : <https://www.gapvelocity.ai/migrate/vb6>
  - Exemples avant/après : <https://github.com/MobilizeNet/VBMigration>
  - Extension Visual Studio : <https://marketplace.visualstudio.com/items?itemName=MobilizeNET.VBUC>
- **VB Migration Partner (VBMP)**, par **Code Architects** (équipe menée par **Francesco Balena**).
  Met l'accent sur la **fidélité** (couverture quasi exhaustive des contrôles/fonctions VB6), les
  **pragmas de migration** et le cycle **convert-test-fix** — qui permet de **continuer à faire
  évoluer** le projet VB6 pendant la migration. Génère notamment de la **finalisation déterministe**
  (`IDisposable`), convertit `On Error` → `Try/Catch`, transforme `GoSub` en méthodes et **détecte
  les `ByRef` convertibles en `ByVal`**.
  - Éditeur / produit : <https://codearchitects.com/en/software-development/visual-basic-applications-migration/>
    · site VBMP : <https://www.vbmigration.com>
  - Exemples de conversion : <https://github.com/codearchitects/vbmigration-code-samples>
  - Extension Visual Studio :
    <https://marketplace.visualstudio.com/items?itemName=CodeArchitects.VBMigrationPartner>

> ⚠️ **La « béquille » à connaître** : l'espace `Microsoft.VisualBasic.Compatibility` (émulation de
> comportements VB6). Pratique pour démarrer, mais à **retirer** ensuite (modules 4.2 et 17.2).
> Les bons outils commerciaux s'en passent et produisent du **code natif**.

---

## 4. Versions et cycle de vie *(pour le module 20)*

- **.NET Framework 4.7.2** — la **cible** de cette formation. En tant que **composant de Windows**
  (depuis 4.5.2), il suit le **cycle de vie de Windows** : supporté tant qu'il est installé sur une
  version de Windows supportée.
- **.NET Framework 4.8 / 4.8.1** — la **dernière** version de la famille .NET Framework. **Gelée
  fonctionnellement** (pas de nouveautés depuis la 4.8.1, août 2022) ; supportée, là encore, via le
  cycle de Windows. Un projet peut légitimement viser **4.7.2 ou 4.8** selon ses contraintes.
- **.NET moderne — .NET 10 (LTS)** — l'actuelle version à **support long** (cadence : une version
  majeure par an en novembre ; **LTS = 3 ans**, **STS = 2 ans**). C'est la **suite optionnelle**
  visée par le module 20 (« VB.NET avec .NET 10 LTS »). Y aller est une **migration distincte**, à
  faire **après** 4.7.2 et **seulement si nécessaire**.

> 🔗 Rappel de cadrage : faire **deux sauts séparés** (VB6 → 4.7.2, puis 4.7.2 → .NET moderne) plutôt
> qu'un seul **multiplie les chances de succès** (modules 1.2 et 20.1).

---

## 5. Communauté et ressources d'écosystème

- **Microsoft Q&A** — questions/réponses officielles (balises Visual Basic, .NET) :
  <https://learn.microsoft.com/en-us/answers/>
- **VBForums** — communauté historique VB6/VB.NET (forums dédiés à la migration) :
  <https://www.vbforums.com>
- **Stack Overflow** — étiquettes `vb6`, `vb.net`, `winforms`, `interop`, `vba` :
  <https://stackoverflow.com>
- **Dépôts d'exemples de migration** (avant/après) :
  - Mobilize.Net : <https://github.com/MobilizeNet/VBMigration>
  - Code Architects : <https://github.com/codearchitects/vbmigration-code-samples>
- **Ouvrage de référence** — *« Programming Microsoft Visual Basic »* (Francesco Balena, Microsoft
  Press) : un classique sur le langage VB, utile pour **comprendre les comportements VB6** que l'on
  cherche à préserver.

> 🤖 **Et l'IA ?** Les assistants (dont Claude) **aident** à convertir un module, **expliquer** un
> comportement VB6 obscur, ou **repérer** des pièges — mais **validation systématique** : ils
> inventent parfois de **faux équivalents** VB6/VB.NET (module 19, et ses limites en 19.5). Toujours
> préciser **« VB6 »** vs **« VB.NET »** et la cible **.NET Framework 4.7.2** dans les invites.

---

## 6. 🔗 Renvois internes

- **Module 1.1-1.2** — Pourquoi migrer ; pourquoi 4.7.2 (support du runtime, le « pont »).
- **Module 4** — Outils de migration (Upgrade Wizard, outils commerciaux, grille de décision).
- **Module 16** — Interop COM & API (documentation P/Invoke, marshaling, RCW/CCW).
- **Module 19** — Migration assistée par IA (et ses limites).
- **Module 20** — De 4.7.2 vers .NET moderne (.NET 10 LTS).
- **Annexe B** — Pièges silencieux (ce que **aucun** outil ne garantit de traiter seul).
- **Annexe G** — Glossaire (définitions des termes et acronymes cités ici).

---

**Juin 2026**  
**Cible** : .NET Framework 4.7.2 (le « pont » depuis VB6)  
**Suite optionnelle** : migration 4.7.2 → .NET 10 LTS (formation dédiée)  
**Licence** : Creative Commons BY-NC-SA 4.0

🔝 [Retour au Sommaire](/SOMMAIRE.md) · 🏁 **Fin de la formation**
