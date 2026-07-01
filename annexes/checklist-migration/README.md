🔝 Retour au [Sommaire](/SOMMAIRE.md)

# E. Checklist de migration (avant / pendant / après)

**Liste de vérification opérationnelle, étape par étape, pour ne rien oublier sur un projet réel.**

> 🧭 **Comment utiliser cette checklist.**
> - Trois phases : **🟦 Avant** (préparer le terrain) → **🟨 Pendant** (migrer) → **🟩 Après**
>   (valider, fiabiliser, déployer).
> - Cocher au fur et à mesure ; **adapter** à votre contexte (tout ne s'applique pas à chaque projet).
> - Les repères : ⭐ **à ne jamais sauter** · ⚠️ **garde-fou anti-piège silencieux** (Annexe B) ·
>   🔗 **interop / cohabitation**.
> - **Fil rouge** : le **golden master** (modules 5.5 et 17.3) est le juge de paix. Tant qu'il n'est
>   pas vert, la migration n'est pas finie.

| Phase | Objectif | Modules clés |
|---|---|---|
| 🟦 **Avant** | Cadrer, outiller, réduire les risques *en amont* | 1-6 |
| 🟨 **Pendant** | Convertir langage, interface, données — sans changer le comportement | 5-16 |
| 🟩 **Après** | Prouver l'iso-comportement, fiabiliser, déployer | 17-20 |

---

## 🟦 AVANT — préparer le terrain

### Cadrage et évaluation *(modules 1-3)*
- [ ] **Inventaire complet** : formulaires `.frm`, modules `.bas`, classes `.cls`, références, OCX, DLL, déclarations d'API. *(3.1)*
- [ ] **Cartographie des dépendances** : COM, ActiveX, bases de données, composants tiers. *(3.2)* 🔗
- [ ] **Estimer l'effort** : taille, complexité, points durs identifiés. *(3.3)*
- [ ] **Repérer les composants sans équivalent direct** : tableaux de contrôles, `Line`/`Shape`/`Data`/`OLE`, `DriveListBox`/`DirListBox`/`FileListBox`. *(Annexe C)* ⚠️
- [ ] **Choisir la stratégie** : *big-bang* vs incrémentale. *(3.4)*
- [ ] Si incrémentale : **décider de la cohabitation COM** pendant la transition. *(3.5)* 🔗
- [ ] **Définir le périmètre, les critères de réussite et les indicateurs.** *(3.6)*

### Outils *(module 4)*
- [ ] **Choisir l'approche** : assistant VS / outil commercial (VBUC, VB Migration Partner) / manuel / hybride / IA. *(4.5)* 🛠️
- [ ] Si outil commercial : **vérifier sa disponibilité et sa licence.** *(4.3)*
- [ ] **Politique vis-à-vis de `Microsoft.VisualBasic.Compatibility`** : décider de **l'éviter** (béquille à retirer). *(4.2)* ⚠️
- [ ] Si IA envisagée : cadrer son usage et la **validation systématique**. *(module 19)* 🤖

### Préparer le code VB6 *(module 5)*
- [ ] **Mettre le source VB6 sous contrôle de version** (point de retour avant toute modification).
- [ ] **Nettoyer** : supprimer le code mort, `Option Explicit` **partout**, retirer le `Variant` superflu. *(5.1)*
- [ ] **Réduire les pièges en amont** : typage explicite, **`ByVal`/`ByRef` explicites côté VB6**. *(5.2)* ⭐ ⚠️
- [ ] **Découpler l'UI de la logique métier** (faciliter la migration progressive). *(5.3)*
- [ ] **Documenter les comportements critiques** : arrondis monétaires, dates, ordre des événements. *(5.4)*
- [ ] **Mettre en place le harnais de tests de référence (*golden master*).** *(5.5)* ⭐
- [ ] **Geler le périmètre fonctionnel** (aucune évolution pendant la migration). *(5.6)* ⭐

### Préparer l'environnement .NET *(module 6)*
- [ ] **Installer/configurer Visual Studio**, **cibler .NET Framework 4.7.2**. *(6.1)*
- [ ] **Définir la structure** de solution et de projets (`.frm`/`.bas`/`.cls` → `.vb`). *(6.2)*
- [ ] **Référencer les composants COM/ActiveX** nécessaires (RCW, *Primary Interop Assemblies*). *(6.3)* 🔗
- [ ] **Fixer les options de compilation initiales** : `Option Strict` **Off** au départ (toléré pour le *late binding* hérité), `Explicit` On. *(6.4)* ⚠️

---

## 🟨 PENDANT — migrer (sans changer le comportement)

### Discipline générale *(transversal)*
- [ ] **Compiler tôt et souvent** ; corriger les erreurs au fur et à mesure.
- [ ] **Committer par petits incréments** (contrôle de version).
- [ ] **Ne PAS introduire d'évolutions fonctionnelles** (périmètre gelé). ⭐
- [ ] **Tenir un journal des décisions** de migration (correspondances non triviales).
- [ ] **Annoter les zones douteuses** (`' TODO migration`) pour la phase de validation.
- [ ] Si stratégie incrémentale : **maintenir la cohabitation COM**. *(3.5, 18.2)* 🔗
- [ ] **Ne pas activer `Option Strict On` trop tôt** (réservé à la phase « après »). ⚠️

### Migration du langage *(modules 5-12, Annexes A/B/D)*
- [ ] **Mapper les types PAR TAILLE** : `Integer`→`Short`, `Long`→`Integer`, `Currency`→`Decimal`. *(Annexe D)* ⭐ ⚠️
- [ ] **Expliciter `ByVal`/`ByRef`** sur chaque procédure. *(Annexe B.1)* ⭐ ⚠️
- [ ] **Convertir `Type` → `Structure`** (avec `<StructLayout>` si interop). *(Annexe A, B.2)* 🔗
- [ ] **Remplacer `Set`/`Let`** et **expliciter les propriétés** (`.Text`, `.Value`…). *(Annexe B.7)* ⚠️
- [ ] **Convertir `Property Get/Let/Set` → `Property` (`Get`/`Set`)**. *(Annexe A §6)*
- [ ] **Remplacer les formes retirées** : `GoSub`/`Return`, `While`/`Wend`, `On…GoTo`, `Eqv`/`Imp`. *(Annexe A §3-4)*
- [ ] **Normaliser les tableaux en 0-based** ; supprimer `Option Base`. *(Annexe B.4)* ⚠️
- [ ] **Traiter chaque `On Error Resume Next`** individuellement (ne **pas** traduire en bloc). *(Annexe B.6)* ⭐ ⚠️
- [ ] **Convertir `On Error` → `Try/Catch`** là où pertinent (sans mélanger les deux modèles). *(2.6)*
- [ ] **Remplacer `Class_Terminate` → `IDisposable`/`Dispose` + `Using`** au point d'appel. *(Annexe B.3)* ⭐ ⚠️
- [ ] **Traiter le `Variant`** : `IsEmpty`/`IsNull`/`IsMissing` → `Is Nothing`/`IsDBNull`/valeur par défaut. *(Annexe B.5)* ⚠️
- [ ] **Vérifier les conversions sensibles à la culture** (`CDbl`/`CDate`/`Format` → invariant pour les données). *(Annexe B.10)* ⚠️
- [ ] **Vérifier les conversions booléennes** (`True = -1` ; cohérence `CInt` vs framework). *(Annexe B.8)* ⚠️
- [ ] **Convertir les fonctions intrinsèques** vers leurs équivalents (chaînes, dates, maths, E/S). *(Annexe A §8)*
- [ ] **API Win32** : `Declare` → `P/Invoke` (`<DllImport>`), gérer **ANSI/Unicode** et le marshaling. *(16.4-16.5)* ⚠️ 🔗
- [ ] **`FileSystemObject` → `System.IO`** ; objet `App`, registre. *(16.6)*

### Migration de l'interface *(modules 13-14, Annexe C)*
- [ ] **`Caption` → `Text`** partout. *(Annexe C §1)*
- [ ] **Convertir les coordonnées twips → pixels** (le `ScaleMode` disparaît). *(Annexe C §1)*
- [ ] **Reconstruire les tableaux de contrôles** : gestionnaire **partagé** + `sender`/`.Tag`. *(Annexe C §3)* ⭐
- [ ] **Remplacer les contrôles sans équivalent** : `Line`/`Shape` → GDI+, `Data`/`OLE` → composants .NET, `Drive`/`Dir`/`FileListBox` → dialogues/`TreeView`/`ListView`. *(Annexe C §4)* ⚠️
- [ ] **Remplacer `CommonDialog`** par les classes dédiées (`OpenFileDialog`, `ColorDialog`…). *(Annexe C §6)*
- [ ] **Couleurs `OLE_COLOR` → `Color`** (`SystemColors`, `ColorTranslator.FromOle`). *(Annexe C §7)*
- [ ] **Convertir le dessin** (`Line`/`Circle`/`PSet`/`Print`) → **GDI+ dans `Paint`** (dessin non persistant !). *(Annexe C §8)* ⚠️
- [ ] **Convertir les menus** (Éditeur de menus → `MenuStrip`/`ContextMenuStrip`). *(Annexe C §4)*
- [ ] **Adapter le cycle de vie du formulaire** : `QueryUnload` → `FormClosing` ; `Show vbModal` → `ShowDialog()`. *(Annexe C §2)*
- [ ] **Mapper les événements** : `Change` → spécifique, `GotFocus` → `Enter`, `KeyPress` `KeyAscii` → `e.KeyChar`. *(Annexe C §2)* ⚠️

### Migration des données & COM *(modules 15-16)*
- [ ] **Remplacer le contrôle `Data` / la liaison** (DAO/RDO/ADO) → **ADO.NET**, `BindingSource`, `DataGridView`. *(15, Annexe C §10)*
- [ ] **Gérer *early* vs *late binding*** (où c'est nécessaire, où l'éliminer). *(16.3)* 🔗
- [ ] **Vérifier les composants COM référencés** (RCW/PIA) à l'exécution. *(6.3)* 🔗

---

## 🟩 APRÈS — valider, fiabiliser, déployer

### Validation (non-régression) *(module 17)*
- [ ] **Exécuter les tests de non-régression contre le golden master.** *(17.3)* ⭐
- [ ] **Vérifier les comportements critiques documentés** (arrondis, dates, ordre des événements) par comparaison. *(5.4 → 17.3)*
- [ ] **Tester sous au moins deux cultures** (ex. `en-US` et `fr-FR`). *(Annexe B.10, 17.3)* ⚠️

### Traque des pièges silencieux *(module 17.4)*
- [ ] **Revue ciblée** : `ByRef`, entiers redimensionnés, finalisation, bornes de tableaux, `Variant`, `On Error Resume Next`, propriétés par défaut, `True = -1`, dates, culture. *(17.4, Annexe B)* ⭐ ⚠️

### Fiabilisation et durcissement *(module 17)*
- [ ] **Activer `Option Strict` progressivement** ; éliminer le *late binding* résiduel. *(17.1)* ⭐
- [ ] **Supprimer les dépendances à `Microsoft.VisualBasic.Compatibility`.** *(17.2)* ⚠️
- [ ] **Vérifier la libération des ressources** : `IDisposable`, `Using`, fuites de *handles*. *(17.6)*
- [ ] **Revue de performance et de mémoire.** *(17.6)*
- [ ] *(Optionnel)* **Refactoring idiomatique** : LINQ, génériques, espace `My`, `Anchor`/`Dock` à la place du `Resize` manuel. *(17.5)*

### Déploiement et bascule *(module 18)*
- [ ] **Packaging .NET Framework** : installeur, **prérequis 4.7.2**, redistribuables. *(18.1)*
- [ ] **Stratégie de bascule** : pilote, *parallel run*, **plan de *rollback***. *(18.3)* ⭐
- [ ] **Cohabitation pendant la transition** / déploiement progressif. *(18.2)* 🔗
- [ ] **Documentation et transfert de compétences** à l'équipe. *(18.4)*

### Clôture et suite *(module 20)*
- [ ] **Archiver** le golden master, le journal des décisions et la documentation de migration.
- [ ] **Décider du second saut** éventuel vers .NET moderne (8/10) — **ou** assumer de **rester sur 4.7.2** (choix légitime). *(20)*

---

## 🚦 Les 10 points à ne jamais sauter

> Si le temps manque, ce sont **ceux-là** qu'il ne faut surtout pas omettre.

1. ⭐ **Geler le périmètre fonctionnel** avant de commencer. *(5.6)*
2. ⭐ **Mettre en place le golden master** *avant* la migration. *(5.5)*
3. ⭐ **Expliciter `ByVal`/`ByRef`** (idéalement dès le source VB6). *(B.1)*
4. ⭐ **Mapper les types par taille** (`Integer`→`Short`, `Long`→`Integer`, `Currency`→`Decimal`). *(B.2, D)*
5. ⭐ **Remplacer `Class_Terminate` par `IDisposable`/`Using`.** *(B.3)*
6. ⭐ **Traiter chaque `On Error Resume Next`** (jamais en bloc). *(B.6)*
7. ⭐ **Reconstruire les tableaux de contrôles** proprement. *(C §3)*
8. ⭐ **Tester sous plusieurs cultures.** *(B.10)*
9. ⭐ **Activer `Option Strict` progressivement** à la fin. *(17.1)*
10. ⭐ **Prévoir un plan de *rollback*** pour la bascule. *(18.3)*

---

## 🔗 Annexes et modules liés

- **Annexe A** — Correspondance VB6 → VB.NET (langage, fonctions, opérateurs).
- **Annexe B** — Catalogue des pièges silencieux (le « pourquoi » de chaque ⚠️ ci-dessus).
- **Annexe C** — Correspondance des contrôles (interface).
- **Annexe D** — Correspondance des types (tailles, plages, conversions).
- **Annexe F** — Modèles de tests *golden master* (outiller le fil rouge).
- **Modules 5-6** — Préparer le code et l'environnement. **17-18** — Valider et déployer. **20** — Et après ?

---

**Juin 2026**  
**Cible** : .NET Framework 4.7.2 (le « pont » depuis VB6)  
**Licence** : Creative Commons BY-NC-SA 4.0

⏭️ [Modèles de tests de non-régression (*golden master*)](/annexes/golden-master/README.md)
