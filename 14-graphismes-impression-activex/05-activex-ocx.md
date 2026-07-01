🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.5 Contrôles **ActiveX/OCX** : réutiliser via interop ou remplacer 🔗

> Beaucoup d'interfaces VB6 reposent sur des contrôles **ActiveX/OCX** — ceux de Microsoft (grilles,
> arborescences, barres d'outils, calendriers…) comme des **composants tiers**. Ces contrôles sont
> des composants **COM**, souvent **32 bits**, parfois **sous licence**, parfois fournis par un
> **éditeur disparu**. C'est pourquoi cette section relève d'une **décision d'architecture**, pas
> d'une traduction de syntaxe — et c'est fréquemment **le vrai point dur** d'une migration.

**Module 14 — Graphismes, impression et contrôles ActiveX** · Cible : .NET Framework 4.7.2 ·
interop COM (`AxHost`, RCW)

---

## 1. Pourquoi c'est souvent **LE** point dur

Un OCX n'est pas du code source que l'on convertit : c'est un **binaire COM** avec son lot de  
contraintes.

- **Enregistrement machine** : il doit être **enregistré** (`regsvr32`) sur chaque poste cible.
- **32 bits** : la plupart sont **x86 uniquement** → ils **interdisent** une exécution 64 bits.
- **Licence de conception** : certains exigent une **clé de licence** présente sur la machine.
- **Pérennité** : éditeur disparu, plus de support, comportement non garanti sur Windows récents.

> ⚠️ Conséquence : avant tout code, **inventoriez** les dépendances OCX (cf. modules 3.1 et 3.2). Pour
> chacune, la question n'est pas « comment la convertir ? » mais « **la réutiliser ou la remplacer ?**».

---

## 2. Deux voies

| | **Voie A — Réutiliser** (interop COM) | **Voie B — Remplacer** (contrôle natif) |
|---|---------------------------------------|------------------------------------------|
| Principe | Héberger l'OCX via un wrapper `AxHost` | Réécrire avec un contrôle Windows Forms |
| Effort court terme | **Faible** | Moyen à élevé |
| Pérennité | **Faible** (couche COM conservée) | **Élevée** (citoyen .NET de plein droit) |
| 64 bits | ❌ Impossible (force x86) | ✅ Possible |
| Déploiement | Enregistrement OCX requis | Aucun |
| Risque | Licence, éditeur, quirks du wrapper | Comportement à revalider |

La bonne réponse est souvent **un mélange** (§7) : remplacer ce qui a un équivalent évident,  
réutiliser temporairement le reste pour **réduire le risque**, puis remplacer plus tard.

---

## 3. Voie A — Réutiliser via interop COM (`AxHost`)

### Comment ça marche

.NET héberge un contrôle ActiveX grâce à un **wrapper** dérivé de `System.Windows.Forms.AxHost`.
Quand vous ajoutez un OCX (depuis le concepteur Visual Studio, ou en ligne de commande), **deux  
assemblys** sont générés :

| Assembly généré | Contenu | Outil |
|-----------------|---------|-------|
| `Interop.<Lib>.dll` | Le **RCW** (interfaces COM, coclasse) | `tlbimp` |
| `AxInterop.<Lib>.dll` | La classe **`Ax<Contrôle>`** (un **vrai contrôle WinForms** dérivant d'`AxHost`) | `aximp` |

```text
aximp Grille.ocx     ' -> AxInterop.Grille.dll (Ax... : AxHost) + Interop.Grille.dll (RCW)
```

Sur le formulaire, on manipule alors `AxMSFlexGrid`, `AxRichTextBox`, etc.

> 🔗 C'est de l'**interopérabilité COM** au sens du **module 16** (RCW). L'état de configuration du
> contrôle (ses propriétés posées dans le concepteur) est **sérialisé dans le `.resx` du formulaire**
> sous forme d'un blob `AxHost.State` — un lien direct avec la **section 14.3**.

### Les quirks du wrapper (à connaître)

L'hébergement n'est **pas** transparent :

- **Événements renommés** : un événement qui entre en collision avec un membre du contrôle reçoit le
  suffixe `Event`.
  ```vb
  Private WithEvents axGrille As AxMSFlexGrid
  Private Sub axGrille_Click(sender As Object, e As EventArgs) Handles axGrille.ClickEvent
      ' ...
  End Sub
  ```
- **`Variant` → `Object`** : beaucoup de membres OCX renvoient `Object` → **transtypages** requis
  (et `Option Strict On` les signalera, cf. modules 16.3 et 17.1).
  ```vb
  Dim valeur As String = CStr(axGrille.TextMatrix(1, 2))   ' TextMatrix renvoie Object
  ```
- **Types image/police COM** : les propriétés `Picture`/`Font` deviennent des types `stdole`
  (`StdPicture`/`IPictureDisp`, `StdFont`/`IFontDisp`) → conversions vers `System.Drawing.Image`/`Font`
  via un helper d'interop (lien section 14.3).
- **Contraintes d'exécution** : cible de plateforme **x86** obligatoire (§6), enregistrement de
  l'OCX, et **licence** éventuelle (§5).

---

## 4. Voie B — Remplacer par un contrôle .NET natif

C'est la voie **propre et pérenne**. Beaucoup de contrôles VB6 ont un **équivalent natif** (sans  
aucune couche COM) :

| Contrôle VB6 (OCX) | Équivalent .NET natif |
|--------------------|------------------------|
| `TreeView` / `ListView` (MSCOMCTL) | `TreeView` / `ListView` |
| `Toolbar` / `StatusBar` | `ToolStrip` / `StatusStrip` |
| `ProgressBar` / `Slider` / `TabStrip` | `ProgressBar` / `TrackBar` / `TabControl` |
| `ImageList` | `ImageList` |
| `DTPicker` / `MonthView` / `UpDown` (MSCOMCT2) | `DateTimePicker` / `MonthCalendar` / `NumericUpDown` |
| `RichTextBox` (RICHTX32) | `RichTextBox` |
| `CommonDialog` (COMDLG32) | `OpenFileDialog`/`SaveFileDialog`/`ColorDialog`/`FontDialog`/`PrintDialog` (cf. 13.8, 14.4) |
| `SSTab` (TABCTL32, Sheridan) | `TabControl` |
| `MSFlexGrid` / `MSHFlexGrid` | `DataGridView` (éditable et plus riche) |
| `DataGrid` / `DataList` / `DBGrid` | `DataGridView` + `BindingSource` (cf. 15.4) |
| `MSChart` | `Chart` (`System.Windows.Forms.DataVisualization.Charting`) |
| `Winsock` (MSWINSCK) | `System.Net.Sockets` (`TcpClient`/`Socket`/`UdpClient`) — **non visuel** |
| `Inet` (MSINET) | `HttpClient` / `WebClient` / `FtpWebRequest` — **non visuel** |
| `Animation` / `MMControl` (MCI32) | `System.Media`, `PictureBox` (GIF), API multimédia |

> ℹ️ Les trois dernières lignes ne sont pas des « contrôles » mais des composants **non visuels** :
> leur migration est un **refactoring de code**, pas un simple échange dans le concepteur.

**Avantages** : 64 bits possible, aucun enregistrement, support assuré, `DataBinding` natif, et un
terrain **prêt pour le saut vers .NET moderne** (module 20). **Coût** : reprendre le code lié à  
l'API du contrôle (noms de propriétés/événements différents, comportements parfois subtils).

> 🔗 L'**Annexe C** détaille la correspondance contrôle par contrôle, avec propriétés et événements
> équivalents — la référence à garder ouverte ici.

---

## 5. ⚠️ Le piège des **licences** (`licenses.licx` / LC.exe)

Les contrôles COM **sous licence de conception** doivent embarquer une **licence d'exécution** : .NET  
le fait via un fichier **`licenses.licx`** traité par **LC.exe** à la compilation. Il faut donc que  
le contrôle soit **correctement installé et licencié sur la machine de build**.

> ⚠️ Si la licence (fichier `.lic` ou clé de registre) **manque**, la compilation **ou** l'exécution
> **échoue** — un blocage fréquent et frustrant, surtout pour des composants tiers anciens dont on n'a
> plus l'installeur ni la clé. **Vérifiez la disponibilité des licences dès l'inventaire** : c'est
> parfois ce qui **impose** le remplacement (voie B).

---

## 6. Bitness et déploiement (renvoi module 18)

- **32 bits = verrou x86.** Héberger un OCX 32 bits **force** la cible de plateforme à **x86** (pas
  d'`AnyCPU` en 64 bits). C'est un piège **silencieux** : tout fonctionne… jusqu'au jour où l'on a
  besoin du 64 bits.
- **Enregistrement** : chaque poste client doit avoir l'OCX **enregistré** (`regsvr32`), ce qui
  alourdit l'installeur. Le **COM sans inscription** (*reg-free*, via manifestes côte à côte) évite
  l'enregistrement machine mais reste délicat — surtout avec des contrôles licenciés.

> 🔗 Ces aspects sont traités au **module 18** (packaging, prérequis, déploiement).

---

## 7. Grille de décision : interop **ou** remplacement

| Critère | Plutôt **remplacer** (natif) | Plutôt **réutiliser** (interop) |
|---------|------------------------------|----------------------------------|
| Équivalent .NET natif existe | ✅ Oui (common controls…) | — |
| Licence / contrôle indisponible | ✅ Obligé de remplacer | — |
| Besoin du **64 bits** | ✅ Oui | ❌ Interop force le x86 |
| Contrôle tiers **complexe**, sans équivalent | — | ✅ Au moins **d'abord** |
| Réduire le **risque court terme** | — | ✅ Cohabitation, remplacer plus tard |
| Pérennité / **saut .NET moderne** | ✅ Oui | ❌ Couche COM à reporter |

**Stratégie pragmatique** (cohérente avec la migration **incrémentale**, modules 3.4/3.5) :

1. **Remplacer tout de suite** les contrôles Microsoft à équivalent évident (common controls,
   dialogues, RichTextBox…).
2. **Réutiliser temporairement** via interop les contrôles tiers complexes, pour **dé-risquer** la
   bascule.
3. **Planifier leur remplacement** ultérieur, une fois l'application stabilisée (et avant un éventuel
   passage à .NET moderne).

---

## 8. Ce que font les outils (renvoi module 4)

- L'**assistant de mise à niveau** (`Upgrade Wizard`) **conserve** les ActiveX via les wrappers `Ax`
  (interop) : vous obtenez une application **qui marche mais reste COM**. Il ne **remplace pas** par
  des contrôles natifs.
- Les **outils commerciaux** (**VBUC** de Mobilize.Net, **VB Migration Partner** de Code Architects)
  vont plus loin : ils **mappent** une partie des contrôles courants vers des équivalents natifs,
  gèrent les **tableaux de contrôles** (cf. 13.6) et suppriment des dépendances de compatibilité.

---

## 9. 🧭 Approche recommandée

1. **Inventorier** chaque OCX (visuel et non visuel), sa **licence**, sa **bitness**, son éditeur
   (modules 3.1/3.2).
2. Trancher **par contrôle** avec la grille du §7 ; remplacer dès que l'équivalent natif existe
   (Annexe C).
3. Si interop : cibler **x86**, régler la **licence** (`licenses.licx`), prévoir le **déploiement**
   (module 18).
4. **Suivre** les contrôles laissés en interop comme **dette technique** à résorber.

---

## 10. ⚠️ Pièges silencieux à retenir

- **Verrou 32 bits** : héberger un OCX 32 bits impose la cible **x86** — invisible jusqu'à un besoin
  64 bits.
- **Licence manquante** (`.lic`/registre) → **build ou exécution en échec**.
- **Enregistrement par poste** (`regsvr32`) oublié → l'application plante chez le client.
- **Quirks du wrapper** : événements suffixés `Event`, `Variant`→`Object` (casts, `Option Strict`),
  propriétés `Picture`/`Font` en types COM `stdole`.
- **« Ça compile » ≠ « ça marche pareil »** : un ActiveX hébergé peut différer (focus, ordre de
  tabulation, rendu, modalité).
- **Tableaux de contrôles** d'OCX (cf. 13.6) → difficulté **cumulée**.
- **Conserver l'interop** **reporte** et **complique** le saut vers .NET moderne (module 20) ; le
  remplacement natif est plus **pérenne**.

> 🔗 Voir l'**Annexe C** (correspondance des contrôles), le **module 16** (RCW, *early/late binding*),
> le **module 17** (`Option Strict`), le **module 18** (déploiement) et le **module 3** (inventaire,
> migration incrémentale).

---

**Section suivante → 14.6 — Contrôles tiers courants (grilles, calendriers…) : stratégies de
remplacement**, pour les composants commerciaux qui n'ont pas d'équivalent direct.

⏭️ [Contrôles tiers courants (grilles, calendriers…) : stratégies de remplacement](/14-graphismes-impression-activex/06-controles-tiers.md)
