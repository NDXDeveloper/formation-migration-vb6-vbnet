🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.7 Reporting : **Crystal Reports** (compatibilité) et alternatives ⚠️

> Beaucoup d'applications VB6 impriment leurs factures, relevés et étiquettes via **Crystal Reports**,
> qui était **livré avec VB6**. Mais cette version **ancienne** n'est ni supportée ni compatible .NET :
> on ne « réutilise » pas l'OCX tel quel. Le vrai choix est : **conserver Crystal** (via *Crystal
> Reports for Visual Studio*, en réutilisant les `.rpt`) **ou** passer à une **alternative**. C'est une
> décision d'**architecture** *et* de **licence** — d'où l'indicateur ⚠️.

**Module 15 — Accès aux données** · Cible : .NET Framework 4.7.2 · reporting

---

## 1. Pourquoi le reporting est un point dur à part

La migration du reporting touche **trois** choses à la fois :

- les **fichiers `.rpt`** (les définitions de rapports : SQL, formules, sous-rapports, paramètres) —
  s'ouvrent-ils encore ? ;
- le **moteur d'exécution** (quelle version de Crystal, **redistribuable**, **licence**, **bitness**) ;
- le **code d'intégration** VB6 — le **contrôle Crystal (OCX)** (`ReportFileName`, `Action`,
  `PrintReport`) ou le **RDC** (modèle objet `Application`/`Report`).

> ⚠️ Le Crystal **fourni avec VB6** est **obsolète et non compatible .NET**. Référencer l'ancien OCX
> n'est pas une option viable : il faut **un autre produit** (une version récente), avec sa propre
> compatibilité de `.rpt`, son runtime et sa licence.

---

## 2. La voie de la **compatibilité** : Crystal Reports sur .NET

Crystal Reports appartient désormais à **SAP**. Il existe **Crystal Reports for Visual Studio
(CR for VS)** — un **téléchargement séparé** (Crystal n'est plus livré avec Visual Studio ni VB). Il
fournit des assemblys .NET (`CrystalDecisions.CrystalReports.Engine`, `CrystalDecisions.Shared`,
`CrystalDecisions.Windows.Forms`) et le contrôle **`CrystalReportViewer`** pour Windows Forms.

Cette voie permet de **réutiliser les `.rpt` existants**, **mais** :

- les `.rpt` de l'**époque VB6** sont anciens : ils doivent généralement être **ouverts et
  re-sauvegardés** dans un designer récent ; certaines **formules/fonctionnalités** peuvent se
  comporter différemment → **tester**.
- il faut **installer le runtime** CR for VS sur **chaque poste**, à la **bonne version et bitness**
  (déploiement, **module 18**).
- ⚠️ **Licence** : CR for VS est gratuit pour le développement et le déploiement du runtime, mais les
  conditions **SAP ont évolué** au fil du temps et certaines fonctionnalités avancées sont payantes —
  **vérifiez les termes actuels**.

### Modèle « push » plutôt que « pull »

| Modèle | Le rapport… | À privilégier ? |
|--------|-------------|-----------------|
| **Pull** | se connecte **lui-même** à la base (connexion stockée dans le `.rpt`) | ❌ fragile, identifiants dans le `.rpt` |
| **Push** | **reçoit** ses données (un `DataSet`/`DataTable` fourni par le code) | ✅ **recommandé** : découple le rapport de la base |

Le modèle **push** s'appuie directement sur les `DataSet`/`DataTable` du chapitre (section 15.3).

---

## 3. Les **alternatives**

Vu les frictions de la voie Crystal (runtime, bitness, licence, dépendance SAP, re-travail des `.rpt`),  
beaucoup évaluent un changement de moteur :

| Option | Nature | Réutilise les `.rpt` ? | Coût / dépendance |
|--------|--------|------------------------|-------------------|
| **Crystal Reports for VS** | Moteur SAP + viewer WinForms | ✅ (avec re-sauvegarde) | Runtime à déployer ; licence SAP à vérifier |
| **RDLC + `ReportViewer`** | Moteur **Microsoft local**, viewer WinForms | ❌ (à **recréer**) | Gratuit ; designer = extension VS, viewer via **NuGet** |
| **SSRS** (`.rdl`) | Reporting **serveur** | ❌ | Infrastructure SQL Server |
| **Suites commerciales** | DevExpress, Telerik, Stimulsoft, **ActiveReports**… | ❌ | **Licence** ; designers + support |
| **Génération directe** | `PrintDocument` (14.4) / PDF / Office | ❌ | **Peu** de dépendances ; **manuel** |

- **RDLC + `ReportViewer`** est l'alternative **native** la plus courante : rapports `.rdlc` rendus
  **localement** (sans serveur), données fournies via `ReportDataSource`. ⚠️ **pas d'import des
  `.rpt`** → les rapports sont à **recréer**.
- **SSRS** convient si l'on dispose déjà de SQL Server et que l'on veut un reporting **centralisé**.
- Les **suites commerciales** offrent designers et exports, contre une licence.
- La **génération directe** (via `PrintDocument`, cf. 14.4, ou une bibliothèque PDF) convient aux
  rapports **simples**.

---

## 4. ⚠️ Grille de décision

| Facteur | Penche vers **garder Crystal** | Penche vers une **alternative** |
|---------|--------------------------------|----------------------------------|
| **Volume/complexité des `.rpt`** | Nombreux et complexes (réutilisation rentable) | Peu nombreux ou simples (recréation acceptable) |
| **Dépendance SAP / licence** | Tolérée | À **éviter** → natif (RDLC) ou commercial |
| **Infrastructure** | — | SQL Server présent → SSRS |
| **Compétences/budget** | — | Budget pour une suite + support |
| **Bitness/déploiement** | Runtime acceptable | Éviter un runtime lourd → **RDLC** natif |
| **Saut .NET moderne** (module 20) | À vérifier | Évaluer la maturité .NET du moteur |

**Position pragmatique :**

1. **Gros investissement en `.rpt`** + risque court terme minimal → **garder Crystal** (CR for VS,
   réutiliser les `.rpt`, modèle **push**), en assumant runtime/licence/bitness.
2. Rapports **peu nombreux/simples** ou volonté de **lâcher la dépendance** → **recréer en RDLC +
   `ReportViewer`** (natif) ou via une suite **commerciale**.
3. **Ne pas** tout réécrire en `PrintDocument` sauf rapports **triviaux**.
4. **Encapsuler** l'appel au reporting derrière une **abstraction** (comme en 14.6) pour pouvoir
   **changer de moteur** plus tard.

---

## 5. Étapes pratiques (voie compatibilité)

1. **Inventorier** les `.rpt` : sources de données, paramètres, sous-rapports, formules (modules
   3.1/3.2).
2. **Installer** CR for VS (version/bitness adaptées) ; **ouvrir et re-sauvegarder** les `.rpt` dans le
   designer récent.
3. **Remplacer** le code OCX/RDC par le modèle objet .NET (`ReportDocument` + `CrystalReportViewer`),
   basculer en **modèle push** (`SetDataSource`), passer les **paramètres** (`SetParameterValue`).
4. **Tester** rendu, **export** (PDF/Excel) et **impression**.
5. Régler le **runtime redistribuable** pour le **déploiement** (module 18).

---

## 6. Aperçu du code (illustratif)

```vb
' Crystal Reports for VS — modèle "push" + visionneuse
Dim rapport As New ReportDocument()
rapport.Load("Factures.rpt")
rapport.SetDataSource(table)                 ' on FOURNIT le DataTable/DataSet
rapport.SetParameterValue("Annee", 2026)
crvViewer.ReportSource = rapport             ' CrystalReportViewer
' ou rapport.ExportToDisk(...) / rapport.PrintToPrinter(...)
```

```vb
' Alternative native : RDLC + ReportViewer (rapport LOCAL, à recréer)
rvViewer.LocalReport.ReportPath = "Factures.rdlc"
rvViewer.LocalReport.DataSources.Clear()
rvViewer.LocalReport.DataSources.Add(New ReportDataSource("DonneesFactures", table))
rvViewer.LocalReport.SetParameters(New ReportParameter("Annee", "2026"))
rvViewer.RefreshReport()
```

> ℹ️ Chaque voie requiert ses **assemblys** : ceux de SAP (CR for VS) d'un côté, le `ReportViewer`
> Microsoft (via **NuGet**) et l'extension de **designer RDLC** de l'autre.

---

## 7. ⚠️ Pièges silencieux à retenir

- **Crystal de VB6 = obsolète et non .NET** : impossible de réutiliser l'OCX d'origine ; il faut
  **CR for VS** (autre produit/version).
- **Compatibilité des `.rpt`** : la re-sauvegarde dans un designer récent peut **altérer** le
  comportement → **tester** systématiquement.
- **Runtime + bitness** à installer **par poste** (déploiement, module 18).
- **Licence SAP** : conditions **évolutives** ; certaines fonctionnalités **payantes** → vérifier.
- **Modèle pull** = identifiants dans le `.rpt` + connexions fragiles → préférer le **push** (`DataSet`).
- **Recréer** des rapports (RDLC/commercial) est un **vrai coût** : **pas** d'import des `.rpt`.
- **`ReportViewer`** = **NuGet** + designer en **extension VS** (plus intégré par défaut) → friction de
  mise en place.
- **Sous-rapports/formules complexes** se transposent mal d'un moteur à l'autre.
- **Sous-estimer l'inventaire** des rapports : il est souvent **plus volumineux** que prévu.

> 🔗 Voir la **section 15.3** (`DataSet` pour le modèle push), la **section 14.6** (encapsulation,
> contrôles tiers), la **section 14.4** (`PrintDocument` pour la génération directe), le **module 18**
> (déploiement/redistribuables) et le **module 3** (inventaire/dépendances).

---

## ✅ Fin du module 15

Le socle de données a changé de **modèle**, pas seulement de noms :

- **trois piles COM** (DAO/RDO/ADO) convergent vers **un** modèle managé, **ADO.NET** (15.1) ;
- les **connexions** s'ouvrent tard/se ferment tôt et les requêtes se **paramètrent** — sécurité
  comprise (15.2) ;
- le `Recordset` se scinde en **`DataReader`** (connecté) et **`DataSet`/`DataTable`** (déconnecté),
  avec un changement de **concurrence** (15.3) ;
- la **liaison de données** passe par le **`BindingSource`** (15.4) ;
- chaque base a ses **chaînes de connexion** et ses **adaptations** (bitness, dialecte) (15.5) ;
- les **transactions** deviennent **explicites** (`DbTransaction`) ou **ambiantes** (`TransactionScope`)
  (15.6) ;
- et le **reporting** est une **décision** d'architecture et de licence (15.7).

Fil rouge : **connecté → déconnecté**, **paramétrer**, et **décider** (forme de données, provider,  
moteur de reporting) plutôt que convertir mécaniquement.

---

**Section suivante → Module 16 — Interopérabilité COM et appels d'API Windows** 🔗 ⭐, dernier volet de
la Partie 5, pour faire cohabiter et appeler l'existant COM/Win32 depuis .NET.

⏭️ [Interopérabilité COM et appels d'API Windows](/16-interop-com-api/README.md)
