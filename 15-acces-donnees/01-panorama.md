🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.1 Panorama : DAO, RDO, ADO (VB6) vs **ADO.NET**

> Trois piles d'accès aux données ont coexisté en VB6 — **DAO**, **RDO**, **ADO** — empilées au fil du
> temps. Toutes convergent vers **une seule** technologie managée : **ADO.NET**. Mais attention au
> faux-ami : **ADO.NET n'est pas « ADO version 2 »**. C'est une **architecture différente**, sans
> COM, organisée autour de **providers** et d'un modèle **déconnecté**.

**Module 15 — Accès aux données** · Cible : .NET Framework 4.7.2 · `System.Data`, `System.Data.Common`

---

## 1. Trois piles, une histoire en couches

| Pile VB6 | Apparue pour | Repose sur | Modèle d'objets (résumé) |
|----------|--------------|-----------|--------------------------|
| **DAO** | Bases **Access/Jet** (`.mdb`) locales | Le moteur **Jet** | `DBEngine` → `Workspace` → `Database` → `Recordset`, `TableDef`, `QueryDef`, `Relation` |
| **RDO** | Bases **serveur** (SQL Server, Oracle) | **ODBC** | `rdoEngine` → `rdoEnvironment` → `rdoConnection` → `rdoResultset`, `rdoQuery` |
| **ADO** | La pile « stratégique » VB6 | **OLE DB** (peut envelopper ODBC) | `Connection`, `Recordset`, `Command`, `Field`, `Parameter`, `Errors` |

- **DAO** brille sur Jet (relations, DDL, fonctions Access) mais s'adapte mal au client/serveur.
- **RDO** est une fine couche objet sur ODBC, pensée serveur ; largement **supplantée par ADO**.
- **ADO** est le modèle **plat et simple** dominant dans les applications VB6 **tardives**, via la
  vision « accès universel aux données » (OLE DB).

> ℹ️ En pratique, un *legacy* mélange souvent **plusieurs** piles (de l'ADO récent côtoyant du DAO ou
> du RDO résiduel), avec la couche d'accès **tissée dans les formulaires** (séparation faible — un
> point à corriger, cf. module 5.3).

---

## 2. ADO.NET : une autre architecture (pas « ADO version 2 »)

ADO.NET est **managé** (aucun COM) et repose sur un **modèle de providers** : chaque fournisseur  
implémente un jeu de classes dérivant de **classes de base communes** (`System.Data.Common`).

Il s'articule en **deux piliers** :

- **Couche connectée** : `Connection`, `Command`, `DataReader`, `Transaction`, `Parameter`.
  Lecture **avant seulement**, rapide, connexion **maintenue** le temps de l'opération.
- **Couche déconnectée** : `DataSet`, `DataTable`, `DataRelation`, `DataView`, synchronisés par un
  **`DataAdapter`**. On travaille **en mémoire**, **connexion fermée**.

### Les providers fournis avec .NET Framework 4.7.2

| Provider | Espace de noms | Cible |
|----------|----------------|-------|
| **SqlClient** | `System.Data.SqlClient` | **SQL Server** (natif, rapide) — **à privilégier** |
| **OLE DB** | `System.Data.OleDb` | Providers OLE DB (**Access/ACE**, autres) |
| **ODBC** | `System.Data.Odbc` | Pilotes **ODBC** (portée de RDO) |
| **OracleClient** | `System.Data.OracleClient` | Oracle — **déprécié** (préférer ODP.NET) |

> ℹ️ Pour d'autres bases (Oracle, MySQL, PostgreSQL, SQLite), on ajoute un **provider de l'éditeur**
> via NuGet (ODP.NET, Connector/NET, **Npgsql**, Microsoft.Data.Sqlite). La classe
> **`DbProviderFactory`** permet d'écrire du code **agnostique** au provider (§7).

---

## 3. Le contraste de fond

| Aspect | VB6 (DAO / RDO / ADO) | ADO.NET |
|--------|------------------------|---------|
| **Technologie** | COM / OLE DB / ODBC | **Managé** (.NET), sans COM |
| **Modèle dominant** | `Recordset` **connecté**, curseurs | **Déconnecté** (`DataSet`) **+** lecteur rapide (`DataReader`) |
| **Connexion** | souvent **gardée ouverte** | **mutualisée** (*pooling*), ouvrir tard / fermer tôt |
| **Navigation** | bidirectionnelle, `RecordCount`, édition en place | `DataReader` avant-seulement **ou** `DataTable` en mémoire |
| **Paramètres** | `Command`/`Parameters` (ADO) | `Command`/`Parameters` **typés** |
| **Erreurs** | `Err` + collection `Errors` | **Exceptions** (`SqlException`…) |
| **Valeur nulle** | `Variant` `Null`, `IsNull` | **`DBNull.Value`** |
| **Provider** | un par technologie | **modèle de providers** unifié |

C'est cette ligne **« connecté/curseur → déconnecté »** qui structure tout le chapitre (et que  
détaille la **section 15.3**).

---

## 4. Le modèle d'objets, pile par pile

### ADO → ADO.NET (le cas le plus fréquent)

| ADO | ADO.NET |
|-----|---------|
| `Connection` | `DbConnection` (`SqlConnection` / `OleDbConnection`) |
| `Command` | `DbCommand` (`SqlCommand`…) |
| `Recordset` *forward-only* | `DataReader` |
| `Recordset` client / déconnecté | `DataTable` / `DataSet` + `DataAdapter` |
| `Field` | colonne d'un `DataReader` / `DataColumn` |
| `Parameter` | `DbParameter` |
| `Errors` | `Exception` (`SqlException.Errors`) |
| `BeginTrans` / `CommitTrans` | `DbTransaction` |

### DAO → ADO.NET

| DAO | ADO.NET |
|-----|---------|
| `DBEngine` / `Workspace` | (plus d'équivalent : provider + connexion) |
| `Database` | `DbConnection` (OLE DB vers **ACE/Jet**) |
| `Recordset` (`Table`/`Dynaset`/`Snapshot`) | `DataReader` **ou** `DataTable` selon l'usage |
| `TableDef` / `QueryDef` | requêtes/commandes SQL ; schéma via `GetSchema` |
| `Relation` | `DataRelation` (en mémoire) ; contraintes **en base** |

> ⚠️ DAO expose des **fonctions propres à Jet** (DDL Jet, dialecte SQL Access, propriétés Access) qui
> **ne se transposent pas** directement. À auditer lors de la migration.

### RDO → ADO.NET

| RDO | ADO.NET |
|-----|---------|
| `rdoEngine` / `rdoEnvironment` | (provider + connexion) |
| `rdoConnection` | `DbConnection` (`Odbc` / `SqlClient`) |
| `rdoResultset` | `DataReader` **ou** `DataTable` |
| `rdoQuery` | `DbCommand` (paramétrée) |

---

## 5. Connecté vs déconnecté : l'idée en deux images

Le **même** besoin de lecture se traduit de **deux** façons en ADO.NET — c'est le choix central de la  
section 15.3 :

```vb
' (a) CONNECTÉ : lecture rapide, avant-seulement
Using lecteur = cmd.ExecuteReader()
    While lecteur.Read()
        ' ... lecteur(0), lecteur("Nom") ...
    End While
End Using                      ' la connexion reste ouverte le temps de la lecture

' (b) DÉCONNECTÉ : on remplit en mémoire, puis la connexion se ferme
Dim table As New DataTable()
Using adaptateur As New SqlDataAdapter(cmd)
    adaptateur.Fill(table)     ' ouvre, remplit, PUIS FERME la connexion
End Using
' 'table' est exploitable hors connexion (accès aléatoire, tri, modifications)
```

> 🔗 Le détail (curseurs, `RecordCount`, édition en place, `RowState`, renvoi des modifications) est
> traité en **15.3**.

---

## 6. Quel provider pour quelle situation (orientation ; détails en 15.5)

| Base de données | Provider recommandé |
|-----------------|---------------------|
| **SQL Server** | `System.Data.SqlClient` (**natif**) |
| **Access** (`.mdb`/`.accdb`) | `System.Data.OleDb` avec le provider **ACE** (ou **Jet** pour les anciens `.mdb`) |
| **Oracle** | **ODP.NET** (le `OracleClient` intégré est **déprécié**) |
| **MySQL / PostgreSQL / SQLite** | Provider de l'éditeur via NuGet (Connector/NET, **Npgsql**, Microsoft.Data.Sqlite) |
| Autre, via pilote ODBC | `System.Data.Odbc` |

> ℹ️ Règle simple : **préférez le provider natif** d'une base à un pont OLE DB/ODBC quand il existe
> (performances et fidélité). ⚠️ Le provider **ACE/Jet** d'Access doit correspondre à la **bitness**
> du processus (32/64 bits) — un piège de déploiement traité en **15.5** et au **module 18**.

---

## 7. Code agnostique au provider (`DbProviderFactory`) — une option

Si l'application doit viser **plusieurs back-ends**, les classes de base de `System.Data.Common`
(`DbConnection`, `DbCommand`, `DbParameter`…) et la fabrique **`DbProviderFactory`** permettent
d'écrire du code **neutre vis-à-vis du provider**. C'est l'équivalent .NET de la souplesse qu'offrait  
déjà ADO via OLE DB. Optionnel : pour une cible unique, le provider natif direct suffit.

---

## 8. Ce que cela implique pour la migration

- **Uniformiser** : un *legacy* mêlant DAO/RDO/ADO peut être **standardisé** sur **un seul** modèle
  (ADO.NET) — une vraie simplification.
- **Décider au cas par cas** entre `DataReader` (connecté) et `DataSet`/`DataTable` (déconnecté) pour
  chaque opération (→ **15.3**).
- **Paramétrer** systématiquement les requêtes (sécurité, → **15.2**) et adopter le réflexe
  **`Using`/pooling** (cf. README du module et **module 12**).
- **Pont temporaire** : on peut **conserver ADO via interop COM** le temps de la transition
  (cohabitation, **module 16**), la cible restant **ADO.NET managé**.
- **Découpler** la couche de données de l'UI au passage (**module 5.3**) ; envisager une **fine
  couche d'accès** dédiée.

> ℹ️ Des approches de **plus haut niveau** (ORM type Entity Framework) existent, mais elles
> **sortent du périmètre** de cette migration « pont » : ADO.NET est la cible **directe et à faible
> risque** ici (on minimise les changements simultanés — philosophie du module 1). Un ORM relève d'une
> étape **ultérieure et optionnelle**.

---

## 9. ⚠️ Points de vigilance du panorama

- **Ne pas traiter ADO.NET comme « ADO renommé »** : le modèle **déconnecté** et l'absence des
  comportements de **curseur vivant** se rappelleront à vous (→ 15.3).
- **Fonctionnalités propres à Jet** (DAO) : DDL Jet, dialecte SQL Access, propriétés Access → pas de
  transposition directe.
- **Préférer le provider natif** aux ponts OLE DB/ODBC quand il existe.
- **`OracleClient` déprécié** → ODP.NET.
- **Bitness du provider ACE/Jet** (32/64) → piège de déploiement (15.5 / module 18).
- **Mentalité de connexion** : abandonner la connexion globale ouverte au profit du *pooling*.

> 🔗 Voir le **README du module** (réflexes `Using` et `DBNull`), la **section 15.3** (curseurs vs
> déconnecté), la **section 15.5** (providers et chaînes de connexion), le **module 16** (interop) et
> l'**Annexe G** (glossaire DAO/RDO/ADO/ADO.NET).

---

**Section suivante → 15.2 — Connexions et commandes paramétrées (et la sécurité au passage)**, pour
ouvrir/fermer proprement et remplacer la concaténation de SQL par des **paramètres**.

⏭️ [Connexions et **commandes paramétrées** (et la sécurité au passage)](/15-acces-donnees/02-connexions-commandes.md)
