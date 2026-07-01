🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 🗄️ 15. Accès aux données — d'ADO/DAO/RDO à ADO.NET ⭐

> **Le changement ne porte pas que sur des noms d'objets : il porte sur le *modèle*.** Les
> applications VB6 travaillent presque toujours avec un `Recordset` **connecté** et **vivant** ;
> ADO.NET privilégie un modèle **déconnecté** (on remplit en mémoire, on ferme la connexion, on
> renvoie les modifications). Comprendre cette bascule **connecté → déconnecté** est la clé de tout
> le chapitre.

**Partie 5 — Migrer les données et l'interopérabilité** · Cible : .NET Framework 4.7.2 ·
`System.Data` (ADO.NET)

---

## 🧭 Où se situe ce chapitre

Après l'interface utilisateur (modules 13–14), on attaque le **socle de données**. Ce module 15  
migre l'**accès aux données** ; le module 16 enchaînera sur l'**interopérabilité COM et l'API  
Windows**. Les deux forment la **Partie 5**.

La plupart des applications VB6 reposent sur **l'une des trois piles COM** d'accès aux données. Toutes  
convergent vers **un seul** modèle managé : **ADO.NET**.

| Pile VB6 (COM) | Pour quoi | Convergence .NET |
|----------------|-----------|------------------|
| **DAO** (Data Access Objects) | Bases **Access/Jet** (`.mdb`) | **ADO.NET** (provider OLE DB / ACE) |
| **RDO** (Remote Data Objects) | Bases **serveur** via ODBC | **ADO.NET** (`System.Data.Odbc` / `SqlClient`) |
| **ADO** (ActiveX Data Objects) | La pile « moderne » VB6, via OLE DB | **ADO.NET** (`SqlClient` / `OleDb`) |

---

## ⚠️ Le piège central : du `Recordset` **connecté** au modèle **déconnecté**

C'est **le** point à intégrer avant le reste.

En VB6, un `Recordset` est en général **ouvert sur la connexion**, avec un **curseur** (statique,  
dynamique, *keyset*, *forward-only*), un **type de verrou**, une navigation **vivante**
(`MoveNext`/`MovePrevious`/`MoveFirst`/`RecordCount`/`BOF`/`EOF`) et des **modifications en place**
(`AddNew`/`Update`/`Delete`).

ADO.NET propose **deux formes distinctes**, et il faut **choisir laquelle** correspond à chaque usage :

- **Connecté, rapide, en lecture seule** : un **`DataReader`** — flux **avant seulement**, connexion
  ouverte, idéal pour lire vite et peu de mémoire. **Pas** de `MovePrevious`, **pas** de
  `RecordCount`, **pas** d'édition en place.
- **Déconnecté** : un **`DataTable`/`DataSet`** rempli par un **`DataAdapter`**, puis **la connexion
  se ferme** ; on travaille **en mémoire** (accès aléatoire, modifications, `RowState`) et l'on
  **renvoie** les changements via l'adaptateur.

> ⚠️ La correspondance **n'est pas 1:1**. Un même `Recordset` VB6 peut devenir un `DataReader` **ou**
> un `DataTable` selon ce qu'on en fait. Ce choix — et les comportements de curseur qui **ne se
> transposent pas** — est l'objet de la **section 15.3**.

---

## 🔒 Le second fil rouge : en profiter pour **sécuriser**

Le code VB6 construit souvent ses requêtes par **concaténation de chaînes**
(`"… WHERE nom='" & txt & "'"`), ouvrant la porte à l'**injection SQL**. La migration est le **bon
moment** pour adopter les **commandes paramétrées** — non comme une option, mais comme la **règle**.  
C'est l'objet de la **section 15.2** (« et la sécurité au passage »).

---

## 📋 Vue d'ensemble des correspondances

| VB6 (ADO) | ADO.NET | Remarque |
|-----------|---------|----------|
| `ADODB.Connection` | `SqlConnection` / `OleDbConnection` (`DbConnection`) | Provider à **choisir** |
| `ADODB.Command` + `Parameters` | `SqlCommand` + `Parameters` | **Paramétrer** (sécurité) |
| `Recordset` *forward-only* | `DataReader` | Connecté, lecture seule |
| `Recordset` client / déconnecté | `DataTable` / `DataSet` + `DataAdapter` | En mémoire |
| `Field.Value` (`Variant`, `Null`) | `reader(i)` / `Cells` (`Object`, **`DBNull.Value`**) | `DBNull` ≠ `Nothing` ⚠️ |
| `BeginTrans` / `CommitTrans` | `DbTransaction` / `TransactionScope` | Section 15.6 |
| `Data` control / *data binding* | `BindingSource` + contrôles liés | Section 15.4 |
| DAO/Jet (`.mdb`) | Provider OLE DB **ACE/Jet** | Section 15.5 |
| RDO / ODBC | `System.Data.Odbc` / `SqlClient` | Section 15.5 |

---

## 📑 Ce que vous allez voir

- **15.1 — Panorama : DAO, RDO, ADO → ADO.NET**
  Situer les trois piles VB6 et le modèle managé, avec le **modèle de providers** d'ADO.NET.
- **15.2 — Connexions et commandes paramétrées (et la sécurité)**
  Ouvrir/fermer proprement, et remplacer la concaténation par des **paramètres** (anti-injection).
- **15.3 — `Recordset` → `DataReader` / `DataSet` / `DataTable`** ⚠️
  Le cœur du chapitre : **curseurs vs déconnecté**, et comment **choisir** la bonne forme.
- **15.4 — Liaison de données (`Data` control → `BindingSource`)**
  Du *data binding* VB6 à la liaison .NET, avec `DataGridView` et contrôles liés.
- **15.5 — Access (`.mdb`/`.accdb`) et SQL Server**
  **Chaînes de connexion** et adaptations selon la base et le provider.
- **15.6 — Transactions (`ADO` → `DbTransaction`)**
  De `BeginTrans`/`CommitTrans` aux transactions managées (et `TransactionScope`).
- **15.7 — Reporting : Crystal Reports et alternatives** ⚠️
  Compatibilité de Crystal Reports et options de remplacement.

---

## ♻️ Un réflexe à acquérir tout de suite : connexions et `Using`

Beaucoup d'applications VB6 gardent **une connexion globale ouverte** pendant toute la durée de vie de  
l'application. En ADO.NET, c'est l'**anti-pattern** : les connexions sont **mutualisées** (*pooling*)  
et la règle est **ouvrir au plus tard, fermer au plus tôt**, avec `Using` :

```vb
Using cnx As New SqlConnection(chaineConnexion),
      cmd As New SqlCommand("SELECT Nom FROM Clients WHERE Id = @id", cnx)
    cmd.Parameters.AddWithValue("@id", idClient)
    cnx.Open()                          ' ouvrir au plus tard
    Using lecteur = cmd.ExecuteReader()
        While lecteur.Read()
            ' ... lire lecteur(0), etc.
        End While
    End Using
End Using                                ' fermé/libéré au plus tôt -> le pool gère le reste
```

> 🔗 `Connection`, `Command`, `DataReader`, `DataAdapter` sont **`IDisposable`** : c'est l'application
> directe du module 12 (`IDisposable`/`Using`). Ne pas les libérer **épuise le pool de connexions** —
> un grand classique des applications migrées qui « se figent » sous charge.

---

## ⚠️ Un piège de données à connaître d'emblée : `DBNull` n'est pas `Nothing`

En VB6, un champ `Null` se teste via `IsNull(champ.Value)`. En ADO.NET, une valeur **SQL NULL** est  
représentée par **`DBNull.Value`**, qui **n'est ni `Nothing` ni une chaîne vide**. On la teste  
explicitement :

```vb
If Not lecteur.IsDBNull(0) Then nom = lecteur.GetString(0)
```

> 🔗 C'est le prolongement direct de la fin du `Variant` et de ses états (`Null`/`Empty`/`Missing`,
> modules 2.7 et 7.7). Ce piège figure aussi à l'**Annexe B**.

---

## 🎯 Objectifs du chapitre

À l'issue de ce module, vous saurez :

- ✅ **Situer** DAO/RDO/ADO et choisir le **bon provider** ADO.NET (`SqlClient`, `OleDb`, `Odbc`).
- ✅ Décider, pour chaque `Recordset`, entre **`DataReader`** (connecté) et **`DataSet`/`DataTable`**
  (déconnecté).
- ✅ Écrire des **commandes paramétrées** et fermer la porte à l'**injection SQL**.
- ✅ Gérer les **connexions** (`Using`, *pooling*), les valeurs **`DBNull`** et les **transactions**.
- ✅ Migrer la **liaison de données** (`BindingSource`) et aborder le **reporting**.

---

## 🔑 Prérequis et liens utiles

- **Modules 7 et 2.7** (types, `Variant`/`Null`) : pour `DBNull` et le typage des valeurs lues.
- **Module 11** (gestion d'erreurs) : la collection `Errors` d'ADO et `Err` → `Try`/`Catch` et
  `SqlException`.
- **Module 12** (`IDisposable`/`Using`) : connexions, commandes et lecteurs à **libérer**.
- **Module 3.2** (cartographie des dépendances) : repérer les bases, pilotes et composants de données.
- **Module 16** (interopérabilité COM) : possibilité de **conserver ADO via interop** pendant la
  transition (cohabitation).
- **Annexe B** (pièges silencieux, dont **B.9** dates et **B.10** culture pour le formatage des
  données) ; **Annexe G** (glossaire DAO/RDO/ADO/ADO.NET).

---

> **En résumé**, ce chapitre remplace **trois piles COM** par **un** modèle managé, et fait passer
> d'un accès **connecté et vivant** à un accès **soit rapide en lecture (`DataReader`), soit
> déconnecté (`DataSet`/`DataTable`)**. Gardez les deux fils rouges : **choisir la bonne forme** pour
> chaque `Recordset`, et **paramétrer** systématiquement les requêtes.

Commençons par le **panorama** : **15.1 — DAO, RDO, ADO (VB6) vs ADO.NET**.

⏭️ [Panorama : DAO, RDO, ADO (VB6) vs **ADO.NET**](/15-acces-donnees/01-panorama.md)
