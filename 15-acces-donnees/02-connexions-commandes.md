🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.2 Connexions et **commandes paramétrées** (et la sécurité au passage)

> Une règle gouverne cette section : **on ne concatène jamais de SQL avec des données**. Les
> **commandes paramétrées** ne sont pas une option de confort — elles ferment la porte à l'**injection
> SQL** *et* corrigent au passage des bugs de données (apostrophes, séparateurs décimaux, formats de
> date). Côté connexions, on abandonne la connexion **globale ouverte** de VB6 au profit du couple
> **`Using` + *pooling***.

**Module 15 — Accès aux données** · Cible : .NET Framework 4.7.2 · `System.Data.SqlClient`,
`System.Data.OleDb`

---

## 1. Connexions : **ouvrir tard, fermer tôt**

En VB6, on ouvrait souvent **une** connexion ADO au démarrage et on la gardait ouverte. En ADO.NET,  
c'est l'inverse : les connexions sont **mutualisées** (*pooling*), et le bon réflexe est d'**ouvrir au  
plus tard, fermer au plus tôt**, avec `Using` :

```vb
Using cnx As New SqlConnection(chaine)
    cnx.Open()
    ' ... opérations ...
End Using        ' Close() + Dispose() garantis, même en cas d'exception
```

- Le *pooling* est **actif par défaut** (SqlClient) et **indexé sur la chaîne de connexion** : rouvrir
  une connexion identique **réutilise** une connexion du pool — c'est **bon marché**.
- La chaîne de connexion (détaillée en **15.5**) ne doit **pas** contenir d'identifiants en clair dans
  le code : placez-la en **configuration** (`app.config`, section `connectionStrings`) :

```vb
' <connectionStrings><add name="Db" connectionString="..." /></connectionStrings>
Dim chaine = ConfigurationManager.ConnectionStrings("Db").ConnectionString
```

> ⚠️ **Ne pas libérer** une connexion (oubli de `Using`/`Dispose`) **épuise le pool** → erreur
> « *Timeout expired… pool* » sous charge. C'est l'application directe d'`IDisposable` (**module 12**),
> et l'un des bugs les plus fréquents des applications migrées.

---

## 2. Commandes et exécution

L'objet **`Command`** porte le SQL, son **type** et ses **paramètres** :

| VB6 (ADO) | ADO.NET | Renvoie |
|-----------|---------|---------|
| `cn.Execute sql` (action) | `cmd.ExecuteNonQuery()` | **Nombre de lignes** affectées (`Integer`) |
| `rs.Open sql, cn` (lignes) | `cmd.ExecuteReader()` | Un **`DataReader`** |
| valeur unique (`rs.Fields(0)`) | `cmd.ExecuteScalar()` | **Une seule** valeur (`Object`) |

```vb
' Valeur unique : COUNT, agrégat, identité…
cmd.CommandText = "SELECT COUNT(*) FROM Clients WHERE Actif = 1"
Dim nb As Integer = CInt(cmd.ExecuteScalar())
```

Le **type** de commande se précise via `CommandType` : `Text` (défaut), `StoredProcedure`,
`TableDirect`. Un `CommandTimeout` règle le délai d'exécution.

---

## 3. ⚠️ Le cœur : les commandes **paramétrées** (sécurité)

### Le problème : la concaténation

```vb
' ❌ DANGEREUX : injection SQL + bugs (apostrophes, séparateurs, dates)
Dim sql = "SELECT * FROM Clients WHERE Nom = '" & txtNom.Text & "'"
```

Si `txtNom` vaut `' OR '1'='1`, la requête change de sens. Et un simple nom comme `O'Connor` la
**casse**. La concaténation mélange **code** et **données** — c'est la faille.

### La solution : des **paramètres**

On insère des **marqueurs** dans le SQL et on fournit les valeurs **séparément**. Le provider  
transmet alors les données **hors** du texte SQL → injection **impossible**, apostrophes gérées, types  
corrects :

```vb
Using cnx As New SqlConnection(chaine),
      cmd As New SqlCommand("SELECT Id, Nom FROM Clients WHERE Nom = @nom", cnx)
    cmd.Parameters.Add("@nom", SqlDbType.NVarChar, 100).Value = txtNom.Text
    cnx.Open()
    Using lecteur = cmd.ExecuteReader()
        While lecteur.Read()
            ' ...
        End While
    End Using
End Using
```

### ⚠️ La syntaxe dépend du provider (nommé `@` vs positionnel `?`)

C'est un **piège classique** lors d'une migration multi-bases :

```vb
' SQL Server (SqlClient) : paramètres NOMMÉS, préfixe @
"... WHERE Nom = @nom AND Ville = @ville"

' OLE DB / ODBC : paramètres POSITIONNELS, marqueur ?  (l'ORDRE COMPTE !)
"... WHERE Nom = ? AND Ville = ?"
' -> ajouter les paramètres DANS L'ORDRE des ? (les noms sont ignorés pour la liaison) :
cmd.Parameters.AddWithValue("p1", nom)     ' 1er ?
cmd.Parameters.AddWithValue("p2", ville)   ' 2e ?
```

### Ajouter des paramètres : `AddWithValue` vs `Add` typé

```vb
' ⚠️ AddWithValue INFÈRE le type et la taille depuis la valeur
cmd.Parameters.AddWithValue("@nom", txtNom.Text)

' ✅ Préférer un type ET une taille EXPLICITES (types exacts, plans de requête stables)
cmd.Parameters.Add("@nom", SqlDbType.NVarChar, 100).Value = txtNom.Text

' NULL : DBNull.Value, JAMAIS Nothing
Dim p = cmd.Parameters.Add("@note", SqlDbType.NVarChar, 200)
p.Value = If(note IsNot Nothing, CObj(note), CObj(DBNull.Value))
```

> ℹ️ `AddWithValue` est pratique mais peut déduire un **type/taille inadaptés** (par ex. `nvarchar`
> face à une colonne `varchar`, ou une taille variable qui **fragmente le cache de plans**). Pour du
> code robuste, **typez** vos paramètres.

### Procédures stockées et paramètres de sortie

```vb
Using cmd As New SqlCommand("dbo.ObtenirSolde", cnx)
    cmd.CommandType = CommandType.StoredProcedure
    cmd.Parameters.Add("@compteId", SqlDbType.Int).Value = compteId
    Dim sortie = cmd.Parameters.Add("@solde", SqlDbType.Decimal)
    sortie.Direction = ParameterDirection.Output      ' (ou InputOutput / ReturnValue)
    cnx.Open()
    cmd.ExecuteNonQuery()
    Dim solde As Decimal = CDec(sortie.Value)
End Using
```

C'est l'équivalent typé des `adParamOutput`/`Parameters.Refresh` d'ADO.

### ⚠️ Le piège de la clause `IN`

```vb
' ❌ Ne fonctionne pas : un paramètre ne se "déplie" pas en liste
"... WHERE Id IN (@ids)"
```

Il faut **générer un paramètre par valeur** (`@id0, @id1, …`) ou utiliser un **paramètre table**
(TVP, SQL Server). À repérer dans le code VB6 qui construisait des listes `IN` par concaténation.

---

## 4. Le motif idiomatique complet

```vb
Dim chaine = ConfigurationManager.ConnectionStrings("Db").ConnectionString

' Écriture (action)
Using cnx As New SqlConnection(chaine),
      cmd As New SqlCommand(
          "UPDATE Clients SET Ville = @ville WHERE Id = @id", cnx)
    cmd.Parameters.Add("@ville", SqlDbType.NVarChar, 60).Value = ville
    cmd.Parameters.Add("@id", SqlDbType.Int).Value = idClient
    cnx.Open()
    Dim lignes As Integer = cmd.ExecuteNonQuery()    ' nb de lignes modifiées
End Using
```

Le triptyque est toujours le même : **`Using` connexion + commande**, **paramètres typés**, puis  
l'**Execute** adapté (`Reader`/`NonQuery`/`Scalar`).

---

## 5. Lire les résultats (renvoi 15.3)

Pour mémoire, `ExecuteReader` renvoie un **`DataReader`** que l'on parcourt avec `Read()` ; on accède  
aux colonnes par index/nom, avec des **accesseurs typés** et un test de nullité :

```vb
While lecteur.Read()
    Dim id As Integer = lecteur.GetInt32(0)
    Dim nom As String = If(lecteur.IsDBNull(1), "", lecteur.GetString(1))
End While
```

> 🔗 Le détail (lecteur **vs** `DataSet`/`DataTable`, curseurs, modifications, renvoi en base) est
> l'objet de la **section 15.3**.

---

## 6. Erreurs et asynchrone

- **Erreurs** : la collection `Errors` d'ADO et `Err` laissent place aux **exceptions** — typiquement
  `SqlException` (avec `.Number` et `.Errors`). On enveloppe les opérations dans `Try`/`Catch`
  (détails au **module 11**).
- **Asynchrone** (optionnel) : ADO.NET propose `OpenAsync`/`ExecuteReaderAsync` (depuis .NET 4.5,
  disponibles en 4.7.2). Pour une première migration d'un code VB6 **synchrone**, **restez
  synchrone** ; l'asynchrone est un raffinement **ultérieur**.

---

## 7. La sécurité, au-delà de l'injection

Profitez de la migration pour durcir l'accès aux données :

- **Chaînes de connexion en configuration** (pas en dur), et **chiffrées** si elles contiennent des
  identifiants (DPAPI/section protégée) ; privilégier l'**authentification intégrée** quand c'est
  possible.
- **Moindre privilège** : un compte applicatif aux droits **strictement nécessaires**.
- **Bénéfice de correction** : la paramétrisation règle aussi des **bugs de données** — apostrophes
  dans les noms, **séparateurs décimaux** et **formats de date** selon les paramètres régionaux. Ce
  dernier point rejoint le piège **culture** de l'**Annexe B.10**.

---

## 8. ⚠️ Pièges silencieux à retenir

- **SQL concaténé** → injection **et** bugs (apostrophes, culture). **Toujours paramétrer.**
- **Syntaxe des marqueurs** : `@nom` **nommé** (SqlClient) vs `?` **positionnel** (OLE DB/ODBC, où
  **l'ordre des paramètres compte**).
- **`AddWithValue`** infère type/taille → types inexacts, cache de plans fragmenté. Préférer
  **`Add(nom, type, taille)`**.
- **NULL** : passer **`DBNull.Value`**, jamais `Nothing`.
- **Connexion globale ouverte** → épuisement du pool, mauvaise montée en charge. **`Using` + ouvrir
  tard/fermer tôt.**
- **Oubli de `Dispose`** (connexion/commande/lecteur) → fuites.
- **Clause `IN (liste)`** non paramétrable directement → un paramètre par valeur, ou TVP.
- **Identifiants en dur** dans la chaîne de connexion.

> 🔗 Voir le **module 11** (exceptions), le **module 12** (`IDisposable`/`Using`), la **section 15.3**
> (lecteur vs déconnecté), la **section 15.5** (chaînes de connexion/providers) et l'**Annexe B.10**
> (culture).

---

**Section suivante → 15.3 — `Recordset` → `DataReader` / `DataSet` / `DataTable` (curseurs vs
déconnecté)** ⚠️, le cœur de la migration des données.

⏭️ [`Recordset` → `DataReader` / `DataSet` / `DataTable` (curseurs vs déconnecté)](/15-acces-donnees/03-recordset-vers-adonet.md)
