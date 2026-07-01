🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.5 Bases Access (`.mdb`/`.accdb`) et SQL Server : chaînes de connexion et adaptations

> Deux migrations très différentes se cachent ici. **Garder la même base** et ne changer que le code
> d'accès : la dialecte SQL reste la même, l'essentiel est la **chaîne de connexion**, la **bitness**
> et la syntaxe des paramètres. **Déplacer la base vers SQL Server** : il faut **en plus** adapter le
> **dialecte SQL** (dates, booléens, fonctions, identité). Savoir dans quel scénario on est dicte la
> charge de travail.

**Module 15 — Accès aux données** · Cible : .NET Framework 4.7.2 · `System.Data.SqlClient`,
`System.Data.OleDb`

---

## 1. Deux scénarios à distinguer

| Scénario | Ce qui change | Charge |
|----------|---------------|--------|
| **A. Même base, nouveau code** (ex. `.mdb`/`.accdb` conservé) | Code d'accès (ADO → ADO.NET OLE DB), **bitness**, marqueur `?` | Modérée |
| **B. Migrer aussi vers SQL Server** | Provider (`OleDb` → `SqlClient`), chaîne de connexion **et** **dialecte SQL** | Plus lourde |

> ℹ️ Le scénario B est une **migration distincte** (de la base, pas seulement du code). Outil dédié :
> **SSMA for Access** (SQL Server Migration Assistant), qui migre **schéma + données** (§7).

---

## 2. Choisir le provider

| Base | Provider .NET | Connexion |
|------|---------------|-----------|
| **SQL Server** | `System.Data.SqlClient` (**natif**) | `SqlConnection` |
| **Access** (`.accdb`/`.mdb`) | `System.Data.OleDb` + provider **ACE**/**Jet** | `OleDbConnection` |

Il n'existe **pas** de provider managé natif pour Access : on passe par **OLE DB**.

---

## 3. SQL Server : chaînes de connexion

```text
' Authentification Windows (recommandée : pas d'identifiants en clair)
Server=.\SQLEXPRESS;Database=Ventes;Integrated Security=SSPI;

' Authentification SQL (identifiants en CONFIGURATION, idéalement chiffrés)
Server=SRV01;Database=Ventes;User ID=app;Password=********;

' LocalDB (développement)
Server=(localdb)\MSSQLLocalDB;Database=Ventes;Integrated Security=true;
```

Quelques clés utiles :

- **`MultipleActiveResultSets=True`** (MARS) si vous devez ouvrir **plusieurs `DataReader`** sur une
  **même** connexion (sinon, ouvrez une **seconde** connexion — le pool s'en charge).
- **`Connect Timeout`**, **`Application Name`**, **`Pooling`**, **`Min/Max Pool Size`**.

> ⚠️ **Chiffrement de la connexion (`Encrypt`/`TrustServerCertificate`) : un défaut qui a changé.** Avec
> **`System.Data.SqlClient`** (le provider de .NET Framework 4.7.2), `Encrypt` vaut **`False`** par
> défaut. Mais **`Microsoft.Data.SqlClient` 4.0+** (le provider moderne, cf. module 20) l'a **passé à
> `True`** : une base **sans certificat valide** fait alors **échouer** la connexion — une **rupture**
> classique lors du saut vers .NET récent. **Positionnez `Encrypt`/`TrustServerCertificate`
> explicitement** plutôt que de vous fier au défaut.

---

## 4. Access : chaînes de connexion **et** le piège de la **bitness** ⚠️

```text
' .accdb (et .mdb modernes) -> ACE
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Data\Base.accdb;

' ancien .mdb -> Jet
Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Data\Base.mdb;

' base protégée par mot de passe
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Data\Base.accdb;Jet OLEDB:Database Password=********;
```

> ⚠️ **LE piège de cette section — la bitness.**
> - **Jet 4.0 est 32 bits UNIQUEMENT** : un processus **64 bits** ne peut **pas** le charger. → cibler
>   **x86**, ou migrer le `.mdb` vers **ACE**.
> - **ACE** existe en **32 et 64 bits** (redistribuable « Microsoft Access Database Engine »), et la
>   **bitness installée doit correspondre à celle du processus**. On ne peut généralement **pas**
>   installer les deux facilement (conflit avec la bitness d'Office).
> - Conséquence **déploiement** : le **redistribuable ACE** doit être présent sur **chaque poste**, à
>   la **bonne bitness** (**module 18**).

> ℹ️ La **sécurité utilisateur** d'Access (fichier `.mdw`, ancien Jet) a **disparu** avec `.accdb` :
> seul un **mot de passe de base** subsiste. À vérifier si l'application VB6 en dépendait.

---

## 5. Mettre les chaînes en **configuration**

Stockez les chaînes (et le **`providerName`**) dans `app.config`, et construisez-les avec un
**builder** pour éviter les erreurs d'échappement :

```vb
' <connectionStrings>
'   <add name="Db" providerName="System.Data.SqlClient" connectionString="..." />
' </connectionStrings>
Dim cs = ConfigurationManager.ConnectionStrings("Db")
Dim fabrique = DbProviderFactories.GetFactory(cs.ProviderName)   ' code agnostique (cf. 15.1)
Using cnx = fabrique.CreateConnection()
    cnx.ConnectionString = cs.ConnectionString
    cnx.Open()
End Using
```

```vb
' Construire une chaîne proprement
Dim b As New SqlConnectionStringBuilder() With {
    .DataSource = ".\SQLEXPRESS", .InitialCatalog = "Ventes", .IntegratedSecurity = True}
Dim chaine = b.ConnectionString
```

> 🔗 Identifiants en **configuration** (chiffrés si besoin), compte au **moindre privilège** — rappel
> de la **section 15.2**.

---

## 6. ⚠️ Adaptations de **dialecte SQL** (Access/Jet → T-SQL)

Indispensables **uniquement dans le scénario B** (déplacement vers SQL Server). Si vous **conservez**  
Access (ACE), le dialecte Jet/ACE continue de fonctionner via le provider.

| Élément | Access / Jet (ACE) | SQL Server (T-SQL) |
|---------|--------------------|--------------------|
| **Littéral date** | `#2024-01-31#` | `'2024-01-31'` |
| **Booléen Vrai/Faux** | `True`/`False` = **0 / -1** | `bit` = **0 / 1** |
| **Jokers `LIKE`** | `*` / `?` (UI Access) — `%`/`_` en ANSI/OLE DB | `%` / `_` |
| **Concaténation** | `&` (et `+`) | `+` |
| **Conditionnel** | `IIf(c, a, b)` | `CASE WHEN c THEN a ELSE b END` |
| **Null → valeur** | `Nz(x, 0)` | `ISNULL(x, 0)` / `COALESCE(...)` |
| **Sous-chaîne** | `Mid`, `Left`, `Len`, `InStr` | `SUBSTRING`, `LEFT`, `LEN`, `CHARINDEX` |
| **Date/heure** | `Now()`, `Date()`, `DateAdd`, `DatePart` | `GETDATE()`/`SYSDATETIME()`, `DATEADD`, `DATEPART` |
| **Marqueur de paramètre** | `?` (positionnel) | `@nom` (nommé) |
| **Nouvelle clé auto** | `SELECT @@IDENTITY` | **`SCOPE_IDENTITY()`** / clause `OUTPUT` |

> ⚠️ **Le piège des littéraux date `#…#`** (et des apostrophes) **disparaît dès qu'on utilise des
> paramètres** (section 15.2) : on n'écrit plus la date dans le texte SQL. C'est une raison de plus de
> **tout paramétrer**.

> ℹ️ Pour récupérer la clé d'une insertion, **préférez `SCOPE_IDENTITY()`** à `@@IDENTITY` sur SQL
> Server : `@@IDENTITY` peut renvoyer la clé d'un **déclencheur** déclenché en cascade. Le booléen
> **-1 vs 1** prolonge le thème « `True = -1` » de VB6 (**module 7.6**) — attention aux données
> migrées.

---

## 7. Migrer la base elle-même vers SQL Server (SSMA)

Le scénario B passe par **SSMA for Access** :

- migre **schéma** et **données**, avec un **mappage de types** Access → SQL Server (Texte → `nvarchar`,
  Mémo → `nvarchar(max)`, Numérique/Long/Double → `int`/`bigint`/`float`, Monétaire → `money`/`decimal`,
  Date/Heure → `datetime2`/`date`, Oui/Non → `bit`…) ;
- produit un **rapport** des points à adapter (requêtes, fonctions Access).

Après la migration de la base, il reste à **adapter les requêtes** (tableau du §6) et à **basculer le  
provider** (`OleDb` → `SqlClient`) et la chaîne de connexion.

---

## 8. ⚠️ Pièges silencieux à retenir

- **Bitness** : **Jet** = 32 bits **uniquement** ; **ACE** doit **correspondre** à la bitness du
  processus ; **redistribuable ACE** par poste (déploiement, module 18).
- **Littéraux date** `#…#` (Access) vs `'…'` (SQL Server) → **disparaissent avec les paramètres**.
- **Booléens** : **-1** (Access) vs **1** (SQL Server `bit`).
- **Jokers `LIKE`** : `*`/`?` vs `%`/`_`.
- **Concaténation** : `&` vs `+`.
- **`IIf`/`Nz`** → `CASE`/`ISNULL`/`COALESCE`.
- **Marqueur de paramètre** : `?` (OLE DB) vs `@nom` (SqlClient) — rappel 15.2.
- **Récupération d'identité** : `@@IDENTITY` vs **`SCOPE_IDENTITY()`**.
- **Identifiants en dur** dans la chaîne → **configuration** + chiffrement.
- **`Encrypt`/`TrustServerCertificate`** : défauts variables → **explicites**.
- **MARS** : nécessaire pour plusieurs lecteurs sur **une** connexion.
- **Sécurité Access `.mdw`** : disparue avec `.accdb`.

> 🔗 Voir la **section 15.2** (paramètres, configuration, sécurité), la **section 15.1** (providers,
> `DbProviderFactory`), le **module 7.6** (`True = -1`), le **module 18** (déploiement, redistribuable
> ACE) et l'**Annexe B** (culture des dates/nombres).

---

**Section suivante → 15.6 — Transactions (de l'`ADO` au `DbTransaction`)**, pour fiabiliser les
écritures groupées.

⏭️ [Transactions (de l'`ADO` au `DbTransaction`)](/15-acces-donnees/06-transactions.md)
