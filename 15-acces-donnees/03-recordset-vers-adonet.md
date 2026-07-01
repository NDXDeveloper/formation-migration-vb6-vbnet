🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.3 `Recordset` → `DataReader` / `DataSet` / `DataTable` (curseurs vs déconnecté) ⚠️

> Le `Recordset` de VB6 était un **couteau suisse** : lecture, navigation avant/arrière, comptage,
> tri, filtre, **édition en place**, verrouillage. ADO.NET **éclate** ce rôle en **deux formes** —
> `DataReader` (connecté, avant-seulement) et `DataTable`/`DataSet` (déconnecté, en mémoire) — et **la
> correspondance n'est pas 1:1**. Le travail consiste à **choisir la bonne forme** selon l'usage, en
> sachant que **certains comportements de curseur ne se transposent pas**.

**Module 15 — Accès aux données** · Cible : .NET Framework 4.7.2 · `System.Data`

---

## 1. La question à se poser : **connecté** ou **déconnecté** ?

| Forme | Caractère | Idéal pour |
|-------|-----------|------------|
| **`DataReader`** | **Connecté**, **avant seulement**, **lecture seule**, rapide, peu de mémoire | Lecture séquentielle unique (remplir une liste, un état), gros volumes en flux |
| **`DataTable` / `DataSet`** (via `DataAdapter`) | **Déconnecté**, accès **aléatoire**, triable/filtrable, **éditable** en mémoire | Grilles, navigation, **édition**, liaison, cache, jeux multiples/relations |

**Grille de choix selon l'usage du `Recordset`** :

| Usage VB6 | Forme ADO.NET |
|-----------|---------------|
| Lecture séquentielle unique | **`DataReader`** |
| Revenir en arrière / accès aléatoire | **`DataTable`** |
| `RecordCount` connu **d'avance** | **`DataTable`** (ou requête `COUNT`) |
| Tri/filtre client (`Sort`/`Filter`) | **`DataTable`** + `DataView` |
| Édition en mémoire puis renvoi | **`DataTable`** + `DataAdapter` |
| Plusieurs jeux de résultats / relations | **`DataSet`** (plusieurs `DataTable`) |
| Très gros volume lu une fois | **`DataReader`** (flux) |

---

## 2. Correspondance des **curseurs** et **verrous**

| ADO (curseur / verrou) | ADO.NET | Note |
|------------------------|---------|------|
| `adOpenForwardOnly` + `adLockReadOnly` | `DataReader` | Correspondance **directe** |
| `adUseClient` + `adLockBatchOptimistic` (recordset déconnecté) | `DataTable`/`DataSet` + `DataAdapter` | L'**analogue direct** |
| `adOpenStatic` / `adOpenKeyset` (instantané défilable) | `DataTable` | Accès **aléatoire** en mémoire |
| `adOpenDynamic` (voit les changements **live**) | **Aucun équivalent** | Re-requêter ; pas de vue « vivante » |
| `adLockPessimistic` (verrou pessimiste) | **Aucun équivalent** | Modèle **optimiste** (§7) |

> ⚠️ Les **deux dernières lignes** sont des **ruptures**, pas des correspondances : ADO.NET déconnecté
> **ne voit pas** les modifications concurrentes et **ne verrouille pas** les lignes pendant l'édition.

---

## 3. Navigation et propriétés : ce qui survit, ce qui **disparaît**

| `Recordset` (ADO) | `DataReader` | `DataTable` |
|-------------------|--------------|-------------|
| `MoveNext` | `Read()` (la boucle) | itérer `Rows` |
| `MoveFirst`/`MovePrevious`/`MoveLast` | **❌ impossible** (avant seulement) | **✅** index dans `Rows` |
| `BOF`/`EOF` | `Read()` renvoie `False` à la fin ; `HasRows` | `Rows.Count` |
| `RecordCount` | **❌ inexistant** | **✅** `Rows.Count` |
| `Filter` / `Sort` | ❌ | **✅** `DataView` (`RowFilter`/`Sort`) |
| `Find` | ❌ | **✅** `Select(filtre)` / `DataView` |
| `Fields("n").Value` / `Fields(0)` | `lecteur("n")` / `lecteur(0)` (Object) ou accesseurs typés | `row("n")` / `row(0)` |
| `IsNull(champ)` | `lecteur.IsDBNull(i)` | `row.IsNull(col)` |
| `GetRows` (tableau) | boucle `Read()` | `table` directement |
| `NextRecordset` | `NextResult()` | plusieurs `DataTable` |

> ⚠️ **Deux pièges très fréquents** :
> 1. **`RecordCount`** : un `DataReader` **ne connaît pas** le nombre de lignes (et même en ADO,
>    `RecordCount` valait souvent **-1** en *forward-only*). Si le code en a besoin **d'avance**,
>    utilisez un **`DataTable`** ou une requête **`COUNT`** séparée.
> 2. **`MovePrevious`/`MoveLast`** : impossibles sur un `DataReader`. Tout code qui **revient en
>    arrière** impose un **`DataTable`**.

---

## 4. Lire en flux : le `DataReader` (connecté)

```vb
' VB6 — Recordset forward-only, lecture seule
Set rs = cn.Execute("SELECT Id, Nom FROM Clients")
Do Until rs.EOF
    Debug.Print rs!Id, rs!Nom
    rs.MoveNext
Loop

' ADO.NET — DataReader (même rôle : flux avant-seulement)
Using cmd As New SqlCommand("SELECT Id, Nom FROM Clients", cnx)
    cnx.Open()
    Using lecteur = cmd.ExecuteReader()
        While lecteur.Read()
            Dim id As Integer = lecteur.GetInt32(0)
            Dim nom As String = If(lecteur.IsDBNull(1), "", lecteur.GetString(1))
        End While
    End Using
End Using
```

- `Read()` **avance ET signale la fin** (pas de `BOF`/`EOF` séparés, pas de `MoveFirst` initial).
- **Accesseurs typés** (`GetInt32`, `GetString`, `GetDecimal`, `GetDateTime`) plus sûrs que
  l'indexeur (qui renvoie `Object`).
- `NextResult()` passe au **jeu de résultats suivant** (≈ `NextRecordset`).

> ⚠️ Le `DataReader` **tient la connexion ouverte** : **ne faites pas** de traitement lent en plein
> milieu de la boucle (appels réseau, UI bloquante). Si une méthode **renvoie** un lecteur, utilisez
> `ExecuteReader(CommandBehavior.CloseConnection)` pour que la **fermeture du lecteur ferme la
> connexion**. Et **toujours** un `Using` (sinon la connexion reste ouverte).

---

## 5. Travailler en mémoire : `DataTable` / `DataSet` (déconnecté)

```vb
Dim table As New DataTable()
Using adaptateur As New SqlDataAdapter("SELECT Id, Ville FROM Clients", cnx)
    adaptateur.Fill(table)          ' ouvre, remplit, PUIS FERME la connexion
End Using
' 'table' est exploitable hors connexion : accès aléatoire, tri, filtre, modifications
```

- **Accès aléatoire** : `table.Rows(i)("Ville")`.
- **Tri/filtre** via `DataView` (l'équivalent de `Sort`/`Filter`) :
  ```vb
  Dim vue As New DataView(table) With {.RowFilter = "Ville = 'Paris'", .Sort = "Nom"}
  ```
- On peut aussi **charger** un `DataTable` depuis un lecteur : `table.Load(lecteur)`.
- Le **`DataSet`** regroupe **plusieurs `DataTable`** et leurs **`DataRelation`** (utile pour des jeux
  liés maître/détail).

---

## 6. Éditer et **renvoyer** les modifications

Le `Recordset` éditait **en place** ; en ADO.NET déconnecté, on modifie **la mémoire** puis on
**pousse** les changements :

```vb
' VB6 — édition en place + renvoi par lot
rs.Open "SELECT Id, Ville FROM Clients", cn, adOpenStatic, adLockBatchOptimistic
rs!Ville = "Lyon"
rs.Update
rs.UpdateBatch

' ADO.NET — modifications EN MÉMOIRE puis DataAdapter.Update
Dim table As New DataTable()
Using adaptateur As New SqlDataAdapter("SELECT Id, Ville FROM Clients", cnx)
    adaptateur.Fill(table)
    table.Rows(0)("Ville") = "Lyon"                 ' modifie en mémoire
    ' table.Rows.Add(...) / row.Delete() pour insérer / supprimer
    Using builder As New SqlCommandBuilder(adaptateur)   ' génère Insert/Update/Delete
        adaptateur.Update(table)                    ' renvoie les modifications en base
    End Using
End Using
```

- Chaque ligne porte un **`RowState`** (`Added`/`Modified`/`Deleted`/`Unchanged`) ; `GetChanges()`
  isole les modifications.
- Le **`SqlCommandBuilder`** génère automatiquement les commandes pour un **cas simple** : **une seule
  table** et une **clé primaire**. ⚠️ Il **ne gère pas** les jointures, peut être **moins performant**
  et impose une clé — beaucoup d'équipes écrivent des **commandes explicites** paramétrées (section
  15.2) à la place.
- **Rappel** : le `DataReader` est **en lecture seule** — toute édition impose un `DataTable`.

---

## 7. ⚠️ Le vrai changement : la **concurrence** (pessimiste → optimiste)

C'est une **rupture de modèle**, pas une question de syntaxe.

VB6 utilisait souvent un **verrouillage pessimiste** (`adLockPessimistic`) : la ligne était
**verrouillée** côté serveur pendant l'édition. ADO.NET déconnecté est **optimiste** : on lit, **la
connexion se ferme**, un tiers peut modifier la ligne, puis votre `Update` **vérifie** (clause `WHERE`  
sur les valeurs d'origine) et **peut échouer** :

- en cas de conflit, `DataAdapter.Update` lève **`DBConcurrencyException`** (zéro ligne affectée) ;
- vous devez **concevoir** la résolution : recharger, fusionner, « dernier qui écrit gagne », etc.

> ⚠️ Les schémas reposant sur le **verrou pessimiste** doivent être **repensés**. Ne réintroduisez pas
> un verrouillage long « comme avant » : préférez des transactions **courtes** (section 15.6) et une
> gestion **explicite** des conflits.

---

## 8. Performances et mémoire

| | `DataReader` | `DataTable`/`DataSet` |
|---|--------------|------------------------|
| Mémoire | **Faible** (flux) | **Tout en mémoire** |
| Connexion | **Tenue** pendant la lecture | **Libérée** après `Fill` |
| Accès | Avant seulement | Aléatoire |
| Gros volumes | ✅ (streaming) | ⚠️ coûteux |

> ⚠️ **Ne chargez pas** un million de lignes dans un `DataSet` pour ne les parcourir **qu'une fois** :
> utilisez un **`DataReader`** (ou une **pagination**). À l'inverse, pour des données **réutilisées**,
> triées, filtrées ou éditées, le `DataTable` libère vite la connexion et simplifie le code.

---

## 9. Exemple : un cas qui **impose** le `DataTable`

```vb
' VB6 — utilise RecordCount ET MovePrevious -> accès ALÉATOIRE nécessaire
n = rs.RecordCount
rs.MoveLast : rs.MovePrevious

' ADO.NET — le DataReader ne le permet pas. -> DataTable
Dim t As New DataTable()
Using ad As New SqlDataAdapter("SELECT Id, Nom FROM Clients", cnx)
    ad.Fill(t)
End Using
Dim n As Integer = t.Rows.Count
Dim avantDernier As DataRow = t.Rows(n - 2)
```

---

## 10. 🧭 Grille de décision et ⚠️ pièges

**Choisir**

1. Lecture **séquentielle unique** ou **gros volume** → **`DataReader`**.
2. **Revenir en arrière**, **compter d'avance**, **trier/filtrer/éditer** en mémoire → **`DataTable`**.
3. **Jeux multiples / relations** → **`DataSet`**.

**Pièges silencieux**

- **`RecordCount` sur un lecteur** : n'existe pas → `DataTable.Rows.Count` ou `COUNT`.
- **`MovePrevious`/`MoveLast`** : impossibles sur un lecteur → `DataTable`.
- **Lecteur = connexion tenue** : pas de traitement lent dans la boucle ; `Using` obligatoire ;
  `CommandBehavior.CloseConnection` si on renvoie le lecteur.
- **Édition** : impossible sur un lecteur (lecture seule) → `DataTable` + `DataAdapter`.
- **`DBNull`** ≠ `Nothing` : tester `IsDBNull`/`IsNull` (rappel README et **Annexe B**).
- **Concurrence** : pessimiste → **optimiste** ; gérer **`DBConcurrencyException`**.
- **`CommandBuilder`** : une seule table + clé, limites de performance → commandes explicites.
- **Gros volume dans un `DataSet`** : coûteux → `DataReader`/pagination.

> 🔗 Voir la **section 15.2** (commandes/paramètres pour des `Update` explicites), la **section 15.4**
> (liaison de `DataTable` aux contrôles), la **section 15.6** (transactions courtes), le **module 12**
> (`IDisposable`) et l'**Annexe B** (`DBNull`, culture).

---

**Section suivante → 15.4 — Liaison de données (`Data` control / *data binding* VB6 → `BindingSource`)**,
pour relier ces `DataTable` aux contrôles de l'interface.

⏭️ [Liaison de données (`Data` control / *data binding* VB6 → `BindingSource`)](/15-acces-donnees/04-data-binding.md)
