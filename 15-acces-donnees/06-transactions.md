🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.6 Transactions (de l'`ADO` au `DbTransaction`)

> En VB6, une transaction ADO est **ambiante** sur la connexion : `cn.BeginTrans`, et **toutes** les
> commandes suivantes y participent. En ADO.NET avec `DbTransaction`, c'est **explicite** : on crée la
> transaction **et l'on enrôle chaque commande** (`cmd.Transaction = tx`). Pour retrouver le modèle
> ambiant — voire des transactions **distribuées** —, on dispose de **`TransactionScope`**.

**Module 15 — Accès aux données** · Cible : .NET Framework 4.7.2 · `System.Data`,
`System.Transactions`

---

## 1. Le modèle ADO (VB6)

Les transactions sont portées par l'objet **`Connection`** :

```vb
cn.BeginTrans
' ... commandes ...
If ok Then cn.CommitTrans Else cn.RollbackTrans
```

- **Ambiant** : une fois `BeginTrans` appelé, tout ce qui passe par la connexion est dans la
  transaction jusqu'au `CommitTrans`/`RollbackTrans`.
- **Imbrication** possible (`BeginTrans` renvoie le niveau).
- Le `RollbackTrans` se fait typiquement dans le gestionnaire **`On Error`**.

---

## 2. Le modèle ADO.NET : le `DbTransaction` **explicite**

On crée la transaction **depuis la connexion**, et l'on **enrôle chaque commande** :

```vb
Using tx = cnx.BeginTransaction()
    Using cmd As New SqlCommand("UPDATE ...", cnx, tx)   ' 3e argument = la transaction
        cmd.ExecuteNonQuery()
    End Using
    tx.Commit()
End Using
```

> ⚠️ **LA différence avec ADO.** Le modèle n'est **pas** ambiant : si vous **oubliez** d'enrôler une
> commande (`cmd.Transaction = tx`, ou le constructeur à 3 arguments), SqlClient **lève une erreur**
> (la commande exige une transaction quand la connexion en a une en cours). Chaque commande doit
> **explicitement** désigner la transaction.

- **`tx.Commit()`** / **`tx.Rollback()`**.
- `SqlTransaction` est **`IDisposable`** : un `Dispose` **sans `Commit`** effectue un **rollback**. Le
  `Using` garantit donc l'annulation sur le chemin d'exception.
- **Niveau d'isolation** : `cnx.BeginTransaction(IsolationLevel.Serializable)` (défaut **`ReadCommitted`** ;
  aussi `ReadUncommitted`, `RepeatableRead`, `Snapshot`).
- `BeginTransaction` exige une connexion **ouverte**.

---

## 3. La correspondance

| ADO (VB6) | ADO.NET `DbTransaction` | `TransactionScope` |
|-----------|--------------------------|--------------------|
| `cn.BeginTrans` | `tx = cn.BeginTransaction()` | `New TransactionScope()` |
| `cn.CommitTrans` | `tx.Commit()` | `scope.Complete()` |
| `cn.RollbackTrans` | `tx.Rollback()` (ou `Dispose` sans `Commit`) | **ne pas** appeler `Complete` |
| `IsolationLevel` | `BeginTransaction(IsolationLevel…)` | `TransactionOptions.IsolationLevel` |
| Transactions **imbriquées** | **Points de sauvegarde** (`Save`/`Rollback(nom)`) | Scopes imbriqués (`Required`/`RequiresNew`/`Suppress`) |
| Enrôler une commande | **`cmd.Transaction = tx`** (obligatoire) | **Automatique** (ambiant) |

---

## 4. Le motif idiomatique (`DbTransaction`)

```vb
Using cnx As New SqlConnection(chaine)
    cnx.Open()
    Using tx = cnx.BeginTransaction()
        Try
            Using debit As New SqlCommand(
                  "UPDATE Comptes SET Solde = Solde - @m WHERE Id = @id", cnx, tx)
                debit.Parameters.AddWithValue("@m", montant)
                debit.Parameters.AddWithValue("@id", idDebit)
                debit.ExecuteNonQuery()
            End Using
            Using credit As New SqlCommand(
                  "UPDATE Comptes SET Solde = Solde + @m WHERE Id = @id", cnx, tx)
                credit.Parameters.AddWithValue("@m", montant)
                credit.Parameters.AddWithValue("@id", idCredit)
                credit.ExecuteNonQuery()
            End Using
            tx.Commit()
        Catch
            tx.Rollback()
            Throw
        End Try
    End Using
End Using
```

> ℹ️ Le `Using` sur `tx` **annulerait déjà** la transaction si une exception court-circuitait le
> `Commit`. Le `Rollback` explicite + `Throw` rend l'intention claire et **relance** l'erreur.

---

## 5. Transactions imbriquées → **points de sauvegarde**

ADO.NET **ne reproduit pas** l'imbrication d'ADO via `BeginTransaction`. À la place, on utilise des
**points de sauvegarde** (savepoints) :

```vb
Using tx = cnx.BeginTransaction()
    ' ... opérations principales ...
    tx.Save("avantDetail")                ' point de sauvegarde
    Try
        ' ... opérations de détail ...
    Catch
        tx.Rollback("avantDetail")        ' annule UNIQUEMENT depuis le point (pas tout)
    End Try
    tx.Commit()
End Using
```

---

## 6. `TransactionScope` : le modèle **ambiant** (et le distribué)

`TransactionScope` rétablit le modèle **implicite** : les connexions **ouvertes à l'intérieur** du
scope s'**enrôlent automatiquement** — pas de `cmd.Transaction` à gérer.

```vb
Using scope As New TransactionScope()
    Using cnx As New SqlConnection(chaine)
        cnx.Open()                          ' s'enrôle AUTOMATIQUEMENT dans la transaction ambiante
        Using cmd As New SqlCommand("UPDATE ...", cnx)
            cmd.ExecuteNonQuery()
        End Using
    End Using
    scope.Complete()                        ' valide ; SANS cet appel -> annulation au Dispose
End Using
```

- **Plusieurs ressources** : un `TransactionScope` peut englober **plusieurs connexions/ressources**
  de façon atomique.
- **`TransactionScopeOption`** (`Required`/`RequiresNew`/`Suppress`) gère l'imbrication de scopes.
- Nécessite une référence à **`System.Transactions`** ; pour l'asynchrone, `TransactionScopeAsyncFlowOption`.

> ⚠️ **Piège d'isolation : le défaut est `Serializable`.** Contrairement à `BeginTransaction` (défaut
> **`ReadCommitted`**, §2), un `TransactionScope` construit **sans options** utilise le niveau
> **`Serializable`** — le **plus strict**, propice aux **interblocages** sous charge. Positionnez-le
> **explicitement** via `TransactionOptions` :
> ```vb
> Dim options As New TransactionOptions With {.IsolationLevel = IsolationLevel.ReadCommitted}
> Using scope As New TransactionScope(TransactionScopeOption.Required, options)
>     ' ... connexions/commandes ...
>     scope.Complete()
> End Using
> ```

> ⚠️ **Le piège du distribué (MSDTC).** Dès que **plusieurs ressources durables** s'enrôlent, la
> transaction peut être **promue** en transaction **distribuée** (MSDTC), qui doit être **configurée
> et active**, avec un **surcoût**. Une **seule** `SqlConnection` reste locale (enrôlement
> mono-phase promouvable), mais l'ouverture d'une **deuxième** connexion peut déclencher la promotion
> — souvent une **surprise**. À surveiller.

---

## 7. Choisir : `DbTransaction` ou `TransactionScope`

| Critère | `DbTransaction` (`BeginTransaction`) | `TransactionScope` |
|---------|--------------------------------------|--------------------|
| Portée | **Une** connexion | **Plusieurs** ressources possibles |
| Enrôlement | **Explicite** (`cmd.Transaction`) | **Ambiant** (automatique) |
| MSDTC | Jamais | **Possible** (promotion) |
| Contrôle fin (isolation, savepoints) | ✅ | Moindre |
| Cas par défaut | **Transaction locale mono-base** | Atomicité **multi-ressources** / modèle ambiant |

---

## 8. Transactions et `DataAdapter.Update`

Pour qu'une **mise à jour par lot** déconnectée (section 15.3) participe à une transaction
`DbTransaction`, affectez la transaction aux **commandes** de l'adaptateur :

```vb
adaptateur.SelectCommand.Transaction = tx
adaptateur.InsertCommand.Transaction = tx
adaptateur.UpdateCommand.Transaction = tx
adaptateur.DeleteCommand.Transaction = tx
adaptateur.Update(table)                   ' le lot participe à la transaction
```

> ℹ️ Avec `TransactionScope`, l'enrôlement étant **ambiant**, ce câblage est inutile : ouvrez la
> connexion **dans** le scope.

---

## 9. ⚠️ Pièges silencieux à retenir

- **Enrôler les commandes** (`cmd.Transaction = tx`) avec `DbTransaction` — sinon **erreur**. (C'est
  automatique avec `TransactionScope`.)
- **`Dispose` sans `Commit` = rollback** : utiliser `Using` pour annuler proprement sur exception ;
  **ne pas oublier `Commit`** (ou `scope.Complete()`) sur le chemin nominal.
- **Garder les transactions COURTES** : ne **jamais** tenir une transaction ouverte pendant une
  interaction utilisateur ou un traitement lent (verrous, blocages) — prolonge le thème **concurrence**
  de la section 15.3.
- **Promotion MSDTC** (`TransactionScope`) avec plusieurs connexions/ressources → configuration et
  surcoût ; souvent inattendue.
- **Niveau d'isolation** : défaut **`ReadCommitted`** pour `BeginTransaction`, mais **`Serializable`**
  (le plus strict) pour `TransactionScope` → le **fixer explicitement** pour éviter les interblocages.
- **`DataAdapter.Update`** dans une transaction `DbTransaction` → affecter `Transaction` aux
  **commandes** de l'adaptateur.
- **Imbrication** : pas de vraie imbrication via `BeginTransaction` → **points de sauvegarde**
  (`DbTransaction`) ou **scopes imbriqués** (`TransactionScope`).

> 🔗 Voir la **section 15.2** (commandes/connexions), la **section 15.3** (concurrence,
> `DataAdapter.Update`), le **module 11** (exceptions et `Try`/`Catch`/`Finally`) et le **module 12**
> (`IDisposable`/`Using`).

---

**Section suivante → 15.7 — Reporting : Crystal Reports (compatibilité) et alternatives** ⚠️, dernier
volet du chapitre données.

⏭️ [Reporting : **Crystal Reports** (compatibilité) et alternatives](/15-acces-donnees/07-reporting.md)
