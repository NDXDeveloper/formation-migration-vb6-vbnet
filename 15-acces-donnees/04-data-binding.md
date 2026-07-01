🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.4 Liaison de données (`Data` control / *data binding* VB6 → `BindingSource`)

> En VB6, un **contrôle de données** (`Data`, `Adodc`) servait d'intermédiaire entre un `Recordset` et
> les contrôles liés, et ses **flèches** déplaçaient l'enregistrement courant pour tout le monde. En
> Windows Forms, ce rôle revient au **`BindingSource`** : un composant **non visuel** placé **entre**
> la source de données et les contrôles, qui gère l'**élément courant**, la **navigation**, le **tri/
> filtre** et la **validation des modifications**.

**Module 15 — Accès aux données** · Cible : .NET Framework 4.7.2 · `System.Windows.Forms` (liaison de
données)

---

## 1. Le modèle VB6

- Le **`Data` control** (DAO/Jet) : `DatabaseName` + `RecordSource` ; les contrôles liés posent
  `DataSource = Data1` et `DataField = "Colonne"`. Des **flèches** intégrées font
  `MoveNext`/`Previous`/`First`/`Last`.
- L'**`Adodc`** (ADO Data Control) : même principe, mais via ADO (`ConnectionString` + `RecordSource`),
  souvent avec `DataGrid`/`DataCombo`/`DataList`.
- **Contrôles liés** : `TextBox`/`Label`/`CheckBox` liés à **un champ** (liaison simple) ;
  `DBGrid`/`DataGrid` liés à **tout** le recordset (liaison complexe) ; `DataCombo`/`DataList` pour les
  listes de **référence**.

Le modèle : **le `Recordset` est la source, le contrôle de données le conduit, et la position  
courante est partagée** par tous les contrôles liés.

---

## 2. Le modèle .NET : le `BindingSource` au centre

Le **`BindingSource`** est le successeur direct du contrôle de données. Il s'intercale entre la
**source** (`DataTable`, `DataView`, `List(Of T)`, `BindingList(Of T)`, `DataSet` + `DataMember`…) et
les **contrôles**, et gère :

- l'**élément courant** et la **navigation** (`Position`, `Current`, `MoveNext`/`MoveFirst`…) ;
- le **tri/filtre** (`Sort`, `Filter`) **quand la source le supporte** ;
- l'**ajout/suppression** (`AddNew`, `RemoveCurrent`) et la **validation** (`EndEdit`, `CancelEdit`).

Deux types de liaison :

- **Liaison simple** : une **propriété** d'un contrôle ↔ **un champ** (`TextBox.Text` ↔ `"Nom"`).
- **Liaison complexe** : un contrôle de **liste** ↔ **toute** la source (`DataGridView`, `ComboBox`,
  `ListBox`).

Et deux compagnons visuels : le **`BindingNavigator`** (la barre de navigation toute faite,  
l'équivalent des flèches du `Data` control) et le **`DataGridView`** (la grille, cf. 14.6).

---

## 3. La correspondance

| VB6 | .NET |
|-----|------|
| `Data` control / `Adodc` | **`BindingSource`** |
| `DataSource` + `DataField` (contrôle lié) | `ctrl.DataBindings.Add("Text", bs, "Champ")` (liaison **simple**) |
| `DBGrid` / `DataGrid` lié | `DataGridView.DataSource = bs` (liaison **complexe**) |
| `DataCombo` / `DataList` | `ComboBox`/`ListBox` + `DataSource`/`DisplayMember`/`ValueMember` |
| Flèches de navigation du `Data` control | **`BindingNavigator`** (ou `bs.MoveNext`…) |
| `Recordset` (source) | `DataTable` / `DataView` / `BindingList(Of T)` |
| `Recordset.Filter` / `Sort` | `bs.Filter` / `bs.Sort` (si la source le supporte) |

---

## 4. Liaison simple et complexe

```vb
Private _table As New DataTable()
Private _bs As New BindingSource()
Private _adaptateur As New SqlDataAdapter("SELECT Id, Nom, Ville FROM Clients", cnx)

Private Sub Charger()
    _adaptateur.Fill(_table)
    _bs.DataSource = _table                     ' la source du BindingSource

    ' Liaison SIMPLE : une propriété d'un contrôle <-> un champ
    txtNom.DataBindings.Add("Text", _bs, "Nom")
    txtVille.DataBindings.Add("Text", _bs, "Ville")

    ' Liaison COMPLEXE : une grille <-> toute la liste
    dgvClients.DataSource = _bs

    ' Navigation prête à l'emploi (flèches + ajouter/supprimer)
    navClients.BindingSource = _bs
End Sub
```

> ℹ️ Tous les contrôles liés au **même** `BindingSource` partagent l'**enregistrement courant** — comme
> les contrôles liés au `Data` control VB6. Pour lier un `DataSet`, posez `bs.DataSource = jeu` **et**
> `bs.DataMember = "Clients"` (le nom de la table).

---

## 5. Le flux de mise à jour : du contrôle à la base

La chaîne est : **contrôle → `BindingSource` → `DataTable`/objet → (`DataAdapter.Update`) → base**.

```vb
Private Sub Enregistrer()
    _bs.EndEdit()                          ' valide la modification EN COURS dans la source
    Using builder As New SqlCommandBuilder(_adaptateur)
        _adaptateur.Update(_table)         ' renvoie les modifications en base (cf. 15.3)
    End Using
End Sub
```

> ⚠️ **Piège fréquent** : appeler `Update` **sans** `EndEdit()` au préalable. La **dernière**
> modification (celle de la ligne en cours d'édition, si le focus n'a pas quitté le contrôle) n'est
> alors **pas** prise en compte. Toujours `EndEdit()` **avant** `Update`.

---

## 6. Listes déroulantes de référence (`DataCombo` → `ComboBox` lié)

Un grand classique de migration : remplacer un `DataCombo`/`DataList` par une `ComboBox` liée à une  
liste de **référence**, dont la **sélection** alimente un champ de l'enregistrement courant.

```vb
' Liste de référence (ex. villes)
cboVille.DataSource = _tableVilles          ' source de référence
cboVille.DisplayMember = "Libelle"          ' ce qui s'AFFICHE
cboVille.ValueMember = "VilleId"            ' la CLÉ stockée
' Lier la sélection au champ de l'enregistrement courant
cboVille.DataBindings.Add("SelectedValue", _bs, "VilleId")
```

> ℹ️ Renseignez `DisplayMember`/`ValueMember` **puis** `DataSource`, **avant** d'ajouter la liaison de
> `SelectedValue` (l'ordre évite des comportements surprenants). Activez `FormattingEnabled` pour bien
> gérer les valeurs absentes/`DBNull`.

---

## 7. Lier des **objets métier** (vers l'avenir)

Au-delà des `DataTable`, le `BindingSource` lie aussi des **collections d'objets**. Pour du code neuf,  
c'est souvent plus propre :

```vb
' Une collection d'objets métier (notifications d'ajout/suppression)
Private _clients As New BindingList(Of Client)()
_bs.DataSource = _clients
```

- **`BindingList(Of T)`** notifie les **ajouts/suppressions** (contrairement à `List(Of T)`).
- Pour une **liaison bidirectionnelle** des propriétés, la classe `Client` implémente
  **`INotifyPropertyChanged`** (sinon l'UI ne se rafraîchit pas quand une propriété change).

> 🔗 Cette voie **découple** l'UI de la base (module 5.3) et **prépare** le saut vers .NET moderne
> (module 20). Elle reste optionnelle : un `DataTable` lié suffit pour une première migration.

---

## 8. Format/Parse et **culture**

La liaison peut **convertir** entre la donnée et l'affichage (formatage de nombres/dates, code →  
libellé) — là où VB6 mettait du code dans les événements du contrôle de données :

```vb
Dim b = txtMontant.DataBindings.Add("Text", _bs, "Montant")
b.FormattingEnabled = True
b.FormatString = "C2"          ' format monétaire
' ou b.Format / b.Parse pour des conversions sur mesure
```

- `DataSourceUpdateMode` règle **quand** le contrôle réécrit dans la source (`OnValidation` par
  défaut, `OnPropertyChanged`, `Never`).
- ⚠️ Le formatage des valeurs liées (montants, dates) est **sensible à la culture** → cohérence avec
  les paramètres régionaux (**Annexe B.10**).

---

## 9. ⚠️ Pièges silencieux à retenir

- **`EndEdit()` avant `Update`** : sinon la dernière modification est perdue.
- **`DataMember`** : obligatoire pour lier un `DataSet` (nom de la table) ; inutile pour un `DataTable`
  lié directement.
- **`Filter`/`Sort`** ne marchent que si la source les supporte (`DataTable`/`DataView` oui ;
  `List(Of T)` non → utiliser une vue triable).
- **`List(Of T)` ne notifie pas** les ajouts/suppressions → préférer **`BindingList(Of T)`** ; et
  **`INotifyPropertyChanged`** pour les modifications de propriétés.
- **Listes de référence** : confusion `SelectedValue`/`SelectedItem` ; respecter l'ordre
  `DisplayMember`/`ValueMember`/`DataSource` ; gérer le `DBNull`.
- **Élément courant partagé** : plusieurs contrôles sur **un** `BindingSource` partagent la position
  (voulu) ; séparez les sources si ce n'est pas le but.
- **Thread d'UI** : la liaison se fait sur le **thread de l'interface**.
- **Colonnes auto** du `DataGridView` : `AutoGenerateColumns` peut exposer des colonnes non voulues →
  définir/masquer les colonnes.
- **Culture/format** des valeurs liées (Annexe B.10).

> 🔗 Voir la **section 15.3** (`DataTable`/`DataAdapter.Update`), la **section 14.6** (`DataGridView`),
> le **module 5.3** (découplage), le **module 20** (.NET moderne) et l'**Annexe B.10** (culture).

---

**Section suivante → 15.5 — Bases Access (`.mdb`/`.accdb`) et SQL Server : chaînes de connexion et
adaptations**, pour relier concrètement chaque type de base.

⏭️ [Bases Access (`.mdb`/`.accdb`) et SQL Server : chaînes de connexion et adaptations](/15-acces-donnees/05-access-sqlserver.md)
