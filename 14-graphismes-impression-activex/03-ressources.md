🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.3 Images, icônes et ressources (`.frx`, `.res` → `.resx`)

> Les images et icônes embarquées dans vos formulaires suivent généralement la migration **toutes
> seules** (l'outil les dépose dans les ressources du concepteur). Le travail réel concerne le
> fichier **`.res`** du projet et les appels `LoadResString`/`LoadResPicture`/`LoadResData` : il faut
> passer d'un système indexé par **numéros** à un système .NET indexé par **noms**, typé et
> localisable.

**Module 14 — Graphismes, impression et contrôles ActiveX** · Cible : .NET Framework 4.7.2 ·
`System.Drawing`, `My.Resources`

---

## 1. Le modèle de ressources VB6

### Les fichiers `.frx` (ressources de formulaire)

Chaque formulaire (`.frm`) qui contient des propriétés **binaires** — `Picture`, `Icon`, image d'un
`PictureBox`/`Image`, `MouseIcon`, `DragIcon`, images d'un `ImageList`, longues listes… — possède un
fichier compagnon **`.frx`**. Le `.frm` (texte) y renvoie par un **décalage** :

```text
Picture = "Form1.frx":0000
```

Le `.frx` est **binaire et opaque**, géré par l'EDI : on ne l'édite pas à la main.

### Le fichier `.res` (ressources de projet) et les fonctions `LoadRes…`

Un projet VB6 peut contenir **un seul** fichier `.res` (ressource Win32), édité avec un éditeur de  
ressources. Il regroupe des **tables de chaînes** (souvent pour la **localisation**), des bitmaps,  
icônes, curseurs et données binaires, indexés par **numéro**. On y accède à l'exécution :

```vb
lblTitre.Caption = LoadResString(101)
Set imgLogo.Picture = LoadResPicture(201, vbResBitmap)   ' vbResBitmap=0, vbResIcon=1, vbResCursor=2
Dim octets() As Byte
octets = LoadResData(301, "CUSTOM")
```

### Charger/enregistrer des images à l'exécution

```vb
Set Picture1.Picture = LoadPicture("C:\img\logo.bmp")
SavePicture Picture1.Picture, "C:\img\copie.bmp"
```

Les images VB6 sont des objets **`StdPicture`** (`IPictureDisp`).

---

## 2. Le modèle de ressources .NET

### Ressources de **projet** : `My.Resources` (typé, recommandé)

Les ressources applicatives partagées vivent dans `Resources.resx` (sous *My Project*/*Properties*).  
Un fichier `Resources.Designer.vb` généré expose un **accès fortement typé** — plus de numéros, plus  
de transtypage :

```vb
lblTitre.Text = My.Resources.TitreAccueil   ' String
imgLogo.Image = My.Resources.Logo           ' System.Drawing.Bitmap
Me.Icon = My.Resources.AppIcone             ' System.Drawing.Icon
Dim octets() As Byte = My.Resources.DonneesBinaires   ' Byte()
```

`My.Resources.X` renvoie le **bon type** selon la ressource : `Bitmap` pour une image, `Icon` pour
une icône, `String` pour une chaîne, `Byte()` ou un flux pour un fichier binaire/audio.

### Ressources de **formulaire** : `Form1.resx` (géré par le concepteur)

Chaque formulaire/`UserControl` a un `.resx` compagnon que le **concepteur** alimente  
automatiquement (images de contrôles, icône du formulaire, chaînes localisées). Il est lu dans
`InitializeComponent` via un `ComponentResourceManager`. **C'est l'équivalent direct du `.frx`.**

### Embarqué, `ResourceManager` et assemblys satellites

- Une ressource peut être **embarquée** dans l'assembly (compilée dans l'EXE/DLL) ou **liée** à un
  fichier externe.
- À bas niveau, tout passe par un **`ResourceManager`** ; `My.Resources` n'en est qu'un emballage
  typé. Pour un accès dynamique : `My.Resources.ResourceManager.GetObject("Logo")`.
- Pour la **localisation**, on crée un `.resx` **par culture** (`Resources.fr.resx`,
  `Resources.de.resx`…), compilé en **assembly satellite** ; le `ResourceManager` choisit le bon
  selon `CurrentUICulture` (voir §7).

---

## 3. Tableaux de correspondance

### Fichiers et formats

| VB6 | .NET | Remarque |
|-----|------|----------|
| `.frx` (binaire, par formulaire) | `Form1.resx` (XML, ressources du concepteur) | Géré automatiquement |
| `.res` (**un** par projet, Win32) | `Resources.resx` (ressources de projet, `My.Resources`) | **Plusieurs** `.resx` possibles |
| Table de chaînes localisée (`.res`) | `Resources.<culture>.resx` → **satellites** | Modèle entièrement différent (§7) |
| `StdPicture` / `IPictureDisp` | `System.Drawing.Image` / `Bitmap` / `Icon` | Modèle objet distinct |

### API d'exécution

| VB6 | .NET | Type renvoyé |
|-----|------|--------------|
| `LoadResString(id)` | `My.Resources.NomChaine` | `String` |
| `LoadResPicture(id, vbResBitmap)` | `My.Resources.NomImage` | `Bitmap` |
| `LoadResPicture(id, vbResIcon)` | `My.Resources.NomIcone` | `Icon` |
| `LoadResPicture(id, vbResCursor)` | `My.Resources.NomCurseur` | `Cursor` |
| `LoadResData(id, format)` | `My.Resources.NomBinaire` | `Byte()` / flux |
| `LoadPicture("...")` | `Image.FromFile(...)` / `New Bitmap(...)` | `Image` / `Bitmap` |
| `LoadPicture()` (effacer) | `ctrl.Image = Nothing` | — |
| `SavePicture pic, "..."` | `image.Save("...", ImageFormat.Png)` | — |

---

## 4. Du **numéro** au **nom** : la bascule de clé

En VB6, une ressource se désigne par un **entier** (`LoadResString(40001)`). En .NET, par un **nom**
(`My.Resources.MessageBienvenue`), vérifié à la compilation.

> ⚠️ **Travail manuel à prévoir.** Migrer un `.res` consiste à : recréer chaque entrée **nommée**
> dans `Resources.resx`, puis remplacer chaque `LoadRes…(numéro)` par l'accès typé correspondant. Le
> **mappage numéro → nom** est une tâche de fond : repérez les **numéros magiques** dispersés dans le
> code et constituez une table de correspondance avant de convertir.

---

## 5. Images, icônes et curseurs : `StdPicture` → `Image` / `Icon` / `Cursor`

Le type unique `StdPicture` de VB6 éclate en **trois types distincts** en .NET, **non  
interchangeables** :

| Propriété VB6 | Type .NET | Espace de noms |
|---------------|-----------|----------------|
| `.Picture` (image) | `System.Drawing.Image` / `Bitmap` | `System.Drawing` |
| `.Icon` (formulaire), icône de notification | `System.Drawing.Icon` | `System.Drawing` |
| `.MouseIcon` / `.DragIcon` (curseur) | `System.Windows.Forms.Cursor` | `System.Windows.Forms` |

Conversions utiles :

```vb
' Icône -> image : direct et sans souci
Dim img As Bitmap = monIcone.ToBitmap()
' Image -> icône : Bitmap.GetHicon() crée un HICON NATIF qu'il faut détruire
' (DestroyIcon via P/Invoke, cf. module 16), sinon FUITE de handle.
```

> ⚠️ **Verrouillage de fichier.** `Image.FromFile(chemin)` **garde le fichier verrouillé** pendant
> toute la vie de l'image (impossible alors de le supprimer/remplacer). Pour s'en affranchir, chargez
> via un flux et faites une **copie indépendante** :
> ```vb
> Private Function ChargerSansVerrou(chemin As String) As Bitmap
>     Using fs As New IO.FileStream(chemin, IO.FileMode.Open, IO.FileAccess.Read),
>           src As New Bitmap(fs)
>         Return New Bitmap(src)        ' copie autonome : flux et fichier libérables
>     End Using
> End Function
> ```

---

## 6. ⚠️ Le piège de la **copie** : `My.Resources` renvoie une **nouvelle instance**

Point crucial, souvent méconnu : **chaque accès** à `My.Resources.X` **désérialise et renvoie un  
nouvel objet**. Deux conséquences :

```vb
' À ÉVITER : un nouveau Bitmap créé à CHAQUE rendu -> churn mémoire/handles
Private Sub Panneau_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
    e.Graphics.DrawImage(My.Resources.Logo, 0, 0)
End Sub

' À PRIVILÉGIER : mettre en cache une seule instance
Private ReadOnly _logo As Bitmap = My.Resources.Logo
Private Sub Panneau_Paint2(sender As Object, e As PaintEventArgs) Handles Me.Paint
    e.Graphics.DrawImage(_logo, 0, 0)
End Sub
```

> 🔗 De plus, **un contrôle ne libère pas** l'image que vous lui affectez : `pictureBox.Image` n'est
> **pas** détruit quand le contrôle l'est. Vous en êtes **propriétaire**. C'est une application
> directe d'`IDisposable`/`Using` (module 12) : conservez une référence si l'image est réutilisée, et
> appelez `Dispose` quand vous avez fini. Comme chaque accès à `My.Resources` est une copie, **mettre
> en cache une référence unique** est le bon réflexe pour éviter à la fois les fuites et les erreurs
> « objet déjà libéré ».

---

## 7. Localisation : du *string-table* `.res` aux **assemblys satellites**

C'est un **changement de modèle**, pas une simple correspondance.

| | VB6 | .NET |
|---|-----|------|
| Stockage | Tables de chaînes par locale dans **`.res`** | `.resx` **par culture** → assemblys satellites |
| Sélection | Manuelle (`LoadResString` + logique) | Automatique via `ResourceManager` + `CurrentUICulture` |
| Formulaires | Code de rechargement à la main | `Localizable = True` + `Language` dans le concepteur |

Pour appliquer une autre langue à un formulaire **à l'exécution**, on réapplique les ressources :

```vb
Threading.Thread.CurrentThread.CurrentUICulture = New Globalization.CultureInfo("fr")
Dim res As New ComponentResourceManager(GetType(Form1))
res.ApplyResources(Me, "$this")
For Each ctrl As Control In Me.Controls
    res.ApplyResources(ctrl, ctrl.Name)
Next
```

> ℹ️ Choisissez la culture via `CurrentUICulture` (ressources d'interface), à distinguer de
> `CurrentCulture` (formats de nombres/dates) — un point repris au catalogue des pièges (séparateurs
> et formats culture-dépendants, **Annexe B.10**).

---

## 8. Ce que font — et ne font pas — les outils

- ✅ Les assistants (`Upgrade Wizard`, VBUC ; cf. **module 4**) migrent en général les **images du
  concepteur** : pictures, icônes, contenu des `ImageList`, vers les `.resx` (formulaire ou projet).
- ❌ Le **fichier `.res`** et les appels `LoadResString`/`LoadResPicture`/`LoadResData` ne sont
  **pas** convertis de façon fiable : il faut migrer le contenu vers `Resources.resx` et **réécrire**
  les appels en `My.Resources.*`. C'est le gros du travail manuel de cette section.
- ➕ Bénéfice de plateforme : les `.resx` sont du **XML lisible et comparable** (contrairement aux
  `.frx`/`.res` binaires), donc **agréables sous contrôle de version** (Git, module 6).

---

## 9. 🧭 Approche recommandée

1. Placez les ressources **partagées** (images, icônes, chaînes communes) dans les **ressources de
   projet** (`My.Resources`) pour un accès **typé** et global.
2. Laissez le **concepteur** gérer les ressources **propres à un formulaire** (`Form1.resx`).
3. Dressez la table **numéro → nom** du `.res`, puis remplacez chaque `LoadRes…` par l'accès
   `My.Resources.*` correspondant.
4. Adoptez le modèle `.resx`-par-culture + `ResourceManager` pour la **localisation** ; activez
   `Localizable = True` sur les formulaires concernés.
5. **Mettez en cache** les objets ressources réutilisés ; **libérez** ce dont vous êtes propriétaire ;
   chargez les fichiers externes **via un flux** pour éviter le verrouillage.

---

## 10. ⚠️ Pièges silencieux à retenir

- **Clés numériques → noms** : repérer les **numéros magiques** du `.res` et bâtir le mappage.
- **`My.Resources.X` = nouvelle copie à chaque accès** → mettre en cache ; **ne pas** y appeler dans
  une boucle ou dans `Paint`.
- **Le contrôle ne libère pas** l'image affectée → gérer soi-même la durée de vie (`Dispose`).
- **`Image.FromFile` verrouille le fichier** → charger via un flux et copier.
- **`Image` ≠ `Icon` ≠ `Cursor`** : types distincts ; conversions explicites (et `GetHicon` fuit si
  l'on n'appelle pas `DestroyIcon`).
- **Localisation** : modèle entièrement différent (satellites + `CurrentUICulture`), pas un simple
  `LoadResString`.
- **`LoadRes…` non migrés** par les outils → réécriture manuelle vers `My.Resources`.

> 🔗 Voir l'**Annexe B** (pièges, dont B.10 sur la culture) et le **module 12** (`IDisposable`/`Using`
> pour les images et icônes).

---

**Section suivante → 14.4 — Impression : objet `Printer` → `System.Drawing.Printing`
(`PrintDocument`)**, où le modèle d'impression passe de *séquentiel* à *événementiel*.

⏭️ [Impression : objet `Printer` → `System.Drawing.Printing` (`PrintDocument`)](/14-graphismes-impression-activex/04-impression.md)
