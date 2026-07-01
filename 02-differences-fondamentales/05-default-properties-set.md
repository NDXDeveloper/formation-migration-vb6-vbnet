🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2.5 Propriétés par défaut et mot-clé `Set` (disparu) ⚠️

> **Chapitre 2 — Différences fondamentales VB6 ↔ VB.NET**
> Deux changements qui n'en font qu'un : le `Set` de VB6 n'existait **que** pour lever l'ambiguïté créée par les propriétés par défaut. Supprimez l'une, l'autre devient inutile.

---

## Deux sujets, une seule histoire

On pourrait croire que « propriétés par défaut » et « mot-clé `Set` » sont deux questions séparées.  
Ils forment en réalité **les deux faces d'une même pièce**, et c'est en comprenant leur lien qu'on  
saisit le changement.

En VB6, une variable objet pouvait être lue de **deux façons** : soit comme une **référence** vers  
l'objet, soit comme la **valeur de sa propriété par défaut**. Il fallait donc un signal pour dire  
laquelle des deux on voulait — ce signal, c'était le mot-clé **`Set`**. VB.NET fait le ménage à la  
source : il **restreint drastiquement les propriétés par défaut**, ce qui **supprime l'ambiguïté**,  
ce qui rend **`Set` inutile** — et `Set` (comme instruction d'affectation) **disparaît**.

> ⚠️ **La grille du chapitre, appliquée ici.** Le sujet mêle du **visible** et du **silencieux** :
> le `Set` d'affectation laissé en place provoque une **erreur de compilation** (visible, le
> compilateur vous guide), mais une lecture de propriété par défaut peut **changer de sens sans
> erreur** — surtout sous `Option Strict Off`, le mode dans lequel on **démarre** souvent la
> migration (→ 6.4). C'est là que se loge le risque.

---

## 🔵 Côté VB6 : les propriétés par défaut, partout

En VB6, beaucoup d'objets exposent une **propriété par défaut** : celle qui est utilisée quand on  
manipule l'objet **sans nommer de membre**. Quelques exemples emblématiques :

| Objet VB6 | Propriété par défaut |
|---|---|
| `TextBox` | `Text` |
| `Label` | `Caption` |
| `CheckBox` / `OptionButton` | `Value` |
| `Field` (ADO) | `Value` |
| `Collection` | `Item` (paramétrée) |

Conséquence : on pouvait **omettre** le nom de la propriété par défaut, et VB6 le sous-entendait.

```vb
' VB6 — la propriété par défaut est implicite
txtNom = "Dupont"           ' signifie en réalité  txtNom.Text = "Dupont"
lblTitre = "Bonjour"        ' signifie en réalité  lblTitre.Caption = "Bonjour"

Dim s As String
s = txtNom                  ' lit  txtNom.Text
If txtA = txtB Then ...     ' compare  txtA.Text  et  txtB.Text

Dim v As Variant
v = rs("nom")               ' lit  rs("nom").Value  (propriété par défaut du Field)
```

C'est **commode**, mais **ambigu** : quand on écrit `x = monObjet`, veut-on la **valeur** (propriété  
par défaut) ou la **référence** ? D'où la nécessité d'un mot-clé pour trancher.

---

## 🔵 Côté VB6 : `Set` (et `Let`), le mot-clé qui tranche

C'est précisément le rôle de **`Set`** : distinguer l'affectation d'une **référence** de  
l'affectation d'une **valeur**.

```vb
' VB6 — deux affectations, deux significations
Dim a As Object

Set a = monObjet            ' affecte la RÉFÉRENCE (a et monObjet désignent le même objet)
a = monObjet                ' affecte la VALEUR de la propriété par défaut de monObjet
```

Symétriquement, l'affectation de **valeur** avait, elle aussi, son mot-clé — **`Let`** (`Let x = y`)
— presque toujours **omis** en pratique, mais conceptuellement présent. VB6 fonctionnait donc avec
un **trio** : `Let` (valeur), `Set` (référence), et la propriété par défaut comme cas implicite.

> 💡 Retenez le lien de cause à effet : **`Set` existe à cause des propriétés par défaut.** Sans
> propriétés par défaut ambiguës, `x = monObjet` ne pourrait signifier qu'**une** chose, et `Set`
> n'aurait plus de raison d'être.

---

## 🟢 Côté VB.NET : propriétés par défaut restreintes, `Set` supprimé

VB.NET coupe le nœud à la racine. La règle est nette :

- **Les propriétés par défaut *sans paramètre* sont supprimées.** Un `TextBox` n'a plus de propriété par défaut implicite : on **doit** écrire `.Text`. Idem pour `Caption` (devenue `.Text` en Windows Forms), `Value`, etc.
- **Les propriétés par défaut *paramétrées* (indexées) survivent.** Une collection garde son `Item` par défaut : `maCollection(1)` reste valide, car l'index `(1)` lève à lui seul l'ambiguïté.

```vb
' VB.NET — la propriété par défaut SANS paramètre n'existe plus → on la nomme
txtNom.Text = "Dupont"      ' obligatoire (txtNom = "Dupont" ne signifie plus la même chose)

Dim s As String = txtNom.Text

' …mais la propriété par défaut PARAMÉTRÉE survit :
Dim x = maListe(0)          ' OK — Item(0), l'index lève l'ambiguïté
```

Du coup, **l'ambiguïté disparaît** : `a = monObjet` ne peut plus vouloir dire « valeur par défaut ».  
Le mot-clé **`Set` (comme instruction d'affectation) n'a plus d'utilité — et il est supprimé** :

```vb
' VB6
Set a = monObjet

' VB.NET — plus de Set : affectation d'objet directe
a = monObjet
```

> 💡 **Une simplification réelle.** Sur ce point, VB.NET est **plus simple** que VB6 : une seule
> forme d'affectation, plus de `Set`/`Let` à manier. Le modèle « valeur ou référence ? » est tranché
> par la **nature du type** (type valeur vs type référence, → 2.1/2.4), non plus par un mot-clé.

---

## ⚠️ Où est le piège ? Le visible et le silencieux

### Le `Set` d'affectation laissé en place → **visible**

Si une conversion oublie un `Set`, le résultat est une **erreur de compilation** : `Set a = ...`  
n'est plus une instruction d'affectation valide. Le compilateur vous arrête — pas de danger  
silencieux ici. **Bonne nouvelle.**

### La lecture de propriété par défaut → **potentiellement silencieux**

C'est ici que se concentre le risque. Tout le code VB6 qui **lisait** une propriété par défaut  
implicite (`s = txtNom`, `If txtA = txtB`, `v = rs("nom")`) change de sens — et la détection
**dépend du réglage `Option Strict`** :

- **Sous `Option Strict On`** : la plupart de ces lectures deviennent des **erreurs de compilation** (on ne peut pas affecter un `TextBox` à une `String`). Le problème est **visible** → détecté.
- **Sous `Option Strict Off`** (le mode dans lequel on **démarre** souvent, → 6.4) : le **late binding** tente d'interpréter l'expression à l'exécution. Selon les cas, cela **échoue au runtime** ou **se comporte différemment** — sans la moindre erreur de compilation. **Silencieux.**

> ⚠️ **Le piège du démarrage en `Option Strict Off`.** En tolérant le *late binding* pour faire
> passer le code hérité, on **masque** précisément les problèmes de propriétés par défaut que
> `Option Strict On` aurait signalés. Ils restent tapis jusqu'à ce qu'on active `Option Strict On`
> (→ 17.1) — ou, pire, jusqu'à ce qu'ils **se manifestent en production**. C'est l'illustration
> exacte du fil conducteur du chapitre : un changement que l'outillage **peut** rendre visible, mais
> que la configuration de départ **laisse silencieux**.

---

## 🪞 Le faux-ami : `Set` n'a pas totalement disparu

Attention à une source de confusion fréquente : le mot-clé **`Set` existe toujours en VB.NET** — mais
**ailleurs et pour autre chose**. On le retrouve dans la **définition des propriétés**, comme
en-tête de l'**accesseur d'écriture** :

```vb
' VB.NET — "Set" ici n'est PAS l'affectation d'objet de VB6 : c'est l'accesseur d'écriture
Public Property Nom As String
    Get
        Return _nom
    End Get
    Set(value)              ' bloc exécuté quand on AFFECTE la propriété
        _nom = value
    End Set
End Property
```

Ce `Set`-là est un **faux-ami** du `Set` de VB6 :

- le **`Set` de VB6** = affecter une **référence d'objet** (instruction) ;
- le **`Set` de VB.NET** = l'**accesseur d'écriture** d'une propriété (bloc de définition).

Plus subtil encore : en VB6, une propriété pouvait avoir **deux** accesseurs d'écriture distincts —
`Property Let` (pour une valeur) **et** `Property Set` (pour une référence). VB.NET les **fusionne en
un seul** accesseur `Set`. Cette fusion `Let` + `Set` → `Set` unique est un point de migration à  
part entière, **détaillé en section 10.6** (et c'est de là que vient le « faux-ami » signalé dans le  
SOMMAIRE).

> 💡 **À mémoriser sans se tromper** : `Set` **disparaît** en tant qu'**instruction d'affectation**,
> mais **réapparaît** en tant qu'**accesseur de propriété** — avec un sens **différent**. Voir un
> `Set` dans du code VB.NET ne veut donc *jamais* dire la même chose que dans du code VB6.

---

## 🛠️ Ce que font les outils

Les assistants de migration traitent ce sujet de façon largement **mécanique**, et plutôt bien :

- ils **explicitent** les propriétés par défaut → `.Text`, `.Value`, `Caption` devient `.Text`, etc. ;
- ils **suppriment** les `Set` d'affectation devenus inutiles → `Set a = b` devient `a = b` ;
- ils **convertissent** les blocs `Property Get`/`Let`/`Set` vers la syntaxe de propriété de VB.NET (avec la fusion `Let`+`Set`, → 10.6).

C'est l'un des points où l'outillage est **fiable**. Le risque résiduel n'est donc pas tant la  
conversion elle-même que les cas que l'outil **n'a pas pu trancher** et qui **survivent sous
`Option Strict Off`** — d'où l'importance d'activer `Option Strict On` ensuite (→ 17.1) pour les
**faire remonter**.

---

## ↔️ Cas connexes (renvois)

- **`Property Get`/`Let`/`Set` → propriétés VB.NET** (la fusion `Let`+`Set` et le faux-ami `Set`) : **section 10.6**.
- **`Option Strict` On/Off** (ce qui rend ces problèmes visibles ou silencieux) : **sections 6.4** (réglage de départ) et **17.1** (activation progressive).
- **Types valeur vs types référence** (ce qui remplace la distinction `Let`/`Set`) : **sections 2.1 et 2.4**.
- **`Nothing`** (l'affectation d'une référence « vide » remplace `Set obj = Nothing`) : **section 7.7**.

---

## 🧭 La règle de migration à retenir

| Écriture VB6 | Cible VB.NET |
|---|---|
| `Set a = obj` (affectation de référence) | `a = obj` (le `Set` disparaît) ✅ |
| `txtNom = "x"` (propriété par défaut implicite, écriture) | `txtNom.Text = "x"` (nommer la propriété) |
| `s = txtNom` (propriété par défaut implicite, lecture) | `s = txtNom.Text` (nommer la propriété) ⚠️ |
| `maCollection(i)` (propriété par défaut **paramétrée**) | `maCollection(i)` — **inchangé** (l'index lève l'ambiguïté) |
| `Property Get/Let/Set` (définition) | Propriété VB.NET avec `Get`/`Set` (Let+Set fusionnés, → 10.6) |

> ⭐ **Le principe directeur.** Partout où VB6 s'appuyait sur une propriété par défaut **implicite**,
> rendez-la **explicite** en la nommant. C'est non seulement nécessaire, mais **bénéfique** : le code
> devient lisible et compatible avec `Option Strict On`. Et rappelez-vous que **voir `Set` ≠
> comprendre `Set`** : son sens diffère entre les deux langages.

---

## ✅ À retenir

- **Propriétés par défaut et `Set` sont liés** : `Set` n'existait en VB6 que pour distinguer l'affectation d'une **référence** de celle de la **valeur** d'une propriété par défaut.
- VB.NET **supprime les propriétés par défaut *sans paramètre*** (on doit nommer `.Text`, `.Value`…) mais **conserve les *paramétrées*** (`collection(i)` reste valide).
- L'ambiguïté disparue, **`Set` (instruction d'affectation) est supprimé** : `Set a = b` devient `a = b`. Un `Set` oublié est une **erreur de compilation** → **visible**.
- Le **vrai risque silencieux** est la **lecture** d'une propriété par défaut : signalée sous **`Option Strict On`**, mais **masquée** sous **`Option Strict Off`** (le mode de démarrage fréquent → 6.4). D'où l'activation ultérieure d'`Option Strict On` (→ 17.1).
- **Faux-ami** : `Set` **survit** en VB.NET comme **accesseur d'écriture** d'une propriété — sens **différent** du `Set` de VB6. La fusion `Property Let` + `Property Set` → `Set` unique est traitée en **10.6**.
- Les outils gèrent ce sujet de façon **fiable et mécanique** ; le résiduel se traque via `Option Strict On`.

---

> 🧭 **Section suivante** → [2.6 Gestion d'erreurs : `On Error` vs exceptions structurées](06-erreurs-apercu.md)  
> 🔝 **Retour** → [Sommaire du chapitre 2](README.md)

⏭️ [Gestion d'erreurs : `On Error` vs exceptions structurées](/02-differences-fondamentales/06-erreurs-apercu.md)
