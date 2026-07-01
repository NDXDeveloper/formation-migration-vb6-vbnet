🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.5 — Déclarations : `Dim` multiple, `DefInt`/`DefStr`…, portée, `As New`

> **Chapitre 7 — Types, variables et déclarations** · Partie 3 — Migrer le langage
>
> La syntaxe de déclaration semble identique d'un langage à l'autre. Pourtant, **quatre détails
> changent** — et deux d'entre eux sont de vrais **faux-amis** : la façon dont un `Dim` multiple
> **type** ses variables, et la façon dont `As New` **instancie** un objet. S'y ajoutent la
> **disparition** des instructions `Def*` et l'arrivée de la **portée de bloc**.

---

## ⚠️ `Dim` multiple : le faux-ami du typage

C'est le piège le plus connu de cette section. Sur une ligne déclarant plusieurs variables avec un  
seul `As Type` à la fin :

```vb
' VB6 — le type ne s'applique QU'À la dernière variable
Dim a, b, c As Integer    ' a et b sont des Variant ; seul c est Integer !

' VB.NET — le type s'applique à TOUTES les variables
Dim a, b, c As Integer    ' a, b ET c sont des Integer
```

La **même ligne** ne déclare donc **pas les mêmes types** selon le langage.

> 🪤 **Conséquence de migration.** Beaucoup de code VB6 écrivait `Dim i, j, k As Integer` en
> *croyant* tout typer — alors que `i` et `j` étaient en réalité des `Variant`. La migration vers
> VB.NET « corrige » silencieusement ce malentendu (tout devient `Integer`), ce qui est **en
> général souhaitable**… sauf si le caractère `Variant` de `i`/`j` était **exploité** quelque part
> (valeur non numérique, état `Empty`). À vérifier au cas par cas.

VB.NET ouvre par ailleurs des écritures plus sûres :

```vb
Dim x As Integer = 5                  ' déclaration + initialisation
Dim nom As String, age As Integer     ' types mixtes, chacun explicite
Dim u, v As New Client()              ' ⚠️ DEUX instances distinctes (un New par variable)
```

> 💡 Avec **`Option Infer On`** (par défaut dans un projet neuf), `Dim x = 5` **infère** le type.
> En migration, le projet démarre souvent en `Option Infer Off`, et des **types explicites** restent
> plus lisibles tant que tout n'est pas stabilisé.

---

## 🗑️ `DefInt`, `DefStr`… : des instructions **supprimées**

En VB6, les instructions `Def*` placées en tête de module fixaient le **type par défaut** des  
variables (et des paramètres/fonctions) dont le **nom commence** par certaines lettres, faute de  
type explicite :

```vb
' VB6 — toute variable non typée commençant par I..N devient Integer
DefInt I-N
' ...
k = 0        ' k est un Integer implicite, sans déclaration
```

La famille complète (`DefBool`, `DefByte`, `DefInt`, `DefLng`, `DefCur`, `DefSng`, `DefDbl`,
`DefDate`, `DefStr`, `DefObj`, `DefVar`) **n'existe plus en VB.NET**.

> 🎯 **Migration :** chaque variable autrefois gouvernée par un `Def*` doit recevoir un **type
> explicite**. C'est l'un des nettoyages à mener idéalement **en amont**, côté VB6 (ajouter
> `Option Explicit`, typer partout — voir **chapitre 5.1**), pour que la conversion produise un code
> sans ambiguïté.

```vb
' VB.NET — type explicite obligatoire
Dim k As Integer = 0
```

---

## 🌐 Portée et visibilité

Plusieurs règles de portée diffèrent. La plus structurante est nouvelle : la **portée de bloc**.

### Portée de **bloc** (nouveau) vs portée de **procédure**

En VB6, une variable déclarée n'importe où dans une procédure est visible dans **toute** la  
procédure. En VB.NET, elle n'est visible que dans le **bloc** (`If`, `For`, `While`…) où elle est  
déclarée.

```vb
' VB6 — portée PROCÉDURE
Sub Traiter()
    If condition Then
        Dim y As Integer
        y = 10
    End If
    Debug.Print y          ' ✅ visible en VB6
End Sub

' VB.NET — portée BLOC
Sub Traiter()
    If condition Then
        Dim y As Integer = 10
    End If
    ' Debug.Print(y)        ' ❌ erreur : y est hors de portée
End Sub
```

> 🪤 **Conséquence de migration :** si du code VB6 déclarait une variable **dans** un bloc et
> l'utilisait **au-delà**, il faut **remonter la déclaration** au niveau de la procédure. Les outils
> de migration le font généralement automatiquement (« hoisting »).

### Mots-clés de visibilité

| VB6 | VB.NET | Remarque |
|-----|--------|----------|
| `Dim` (au niveau module) | `Private` *(par défaut)* | Dans une **classe**, un champ `Dim` est **`Private`** ; dans une **structure**, il est **`Public`** |
| `Private` | `Private` | inchangé |
| `Public` | `Public` | inchangé |
| `Global` | *(supprimé)* | utiliser **`Public`** |
| `Static` (local persistant) | `Static` | conservé |
| `Friend` | `Friend` | portée **assembly** en .NET (≠ « projet ») — voir **12.6** ⚠️ |

> 💡 Préférer désormais des modificateurs **explicites** (`Private`/`Public`) plutôt que `Dim` au
> niveau d'une classe : l'intention de visibilité est plus claire. Les variables globales des
> modules `.bas` et leur portée sont traitées au **module 12.4**.

---

## ⚠️ `As New` : l'instanciation implicite change de sens

Deuxième faux-ami majeur de la section. `Dim x As New Classe` n'a **pas** le même comportement dans  
les deux langages.

```vb
' VB6 — instanciation PARESSEUSE + RE-création automatique
Dim c As New Collection
Set c = Nothing
c.Add "x"                ' c est SILENCIEUSEMENT recréé — aucune erreur
If c Is Nothing Then ... ' ⚠️ toujours False : évaluer c force son instanciation

' VB.NET — instanciation À LA DÉCLARATION, sans re-création
Dim c As New Collection
c = Nothing
c.Add("x")               ' ❌ NullReferenceException
If c Is Nothing Then ... ' ✅ test fiable (n'instancie rien)
```

Trois différences à mémoriser :

| | **VB6** | **VB.NET** |
|---|---|---|
| **Quand l'objet est créé** | au **premier usage** (paresseux) | à la **déclaration** |
| **Après mise à `Nothing`** | **recréé** automatiquement au prochain usage | **non recréé** → `NullReferenceException` |
| **Test `Is Nothing`** | non fiable (force la création, ~toujours False) | fiable |

> 🪤 **Conséquences de migration :**
> - Le code VB6 qui posait `Set x = Nothing` pour **« réinitialiser »** un objet puis le réutilisait
>   **casse** en .NET : il faut réinstancier **explicitement** (`x = New Classe()`).
> - À l'inverse, un `If x Is Nothing Then Set x = New …` qui, en VB6, **ne faisait rien** (le test
>   forçait déjà la création) devient en .NET une **vraie initialisation paresseuse** qui fonctionne
>   — la sémantique « s'inverse » dans le bon sens, mais il faut en être conscient.
>
> 🎯 **Recommandation :** remplacer `As New` par une **déclaration simple + `New` explicite** là où
> l'objet doit réellement être créé (dans le constructeur, ou au point d'usage). L'**instant de
> création** et la **durée de vie** deviennent visibles — ce qui prépare aussi la bonne gestion de
> `IDisposable`/`Using` (chapitres **12.2** et **12.3**).

> 💡 Rappel lié au `Dim` multiple : `Dim a, b As New Classe` crée **deux** instances distinctes
> (un `New` par variable), pas une seule partagée.

---

## ⚠️ Pièges silencieux — récapitulatif

- **`Dim` multiple** : en VB6 le type ne s'applique qu'à la **dernière** variable (les autres sont
  `Variant`) ; en VB.NET, **à toutes**.
- **`Def*` supprimés** : tout repose désormais sur des **types explicites**.
- **Portée de bloc** : une variable déclarée dans un `If`/`For` n'est plus visible **hors** du bloc.
- **`Dim` au niveau classe** = **`Private`** (et **`Public`** dans une structure) ; **`Global`**
  disparaît.
- **`As New`** : création **à la déclaration** (et non au premier usage) et **plus de re-création**
  après `Nothing` ; le test `Is Nothing` redevient **fiable**.

---

## 🧭 Stratégie de migration

1. **Lire chaque `Dim` multiple** à la lumière de la règle VB6 (seule la dernière variable était
   typée) pour confirmer l'intention réelle avant de tout typer.
2. **Éliminer les `Def*`** en déclarant explicitement les variables concernées — idéalement **en
   amont**, côté VB6 (chapitre 5.1).
3. **Remonter au niveau procédure** les déclarations utilisées au-delà de leur bloc d'origine.
4. **Expliciter la visibilité** (`Private`/`Public`) et remplacer `Global` par `Public`.
5. **Convertir `As New`** en déclaration + `New` explicite quand on veut maîtriser la création et la
   durée de vie (notamment pour les objets jetables).
6. **Filet de sécurité** : le ***golden master*** (chapitre 5.5) révèle un changement de
   comportement issu d'une variable autrefois `Variant` ou d'un objet autrefois auto-recréé.

---

## ✅ À retenir

- **`Dim a, b, c As Integer`** : `a`/`b` sont **`Variant`** en VB6, **`Integer`** en VB.NET — la
  même ligne, deux sens.
- Les instructions **`DefInt`/`DefStr`/…** **disparaissent** : place aux **types explicites**.
- La **portée de bloc** remplace la portée de procédure ; **remonter** les déclarations utilisées
  au-delà de leur bloc.
- **`Dim` au niveau classe = `Private`** ; **`Global` → `Public`** ; **`Static`** et **`Friend`**
  subsistent (`Friend` = portée **assembly**).
- **`As New`** : **création à la déclaration**, **pas de re-création** après `Nothing` — préférer
  l'**instanciation explicite**.

---

## 🔗 Pour aller plus loin

- **7.1 — `Variant` → `Object`** et **7.7 — `Nothing`, `Empty`, `Null`** : le sort des `Variant`
  implicites et de l'« objet nul ».
- **7.6 — Constantes, énumérations et booléen** : les autres formes de déclaration.
- **Chapitres 12.1–12.3** — classes, constructeurs (`Class_Initialize`) et `IDisposable`/`Using` :
  où la création explicite d'objets prend tout son sens.
- **Chapitre 12.4** (modules `.bas`, variables globales) et **12.6** (portées, `Friend` = assembly).
- **Chapitre 5.1** — nettoyer le code VB6 (`Option Explicit`, typage) **avant** de migrer.
- **Annexe D** — correspondance des types de données.

➡️ Section suivante : **[7.6 — Constantes, énumérations et booléen](06-constantes-enums-bool.md)**  
⬆️ Retour au **[sommaire du chapitre 7](README.md)**

⏭️ [Constantes, énumérations, et le booléen (`True` = -1 en VB6 ?)](/07-types-variables/06-constantes-enums-bool.md)
