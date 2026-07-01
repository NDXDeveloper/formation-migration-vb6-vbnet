🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.6 — `Property Get`/`Let`/`Set` → propriétés (le faux-ami du `Set`) ⚠️

> **Trois procédures VB6 fusionnent en un seul bloc.** En VB.NET, `Property Let` (écriture d'une **valeur**) **et** `Property Set` (écriture d'un **objet**) se rejoignent dans **l'unique** accesseur `Set`. Et comme le mot `Set` **survit** avec un sens différent, c'est un **faux-ami** redoutable.

📍 *Module 10 · § 10.6 · [↑ Introduction du chapitre](README.md) · [← § 10.5](05-surcharge-retour.md)*

---

## ⚡ En bref

- VB6 implémente une propriété avec **jusqu'à trois procédures** : `Property Get` (lecture), `Property Let` (écriture d'une **valeur**), `Property Set` (écriture d'un **objet**).
- VB.NET implémente une propriété avec **une seule déclaration** dotée de deux accesseurs : `Get` et `Set`. **`Let` et `Set` fusionnent** dans ce **`Set` unique**, qui gère **toutes** les écritures (valeur **comme** objet).
- **Le faux-ami ⚠️** : le mot `Set` existe dans les deux langages mais ne désigne pas la même chose. En VB6, `Set` est aussi une **instruction** d'affectation d'objet (`Set x = obj`) — **supprimée** en VB.NET (**[§ 2.5](../02-differences-fondamentales/05-default-properties-set.md)**).
- **Le cas piégeux** : une propriété VB6 possédant **à la fois** un `Let` et un `Set` **n'a aucune traduction directe** et exige un **refactoring**.

---

## 1. Deux modèles de propriété

| Aspect | VB6 | VB.NET (4.7.2) |
|---|---|---|
| Structure | **3 procédures séparées** (`Get` / `Let` / `Set`) | **1 déclaration**, accesseurs `Get` / `Set` |
| Écriture d'une **valeur** | `Property Let` | accesseur `Set` |
| Écriture d'un **objet** | `Property Set` (+ mot-clé `Set` côté appelant) | accesseur `Set` (sans mot-clé) |
| Lecture seule | `Get` seul | `ReadOnly Property` |
| Écriture seule | `Let` **ou** `Set` seul | `WriteOnly Property` |
| Type | un par procédure (à accorder) | **un seul** type pour la propriété |

---

## 2. Le rôle des trois procédures VB6

En VB6, la procédure d'écriture dépend de **ce qu'on assigne** : une valeur ou une référence d'objet.

```vb
' VB6 — propriété VALEUR : Get + Let
Private m_Nom As String

Property Get Nom() As String
    Nom = m_Nom
End Property
Property Let Nom(ByVal valeur As String)      ' écriture d'une VALEUR
    m_Nom = valeur
End Property
```

```vb
' VB6 — propriété OBJET : Get + Set
Private m_Connexion As Object

Property Get Connexion() As Object
    Set Connexion = m_Connexion               ' lecture d'un objet → mot-clé Set
End Property
Property Set Connexion(ByVal obj As Object)   ' écriture d'un OBJET
    Set m_Connexion = obj
End Property
```

Côté appelant, VB6 **exige** le mot-clé `Set` pour l'affectation d'objet :

```vb
obj.Nom = "Dupont"               ' → Property Let
Set obj.Connexion = maConnexion  ' → Property Set (mot-clé Set obligatoire)
```

---

## 3. ⚠️ Le faux-ami : « Set » ne désigne pas la même chose

C'est le cœur de cette section. Le mot `Set` existe dans les deux langages, mais avec des sens **différents** — d'où des confusions à la lecture comme à l'écriture.

| « Set » | VB6 | VB.NET |
|---|---|---|
| **Instruction** `Set x = obj` | Affectation d'un **objet** à une variable | **Supprimée** → on écrit `x = obj` (**[§ 2.5](../02-differences-fondamentales/05-default-properties-set.md)**) |
| **`Property Set`** | Procédure d'écriture pour les **objets** seulement | N'existe plus en tant que telle |
| **Accesseur `Set`** | — | **L'unique** accesseur d'écriture (valeur **et** objet) = fusion `Let` + `Set` |

Autrement dit : en VB6, `Set` = « j'affecte un **objet** » ; en VB.NET, `Set` = « **le** bloc d'écriture de la propriété, quel que soit le type ». Un développeur VB6 qui lit `Set` en VB.NET pourrait croire à une écriture **réservée aux objets** — c'est faux.

Voici les mêmes propriétés en VB.NET. **Un seul** `Set` dans chaque cas :

```vb
' VB.NET — propriété valeur : Get + Set
Private m_Nom As String
Public Property Nom() As String
    Get
        Return m_Nom
    End Get
    Set(ByVal value As String)
        m_Nom = value
    End Set
End Property

' VB.NET — propriété objet : LE MÊME mécanisme Get + Set
Private m_Connexion As Object
Public Property Connexion() As Object
    Get
        Return m_Connexion
    End Get
    Set(ByVal value As Object)
        m_Connexion = value
    End Set
End Property
```

Et côté appelant, le mot-clé `Set` **disparaît** :

```vb
obj.Nom = "Dupont"               ' accesseur Set
obj.Connexion = maConnexion      ' accesseur Set — PLUS de mot-clé Set
```

---

## 4. Les cas de migration

La plupart des propriétés se migrent **mécaniquement**, selon le tableau suivant.

| VB6 | VB.NET | Remarque |
|---|---|---|
| `Get` + `Let` | `Property` avec `Get` / `Set` | Propriété **valeur** (lecture/écriture) |
| `Get` + `Set` | `Property` avec `Get` / `Set` | Propriété **objet** ; **retirer le `Set`** au site d'appel |
| `Get` seul | `ReadOnly Property` (`Get`) | Lecture seule |
| `Let` seul | `WriteOnly Property` (`Set`) | Écriture seule (valeur) |
| `Set` seul | `WriteOnly Property` (`Set`) | Écriture seule (objet) |
| `Get` + `Let` + `Set` | ⚠️ **Aucune traduction directe** | Refactoring requis (§ 5) |

> 💡 Pour les propriétés **objet** (`Get` + `Set`), n'oubliez pas de **convertir les sites d'appel** : tout `Set obj.Prop = …` devient `obj.Prop = …`.

---

## 5. ⚠️ Le cas qui ne se traduit pas : `Let` *et* `Set` sur la même propriété

VB6 autorise — c'est rare, mais légal — une propriété (typiquement `Variant`) dotée **à la fois** d'un `Property Let` **et** d'un `Property Set`, afin de réagir **différemment** selon qu'on lui affecte une valeur ou un objet :

```vb
' VB6 — la propriété distingue valeur et objet via Let ET Set
Property Let Donnee(ByVal v As Variant)    ' obj.Donnee = 42        → traitement "valeur"
    ' ...
End Property
Property Set Donnee(ByVal o As Object)     ' Set obj.Donnee = autre → traitement "objet"
    ' ...
End Property
```

VB.NET n'ayant **qu'un** accesseur `Set`, **impossible** d'aiguiller l'écriture vers deux traitements distincts selon le type. Ce cas **doit être refactorisé**. Trois pistes :

- **Option A — un `Set` unique typé `Object`.** La propriété devient `As Object` et le `Set` unique teste éventuellement le type à l'exécution (`If TypeOf value Is …`). On **perd la distinction syntaxique** valeur/objet et l'on retombe sur du *late binding* → voir **[§ 7.1](../07-types-variables/01-variant-object.md)**.
- **Option B — deux membres distincts.** Séparer en deux propriétés nommées clairement (`DonneeValeur` / `DonneeObjet`), ce qui rend l'intention explicite.
- **Option C — des méthodes surchargées.** Remplacer la propriété par des méthodes `DefinirDonnee(...)` **surchargées** (**[§ 10.5](05-surcharge-retour.md)**), qui distinguent naturellement les types par leurs signatures.

> ⚠️ C'est le **seul** cas de migration de propriété qui demande une véritable décision de conception. Repérez-le en cherchant les propriétés possédant **à la fois** un `Let` et un `Set`.

---

## 6. Détails de signature

- **La valeur écrite.** En VB6, la valeur assignée est le **dernier paramètre** du `Let`/`Set` (`Property Let Nom(ByVal valeur As String)`). En VB.NET, elle devient le **paramètre implicite `value`** de l'accesseur `Set` (`Set(ByVal value As String)`).
- **Un seul type.** En VB.NET, le type de retour du `Get` et le type du `value` du `Set` doivent être **identiques** : c'est l'unique type de la propriété. (En VB6, c'étaient des déclarations séparées qu'il fallait accorder à la main.)
- **Propriétés paramétrées (indexées).** Les paramètres d'index passent **sur la déclaration** de la propriété ; `value` reste implicite dans le `Set` :

  ```vb
  ' VB.NET — propriété paramétrée
  Property Element(ByVal i As Integer) As String
      Get
          Return m_Tableau(i)
      End Get
      Set(ByVal value As String)
          m_Tableau(i) = value
      End Set
  End Property
  ```

- **Propriétés par défaut.** VB.NET **restreint** les propriétés par défaut à celles qui prennent **au moins un paramètre** (indexées). Les propriétés par défaut **sans paramètre** très courantes en VB6 (sur les contrôles, etc.) ne survivent **pas** comme « défaut » et doivent être **qualifiées explicitement** → voir **[§ 2.5](../02-differences-fondamentales/05-default-properties-set.md)** et le **[module 12](../12-poo/01-classes.md)**.

---

## 7. Une modernisation possible : les propriétés auto-implémentées

Quand l'accès ne contient **aucune logique** (ni validation, ni effet de bord), VB.NET permet une **propriété auto-implémentée** : le compilateur génère le champ privé de stockage.

```vb
' VB.NET — propriété auto-implémentée (Get/Let triviaux en VB6)
Public Property Nom As String

' avec valeur initiale
Public Property Statut As String = "Actif"
```

> 💡 Réservez-la aux propriétés **triviales**. Si le `Property Let` VB6 d'origine **validait** la valeur ou déclenchait un **effet de bord**, conservez un accesseur `Set` **explicite** pour préserver ce comportement.

---

## 8. Migration : outils et travail manuel

- L'**[*Upgrade Wizard*](../04-outils-migration/01-upgrade-wizard.md)** **fusionne** `Get`/`Let`/`Set` en propriétés VB.NET et **convertit les sites d'appel** (`Set x = obj` → `x = obj`). Les cas `Get`+`Let` et `Get`+`Set` se convertissent **proprement**.
- Le cas **`Let` + `Set`** sur une même propriété est celui qui **résiste** à la conversion automatique : aucun outil ne peut reproduire deux traitements d'écriture dans un accesseur unique. Il **doit** passer en **revue manuelle** (§ 5).
- **Vérification** : l'écriture d'une propriété pouvant cacher une **logique** (validation, événement, finalisation indirecte), confirmez par les tests que la migration n'a pas modifié les **effets de bord** — cohérent avec la démarche des **[§ 10.2](02-byref-byval-audit.md)** et du golden master.

> 🔗 Les propriétés reviennent dans leur contexte complet — classes, encapsulation, événements — au **[module 12 (Programmation orientée objet)](../12-poo/01-classes.md)**. La présente section ne traite que leur **syntaxe de migration**.

---

## 📋 Tableau de synthèse

| Aspect | VB6 | VB.NET (4.7.2) | Nature du changement |
|---|---|---|---|
| Implémentation | 3 procédures `Get`/`Let`/`Set` | 1 propriété, accesseurs `Get`/`Set` | Structurel |
| Écriture valeur **vs** objet | `Let` **vs** `Set` (distincts) | **Un seul** `Set` (fusion) | ⚠️ Fusion |
| Mot-clé `Set` (instruction) | Affectation d'objet `Set x = obj` | **Supprimé** (`x = obj`) | ⚠️ Faux-ami (→ § 2.5) |
| Valeur écrite | Dernier paramètre du `Let`/`Set` | Paramètre implicite `value` du `Set` | Mécanique |
| Type | Un par procédure | **Un seul** pour la propriété | Mécanique |
| `Let` **et** `Set` sur une propriété | Légal (rare) | **Aucune traduction directe** | ⚠️ Refactoring |
| Propriétés par défaut | Souvent sans paramètre | **Doivent** être paramétrées | → § 2.5 / module 12 |
| Auto-implémentation | — | Disponible (accès triviaux) | ✅ Opportunité |

---

## ✅ À retenir

1. **Trois procédures → un bloc** : `Property Get`/`Let`/`Set` deviennent une propriété VB.NET avec `Get`/`Set`. **`Let` et `Set` fusionnent** dans **l'unique** `Set`.
2. **Le faux-ami du `Set`** : le mot survit, le sens change. L'**instruction** `Set x = obj` **disparaît** (on écrit `x = obj`) ; l'**accesseur** `Set` n'est **pas** réservé aux objets — il gère **toutes** les écritures.
3. **Convertissez les sites d'appel** des propriétés objet : `Set obj.Prop = …` → `obj.Prop = …`.
4. **Le seul cas dur** : une propriété ayant **à la fois** un `Let` et un `Set` n'a **pas** de traduction directe → refactoring (objet unique, deux membres, ou méthodes surchargées).
5. **Modernisez avec prudence** : propriété auto-implémentée pour les accès **triviaux** seulement ; gardez un `Set` explicite si le `Let` VB6 validait ou avait des effets de bord.

---

## 🏁 Fin du module 10

Vous avez parcouru l'ensemble des **procédures, fonctions et paramètres** : appels et `Call` (§ 10.1), le piège-vedette **ByRef → ByVal** (§ 10.2), `Optional` et la fin de `IsMissing` (§ 10.3), `ParamArray` (§ 10.4), surcharge et valeurs de retour (§ 10.5), et les propriétés (§ 10.6).

*Chapitre suivant : **[11 — Gestion des erreurs : de `On Error` aux exceptions](../11-gestion-erreurs/README.md)** ⭐*
*Pour aller plus loin sur les propriétés et les classes : **[module 12 — Programmation orientée objet](../12-poo/README.md)** ⭐ ⚠️*

**Cible** : .NET Framework 4.7.2 · **Indicateurs** : ⭐ point clé · ⚠️ piège · 🛠️ outils

⏭️ [Gestion des erreurs — de `On Error` aux exceptions](/11-gestion-erreurs/README.md)
