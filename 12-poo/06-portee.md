🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.6 — `Public` / `Private` / `Friend` : `Friend` = portée **assembly** en .NET ⚠️

> **Chapitre 12 — Programmation orientée objet** · Section 12.6
> Les modificateurs d'accès se ressemblent… mais **`Friend` a changé de frontière** : du **projet** (VB6) à l'**assembly** (.NET).

---

## 🧭 Périmètre

Trois mots-clés (`Public`, `Private`, `Friend`) existent dans les deux langages, et la traduction **paraît** triviale. Le piège est ailleurs : **ce que `Friend` protège n'est plus la même chose**. À cela s'ajoutent des modificateurs **nouveaux** (`Protected` et ses variantes), apparus avec l'héritage.

---

## 1. Les modificateurs, côte à côte

| Modificateur | VB6 | VB.NET |
|--------------|-----|--------|
| **`Public`** | accessible partout (et exposé aux clients COM externes) | accessible partout, **toutes assemblies** confondues |
| **`Private`** | accessible **dans le module** déclarant | accessible **dans le type** déclarant |
| **`Friend`** | accessible **dans le projet**, **invisible** aux clients COM externes | accessible **dans l'assembly**, invisible aux **autres assemblies** |
| **`Protected`** | *(inexistant — pas d'héritage)* | accessible dans la classe **et ses dérivées** (toutes assemblies) |
| **`Protected Friend`** | *(inexistant)* | dérivées **OU** même assembly (union) |
| **`Private Protected`** | *(inexistant)* | dérivées **ET** même assembly (intersection) |

---

## 2. Le cœur du piège : `Friend`, de « projet » à « assembly » ⚠️

En **VB6**, `Friend` signifie : *« utilisable partout dans mon **projet**, mais **invisible** aux clients (COM) qui consomment mon composant »*. C'est le moyen de partager des membres **entre classes d'un même projet** sans les exposer au monde extérieur.

En **VB.NET**, `Friend` signifie : *« utilisable partout dans mon **assembly** (le `.dll`/`.exe` compilé), mais **invisible** aux **autres assemblies** »*.

### 2.1 Le cas courant : projet = assembly → migration **propre**

La plupart des projets VB6 compilent **un** composant ; la plupart des projets VB.NET compilent **un** assembly. Quand **projet ≈ assembly**, les deux notions de `Friend` **coïncident** : l'assistant de mise à niveau garde `Friend` tel quel, et le comportement est **identique**.

```vb
' VB6 — Compte.cls (projet "MaBanque")
Friend Sub AppliquerInteret(ByVal taux As Double)   ' visible dans le projet,
    ' …                                              ' invisible aux clients COM
End Sub
```

```vb
' VB.NET — Compte.vb (assembly "MaBanque.dll")
Public Class Compte
    Friend Sub AppliquerInteret(taux As Double)     ' visible dans MaBanque.dll,
        ' …                                          ' invisible aux autres assemblies
    End Sub
End Class
```

Tant que l'appelant est **dans le même assembly**, l'accès fonctionne — exactement comme en VB6.

### 2.2 Le piège : quand la **structure de la solution** change

La frontière de `Friend` étant désormais l'**assembly**, tout **redécoupage** la déplace :

- **Scinder** un projet VB6 en **plusieurs** assemblies .NET : un membre `Friend` (jadis visible dans tout le projet) devient **inaccessible** entre deux assemblies différentes → **erreur de compilation**.
- **Fusionner** plusieurs projets VB6 en **un** assembly : des membres autrefois `Public` *entre projets* deviennent de simples membres **internes** à l'assembly — leur visibilité **se réduit** de fait.

```vb
' Si Compte est dans MaBanque.Core.dll
' et que l'appelant est dans MaBanque.Services.dll :
compte.AppliquerInteret(0.02)   ' ❌ ne compile plus : Friend ne traverse pas l'assembly
```

> ⚠️ **Le mot ne change pas, sa portée si.** `Friend` se migre à l'identique **sans alerte**, mais sa frontière est passée du **projet** (notion d'organisation du code + exposition COM) à l'**assembly** (frontière de **déploiement** et de **sécurité**). Si votre **architecture de solution** diffère de l'organisation VB6 d'origine, **auditez chaque `Friend`** aux frontières entre assemblies.

---

## 3. `InternalsVisibleTo` : rouvrir `Friend` entre assemblies

Quand un membre `Friend` doit rester accessible **depuis un autre assemblie précis** (sans le rendre `Public`), .NET fournit l'attribut d'assembly **`InternalsVisibleTo`** : l'assembly qui **possède** les membres `Friend` **autorise** nommément un autre assembly à les voir.

```vb
' VB.NET — dans MaBanque.Core (ex. fichier AssemblyInfo.vb)
<Assembly: InternalsVisibleTo("MaBanque.Services")>
```

Deux usages typiques :

- **Solutions multi-assemblies** issues d'un découpage : exposer les internes à un assembly partenaire.
- **Tests unitaires** : donner à l'assembly de tests l'accès aux membres `Friend` à vérifier (`InternalsVisibleTo("MaBanque.Tests")`).

> 💡 Pour un assembly **à nom fort** (*strong-named*), `InternalsVisibleTo` doit inclure la **clé publique** de l'assembly cible. À garder en tête si votre projet est signé.

---

## 4. `Protected` & variantes : nouveaux **parce que l'héritage est nouveau**

VB6 n'avait pas d'héritage : il n'avait donc **pas** de `Protected`. VB.NET introduit :

- **`Protected`** — accessible dans la classe **et ses classes dérivées** ;
- **`Protected Friend`** — dérivées **ou** même assembly (union) ;
- **`Private Protected`** — dérivées **et** même assembly (intersection), disponible avec le compilateur VB moderne (rarement nécessaire en migration).

Comme l'héritage (voir **12.5**), on **n'introduit pas** `Protected` pendant la migration — mais on le **rencontre** dès qu'on dérive d'une classe du **framework**. C'est précisément pourquoi le finaliseur s'écrit `Protected Overrides Sub Finalize` (voir **12.3**) : on **redéfinit** un membre `Protected` de `Object`.

---

## 5. Les valeurs par défaut (à connaître — mais restez explicite)

Omettre le modificateur n'a pas toujours le même effet, et **un défaut surprend** :

> ⚠️ **Un type sans modificateur est `Friend`, pas `Public`.** En VB.NET, une `Class`, un `Module` ou une `Structure` déclaré **sans** modificateur au niveau de l'espace de noms est **`Friend`** (interne à l'assembly) — **pas** `Public`. Si une classe doit être utilisable depuis un **autre assembly**, marquez-la **explicitement `Public`**.

Pour les **membres**, les défauts sont **cohérents** entre VB6 et VB.NET, ce qui limite le risque :

- un champ déclaré avec **`Dim`** dans une classe est **`Private`** (dans les deux langages) ;
- une **procédure** sans modificateur dans une classe est **`Public`** (dans les deux langages).

> ✅ Recommandation : **soyez explicite** sur chaque modificateur. L'assistant de mise à niveau rend d'ailleurs la plupart des accès explicites — c'est une bonne chose, conservez-la.

Note connexe : la **manière dont une classe VB6 était exposée** via COM (sa propriété **`Instancing`** : `Private`, `PublicNotCreatable`, `MultiUse`…) se traduit en .NET par une combinaison de **modificateurs d'accès** et, si l'on continue d'exposer du COM, d'**attributs COM** (voir **chapitre 16**).

---

## 6. Conseils de migration

- **Projet = assembly ?** Gardez `Friend` tel quel : le comportement est identique (le défaut de l'assistant convient).
- **Découpage en plusieurs assemblies ?** **Auditez les `Friend`** aux frontières : promouvez en `Public` ce qui doit traverser, ou ouvrez ciblé avec `InternalsVisibleTo`.
- **N'introduisez pas `Protected`** (ni d'héritage) pendant la migration ; laissez cela au refactoring (**17.5**).
- **Marquez `Public`** explicitement les types réellement destinés à être consommés hors de l'assembly (rappel : le défaut est `Friend`).

---

## 🔁 Récapitulatif VB6 → VB.NET

| VB6 | VB.NET |
|-----|--------|
| `Public` (exposé COM) | `Public` (toutes assemblies) |
| `Private` (module) | `Private` (type) |
| `Friend` (**projet**) | `Friend` (**assembly**) — frontière différente ⚠️ |
| *(pas de `Protected`)* | `Protected`, `Protected Friend`, `Private Protected` |
| accès `Friend` entre deux parties d'un projet | même assembly, **ou** `InternalsVisibleTo` entre assemblies |
| propriété `Instancing` (COM) | modificateurs d'accès **+** attributs COM si exposition (16) |
| type sans modificateur | **`Friend`** par défaut (pas `Public`) ⚠️ |

---

## ✅ Points clés

- `Public` / `Private` se transposent simplement ; **`Friend` aussi en apparence** — mais sa frontière passe du **projet** à l'**assembly**.
- Tant que **projet = assembly**, `Friend` se comporte **à l'identique** ; le risque n'apparaît qu'en **redécoupant** la solution.
- En multi-assemblies, **auditez les `Friend`** : `Public` ou **`InternalsVisibleTo`** selon le besoin (utile aussi pour les **tests**).
- ⚠️ Un **type sans modificateur** est **`Friend`** par défaut en VB.NET, **pas `Public`**.
- `Protected` (et variantes) sont **nouveaux** : on ne les introduit pas en migration, mais ils **surgissent** via l'héritage du framework (ex. `Protected Overrides Sub Finalize`).

---

## 🔗 Renvois

- **12.1** — structure de classe. · **12.4** — `Public` dans un `Module` (global) vs dans une classe (membre).
- **12.5** — héritage (d'où viennent `Protected` & co.). · **12.3** — `Protected Overrides Sub Finalize`.
- **16** — interop COM et exposition (attributs COM, `Instancing`).
- **17.5** — refactoring idiomatique.
- **Suite → [12.7 — `Type…End Type` → `Structure` (sémantique de valeur, attention au comportement)](07-type-vers-structure.md)** ⚠️

⏭️ [`Type…End Type` → `Structure` (sémantique de valeur, attention au comportement)](/12-poo/07-type-vers-structure.md)
