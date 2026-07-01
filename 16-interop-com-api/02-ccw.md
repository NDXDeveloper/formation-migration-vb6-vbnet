🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.2 — Exposer du .NET à du VB6 (CCW) : la clé d'une migration incrémentale 🔗

**Comment rendre un composant .NET appelable depuis du code VB6 encore en place — pour réécrire l'application classe par classe, sans jamais tout casser d'un coup.**

> 🔗 Second sens de l'interopérabilité COM, et le plus stratégique. Ici, **du vieux consomme du
> neuf** : votre application VB6 *pas encore migrée* appelle un nouveau composant .NET, exactement
> comme elle appelait autrefois une DLL ActiveX VB6. C'est le miroir du RCW (module 16.1) — et c'est
> ce mécanisme qui transforme une migration *big-bang* à haut risque en une migration **progressive**
> et maîtrisée.

---

## Pourquoi le CCW est **le** levier de la migration incrémentale

Reprenons le choix stratégique du module 3.4 : *big-bang* (tout réécrire d'un coup) ou
**incrémentale** (remplacer morceau par morceau). L'approche incrémentale est presque toujours la plus
sage sur un *legacy* d'entreprise — mais elle suppose une chose : que l'ancien et le nouveau puissent
**cohabiter et communiquer** pendant toute la transition (modules 3.5 et 18.2).

Le CCW rend cette cohabitation possible **dans le sens qui compte le plus**. Le principe :

1. Vous isolez un ensemble cohérent (par exemple, le moteur de calcul de TVA).
2. Vous le **réécrivez en .NET**, proprement, avec toute la richesse du framework.
3. Vous l'**exposez au COM** (c'est l'objet de cette section).
4. Votre application VB6, toujours en production, l'appelle **comme si c'était l'ancienne DLL VB6** —
   souvent en gardant le **même ProgID**, si bien que le code appelant ne change presque pas.
5. Vous validez par rapport au *golden master* (modules 5.5 et 17.3), puis vous recommencez avec le
   bloc suivant.

> 🎯 **Le sens de la marche.** À chaque itération, la part .NET **grossit** et la part VB6 **maigrit**.
> L'application VB6 se vide progressivement de sa logique jusqu'à n'être plus qu'une coquille mince —
> que vous retirez enfin. C'est l'esprit du remplacement progressif : on **étrangle** doucement
> l'ancien plutôt que de tenter une réécriture totale d'un seul tenant. Et rien n'empêche de combiner
> les deux sens pendant la transition : du CCW pour que VB6 appelle vos nouveaux composants, du RCW
> (16.1) pour que vos composants .NET réutilisent encore quelques briques COM legacy.

---

## Qu'est-ce qu'un CCW, concrètement ?

**CCW = *COM Callable Wrapper*** — l'« emballage appelable par COM ». C'est l'exacte symétrie du RCW :
quand du code COM (votre VB6) appelle un objet .NET, le **CLR fabrique un proxy** qui présente votre  
objet managé **comme un véritable objet COM** (il implémente `IUnknown`, `IDispatch`, etc., et gère  
le comptage de références côté COM).

```
   Monde non managé (COM / VB6)              Monde managé (.NET / CLR)
   ────────────────────────────              ─────────────────────────

   Code VB6  ──appel COM──►  [  CCW  ]  ──appel .NET──►  Votre objet .NET
                              proxy                       (classe ComVisible)
             ◄──résultat────             ◄──────────────
```

Côté VB6, **rien ne trahit** que l'objet est en réalité du .NET : on l'instancie avec `New` ou
`CreateObject`, on appelle ses méthodes, on libère la référence avec `Set … = Nothing`. Tout le travail
consiste donc, côté .NET, à **rendre la classe correctement visible et stable** pour le monde COM.

---

## Rendre une classe .NET visible au COM

Tout repose sur quelques attributs de l'espace **`System.Runtime.InteropServices`**. Pris dans le bon  
ordre, ils n'ont rien d'intimidant.

### `ComVisible` — rendre le type accessible

Par défaut, le modèle de projet ajoute généralement `<Assembly: ComVisible(False)>` dans les  
informations d'assembly : **rien n'est exposé** tant que vous ne le demandez pas explicitement. Vous  
appliquez donc `<ComVisible(True)>` aux **types précis** que vous voulez offrir à VB6 — et à eux  
seuls. C'est une bonne discipline : la surface COM doit rester **petite et intentionnelle**.

### Des **GUID figés** — le point le plus critique ⚠️

Le monde COM identifie les types par des **GUID** (le `CLSID` pour une classe, l'`IID` pour une  
interface). Si vous **laissez le CLR les générer automatiquement**, ils **changent à chaque  
recompilation** — et tous les clients VB6 déjà enregistrés se retrouvent à pointer vers un CLSID qui  
n'existe plus. Régression garantie, et particulièrement vicieuse.

> ⚠️ **Règle d'or : attribuez vous-même un `<Guid("…")>` explicite à chaque classe et chaque interface
> exposée, et ne le régénérez JAMAIS.** Ce GUID est un contrat figé pour toute la durée de vie du
> composant.

### `ProgId` — un nom convivial pour VB6

`<ProgId("MaSociete.TaxCalculator")>` donne à VB6 un nom lisible et **stable** pour `CreateObject` et
`New`. Astuce de migration : si vous **réutilisez le ProgID de l'ancienne DLL VB6**, le code appelant
n'a parfois *aucune* ligne à changer.

### Le patron **« interface d'abord »** — la meilleure pratique ⭐

C'est **la** décision qui sépare un CCW robuste d'un CCW fragile. Si vous laissez .NET générer  
automatiquement une « interface de classe » (`ClassInterfaceType.AutoDual` ou `AutoDispatch`), la  
forme COM de votre objet (l'ordre des membres, leurs identifiants) **change dès que vous modifiez la  
classe** — ce qui casse silencieusement les clients VB6, surtout en liage anticipé. Microsoft
**déconseille** explicitement l'`AutoDual` pour cette raison.

La parade : **définir vous-même une interface explicite**, et faire de votre classe une simple  
implémentation derrière ce contrat stable.

- Sur l'**interface** : `<ComVisible(True)>`, un `<Guid>` figé, et
  `<InterfaceType(ComInterfaceType.InterfaceIsDual)>` (le mode *dual* autorise à la fois le liage
  anticipé **et** tardif — le plus confortable pour VB6).
- Sur la **classe** : `<ClassInterface(ClassInterfaceType.None)>` (on supprime l'interface
  auto-générée, fragile) et `<ComDefaultInterface(GetType(...))>` pour désigner l'interface par défaut.

Résultat : la table des méthodes vue par COM devient **stable**. Vous pouvez enrichir l'implémentation
.NET librement, tant que vous ne **modifiez pas** l'interface publiée (pour ajouter des membres, on
crée plutôt une **nouvelle** interface — `ITaxCalculator2` — afin de ne jamais casser l'existante).

---

## Enregistrer et déployer

L'assembly ne suffit pas : COM doit **savoir** qu'il existe. C'est le rôle de **`regasm`** :

```text
regasm MaSociete.Tax.dll /tlb:MaSociete.Tax.tlb /codebase
```

- `/tlb:` génère la **bibliothèque de types** (le `.tlb`), nécessaire au **liage anticipé** depuis VB6
  (on l'ajoutera dans *Projet ▸ Références*).
- `/codebase` inscrit le **chemin** de l'assembly dans le registre — indispensable quand il **n'est
  pas** dans le **GAC**. (regasm émet un avertissement si l'assembly n'est pas signé : pour de la
  production, la voie propre est de **signer** l'assembly et de l'installer dans le **GAC**, puis
  d'exécuter `regasm` *sans* `/codebase`.)

> 🛠️ **En développement**, l'option de projet *Propriétés ▸ Compilation ▸ « Inscrire pour
> l'interopérabilité COM »* exécute cet enregistrement **à chaque build** (pratique, mais requiert les
> droits administrateur).

> ⚠️ **Le piège de la « bitness » à l'enregistrement.** VB6 produit toujours des processus **32 bits**.
> Sur un Windows 64 bits, vous devez enregistrer avec le **`regasm` 32 bits** (celui du dossier
> `Framework`, **pas** `Framework64`) pour que l'inscription atterrisse dans la **vue de registre**
> que le processus VB6 32 bits consultera. Enregistrer côté 64 bits seulement = composant introuvable
> côté VB6.

---

## Exemple complet : un composant .NET appelé depuis VB6

### Côté .NET — le composant exposé

```vbnet
Imports System.Runtime.InteropServices

' ── Le CONTRAT exposé à COM : une interface explicite et STABLE ──────────────
<ComVisible(True)>
<Guid("B2C3D4E5-1111-2222-3333-444455556666")>     ' IID figé : ne JAMAIS le régénérer
<InterfaceType(ComInterfaceType.InterfaceIsDual)>   ' liage anticipé ET tardif (confort VB6)
Public Interface ITaxCalculator
    Function CalculerTVA(montantHT As Decimal, tauxPourcent As Decimal) As Decimal
    ReadOnly Property DerniereErreur As String
End Interface

' ── L'IMPLÉMENTATION .NET, vue par VB6 comme un objet COM ────────────────────
<ComVisible(True)>
<Guid("A1B2C3D4-9999-8888-7777-666655554444")>      ' CLSID figé
<ProgId("MaSociete.TaxCalculator")>                 ' nom convivial pour CreateObject / New
<ClassInterface(ClassInterfaceType.None)>           ' pas d'interface auto-générée (fragile)
<ComDefaultInterface(GetType(ITaxCalculator))>
Public Class TaxCalculator
    Implements ITaxCalculator

    ' Constructeur PUBLIC et SANS paramètre : indispensable pour New / CreateObject côté COM
    Public Sub New()
    End Sub

    Public Function CalculerTVA(montantHT As Decimal, tauxPourcent As Decimal) _
            As Decimal Implements ITaxCalculator.CalculerTVA
        If tauxPourcent < 0D Then
            ' Une exception .NET traversera la frontière sous forme de HRESULT (voir plus bas)
            Throw New ArgumentOutOfRangeException(NameOf(tauxPourcent),
                                                  "Le taux de TVA ne peut pas être négatif.")
        End If
        Return Decimal.Round(montantHT * tauxPourcent / 100D, 2)   ' Decimal : précision monétaire (7.2)
    End Function

    Private _derniereErreur As String = String.Empty
    Public ReadOnly Property DerniereErreur As String _
            Implements ITaxCalculator.DerniereErreur
        Get
            Return _derniereErreur
        End Get
    End Property

End Class
```

### Côté VB6 — la consommation

```vb
' ── Liage tardif : aucune référence à ajouter ────────────────────────────────
Dim calc As Object
Set calc = CreateObject("MaSociete.TaxCalculator")

Dim tva As Variant          ' VB6 n'a pas de type Decimal natif : on reçoit dans un Variant
tva = calc.CalculerTVA(1000, 20)        ' → 200,00

Set calc = Nothing          ' libère la référence COM (cf. durée de vie, module 16.1)
```

```vb
' ── Liage anticipé : Projet ▸ Références ▸ cocher la bibliothèque générée (.tlb) ─
Dim calc As New MaSociete.TaxCalculator
Dim tva As Variant
tva = calc.CalculerTVA(1000, 20)
Set calc = Nothing
```

> 🧩 **Détail de marshaling** : le `Decimal` .NET traverse la frontière comme un sous-type
> `Variant` (VT_DECIMAL) ; côté VB6 on le reçoit donc dans un `Variant`. C'est exactement le genre de
> subtilité que détaille le module 16.5. Et notez, au passage, le `Set` de VB6 pour l'affectation
> d'objet — un mot-clé disparu en VB.NET (rappel du module 2.5).

---

## Concevoir la frontière : ce que COM peut voir

Le CCW n'expose pas n'importe quoi. La surface COM impose des contraintes — qui doivent **guider votre  
conception** :

- Les types exposés doivent être **publics**, **non statiques** et posséder un **constructeur public
  sans paramètre** (sinon, ni `New` ni `CreateObject` côté VB6).
- Les **génériques** (`List(Of T)`, `Task(Of T)`…) ne sont **pas** visibles au COM. Les membres
  **`Shared`** non plus.
- Les **types des paramètres et des valeurs de retour** doivent eux-mêmes être compatibles COM
  (primitifs, chaînes, tableaux, ou autres interfaces COM-visibles).

> 💡 **Conséquence de conception : gardez la frontière simple et stable.** Exposez des signatures
> « amies de COM » (primitifs, chaînes, tableaux, interfaces explicites) ; cachez toute la richesse
> .NET (génériques, LINQ, async…) **derrière** cette façade. La frontière est un contrat figé ;
> l'implémentation, elle, reste libre d'évoluer.

---

## Le mode d'emploi d'une migration incrémentale par CCW

En réunissant tout ce qui précède, voici la **boucle** type — à répéter bloc après bloc :

1. **Choisir un module cohésif** à faible couplage UI (le découplage préparé au module 5.3 paie ici).
2. **Réécrire ce module en .NET**, idiomatique, derrière une **interface explicite** (`<Guid>` figé).
3. **Exposer et enregistrer** le composant (`ComVisible`, `regasm` 32 bits, `.tlb`), en **réutilisant
   le ProgID** de l'ancienne brique VB6 quand c'est possible.
4. **Repointer** le code VB6 vers le nouveau composant (souvent : presque aucune modification, grâce au
   ProgID conservé).
5. **Valider** par non-régression contre le *golden master* (modules 5.5 / 17.3).
6. **Recommencer** avec le bloc suivant — la part .NET grandit, la part VB6 se réduit.

Au terme du processus, il ne reste qu'une mince coquille VB6 (souvent l'UI résiduelle, traitée en  
partie 4) que l'on remplace en dernier. La migration n'a **jamais** exigé un saut total et risqué.

---

## ⚠️ Pièges à connaître

- **GUID auto-générés** → CLSID/IID instables → clients VB6 cassés à chaque build. **Figez** toujours
  les `<Guid>`. C'est le piège n°1 du CCW.
- **Interface de classe `AutoDual`** → forme COM fragile. **Définissez une interface explicite** et
  utilisez `ClassInterfaceType.None`. Pour faire évoluer un contrat, **ajoutez** une nouvelle interface
  plutôt que de modifier l'ancienne.
- **« Bitness » et enregistrement.** Hôte VB6 = 32 bits. Côté .NET, un serveur **en processus** chargé
  par cet hôte 32 bits s'exécute en 32 bits : `AnyCPU` comme `x86` conviennent (c'est l'image inversée
  de la contrainte du module 16.1, où c'était l'**hôte** .NET 64 bits qui ne pouvait pas charger un COM
  32 bits). En revanche, **enregistrez avec le `regasm` 32 bits** (voir plus haut), sinon VB6 ne trouve
  pas le composant.
- **Nettoyage déterministe depuis VB6.** Côté VB6, `Set … = Nothing` libère la **référence COM** ;
  l'objet .NET ne devient éligible au GC qu'ensuite — donc **sans finalisation déterministe** (le
  problème des modules 2.2 / 12.3, vu depuis l'autre rive). Si votre composant tient une ressource
  lourde, **n'attendez pas le GC** : exposez une méthode de nettoyage explicite (p. ex.
  `FermerConnexion()`) que le code VB6 appellera, puisqu'il ne peut pas invoquer `IDisposable.Dispose`
  de façon idiomatique.
- **Exceptions → HRESULT.** Une exception .NET lancée dans une méthode exposée traverse la frontière
  sous forme de **HRESULT en échec** ; côté VB6, elle est capturée par `On Error`, avec le message dans
  **`Err.Description`** et un code dans `Err.Number`. Lancez donc des exceptions **explicites et
  parlantes** côté .NET ; le code VB6 conserve son `On Error` habituel (à relier au module 11).
- **Événements .NET → VB6.** Pour qu'un formulaire VB6 capte des événements via `WithEvents`, le
  composant .NET doit publier une **interface source** (`<ComSourceInterfaces(...)>`). C'est plus
  avancé ; le modèle d'événements est traité au module 12.8.
- **Versionnement et ré-enregistrement.** Chaque déploiement réenregistre le composant ; privilégiez le
  **GAC + signature** en production, et gérez les versions avec soin (à coordonner avec le packaging du
  module 18.1).

---

## En résumé

- Le **CCW** (*COM Callable Wrapper*) est le proxy que le CLR crée pour qu'un objet .NET soit
  appelable **comme un objet COM** : **du vieux consomme du neuf**. C'est le symétrique du RCW.
- C'est **le levier de la migration incrémentale** : on réécrit en .NET, on expose via CCW (souvent
  avec le **même ProgID**), VB6 appelle le nouveau composant presque sans changement, et l'on
  **étrangle** l'ancien bloc après bloc.
- La mécanique tient en quelques attributs : `<ComVisible(True)>`, des **`<Guid>` figés** (le point le
  plus critique), un `<ProgId>`, et surtout le patron **« interface d'abord »**
  (`ClassInterfaceType.None` + interface explicite *dual*) pour une forme COM **stable**.
- On **enregistre** avec `regasm` (`/tlb`, `/codebase` ou GAC) — **en 32 bits** pour rester visible de
  l'hôte VB6.
- On **conçoit la frontière** simple et stable (pas de génériques ni de `Shared` ; signatures
  COM-compatibles), en cachant la richesse .NET derrière.
- Pièges : GUID instables, interface `AutoDual` fragile, enregistrement 32 bits, nettoyage non
  déterministe (prévoir une méthode explicite), exceptions converties en HRESULT, événements via
  interface source.

> ➡️ **Les ponts sont posés dans les deux sens (RCW et CCW). Reste à choisir, à chaque appel, entre
> liage anticipé et liage tardif** — où le second est nécessaire, et où il faut l'éliminer pour
> retrouver la sécurité de type. (Module 16.3)

---

🏷️ **Indicateurs** : 🔗 Interop · ⭐ Point clé (migration incrémentale)
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026 · `System.Runtime.InteropServices` · `regasm`

⏭️ [*Early binding* vs *late binding* (`Option Strict Off`) : où c'est nécessaire, où l'éliminer](/16-interop-com-api/03-early-late-binding.md)
