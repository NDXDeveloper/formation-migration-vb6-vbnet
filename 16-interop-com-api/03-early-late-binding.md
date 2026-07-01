🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.3 — *Early binding* vs *late binding* (`Option Strict Off`) : où c'est nécessaire, où l'éliminer

**Choisir, à chaque appel d'objet, entre un type connu à la compilation (liage anticipé) et une résolution à l'exécution (liage tardif) — et comprendre pourquoi l'objectif est de revenir au premier partout où c'est possible.**

> 🔗 Les ponts COM sont posés dans les deux sens (RCW au 16.1, CCW au 16.2). Reste une décision
> transversale, qui conditionne la sûreté de tout le code migré : **comment vos appels d'objets
> sont-ils résolus ?** Cette question est indissociable du commutateur **`Option Strict`** (introduit
> au module 6.4) et de l'objectif de validation du module 17.1.

---

## Deux façons de résoudre un appel

### Liage anticipé (*early binding*)

Le **type est connu à la compilation**. Le compilateur sait que `calc` est un `ITaxCalculator`, vérifie  
que `CalculerTVA` existe avec la bonne signature, et génère un **appel direct** (par la table des  
méthodes). Cela suppose une **référence** au type — au composant .NET, à l'assembly d'interop, ou à la  
bibliothèque de types COM.

### Liage tardif (*late binding*)

Le type est **`Object`**. Le compilateur ne sait rien de ses membres : la résolution de l'appel est
**reportée à l'exécution**. Pour un objet COM, cela passe par `IDispatch` (recherche du membre par son
nom, puis invocation) ; pour un objet .NET, par la **réflexion**. Souple, mais **non vérifié**.

| | **Liage anticipé** | **Liage tardif** |
|---|---|---|
| Résolution | À la **compilation** | À l'**exécution** |
| Vérification de type | ✅ Oui | ❌ Non |
| IntelliSense / découvrabilité | ✅ Oui | ❌ Non |
| Performance | Appel direct (rapide) | `IDispatch`/réflexion (plus lent) |
| Erreur de nom/signature | **Erreur de compilation** | **Plantage à l'exécution** |
| Prérequis | Une référence + un type | Rien (un simple `Object`) |
| `Option Strict` | Compatible **`On`** | Exige **`Off`** |

---

## L'héritage VB6 : du liage tardif partout (souvent par accident)

VB6 pratiquait massivement le liage tardif — et **fréquemment sans nécessité** :

- `Dim x As Object`, ou pire `Dim x` (un `Variant` non typé), produisaient du liage tardif.
- `CreateObject("…")` renvoie un `Object` → liage tardif systématique.
- La culture du **`Variant` omniprésent** et les **propriétés par défaut** (module 2.5) encourageaient
  un style « non typé » même quand le type réel était parfaitement connu.

> ⚠️ **Conséquence directe pour la migration.** Une conversion naïve **reconduit** tout ce liage tardif
> tel quel — et, avec lui, l'absence de vérification. C'est précisément le terreau des **changements
> silencieux** que cette formation combat (annexe B) : une faute de frappe ou un mauvais type d'argument
> qui ne se révèle qu'en production.

---

## Le commutateur central : `Option Strict`

Tout se joue sur une seule directive :

- **`Option Strict On`** : le liage tardif est **interdit** (appeler un membre sur un `Object` devient
  une **erreur de compilation**). En prime, les **conversions implicites avec perte** sont également
  interdites — il faut convertir explicitement (`CType`, `CInt`, `DirectCast`…).
- **`Option Strict Off`** : le liage tardif est **toléré**, ainsi que les conversions implicites. C'est
  l'état de départ recommandé pour absorber l'héritage VB6 (module 6.4 : « démarrer en `Option Strict
  Off` puis l'activer **progressivement** »).

> 🧨 **Le double danger d'`Option Strict Off`.** On le réduit souvent au liage tardif, mais il ouvre en
> réalité **deux** portes :
> 1. **Le liage tardif** — résolution non vérifiée, erreurs repoussées à l'exécution.
> 2. **Les conversions implicites avec perte** — un `Long` glissé dans un `Integer`, un `Double` dans un
>    `Single`… silencieusement (le terrain exact des entiers redimensionnés du module 7.2).
>
> Activer `Option Strict On` apporte donc un **bénéfice double** : il élimine le liage tardif **et**
> force les conversions explicites. C'est l'une des actions les plus rentables de toute la migration
> (pilotée globalement au module 17.1).

> ℹ️ À ne pas confondre avec `Option Explicit` (qui force la **déclaration** des variables) :
> ce sont deux réglages distincts, tous deux traités au module 6.4.

---

## Pourquoi le liage anticipé est presque toujours préférable

L'argumentaire en faveur de l'élimination du liage tardif :

- **Vérification à la compilation** — un membre mal orthographié ou une signature erronée est attrapé
  *avant* l'exécution, pas découvert par un utilisateur en production. C'est la meilleure défense
  contre les régressions silencieuses.
- **IntelliSense et découvrabilité** — vous voyez les membres disponibles ; productivité immédiate.
- **Performance** — un appel direct par la table des méthodes, au lieu d'une recherche `IDispatch`
  (`GetIDsOfNames` puis `Invoke`) ou par réflexion à chaque appel. Dans une boucle chaude, l'écart est
  net.
- **Refactoring sûr** — renommage et « rechercher les références » fonctionnent réellement.

> 🎯 **L'état cible est donc : liage anticipé + `Option Strict On`.** Le liage tardif n'est qu'une
> **exception** que l'on assume au cas par cas — jamais le régime par défaut.

---

## Où le liage tardif reste **légitime**

Soyons honnêtes : il existe de vrais cas où le liage tardif est **nécessaire**, ou au moins le moindre  
mal.

| Situation | Pourquoi le liage tardif |
|-----------|--------------------------|
| **Pas de bibliothèque de types** disponible (composant atteignable seulement par ProgID via `CreateObject`) | Aucun type à la compilation → liage tardif **forcé** (c'était le cas du FSO au module 16.1) |
| **Automatisation tolérante aux versions** (Office, typiquement) | Se lier à un PIA d'une version précise (Excel 2016…) casse l'app sur une machine équipée d'une autre version ; `CreateObject("Excel.Application")` reste **indépendant de la version** |
| **Dispatch réellement dynamique** (membre choisi à l'exécution, scénarios de script/plugin) | Le type n'est connu qu'au moment de l'appel |
| **API renvoyant un `Object`/`Variant` générique** | Le type concret n'est pas exprimable statiquement |
| **Transition de migration** | Maintenir, *temporairement*, un large corpus de code hérité qui compile (Option Strict `Off` transitoire, modules 6.4 / 17.1) |

Le cas de l'**automatisation Office** mérite d'être souligné : beaucoup d'équipes choisissent
**délibérément** le liage tardif pour ne pas dépendre d'une version d'Office installée. C'est un
arbitrage légitime — à condition de **confiner** ce liage tardif (voir ci-dessous).

---

## Où — et comment — l'éliminer

Partout ailleurs, l'objectif est de **revenir au liage anticipé**. Concrètement :

- **Donner le vrai type.** Là où vous disposez d'une bibliothèque de types ou d'un type .NET, **ajoutez
  la référence** et déclarez le type réel. Remplacez le style « tout `Object` » par les types concrets.
- **Pour vos propres composants exposés (CCW, module 16.2)** : vous **maîtrisez** la bibliothèque de
  types — fournissez donc toujours un `.tlb` et laissez VB6 se lier **en anticipé** à votre interface
  explicite (le mode `InterfaceIsDual` choisi au 16.2 le permet). Côté .NET **consommateur**, référencez
  l'assembly d'interop et utilisez le RCW typé (16.1) plutôt qu'un `Object`.
- **Activer `Option Strict On` progressivement.** C'est l'objet du module 17.1 : à chaque appel signalé
  par le compilateur, soit vous donnez un vrai type (préféré), soit — si le liage tardif est réellement
  nécessaire — vous l'**isolez**.

### Le levier du **niveau fichier**

`Option Strict On` peut être placé en tête d'un **fichier individuel**, et **prime sur le réglage du
projet** (qui peut rester `Off` pendant la migration). C'est ce qui rend l'activation **progressive**  
possible : on bascule un fichier à la fois en mode strict, sans devoir tout traiter d'un coup.

### Le patron d'**isolation** ⭐

Quand le liage tardif est **inévitable** (automatisation Office, composant sans `.tlb`), la bonne  
technique n'est pas de relâcher `Option Strict` sur tout le projet, mais de **confiner** ce liage  
tardif derrière une **interface stricte**, dans un **unique fichier** `Option Strict Off`. Le reste du  
code reste entièrement strict et anticipé.

```vbnet
' ── Fichier 1 (Option Strict On) : le CONTRAT, vu par tout le reste de l'application ──
Option Strict On

Public Interface IExporteurClasseur
    Sub ExporterVers(chemin As String, donnees As String(,))
End Interface
```

```vbnet
' ── Fichier 2 : le SEUL fichier en Option Strict Off — le liage tardif est confiné ici ──
Option Strict Off
Imports System.Runtime.InteropServices

Public NotInheritable Class ExporteurExcel
    Implements IExporteurClasseur

    Public Sub ExporterVers(chemin As String, donnees As String(,)) _
            Implements IExporteurClasseur.ExporterVers

        ' Liage tardif assumé : automatisation Excel tolérante aux versions
        Dim excel As Object = CreateObject("Excel.Application")
        Try
            ' ... pilotage d'Excel en liage tardif, isolé du reste du code ...
        Finally
            excel.Quit()
            Marshal.FinalReleaseComObject(excel)   ' durée de vie des objets COM (module 16.1)
            excel = Nothing
        End Try
    End Sub

End Class
```

Le reste de l'application ne connaît jamais que `IExporteurClasseur` — **strict, anticipé, vérifié**. Le  
liage tardif vit dans une seule cellule, petite et facile à tester. C'est ainsi qu'on **élimine le  
liage tardif du gros du code** même lorsqu'on ne peut pas l'éliminer totalement.

---

## Comparaison directe : tardif vs anticipé

```vbnet
' ❌ Liage tardif (exige Option Strict Off) — aucune vérification
Option Strict Off

Dim calc As Object = CreateObject("MaSociete.TaxCalculator")
Dim tva As Object = calc.CalculerTVA(1000D, 20D)   ' résolu à l'exécution via IDispatch
' Une faute de frappe (calc.CalculerTAV) COMPILE sans broncher → plantage en production
```

```vbnet
' ✅ Liage anticipé (compatible Option Strict On) — vérifié, rapide, découvrable
Option Strict On
Imports MaSociete                                  ' référence au composant / à son interop

Dim calc As ITaxCalculator = New TaxCalculator()
Dim tva As Decimal = calc.CalculerTVA(1000D, 20D)  ' vérifié par le compilateur, appel direct
' calc.CalculerTAV → ERREUR DE COMPILATION, attrapée immédiatement
```

Même composant (celui exposé au module 16.2), même opération : seul le **moment où l'erreur se révèle**  
change — et c'est tout ce qui compte pour la fiabilité.

---

## ⚠️ Pièges à connaître

- **Le liage tardif masque les erreurs jusqu'à l'exécution.** Un membre mal nommé ou un mauvais type
  d'argument se manifeste tardivement par un `MissingMemberException` ou une `COMException` peu
  parlante — souvent en production. C'est le cœur du danger des changements silencieux.
- **`Option Strict Off` n'est pas qu'une affaire de liage tardif.** Il autorise aussi les **conversions
  implicites avec perte** (module 7.2). Le laisser `Off` « juste pour le liage tardif » réintroduit
  *par la bande* la deuxième famille de bugs silencieux.
- **Coût en performance** dans les boucles : chaque appel tardif refait une résolution. À early-binder
  en priorité dans le code chaud.
- **Du liage tardif resté là « par oubli ».** Très souvent, un `Object` traîne simplement parce qu'on a
  oublié d'**ajouter la référence** au type. Le correctif est immédiat : référencer, puis typer.
- **Propriétés par défaut et liage tardif** (module 2.5) : en liage tardif, la résolution d'une
  propriété par défaut peut réserver des surprises ; le liage anticipé rend ces accès explicites et
  sûrs.

---

## En résumé

- **Liage anticipé** = type connu à la compilation → vérifié, rapide, découvrable. **Liage tardif** =
  résolution à l'exécution (`IDispatch` ou réflexion) → souple mais non vérifié, et **exige
  `Option Strict Off`**.
- VB6 pratiquait le liage tardif **massivement et souvent par accident** ; une migration naïve le
  reconduit, avec son absence de garde-fous.
- **`Option Strict On`** interdit le liage tardif **et** les conversions implicites : double bénéfice
  (le second rejoint le module 7.2). L'état cible est **liage anticipé + `Option Strict On`**.
- Le liage tardif reste **légitime** dans quelques cas : pas de bibliothèque de types, automatisation
  **tolérante aux versions** (Office), dispatch dynamique, ou transition de migration.
- Pour l'**éliminer** : donner les vrais types, fournir un `.tlb` pour vos propres composants (CCW),
  activer `Option Strict On` **au niveau fichier** progressivement (module 17.1), et — quand le liage
  tardif est inévitable — le **confiner** derrière une interface stricte dans un unique fichier
  `Option Strict Off` (patron d'isolation).

> ➡️ **Quittons les composants COM pour l'autre frontière : les appels directs à l'API Windows**, où
> l'instruction `Declare` de VB6 cède la place au **P/Invoke** (`DllImport`). (Module 16.4)

---

🏷️ **Indicateurs** : 🔗 Interop · lié au module 7.2 (conversions) et au module 17.1 (activation d'`Option Strict`)
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026

⏭️ [Déclarations d'API Win32 : `Declare` → `P/Invoke` (`DllImport`)](/16-interop-com-api/04-declare-vers-pinvoke.md)
