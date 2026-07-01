🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.1 — Réutiliser des composants COM/DLL existants depuis .NET (RCW)

**Comment appeler, depuis votre code .NET, un composant COM que vous n'avez pas (encore) migré : un OCX tiers, une DLL ActiveX, votre propre bibliothèque VB6 compilée — sans la réécrire.**

> 🔗 Premier des deux sens de l'interopérabilité COM (voir l'introduction du module). Ici, **du neuf
> consomme du vieux** : votre application .NET réutilise un composant COM existant. Le sens inverse —
> exposer du .NET à du code VB6 encore en place — est le **CCW**, traité au module 16.2.

---

## Le problème que résout le RCW

Vous migrez votre application par étapes. Tout ne peut pas (ni ne doit) basculer en même temps :

- Un **contrôle tiers** (grille, calendrier, composant graphique) dont vous n'avez pas l'équivalent
  .NET sous la main.
- Une **DLL ActiveX maison**, écrite en VB6, qui contient des règles métier précieuses et testées.
- Un **moteur externe** (Crystal Reports, un automate Office, un SDK fourni en COM).

Dans tous ces cas, la question est la même : *comment mon code .NET peut-il parler à un objet qui,  
lui, est resté du COM ?* La réponse est le **RCW**.

---

## Qu'est-ce qu'un RCW, concrètement ?

**RCW = *Runtime Callable Wrapper*** — littéralement « emballage appelable par le runtime ». Quand
votre code .NET instancie ou reçoit un objet COM, le **CLR fabrique automatiquement un objet managé  
intermédiaire** (le RCW) qui se place entre les deux mondes :

```
   Monde managé (.NET / CLR)                 Monde non managé (COM)
   ─────────────────────────                 ──────────────────────

   Votre code  ──appel .NET──►  [  RCW  ]  ──appel COM (IUnknown…)──►  Objet COM
                                  proxy                                 (DLL / OCX)
                ◄──résultat────             ◄────────résultat───────
```

Le RCW joue trois rôles :

1. **Il traduit les appels** : il convertit vos arguments .NET en types COM et inversement (c'est le
   *marshaling*, automatique pour les appels COM — détaillé au module 16.5).
2. **Il présente une façade .NET** : grâce à lui, l'objet COM ressemble à un objet .NET ordinaire
   (propriétés, méthodes, événements).
3. **Il gère le comptage de références COM** à votre place — et c'est précisément là que se cache le
   piège majeur de la section (voir plus bas).

> 💡 **Un RCW par identité COM.** Au sein d'un même *AppDomain*, le CLR garantit qu'un même objet COM
> est toujours enveloppé par **le même** RCW. Cette identité unique a une conséquence directe sur la
> libération : ce n'est pas l'objet COM que l'on relâche, c'est le **compteur interne du RCW**.

---

## Quand y recourir — et quand s'en abstenir

Le RCW est un outil **stratégique**, pas un réflexe. Avant de l'employer, posez-vous la question du  
remplacement pur et simple :

| Situation | Recommandation |
|-----------|----------------|
| Composant tiers complexe, **sans équivalent .NET** raisonnable | ✅ RCW — c'est exactement son rôle |
| **DLL VB6 métier** que vous migrerez plus tard | ✅ RCW — gagnez du temps, migrez-la après |
| Besoin **transitoire** pendant une migration incrémentale | ✅ RCW — pont assumé, à retirer ensuite |
| `Scripting.FileSystemObject`, objet `App`, registre… | ❌ Pas de RCW : utilisez `System.IO` / l'espace `My` (module 16.6) |
| Fonction déjà couverte par le framework (.NET natif) | ❌ Réécrivez : moins de dépendances, pas de COM à gérer |

> ⚖️ Rappel de la philosophie du chapitre : **l'interop est un pont, pas une destination.** Chaque RCW
> est une dépendance COM de plus à enregistrer, déployer et libérer correctement. On y recourt pour
> *réutiliser ce qui a de la valeur*, pas pour ré-emballer ce que .NET sait déjà faire nativement.

---

## Référencer un composant COM (liage anticipé)

Le cas le plus sûr et le plus confortable est le **liage anticipé** (*early binding*) : vous disposez  
de la **bibliothèque de types** du composant (intégrée à la DLL/OCX, ou dans un fichier `.tlb`), et  
Visual Studio vous génère des types .NET fortement typés.

**Dans Visual Studio 2026** : *Projet → Ajouter → Référence COM…*, puis cochez le composant dans
l'onglet **COM** (ou parcourez jusqu'au fichier). Cette opération de mise en place a déjà été décrite  
au **module 6.3** ; voici ce qui se passe en coulisses :

- Visual Studio appelle l'outil **`tlbimp.exe`** et génère un **assembly d'interopérabilité**, nommé
  par convention `Interop.{NomDeLaLib}.dll`. Cet assembly **ne contient pas** le composant : il
  contient seulement les **types managés** (les classes et interfaces RCW) qui en décrivent la forme.
  L'objet COM réel, lui, reste activé via COM à l'exécution (par son **CLSID**, donc le composant doit
  être **enregistré** — voir plus loin).

- Si un **PIA** (*Primary Interop Assembly* — assembly d'interop **officiel** fourni par l'éditeur,
  signé) est installé pour ce composant, Visual Studio le référence **au lieu** d'en générer un. C'est
  le cas, par exemple, des bibliothèques Microsoft Office.

- L'option de référence **« Incorporer les types d'interop »** (*Embed Interop Types*, propriété
  `Embed Interop Types = True`, disponible depuis .NET 4) permet d'**intégrer** les types d'interop
  nécessaires directement dans votre assembly : vous n'avez alors plus besoin de **distribuer** le
  fichier `Interop.*.dll` séparément. C'est souvent le réglage par défaut et le plus pratique.
  ⚠️ Quelques composants COM ne supportent pas l'incorporation (équivalence de type impossible) :
  pour ceux-là, repassez la propriété à `False` et livrez l'assembly d'interop (ou le PIA).

> 🛠️ **En ligne de commande**, l'équivalent direct est :
> `tlbimp MonComposant.dll /out:Interop.MonComposant.dll`
> Utile pour scripter une génération reproductible (intégration continue, jalons de migration).

---

## Premier exemple : appeler une DLL ActiveX VB6 depuis .NET

Supposons une ancienne DLL ActiveX VB6, `LegacyBilling`, exposant une classe `Calculator`. Après avoir  
ajouté la référence COM (donc avec `Interop.LegacyBilling.dll` disponible), on l'utilise comme un objet
.NET — à un détail crucial près, la **libération** (voir la section suivante) :

```vbnet
Imports System.Runtime.InteropServices
Imports LegacyBilling                          ' espace de noms issu de l'interop

Public Module FacturationLegacy

    Public Function CalculerTotal(lignes As IEnumerable(Of LigneFacture)) As Decimal
        Dim calc As Calculator = Nothing
        Try
            calc = New Calculator()            ' le CLR crée ici le RCW
            For Each ligne In lignes
                calc.AjouterLigne(ligne.Reference, ligne.Quantite, ligne.PrixUnitaire)
            Next
            Return CDec(calc.Total)            ' Currency (COM) → Decimal (.NET) — cf. module 7.2

        Catch ex As COMException
            ' Une erreur COM remonte sous forme de COMException, HRESULT inclus
            Throw New ApplicationException(
                $"Le composant de facturation a échoué (HRESULT 0x{ex.ErrorCode:X8}).", ex)

        Finally
            ' Libération déterministe de l'objet COM (essentiel — voir ci-dessous)
            If calc IsNot Nothing Then
                Marshal.FinalReleaseComObject(calc)
                calc = Nothing
            End If
        End Try
    End Function

End Module
```

Deux points méritent déjà l'attention : la conversion de type au retour (`Currency` COM devient
`Decimal` .NET, conformément au module 7.2), et le bloc `Finally` qui libère l'objet COM — le sujet
central de cette section.

---

## Le liage tardif, en l'absence de bibliothèque de types

Si vous n'avez **pas** de bibliothèque de types (ou si le ProgID n'est connu qu'à l'exécution), vous  
devez recourir au **liage tardif** (*late binding*), qui impose `Option Strict Off` :

```vbnet
' Option Strict Off requis pour appeler des membres sur un Object par liage tardif
Imports System.Runtime.InteropServices

Dim fso As Object = CreateObject("Scripting.FileSystemObject")
Try
    If fso.FileExists("C:\app\config.ini") Then
        Dim flux As Object = fso.OpenTextFile("C:\app\config.ini")   ' objet COM intermédiaire
        Dim contenu As String = flux.ReadAll()
        flux.Close()
        Marshal.ReleaseComObject(flux)                                ' on le libère aussi !
        ' ... exploiter "contenu" ...
    End If
Finally
    Marshal.FinalReleaseComObject(fso)
    fso = Nothing
End Try
```

> ❌ **Mais ne faites pas ça en pratique pour le FSO.** Cet exemple n'illustre que la *mécanique* du
> liage tardif. Conformément au **module 16.6**, le bon réflexe est de **remplacer** ce code par du
> `System.IO` natif — sans COM, sans RCW, sans libération à gérer :
> ```vbnet
> If IO.File.Exists("C:\app\config.ini") Then
>     Dim contenu As String = IO.File.ReadAllText("C:\app\config.ini")
> End If
> ```
> Le *late binding* d'interop est légitime quand vous **n'avez pas le choix** ; le module **16.3**
> détaille précisément où il reste nécessaire et où il faut l'éliminer pour retrouver la sécurité de
> type.

---

## ⚠️ Le piège n°1 : la durée de vie des objets COM

C'est **le** point à retenir de toute la section, et il prolonge directement le piège de la  
finalisation déterministe perdue (modules 2.2 et 12.3).

**En VB6**, un objet COM disparaissait de façon **déterministe** : dès que son compteur de références
tombait à zéro (variable mise à `Nothing` ou sortie de portée), il était détruit *immédiatement*, et
`Class_Terminate` s'exécutait sur-le-champ.

**En .NET**, l'objet COM n'est pas tenu par votre variable, mais par le **RCW**. Et le RCW est un
objet managé : il n'est libéré que lorsque le **ramasse-miettes** le collecte… **quand il le décide**.  
Tant que le RCW n'est pas finalisé, l'objet COM sous-jacent **reste vivant**.

Pour la plupart des composants en cours de processus, c'est sans conséquence visible. Mais dès qu'un  
objet COM tient une **ressource coûteuse** — un fichier verrouillé, une connexion, et surtout un
**processus externe** — le non-déterminisme devient un vrai bug. L'illustration la plus célèbre est
l'**automatisation d'Excel** : tant que tous les RCW associés ne sont pas libérés, le processus
`EXCEL.EXE` **reste actif en mémoire**, invisible et tenace, après la fin de votre traitement.

### La parade : libérer explicitement

Deux méthodes de l'objet `Marshal` (espace `System.Runtime.InteropServices`) :

| Méthode | Effet |
|---------|-------|
| `Marshal.ReleaseComObject(o)` | Décrémente de **1** le compteur interne du RCW et renvoie le compteur restant. À appeler autant de fois que l'objet a été « enveloppé ». |
| `Marshal.FinalReleaseComObject(o)` | Force le compteur à **0** et libère **immédiatement** l'objet COM sous-jacent. Pratique quand vous savez en avoir fini. |

Après libération, mettez toujours la variable à **`Nothing`** pour ne pas réutiliser un RCW relâché.

> 🧩 **Attention : un RCW n'est pas `IDisposable`.** Le réflexe `Using` acquis au module 12.3 ne
> s'applique donc **pas** directement à un objet COM brut. La discipline correcte est le couple
> `Try … Finally` + `Marshal.ReleaseComObject`, comme dans les exemples ci-dessus. (Vous pouvez, si
> vous le souhaitez, encapsuler un objet COM dans votre propre classe `IDisposable` qui appelle
> `ReleaseComObject` dans son `Dispose` — un patron utile pour les composants COM réutilisés souvent.)

### La règle du « point unique » (l'erreur du double point)

Conséquence directe de l'identité RCW : **chaque** objet COM intermédiaire que vous traversez reçoit  
son **propre** RCW. Si vous enchaînez les appels en chaîne, ces intermédiaires n'ont **aucune  
variable** qui les référence… et deviennent donc **impossibles à libérer**.

```vbnet
' ❌ À ÉVITER : "Documents", le document ouvert et "Paragraphs" créent
'    chacun un RCW orphelin que vous ne pourrez jamais relâcher → fuite
Dim nombre As Integer = app.Documents.Open(chemin).Paragraphs.Count
```

```vbnet
' ✅ Un objet COM = une variable. Chaque intermédiaire est libérable.
Dim documents As Object = app.Documents
Dim doc As Object = documents.Open(chemin)
Dim paragraphes As Object = doc.Paragraphs
Dim nombre As Integer = paragraphes.Count

Marshal.ReleaseComObject(paragraphes)
doc.Close()
Marshal.ReleaseComObject(doc)
Marshal.ReleaseComObject(documents)
```

Retenez la formule : **un objet COM, une variable, une libération.** C'est verbeux, mais c'est le prix  
d'un traitement propre — et c'est l'une des causes les plus fréquentes de processus fantômes.

### L'« option nucléaire » `GC.Collect` — à connaître, à manier avec lucidité

Quand on n'a pas le contrôle de tous les RCW (objets enfouis, code tiers), on voit souvent ceci :

```vbnet
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()                       ' parfois répété pour forcer la finalisation des RCW
```

Cela **force** le ramasse-miettes à finaliser les RCW en attente, donc à libérer les objets COM. C'est  
la rustine classique pour faire mourir un `EXCEL.EXE` récalcitrant. **Mais** c'est un **aveu de perte  
de contrôle** : on déclenche un cycle de GC global pour un problème local. À **préférer** : la  
discipline `Marshal.ReleaseComObject` + « un objet, une variable » ci-dessus. Réservez `GC.Collect` au  
dernier recours, quand vous ne pouvez vraiment pas atteindre tous les RCW.

---

## Les autres pièges à connaître

- **« Bitness » 32/64 bits.** ⚠️ Un composant COM/OCX **32 bits** ne peut être chargé que dans un
  processus **32 bits**. Si votre application est compilée en *AnyCPU* et tourne sur un Windows 64 bits,
  elle s'exécute en 64 bits et **échouera** à charger un composant 32 bits (souvent
  « *Classe non enregistrée* », HRESULT `0x80040154`, ou une `BadImageFormatException`). La parade :
  régler la **plateforme cible sur x86**. (Les serveurs COM **hors processus** — exécutables — peuvent
  contourner ce point via un substitut ; les DLL/OCX **en processus**, non : la « bitness » doit
  correspondre.)

- **Enregistrement et déploiement.** L'assembly d'interop ne fait que **décrire** le composant ;
  l'objet réel est activé via COM par son CLSID, donc le composant **doit être enregistré** sur la
  machine (par ex. `regsvr32 MonComposant.ocx`). Absent du registre → « *Classe non enregistrée* ».
  Ce qui « marche sur ma machine » peut donc échouer en production : à anticiper dès le **module 18.1**
  (packaging et prérequis).

- **Cloisonnement de *thread* (STA / MTA).** ⚠️ Beaucoup de composants COM — et toute l'UI Windows
  Forms — exigent un *thread* **STA**. C'est pourquoi le point d'entrée WinForms est marqué
  `<STAThread>`. Appeler un tel composant depuis un *thread* d'arrière-plan (MTA) peut provoquer des
  blocages ou des comportements erratiques. Si vous parallélisez un traitement qui touche du COM,
  vérifiez le modèle de cloisonnement attendu.

- **Les erreurs COM deviennent des exceptions.** Une défaillance côté COM remonte sous forme de
  **`COMException`**, dont la propriété `ErrorCode` porte le **HRESULT**. Le réflexe VB6 « tester
  `Err.Number` » se traduit donc par un `Try … Catch ex As COMException` (à relier au module 11 sur la
  gestion des erreurs).

- **Le marshaling des arguments est automatique… pour les appels COM.** Le RCW convertit seul vos
  types entrants/sortants. C'est l'appel d'**API Win32 brute** (P/Invoke) qui, lui, vous oblige à
  décrire le marshaling explicitement — sujet des modules 16.4 et **16.5**.

---

## En résumé

- Le **RCW** (*Runtime Callable Wrapper*) est le proxy managé que le CLR crée automatiquement pour que
  votre code .NET appelle un objet **COM** existant : **du neuf consomme du vieux**.
- Privilégiez le **liage anticipé** (référence COM → `Interop.*.dll` généré, ou **PIA**, avec
  l'option **Incorporer les types d'interop**) ; n'utilisez le **liage tardif** (`Option Strict Off`)
  que faute de bibliothèque de types (détaillé au module 16.3).
- **Le piège majeur** est la **durée de vie** : le RCW maintient l'objet COM vivant jusqu'à sa
  collecte par le GC — donc de façon **non déterministe**. Libérez explicitement avec
  `Marshal.ReleaseComObject` / `FinalReleaseComObject`, et appliquez la règle **« un objet COM, une
  variable, une libération »**.
- Un **RCW n'est pas `IDisposable`** : le `Using` du module 12.3 ne s'applique pas directement ;
  reposez-vous sur `Try … Finally`.
- Pièges associés : **bitness 32/64** (cibler x86 si besoin), **enregistrement/déploiement**,
  **cloisonnement STA/MTA**, et les erreurs qui deviennent des **`COMException`**.
- N'employez le RCW que pour **réutiliser ce qui a de la valeur** ; pour le FSO, l'objet `App` ou le
  registre, préférez les équivalents .NET natifs (module 16.6).

> ➡️ **Voyons maintenant le sens inverse — le pilier de la migration incrémentale :** exposer un
> composant .NET à du code VB6 encore en place, grâce au **CCW**. (Module 16.2)

---

🏷️ **Indicateurs** : 🔗 Interop · ⚠️ Piège (durée de vie des objets COM)
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026 · espace `System.Runtime.InteropServices`

⏭️ [Exposer du .NET à du VB6 (**CCW**) : la clé d'une migration **incrémentale**](/16-interop-com-api/02-ccw.md)
