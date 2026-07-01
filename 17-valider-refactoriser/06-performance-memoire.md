🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.6 — Revue de performance et de mémoire (`IDisposable`, `Using`, fuites de *handles*)

**Vérifier que le changement de modèle mémoire — du comptage de références déterministe de VB6 au ramasse-miettes de .NET — ne laisse rien fuir : fichiers, connexions, objets GDI, références COM, abonnements d'événements. Puis passer la performance en revue.**

> 🧭 La dernière section du module, et l'**aboutissement opérationnel** d'un fil rouge tenu depuis le
> début : la **finalisation déterministe perdue** (modules 2.2 → 12.2/12.3 → 16.1/16.6 → 17.4). Le
> module 17.4 en a traité la face **comportementale** (timing, ordre) ; ici on traite sa face
> **ressources** : les **fuites**. C'est aussi le dernier critère de sortie du module (« aucune fuite »)
> avant le déploiement (module 18).

---

## Pourquoi cette revue : le revers de la finalisation non déterministe

En VB6, une ressource était libérée **immédiatement** quand l'objet sortait de portée (compteur de  
références à 0, `Class_Terminate`). Les développeurs s'appuyaient sur ce comportement **sans y penser** :  
c'était toujours « maintenant ».

En .NET, le **ramasse-miettes** est **non déterministe**. La **mémoire managée** est récupérée  
automatiquement — mais les **ressources non managées** tenues par un objet (handles de fichiers, objets  
GDI, connexions, références COM, handles Windows) ne sont libérées qu'à la **finalisation** de l'objet…
« un jour », à la discrétion du GC.

> ⚠️ **Conséquence** : une application migrée **fidèlement** peut être *correcte* (le *golden master* est
> vert) et **fuir** néanmoins sous **usage prolongé** — les handles s'accumulent, le pool de connexions
> s'épuise, les objets GDI atteignent la limite Windows, un processus COM externe reste en vie. Ces
> problèmes n'apparaissent **pas** lors d'un test fonctionnel rapide : ils émergent **sous charge, dans
> la durée**. D'où la nécessité d'une revue **délibérée**.

---

## Mémoire managée vs ressources non managées

La distinction est cruciale :

- **Mémoire managée** — le GC s'en charge ; on ne « fuit » pas au sens C/C++. Mais on peut créer des
  **fuites logiques** : des objets **maintenus en vie** involontairement (références encore enracinées).
  Le cas n°1 : les **abonnements d'événements non désabonnés** (un éditeur de longue durée retient ses
  abonnés), suivi des **collections statiques** qui grossissent sans fin et des **caches sans éviction**.
- **Ressources non managées** — handles, objets GDI+, connexions, références COM : elles exigent une
  libération **déterministe** via `IDisposable`/`Using`. Le finaliseur du GC n'est qu'un **filet de
  rattrapage**, pas une stratégie.

---

## L'outil central : `IDisposable` / `Using`

Le patron `IDisposable` (module 12.3, déjà mobilisé aux 16.1/16.6/17.5) est l'outil de la libération  
déterministe. **`Using` garantit l'appel de `Dispose`** à la fin du bloc, **même en cas d'exception** —  
retrouvant le *timing* que VB6 offrait implicitement.

Cette revue consiste à **vérifier que chaque `IDisposable` est bien libéré**, de préférence par `Using` :

- Quand la durée de vie tient **dans une méthode** → `Using`.
- Quand la ressource est un **champ** (durée de vie plus longue) → la **classe propriétaire** doit
  elle-même être `IDisposable` et libérer ce champ dans son propre `Dispose` (chaîne de **propriété**
  des ressources).
- Pour vos classes qui enveloppent une ressource non managée → implémenter **correctement** le patron
  complet (`Dispose`, `Dispose(disposing)`, finaliseur de secours, `GC.SuppressFinalize`) — détail au
  module 12.3.

---

## Où fuient les applications VB6 migrées

La chasse, par catégorie (avec ce qu'il faut vérifier) :

### Connexions et commandes de base (module 15)

Les `Recordset` ADO de VB6 se fermaient par comptage de références ; en ADO.NET, `SqlConnection`,
`SqlCommand`, `SqlDataReader` **doivent** être libérés, sinon le **pool de connexions s'épuise** sous
charge.

```vbnet
' ✅ Using imbriqués : fermeture et retour au pool garantis, même en cas d'exception
Using cn As New SqlConnection(chaine)
    cn.Open()
    Using cmd As New SqlCommand(sql, cn)
        ' ... exécution ...
    End Using
End Using
```

### Objets GDI+ (module 14.1) — la fuite classique du portage

Le dessin VB6 (`Line`, `Circle`, `PSet`) ne montrait pas ce problème ; le code GDI+ migré qui crée des
`Pen`/`Brush` à chaque rendu **sans les libérer** épuise vite les **objets GDI** (limite Windows
d'environ 10 000 par processus).

```vbnet
' ❌ Fuite : stylo et pinceau créés à chaque rendu, jamais libérés
Protected Overrides Sub OnPaint(e As PaintEventArgs)
    Dim stylo As New Pen(Color.Blue, 2)
    Dim pinceau As New SolidBrush(Color.LightYellow)
    e.Graphics.FillRectangle(pinceau, ClientRectangle)
    e.Graphics.DrawRectangle(stylo, 0, 0, Width - 1, Height - 1)
End Sub

' ✅ Using (plusieurs ressources) : libération déterministe
Protected Overrides Sub OnPaint(e As PaintEventArgs)
    Using stylo As New Pen(Color.Blue, 2),
          pinceau As New SolidBrush(Color.LightYellow)
        e.Graphics.FillRectangle(pinceau, ClientRectangle)
        e.Graphics.DrawRectangle(stylo, 0, 0, Width - 1, Height - 1)
    End Using
End Sub
```

> 💡 **Ne libérez pas `e.Graphics`** : vous ne le **possédez** pas (il est fourni par l'événement). On ne
> libère que ce que l'on a créé.

### Flux de fichiers (module 16.6)

`StreamReader`/`StreamWriter`/`FileStream` → `Using`, sinon les **handles de fichiers fuient** et les
fichiers **restent verrouillés**.

### Références COM / RCW (module 16.1)

La discipline `Marshal.ReleaseComObject` / `FinalReleaseComObject` et la règle **« un objet COM, une  
variable, une libération »**. Symptôme typique : un **processus externe** (le fameux `EXCEL.EXE`) qui
**persiste**. C'est la fuite spécifique à l'interop.

### Abonnements d'événements (module 12.8) — la fuite **managée** n°1

Un `AddHandler` sans `RemoveHandler` : l'éditeur (de longue durée) **retient** l'abonné, qui ne peut plus
être ramassé.

```vbnet
' ❌ Abonnement sans désabonnement : le modèle (longue durée) retient le formulaire
AddHandler modele.DonneesModifiees, AddressOf RafraichirAffichage

' ✅ Se désabonner (par ex. à la fermeture du formulaire)
Private Sub Form_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
    RemoveHandler modele.DonneesModifiees, AddressOf RafraichirAffichage
End Sub
```

> 💡 Les clauses **`Handles`** (`WithEvents`) lient l'abonnement à la durée de vie de l'objet ; la fuite
> concerne surtout les **`AddHandler` manuels** vers des éditeurs de **longue durée**.

### Handles Win32 issus du P/Invoke (module 16.4)

Un handle obtenu via une API (`CreateXxx`…) doit être libéré par l'API correspondante
(`CloseHandle`/`DeleteObject`…) : un `IntPtr` n'est **pas** libéré automatiquement. Envisagez un
**`SafeHandle`** pour que le GC serve de filet.

---

## Détecter les fuites

La méthode : un **test d'endurance** (*soak test*) — faire tourner l'application sous charge **réaliste**,
**dans la durée**, en surveillant les compteurs de ressources. Un test rapide ne révèle pas une fuite
lente.

**La signature de chaque fuite** oriente le diagnostic :

| Indicateur qui **monte** sous charge constante | Type de fuite | Où regarder |
|----|----|----|
| Nombre de **handles** du processus | Handles non managés : fichiers, connexions, Win32 | 15 · 16.4 · 16.6 |
| Nombre d'**objets GDI** | Objets GDI+ non libérés (`Pen`, `Brush`, `Bitmap`, `Font`) | 14.1 |
| **Tas managé** / nombre d'instances d'un type donné | Fuite **logique** : abonnements d'événements, caches sans éviction | 12.8 |
| Un **processus externe** qui persiste | Référence **COM** non libérée (RCW) | 16.1 |

**Outils** : le **Gestionnaire des tâches** / **Process Explorer** (compteurs de **handles**, d'objets  
**GDI**/**USER**, mémoire) pour repérer une montée régulière ; **PerfMon** (compteurs *.NET CLR Memory*,
nombre de handles) ; un **profileur mémoire** (dotMemory, les *outils de diagnostic* de Visual Studio,  
ANTS) pour les fuites managées — prendre des **instantanés**, les **comparer**, et identifier les
**racines** qui retiennent les objets (souvent un gestionnaire d'événements).

---

## La revue de performance

L'autre moitié du titre. La migration peut introduire des régressions, et le modèle GC change les  
caractéristiques. À passer en revue :

- **Retirer les `GC.Collect` gratuits.** Appeler `GC.Collect()` en routine (l'« option nucléaire » du
  module 16.1) **nuit** à la performance et est presque toujours une erreur. Laissez le GC faire son
  travail ; supprimez les appels superflus.
- **Limiter les allocations** (*churn*). Les idiomes VB6 (concaténation par `&` en boucle, création
  d'objets dans une boucle serrée) engendrent une pression GC excessive : **`StringBuilder`** pour les
  constructions lourdes de chaînes, pas d'allocation par itération dans les chemins chauds (et attention
  au coût de LINQ en boucle chaude, module 17.5).
- **Éviter le *boxing*** : les collections non typées (`Object`, `Collection`) **emballent** les types
  valeur → les **génériques** (17.5 / 8.4) l'éliminent.
- **Éliminer le liage tardif résiduel** : chaque appel tardif (`IDispatch`/réflexion) est **lent** — le
  travail du module 17.1 améliore aussi la performance.
- **Réduire les franchissements d'interop** : le marshaling et les appels COM ont un coût ; minimisez les
  passages de frontière dans les chemins chauds (modules 16.5).

> 🎯 **Mesurer, ne pas deviner.** Profilez **avant** d'optimiser (profileur de performance / outils de
> diagnostic VS) : l'optimisation prématurée gaspille l'effort. Et — comme tout changement de code —
> **validez chaque optimisation au *golden master*** (17.3).

---

## Boucler la boucle de la finalisation

Cette section **clôt** un thème récurrent de toute la formation : la perte de la finalisation  
déterministe, abordée au module 2.2, outillée au 12.3 (`IDisposable`/`Using`), rencontrée à l'interop
(16.1/16.6), traitée sous l'angle comportemental au 17.4, et ici sous l'angle **ressources**. La
question finale — *« avons-nous réellement absorbé la principale différence structurelle entre VB6 et
.NET ? »* — trouve sa réponse opérationnelle dans cette revue.

---

## En résumé

- Le passage du **comptage de références** (VB6) au **GC** (.NET) impose une revue **dédiée** : une
  application correcte peut **fuir** sous usage prolongé — invisible aux tests rapides.
- **Mémoire managée** (gérée par le GC, mais fuites **logiques** : abonnements d'événements, caches) vs
  **ressources non managées** (handles, GDI+, connexions, COM) qui exigent une libération
  **déterministe**.
- **Outil central : `IDisposable`/`Using`** — généraliser `Using` (et chaîner la propriété des
  ressources via le `Dispose` des classes propriétaires).
- **Points de fuite** : connexions (15), **GDI+** (14.1, la fuite classique), flux de fichiers (16.6),
  **références COM** (16.1), **abonnements d'événements** (12.8, la fuite managée n°1), handles Win32
  (16.4).
- **Détecter** par *soak test* et **signature** (handles ↑, objets GDI ↑, tas managé ↑, processus
  persistant) ; outils : Process Explorer, PerfMon, profileur mémoire.
- **Performance** : retirer les `GC.Collect` gratuits, limiter allocations/*boxing*, éliminer le liage
  tardif et les franchissements d'interop — **en mesurant** et en validant au *golden master*.

---

## Le module 17 en perspective

La **porte** est franchie. L'application ne se contente plus de *compiler et tourner* : sa **correction  
est prouvée** (17.3), elle est **durcie** (`Option Strict` 17.1, `Compatibility` retiré 17.2), ses
**pièges silencieux** sont neutralisés (17.4), son code est **idiomatique** (17.5), et elle ne **fuit
pas** (17.6). Les **critères de sortie** définis dans l'introduction du module sont remplis.

> ➡️ **Cap sur la mise en production : packaging .NET Framework, cohabitation VB6/.NET pendant la
> transition, stratégie de bascule (pilote, *parallel run*) et plan de *rollback*.** (Module 18)

---

🏷️ **Indicateurs** : aboutissement du thème de la finalisation (2.2 → 12.3 → 16.1/16.6 → 17.4) ; satisfait le critère « aucune fuite » du module ; lié aux modules 14.1, 15, 16.4/16.5, 12.8 ; sous filet du module 17.3
**Cible** : .NET Framework 4.7.2 · Visual Studio 2026 · `System.IDisposable` · `Using`

⏭️ [Déploiement et bascule en production](/18-deploiement-bascule/README.md)
