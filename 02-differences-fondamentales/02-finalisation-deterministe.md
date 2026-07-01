🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2.2 La finalisation déterministe perdue — le piège n°1 ⭐ ⚠️

> **Chapitre 2 — Différences fondamentales VB6 ↔ VB.NET**
> S'il ne fallait comprendre **qu'une seule** différence avant de migrer, ce serait celle-ci.

---

## Le piège qui mérite son numéro 1

La section 2.1 a établi que VB6 et VB.NET reposent sur deux moteurs de mémoire opposés : le
**comptage de références** (hérité de COM) d'un côté, le **ramasse-miettes** (GC) de l'autre. Cette
section examine **la conséquence la plus structurante** de cette différence — celle qui casse le  
plus de code, et le casse **en silence**.

Pourquoi « le piège n°1 » ? Parce qu'il réunit les trois caractéristiques du danger maximal :

- il touche **énormément** de code VB6 (toute libération de ressource y est concernée) ;
- il est **parfaitement silencieux** : le code converti **compile**, et souvent **fonctionne en test** ;
- il se manifeste **plus tard**, en production, sous une forme déroutante (connexions épuisées, fichiers verrouillés) **sans lien apparent** avec la cause.

> ⚠️ **Le cœur du problème en une phrase.** En VB6, vous saviez **exactement quand** un objet était
> détruit et **quand** son code de nettoyage s'exécutait. En VB.NET, **vous ne le savez plus** — et
> tout le code qui comptait sur ce « quand » change de comportement.

---

## 🔵 Côté VB6 : comptage de références et destruction déterministe

En VB6, les objets sont des **objets COM**, et COM gère leur durée de vie par **comptage de  
références**. Le principe est simple et **prévisible** :

- chaque fois qu'une variable référence l'objet, un **compteur** s'incrémente ;
- chaque fois qu'une référence disparaît, le compteur **se décrémente** ;
- dès que le compteur atteint **zéro**, l'objet est détruit **immédiatement** — et **`Class_Terminate` s'exécute à cet instant précis**.

C'est ce qu'on appelle la **finalisation déterministe** : la destruction survient à un **moment  
connu, prévisible, contrôlé par votre code**.

Deux événements font tomber le compteur à zéro :

```vb
' VB6 — une classe qui détient une ressource (.cls)
Private mConn As ADODB.Connection

Private Sub Class_Initialize()
    Set mConn = New ADODB.Connection
    mConn.Open "..."            ' ressource acquise à la naissance
End Sub

Private Sub Class_Terminate()
    mConn.Close                 ' ressource libérée à la destruction…
    Set mConn = Nothing
End Sub
```

```vb
Sub Traiter()
    Dim svc As CService
    Set svc = New CService       ' Class_Initialize → connexion ouverte
    svc.FaireQuelqueChose
    Set svc = Nothing            ' compteur = 0 → Class_Terminate s'exécute ICI, maintenant
End Sub                          ' (et si on avait omis "= Nothing", à la fin du Sub : même effet)
```

**C'est sur ce comportement que des millions de lignes de VB6 ont été écrites.** Le développeur VB6
plaçait son nettoyage dans `Class_Terminate` en *sachant* qu'il s'exécuterait au bon moment :  
fermeture de fichiers, libération de verrous, **fermeture de connexions à une base**, relâchement  
d'objets COM, libération de *handles*. Le réflexe `Set obj = Nothing` faisait partie de la culture  
du langage, précisément parce qu'il **déclenchait** cette libération.

> 💡 En somme, VB6 offrait gratuitement un mécanisme proche du *RAII* : « l'objet disparaît → la
> ressource est libérée, **tout de suite** ». C'est confortable — et c'est exactement ce que
> .NET **ne** fait **pas**.

---

## 🟢 Côté .NET : le ramasse-miettes décide *quand*

En VB.NET, la mémoire des objets managés est gérée par le **ramasse-miettes** (GC). Et le GC  
fonctionne sur un principe **radicalement différent** :

- **il n'y a pas de comptage de références** pour les objets managés ;
- le GC repère les objets **devenus inaccessibles** (plus aucune référence ne permet de les atteindre) ;
- mais il ne les collecte **pas immédiatement** : il se déclenche **quand il le juge utile**, principalement sous la **pression mémoire** (quand les allocations s'accumulent) — **pas** quand vous cessez d'utiliser un objet.

Autrement dit : **le moment de la destruction n'est plus déterminé par votre code, mais par le GC.**  
C'est la **finalisation non déterministe**.

L'équivalent approximatif de `Class_Terminate` existe — c'est la méthode **`Finalize`** (le  
destructeur). Mais elle hérite de ce non-déterminisme :

```vb
' VB.NET — conversion NAÏVE (le piège) : Class_Terminate → Finalize
Public Class Service
    Private conn As SqlConnection

    Public Sub New()
        conn = New SqlConnection("...")
        conn.Open()
    End Sub

    Protected Overrides Sub Finalize()
        conn.Close()        ' ⚠️ s'exécute quand le GC le décide — PAS à la fin de Traiter()
        MyBase.Finalize()
    End Sub
End Class
```

```vb
Sub Traiter()
    Dim svc As New Service       ' connexion ouverte
    svc.FaireQuelqueChose()
    svc = Nothing                ' ⚠️ ne ferme RIEN ici ; déréférence seulement
End Sub                          ' la connexion reste OUVERTE après la fin du Sub —
                                 ' jusqu'à ce que le GC se déclenche, un jour
```

Trois pièges concentrés dans cet exemple, à bien intérioriser :

- **`Finalize` ne s'exécute pas à la fin du `Sub`.** Elle s'exécutera à un **moment indéterminé**, sur un **thread de finalisation** distinct, potentiellement **bien plus tard** — voire, dans des cas extrêmes (arrêt brutal du processus), **pas du tout**.
- **`svc = Nothing` ne détruit rien.** Le réflexe VB6 est ici **trompeur** : en .NET, mettre une variable à `Nothing` ne fait que retirer une référence ; **le GC reste seul maître** du moment de la collecte. Mettre à `Nothing` ne « ferme » aucune ressource.
- **Bonus coûteux** : un objet doté d'un `Finalize` **vit plus longtemps** (il faut deux cycles de GC pour le récupérer) et **pèse** sur le ramasse-miettes. Mettre du nettoyage dans `Finalize` est donc non seulement tardif, mais **contre-performant**.

> 💡 Le contraste tient en un mot. VB6 : **vous** décidez quand l'objet meurt (comptage de
> références). VB.NET : **le GC** décide (ramasse-miettes). Tout le malentendu de migration naît de
> ce transfert de décision.

---

## ⚠️ Pourquoi c'est *silencieux* — et donc dangereux

Reprenons le fil conducteur du chapitre : ce changement ne déclenche **aucune erreur de  
compilation**. La conversion naïve `Class_Terminate` → `Finalize` **compile parfaitement**, et le  
code de nettoyage **est toujours là** — il s'exécute simplement **au mauvais moment**.

Le résultat, c'est une famille de bugs où la **ressource n'est pas libérée à temps** :

| Ressource non libérée à temps | Symptôme observé en production |
|---|---|
| **Connexions à une base** | **Épuisement du pool** de connexions, erreurs de *timeout* / « trop de connexions » **sous charge** |
| **Fichiers** | Fichier **verrouillé** (« in use »), impossible à rouvrir, déplacer ou supprimer |
| **Verrous / sections critiques** | Blocages, contention, comportements erratiques |
| **Objets COM** (interop) | Référence COM non relâchée → processus serveur (Excel, Word…) qui **reste en mémoire** |
| **Handles système** | **Fuite de *handles***, dégradation progressive, épuisement à long terme |

Le piège est d'autant plus sournois qu'il présente une **asymétrie test/production** redoutable :

> ⚠️ **« Ça marchait pourtant en test… »** En développement, sous **faible charge**, le GC se
> déclenche assez souvent — et les connexions sont assez peu nombreuses — pour que **rien ne se
> voie**. En **production**, sous charge réelle, le **pool s'épuise bien avant** que le GC ne libère
> les objets : l'application tombe **là où aucun test ne l'avait montrée**. C'est la signature d'un
> changement silencieux.

---

## ⏱️ Les deux chronologies, côte à côte

```
   VB6 — comptage de références              VB.NET — ramasse-miettes
   ─────────────────────────────            ──────────────────────────────
   New CService   → conn ouverte            New Service    → conn ouverte
     refs = 1                                 (référencé)
   ...                                       ...
   Set svc = Nothing → refs = 0             svc = Nothing  → déréférencé
     └─► Class_Terminate ICI                  (rien ne se passe)
         conn fermée ✔ (déterministe)         │
                                              ▼  ... délai indéterminé ...
                                             le GC se déclenche un jour
                                               └─► Finalize
                                                   conn fermée ✖ (trop tard)
```

À gauche, la libération est **ancrée à un point précis** du déroulement. À droite, elle **flotte**,
détachée de votre logique, à la merci du GC. Toute la migration consiste à **réancrer** cette  
libération — ce que permet la solution ci-dessous.

---

## ✅ La parade : `IDisposable` et `Using`

.NET a parfaitement conscience du problème — au point d'avoir **conçu un mécanisme dédié** pour le
résoudre. C'est l'interface **`IDisposable`** et l'instruction **`Using`**, qui **rétablissent la  
libération déterministe** que VB6 offrait nativement.

**Le principe.** Une classe qui détient une ressource implémente `IDisposable` et expose une méthode  
**`Dispose()`** : c'est **le** point de libération **explicite et déterministe**. On y met le
nettoyage qui était dans `Class_Terminate`.

```vb
' VB.NET — la solution : IDisposable + Using
Public Class Service
    Implements IDisposable

    Private conn As SqlConnection

    Public Sub New()
        conn = New SqlConnection("...")
        conn.Open()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        conn.Close()             ' libération DÉTERMINISTE, déclenchée explicitement
        conn.Dispose()
    End Sub
End Class
```

**Le bloc `Using`** garantit alors l'appel de `Dispose()` **à la fin du bloc**, de façon **certaine**
— y compris si une **exception** survient (sous le capot, c'est un `Try…Finally`) :

```vb
Sub Traiter()
    Using svc As New Service     ' connexion ouverte
        svc.FaireQuelqueChose()
    End Using                    ' ✅ Dispose() appelé ICI, garanti — même en cas d'exception
End Sub
```

Comparez avec la chronologie précédente : le `End Using` joue **exactement le rôle** que jouait la  
chute du compteur de références à zéro en VB6. **La libération est de nouveau ancrée à un point  
précis** — déterministe.

On peut englober **plusieurs ressources** dans un même `Using` (libérées dans l'ordre inverse) :

```vb
Using cn As New SqlConnection("..."), cmd As New SqlCommand("...", cn)
    ' ... utilisation ...
End Using                        ' cmd puis cn libérés, automatiquement
```

> 💡 **Bonne nouvelle : la moitié du travail est déjà faite.** Tous les types « à ressource » de
> .NET (`SqlConnection`, `StreamReader`, `Font`, `Graphics`…) implémentent **déjà** `IDisposable`. Il
> suffit de les **consommer correctement**, c'est-à-dire dans un `Using`. Le réflexe à acquérir est
> simple : **« un objet à ressource ⇒ un `Using` »**.

> ⚠️ **À noter** : le `Dispose()` montré ici est la version **simplifiée** (le chemin déterministe
> seul). Le **motif `Dispose` complet** — avec `Finalize` en **filet de sécurité** pour les
> ressources **non managées**, la méthode protégée `Dispose(disposing)` et `GC.SuppressFinalize` —
> est détaillé en **section 12.3**. Pour le cas courant (envelopper un objet .NET déjà
> `IDisposable`), la version simplifiée suffit largement.

---

## 🧭 La règle de migration à retenir

Face à un `Class_Terminate` en VB6, posez-vous **une seule question** : *« libère-t-il une  
ressource ? »*.

| Situation VB6 | La bonne cible .NET |
|---|---|
| `Class_Terminate` qui **libère une ressource** (fichier, connexion, verrou…) | **`IDisposable` + `Using`** (déterministe) ✅ |
| `Class_Terminate` **vide** ou purement mémoire (rien à libérer) | **Rien à faire** — laisser le GC gérer la mémoire |
| `Set obj = Nothing` utilisé pour **« détruire »** l'objet | **Inutile** pour la mémoire ; utiliser **`Using`** pour libérer une ressource |
| Détention **directe** d'une ressource **non managée** (handle brut, COM bas niveau) | **Motif `Dispose` complet** avec `Finalize` en filet (**→ 12.3**) |

> ⭐ **Le principe directeur.** Ne traduisez **jamais** mécaniquement `Class_Terminate` → `Finalize`.
> La cible naturelle d'un `Class_Terminate` qui libère une ressource est **`IDisposable` + `Using`**,
> **pas** `Finalize`. `Finalize` n'est qu'un **filet de sécurité** de dernier recours, non un
> remplacement.

---

## 💡 Les réflexes VB6 à désapprendre

Trois automatismes hérités de VB6 deviennent **faux** — ou inutiles — en VB.NET :

- **« Je mets `Set obj = Nothing` pour libérer. »** → En .NET, `obj = Nothing` ne libère **rien** de déterministe. Pour une ressource : **`Using`**. Pour la mémoire seule : **ne rien faire**.
- **« Je mets le nettoyage dans le destructeur. »** → En .NET, le destructeur (`Finalize`) est **tardif et non déterministe**. Le nettoyage déterministe va dans **`Dispose()`**, déclenché par **`Using`**.
- **« La ressource se ferme toute seule quand l'objet sort de portée. »** → Vrai en VB6, **faux** en .NET. Rien ne se ferme « tout seul au bon moment » : c'est **`Using`** qui rétablit cette garantie.

---

## ✅ À retenir

- **VB6** détruit les objets par **comptage de références** : dès que la dernière référence disparaît, l'objet meurt **immédiatement** et `Class_Terminate` s'exécute **à cet instant** → **finalisation déterministe**.
- **VB.NET** confie la mémoire au **ramasse-miettes**, qui décide **quand** collecter (sous pression mémoire) → **finalisation non déterministe** : `Finalize` s'exécute **plus tard**, à un moment imprévisible.
- C'est **le piège n°1** parce qu'il est **silencieux** : le code **compile**, **passe les tests**, puis **échoue en production** (pool épuisé, fichiers verrouillés, *handles* fuités) — l'asymétrie test/production est sa signature.
- **`obj = Nothing` ne libère rien** de déterministe en .NET, et **mettre le nettoyage dans `Finalize` est une erreur**.
- La parade est **`IDisposable` + `Using`** : le `End Using` **réancre** la libération à un point précis, comme le faisait la chute du compteur à zéro en VB6. Réflexe : **« objet à ressource ⇒ `Using` »**.
- Le **motif `Dispose` complet** (filet `Finalize`, `SuppressFinalize`) est traité en **12.3** ; il ne concerne que les ressources **non managées** détenues directement.

---

> 🧭 **Section suivante** → [2.3 Le typage : `Variant`, entiers redimensionnés, `Currency`](03-typage.md) ⚠️  
> 🔝 **Retour** → [Sommaire du chapitre 2](README.md)

⏭️ [Le typage : `Variant`, entiers redimensionnés, `Currency`](/02-differences-fondamentales/03-typage.md)
