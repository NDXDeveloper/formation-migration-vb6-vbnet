🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 12.8 — Événements : `Event` / `RaiseEvent` / `WithEvents` / `Handles` ; `AddHandler` / `RemoveHandler`

> **Chapitre 12 — Programmation orientée objet** · Section 12.8 *(dernière section du chapitre)*
> Le modèle déclaratif d'événements **évolue** (la clause `Handles`), un modèle **dynamique** **apparaît** (`AddHandler`), et un **piège mémoire** propre à .NET fait son entrée.

---

## 🧭 Périmètre

Les événements existaient déjà en VB6, et **l'essentiel se transpose sans douleur**. Trois nuances méritent l'attention : le **câblage** des abonnés passe d'une **convention de nommage** à une **clause `Handles`** ; un **câblage dynamique** (`AddHandler`/`RemoveHandler`) **nouveau** apparaît ; et les **abonnements** peuvent provoquer des **fuites mémoire** — un sujet qui renoue avec le GC (12.2/12.3).

---

## 1. Déclarer et déclencher : `Event` / `RaiseEvent` (quasi inchangé)

La **déclaration** et le **déclenchement** d'un événement sont **identiques** d'un langage à l'autre.

```vb
' VB6 — CCompteur déclenche un événement au-delà d'un seuil
Public Event SeuilAtteint(ByVal valeur As Long)

Private mValeur As Long

Public Sub Incrementer()
    mValeur = mValeur + 1
    If mValeur >= 100 Then
        RaiseEvent SeuilAtteint(mValeur)
    End If
End Sub
```

```vb
' VB.NET — équivalent (signature VB6 conservée)
Public Class CCompteur
    Public Event SeuilAtteint(valeur As Integer)   ' Long → Integer (7.2)

    Private _valeur As Integer

    Public Sub Incrementer()
        _valeur += 1
        If _valeur >= 100 Then
            RaiseEvent SeuilAtteint(_valeur)
        End If
    End Sub
End Class
```

Deux remarques :

- **`RaiseEvent` est sûr même sans abonné** : inutile de vérifier la présence d'un destinataire, VB.NET s'en charge.
- **Forme idiomatique .NET** : on déclare souvent l'événement via un **délégué** et la signature `(sender, e)` — `Public Event SeuilAtteint As EventHandler(Of Integer)`. Pour une **migration fidèle**, **conservez la signature VB6** ; la convention `(sender, e)` est un **refactoring** (17.5).
- Rappel VB6 : les événements ne peuvent être déclarés/déclenchés que dans des **modules de classe**.

---

## 2. Recevoir : `WithEvents` + `Handles` (l'évolution du modèle déclaratif)

En **VB6**, on s'abonne en déclarant la variable `WithEvents`, et le handler est relié par **convention de nommage** (`Variable_Événement`) :

```vb
' VB6 — abonné via WithEvents + convention de nommage
Private WithEvents mCompteur As CCompteur

Private Sub Form_Load()
    Set mCompteur = New CCompteur
End Sub

Private Sub mCompteur_SeuilAtteint(ByVal valeur As Long)   ' nom IMPOSÉ
    MsgBox "Seuil : " & valeur
End Sub
```

En **VB.NET**, on garde `WithEvents`, mais le lien se fait par une **clause `Handles`** — et le **nom du handler devient libre** :

```vb
' VB.NET — abonné via WithEvents + clause Handles (nom de méthode LIBRE)
Private WithEvents _compteur As New CCompteur()

Private Sub QuandSeuil(valeur As Integer) Handles _compteur.SeuilAtteint
    MsgBox("Seuil : " & valeur)
End Sub
```

Ce que VB.NET apporte :

- le **nom** du handler est **libre** (la clause `Handles` fait le lien, plus la convention) ;
- un **même handler** peut traiter **plusieurs** événements : `… Handles _compteur.SeuilAtteint, _autre.SeuilAtteint` ;
- comme en VB6, `WithEvents` reste réservé aux variables **de niveau classe** (pas de variable **locale**).

---

## 3. Le modèle **dynamique** : `AddHandler` / `RemoveHandler` (nouveauté .NET)

VB.NET ajoute un câblage **à l'exécution**, **sans équivalent VB6**. Il fonctionne sur **n'importe quelle** référence — y compris une variable **locale** ou un objet **créé dynamiquement**.

```vb
' VB.NET — câblage dynamique
Dim c As New CCompteur()
AddHandler c.SeuilAtteint, AddressOf QuandSeuil      ' on relie
' …
RemoveHandler c.SeuilAtteint, AddressOf QuandSeuil   ' on délie

' Avec une lambda :
AddHandler c.SeuilAtteint, Sub(v) Console.WriteLine($"Seuil : {v}")
```

Points utiles :

- `AddressOf` crée un **délégué** vers la méthode handler ; on peut aussi passer une **lambda**.
- On peut attacher **plusieurs** handlers au même événement, et les retirer **individuellement**.
- ⚠️ Une **lambda** attachée par `AddHandler` est **difficile à retirer** (pas de référence stockée). Si vous devez vous désabonner, **conservez le délégué** dans une variable plutôt qu'une lambda anonyme.

Ce modèle **résout** le cas VB6 épineux des objets **créés dynamiquement** (qu'on ne pouvait pas câbler par `WithEvents`).

---

## 4. ⚠️ Le piège : abonnement et **durée de vie** (fuite mémoire)

C'est le point de vigilance majeur des événements en .NET — et il prolonge le thème du **GC** (12.2/12.3).

Quand un objet **s'abonne** à un événement, la **source** détient une **référence** vers l'**abonné** (via le délégué). **Conséquence** : tant que la source vit **et** que le handler reste attaché, l'abonné **ne peut pas être collecté** par le GC. Oublier de se **désabonner** crée une **fuite mémoire managée** (le « *lapsed listener* ») : les abonnés s'**accumulent**, jamais récupérés.

```vb
' ⚠️ Fuite : un abonnement jamais retiré sur une source à longue durée de vie
AddHandler sourceLongueDuree.Evt, AddressOf Me.Handler
' Tant que 'sourceLongueDuree' vit et que le handler est attaché,
' CET objet (Me) reste en vie — il ne sera jamais collecté.
```

Bonnes pratiques :

- **`WithEvents` / `Handles`** : le câblage est **géré par la variable** ; lui réaffecter un objet (ou `Nothing`) **délie** automatiquement l'ancien. Ce modèle est **moins** sujet aux fuites.
- **`AddHandler`** : **vous** êtes responsable du **`RemoveHandler`**. Pour un abonné à courte durée de vie relié à une source durable, désabonnez-vous — souvent dans **`Dispose`** (voir **12.3**) :

```vb
' VB.NET — se désabonner dans Dispose (rejoint 12.3)
Public Sub Dispose() Implements IDisposable.Dispose
    RemoveHandler sourceLongueDuree.Evt, AddressOf Me.Handler
End Sub
```

---

## 5. La traduction, en pratique

| VB6 | VB.NET |
|-----|--------|
| `Public Event Evt(...)` | `Public Event Evt(...)` (idem ; ou délégué `EventHandler(Of T)`) |
| `RaiseEvent Evt(...)` | `RaiseEvent Evt(...)` (sûr même sans abonné) |
| `Private WithEvents v As C` | `Private WithEvents v As C` (toujours niveau classe) |
| `Private Sub v_Evt(...)` (nom imposé) | `Private Sub NomLibre(...) Handles v.Evt` |
| *(câblage statique uniquement)* | **`AddHandler` / `RemoveHandler`** (dynamique, nouveau) |
| objet créé dynamiquement non câblable | câblé par **`AddHandler`** |

> 💡 L'assistant de mise à niveau convertit les handlers nommés `v_Evt` en méthodes munies d'une **clause `Handles v.Evt`**, en gardant souvent le **même nom** — le code reste lisible.

**Lien avec les *control arrays*** (chapitre 13) : un *tableau de contrôles* VB6 utilisait **un** handler partagé par plusieurs contrôles (via `Index`). En Windows Forms, on relie plusieurs contrôles à **un** handler avec `Handles bouton1.Click, bouton2.Click` ou par `AddHandler` — c'est le **mécanisme de remplacement** (voir **13.6**).

---

## 6. Pour mémoire : `Custom Event`

VB.NET permet aussi un **`Custom Event`** avec des accesseurs explicites (`AddHandler` / `RemoveHandler` / `RaiseEvent`), utile pour contrôler finement le stockage des abonnements. C'est un mécanisme **avancé**, **rarement nécessaire** en migration — à connaître, sans s'y attarder.

---

## 🔁 Récapitulatif VB6 → VB.NET

| Aspect | VB6 | VB.NET |
|--------|-----|--------|
| Déclarer / déclencher | `Event` / `RaiseEvent` | identique |
| Recevoir (déclaratif) | `WithEvents` + nom `v_Evt` | `WithEvents` + **clause `Handles`** (nom libre) |
| Recevoir (dynamique) | *(impossible)* | **`AddHandler` / `RemoveHandler`** |
| Désabonnement | réaffecter / `Set … = Nothing` | géré par `WithEvents`, ou **`RemoveHandler`** explicite |
| Risque mémoire | faible (refcount + câblage statique) | **fuite** si abonnement non retiré ⚠️ |

---

## ✅ Points clés

- `Event` et `RaiseEvent` se transposent **à l'identique** (`RaiseEvent` est sûr **sans abonné**).
- Le modèle déclaratif évolue : `WithEvents` + **clause `Handles`** remplace la **convention de nommage** ; le nom du handler devient **libre**, et un handler peut couvrir **plusieurs** événements.
- **`AddHandler` / `RemoveHandler`** ajoutent un câblage **dynamique** **nouveau**, sur n'importe quelle référence — idéal pour les **objets créés à l'exécution**.
- ⚠️ Un **abonnement** maintient l'abonné **en vie** : pensez au **`RemoveHandler`** (souvent dans **`Dispose`**) pour éviter une **fuite mémoire**.
- Le modèle d'événements est aussi le **remplaçant** des *control arrays* (un handler pour plusieurs contrôles, voir **13.6**).

---

## 🔗 Renvois

- **12.3** — `IDisposable` / `Dispose` (où retirer les abonnements). · **12.2** — durée de vie et GC.
- **7.2** — entiers redimensionnés (signatures d'événements). · **17.5** — refactoring (signature `(sender, e)`).
- **13.6** — *control arrays* disparus → handlers partagés (`Handles …, …` / `AddHandler`).
- **Fin du chapitre 12.** Suite → **[13. Des formulaires VB6 à Windows Forms](../13-formulaires-winforms/README.md)** ⭐ ⚠️

⏭️ [Des formulaires VB6 à Windows Forms](/13-formulaires-winforms/README.md)
