🔝 Retour au [Sommaire](/SOMMAIRE.md)

# F. Modèles de tests de non-régression (*golden master*)

**Patrons pour capturer le comportement de référence de l'application VB6 et le comparer
automatiquement après migration.**

> 🎯 **Pourquoi cette technique.** Le *golden master* (ou test de **caractérisation**) capture la
> **sortie** du système VB6 existant pour un large jeu d'entrées, **fige** cette sortie comme
> référence approuvée, puis rejoue **les mêmes entrées** sur la version VB.NET et **compare**. Tout
> écart est une régression à instruire. C'est **la** technique adaptée à une migration : on cherche
> à **préserver le comportement**, pas à le **comprendre** — on n'a donc pas besoin de connaître le
> détail du code pour le sécuriser. C'est le **fil rouge** posé au module 5.5 et exploité au
> module 17.3, et le détecteur principal des **pièges silencieux** (Annexe B).

> ⚠️ **Règle d'or du golden master** : on ne met **jamais** à jour la référence pour faire passer un
> écart au vert **sans comprendre pourquoi**. Mettre à jour le baseline, c'est **approuver un
> changement de comportement** — décision à tracer (voir §9).

---

## 1. Le principe et le flux de travail

Six étapes, à outiller une fois puis à rejouer à volonté :

1. **Délimiter la surface de sortie** à capturer (§2).
2. **Rassembler des entrées représentatives** — y compris les cas limites liés aux pièges (dates,
   montants, culture, bornes, sorties modifiées par `ByRef`).
3. **Capturer la référence** depuis VB6 → fichiers **approuvés** (`*.approved.txt`).
4. **Rejouer** les **mêmes** entrées sur VB.NET → sortie **obtenue**.
5. **Comparer** obtenu vs approuvé (§5).
6. **Trier** les écarts : régression à corriger, ou changement voulu à ré-approuver (§9).

> 💡 Pré-requis qui facilite tout : **découpler l'UI de la logique métier** (module 5.3). Le golden
> master est bien plus simple et robuste sur de la **logique testable** que sur des formulaires
> (voir §7).

---

## 2. Que capturer (la surface de sortie)

| Type d'application | Surface à capturer |
|---|---|
| Calcul / logique métier | Valeurs de retour, champs calculés |
| Générateur de rapports | Couche **texte** du rapport (ou PDF → texte) |
| Export de données | Fichiers **CSV / XML / largeur fixe** produits |
| Écritures en base | **Vidage ordonné** des lignes affectées |
| Traitement de fichiers | Fichiers de sortie |
| Gestion d'erreurs | **Code/message** d'erreur **et** branche empruntée |
| Logique d'IHM (découplée) | Valeurs de propriétés résultantes / état du « *view-model* » |

> Capturer une surface **observable et stable** : préférer des **valeurs** (texte sérialisé) à des
> objets vivants. Sérialiser dans un format **lisible et diffable** (texte ligne à ligne, CSV, JSON).

---

## 3. Rendre la capture déterministe *(crucial)*

Sans cela, le golden master produit de **faux écarts**. Neutraliser **toutes** les sources de  
non-déterminisme, des deux côtés, **à l'identique** :

| Source de non-déterminisme | Côté VB6 | Côté VB.NET |
|---|---|---|
| Date/heure courante (`Now`, `Date`, `Time`) | Injecter une date **fixe** ; **bannir** `Now` dans les cas | Injecter un `DateTime` fixe (paramètre / horloge simulée) |
| **Culture / régional** ⚠️ | Poste à culture fixe ; **`Val`/`Str`** (invariant) | **`CultureInfo.InvariantCulture`** explicite ; fixer `CurrentCulture` du thread |
| Aléatoire (`Rnd`) | `Rnd -1 : Randomize 42` (réinitialise puis fixe la graine) | `New Random(42)` |
| Ordre (collections, dictionnaires) | **Trier** avant écriture | **`OrderBy`** avant écriture |
| Chemins / fichiers temporaires | **Masquer** (§4) | **Masquer** (§4) |
| Identifiants volatils (GUID, ID auto) | Masquer | Masquer |
| Encodage de fichier | Fixe (ANSI/UTF-8 défini) | `Encoding` **explicite** |

> ⚠️ La **culture** est le piège n°1 du golden master (et de la migration : Annexe B.10). La parade
> retenue dans les patrons ci-dessous : **sérialiser le baseline en invariant** — parser les entrées
> avec `Val` (VB6) / `InvariantCulture` (.NET) et formater les sorties avec `Str`/`InvariantCulture`.
> Le baseline devient **neutre**, la comparaison **non ambiguë**.

---

## 4. Normaliser / masquer avant comparaison

Avant de comparer, **gommer** les valeurs volatiles résiduelles et harmoniser les fins de ligne :

```vbnet
Imports System.Text.RegularExpressions

' Neutralise les valeurs volatiles, puis normalise les fins de ligne
Private Function Normaliser(texte As String) As String
    texte = Regex.Replace(texte, "\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2}", "<HORODATAGE>")
    texte = Regex.Replace(texte, "[A-Za-z]:\\[^\s""]+", "<CHEMIN>")
    texte = Regex.Replace(texte,
        "\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b",
        "<GUID>")
    Return texte.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).TrimEnd()
End Function
```

> 📌 VB6 (`Print #`) écrit des fins de ligne **CRLF** ; la normalisation évite les faux écarts si un
> baseline est récupéré avec des fins de ligne LF (Git, autre éditeur).

---

## 5. Stratégies de comparaison et tolérances

| Type de sortie | Comparaison | Tolérance |
|---|---|---|
| Texte / chaînes | **exacte** | aucune |
| Entiers | **exacte** | aucune |
| `Decimal` (ex-`Currency`, montants) | **exacte** | **aucune** ⚠️ (un écart d'un centime = un vrai bug) |
| `Single` / `Double` | arrondi à **N décimales** *avant* comparaison, ou *epsilon* | *epsilon* réduit |
| Dates | **exacte** (après neutralisation de l'heure courante) | aucune |
| Fichiers / exports | **octet à octet** après masquage des champs volatils | — |

> ⚠️ Réserver la **tolérance** au **flottant binaire** (`Single`/`Double`), où la représentation peut
> varier. Pour les **montants**, viser `Decimal` (Annexe B.9) et comparer **exactement**.

---

## 6. Les patrons

> Convention commune aux patrons : un fichier d'entrées **partagé** `golden/entree.csv` (lu par les
> deux côtés), un baseline `golden/<cas>.approved.txt` produit par VB6, et un rejeu .NET comparé à
> ce baseline. Les deux côtés **parsent en invariant** et **formatent en invariant**.

### 6.1 Patron — Harnais de capture (VB6, produit le baseline)

```vb
' VB6 — Harnais de capture du golden master
' Lit golden\entree.csv, écrit golden\facturation.approved.txt (lignes "entrée => sortie")
Public Sub CapturerReference()
    Dim nIn As Integer  : nIn = FreeFile
    Open App.Path & "\golden\entree.csv" For Input As #nIn
    Dim nOut As Integer : nOut = FreeFile
    Open App.Path & "\golden\facturation.approved.txt" For Output As #nOut

    Dim ligne As String, p() As String
    Do Until EOF(nIn)
        Line Input #nIn, ligne
        p = Split(ligne, ";")
        Dim montant As Currency : montant = CCur(Val(p(0)))   ' Val = invariant (point décimal)
        Dim taux    As Double   : taux    = Val(p(1))
        ' Sérialisation invariante de la sortie (Trim$(Str$()) : pas d'espace, point décimal)
        Print #nOut, ligne & " => " & Trim$(Str$(Facturer(montant, taux)))
    Loop

    Close #nIn : Close #nOut
End Sub
```

### 6.2 Patron — Harnais de rejeu + comparaison (VB.NET)

```vbnet
Imports System.Globalization
Imports System.IO
Imports System.Text

' Rejoue les MÊMES entrées et produit la même sérialisation (invariante)
Private Function Rejouer() As String
    Dim ci = CultureInfo.InvariantCulture
    Dim sb As New StringBuilder()
    For Each ligne In File.ReadAllLines("golden\entree.csv")
        Dim p = ligne.Split(";"c)
        Dim montant = Decimal.Parse(p(0), ci)     ' Currency → Decimal (Annexe B.9)
        Dim taux    = Double.Parse(p(1), ci)
        sb.AppendLine($"{ligne} => {Facturer(montant, taux).ToString(ci)}")
    Next
    Return sb.ToString()
End Function

<TestMethod>
Public Sub Facturation_NonRegression()
    Approuver("facturation", Rejouer())          ' voir le patron d'approbation (§6.3)
End Sub
```

### 6.3 Patron — Approbation (*received* / *approved*)

Compare la sortie obtenue au fichier **approuvé**. En cas d'écart (ou d'absence de référence),  
le test **échoue** et laisse un fichier `*.received.txt` à examiner, puis à promouvoir si l'écart est  
voulu (§9).

```vbnet
Private Sub Approuver(nomCas As String, obtenu As String)
    Dim dossier  = "golden"
    Dim approuve = Path.Combine(dossier, nomCas & ".approved.txt")
    Dim recu     = Path.Combine(dossier, nomCas & ".received.txt")

    File.WriteAllText(recu, obtenu, Encoding.UTF8)   ' encodage explicite

    If Not File.Exists(approuve) Then
        Assert.Fail($"Aucune référence approuvée pour « {nomCas} ». " &
                    $"Examiner {recu}, puis le renommer en {nomCas}.approved.txt.")
        Return
    End If

    Dim attendu = File.ReadAllText(approuve)
    If Normaliser(attendu) <> Normaliser(obtenu) Then
        Assert.Fail($"Écart détecté pour « {nomCas} ». Comparer {recu} et {approuve}.")
    Else
        File.Delete(recu)   ' propre : on supprime le « received » quand tout concorde
    End If
End Sub
```

> 🛠️ Des bibliothèques **d'*approval testing*** (ex. *ApprovalTests* pour .NET) automatisent ce
> motif *received/approved* et l'ouverture d'un outil de *diff*. Le patron ci-dessus en montre la
> logique, sans dépendance.

### 6.4 Patron — Piloté par les données (un fichier d'entrées → un fichier de sortie)

Idéal pour **beaucoup** de cas : on ajoute des lignes à `entree.csv` sans toucher au code de test.  
C'est exactement la structure des §6.1-6.2 (boucle sur `entree.csv`). Avantages : couverture qui
**grandit** avec les cas réels rencontrés, et **un seul** fichier à *differ*.

### 6.5 Patron — Épingler une fonction (*characterization*, cas anti-pièges)

Pour **figer** le comportement d'une fonction sensible, balayer des entrées qui **révèlent les  
pièges silencieux** (arrondi banquier, bornes, négatifs, dates) :

```vbnet
' Balayage ciblant les pièges : arrondi (B.9), troncature (A §8.4), bornes
<TestMethod>
Public Sub Arrondis_NonRegression()
    Dim ci = CultureInfo.InvariantCulture
    Dim cas = {-2.5D, -0.5D, 0D, 0.5D, 1.5D, 2.5D, 99999999.99D}
    Dim sb As New StringBuilder()
    For Each x In cas
        sb.AppendLine($"{x.ToString(ci)} => " &
                      $"CInt={CInt(x)} Int={Int(x)} Fix={Fix(x)} " &
                      $"Round2={Math.Round(x, 2, MidpointRounding.AwayFromZero).ToString(ci)}")
    Next
    Approuver("arrondis", sb.ToString())
End Sub
```

> Le baseline correspondant doit d'abord être **capturé depuis VB6** (mêmes entrées, même
> sérialisation invariante) pour servir de juge : c'est lui qui révèle, par exemple, qu'un `Int`
> appliqué à un négatif ne donne pas le même résultat qu'un `Fix` (Annexe A §8.4, Annexe B.9).

---

## 7. Cas particulier : l'interface (formulaires)

Le golden master est **le plus efficace sur la logique découplée** (module 5.3). Pour l'IHM  
elle-même :

- **Préférer extraire la logique** derrière le formulaire et capturer les **valeurs résultantes**
  (propriétés, état du *view-model*) plutôt que des pixels.
- À défaut, un **vidage de propriétés** (libellés, états activé/visible, contenu des listes) reste
  diffable ; une **capture d'écran** est possible mais **fragile** (DPI, polices, thème — voir
  Annexe C : twips→pixels).
- ⚠️ Éviter de faire dépendre le test de l'ordre d'événements `GotFocus`/`Enter` ou d'horodatages
  (Annexe C §2) : neutraliser comme en §3.

---

## 8. Organisation, versionnement, intégration continue

- **Versionner les fichiers `*.approved.txt`** (et `entree.csv`) dans le dépôt : ils **font partie**
  du code et tracent le comportement de référence.
- **Arborescence** suggérée : `golden/<module>/<cas>.approved.txt` (un dossier par module migré).
- **Granularité** : *un fichier par cas* (localisation facile de l'écart) **ou** *un fichier
  agrégé* (diff global plus simple) — choisir selon la taille.
- **Intégration continue** : **rejouer la comparaison à chaque build** et **échouer** au moindre
  écart non approuvé. Le golden master devient une **barrière** automatique de non-régression.

---

## 9. Triage des écarts

| Verdict | Signification | Action |
|---|---|---|
| 🟢 **Vert** | Sortie identique | Rien — comportement préservé. |
| 🔴 **Rouge** | Écart **non voulu** | **Bug de migration** : corriger le **code** (souvent un piège de l'Annexe B). |
| 🟠 **Ambre** | Écart **voulu** (ex. correction d'un bug VB6 historique) | **Justifier**, **consigner** dans le journal des décisions (module 5.4 / Annexe E), **puis** ré-approuver le baseline (promouvoir `*.received.txt`). |

> ⚠️ Ne jamais transformer un 🔴 en 🟢 en **éditant la référence** : ce serait masquer une régression.
> La référence ne change **que** sur un écart **ambre** assumé et tracé.

---

## 10. 🔗 Annexes et modules liés

- **Module 5.5** — Mettre en place le harnais de tests de référence (*golden master*).
- **Module 5.3** — Découpler l'UI de la logique métier (rend le golden master simple et robuste).
- **Module 5.4** — Documenter les comportements critiques (et le **journal des décisions** du §9).
- **Module 17.3** — Tests de non-régression : comparer avec le golden master.
- **Module 17.4** — Traquer les pièges silencieux (ce que le golden master fait **remonter**).
- **Annexe B** — Catalogue des pièges silencieux (le « pourquoi » des cas anti-pièges du §6.5).
- **Annexe E** — Checklist (où le golden master est le **fil rouge** des phases avant/après).

---

**Juin 2026**  
**Cible** : .NET Framework 4.7.2 (le « pont » depuis VB6)  
**Licence** : Creative Commons BY-NC-SA 4.0

⏭️ [Glossaire et acronymes](/annexes/glossaire/README.md)
