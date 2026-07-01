🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.6 Instructions multiples, séparateur « `:` », étiquettes de ligne

> **Module 9 — Opérateurs, chaînes et structures de contrôle** · [↩ Retour au chapitre](README.md)

Voici la **dernière section** du module 9, et la plus **paisible**. Tout ce qu'elle couvre — plusieurs instructions par ligne, le séparateur `:`, les étiquettes de ligne — est **conservé à l'identique** en VB.NET. Il n'y a donc **aucune rupture** à gérer. Son intérêt est ailleurs : **confirmer** cette continuité, **clarifier le double rôle** du `:`, et signaler deux points de **lisibilité**.

---

## Le séparateur « `:` » : plusieurs instructions sur une ligne

Le `:` sépare plusieurs instructions sur une même ligne, exactement comme en VB6 :

```vb
' Conservé à l'identique
x = 1 : y = 2 : z = 3
```

C'est valide, mais **déconseillé** en code moderne : une **instruction par ligne** se lit et se débogue mieux (points d'arrêt, couverture, lecture des différences en revue de code). Lors du refactoring (module **17.5**), c'est typiquement ce qu'on déplie.

---

## Le piège de lisibilité : `If` sur une ligne **+** « `:` »

Un cas mérite attention, non parce qu'il change à la migration (il ne change **pas**), mais parce qu'il se **lit mal**. Dans un `If` sur une seule ligne, **tout ce qui suit `Then`** — y compris les instructions séparées par `:` — appartient à la **branche conditionnelle** :

```vb
' VB6 ET VB.NET : même comportement
If estValide Then x = 1 : Enregistrer()
'   -> x = 1 ET Enregistrer() ne s'exécutent QUE si estValide est vrai
'   (et NON : « x = 1 si valide, Enregistrer() toujours » — erreur de lecture fréquente)
```

Le comportement est **préservé**, mais comme on lit souvent `Enregistrer()` comme une instruction *inconditionnelle*, préférez la **forme bloc**, sans ambiguïté :

```vb
If estValide Then
    x = 1
    Enregistrer()
End If
```

> 🔗 La forme sur une ligne a déjà été évoquée au module **9.3** (structures conditionnelles). Le point reste valable ici sous l'angle du séparateur `:`.

---

## Le **double rôle** du « `:` » : séparateur **et** étiquette

Le caractère `:` a **deux usages distincts** en VB — une source de confusion fréquente :

```vb
Reprise:                          ' « : » = fin d'une ÉTIQUETTE de ligne
    Ouvrir() : Lire() : Fermer()  ' « : » = SÉPARATEUR d'instructions
```

Sur la première ligne, `Reprise:` **définit une étiquette** (cible de saut). Sur la seconde, les `:` **séparent des instructions**. Les deux rôles sont **conservés** en VB.NET ; il suffit de les distinguer : une étiquette est un identifiant **seul** suivi de `:`, en tête de ligne.

---

## Étiquettes de ligne

Les **étiquettes de ligne** (`NomEtiquette:`) sont **conservées** en VB.NET. Elles servent de cibles à :

- `GoTo label` — **conservé** (toujours déconseillé) ;
- `On Error GoTo label` — **conservé** (gestion d'erreurs non structurée, module **11**).

```vb
' Étiquette préservée ; cible de On Error GoTo
Sub Traiter()
    On Error GoTo Gestion
    Ouvrir()
    Exit Sub
Gestion:                          ' étiquette de ligne
    Journaliser(Err.Description)
End Sub
```

Les étiquettes ne disparaissent donc pas — mais un code **structuré** (méthodes, `Try`/`Catch`, `Select Case`) en a **beaucoup moins besoin**. Deux précisions :

- **Numéros de ligne** (`10`, `20`…) : leur rôle historique d'étiquette/cible est **obsolète** en VB.NET. Voir module **9.5**.
- **Étiquettes générées par l'assistant** : en convertissant la gestion d'erreurs, l'assistant de migration **produit parfois** des étiquettes et des `GoTo`. C'est fonctionnel mais peu idiomatique : à remplacer par `Try`/`Catch` lors de la fiabilisation (module **11**).

---

## L'inverse : la continuation de ligne (« `_` » et continuation implicite)

Là où `:` **regroupe** plusieurs instructions sur une ligne, on a parfois le besoin **inverse** : **étaler** une instruction sur plusieurs lignes. VB6 imposait le **trait de soulignement** `_` en fin de ligne ; il est **conservé** en VB.NET, mais souvent **inutile** grâce à la **continuation implicite** (depuis VB 2010) après une virgule, un opérateur, une parenthèse ouvrante, etc. :

```vb
' Continuation explicite (« _ ») — conservée
total = prixHT _
      + taxe

' Continuation IMPLICITE — souvent, plus besoin du « _ »
total = prixHT +
        taxe
```

Le style moderne privilégie donc **une instruction par ligne** (avec continuation implicite si nécessaire) plutôt que le **regroupement par `:`**.

---

## 🎯 Synthèse — points d'attention de la section

| Élément | Statut | Note de style |
|---------|--------|---------------|
| Séparateur « `:` » | Conservé | Éviter ; une instruction par ligne |
| `If` sur une ligne + « `:` » | Conservé (tout dans le `Then`) | Préférer la forme **bloc** (lisibilité) |
| « `:` » d'étiquette vs de séparation | Conservés (deux rôles) | Distinguer : étiquette = identifiant seul + `:` |
| Étiquettes de ligne (`Label:`) | Conservées | Cibles de `GoTo`/`On Error GoTo` ; structurer si possible |
| Numéros de ligne | Sans rôle (voir 9.5) | Supprimer |
| Continuation « `_` » | Conservée | Souvent inutile (continuation implicite) |

---

## 🔗 Pour aller plus loin

- Module **9.3** — Structures conditionnelles (la forme `If` sur une ligne)
- Module **9.5** — Numéros de ligne, `GoTo`, `On…GoTo` (le contexte des étiquettes)
- Module **11** — `On Error GoTo` → `Try`/`Catch` (les étiquettes de gestion d'erreurs)
- Module **17.5** — Refactoring idiomatique (déplier les `:`, supprimer les étiquettes superflues)
- Annexe **A** — Tableau de correspondance VB6 → VB.NET

---

## ✅ Fin du module 9

Ce module a couvert la **grammaire quotidienne** du code — opérateurs, chaînes, structures de contrôle — et ses **pièges silencieux** : l'**évaluation court-circuit** (`And`/`Or` vs `AndAlso`/`OrElse`, `IIf` vs `If()`), l'**indexation des chaînes** (1-based vs 0-based, `Microsoft.VisualBasic` vs `String`), les **constructions supprimées** (`GoSub`/`Return`, `On…GoTo`, numéros de ligne) et, ici, des éléments **entièrement conservés** mais à soigner pour la lisibilité.

Le **module 10** enchaîne sur les **procédures, fonctions et paramètres** — où resurgit le piège majeur **ByRef par défaut → ByVal par défaut**, déjà entrevu au module 2.4.

---

> **Section précédente** → [9.5 — `GoSub`/`Return`, `On…GoTo`, numéros de ligne : supprimés](05-gosub-ongoto-supprimes.md)
> **Module suivant** → [10. Procédures, fonctions et paramètres](../10-procedures-fonctions/README.md) ⭐ ⚠️

⏭️ [Procédures, fonctions et paramètres](/10-procedures-fonctions/README.md)
