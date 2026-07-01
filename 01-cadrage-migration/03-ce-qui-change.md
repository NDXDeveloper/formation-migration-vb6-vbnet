🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 1.3 VB6 vs VB.NET : ce qui change vraiment ⭐

> **Chapitre 1 — Introduction : pourquoi et comment migrer**
> Le nom se ressemble — « VB » d'un côté, « VB » de l'autre. C'est précisément ce qui trompe.

---

## La continuité du nom cache une rupture

« On reste en Visual Basic, c'est juste une version plus récente. » Voilà sans doute l'idée reçue
la **plus coûteuse** de toute la migration. Le nom rassure ; il ne devrait pas. Entre VB6 et VB.NET,  
le **langage** se ressemble en surface, mais en dessous, **le moteur d'exécution et le modèle objet  
sont entièrement différents**. C'est moins un « passage à la version suivante » qu'un changement de  
fondations sur lequel on aurait gardé une façade familière.

Cette section ne détaille pas encore chaque écart — c'est le rôle du **chapitre 2** et de  
l'**annexe B**. Son objectif est de vous donner le **bon modèle mental** : comprendre *quelles  
catégories de choses changent*, et surtout **lesquelles sont dangereuses**. Car tous les changements  
ne se valent pas.

---

## 🧊 Le modèle de l'iceberg : deux catégories de changements

Pour aborder la migration sans se faire surprendre, il faut séparer les changements en **deux  
familles radicalement différentes** :

| | **Changements visibles** | **Changements silencieux** |
|---|---|---|
| **Nature** | Syntaxe, mots-clés, structures | Comportement à l'exécution |
| **Qui les détecte ?** | **Le compilateur** | **Personne** (le code compile) |
| **Symptôme** | « Ça ne compile pas » | « Ça compile, mais ça ne fait plus pareil » |
| **Danger** | Faible — le compilateur vous guide | **Élevé** — rien ne vous alerte |
| **Effort** | Mécanique, fastidieux | **Jugement**, investigation, tests |

> ⚠️ **C'est tout l'enjeu de la migration.** Les changements **visibles** sont au-dessus de la
> ligne de flottaison : pénibles, mais sûrs, car le compilateur refuse de générer un programme
> incorrect. Les changements **silencieux** sont sous la surface : ils n'arrêtent pas la
> compilation — ils déforment le comportement. **C'est la coque, pas le sommet de l'iceberg, qui
> coule le navire.**

Gardez cette grille en tête en parcourant les quatre axes qui suivent : à chaque fois, la question  
utile est « ce changement est-il visible ou silencieux ? ».

---

## 1️⃣ Le langage — le changement visible (et le moins dangereux)

C'est la partie la plus apparente, et paradoxalement la moins risquée. Mots-clés, types, structures  
de contrôle, syntaxe : beaucoup de choses changent.

- Le mot-clé **`Set`** disparaît (affectation d'objet) ; `GoSub`/`Return` et `On…GoTo` sont **supprimés** ; `While…Wend` devient `While…End While`.
- Les **types entiers sont redimensionnés** (l'`Integer` de VB6 fait 16 bits, celui de .NET en fait 32) ; `Currency` cède la place à `Decimal` ; `Variant` devient `Object`.
- De nouvelles possibilités apparaissent (surcharge, héritage, génériques…).

Pour l'essentiel, ces écarts relèvent du **changement visible** : le code ne compile pas tant qu'on  
ne les a pas traités. Fastidieux, mais sûr.

> ⚠️ **L'exception piégeuse.** Une **minorité** de changements de langage sont **silencieux** — et
> ce sont les plus dangereux de tous, car ils se cachent au milieu d'une zone réputée « sûre ».
> Trois exemples emblématiques (détaillés au chapitre 2) :
> - **`ByRef` par défaut → `ByVal` par défaut** : le même appel de procédure ne modifie plus la variable de l'appelant. *Le code compile, le résultat change.*
> - **Entiers redimensionnés** : un calcul qui tenait dans un `Integer` 16 bits peut se comporter différemment une fois en 32 bits (et inversement lors d'appels d'API).
> - **`True` vaut `-1`** en VB6 (tous les bits à 1) : dès qu'un booléen est converti en nombre — stocké en base, passé à une API, utilisé dans un masque de bits — la valeur `-1` (et non `1`) réserve des surprises au reste du monde .NET, qui attend `1`.

**Un exemple, pour rendre le danger concret.** Le même code, migré tel quel, ne se comporte plus
pareil — et *rien* ne vous prévient :

```vb
' VB6 — le paramètre est ByRef PAR DÉFAUT
Sub Doubler(n As Integer)   ' « ByRef » implicite
    n = n * 2
End Sub

Dim x As Integer
x = 10
Doubler x                   ' appel sans parenthèses
Debug.Print x               ' ➜ 20  : x a bien été modifié
```

```vb
' VB.NET — le même paramètre est ByVal PAR DÉFAUT
Sub Doubler(n As Integer)   ' « ByVal » implicite
    n = n * 2
End Sub

Dim x As Integer = 10
Doubler(x)
Console.WriteLine(x)        ' ➜ 10  : x est resté intact !
```

Les deux versions **compilent sans la moindre erreur ni le moindre avertissement**. Seul le
*résultat* diffère — et aucun outil ne peut deviner à votre place si l'appelant *comptait* sur la
modification. C'est toute la signature du changement silencieux. (Audit systématique : sections
**2.4** et **10.2**.)

👉 **Le détail** : Partie 3 (chapitres 7 à 12), et l'aide-mémoire de l'**annexe A**.

---

## 2️⃣ Le runtime — VBVM contre CLR : deux moteurs, deux mondes

Voici la première rupture **invisible dans le code source**, et pourtant fondamentale.

- VB6 s'exécute sur le **VBVM** (*Visual Basic Virtual Machine*, `msvbvm60.dll`) : le moteur historique, fondé sur **COM**.
- VB.NET s'exécute sur le **CLR** (*Common Language Runtime*) : le moteur de .NET, avec **exécution managée**, compilation **JIT** (*Just-In-Time*, à la volée), système de types commun (CTS), assemblies, sécurité de type.

Ce n'est **pas** « le même moteur en plus récent » : c'est un **moteur différent, avec une  
philosophie différente**. Et de ce changement de moteur découlent les conséquences les plus  
structurantes de la migration : la **gestion automatique de la mémoire** (le point suivant), les
**exceptions** à la place des codes d'erreur, la sécurité de type, le modèle d'assemblies.

> 💡 En une phrase : on ne change pas seulement la **langue** qu'on parle (le langage), on change le
> **pays** dans lequel on la parle (le runtime). Et les coutumes locales ne sont plus les mêmes.

👉 **Le détail** : section **2.1 — Deux runtimes, deux mondes (VBVM vs CLR)**.

---

## 3️⃣ Le ramasse-miettes (GC) — la fin de la finalisation déterministe ⭐ ⚠️

Si vous ne deviez retenir **qu'un seul** changement de fond de toute cette section, ce serait  
celui-ci. C'est **le piège n°1** de la migration.

**En VB6**, la durée de vie d'un objet repose sur le **comptage de références** (hérité de COM) :
dès que la **dernière référence** vers un objet tombe à zéro, l'objet est détruit **immédiatement**,  
de façon **déterministe** — et son `Class_Terminate` se déclenche **à cet instant précis**.

**En VB.NET**, c'est le **ramasse-miettes** (*garbage collector*, GC) du CLR qui gère la mémoire. Et
le GC décide **quand** il collecte les objets — **pas vous**. La destruction n'est plus déterministe :  
l'équivalent de `Class_Terminate` (la méthode `Finalize`) **ne se déclenche pas à un moment  
prévisible**, parfois bien après que l'objet est devenu inutile.

**Pourquoi c'est un piège silencieux.** Tout le code VB6 qui comptait sur « l'objet sort de portée
→ la ressource est libérée **maintenant** » change de comportement **sans erreur de compilation** :

- fermeture de fichiers, libération de verrous ;
- fermeture de **connexions à une base de données** ;
- libération d'objets **COM** ;
- libération de *handles* système.

Ces ressources ne sont plus libérées « tout de suite », mais « plus tard », voire « on ne sait  
quand ». Résultat : fichiers qui restent verrouillés, connexions qui s'épuisent, fuites de
*handles* — **alors que le programme compile parfaitement**.

**Concrètement.** Prenons une classe qui ouvre un fichier à sa création et le referme à sa
destruction :

```vb
' VB6 — la fermeture est DÉTERMINISTE
Dim log As FichierLog
Set log = New FichierLog    ' Class_Initialize ouvre le fichier
log.Ecrire "événement"
Set log = Nothing           ' dernière référence perdue → Class_Terminate
                            ' ferme le fichier ICI, tout de suite
```

```vb
' VB.NET — même logique migrée telle quelle : la fermeture PART À LA DÉRIVE
Dim log As New FichierLog   ' le constructeur ouvre le fichier
log.Ecrire("événement")
log = Nothing               ' l'objet devient « collectable »… mais le fichier
                            ' reste ouvert jusqu'au prochain passage du GC — quand ?
```

Le code VB.NET compile et *semble* correct. Pourtant le fichier reste verrouillé un temps  
indéterminé — le genre de bug qui ne se manifeste qu'en production, sous charge.

> ⭐ **La parade** (à internaliser dès maintenant) : pour toute libération qui doit être
> **déterministe**, .NET fournit le motif **`IDisposable` / `Using`**. Le bloc `Using` garantit la
> libération **à la fin du bloc**, de façon prévisible — c'est le remplaçant attitré du
> `Class_Terminate` déterministe de VB6.

```vb
' VB.NET — Using rétablit le déterminisme perdu
Using log As New FichierLog   ' FichierLog implémente IDisposable
    log.Ecrire("événement")
End Using                      ' Dispose() appelé ICI, à coup sûr, à la fin du bloc
```

👉 **Le détail** : section **2.2** (le problème), puis sections **12.2 et 12.3** (le motif `Dispose`,
`Using`, `Finalize`).

---

## 4️⃣ Le modèle objet — instanciation, durée de vie, événements

Au-delà du GC, c'est toute la **façon dont les objets naissent, vivent et meurent** qui évolue.

- **`Class_Initialize`** devient un **constructeur** ; **`Class_Terminate`** pose le problème de finalisation vu ci-dessus.
- L'**héritage de classe** est une **nouveauté** de VB.NET (VB6 n'offrait que l'implémentation d'interfaces via `Implements`, jamais l'héritage).
- Les **propriétés par défaut** disparaissent en grande partie, et le mot-clé **`Set`** n'existe plus : `Property Get`/`Let`/`Set` se réorganisent (avec un **faux-ami** au passage — le `Set` de VB.NET n'est pas celui de VB6).
- Le modèle d'**événements** s'enrichit : `Event`/`RaiseEvent`/`WithEvents`/`Handles` demeurent, mais `AddHandler`/`RemoveHandler` apparaissent.
- La **portée** change : `Friend` désigne désormais la portée **assembly**.

Beaucoup de ces points sont **visibles** (le code ne compile pas), mais certains sont **silencieux**
(sémantique de `Type…End Type` devenu `Structure`, comportement des événements, finalisation) — d'où
la vigilance requise.

👉 **Le détail** : chapitre **12 — Programmation orientée objet**.

---

## 🗺️ Synthèse : les quatre axes en un coup d'œil

| Axe | Ce qui change | Plutôt… | Où c'est détaillé |
|---|---|---|---|
| **Langage** | Syntaxe, mots-clés, types | **Visible** (sauf exceptions ⚠️) | Chap. 7-12, annexe A |
| **Runtime** | VBVM → CLR | **Silencieux** (invisible en source) | Section 2.1 |
| **GC** | Finalisation déterministe perdue | **Silencieux** ⭐ — le piège n°1 | Sections 2.2, 12.2-12.3 |
| **Modèle objet** | Cycle de vie, héritage, événements | **Mixte** (visible + silencieux) | Chapitre 12 |

La leçon de ce tableau : **le danger est concentré dans tout ce qui est silencieux** — le runtime,  
le GC, une partie du modèle objet, et la minorité piégeuse des changements de langage. C'est  
exactement la raison d'être du **chapitre 2** (les différences fondamentales) et de l'**annexe B**
(le catalogue des pièges silencieux), les deux ressources les plus importantes de la formation.

---

## ✅ À retenir

- La **continuité du nom** (« VB » → « VB ») masque une **rupture** : le langage se ressemble, mais le **runtime** et le **modèle objet** sont entièrement différents.
- Distinguez **deux familles de changements** : les **visibles** (syntaxe — le compilateur les attrape, donc peu dangereux) et les **silencieux** (comportement — rien ne vous alerte, donc très dangereux).
- Le **runtime change de moteur** : VBVM (COM) → CLR (.NET). Invisible dans le code, structurant en dessous.
- Le **ramasse-miettes** fait perdre la **finalisation déterministe** : c'est **le piège n°1**. La parade est le motif **`IDisposable` / `Using`**.
- Le **modèle objet** évolue (constructeurs, héritage, événements, `Set` disparu, portée `Friend`).
- **Le danger est dans le silencieux** : c'est pourquoi le **chapitre 2** et l'**annexe B** existent.

---

> 🧭 **Section suivante** → [1.4 Panorama complet d'une migration : langage, formulaires, données, COM, API](04-panorama.md)

⏭️ [Panorama complet d'une migration : langage, formulaires, données, COM, API](/01-cadrage-migration/04-panorama.md)
