🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 🧹 5.1 — Nettoyer le code VB6

**Module 5 — Préparer le code VB6 avant la migration** · Section 5.1
Cible : .NET Framework 4.7.2

> 🧹 **En une phrase** : avant de déménager, on jette ce qu'on ne transporte pas et on étiquette
> les cartons. Nettoyer, c'est réduire la quantité de code à migrer **et** lever les ambiguïtés que
> le convertisseur — outil ou humain — interpréterait de travers.

---

## 🧭 De quoi parle cette section

C'est l'étape la plus **mécanique** du module : l'hygiène de base, à faire pendant que le code VB6  
compile et tourne encore. Trois chantiers, du plus simple au plus subtil :

1. **`Option Explicit` partout** — forcer la déclaration de toutes les variables ;
2. **Supprimer le code mort** — ce qu'on ne migrera pas, autant ne pas le migrer ;
3. **Supprimer le `Variant` superflu** — typer ce qui n'a jamais eu besoin d'être polymorphe.

> ⭐ **Distinction importante avec la [section 5.2](02-reduire-pieges-amont.md)** : ici, on fait le
> *ménage* (retirer les `Variant` manifestement inutiles, déclarer ce qui ne l'était pas). La 5.2
> va plus loin et de façon **systématique** : choisir le bon type pour *chaque* variable et rendre
> chaque passage de paramètre explicite (`ByVal`/`ByRef`). 5.1 dégrossit ; 5.2 affine et sécurise.

> ⚠️ **Filet de sécurité** : nettoyer modifie le code. Le risque de régression est faible mais pas
> nul (voir les fausses pistes plus bas). Travaillez par **petites passes**, en recompilant et en
> relançant l'application après chaque lot. Idéalement, le harnais de référence
> ([§5.5, *golden master*](05-golden-master.md)) est déjà en place — même s'il apparaît plus loin
> dans le module, en pratique on veut ce filet **le plus tôt possible**.

---

## 1️⃣ `Option Explicit` partout

### Le problème qu'il résout

En VB6, **sans** `Option Explicit`, utiliser un nom de variable jamais déclaré ne provoque aucune  
erreur : VB6 **crée la variable à la volée**, comme un `Variant` vide. Conséquence : la moindre  
faute de frappe fabrique silencieusement une nouvelle variable, et le bug ne se voit pas.

```vb
' Sans Option Explicit — le piège classique
Sub CalculerTotal()
    montantHT = 100
    tva = montantHT * 0.2
    montantTTC = montantHT + tav   ' faute de frappe : "tav" au lieu de "tva"
    ' "tav" est créée à la volée comme Variant vide (= 0)
    ' montantTTC vaut 100 au lieu de 120 — et AUCUNE erreur n'est signalée
End Sub
```

Avec `Option Explicit` en tête de module, le compilateur s'arrête net sur `tav` : *variable non  
déclarée*. Le bug devient visible **avant** la migration, dans l'environnement où on sait encore ce  
qui est correct.

### Pourquoi le faire maintenant

- **VB.NET impose `Option Explicit On` par défaut.** Autant régulariser pendant qu'on dispose de
  l'oracle vivant qu'est l'application VB6.
- Chaque variable « fantôme » débusquée est soit une faute de frappe (un vrai bug latent), soit une
  variable réelle qu'on va enfin **déclarer** — et donc pouvoir typer correctement (ce qui prépare
  la 5.2).

### Comment procéder

1. Activez **Tools → Options → onglet Editor → « Require Variable Declaration »**. ⚠️ Attention :
   cette option insère `Option Explicit` automatiquement dans les **nouveaux** modules seulement —
   elle ne touche **pas** au code existant.
2. Ajoutez donc **à la main** la ligne `Option Explicit` en tête de **chaque** module existant
   (`.bas`, `.cls`, `.frm`). C'est l'unique façon de couvrir tout le projet — d'où le « partout ».
3. Recompilez (**Debug → Compile**). Le compilateur s'arrête à la première variable non déclarée.
   Corrigez, recompilez, recommencez. Sur un gros projet, avancez **module par module**.

> 💡 `Option Explicit` est purement un **contrôle à la compilation** : sur un module qui passe déjà
> proprement, l'ajouter ne change rien au comportement. Le seul cas où « ça bouge », c'est quand le
> code **dépendait** d'une variable implicite — et alors la correction révèle un bug réel. C'est
> exactement ce qu'on veut traiter ici, pas en VB.NET.

---

## 2️⃣ Supprimer le code mort

### Ce qu'on vise

Tout ce qui ne sert plus mais alourdit la migration :

- **procédures, fonctions, classes ou modules** jamais appelés ;
- **variables, constantes, énumérations** déclarées et inutilisées ;
- **blocs commentés** laissés « au cas où » (versions précédentes d'un algorithme, essais) ;
- **code inaccessible**, situé après un `Exit Sub`/`Exit Function`, un `GoTo` ou une condition
  toujours fausse ;
- **formulaires ou contrôles** présents dans le projet mais plus utilisés.

```vb
Sub Traiter()
    Exit Sub
    MsgBox "Jamais affiché"   ' code mort : inaccessible après Exit Sub
End Sub
```

### Pourquoi c'est rentable avant migration

Chaque ligne migrée a un **coût** (effort) et un **risque** (régression possible). Migrer du code  
mort, c'est payer ce coût et ce risque pour rien. Pire : le code mort **brouille la lecture** des  
outils automatiques et de l'IA (module 19), qui peuvent s'épuiser à convertir fidèlement ce que  
vous alliez jeter. Et il dilue la compréhension du comportement réel de l'application.

### ⚠️ Attention aux « faux morts »

Du code peut **paraître** inutilisé sans l'être. Vérifiez avant de supprimer :

- les membres appelés par **liaison tardive** ou par leur **nom** (`CallByName`, dispatch
  dynamique) — invisibles pour une recherche de références classique ;
- les membres **`Public`** d'un composant **consommé par d'autres applications** (clients COM) :
  « inutilisé dans ce projet » ne veut pas dire « inutilisé tout court » ;
- le code sous **compilation conditionnelle** (`#If…#End If`) : mort sous les drapeaux actuels,
  vivant sous d'autres ;
- les gestionnaires d'événements câblés ailleurs que dans le code visible.

> 🛠️ **Outillage** : l'IDE VB6 ne liste pas les procédures inutilisées. Des analyseurs dédiés
> (type *Project Analyzer*, *MZ-Tools*) repèrent code mort, variables et constantes inutilisées,
> et procédures jamais appelées. Voir le [module 4 — Outils de migration](../04-outils-migration/README.md).
> Traitez leurs résultats comme des **pistes**, pas comme des ordres : la décision de supprimer
> reste humaine, surtout pour les membres publics.

---

## 3️⃣ Supprimer le `Variant` superflu

### Le `Variant`, ce qu'il est en VB6

C'est le type **fourre-tout** : il accepte n'importe quelle valeur et, en prime, des **états  
spéciaux** propres à VB6 — `Empty`, `Null`, `Missing`, `Error`. C'est aussi le type **par défaut**  
quand on ne précise rien (et le type des variables implicites, d'où le lien avec `Option Explicit`  
ci-dessus).

### Pourquoi le `Variant` superflu pose problème à la migration

À la bascule, **`Variant` devient `Object`** ([§7.1](../07-types-variables/01-variant-object.md)),
avec deux ennuis :

- les **états spéciaux** (`Empty`/`Null`/`Missing`) **ne se traduisent pas proprement** en .NET —
  c'est l'un des pièges silencieux catalogués (voir
  [Annexe B.5](../annexes/pieges-silencieux/README.md) et
  [§2.3](../02-differences-fondamentales/03-typage.md),
  [§7.7](../07-types-variables/07-nothing-empty-null.md)) ;
- chaque `Variant` entretient une **conversion/boxing** et une incertitude de type que l'outil de
  migration doit deviner.

Moins il reste de `Variant` *inutiles* au moment de migrer, moins il y a de surface où ces pièges  
peuvent surgir.

### Le critère : superflu ≠ nécessaire

L'objectif n'est **pas** d'éradiquer tout `Variant`, mais de retirer ceux qui n'ont **jamais** eu  
besoin de l'être :

```vb
' Variant par paresse — alors qu'on ne manipule qu'un entier
Dim compteur As Variant
compteur = 0
' ... compteur ne contiendra jamais qu'un nombre

' ✅ Mieux : un type explicite (le choix précis du type est traité en 5.2 et au module 7)
Dim compteur As Long
```

| `Variant` **superflu** (à typer) | `Variant` **légitime** (à conserver pour l'instant) |
|----------------------------------|------------------------------------------------------|
| Ne contient toujours qu'un nombre, qu'une chaîne, qu'une date | Données réellement polymorphes (contenu hétérogène assumé) |
| Né d'une déclaration paresseuse ou implicite | Retours d'API/COM **typés `Variant`** (interop) |
| N'utilise jamais `Empty`/`Null`/`Missing` | Cas qui **s'appuient** sur ces états (à traiter, pas à improviser) |

> ⭐ Ici, on retire les `Variant` **évidemment** inutiles. Le passage au crible *systématique* de
> chaque variable — et le choix du bon type cible (sachant que `Integer`→`Short`, `Long`→`Integer`,
> etc.) — relève de la [section 5.2](02-reduire-pieges-amont.md) et du
> [module 7](../07-types-variables/README.md). N'allez pas plus loin que le ménage à ce stade.

---

## 🔁 Un ordre de passage qui marche

Les trois chantiers se renforcent dans cet enchaînement :

1. **`Option Explicit` d'abord.** Forcer la déclaration fait remonter les variables fantômes et les
   fautes de frappe : c'est déjà du nettoyage, et ça révèle ce qu'il faudra typer.
2. **Code mort ensuite.** Inutile de déclarer, typer ou relire ce qu'on va supprimer : on réduit
   d'abord le volume.
3. **`Variant` superflu enfin.** Sur le code restant — vivant et explicitement déclaré — on resserre
   les types les plus manifestes.

À chaque étape : **petites passes → recompilation → test contre l'application de référence**.

---

## ⚠️ Les pièges de cette section

- **Tout vouloir parfaire.** Nettoyer n'est pas réécrire. On vise un code *prêt à migrer*, pas un
  code idéal : on va de toute façon le transformer.
- **Supprimer trop vite.** Les « faux morts » (liaison tardive, membres publics consommés ailleurs,
  compilation conditionnelle) doivent être vérifiés avant la suppression.
- **Confondre 5.1 et 5.2.** Si vous vous surprenez à arbitrer finement *quel* type entier choisir ou
  à basculer des paramètres en `ByVal`, vous êtes déjà dans la 5.2 — notez-le et gardez-le pour là.
- **Nettoyer sans filet.** Sans recompilation fréquente ni *golden master*, une régression
  introduite par le ménage se découvrira trop tard.

---

## 🔗 Pour aller plus loin

- **Suite logique** : [5.2 — Réduire les pièges en amont](02-reduire-pieges-amont.md) ⭐ (typage
  explicite systématique + `ByVal` explicite côté VB6).
- **Filet de sécurité** : [5.5 — Harnais de tests de référence (*golden master*)](05-golden-master.md).
- **Le « pourquoi » des `Variant`** : [Annexe B.5](../annexes/pieges-silencieux/README.md),
  [§2.3](../02-differences-fondamentales/03-typage.md),
  [§7.1](../07-types-variables/01-variant-object.md),
  [§7.7](../07-types-variables/07-nothing-empty-null.md).
- **`Option Explicit` côté .NET** (et ses cousins `Strict`/`Infer`) :
  [§6.4](../06-preparer-environnement/04-options-compilation.md).
- **Outils d'analyse** : [module 4](../04-outils-migration/README.md).

---

> ✅ **À retenir** : `Option Explicit` partout pour rendre le code **honnête** (plus de variables
> fantômes), suppression du code mort pour migrer **moins**, retrait des `Variant` superflus pour
> migrer **plus clair**. Trois gestes simples, faits dans l'environnement stable de VB6, qui
> retirent autant d'occasions de régression silencieuse au moment de la bascule.

---

**Section précédente** : [5. Préparer le code VB6 (introduction)](README.md)  
**Section suivante** : [5.2 — Réduire les pièges en amont →](02-reduire-pieges-amont.md)

⏭️ [Réduire les pièges **en amont** (typage explicite, `ByVal` explicite côté VB6)](/05-preparer-code-vb6/02-reduire-pieges-amont.md)
