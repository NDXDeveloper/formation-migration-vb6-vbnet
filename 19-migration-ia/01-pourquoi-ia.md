🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.1 Pourquoi l'IA aide pour VB6 → VB.NET (et ses limites) ⭐

> 🧭 *Module 19 — Migrer avec l'aide de l'IA · Section 1 / 5*

Avant d'apprendre à **piloter** l'IA (section 19.2), il faut comprendre **pourquoi** elle est utile sur cette migration précise — et **où** elle cesse de l'être. Ce n'est pas une question de mode : l'IA a des forces réelles et des angles morts tout aussi réels, et les deux découlent directement de la **nature** de la tâche VB6 → VB.NET. L'objet de cette section est de dresser une **carte honnête** : que peut-on raisonnablement déléguer, et que doit-on absolument garder sous contrôle humain ?

> ⭐ **La thèse de cette section, en une ligne :** l'IA est **excellente sur ce qui est régulier et explicite**, et **dangereuse sur ce qui est caché et global** — or les pires risques de cette migration sont justement **cachés**.

---

## Pourquoi cette migration est « favorable » à l'IA

Plusieurs caractéristiques de la tâche jouent en faveur de l'IA :

- **Deux dialectes proches.** VB6 et VB.NET appartiennent à la **même famille** : la conversion est en grande partie une **traduction entre langages voisins**, riche en motifs répétitifs.
- **Des langages anciens et abondamment documentés.** Le corpus d'entraînement des modèles regorge d'exemples, de discussions et de guides de portage. Les transformations courantes (mots-clés, fonctions intrinsèques, structures de contrôle) sont **connues et stéréotypées**.
- **Un gros volume mécanique et régulier.** C'est exactement le profil de tâche où un modèle de langage est performant.
- **Une grande aptitude à lire et expliquer du code.** Un atout décisif face à un *legacy* peu ou pas documenté.

---

## Là où l'IA aide vraiment

### 1. La conversion mécanique de masse
Traduction de syntaxe et de mots-clés, réécriture de déclarations, transposition de structures, mise en forme. C'est le **travail ingrat** — les 60-90 % de l'effort de conversion évoqués dès l'introduction de la formation — que l'IA abat vite. À condition de **relire** : « vite » ne veut pas dire « juste ».

### 2. L'explication de code VB6 obscur — l'atout sous-estimé
Donnez à l'IA une routine cryptique de 1998 et demandez-lui ce qu'elle **fait** : effets de bord, **propriétés par défaut** implicites, ordre des événements, intention probable. Sur du *legacy* dont l'auteur est parti depuis longtemps, cette capacité de **narration du code** vaut souvent plus que la conversion elle-même. (On y revient en 19.2.)

### 3. La génération d'échafaudage
Squelettes de **tests de non-régression**, **documentation XML** (`''' <summary>`), stubs, commentaires. Une matière première utile — toujours **à valider**, jamais à prendre pour argent comptant (détaillé en 19.4).

### 4. La suggestion d'idiomes modernes
Passer de `On Error` à `Try/Catch`, proposer `Using`/`IDisposable` là où VB6 comptait sur `Class_Terminate`, esquisser un `DllImport` à partir d'un `Declare`. L'IA connaît les **patrons de portage** et les rappelle volontiers.

### 5. Le rappel de correspondances
« Quel est l'équivalent .NET de telle fonction VB6 ? » L'IA fait gagner du temps comme **aide-mémoire** — à recouper avec l'[Annexe A](../annexes/correspondance-vb6-vbnet/README.md), qui, elle, ne se trompe pas.

> 💡 **Fil conducteur des forces :** l'IA est forte quand la réponse est **déjà connue, régulière et largement documentée**.

---

## Là où l'IA atteint ses limites

C'est ici que la lucidité de cette formation prend tout son sens. Les limites de l'IA recouvrent **presque exactement** la liste des pièges de l'[Annexe B](../annexes/pieges-silencieux/README.md).

### 1. Les pièges silencieux — l'angle mort majeur
L'IA raisonne sur la **syntaxe de surface**, pas sur la **sémantique d'exécution**. Elle ne « sent » pas qu'un `ByRef` devenu `ByVal` casse une routine qui modifiait son argument, qu'un `Integer` redimensionné déborde, qu'un `Class_Terminate` ne se déclenche plus de façon déterministe, qu'un tableau passe en **0-based**, ou qu'une conversion `CDbl`/`CDate` devient **sensible à la culture**. Résultat : du code qui **compile et paraît correct**, mais ne se comporte plus pareil — la première cause de régressions, produite avec assurance.

Exemple typique : une procédure VB6 qui s'appuie sur le passage **par référence** implicite.

```vb
' VB6 — ByRef est IMPLICITE : la procédure modifie réellement « total »
Sub Ajouter(montant As Currency, total As Currency)
    total = total + montant
End Sub
```

Convertie « littéralement », l'IA peut produire une signature **`ByVal` par défaut** (la nouvelle norme VB.NET) : le code compile, mais `total` n'est **plus** mis à jour chez l'appelant. Aucun message d'erreur — juste un résultat faux (voir Annexe B.1).

### 2. Les faux équivalents
L'IA propose volontiers un « équivalent » **plausible mais sémantiquement différent** : une méthode .NET dont le comportement aux **cas limites** diffère (arrondi monétaire, gestion des `Null`/`Empty`, format de date), ou une correspondance qui « marche » sur l'exemple mais trahit un invariant ailleurs. C'est l'objet central de la vigilance du module (repris en 19.5).

### 3. Les hallucinations
Méthodes, surcharges ou espaces de noms **inventés** mais crédibles. Pire dans ce contexte : l'IA peut recommander des **béquilles dépréciées** comme l'espace `Microsoft.VisualBasic.Compatibility` — précisément ce que la formation vous apprend à **éviter** ([section 4.2](../04-outils-migration/02-compatibility-namespace.md)).

### 4. L'absence de contexte global
L'IA ne voit que **ce que vous lui fournissez**, pas votre solution entière : dépendances inter-modules, références **COM/ActiveX**, *control arrays*, invariants à l'échelle du projet. Elle ne peut pas raisonner sur ce qu'elle ne voit pas — d'où des conversions localement correctes mais globalement fausses.

### 5. La confusion VB6 / VB.NET
Sans **désambiguïsation explicite**, l'IA mélange les deux dialectes et propose des tournures qui n'existent que dans l'un ou l'autre. C'est exactement pourquoi 19.2 impose de **toujours** préciser « VB6 » vs « VB.NET » **et** la cible **.NET Framework 4.7.2**.

### 6. La cible qui dérive vers le « .NET moderne »
Spontanément, l'IA penche pour les idiomes et API du **.NET récent (8/10)** : elle peut suggérer des méthodes absentes de **4.7.2**, ou des choix qui relèvent du **second saut**, distinct, traité au [module 20](../20-apres-net-moderne/README.md). Mélanger les deux sauts est exactement ce que la formation demande d'éviter.

### 7. L'UI et les *control arrays*
Concepteur Windows Forms, spécificités des `.frm`, *control arrays* sans équivalent direct : un terrain **mal maîtrisé** par les outils en général, IA comprise.

### 8. Le non-déterminisme
Même *prompt*, réponses **différentes** d'une fois sur l'autre. Aucune garantie de **cohérence** sur l'ensemble d'une base de code volumineuse — un problème pour une migration qui doit rester **reproductible** et **auditable**.

---

## Pourquoi l'IA se trompe (en bref)

Un modèle de langage **prédit du texte plausible** à partir de motifs. Il n'a **pas de modèle d'exécution**, **pas d'accès à votre *runtime***, **pas de vue d'ensemble** de votre projet, et il optimise une réponse **convaincante**, pas nécessairement **correcte**. D'où un profil très net : **fort sur la régularité syntaxique, faible sur la sémantique cachée**. Or, dans cette migration, le risque vit précisément dans la **sémantique cachée**.

> ⚠️ **Le constat à graver :** *l'IA est la plus faible là où cette migration est la plus dangereuse.*
> C'est pourquoi aucune sortie d'IA n'est validée sans **confrontation au *golden master*** ([module 5.5](../05-preparer-code-vb6/05-golden-master.md), [Annexe F](../annexes/golden-master/README.md)) ni **relecture à la lumière de l'Annexe B**.

---

## Le bon partage des rôles

| Tâche | Confier à l'IA ? | Pourquoi |
|-------|------------------|----------|
| Conversion mécanique de masse | ✅ Oui, **avec relecture** | régulière, bien documentée |
| Expliquer du VB6 obscur | ✅ Oui (point fort) | narration de code |
| Générer tests / documentation | ✅ Oui, **à valider** | échafaudage utile |
| Suggérer un idiome moderne | 🟡 Oui, **à vérifier** | bons patrons, mais faux équivalents possibles |
| Trancher un **piège silencieux** | ❌ Non | jugement humain + *golden master* |
| Décider **architecture / périmètre** | ❌ Non | hors de portée de l'IA |
| Gérer **UI / *control arrays*** | ❌ Non (ou forte supervision) | mal maîtrisé |

---

## Une précaution préalable : où va votre code ?

Avant même de coller la moindre ligne dans un assistant, une question de **gouvernance** se pose : le code d'un *legacy* d'entreprise est un **actif propriétaire**, et l'envoyer à une IA **hébergée dans le cloud** revient à le faire **sortir** de l'organisation.

- **Confidentialité et propriété intellectuelle** — code métier, algorithmes, secrets : vérifiez ce que le fournisseur fait des données soumises (**rétention**, **entraînement** éventuel sur vos échanges).
- **Données personnelles (RGPD)** — chaînes, jeux d'essai ou commentaires peuvent contenir des données réelles, à **ne pas** transmettre sans base légale.
- **Politique interne** — beaucoup d'organisations **encadrent** ou **interdisent** l'usage d'IA externes sur le code source.

**Parades** : privilégier une offre **« entreprise »** garantissant la **non-rétention** et la **non-réutilisation** des données ; ou une IA **auto-hébergée / locale** (le code ne quitte pas le réseau) ; à défaut, **anonymiser** (retirer secrets, identifiants et données réelles) avant de soumettre. En un mot : **connaître et respecter la politique de l'organisation avant tout usage.**

---

## À retenir

L'IA est un **multiplicateur de force** sur ce qui est régulier et explicable, et un **passif** sur ce qui est silencieux et global. Le bon usage découle directement de cette carte : lui confier le **mécanique** et l'**explication**, et garder pour l'humain — armé du *golden master* et de l'Annexe B — tout ce qui touche au **comportement caché** et aux **décisions de fond**.

La section suivante passe de la théorie à la pratique : comment **rédiger des prompts** qui maximisent ces forces et **contiennent** ces faiblesses.

---

> 🧭 **Navigation**  
> ⬅️ [19. Migrer avec l'aide de l'IA — introduction](README.md)  
> ⬆️ [Sommaire de la formation](../SOMMAIRE.md)  
> ➡️ Suivant : [19.2 — Prompting : convertir, expliquer](02-prompting.md)

---

**Juin 2026** · Cible : .NET Framework 4.7.2 · Licence : Creative Commons BY-NC-SA 4.0

⏭️ [Prompting : convertir un module/une classe, **expliquer** un comportement VB6 obscur](/19-migration-ia/02-prompting.md)
