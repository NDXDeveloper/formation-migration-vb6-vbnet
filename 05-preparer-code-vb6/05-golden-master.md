🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 5.5 Mettre en place un harnais de tests de référence (*golden master*) ⭐

> *Partie 2 — Préparer le terrain · Module 5 — Préparer le code VB6 avant la migration*

---

## Le problème que résout cette section

Rappel de la philosophie de cette formation : **le vrai danger n'est pas le code qui ne compile pas — c'est le code qui compile mais ne se comporte plus pareil.** Un `ByRef` devenu `ByVal`, un `Integer` qui passe de 16 à 32 bits, un arrondi monétaire qui dévie d'un centime, un `Class_Terminate` qui ne se déclenche plus au bon moment : ces changements sont **silencieux**. Ils ne lèvent aucune erreur de compilation. Ils ne se voient pas à la relecture. Ils ne se révèlent qu'en production, parfois des mois plus tard, sous la forme d'un écart de chiffre que personne ne sait expliquer.

Relire le code ligne à ligne ne suffit pas : l'œil humain ne peut pas suivre simultanément les dizaines de règles qui changent entre VB6 et VB.NET (voir le catalogue de l'**Annexe B**). Il faut un filet **comportemental**, automatique, qui réponde à une seule question :

> **« Pour les mêmes entrées, l'application migrée produit-elle exactement les mêmes sorties que l'originale ? »**

C'est exactement le rôle du **golden master**.

---

## Qu'est-ce qu'un golden master ?

Un **golden master** (aussi appelé *test de caractérisation*, *approval test* ou *snapshot test*) est une **capture figée du comportement observé** de l'application existante.

Le principe tient en quatre temps :

1. On exécute l'application **VB6** sur un jeu d'entrées fixe.
2. On **enregistre** toutes les sorties produites : c'est la **référence** (le « *master* », l'étalon doré).
3. Après migration, on exécute l'application **VB.NET** sur **exactement le même jeu d'entrées**.
4. On **compare** les nouvelles sorties à la référence. **Toute différence est un signal** : soit un bug de migration, soit un changement assumé.

```
   ENTRÉES FIXES
        │
        ├──────────────┐
        ▼              ▼
   ┌─────────┐    ┌──────────┐
   │  VB6    │    │  VB.NET  │
   │(origine)│    │ (migré)  │
   └────┬────┘    └────┬─────┘
        ▼              ▼
   reference.txt   candidat.txt
        │              │
        └──────┬───────┘
               ▼
          COMPARAISON
        identique ? ✅   →  pas de régression
        différent ? ⚠️   →  à investiguer
```

La force de l'approche : elle ne teste pas **un** piège à la fois, elle les attrape **tous d'un coup**, simplement parce qu'ils finissent par modifier une sortie observable.

---

## Pourquoi capturer la référence AVANT de migrer ?

C'est le point de **timing** le plus important de toute la démarche, et c'est pour cela que cette section est placée dans la partie « préparer le terrain » :

> ⚠️ **La référence doit être capturée sur l'application VB6 qui tourne encore, intacte.**
> Une fois la migration commencée, vous n'avez plus de source fiable pour produire l'étalon. Si vous attendez d'avoir « un peu migré » pour capturer, vous figez déjà un comportement potentiellement altéré : vous ne mesurez plus rien.

Le golden master fait partie de l'**outillage que l'on prépare en amont**, au même titre que le nettoyage du code (§ 5.1) ou le gel du périmètre (§ 5.6). Concrètement, l'ordre est : **on prépare le harnais → on capture la référence sur VB6 → on gèle → on migre → on compare**.

---

## Le principe clé : caractériser, pas spécifier

Voici la nuance que beaucoup d'équipes manquent au début.

Un test classique vérifie ce que le code **devrait** faire (sa spécification). Un golden master capture ce que le code **fait réellement**, aujourd'hui — **bugs et bizarreries compris**.

> 💡 Si l'application VB6 arrondit d'une certaine manière « discutable », vous capturez cet arrondi **tel quel**. Si elle traite un cas limite « bizarrement », vous capturez ce comportement bizarre.

Pourquoi ? Parce que l'objectif d'une migration est la **préservation du comportement**, pas son amélioration. Tant qu'on migre, le contrat est : *« l'application migrée doit se comporter comme l'ancienne »*. Corriger un bug historique est une **décision fonctionnelle distincte**, à prendre **après** la migration, en pleine conscience — pas un effet de bord involontaire glissé dans la conversion. (C'est aussi pourquoi le périmètre est gelé au § 5.6 : on ne corrige et on n'améliore pas pendant qu'on migre.)

Le golden master **opérationnalise** les comportements critiques que vous avez documentés au § 5.4 (arrondis, dates, ordre d'événements) : au lieu de les décrire dans un document, vous les figez dans des sorties comparables automatiquement.

---

## Anatomie d'un harnais golden master

Un harnais se compose toujours de trois briques :

| Brique | Rôle | Exigence principale |
|--------|------|---------------------|
| **Le jeu d'entrées** | Stimuler l'application de façon reproductible | Fixe, représentatif, versionné |
| **La capture des sorties** | Sérialiser ce que produit le code | **Déterministe** et lisible |
| **Le mécanisme de comparaison** | Détecter le moindre écart | Automatique, rejouable |

Le *driver* (programme pilote) qui relie ces briques est souvent un petit module à part : il parcourt les cas d'entrée, appelle la logique métier, et écrit chaque résultat dans un fichier au format stable. On en écrit **une version VB6** (pour la référence) et **une version VB.NET** qui produit **le même format** (pour le candidat).

---

## Que capturer ? Les bons et les mauvais candidats

Tout n'est pas également facile à mettre sous golden master. Visez les sorties **déterministes et isolables**.

| Candidat | Adapté ? | Commentaire |
|----------|:--------:|-------------|
| Calculs purs, logique métier (taxes, tarifs, intérêts, règles de gestion) | ✅✅ | Cible idéale : entrée → sortie, sans état caché |
| Génération de fichiers (exports CSV, fichiers à plat, rapports texte) | ✅✅ | Le fichier produit **est** la sortie à comparer |
| Transformations de données (parsing, formatage, conversions) | ✅ | Excellent terrain de chasse aux pièges de typage |
| États de base de données après traitement | ✅ | Possible, mais demande d'isoler/réinitialiser la base |
| Interface utilisateur (écrans, comportements de formulaire) | ⚠️ | Difficile à figer ; traité autrement (voir module 14) |
| Code fortement couplé à l'UI ou au temps réel | ❌ | À refactorer d'abord (voir ci-dessous) |

> 🔗 **Lien direct avec le § 5.3 (découpler l'UI de la logique métier).** Plus votre logique métier est séparée de l'interface, plus elle est facile à piloter directement depuis un harnais. Si toute la logique vit dans les événements de formulaires (`Command1_Click`…), commencez par l'en extraire : c'est ce découplage qui rend le golden master praticable.

---

## Le défi du déterminisme

C'est **le** point technique délicat. Pour comparer deux exécutions, leurs sorties ne doivent différer que si le **comportement métier** diffère — jamais à cause de « bruit ». Toute source de variation involontaire doit être **neutralisée**, à l'identique, des deux côtés.

| Source de non-déterminisme | Risque | Neutralisation |
|----------------------------|--------|----------------|
| Date / heure courante (`Now`, `Date`, `Time`, `Timer`) | Sortie différente à chaque exécution | **Injecter une horloge fixe** (passer la date en paramètre plutôt que d'appeler `Now` au cœur du calcul) |
| Aléatoire (`Rnd`) | Séquences différentes | Voir l'encadré ci-dessous ⚠️ |
| Identifiants générés (GUID, `IDENTITY`/auto-incrément) | Valeurs imprévisibles | Externaliser comme entrée, ou exclure de la sortie comparée |
| Ordre non garanti (collections, `SELECT` **sans** `ORDER BY`) | Mêmes données, ordre différent | Forcer un tri explicite et stable avant de sérialiser |
| Chemins, horodatages de fichiers | Chemins absolus, dates de création | **Normaliser** : retirer les chemins machines, masquer les timestamps |
| Paramètres régionaux (*culture*) | Séparateur décimal, format de date | Fixer la culture (voir encadré) — piège **B.10** |
| Multithreading / timing | Entrelacement non reproductible | Sérialiser l'exécution pendant la capture |

> ⚠️ **Cas particulier de l'aléatoire — un faux ami.**
> On pourrait croire qu'il suffit de semer les deux générateurs avec la même graine (`Randomize 12345` côté VB6, `New Random(12345)` côté .NET). **C'est faux** : `Rnd` (VBVM) et `System.Random` (CLR) reposent sur des **algorithmes différents**. Une même graine ne produit **pas** la même séquence d'un runtime à l'autre. La bonne approche est donc d'**externaliser** les tirages : ou bien on capture la suite de nombres aléatoires comme une **entrée** que l'on rejoue à l'identique des deux côtés, ou bien on isole le générateur derrière une abstraction que l'on peut remplacer par une séquence fixe. Tant que l'aléatoire reste enfoui dans le calcul, le golden master est impossible.

> ⚠️ **Cas particulier de la culture — distinguez deux choses.**
> 1. **Le formatage de la sortie de référence** doit, lui, être **stable et indépendant de la machine** : forcez une culture invariante (`InvariantCulture`) au moment d'écrire le fichier, sinon le même résultat s'affichera `1.5` ici et `1,5` ailleurs.
> 2. **Mais si la logique testée est elle-même sensible à la culture** (par ex. elle parse une saisie utilisateur avec les réglages régionaux), alors ce comportement **fait partie de ce qu'on doit tester** : exercez-le explicitement sous les cultures pertinentes. Ne confondez pas « stabiliser l'affichage du dump » et « masquer un comportement métier réel ».

---

## Le format de la référence

Choisissez un format **texte, lisible et comparable ligne à ligne** :

- **Texte brut, CSV ou JSON** — surtout pas de binaire : un diff doit être lisible par un humain.
- **Une valeur étiquetée par ligne** — chaque ligne se suffit à elle-même (`Libellé = valeur`), pour qu'un écart pointe immédiatement le cas fautif.
- **Stocké dans le gestionnaire de version (Git)** — la référence vit avec le code. On voit son historique, on voit chaque modification proposée, on peut la relire en revue.
- **Auto-documenté** — un fichier de référence qu'on peut lire et comprendre sans lancer le programme est un fichier qu'on saura maintenir.

> 💡 Un bon fichier de référence se lit comme un compte rendu. Mauvais : `42|3.14|-1|X`. Bon : une ligne par cas, libellée, formatée à la précision qui compte pour le métier.

---

## Mettre en œuvre : capturer puis comparer

### Côté VB6 — produire la référence (avant migration)

L'idée : un petit *driver* qui parcourt les cas, appelle la fonction métier réelle, et écrit chaque résultat avec un **format figé**.

```vb
' === Harnais de capture — à exécuter sur l'application VB6 INTACTE ===
Option Explicit

Public Sub CapturerReference()
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\golden\reference.txt" For Output As #f

    ' Jeu d'entrées : cas nominaux + cas limites + (idéalement) données de production
    EcrireCas f, "TVA standard",     CalculerTVA(100#, 0.2)
    EcrireCas f, "TVA taux zero",    CalculerTVA(100#, 0#)
    EcrireCas f, "TVA arrondi fin",  CalculerTVA(99.99, 0.196)
    EcrireCas f, "Montant negatif",  CalculerTVA(-50#, 0.2)
    EcrireCas f, "Montant nul",      CalculerTVA(0#, 0.2)
    ' ... des centaines de cas : c'est la couverture qui fait la valeur (voir plus bas)

    Close #f
End Sub

Private Sub EcrireCas(ByVal f As Integer, ByVal libelle As String, ByVal resultat As Currency)
    ' Format stable et explicite : 4 décimales, point décimal forcé
    Print #f, libelle & " = " & Format$(resultat, "0.0000")
End Sub
```

Points d'attention sur cet exemple :
- La fonction `CalculerTVA` est **isolable** (entrée → sortie, sans dépendre de l'UI ni de l'heure). C'est ce qui la rend testable.
- Le type `Currency` (monétaire, 4 décimales) est volontairement préservé : c'est précisément là que les écarts d'arrondi se logent après migration.
- Le format `"0.0000"` fige la précision **et** force le point décimal, indépendamment des réglages régionaux de la machine de capture.

Le fichier `reference.txt` produit ressemble alors à ceci — une **ligne lisible par cas**, si bien  
qu'un écart pointe immédiatement le cas fautif (ici, `CalculerTVA` renvoie le montant de TVA) :

```text
TVA standard = 20.0000
TVA taux zero = 0.0000
TVA arrondi fin = 19.5980
Montant negatif = -10.0000
Montant nul = 0.0000
```

C'est cet étalon — figé, versionné, relu en revue — que la version migrée devra reproduire
**à l'identique**.

### Côté VB.NET — produire le candidat (après migration), au même format

```vbnet
' === Même harnais, côté VB.NET — APRÈS migration ===
Option Strict On

Module GoldenMaster
    Public Sub CapturerCandidat()
        Using ecrivain As New IO.StreamWriter("golden\candidat.txt", append:=False)
            EcrireCas(ecrivain, "TVA standard",     CalculerTVA(100D, 0.2D))
            EcrireCas(ecrivain, "TVA taux zero",    CalculerTVA(100D, 0D))
            EcrireCas(ecrivain, "TVA arrondi fin",  CalculerTVA(99.99D, 0.196D))
            EcrireCas(ecrivain, "Montant negatif",  CalculerTVA(-50D, 0.2D))
            EcrireCas(ecrivain, "Montant nul",      CalculerTVA(0D, 0.2D))
        End Using
    End Sub

    Private Sub EcrireCas(ecrivain As IO.StreamWriter, libelle As String, resultat As Decimal)
        ' Culture invariante : la sortie ne dépend pas des réglages de la machine
        Dim valeur As String = resultat.ToString("0.0000", Globalization.CultureInfo.InvariantCulture)
        ecrivain.WriteLine($"{libelle} = {valeur}")
    End Sub
End Module
```

> ⚠️ Notez la traduction de type **`Currency` (VB6) → `Decimal` (VB.NET)** : c'est l'équivalent monétaire correct (et **non** `Double`, qui introduirait des erreurs de représentation binaire). Voir l'**Annexe D** (correspondance des types) et le piège **B.9**.

### Comparer

La comparaison la plus simple est un **diff de deux fichiers texte** :

```
fc golden\reference.txt golden\candidat.txt      (Windows)
diff golden/reference.txt golden/candidat.txt    (Unix / Git Bash)
```

Ou, en revue, le **diff de Git** lui-même : on remplace `reference.txt` par la sortie du candidat et l'on regarde ce qui change. **Aucune ligne modifiée = aucune régression** sur les cas couverts. Pour automatiser, un petit comparateur qui lit les deux fichiers et signale la **première ligne divergente** (libellé + valeur attendue + valeur obtenue) suffit à intégrer le golden master dans une exécution répétable. La mécanique complète d'intégration au cycle de validation est détaillée au **§ 17.3**.

---

## Comparaison exacte ou tolérante ?

Question récurrente sur les valeurs numériques : faut-il exiger l'**égalité stricte** ou tolérer un petit écart ?

- **Par défaut : égalité stricte.** Pour tout ce qui est **monétaire**, c'est non négociable — détecter une dérive d'un centième est *l'objectif même* de l'exercice. Une tolérance reviendrait à aveugler le test précisément là où il est le plus utile.
- **Pour certains flottants non monétaires** (calculs scientifiques accumulant des opérations), une **tolérance contrôlée** peut être légitime, mais alors : gardez-la **serrée**, **documentez-la**, et sachez qu'une tolérance trop large **cache** les vraies régressions.
- **Meilleure pratique :** plutôt qu'une tolérance floue, **formatez à la précision qui compte pour le métier** (ici, 4 décimales) puis comparez **exactement** cette représentation. On compare alors ce qui a un sens fonctionnel, pas le 17ᵉ chiffre après la virgule.

---

## La couverture : le vrai facteur limitant

Le golden master ne teste **que ce que vous lui donnez à manger.** Un harnais qui ne couvre que le cas nominal validera une migration… qui casse sur tous les cas limites.

La valeur de votre filet dépend donc directement de la **diversité des entrées** :

- **Cas nominaux** — les flux métier courants.
- **Cas limites** — zéro, valeurs négatives, vide / `Null`, valeurs maximales, dates de bascule (fins de mois, années bissextiles, 29 février), chaînes avec caractères spéciaux ou accentués.
- **Données de production anonymisées** — **l'étalon-or de la couverture.** Rien ne remplace de vrais volumes, avec leurs irrégularités que personne n'aurait imaginées. (Anonymisez : on capture du comportement, pas des données personnelles.)

> 💡 Un repère utile : mesurez **quelle proportion des chemins de code** votre jeu d'entrées exerce réellement. Une fonction jamais appelée par le harnais est une fonction dont la migration n'est **pas** validée — même si le diff est tout vert.

---

## Le workflow complet

1. **Préparer** le harnais (driver + format de sortie) pendant la phase de préparation.
2. **Capturer** la référence sur l'application **VB6 intacte** → `reference.txt`.
3. **Geler** : committer la référence dans le gestionnaire de version (§ 5.6).
4. **Migrer** le code vers VB.NET.
5. **Rejouer** le même harnais côté VB.NET → `candidat.txt`.
6. **Comparer** `candidat.txt` à `reference.txt`.
7. **Pour chaque écart**, trancher :
   - **Bug de migration** → on corrige le code migré, on rejoue.
   - **Changement assumé** → on documente, et on **met à jour la référence** (voir ci-dessous).
8. **Répéter** jusqu'à n'avoir plus que des écarts intentionnels et justifiés.

---

## Gérer les écarts « légitimes »

Tout écart n'est pas un bug. Parfois la migration change **volontairement** un comportement : on corrige enfin un bug historique, on adopte un traitement de culture plus correct, on aligne un arrondi sur la règle officielle. Dans ce cas, le nouveau résultat est **le bon**.

La discipline d'*approval testing* consiste alors à **« bénir » la nouvelle sortie** : on inspecte le diff, on confirme que le changement est voulu, on **met à jour le fichier de référence** avec la nouvelle valeur, et — surtout — **on documente pourquoi** (idéalement dans le message de commit qui modifie la référence).

> ⚠️ La règle d'or : **une référence ne se met jamais à jour « pour faire passer le test ».** On ne la met à jour qu'après avoir **compris et validé consciemment** l'écart. Bénir un diff sans le comprendre, c'est transformer le filet de sécurité en tampon-encreur qui valide n'importe quoi.

---

## Les limites du golden master

Pour rester lucide — c'est une valeur de cette formation :

- **Il prouve l'équivalence, pas la correction.** Il garantit que le code migré se comporte comme le VB6 — **bugs d'origine inclus**. Ce n'est pas un test de validité fonctionnelle.
- **Il ne vaut que par sa couverture.** Ce qui n'est pas exercé n'est pas protégé.
- **Il est fragile face aux sorties bruitées.** Un déterminisme mal maîtrisé produit des faux positifs qui finissent par être ignorés — et un test qu'on ignore ne sert plus à rien.
- **L'UI s'y prête mal.** Les comportements d'interface se valident par d'autres moyens (module 14).
- **Il a un coût d'entretien.** Le harnais et les références sont du code à maintenir.

Cela étant dit : pour une **migration**, c'est le filet de sécurité au **meilleur rapport effort/protection** qui soit. Il transforme une question angoissante (« est-ce qu'on a cassé quelque chose, quelque part ? ») en une réponse **automatique et reproductible**.

---

## À retenir

> - Un **golden master** capture le **comportement observé** de l'application VB6, puis vérifie que la version migrée produit **les mêmes sorties pour les mêmes entrées**.
> - Il **caractérise** (ce que le code fait) plutôt qu'il ne **spécifie** (ce qu'il devrait faire) : on préserve le comportement, **bugs compris**.
> - La référence se capture **impérativement avant** la migration, sur l'application **intacte**, puis se **gèle**.
> - Le succès dépend de deux choses : le **déterminisme** des sorties (neutraliser dates, aléatoire, ordre, culture…) et la **couverture** des entrées (cas limites + données de production anonymisées).
> - Pour le monétaire : **comparaison stricte**, jamais de tolérance.
> - C'est le filet qui attrape **tous les pièges silencieux d'un coup** (ByRef, types redimensionnés, arrondis, finalisation…) — le meilleur rapport effort/protection de la migration.

---

## Liens utiles

- **§ 5.3** — Découpler l'UI de la logique métier *(prérequis pour piloter le métier directement)*
- **§ 5.4** — Documenter les comportements critiques *(que le golden master fige ensuite)*
- **§ 5.6** — Geler le périmètre *(pour que la référence reste valide)*
- **§ 17.3** — Tests de non-régression : comparer avec le *golden master* *(la mise en œuvre côté validation)*
- **§ 17.4** — Traquer les pièges silencieux *(ce que le filet révèle)*
- **Annexe B** — Catalogue des pièges silencieux *(ce que l'on cherche à attraper)*
- **Annexe F** — Modèles de tests de non-régression (*golden master*) *(patrons réutilisables)*

⏭️ [Geler le périmètre (arrêter les évolutions fonctionnelles pendant la migration)](/05-preparer-code-vb6/06-geler-perimetre.md)
