🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2.1 Deux runtimes, deux mondes (VBVM vs CLR)

> **Chapitre 2 — Différences fondamentales VB6 ↔ VB.NET**
> La rupture la plus profonde de la migration — et la plus invisible, car elle n'apparaît **nulle part** dans votre code source.

---

## Le changement qu'on ne voit pas, mais dont tout découle

On a dit au chapitre 1 que migrer, c'est changer **trois choses** : le langage, le runtime et le  
modèle objet. Le langage, on le **voit** (les mots-clés changent). Le modèle objet, on finit par le  
sentir. Mais le **runtime** — le moteur qui exécute le programme — ne se montre **jamais** dans le  
code : aucune ligne ne mentionne « VBVM » ou « CLR ». Et pourtant, **c'est de lui que découlent les  
conséquences les plus structurantes** de toute la migration.

C'est donc le **changement silencieux par excellence** : invisible à la lecture, déterminant à  
l'exécution. Cette section explique ce qu'est ce moteur des deux côtés, en quoi les deux diffèrent,  
et — surtout — **pourquoi cette différence engendre tous les pièges** détaillés dans le reste du  
chapitre.

> 💡 **Une image pour commencer.** Le langage, c'est la **langue** que vous parlez. Le runtime,
> c'est le **pays** où vous la parlez. Migrer de VB6 à VB.NET, ce n'est pas seulement traduire les
> phrases : c'est changer de pays — avec d'autres lois, d'autres habitudes, une autre gestion du
> quotidien. Les mots se ressemblent ; les **règles du lieu**, non.

---

## 🏛️ D'abord, qu'est-ce qu'un « runtime » ?

Un **runtime** (moteur d'exécution) est l'ensemble des services dont un programme a besoin **pendant  
qu'il tourne** : créer et détruire des objets, gérer la mémoire, convertir les types, traiter les  
erreurs, dialoguer avec le système. Votre code ne fait pas tout cela seul — il **délègue** au  
runtime.

C'est précisément pour cela que le changement de runtime est si lourd de conséquences : on ne change  
pas qu'un dialecte, **on change le prestataire qui assure tous ces services en coulisses** — et il  
ne les assure pas de la même façon.

---

## 🔵 Côté VB6 : le VBVM, un monde bâti sur COM

Les applications VB6 s'exécutent avec l'aide du **VBVM** — *Visual Basic Virtual Machine* —, livré  
sous la forme de la bibliothèque **`msvbvm60.dll`**.

Quelques traits essentiels :

- **Le nom « Virtual Machine » est trompeur.** Le VBVM n'est pas une machine virtuelle au sens de Java : VB6 compile généralement vers du **code natif** (ou, en option, du *P-code* interprété). Dans les deux cas, l'exécutable compilé **s'appuie sur `msvbvm60.dll`** pour ses services de runtime (fonctions de chaînes, conversions, moteur de formulaires, gestion des `Variant`, etc.). C'est davantage une **bibliothèque de support** qu'un interpréteur.
- **Tout est bâti sur COM.** C'est le point fondamental : en VB6, **les objets *sont* des objets COM**. Une classe (`.cls`) est, sous le capot, un composant COM. L'identité, les interfaces et la **durée de vie** des objets suivent les règles de COM.
- **La mémoire repose sur le comptage de références.** Hérité de COM, ce mécanisme détruit un objet **immédiatement** lorsque sa dernière référence disparaît — ce qui donne la **finalisation déterministe** (`Class_Terminate` au bon moment). *C'est le cœur du piège n°1, traité en 2.2.*
- **32 bits, mono-thread (STA).** Le runtime VB6 est exclusivement **32 bits**, et son modèle de threads est l'**apartment à thread unique** (STA) : pour l'essentiel, une application VB6 est **mono-thread**.

> 💡 Retenez le mot-clé : **COM**. Le monde VB6 est un monde COM, et beaucoup de ce qui surprend
> lors de la migration vient de là.

---

## 🟢 Côté VB.NET : le CLR, l'exécution « managée »

Les applications VB.NET s'exécutent sur le **CLR** — *Common Language Runtime* —, le moteur commun  
de la plateforme .NET (sur .NET Framework 4.x, il s'agit du CLR 4).

Ses caractéristiques renversent, une à une, celles du VBVM :

- **Compilation en deux temps (IL + JIT).** Le compilateur VB.NET ne produit **pas** directement du code natif : il produit du **langage intermédiaire** (IL / MSIL), stocké dans un **assembly** (`.exe` ou `.dll`). C'est ensuite le **compilateur JIT** (*Just-In-Time*) du CLR qui traduit cet IL en code natif **au moment de l'exécution**.
- **Exécution managée.** Le CLR **gère** activement le programme pendant qu'il tourne : il assure la **gestion automatique de la mémoire**, la **sécurité de type**, la **vérification du code** et la **gestion des exceptions**. On parle de *code managé* (par opposition au code natif « non managé » du monde VB6/Win32).
- **Le ramasse-miettes (GC) à la place du comptage de références.** La mémoire n'est plus libérée par comptage de références mais par un **ramasse-miettes** qui décide **lui-même quand** collecter les objets devenus inutiles. Conséquence directe : **la finalisation n'est plus déterministe** *(→ 2.2)*.
- **Un système de types unifié (CTS).** Le *Common Type System* offre un système de types **commun** à tous les langages .NET, où **tout dérive d'`Object`** et où l'on distingue types **valeur** et types **référence**. Les types VB6 (`Variant`, `Currency`…) n'y ont plus leur place telle quelle *(→ 2.3)*.
- **Les exceptions structurées.** La gestion d'erreurs passe d'un modèle de **codes d'erreur** (`On Error`) à un modèle d'**exceptions** géré par le CLR (`Try`/`Catch`/`Finally`) *(→ 2.6 et chapitre 11)*.
- **Pas de limite au 32 bits, vrai multithreading.** Le CLR s'exécute en **64 bits** (ou 32 selon la configuration) et offre un **véritable modèle multithread**, là où VB6 restait cantonné au STA mono-thread.

> 💡 Le mot-clé symétrique du précédent : **managé**. Là où VB6 vivait dans le monde **COM/natif**,
> VB.NET vit dans le monde **managé** du CLR.

---

## ⚙️ La différence d'exécution, côte à côte

```
   VB6  ─────────────────────────────────────────────────────────────
   Code VB6 ──[compilateur VB6]──► code natif (.exe)
                                        │  au lancement, s'appuie sur…
                                        ▼
                                  msvbvm60.dll  (VBVM)
                                  • services COM   • comptage de références
                                  • Variant        • moteur de formulaires


   VB.NET  ──────────────────────────────────────────────────────────
   Code VB.NET ──[compilateur VB]──► IL (assembly .exe/.dll)
                                        │  à l'exécution…
                                        ▼
                                  CLR  ──[JIT]──► code natif
                                  • ramasse-miettes (GC)  • sécurité de type
                                  • CTS / Object           • exceptions
```

Deux philosophies, donc. VB6 : **compiler à l'avance** vers du natif, puis s'appuyer sur une  
bibliothèque de support COM. VB.NET : **compiler vers un format intermédiaire**, puis confier au CLR  
une **gestion active** de l'exécution. Ce n'est pas « le même moteur en plus récent » — ce sont
**deux moteurs, deux mondes**.

---

## 🆚 Tableau comparatif

| Aspect | **VBVM** (VB6) | **CLR** (VB.NET) |
|---|---|---|
| **Nature** | Bibliothèque de support (`msvbvm60.dll`) | Moteur d'exécution managé |
| **Fondation** | **COM** | Plateforme .NET / CTS |
| **Compilation** | Vers code **natif** (ou P-code) en amont | Vers **IL**, puis **JIT** au runtime |
| **Code** | Non managé (natif) | **Managé** (géré par le CLR) |
| **Mémoire** | **Comptage de références** (déterministe) | **Ramasse-miettes** (non déterministe) → 2.2 |
| **Objets** | Objets **COM** | Objets CLR (tout dérive d'`Object`) |
| **Types** | Types VB6 (`Variant`, `Currency`…) | **CTS** (valeur / référence) → 2.3 |
| **Erreurs** | Codes d'erreur (`On Error`) | **Exceptions** (`Try`/`Catch`) → 2.6 / chap. 11 |
| **Bits / threads** | 32 bits, **STA mono-thread** | 32/64 bits, **multithread** |
| **Déploiement** | `.exe` + `msvbvm60.dll` + COM/OCX enregistrés | Assembly + .NET Framework installé → chap. 18 |

---

## 🧠 Les cinq conséquences qui comptent pour la migration

Inutile de retenir tous les détails internes du CLR. Ce qu'il faut emporter, ce sont les **cinq  
répercussions concrètes** de ce changement de moteur — chacune ouvrant un pan de la formation :

1. **La mémoire change de logique** → comptage de références (déterministe) cède la place au **ramasse-miettes** (non déterministe). C'est **le piège n°1** : `Class_Terminate` ne se déclenche plus au bon moment. *(→ 2.2, puis 12.2-12.3)*
2. **Les objets changent de nature** → d'objets **COM** à objets **CLR**. C'est ce qui rend nécessaire l'**interopérabilité COM** (RCW/CCW) pour réutiliser l'existant et cohabiter pendant la transition. *(→ chapitre 16)*
3. **Les types changent de système** → un `Integer` n'a plus la même taille, `Variant` et `Currency` disparaissent. Débordements et conversions **silencieux** à la clé. *(→ 2.3)*
4. **Les erreurs changent de modèle** → de `On Error` aux **exceptions** structurées. *(→ 2.6, puis chapitre 11)*
5. **Le déploiement change de prérequis** → on ne livre plus `msvbvm60.dll` et des composants COM enregistrés, mais des **assemblies** et un **.NET Framework** installé. *(→ chapitre 18)*

> ⚠️ **Le fil conducteur du chapitre se vérifie ici.** Aucune de ces cinq répercussions n'est
> annoncée par une erreur de compilation. Le moteur a changé **sous** le code, sans que le code le
> dise. C'est exactement pourquoi ce chapitre 2 existe : **rendre visibles**, à l'œil humain, des
> ruptures que ni le code source ni le compilateur ne signalent.

---

## 💡 « Managé » vs « non managé » : un repère qui resservira

Vous rencontrerez sans cesse cette opposition pendant la migration, autant la fixer maintenant :

- **Code non managé** : code natif qui gère lui-même sa mémoire et ses ressources — le monde de VB6, de COM et de l'API Win32.
- **Code managé** : code exécuté **sous le contrôle du CLR**, qui prend en charge mémoire, types, sécurité et exceptions — le monde de .NET.

Toute la difficulté de l'interopérabilité (chapitre 16) vient de la **frontière** entre ces deux  
mondes : appeler du COM ou une API Win32 depuis du code managé, c'est **franchir cette frontière** —  
avec tout ce que cela implique de *marshaling* et de conversions. Comprendre dès maintenant **de  
quel côté** se trouve chaque chose vous fera gagner un temps précieux plus tard.

---

## ✅ À retenir

- Le **runtime** est le moteur qui rend les services d'exécution (mémoire, types, erreurs, objets) ; en changer, c'est changer **tout le socle** sous le code — **sans qu'aucune ligne ne le dise**.
- **VB6 → VBVM** (`msvbvm60.dll`) : un monde **COM**, à **comptage de références** (finalisation déterministe), compilé en **natif**, 32 bits mono-thread.
- **VB.NET → CLR** : l'exécution **managée**, en **IL + JIT**, avec **ramasse-miettes** (finalisation non déterministe), **CTS**, **exceptions**, multithread.
- Ce n'est **pas** « le même moteur en plus récent » : ce sont **deux moteurs, deux mondes** (COM/natif vs managé).
- Ce changement invisible engendre **cinq répercussions concrètes** — mémoire (→ 2.2), objets/COM (→ 16), types (→ 2.3), erreurs (→ 2.6/11), déploiement (→ 18) — qui structurent tout le reste de la formation.

---

> 🧭 **Section suivante** → [2.2 La finalisation déterministe perdue — le piège n°1](02-finalisation-deterministe.md) ⭐ ⚠️  
> 🔝 **Retour** → [Sommaire du chapitre 2](README.md)

⏭️ [La finalisation déterministe perdue — **le piège n°1**](/02-differences-fondamentales/02-finalisation-deterministe.md)
