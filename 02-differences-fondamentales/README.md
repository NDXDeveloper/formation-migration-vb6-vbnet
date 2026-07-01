🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 2. Différences fondamentales VB6 ↔ VB.NET ⭐ ⚠️

> **Partie 1 — Comprendre et cadrer la migration**
> Le chapitre le plus important pour comprendre **pourquoi** la migration est risquée — et **où**.

---

Le chapitre 1 a posé le cadrage : *pourquoi* migrer, *vers quoi*, et la carte des cinq fronts. Il a  
aussi annoncé l'idée maîtresse — migrer, c'est changer **langage + runtime + modèle objet** — et  
introduit la distinction décisive entre **changements visibles** et **changements silencieux**.

Ce chapitre **descend d'un cran**. Il prend chacune de ces différences de fond et l'examine de près :  
non plus « il y a des différences », mais **lesquelles**, **pourquoi** elles existent, et **comment**  
elles se transforment en bugs. C'est le **socle conceptuel** de toute la formation : tout ce que les  
parties suivantes traitent en pratique (types, formulaires, données, COM…) prend racine dans les  
mécanismes expliqués ici.

> ⚠️ **Pourquoi ⭐ *et* ⚠️ ?** Parce que ce chapitre est à la fois **central** (⭐ — sans lui, le
> reste n'a pas de fondations) et **dangereux** (⚠️ — il rassemble les pièges qui causent le plus de
> régressions). Les notions de ce chapitre sont celles que l'on regrette le plus de ne pas avoir
> comprises **avant** de migrer.

---

## 🎯 Ce que vous apprendrez dans ce chapitre

À la fin de ce chapitre, vous serez capable de :

- **Expliquer les deux runtimes** (VBVM et CLR) et pourquoi ce changement de moteur conditionne tout le reste.
- **Comprendre le piège n°1** — la perte de la finalisation déterministe — et savoir ce qui le remplace.
- **Repérer les changements silencieux de type** (entiers redimensionnés, `Variant`, `Currency`) avant qu'ils ne produisent des bugs.
- **Auditer les appels de procédures** à la lumière du basculement **ByRef → ByVal par défaut**.
- **Anticiper** les pièges des propriétés par défaut, du `Set` disparu et de la gestion d'erreurs.
- **Disposer d'une vue de synthèse** des incompatibilités majeures, à garder sous la main pendant toute la migration.

---

## 🧊 Le fil conducteur : traquer le silencieux

Reprenons la grille du chapitre 1 — elle structure **tout** ce chapitre :

| | **Changement visible** | **Changement silencieux** |
|---|---|---|
| Détecté par | le **compilateur** | **personne** (le code compile) |
| Symptôme | « ça ne compile pas » | « ça compile, mais ne fait plus pareil » |
| Danger | faible | **élevé** |

Chaque section de ce chapitre pose, à un moment, la même question : **« ce changement est-il visible  
ou silencieux ? »**. Et la plupart des sections marquées ⚠️ ci-dessous traitent justement de
**changements silencieux** — ceux que le compilateur **ne** signalera **pas**. C'est ce qui justifie
l'existence même de ce chapitre : rendre **visible**, à l'œil humain, ce que l'outillage laisse  
passer.

> 💡 **Lien avec l'annexe B.** Ce chapitre **explique** les mécanismes ; l'**annexe B** (le catalogue
> des pièges silencieux) les transforme en fiches **symptôme → cause → correction**. Lisez ce
> chapitre pour *comprendre*, gardez l'annexe B ouverte pour *corriger*. Les deux se répondent.

---

## 📚 Les sections de ce chapitre

Ce chapitre se décline en sept sections, de la rupture la plus profonde (le runtime) jusqu'à la  
synthèse :

- **2.1 — Deux runtimes, deux mondes (VBVM vs CLR)**
  La rupture fondatrice, invisible dans le code : on change de moteur d'exécution. Tout le reste en découle.

- **2.2 — La finalisation déterministe perdue — le piège n°1** ⭐ ⚠️
  En VB6, `Class_Terminate` se déclenche quand le compteur de références tombe à zéro ; en .NET, c'est le *garbage collector* qui décide **quand**. La conséquence la plus structurante de la migration, et sa parade (`IDisposable`/`Using`).

- **2.3 — Le typage : `Variant`, entiers redimensionnés, `Currency`** ⚠️
  Pourquoi un `Integer` ne fait plus la même taille, où `Variant` et `Currency` disparaissent — et les débordements silencieux qui en résultent.

- **2.4 — ByRef par défaut → ByVal par défaut** ⭐ ⚠️
  Le changement d'appel le plus sournois : la même procédure ne modifie plus la variable de l'appelant. *Le code compile, le résultat change.*

- **2.5 — Propriétés par défaut et mot-clé `Set` (disparu)** ⚠️
  La fin des propriétés par défaut et du `Set` de VB6 — et le faux-ami que constitue le `Set` de VB.NET.

- **2.6 — Gestion d'erreurs : `On Error` vs exceptions structurées**
  Le passage d'un modèle de codes d'erreur (`On Error`) à un modèle d'exceptions (`Try`/`Catch`) — un aperçu avant le traitement détaillé du chapitre 11.

- **2.7 — Tableau de synthèse des incompatibilités majeures**
  La vue d'ensemble récapitulative : l'aide-mémoire des grandes différences, à garder à portée de main.

---

## 🧭 Comment lire ce chapitre

- **La section 2.2 est la plus importante du chapitre** (le piège n°1). Si votre temps est compté, c'est par elle qu'il faut commencer après la 2.1.
- Les sections **2.2, 2.3, 2.4 et 2.5** (marquées ⚠️) sont des **changements silencieux** : lisez-les comme des **mises en garde**, pas seulement comme de la théorie.
- Ce chapitre **explique et alerte** ; il **renvoie** au traitement pratique dans les parties 3 à 5. Vous n'avez pas à tout corriger ici — vous avez à **comprendre** ce qu'il faudra traquer.

> ⚠️ **Une erreur à éviter en lisant.** Ne confondez pas « j'ai compris la différence » avec « mon
> code est corrigé ». Comprendre le piège ByRef ne corrige pas vos appels ; c'est l'**audit**
> (chapitre 10) et les **tests** (chapitre 17) qui le font. Ce chapitre vous arme le **regard** —
> l'application viendra ensuite.

---

## 👤 À qui s'adresse ce chapitre

À **tous les profils techniques** de la formation — c'est un passage **incontournable** pour
quiconque touchera au code.

- **Développeurs** : c'est votre socle. Les parties 3 à 5 supposent ces notions acquises.
- **Chefs de projet** : même sans coder, comprendre **pourquoi** ces différences créent du risque aide à **planifier** le temps de validation à sa juste valeur (et à ne pas croire les mythes du 1.5).
- **Profils de cadrage** : la section de synthèse (2.7) suffit à saisir l'essentiel des incompatibilités pour chiffrer et décider.

Prérequis : avoir lu le **chapitre 1** (en particulier 1.3, dont ce chapitre est le prolongement  
détaillé). Connaître VB6 reste la base utile.

---

## ✅ À retenir

- Ce chapitre est le **socle conceptuel** de la formation : il explique **les mécanismes** dont découlent tous les pièges pratiques des parties suivantes.
- Son fil conducteur est la **traque du silencieux** : rendre visible à l'œil humain ce que le compilateur laisse passer.
- La **section 2.2** (finalisation déterministe perdue) est **le piège n°1** et la notion la plus structurante du chapitre.
- Plusieurs sections (2.2 à 2.5, ⚠️) traitent de **changements silencieux** : à lire comme des **mises en garde**.
- **Comprendre** un piège (ce chapitre) n'est pas le **corriger** : la correction relève des parties 3 à 5, des tests, et de l'**annexe B**.

---

> 🧭 **Section suivante** → [2.1 Deux runtimes, deux mondes (VBVM vs CLR)](01-vbvm-vs-clr.md)  
> 🔝 **Retour** → [Sommaire général](../SOMMAIRE.md)

⏭️ [Deux runtimes, deux mondes (VBVM vs CLR)](/02-differences-fondamentales/01-vbvm-vs-clr.md)
