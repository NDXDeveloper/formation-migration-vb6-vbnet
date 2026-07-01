🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 1. Introduction : pourquoi et comment migrer ⭐

> **Partie 1 — Comprendre et cadrer la migration**
> Module d'entrée de la formation. À lire **en premier**, quel que soit votre profil.

---

Avant d'ouvrir Visual Studio, avant de lancer le moindre assistant de conversion, il faut  
répondre à trois questions simples — et y répondre **honnêtement** :

1. **Pourquoi** migrer ce code VB6 aujourd'hui, plutôt que de le laisser tourner ?
2. **Vers quoi** migrer exactement, et pourquoi cette cible-là ?
3. **Comment** s'y prendre sans casser une application qui, bien souvent, fait tourner l'entreprise depuis quinze ou vingt ans ?

Ce chapitre ne migre pas encore une ligne de code. Il **cadre** le projet. Son rôle est de  
remplacer les idées reçues par une vision réaliste de ce qui vous attend, pour que les décisions  
prises ensuite — stratégie, outils, périmètre, planning — reposent sur du solide plutôt que sur  
des promesses.

> ⚠️ **Le piège le plus coûteux se joue ici, avant le code.** Sous-estimer la nature du chantier
> (« ce n'est qu'un changement de syntaxe »), surestimer les outils (« l'assistant fait tout »),
> ou viser d'emblée la mauvaise cible : ces erreurs de cadrage se paient pendant des mois. Un
> chapitre d'introduction bien digéré fait gagner un temps considérable sur tout le reste.

---

## 🎯 Ce que vous apprendrez dans ce chapitre

À la fin de ce chapitre, vous serez capable de :

- **Justifier** une migration VB6 → VB.NET auprès d'une direction ou d'une équipe (arguments techniques, économiques, humains et de sécurité).
- **Comprendre pourquoi .NET Framework 4.7.2** est retenu comme cible — le « pont » qui minimise le nombre de changements simultanés.
- **Identifier ce qui change réellement** entre VB6 et VB.NET : non pas seulement le langage, mais aussi le runtime et le modèle objet.
- **Avoir une vue d'ensemble** de tous les fronts d'une migration (langage, formulaires, données, COM, API), pour ne pas découvrir les difficultés une par une.
- **Distinguer les mythes des réalités**, afin d'aborder le projet avec des attentes calibrées.

---

## 🧭 L'idée directrice : trois changements en une fois

Le fil rouge de toute la formation est posé dès ce chapitre.

Migrer du Visual Basic 6 vers VB.NET n'est **pas** une conversion de syntaxe. C'est un changement  
de **trois choses simultanément** :

| Ce qui change | De VB6… | …à VB.NET |
|---------------|---------|-----------|
| **Le langage** | mots-clés, types, structures de VB6 | la syntaxe et les types de VB.NET |
| **Le runtime** | le moteur VB6 (« VBVM ») | le **CLR** de .NET, avec un *garbage collector* |
| **Le modèle objet** | instanciation, durée de vie, finalisation déterministe | nouvelle gestion du cycle de vie des objets |

C'est cette **simultanéité** qui rend la migration délicate. Un changement de langage seul serait  
gérable ; un changement de runtime seul aussi. Les trois ensemble produisent des effets de bord  
que le compilateur ne signale pas.

> ⚠️ **Le vrai danger n'est pas le code qui ne compile pas — c'est le code qui compile mais ne se
> comporte plus pareil.** Un `ByRef` qui devient `ByVal`, un `Integer` qui change de taille, un
> `Class_Terminate` qui ne se déclenche plus au bon moment : ces **changements silencieux** sont
> la première cause de régressions. Toute la formation est construite autour de cette priorité.

C'est aussi cette idée qui justifie le choix de la cible : viser **.NET Framework 4.7.2** plutôt  
que .NET moderne, c'est délibérément **réduire le nombre de changements simultanés** pour garder  
le chantier maîtrisable. Le saut vers .NET moderne, lui, est traité comme une **étape séparée et  
optionnelle** (module 20).

---

## 📚 Les sections de ce chapitre

Ce chapitre se décline en cinq sections, qui répondent dans l'ordre aux questions de cadrage :

- **1.1 — Pourquoi migrer VB6 en 2026** ⭐
  Les raisons concrètes de ne plus rester sur VB6 : IDE hors support, situation du runtime, dette technique, difficulté de recrutement, enjeux de sécurité.

- **1.2 — Pourquoi cibler .NET Framework 4.7.2 : le « pont » naturel** ⭐
  Pourquoi cette version précise, et pas .NET 8/10 : compatibilité API, COM/ActiveX, Windows Forms, runtime supporté. La cible qui minimise les risques.

- **1.3 — VB6 vs VB.NET : ce qui change vraiment**
  Au-delà de la syntaxe : langage, runtime, modèle objet et *garbage collector*. La vue d'ensemble des différences de fond, détaillées ensuite au chapitre 2.

- **1.4 — Panorama complet d'une migration**
  La carte des fronts à couvrir : le langage, les formulaires, les données, l'interopérabilité COM et les appels d'API. De quoi anticiper l'ampleur réelle du chantier.

- **1.5 — Mythes et réalités** ⚠️
  Démontage des idées reçues les plus coûteuses (« l'assistant fait tout », « ce n'est que de la syntaxe »), pour partir avec des attentes justes.

---

## 👤 À qui s'adresse ce chapitre

À **tout le monde**, sans exception — c'est le socle commun de la formation.

- **Décideurs et profils de cadrage** y trouveront de quoi évaluer, chiffrer et décider (avec les modules 2 à 4).
- **Développeurs** qui attaquent la migration du langage y poseront les bases conceptuelles indispensables avant les chapitres techniques.
- **Chefs de projet** y gagneront le vocabulaire et la vision d'ensemble nécessaires pour piloter sereinement.

Aucun prérequis .NET n'est nécessaire pour ce chapitre. Une connaissance de VB6 (le code que vous  
allez migrer) suffit largement.

---

## ✅ À retenir

- La migration VB6 → VB.NET est un changement de **langage + runtime + modèle objet**, pas une simple traduction de syntaxe.
- Le risque principal, ce sont les **comportements qui changent en silence** — pas les erreurs de compilation.
- **.NET Framework 4.7.2** est choisi comme **pont** parce qu'il minimise le nombre de changements simultanés ; le passage à .NET moderne est une migration distincte et ultérieure.
- Un **bon cadrage en amont** (ce chapitre) conditionne le succès de tout le reste du projet.

---

> 🧭 **Section suivante** → [1.1 Pourquoi migrer VB6 en 2026](01-pourquoi-migrer.md)

⏭️ [Pourquoi migrer VB6 en 2026 (IDE hors support, runtime, dette technique, recrutement, sécurité)](/01-cadrage-migration/01-pourquoi-migrer.md)
