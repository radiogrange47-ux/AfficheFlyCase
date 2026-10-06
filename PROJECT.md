# AfficheFlyCase — cahier des charges

## 1. Présentation

**AfficheFlyCase** est un logiciel de bureau Windows qui transforme un tableau Excel recensant des fly cases en étiquettes prêtes à imprimer. Chaque fly case sélectionné donne lieu à une étiquette sur une page A4, inspirée de la référence visuelle fournie : un en-tête d’identification, une zone d’informations logistiques, une grande zone « contenu / remarques » et, si nécessaire, un avertissement de manutention très visible.

Le livrable destiné aux utilisateurs est un fichier exécutable Windows (`.exe`). L’application doit fonctionner sans installation préalable de Python ou d’autres dépendances.

## 2. Objectifs

- Éviter la création manuelle et répétitive des étiquettes.
- Rendre l’identification du fly case et ses consignes de manutention immédiatement visibles.
- Permettre à l’utilisateur de compléter les renseignements propres à une production qui ne figurent pas dans le tableau.
- Produire un document imprimable fiable : une page A4 complète par fly case, sans texte tronqué ni débordement.

## 3. Parcours utilisateur

1. L’utilisateur ouvre l’application et sélectionne le fichier Excel (`.xlsx` ou `.xlsm`) à traiter.
2. L’application vérifie que le classeur correspond au modèle décrit ci-dessous, en particulier ses deux premiers onglets. L’utilisateur n’a pas à choisir les feuilles ni à associer les colonnes.
3. Si la structure est conforme, l’application importe les données et affiche les fly cases reconnus ainsi que les avertissements éventuels. Si le fichier ne correspond pas au modèle ou est illisible, l’application affiche un message d’erreur explicite et arrête l’import sans produire d’étiquettes.
4. L’utilisateur renseigne les informations de production absentes du tableau :
   - nom du show / de la production ;
   - date ou période ;
   - lieu / salle.
5. L’utilisateur vérifie la liste des fly cases et le contenu associé, peut sélectionner ceux à éditer et prévisualise les étiquettes.
6. Il génère un PDF, puis peut l’ouvrir pour impression ou choisir un emplacement de sortie.

Les intitulés et la valeur des champs de production sont repris à l’identique sur toutes les étiquettes de la génération. L’application ne modifie jamais le classeur source.

## 4. Format Excel attendu et règles d’import

Le classeur doit contenir les onglets suivants, dans cet ordre, avec les en-têtes sur la **ligne 2**. Les noms d’onglets et les colonnes attendues sont ceux du fichier de référence :

### Onglet 1 : `Listing Materiel`

Les colonnes B à F doivent être présentes dans cet ordre :

| Colonne | En-tête en ligne 2 | Utilisation |
| --- | --- | --- |
| B | `ELEMENT` | Nom de l’élément à placer dans le contenu de l’étiquette. |
| C | `POSITION` | Repère, emplacement ou répartition de l’élément. |
| D | `QTE` | Quantité de l’élément. |
| E | `SPARE` | Quantité de spare/réserve, si renseignée. |
| F | `FLYCASE` | Désignation du ou des fly cases associés à l’élément, par exemple `LX07 - Fly Case - Bois - 84 x 70 x 83`. |

Chaque ligne de matériel est associée au fly case (ou aux fly cases) mentionné(s) en colonne F. L’application extrait l’ID du fly case de cette désignation et l’utilise pour faire le rapprochement avec l’onglet `Listing FlyCase`. Si une cellule F contient plusieurs désignations, chacune doit être reconnue et l’élément associé à chacune. Les champs `ELEMENT`, `POSITION`, `QTE` et `SPARE` constituent le contenu de l’étiquette ; une quantité de spare vide ou indiquée par `-` ne doit pas être inventée ni affichée comme quantité disponible.

Dans le fichier de référence, certaines cellules B à E sont fusionnées verticalement sur plusieurs lignes, tandis que la colonne F contient une désignation de fly case par ligne. La valeur de chaque cellule fusionnée B à E s’applique à **chacune** des lignes couvertes par la fusion : l’import doit la réutiliser pour tous les fly cases correspondants, sans perdre les associations sur les lignes où Excel n’expose qu’une valeur dans la cellule supérieure de la fusion.

### Onglet 2 : `Listing FlyCase`

Les colonnes B à L doivent être présentes dans cet ordre :

| Colonne | En-tête en ligne 2 | Utilisation |
| --- | --- | --- |
| B | `ID` | Identifiant unique, par exemple `LX07` ; affiché en très grand. |
| C | `Type2` | Type de contenant (par exemple Fly Case, Caisse, Tube ou Chariot). |
| D | `Couleur` | Couleur ou finition. |
| E | `Largeur` | Dimension avec l’unité du classeur. |
| F | `Longeur` | Longueur avec l’unité du classeur ; conserver l’orthographe de l’en-tête du modèle. |
| G | `Hauteur` | Hauteur avec l’unité du classeur. |
| H | `Empatement` | Empattement, s’il est renseigné. |
| I | `Cubage` | Volume, s’il est renseigné. |
| J | `Tip` | `OK` signifie basculable ; `NON` signifie ne pas basculer. |
| K | `Gerbable` | `OK` signifie gerbable ; `NON` signifie non gerbable. |
| L | `Commentaire` | Commentaire associé au fly case, s’il est renseigné. |

Les dimensions affichées proviennent des colonnes E à G et conservent leurs valeurs et unités. L’empattement et le cubage sont affichés en complément lorsqu’ils sont renseignés. L’onglet `Listing FlyCase` fait foi pour la liste des fly cases et leurs caractéristiques : chaque fiche avec un ID peut donner lieu à une étiquette, même si aucun élément de `Listing Materiel` ne lui est associé. Les éléments de l’onglet 1 sont ajoutés à l’étiquette correspondante par l’ID ; la désignation textuelle complète de la colonne F ne remplace pas les valeurs de la fiche `Listing FlyCase`.

### Validation du fichier et erreurs

- Lire automatiquement les deux premiers onglets, dans l’ordre ; ne pas demander à l’utilisateur de choisir les onglets ni de remapper les colonnes.
- Avant toute génération, vérifier l’extension et la lisibilité du classeur, la présence des deux onglets attendus dans le bon ordre, les en-têtes requis en ligne 2 et l’existence d’au moins une fiche de fly case avec un ID.
- Refuser l’import si un onglet est absent/mal nommé ou si un en-tête requis est absent, déplacé ou différent du modèle. Ne pas essayer de deviner ou de substituer silencieusement une autre feuille ou colonne.
- Si le fichier ne correspond pas au modèle, afficher un message du type : **« Fichier Excel non conforme : onglet “Listing FlyCase” manquant. Vérifiez que vous avez sélectionné le listing matériel attendu. »** Le message doit préciser le problème détecté (format illisible/non pris en charge, onglet manquant ou inattendu, en-tête manquant, ou données incohérentes) et indiquer, si possible, la feuille, la cellule et la valeur concernées. Ne pas générer de PDF à partir d’un import refusé.
- Refuser les IDs de fly case absents ou en double dans `Listing FlyCase`, les désignations de la colonne F dont aucun ID ne peut être extrait, ainsi que les IDs de matériel qui ne correspondent à aucune fiche. Le message doit permettre de localiser la ligne en cause.
- Ignorer les lignes entièrement vides. Ne pas inventer les valeurs absentes ; signaler les champs de contenu ou de fiche manquants dans l’aperçu.
- Interpréter les valeurs de `Tip` et `Gerbable` indépendamment : `OK` et `NON` sont les valeurs reconnues. Une valeur non vide différente doit être signalée comme incohérente et bloquer la génération jusqu’à correction du fichier. Une cellule vide est un statut inconnu, à signaler sans la transformer en `OK` ou `NON`.
- Les formules Excel ne doivent pas être exécutées. Si une valeur attendue est une formule sans résultat exploitable, signaler la cellule et demander à l’utilisateur de vérifier/recalculer le classeur dans Excel.
- L’application ne doit ni exécuter les macros d’un fichier `.xlsm`, ni enregistrer de modification dans le classeur.
- Si une donnée est trop longue pour la zone prévue, adapter la composition sans la couper ; si elle ne peut toujours pas tenir avec une taille lisible, avertir l’utilisateur et empêcher une génération qui masquerait l’information.

## 5. Maquette de l’étiquette

### Format et disposition

- Une page **A4 paysage** par fly case (297 × 210 mm), conforme au format A4 de la série définie par l’ISO 216.
- Mise en page à bordure visible, avec marges de sécurité de l’ordre de 8 à 10 mm afin de limiter les risques de rognage à l’impression.
- Composition inspirée de l’image de référence, sans reproduire sa marque ni son adresse :
  1. **En-tête pleine largeur** : nom du show / production, puis date(s) et lieu.
  2. **Bandeau d’informations logistiques** : ID du fly, dimensions et statuts de manutention ; poids, zone de scène et case « Checked » si disponibles ou utiles.
  3. **Grande zone “CONTENU / REMARQUES”** : contenu et commentaires clairement séparés et faciles à lire.
  4. **Avertissement de basculement** : bandeau rouge en bas lorsque le fly case ne doit pas être basculé ; aucun bandeau d’interdiction si le statut est autorisé ou inconnu.

### Hiérarchie typographique et lisibilité

Tailles indicatives à ajuster lors des essais d’impression ; la priorité est de rester lisible et de faire tenir toutes les informations :

1. **Avertissement “NE PAS BASCULER / DO NOT TIP”** : environ 32–44 pt, capitales, gras, texte clair sur fond rouge.
2. **ID du fly case** : environ 24–32 pt, gras, contraste élevé.
3. **Nom du show / production** : environ 20–28 pt, gras.
4. **Titres des zones, dimensions et statuts** : environ 14–18 pt, avec libellés explicites.
5. **Contenu, commentaires, dates et lieu** : environ 11–14 pt selon la quantité de texte.

Ces valeurs sont des points de départ, non des tailles fixes. La typographie doit conserver une hiérarchie nette, un contraste suffisant et des libellés écrits : ne jamais communiquer un statut uniquement par une couleur ou une icône. Lorsqu’une zone contient beaucoup de texte, réduire d’abord les espacements et optimiser les retours à la ligne ; diminuer ensuite la police jusqu’à un minimum cible de 10 pt. Au-delà de cette limite, signaler le problème plutôt que produire une étiquette illisible.

### Consignes de manutention

- La colonne `Tip` détermine le bandeau de basculement : `NON` déclenche le bandeau rouge « NE PAS BASCULER / DO NOT TIP » ; `OK` affiche « BASCULABLE ». Le statut doit être impossible à confondre avec le statut de gerbage.
- La colonne `Gerbable` détermine un statut séparé, sous forme textuelle explicite : `GERBABLE` pour `OK` et `NON GERBABLE` pour `NON`.
- Un statut absent, vide ou non reconnu est **inconnu** : ne pas le convertir en « oui » ou « non ». Le signaler dans l’aperçu et, si nécessaire, imprimer « STATUT À VÉRIFIER ».

## 6. Sortie et impression

- Format de sortie principal : PDF standard, une étiquette par page, dans l’ordre du tableau ou de la sélection.
- Chaque page doit déclarer un format A4 paysage. L’impression à « taille réelle » doit respecter ce format ; l’application ne doit pas dépendre de l’imprimante choisie pour composer la page.
- Fournir un aperçu avant génération et un récapitulatif des avertissements.
- Le fichier généré doit être nommé avec le nom du show et une date de génération, sans écraser silencieusement un PDF existant.
- Les étiquettes doivent rester lisibles en niveaux de gris, à l’exception de l’avertissement rouge qui doit également rester identifiable par son texte et sa typographie.

## 7. Interface et fonctionnement

- Interface de bureau en français, avec des libellés simples et des erreurs compréhensibles.
- Sélection du fichier Excel au moyen d’un sélecteur de fichiers ; afficher le chemin retenu et permettre de le remplacer.
- Étapes clairement visibles : fichier et feuille, correspondance des colonnes, informations du show, sélection/aperçu, génération.
- Aucune connexion Internet ne doit être nécessaire à l’utilisation courante ; les données des classeurs ne sont pas envoyées à un service externe.
- Les paramètres de correspondance peuvent être mémorisés localement pour faciliter les générations suivantes, avec possibilité de les modifier. Ne pas conserver les données métier du classeur sans nécessité.

## 8. Livrable Windows

- Fournir un `.exe` exécutable sur Windows sans installation de l’interpréteur ni des bibliothèques de développement.
- Le paquet de production doit inclure les ressources nécessaires à l’interface et à la génération du PDF.
- Afficher une erreur explicite en cas d’échec d’import, d’écriture du PDF ou d’ouverture de l’emplacement de sortie ; ne pas annoncer une génération réussie si le fichier n’a pas été créé.
- À titre de piste d’implémentation si le projet est réalisé en Python : `openpyxl` sait lire les classeurs Excel modernes `.xlsx` et `.xlsm`, et PyInstaller peut empaqueter une application Python dans un exécutable autonome. Vérifier le comportement du paquet sur un poste Windows propre avant livraison.

## 9. Critères d’acceptation

1. L’utilisateur peut sélectionner un fichier Excel conforme au modèle, qui est importé depuis ses deux premiers onglets sans modifier le classeur.
2. Les champs de production sont demandés avant la génération et apparaissent sur chaque page.
3. Pour chaque ligne sélectionnée contenant un fly case valide, le PDF produit exactement une page A4 paysage.
4. Les éléments de `Listing Materiel` sont associés à la fiche correspondante de `Listing FlyCase` par leur ID ; les valeurs de colonnes B à E fusionnées sont appliquées à chaque ligne fusionnée et aucune association de la colonne F n’est perdue.
5. L’ID, les dimensions, le contenu, les commentaires et les statuts importés sont restitués fidèlement.
6. Les statuts de basculement et de gerbage sont traités indépendamment ; « ne pas basculer » déclenche un avertissement très visible.
7. Un contenu long ne déborde pas, n’est pas tronqué silencieusement et déclenche une alerte s’il ne peut tenir de façon lisible.
8. L’aperçu et le PDF ne présentent ni texte coupé, ni page supplémentaire, ni mise à l’échelle qui change le format A4.
9. Le `.exe` démarre et génère un PDF sur un poste Windows où Python n’est pas installé.
10. Un classeur illisible ou non conforme (mauvais onglets, ordre ou en-têtes) provoque un message d’erreur précis et aucune génération ; les erreurs d’écriture du fichier sont également communiquées.

## 10. Hors périmètre initial

- Modifier, compléter ou réenregistrer le tableau Excel.
- Calculer le poids, les dimensions ou les statuts à partir d’autres valeurs.
- Imprimer automatiquement sans validation de l’utilisateur.
- Prendre en charge tous les anciens formats Excel binaires (`.xls`) sans décision et dépendance d’implémentation dédiées.
- Synchroniser les classeurs ou les étiquettes avec un serveur distant.

## 11. Références consultées

- ISO, **ISO 216 — Writing paper and certain classes of printed matter — Trimmed sizes — A and B series** : <https://www.iso.org/standard/36631.html>
- Microsoft Support, **Formats de fichiers pris en charge dans Excel** : <https://support.microsoft.com/en-us/office/excel-file-formats-that-are-supported-in-excel-0943ff2c-6014-4e8d-aaea-b83d51d46247>
- Documentation `openpyxl`, **lecture et écriture des fichiers Excel 2010** (`.xlsx`, `.xlsm`) : <https://openpyxl.readthedocs.io/en/stable/>
- Documentation PyInstaller, **empaquetage d’une application et génération d’un exécutable autonome** : <https://pyinstaller.org/en/stable/>

Les recommandations de tailles de police ci-dessus sont des choix de conception initiaux déduits de la hiérarchie visible sur l’image de référence ; elles doivent être confirmées par un essai d’impression A4 à taille réelle.