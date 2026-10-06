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

Chaque ligne de matériel est associée au fly case (ou aux fly cases) mentionné(s) en colonne F. L’application extrait les ID du fly case de cette désignation et les utilise pour faire le rapprochement avec l’onglet `Listing FlyCase`. Si un même matériel est attribué à plusieurs fly cases, sa quantité est suivie de « (aussi dans <ID> et <ID>…) », indiquant les autres fly cases de l’attribution. Lorsqu’un spare est renseigné, il est accolé à la quantité sous la forme `QTE+SPARE` (par exemple `2+1`) ; il ne donne pas lieu à un libellé séparé. Un spare vide ou indiqué par `-` n’est pas affiché.

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
| J | `Tip` | `NON` affiche un avertissement « NE PAS TIPER » ; les autres valeurs n’affichent pas de statut positif. |
| K | `Gerbable` | `NON` affiche un avertissement « NE PAS GERBER » ; les autres valeurs n’affichent pas de statut positif. |
| L | `Commentaire` | Commentaire associé au fly case, s’il est renseigné. |

Les dimensions affichées proviennent des colonnes E à G et conservent les unités explicites ; lorsque l’unité est absente, elles sont considérées en centimètres. L’empattement est affiché en m² et le cubage en m³, arrondis à une décimale avec une virgule, lorsqu’ils sont renseignés. L’onglet `Listing FlyCase` fait foi pour la liste des fly cases et leurs caractéristiques : chaque fiche avec un ID peut donner lieu à une étiquette, même si aucun élément de `Listing Materiel` ne lui est associé. Les éléments de l’onglet 1 sont ajoutés à l’étiquette correspondante par l’ID ; la désignation textuelle complète de la colonne F ne remplace pas les valeurs de la fiche `Listing FlyCase`.

### Validation du fichier et erreurs

- Lire automatiquement les deux premiers onglets, dans l’ordre ; ne pas demander à l’utilisateur de choisir les onglets ni de remapper les colonnes.
- Avant toute génération, vérifier l’extension et la lisibilité du classeur, la présence des deux onglets attendus dans le bon ordre, les en-têtes requis en ligne 2 et l’existence d’au moins une fiche de fly case avec un ID.
- Refuser l’import si un onglet est absent/mal nommé ou si un en-tête requis est absent, déplacé ou différent du modèle. Ne pas essayer de deviner ou de substituer silencieusement une autre feuille ou colonne.
- Si le fichier ne correspond pas au modèle, afficher un message du type : **« Fichier Excel non conforme : onglet “Listing FlyCase” manquant. Vérifiez que vous avez sélectionné le listing matériel attendu. »** Le message doit préciser le problème détecté (format illisible/non pris en charge, onglet manquant ou inattendu, en-tête manquant, ou données incohérentes) et indiquer, si possible, la feuille, la cellule et la valeur concernées. Ne pas générer de PDF à partir d’un import refusé.
- Refuser les IDs de fly case absents ou en double dans `Listing FlyCase`. Pour une ligne de matériel qui a un élément mais pas d’attribution de fly case (colonne F vide, désignation sans ID valide ou ID absent de `Listing FlyCase`), afficher une erreur d’affectation **distincte pour chaque matériel**, avec le numéro de ligne, sa désignation, sa quantité/position si disponibles et la cause. Présenter le détail dans un onglet « Erreurs d’affectation » et avertir l’utilisateur à l’import. Ignorer ces éléments non affectés pour la création des étiquettes, sans bloquer la génération des fly cases correctement identifiés.
- Ignorer les lignes entièrement vides, y compris celles qui se trouvent dans une plage de cellules fusionnées et dont les cellules afficheraient autrement la valeur de la ligne précédente. Ignorer également les lignes de total/sous-total et les lignes qui répètent les en-têtes des colonnes ; elles ne représentent pas des éléments ou des fiches et leurs éventuelles formules de calcul ne doivent pas bloquer l’import. Ne pas inventer les valeurs absentes ; signaler les champs de contenu ou de fiche manquants dans l’aperçu.
- Les colonnes `Tip` et `Gerbable` restent requises dans le modèle Excel. La valeur `NON` déclenche le bandeau d’avertissement correspondant ; aucun statut positif n’est affiché.
- Les formules Excel ne doivent pas être exécutées. Si une valeur attendue est une formule sans résultat exploitable, signaler la cellule et demander à l’utilisateur de vérifier/recalculer le classeur dans Excel.
- L’application ne doit ni exécuter les macros d’un fichier `.xlsm`, ni enregistrer de modification dans le classeur.
- Si une donnée est trop longue pour la zone prévue, adapter la composition sans la couper ; si elle ne peut toujours pas tenir avec une taille lisible, avertir l’utilisateur et empêcher une génération qui masquerait l’information.

## 5. Maquette de l’étiquette

### Format et disposition

- Une page **A4 paysage** par fly case (297 × 210 mm), conforme au format A4 de la série définie par l’ISO 216.
- Mise en page à bordure visible, avec marges de sécurité de l’ordre de 8 à 10 mm afin de limiter les risques de rognage à l’impression.
- Composition inspirée de l’image de référence, sans reproduire sa marque ni son adresse :
  1. **Première ligne** : ID du fly case en gros caractères à gauche ; pictogrammes d’identification batterie et/ou produit chimique détectés dans le contenu, immédiatement à gauche du logo mémorisé en haut à droite.
  2. **Trois cases** : production, date/période et lieu/salle.
  3. **Bandeau des caractéristiques** : dimensions, puis empattement (m²) et cubage (m³), sans répéter leurs noms. Ne pas afficher le type du contenant ni sa couleur.
  4. **Zone principale** : liste des éléments et de leurs quantités/positions, sans titre « CONTENU » ; si le fly case n’a aucun contenu associé, laisser la zone vide sans texte de remplacement.
  5. **Section “COMMENTAIRES”** séparée en bas de l’étiquette, au-dessus de l’avertissement éventuel. Les fragments de commentaire relatifs au basculement, au gerbage ou à l’empilement sont retirés.
  6. **Avertissement de manutention** : afficher « NE PAS TIPER » lorsque `Tip` vaut `NON`, « NE PAS GERBER » lorsque `Gerbable` vaut `NON`, ou les deux si les deux conditions sont réunies. Aucun bandeau n’est affiché si aucune valeur n’est `NON`.

Le logo batterie fourni par l’utilisateur est affiché à partir d’une ressource PNG incluse dans l’exécutable. Il utilise le même cadre que le logo de production (192 × 96 pt), avec conservation de ses proportions. Le pictogramme chimique reste un repère générique. Ces pictogrammes sont des repères de contenu, pas des pictogrammes réglementaires de danger (GHS/CLP) ni une détermination de classe de transport. La détection s’appuie sur les désignations du matériel ; elle n’infère pas un niveau de danger.

### Hiérarchie typographique et lisibilité

Tailles indicatives à ajuster lors des essais d’impression ; la priorité est de rester lisible et de faire tenir toutes les informations :

1. **Bandeau « NE PAS TIPER / NE PAS GERBER »** : 76 pt de hauteur, texte gras jusqu’à 48 pt, ajusté uniquement si un avertissement combiné doit tenir dans la largeur.
2. **ID du fly case** : environ 72 pt, gras, première ligne, contraste élevé.
3. **Nom du show / production** : environ 15–18 pt, gras, dans la case « Production ».
4. **Titres des zones et dimensions** : environ 10–14 pt, avec libellés explicites seulement lorsque nécessaires.
5. **Contenu, commentaires, dates et lieu** : environ 11–14 pt selon la quantité de texte.

Ces valeurs sont des points de départ, non des tailles fixes. La typographie doit conserver une hiérarchie nette, un contraste suffisant et des libellés écrits : ne jamais communiquer un statut uniquement par une couleur ou une icône. Lorsqu’une zone contient beaucoup de texte, réduire d’abord les espacements et optimiser les retours à la ligne ; diminuer ensuite la police jusqu’à un minimum cible de 10 pt. Au-delà de cette limite, signaler le problème plutôt que produire une étiquette illisible.

## 6. Sortie et impression

- Format de sortie principal : PDF standard, une étiquette par page, dans l’ordre du tableau ou de la sélection.
- Chaque page doit déclarer un format A4 paysage. L’impression à « taille réelle » doit respecter ce format ; l’application ne doit pas dépendre de l’imprimante choisie pour composer la page.
- Fournir un aperçu avant génération et un récapitulatif des avertissements.
- Le nom proposé par défaut au moment de l’enregistrement est `EtiquettesFly_A4_<Nom de la production>.pdf`. Les caractères interdits dans les noms de fichiers Windows sont remplacés et la fenêtre de sauvegarde conserve la confirmation avant écrasement.
- Les étiquettes doivent rester lisibles en niveaux de gris, à l’exception de l’avertissement rouge qui doit également rester identifiable par son texte et sa typographie.

## 7. Interface et fonctionnement

- Interface de bureau en français, avec des libellés simples et des erreurs compréhensibles.
- Sélection du fichier Excel au moyen d’un sélecteur de fichiers ; afficher le chemin retenu et permettre de le remplacer.
- Étapes clairement visibles : fichier et feuille, correspondance des colonnes, informations du show, sélection/aperçu, génération.
- Aucune connexion Internet ne doit être nécessaire à l’utilisation courante ; les données des classeurs ne sont pas envoyées à un service externe.
- Mémoriser localement le dernier nom du show, les dates/période et le lieu/salle saisis ; les préremplir au lancement suivant et les sauvegarder automatiquement lorsqu’ils changent.
- Permettre à l’utilisateur de sélectionner un logo PNG ou JPEG depuis une fenêtre de fichiers. Copier l’image choisie dans le dossier de données utilisateur AfficheFlyCase et mémoriser son emplacement afin de la conserver même si le fichier original est déplacé. Inclure le logo, proportionnellement et sans déformation, dans la zone d’en-tête de chaque étiquette et réserver l’espace nécessaire pour éviter le chevauchement avec les textes.
- Prévoir une action pour supprimer le logo mémorisé. Les préférences et la copie du logo résident sur l’ordinateur de l’utilisateur ; elles ne sont pas intégrées au classeur Excel et ne sont pas envoyées à un service externe.

## 8. Livrable Windows

- Fournir un `.exe` exécutable sur Windows sans installation de l’interpréteur ni des bibliothèques de développement.
- Le paquet de production doit inclure les ressources nécessaires à l’interface et à la génération du PDF.
- Afficher une erreur explicite en cas d’échec d’import, d’écriture du PDF ou d’ouverture de l’emplacement de sortie ; ne pas annoncer une génération réussie si le fichier n’a pas été créé.
- À titre de piste d’implémentation si le projet est réalisé en Python : `openpyxl` sait lire les classeurs Excel modernes `.xlsx` et `.xlsm`, et PyInstaller peut empaqueter une application Python dans un exécutable autonome. Vérifier le comportement du paquet sur un poste Windows propre avant livraison.

## 9. Critères d’acceptation

1. L’utilisateur peut sélectionner un fichier Excel conforme au modèle, qui est importé depuis ses deux premiers onglets sans modifier le classeur.
2. L’ID apparaît en première ligne avec le logo à droite ; production, date et lieu apparaissent chacun dans leur case dédiée.
3. Pour chaque ligne sélectionnée contenant un fly case valide, le PDF produit exactement une page A4 paysage.
4. Les éléments de `Listing Materiel` sont associés à la fiche correspondante de `Listing FlyCase` par leur ID ; les valeurs de colonnes B à E fusionnées sont appliquées à chaque ligne fusionnée et aucune association de la colonne F n’est perdue. Lorsqu’un même matériel est affecté à plusieurs fly cases, la quantité mentionne également les autres ID.
5. Les dimensions apparaissent avant les valeurs d’empattement et de cubage, sans répéter le nom de ces deux mesures ; type et couleur sont absents de l’étiquette.
6. Le logo batterie fourni et/ou le repère chimique apparaît à gauche du logo de l’étiquette lorsque la désignation d’un contenu correspond aux termes détectés ; les ressources nécessaires sont incluses dans l’exécutable et ces pictogrammes ne prétendent pas signaler une classe de danger réglementaire.
7. Les commentaires sont dans leur propre section en bas ; les commentaires de manutention sont exclus.
8. Les valeurs `NON` des colonnes Tip et Gerbable déclenchent respectivement les avertissements « NE PAS TIPER » et « NE PAS GERBER » ; les valeurs positives n’affichent pas de statut.
9. Un contenu long ne déborde pas, n’est pas tronqué silencieusement et déclenche une alerte s’il ne peut tenir de façon lisible.
10. L’aperçu et le PDF ne présentent ni texte coupé, ni page supplémentaire, ni mise à l’échelle qui change le format A4.
11. Le `.exe` démarre et génère un PDF sur un poste Windows où Python n’est pas installé.
12. Les informations du show sont conservées après fermeture et restaurées au prochain lancement.
13. Un logo PNG ou JPEG choisi est mémorisé et apparaît dans chaque page PDF sans déformer le format A4 ni chevaucher les textes ; un logo invalide ou manquant provoque une erreur explicite.
14. Un classeur illisible ou non conforme (mauvais onglets, ordre ou en-têtes) provoque un message d’erreur précis et aucune génération ; les erreurs d’écriture du fichier sont également communiquées.

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