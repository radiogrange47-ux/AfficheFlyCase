# AfficheFlyCase

Application Windows pour générer une étiquette PDF A4 paysage par fly case à partir du classeur Excel `Listing Materiel.xlsx`.

## Utilisation

Lancer `AfficheFlyCase.exe` : la fenêtre de sélection Excel s’ouvre automatiquement dans le dossier où se trouve l’exécutable. Sélectionner le classeur, renseigner le nom du show, les dates et le lieu, vérifier la sélection, puis choisir le nom et l’emplacement du PDF. Toutes les fiches de l’onglet `Listing FlyCase` sont sélectionnées au chargement ; Ctrl+clic permet d’ajuster la sélection.

Le nom du show, les dates/période et le lieu/salle sont mémorisés automatiquement dans `%APPDATA%\AfficheFlyCase`. Le bouton « Choisir un logo… » permet d’ajouter un logo PNG ou JPEG ; une copie est conservée dans ce même dossier utilisateur et incluse sur chaque étiquette PDF. Le bouton « Supprimer » retire le logo mémorisé.

Sur l’étiquette, l’ID apparaît en très gros sur la première ligne avec le logo en haut à droite. La ligne suivante présente trois cases (production, date et lieu), puis les dimensions, l’empattement en m² et le cubage en m³. Le type, la couleur et tous les statuts Tip/Gerbable sont omis ; les commentaires relatifs au basculement, au gerbage ou à l’empilement sont également retirés.

Le classeur doit contenir en première et deuxième position les onglets `Listing Materiel` et `Listing FlyCase`. Les en-têtes attendus sont en ligne 2 ; le détail des colonnes et les règles d’import sont documentés dans [PROJECT.md](PROJECT.md). Un fichier non conforme est refusé avec une erreur indiquant le problème détecté.

Les matériels sans attribution de fly case (attribution absente ou ID inconnu) sont listés individuellement dans l’onglet « Erreurs d’affectation », avec leur ligne Excel et la raison. Ils sont exclus des étiquettes, mais ne bloquent pas la génération des autres fly cases.

## Construire le fichier EXE sous Windows

Prérequis : Python 3.10 ou plus récent. Depuis PowerShell, à la racine du dépôt :

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\build_windows.ps1
```

L’exécutable autonome est créé dans `dist\AfficheFlyCase.exe`. La construction doit être effectuée sous Windows ; PyInstaller ne produit pas un exécutable Windows depuis un autre système d’exploitation.

## Développement et tests

```powershell
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
python app.py
```

L’application traite les classeurs localement et ne requiert aucune connexion Internet à l’usage.
