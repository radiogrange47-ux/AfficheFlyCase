# AfficheFlyCase

Application Windows pour générer une étiquette PDF A4 paysage par fly case à partir du classeur Excel `Listing Materiel.xlsx`.

## Utilisation

Lancer `AfficheFlyCase.exe`, sélectionner le classeur, renseigner le nom du show, les dates et le lieu, vérifier la sélection, puis choisir le nom et l’emplacement du PDF. Toutes les fiches de l’onglet `Listing FlyCase` sont sélectionnées au chargement ; Ctrl+clic permet d’ajuster la sélection.

Le classeur doit contenir en première et deuxième position les onglets `Listing Materiel` et `Listing FlyCase`. Les en-têtes attendus sont en ligne 2 ; le détail des colonnes et les règles d’import sont documentés dans [PROJECT.md](PROJECT.md). Un fichier non conforme est refusé avec une erreur indiquant le problème détecté.

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
