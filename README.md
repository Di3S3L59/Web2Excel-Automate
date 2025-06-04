# Automatisation de traitement de dossiers [NOM_DU_SERVICE]

## Description

Ce projet permet d'automatiser le traitement de fichiers Excel contenant des dossiers, en récupérant automatiquement le nom des agents associés via une interface web interne ([URL_INTERNE]). Il s'appuie sur des scripts Python et PowerShell, ainsi que sur Selenium pour l'automatisation du navigateur.

## Fonctionnalités

- Extraction des numéros de dossier depuis un fichier Excel.
- Récupération automatique des noms d'agents via l'interface web interne.
- Mise à jour du fichier Excel avec les informations récupérées.
- Gestion des logs d'exécution et des erreurs.

## Prérequis

- **Python 3.x**
- **Selenium** (`pip install selenium`)
- **Pandas** (`pip install pandas`)
- **dateutil** (`pip install python-dateutil`)
- **Navigateur Firefox** et le driver GeckoDriver (placé dans le dossier `SeleniumDriverEXE`)
- **PowerShell** (pour l'exécution des scripts .ps1)
- **Microsoft Excel** (pour la manipulation des fichiers .xlsx via COM)

## Installation

1. **Cloner le dépôt** :
   ```sh
   git clone [URL_DU_DEPOT]
   ```

2. **Installer les dépendances Python** :
   ```sh
   pip install selenium pandas python-dateutil
   ```

3. **Placer le driver Selenium** :
   - Télécharger GeckoDriver compatible avec votre version de Firefox.
   - Placer l'exécutable dans le dossier `SeleniumDriverEXE`.

4. **Configurer les chemins** :
   - Adapter les chemins dans les scripts si besoin (remplacer les chemins utilisateurs par des chemins relatifs ou des variables d'environnement).

5. **(Optionnel) Configurer le proxy** :
   - Si vous êtes derrière un proxy, adapter la commande pip :
     ```
     pip install [NOM_PACKAGE] --proxy [ADRESSE_PROXY]
     ```

## Utilisation

### 1. Préparer le fichier Excel

- Placer le fichier Excel à traiter dans le dossier attendu (voir variable dans le script PowerShell).
- Le fichier doit contenir les colonnes attendues (numéro de dossier, etc.).

### 2. Lancer le script PowerShell

```powershell
pwsh [NOM_DU_SCRIPT].ps1
```

- Le script va :
  - Extraire les numéros de dossier.
  - Appeler le script Python pour récupérer les noms d'agents.
  - Mettre à jour le fichier Excel.

### 3. Logs

- Les logs d'exécution sont enregistrés dans le fichier `[NOM_DU_LOG].log`.
- **Attention** : ne pas partager ce fichier publiquement, il peut contenir des informations sensibles.

## Sécurité & Confidentialité

- **Ne partagez jamais les fichiers de logs ou les fichiers Excel contenant des données réelles.**
- **Anonymisez les chemins et les exemples avant toute diffusion externe.**
- **Les accès à l'interface interne sont soumis à authentification et à la politique de sécurité de votre entreprise.**

## Structure du projet

```
.
├── [SCRIPT_PYTHON].py         # Script principal Python
├── [SCRIPT_POWERSHELL].ps1    # Script principal PowerShell
├── SeleniumDriverEXE/         # Dossier pour le driver Selenium
├── [NOM_DU_LOG].log           # Fichier de log (à ne pas versionner)
├── README.md                  # Ce fichier
└── ...
```

## Auteurs

- [Votre Nom ou Équipe]
- Contact : [email@domaine.com]

---

**Note :** Ce projet est destiné à un usage interne. Toute diffusion externe doit faire l'objet d'une anonymisation et d'une validation par le responsable sécurité. 