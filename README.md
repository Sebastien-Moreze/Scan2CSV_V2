# Scan2CSV V2

Application Python pour extraire des informations d'entreprises depuis des fichiers PDF et les exporter en format CSV.

## Fonctionnalités

- Extraction de texte depuis des fichiers PDF (avec fallback OCR)
- Utilisation de l'API Anthropic (Claude) pour extraire les informations structurées
- Interface graphique intuitive avec Tkinter
- Export des données en format CSV
- Support multilingue (français/anglais)

## Prérequis

### 1. Python
- Python 3.7 ou supérieur

### 2. Tesseract OCR
**Sur macOS :**
```bash
brew install tesseract
brew install tesseract-lang  # Pour les langues supplémentaires
```

**Sur Windows :**
- Téléchargez et installez depuis : https://github.com/tesseract-ocr/tesseract
- Le chemin par défaut est : `C:\Program Files\Tesseract-OCR\tesseract.exe`

**Sur Linux (Ubuntu/Debian) :**
```bash
sudo apt-get install tesseract-ocr
sudo apt-get install tesseract-ocr-fra  # Pour le français
```

### 3. Clé API Anthropic
- Créez un compte sur https://console.anthropic.com/
- Générez une clé API
- Créez un fichier `.env` dans le dossier du projet avec :
```
ANTHROPIC_API_KEY=VOTRE_CLE_API_ICI
```

## Installation

1. **Cloner ou télécharger le projet**

2. **Installer les dépendances Python :**
```bash
pip install -r requirements.txt
```

3. **Configurer la clé API :**
- Copiez le fichier `env_example.txt` vers `.env`
- Remplacez `VOTRE_CLE_API_ICI` par votre vraie clé API Anthropic

## Utilisation

1. **Lancer l'application :**
```bash
python scan2csv_gui.py
```

2. **Dans l'interface :**
- Cliquez sur "Sélectionner un PDF" pour choisir votre fichier
- Cliquez sur "Lancer le traitement" pour extraire les informations
- Cliquez sur "Télécharger le CSV" pour sauvegarder les résultats

## Informations extraites

Pour chaque entreprise trouvée, l'application extrait :
- Nom de l'entreprise
- Nom du contact
- Numéro de téléphone
- Adresse de l'entreprise
- URL
- Résumé de la description

## Dépannage

### Erreur "Tesseract not found"
- Vérifiez que Tesseract est installé
- Sur macOS/Linux, assurez-vous qu'il est dans le PATH
- Sur Windows, modifiez la ligne 30 dans `scan2csv_gui.py` avec le bon chemin

### Erreur "Clé API non définie"
- Vérifiez que le fichier `.env` existe
- Vérifiez que la clé API est correctement définie

### Erreur d'extraction PDF
- L'application utilise d'abord PyPDF2, puis l'OCR si nécessaire
- Assurez-vous que le PDF n'est pas protégé par mot de passe

## Dépendances

- **PyPDF2** : Extraction de texte depuis PDF
- **anthropic** : Client pour l'API Anthropic
- **python-dotenv** : Gestion des variables d'environnement
- **pdfplumber** : Extraction avancée de PDF
- **pytesseract** : Interface Python pour Tesseract OCR
- **Pillow** : Traitement d'images
- **tkinter** : Interface graphique (inclus avec Python) 