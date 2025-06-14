# EBIOS RM Generator v2

Générateur d'étude de risque EBIOS Risk Manager modulaire et configurable.

## 🚀 Installation

```bash
# Installation en mode développement
pip install -e .

# Ou avec les dépendances de développement
pip install -e ".[dev]"
```

## 📁 Structure du projet

```
project_root/
├── config/                 # Configuration et données
│   ├── assets.csv          # Actifs valorisés
│   ├── threats.csv         # Menaces et scénarios
│   └── settings.yaml       # Paramètres généraux
├── ebiosrm_core/          # Logique métier
├── cli.py                 # Interface en ligne de commande
└── tests/                 # Tests automatisés
```

## 🔧 Utilisation

### Génération complète

```bash
# Export Excel (par défaut)
ebiosrm export

# Export JSON
ebiosrm export --fmt json

# Répertoires personnalisés
ebiosrm export --cfg ./my_config --out ./my_output --fmt xlsx
```

### Validation des données

```bash
# Vérifier la cohérence des CSV/YAML
ebiosrm validate --cfg ./config
```

## 📊 Format des données

### assets.csv
```csv
id,type,label,criticality
A001,Data,Customer Database,Critical
A002,System,Web Server,High
```

### threats.csv
```csv
sr_id,ov_id,strategic_path,operational_steps
SR001,OV001,External Attack,Step1:High,Step2:Medium
SR002,OV002,Internal Fraud,Step1:Low,Step2:High,Step3:Medium
```

### settings.yaml
```yaml
excel_template: templates/ebiosrm_empty.xlsx
severity_scale: [Low, Medium, High, Critical]
likelihood_scale: [One-shot, Occasional, Probable, Systematic]
output_dir: build/
```

## 🧪 Tests

### Tests unitaires
```bash
# Exécuter tous les tests
pytest

# Tests avec sortie détaillée
pytest -v

# Tester un fichier spécifique
pytest tests/test_loader.py

# Tester une fonction spécifique
pytest tests/test_loader.py::test_load_assets

# Arrêter au premier échec
pytest -x
```

### Couverture de code
```bash
# Tests avec rapport de couverture
pytest --cov=ebiosrm_core

# Rapport de couverture en HTML
pytest --cov=ebiosrm_core --cov-report=html

# Couverture avec détails des lignes manquantes
pytest --cov=ebiosrm_core --cov-report=term-missing
```

### Validation des données
```bash
# Vérifier la cohérence des CSV/YAML
ebiosrm validate --cfg ./config

# Validation avec diagnostic détaillé
python debug_threats.py

# Tester le chargement des assets
python -c "from ebiosrm_core.loader import load_assets; print(load_assets('config/assets.csv'))"

# Tester le chargement des threats
python -c "from ebiosrm_core.loader import load_threats; print(load_threats('config/threats.csv'))"
```

### Qualité du code
```bash
# Linting avec ruff
ruff check .

# Formatage automatique
ruff format .

# Vérification du typage (si mypy installé)
mypy ebiosrm_core/

# Vérification de sécurité (si bandit installé)
bandit -r ebiosrm_core/
```

### Tests d'intégration
```bash
# Test complet de génération
ebiosrm export --cfg ./config --out ./test_output --fmt xlsx

# Vérifier que les fichiers sont générés
ls -la test_output/

# Test avec différents formats
ebiosrm export --fmt json
ebiosrm export --fmt xlsx
```

### Tests de performance
```bash
# Test avec profiling (si cProfile disponible)
python -m cProfile -o profile.stats cli.py export

# Analyser le profil
python -c "import pstats; p = pstats.Stats('profile.stats'); p.sort_stats('cumulative').print_stats(10)"
```

### Commandes de débogage
```bash
# Mode debug avec logs détaillés
python -c "import logging; logging.basicConfig(level=logging.DEBUG); from ebiosrm_core.cli import main; main()"

# Vérifier la structure des données chargées
python -c "
from ebiosrm_core.loader import load_all_data
assets, threats, settings = load_all_data('config')
print(f'Assets: {len(assets)}, Threats: {len(threats)}')
print('Premier asset:', assets[0] if assets else 'Aucun')
"

# Test de la logique de calcul
python -c "
from ebiosrm_core.generator import EbiosRmGenerator
gen = EbiosRmGenerator('config')
risks = gen.generate_risks()
print(f'Risques générés: {len(risks)}')
"
```

## 🐛 Dépannage

### Erreur "keywords must be strings"

Si vous obtenez cette erreur lors de la validation :

```bash
# Diagnostiquer le fichier threats.csv
python debug_threats.py

# Vérifier l'encodage du fichier
file config/threats.csv  # Sur Linux/Mac
```

**Causes communes :**
- Fichier CSV avec BOM (Byte Order Mark)
- Caractères non-UTF8 dans les en-têtes
- Colonnes vides ou avec des noms None
- Format CSV incorrect (virgules supplémentaires)

**Solutions :**
1. Ouvrir le CSV dans un éditeur de texte et vérifier l'encodage
2. Supprimer le BOM si présent
3. Vérifier que toutes les colonnes ont des noms valides
4. Utiliser l'encodage UTF-8 sans BOM

### Commande ebiosrm non trouvée

```bash
# Ajouter le répertoire Python Scripts au PATH
$env:PATH += ";C:\Users\<username>\AppData\Roaming\Python\Python313\Scripts"

# Ou utiliser le module directement
python -m ebiosrm_core.cli validate
```

### Erreurs Excel (formules supprimées)
Si des formules sont corrompues au sein du fichier .xlsx, Excel peut les supprimer automatiquement. 
Pour restaurer ou diagnostiquer ces formules :
1. Ouvrez le fichier dans Excel (mode protégé).
2. Suivez les indications de réparation.
3. Exportez à nouveau avec la commande `ebiosrm export`.

## 🏗️ Architecture

- **models.py**: Modèles Pydantic (Asset, Threat, Settings)
- **loader.py**: Chargement et validation des données
- **generator.py**: Logique de calcul des risques
- **exporters.py**: Export vers différents formats
- **cli.py**: Interface utilisateur

## 📈 Logique métier

1. **Gravité** = Maximum des criticités des actifs impactés
2. **Vraisemblance** = Moyenne pondérée des étapes opérationnelles
3. **Niveau de risque** = Matrice gravité × vraisemblance (4×4)

## 🤝 Contribution

1. Fork le projet
2. Créer une branche feature
3. Implémenter avec tests
4. Soumettre une pull request
