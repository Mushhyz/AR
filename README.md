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

```bash
# Exécuter tous les tests
pytest

# Avec couverture
pytest --cov=ebiosrm_core

# Linting
ruff check .
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
