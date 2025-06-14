"""
Gestionnaire de configuration EBIOS RM.
Lit les paramètres depuis les fichiers de config et les applique aux templates.
"""

import yaml
import json
from pathlib import Path
from typing import Dict, List, Any, Optional
import logging

logger = logging.getLogger(__name__)

class EBIOSConfigManager:
    """Gestionnaire centralisé de la configuration EBIOS RM."""
    
    def __init__(self, config_dir: Path = Path("config")):
        self.config_dir = Path(config_dir)
        self.config_file = self.config_dir / "config_ebios.yaml"
        self.config_data = {}
        
        self.load_config()
    
    def load_config(self) -> None:
        """Charge la configuration depuis le fichier YAML."""
        if self.config_file.exists():
            with open(self.config_file, 'r', encoding='utf-8') as f:
                self.config_data = yaml.safe_load(f) or {}
            logger.info(f"Configuration chargée depuis {self.config_file}")
        else:
            logger.warning(f"Fichier de configuration non trouvé : {self.config_file}")
            self.config_data = self._get_default_config()
    
    def save_config(self) -> None:
        """Sauvegarde la configuration dans le fichier YAML."""
        self.config_file.parent.mkdir(parents=True, exist_ok=True)
        with open(self.config_file, 'w', encoding='utf-8') as f:
            yaml.dump(self.config_data, f, default_flow_style=False, allow_unicode=True)
        logger.info(f"Configuration sauvegardée dans {self.config_file}")
    
    def get_gravity_scale(self) -> List[Dict[str, Any]]:
        """Retourne l'échelle de gravité configurée."""
        echelles = self.config_data.get('echelles', {})
        gravite = echelles.get('gravite', {}).get('niveaux', {})
        
        return [
            {
                "ID": level_id,
                "Libelle": data.get('libelle', f'Niveau {level_id}'),
                "Description": data.get('description', ''),
                "Couleur": data.get('couleur', '#7F8C8D')
            }
            for level_id, data in gravite.items()
        ]
    
    def get_likelihood_scale(self) -> List[Dict[str, Any]]:
        """Retourne l'échelle de vraisemblance configurée."""
        echelles = self.config_data.get('echelles', {})
        vraisemblance = echelles.get('vraisemblance', {}).get('niveaux', {})
        
        return [
            {
                "ID": level_id,
                "Libelle": data.get('libelle', f'Niveau {level_id}'),
                "Description": data.get('description', ''),
                "Frequence": data.get('frequence', '')
            }
            for level_id, data in vraisemblance.items()
        ]
    
    def get_risk_sources(self) -> List[Dict[str, str]]:
        """Retourne les sources de risque configurées."""
        sources = self.config_data.get('sources_risque', {})
        categories = sources.get('categories', [])
        motivations = sources.get('motivations', [])
        
        # Générer des sources par défaut basées sur les catégories
        risk_sources = []
        for i, category in enumerate(categories, 1):
            motivation = motivations[i-1] if i-1 < len(motivations) else "Non définie"
            risk_sources.append({
                "Source_ID": f"RS{i:03d}",
                "Label": f"Source {category}",
                "Category": category,
                "Motivation": motivation,
                "Resources": "À définir",
                "Targeting": "À définir"
            })
        
        return risk_sources
    
    def get_strategic_scenarios(self) -> List[Dict[str, str]]:
        """Retourne les scénarios stratégiques configurés."""
        scenarios = self.config_data.get('scenarios_strategiques', {})
        objectifs = scenarios.get('objectifs_types', [])
        chemins = scenarios.get('chemins_attaque', [])
        
        strategic_scenarios = []
        for i, objectif in enumerate(objectifs, 1):
            chemin = chemins[i-1] if i-1 < len(chemins) else "Non défini"
            strategic_scenarios.append({
                "Scenario_ID": f"SR{i:03d}",
                "Risk_Source": f"RS{((i-1) % 5) + 1:03d}",  # Rotation sur les sources
                "Target_Objective": objectif,
                "Attack_Path": chemin,
                "Motivation": "À définir"
            })
        
        return strategic_scenarios
    
    def get_operational_scenarios(self) -> List[Dict[str, str]]:
        """Retourne les scénarios opérationnels configurés."""
        vecteurs = self.config_data.get('vecteurs_attaque', {})
        techniques = vecteurs.get('techniques', [])
        etapes = vecteurs.get('etapes_kill_chain', [])
        
        operational_scenarios = []
        for i, technique in enumerate(techniques, 1):
            etapes_str = " > ".join(etapes[:4])  # Prendre les 4 premières étapes
            operational_scenarios.append({
                "OV_ID": f"OV{i:03d}",
                "Strategic_Scenario": f"SR{((i-1) % 6) + 1:03d}",
                "Attack_Vector": technique,
                "Operational_Steps": etapes_str
            })
        
        return operational_scenarios
    
    def get_risk_matrix(self) -> Dict[str, Any]:
        """Retourne la matrice de risque configurée."""
        return self.config_data.get('matrice_risque', {
            'type': '4x4',
            'seuils': {'acceptable': [1,2,3], 'attention': [4,6,8], 'critique': [9,12,16]}
        })
    
    def get_export_colors(self) -> Dict[str, str]:
        """Retourne les couleurs d'export configurées."""
        export_config = self.config_data.get('export', {})
        return export_config.get('couleurs_risque', {
            'Faible': '#27AE60',
            'Moyen': '#F39C12', 
            'Élevé': '#E74C3C',
            'Critique': '#C0392B'
        })
    
    def is_pme_profile(self) -> bool:
        """Vérifie si le profil PME est activé."""
        pme_config = self.config_data.get('pme_config', {})
        return pme_config.get('echelles_simplifiees', False)
    
    def update_scale(self, scale_type: str, level_id: int, data: Dict[str, Any]) -> None:
        """Met à jour une échelle de valeurs."""
        if 'echelles' not in self.config_data:
            self.config_data['echelles'] = {}
        
        if scale_type not in self.config_data['echelles']:
            self.config_data['echelles'][scale_type] = {'niveaux': {}}
        
        self.config_data['echelles'][scale_type]['niveaux'][level_id] = data
        logger.info(f"Échelle {scale_type} niveau {level_id} mise à jour")
    
    def add_risk_source(self, source_data: Dict[str, str]) -> None:
        """Ajoute une nouvelle source de risque."""
        # Cette méthode nécessiterait une extension de la structure de config
        logger.info(f"Ajout source de risque : {source_data.get('Source_ID', 'N/A')}")
    
    def export_to_json(self, output_path: Path) -> None:
        """Exporte la configuration en JSON."""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(self.config_data, f, indent=2, ensure_ascii=False)
        logger.info(f"Configuration exportée en JSON : {output_path}")
    
    def _get_default_config(self) -> Dict[str, Any]:
        """Retourne la configuration par défaut."""
        return {
            'metadata': {
                'version': '2.0',
                'organisation': 'Organisation par défaut'
            },
            'echelles': {
                'gravite': {
                    'niveaux': {
                        1: {'libelle': 'Négligeable', 'description': 'Impact minimal'},
                        2: {'libelle': 'Limité', 'description': 'Impact modéré'},
                        3: {'libelle': 'Important', 'description': 'Impact significatif'},
                        4: {'libelle': 'Critique', 'description': 'Impact majeur'}
                    }
                },
                'vraisemblance': {
                    'niveaux': {
                        1: {'libelle': 'Minimal', 'description': 'Très peu probable'},
                        2: {'libelle': 'Significatif', 'description': 'Possible'},
                        3: {'libelle': 'Élevé', 'description': 'Probable'},
                        4: {'libelle': 'Maximal', 'description': 'Quasi-certain'}
                    }
                }
            }
        }

def main():
    """Test du gestionnaire de configuration."""
    config_manager = EBIOSConfigManager()
    
    print("📊 Échelle de gravité:")
    for item in config_manager.get_gravity_scale():
        print(f"  {item['ID']}: {item['Libelle']} - {item['Description']}")
    
    print("\n📈 Échelle de vraisemblance:")
    for item in config_manager.get_likelihood_scale():
        print(f"  {item['ID']}: {item['Libelle']} - {item['Description']}")
    
    print(f"\n🏢 Profil PME: {config_manager.is_pme_profile()}")
    
    print("\n⚠️ Sources de risque:")
    for source in config_manager.get_risk_sources()[:3]:  # Premières 3
        print(f"  {source['Source_ID']}: {source['Label']}")

if __name__ == "__main__":
    main()
