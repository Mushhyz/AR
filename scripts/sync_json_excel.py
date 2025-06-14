"""
Module de synchronisation entre template Excel EBIOS RM et schéma JSON.
Garantit la cohérence des énumérations avec structure JSON-Schema.
"""

import json
import logging
from pathlib import Path
from typing import Dict, List, Any, Union

from openpyxl import load_workbook

logger = logging.getLogger(__name__)

class JSONExcelSyncer:
    """Synchronise les données entre Excel et JSON pour EBIOS RM avec structure JSON-Schema."""
    
    def __init__(self, excel_path: Path, json_path: Path):
        self.excel_path = Path(excel_path)
        self.json_path = Path(json_path)
    
    def extract_enums_from_excel(self) -> Dict[str, List]:
        """Extrait les énumérations depuis l'onglet __REFS en format compatible tests."""
        wb = load_workbook(self.excel_path, data_only=True)
        
        if "__REFS" not in wb.sheetnames:
            raise ValueError("Onglet __REFS introuvable dans le template Excel")
        
        ws = wb["__REFS"]
        enums = {}
        
        # Analyser les colonnes pour extraire les énumérations
        current_col = 1
        while current_col <= ws.max_column:
            header = ws.cell(row=1, column=current_col).value
            if not header:
                current_col += 1
                continue
            
            # Tables de niveaux (Gravité, Vraisemblance, Valeur Métier)
            if header == "ID":
                second_header = ws.cell(row=1, column=current_col + 1).value
                third_header = ws.cell(row=1, column=current_col + 2).value
                
                if second_header == "Libelle":
                    # Extraire les libellés directement
                    labels = []
                    row = 2
                    while ws.cell(row=row, column=current_col + 1).value:
                        label_val = ws.cell(row=row, column=current_col + 1).value
                        labels.append(label_val)
                        row += 1
                    
                    # **CORRECTION 4** : Identifier le type d'échelle et ajouter libellés complets
                    first_label = labels[0] if labels else ""
                    if first_label == "Négligeable":
                        enums["gravity_scale"] = labels  # Labels directement pour compatibilité tests
                        enums["gravity_labels"] = labels  # **CORRECTION 4** : Ajout libellé explicite
                    elif first_label == "Minimal":
                        enums["likelihood_scale"] = labels
                        enums["likelihood_labels"] = labels  # **CORRECTION 4**
                    elif first_label.startswith("Niveau"):
                        enums["business_value_scale"] = list(range(1, len(labels) + 1))
                        enums["business_value_labels"] = labels  # **CORRECTION 4**
                    elif first_label == "Faible" and third_header == "Valeur":
                        # Table Pertinence
                        enums["pertinence_scale"] = labels
                    elif first_label == "Limitée" and third_header == "Valeur":
                        # Table Exposition  
                        enums["exposition_scale"] = labels
            
            # **CORRECTION 1** : Tables avec ID spécifiques
            elif header == "Measure_ID":
                # Extraire les mesures de sécurité
                libelle_col = current_col + 1
                if ws.cell(row=1, column=libelle_col).value == "Libelle":
                    measures = []
                    row = 2
                    while ws.cell(row=row, column=current_col).value:
                        measure_data = {
                            "id": ws.cell(row=row, column=current_col).value,
                            "label": ws.cell(row=row, column=libelle_col).value,
                        }
                        # Ajouter autres colonnes si présentes
                        for extra_col in range(current_col + 2, current_col + 6):
                            extra_header = ws.cell(row=1, column=extra_col).value
                            if extra_header:
                                measure_data[extra_header.lower()] = ws.cell(row=row, column=extra_col).value
                        measures.append(measure_data)
                        row += 1
                    enums["measure_catalog"] = measures
            
            elif header in ["Asset_Type_ID", "Stakeholder_ID"]:
                # **CORRECTION 2** : Extraire ID et libellés pour types d'actifs et parties prenantes
                id_values = []
                labels = []
                
                libelle_col = current_col + 1
                if ws.cell(row=1, column=libelle_col).value == "Libelle":
                    row = 2
                    while ws.cell(row=row, column=current_col).value:
                        id_val = ws.cell(row=row, column=current_col).value
                        label_val = ws.cell(row=row, column=libelle_col).value
                        id_values.append(id_val)
                        labels.append(label_val)
                        row += 1
                    
                    if header == "Asset_Type_ID":
                        enums["asset_type_catalog"] = labels  # Tests attendent les libellés
                        enums["asset_type_ids"] = id_values   # IDs pour référence
                    elif header == "Stakeholder_ID":
                        enums["stakeholder_catalog"] = labels
                        enums["stakeholder_ids"] = id_values
            
            # Tables complexes - extraire en tant qu'objets
            elif header in ["Source_ID", "Scenario_ID", "OV_ID"]:
                # Extraire toute la table comme objets
                table_data = []
                headers_row = []
                
                # Lire les en-têtes de la table
                col = current_col
                while col <= ws.max_column and ws.cell(row=1, column=col).value:
                    headers_row.append(ws.cell(row=1, column=col).value)
                    col += 1
                
                # Lire les données
                row = 2
                while ws.cell(row=row, column=current_col).value:
                    row_data = {}
                    for i, header_name in enumerate(headers_row):
                        cell_value = ws.cell(row=row, column=current_col + i).value
                        row_data[header_name] = cell_value
                    table_data.append(row_data)
                    row += 1
                
                # Stocker selon le type
                if header == "Source_ID":
                    enums["source_catalog"] = table_data
                elif header == "Scenario_ID":
                    enums["scenario_catalog"] = table_data
                elif header == "OV_ID":
                    enums["operational_catalog"] = table_data
            
            # Passer au groupe suivant
            col = current_col
            while col <= ws.max_column and ws.cell(row=1, column=col).value:
                col += 1
            current_col = col + 1
        
        wb.close()
        return enums
    
    def sync_excel_to_json(self) -> None:
        """Synchronise les données Excel vers le fichier JSON avec structure attendue par les tests."""
        enums = self.extract_enums_from_excel()
        
        # **CORRECTION 4** : Structure conforme avec bloc enumerations complet et metadata
        schema_data = {
            "metadata": {
                "generator": "EBIOSTemplateGenerator", 
                "version": "2.0.0",  # **CORRECTION 4** : Version mise à jour
                "date": "2024-01-01",
                "profile": "EBIOS_RM_Complet",  # **CORRECTION 4**
                "total_enumerations": len(enums)
            },
            "enumerations": enums,  # **CORRECTION 4** : Toutes les énumérations dans le bloc principal
            "$defs": {
                "enumerations": {
                    enum_name: {"enum": enum_values} 
                    for enum_name, enum_values in enums.items()
                    if isinstance(enum_values, list) and enum_values  # Seulement listes non-vides
                }
            },
            "schema": {
                "atelier1_socle": {
                    "validations": {
                        "Type": {"enum": enums.get("asset_type_catalog", [])},
                        "Gravité": {"enum": enums.get("gravity_scale", [])},
                        "Confidentialité": {"enum": enums.get("gravity_scale", [])},
                        "Intégrité": {"enum": enums.get("gravity_scale", [])},
                        "Disponibilité": {"enum": enums.get("gravity_scale", [])},
                        "Valeur_Métier": {"enum": enums.get("business_value_labels", [])},
                        "Propriétaire": {"enum": enums.get("stakeholder_catalog", [])}
                    }
                },
                "atelier2_sources": {
                    "validations": {
                        "Source_ID": {"enum": [item.get("Source_ID") for item in enums.get("source_catalog", [])]},
                        "Pertinence": {"enum": enums.get("pertinence_scale", [])},  # **CORRECTION 2**
                        "Exposition": {"enum": enums.get("exposition_scale", [])}   # **CORRECTION 2**
                    }
                },
                "atelier3_scenarios": {
                    "validations": {
                        "Gravité": {"enum": enums.get("gravity_scale", [])},
                        "Vraisemblance": {"enum": enums.get("likelihood_scale", [])},
                        "Valeur_Métier": {"enum": enums.get("business_value_labels", [])}
                    }
                },
                # **CORRECTION 1** : Ajout des nouveaux ateliers
                "atelier4_operationnels": {
                    "validations": {
                        "Vraisemblance_Résiduelle": {"enum": enums.get("likelihood_scale", [])},
                        "Impact": {"enum": enums.get("gravity_scale", [])},
                        "Mesure_Recommandée": {"enum": [item.get("id") for item in enums.get("measure_catalog", [])]}
                    }
                },
                "atelier5_traitement": {
                    "validations": {
                        "Option_Traitement": {"enum": ["Réduire", "Éviter", "Transférer", "Accepter"]},
                        "Mesure_Choisie": {"enum": [item.get("id") for item in enums.get("measure_catalog", [])]},
                        "Responsable": {"enum": enums.get("stakeholder_catalog", [])},
                        "Statut": {"enum": ["Planifiée", "En cours", "Terminée", "Reportée", "Annulée"]}
                    }
                }
            }
        }
        
        # Sauvegarder le JSON
        self.json_path.parent.mkdir(parents=True, exist_ok=True)
        with open(self.json_path, 'w', encoding='utf-8') as f:
            json.dump(schema_data, f, indent=2, ensure_ascii=False)
        
        logger.info(f"Schéma JSON synchronisé : {self.json_path}")
    
    def validate_consistency(self) -> Dict[str, List[str]]:
        """Valide la cohérence entre Excel et JSON avec vérification enum."""
        issues = {"warnings": [], "errors": []}
        
        try:
            # Vérifier que les fichiers existent
            if not self.excel_path.exists():
                issues["errors"].append(f"Fichier Excel introuvable : {self.excel_path}")
                return issues
            
            if not self.json_path.exists():
                issues["warnings"].append(f"Fichier JSON introuvable : {self.json_path}")
                return issues
            
            # Extraire les données Excel
            excel_enums = self.extract_enums_from_excel()
            
            # Charger le JSON
            with open(self.json_path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
            
            json_enums = json_data.get("$defs", {}).get("enumerations", {})
            
            # Comparer les énumérations avec structure enum
            for enum_name in excel_enums:
                if enum_name not in json_enums:
                    issues["warnings"].append(f"Énumération '{enum_name}' présente dans Excel mais absente du JSON")
                elif excel_enums[enum_name] != json_enums[enum_name]:
                    # Comparaison des valeurs enum
                    excel_values = excel_enums[enum_name].get("enum", [])
                    json_values = json_enums[enum_name].get("enum", [])
                    if excel_values != json_values:
                        issues["warnings"].append(f"Énumération '{enum_name}' différente entre Excel ({excel_values}) et JSON ({json_values})")
            
            for enum_name in json_enums:
                if enum_name not in excel_enums:
                    issues["warnings"].append(f"Énumération '{enum_name}' présente dans JSON mais absente d'Excel")
            
        except Exception as e:
            issues["errors"].append(f"Erreur lors de la validation : {str(e)}")
        
        return issues


def main():
    """Point d'entrée pour test."""
    import sys
    
    if len(sys.argv) != 3:
        print("Usage: python sync_json_excel.py <excel_file> <json_file>")
        return
    
    excel_path = Path(sys.argv[1])
    json_path = Path(sys.argv[2])
    
    syncer = JSONExcelSyncer(excel_path, json_path)
    syncer.sync_excel_to_json()
    
    issues = syncer.validate_consistency()
    
    if issues["errors"]:
        print("❌ Erreurs détectées :")
        for error in issues["errors"]:
            print(f"  - {error}")
    
    if issues["warnings"]:
        print("⚠️ Avertissements :")
        for warning in issues["warnings"]:
            print(f"  - {warning}")
    
    if not issues["errors"] and not issues["warnings"]:
        print("✅ Synchronisation réussie, aucun problème détecté")


if __name__ == "__main__":
    main()
