"""Tests de validation du template EBIOS RM avec énumérations et SoA."""

import pytest
import json
from pathlib import Path
from openpyxl import load_workbook
from scripts.generate_template import EBIOSTemplateGenerator
from scripts.sync_json_excel import EBIOSJSONExporter


class TestTemplateValidation:
    """Tests de validation du template et de l'export JSON."""
    
    @pytest.fixture
    def template_path(self, tmp_path):
        """Génère un template de test."""
        template_file = tmp_path / "test_template.xlsx"
        generator = EBIOSTemplateGenerator()
        generator.generate_template(template_file)
        return template_file
    
    def test_template_generation(self, template_path):
        """Test de génération du template."""
        assert template_path.exists()
        
        wb = load_workbook(template_path)
        expected_sheets = [
            "Config_EBIOS", 
            "Atelier1_Socle", 
            "Atelier2_Sources", 
            "Atelier3_Scenarios",
            "Atelier4_Operationnels", 
            "Atelier5_Traitement", 
            "Synthese"
        ]
        
        for sheet in expected_sheets:
            assert sheet in wb.sheetnames, f"Onglet manquant: {sheet}"
        
        # Vérifier que l'onglet __REFS existe et est masqué
        assert "__REFS" in wb.sheetnames
        refs_sheet = wb["__REFS"]
        assert refs_sheet.sheet_state == "veryHidden"
    
    def test_reference_tables(self, template_path):
        """Test de présence des tables de référence."""
        wb = load_workbook(template_path)
        
        # Vérifier les plages nommées essentielles
        expected_ranges = [
            "Gravite", "Vraisemblance", "Pertinence", "Exposition",
            "Measure_ID", "Asset_Type", "Stakeholder_ID",
            "tbl_Gravite_Valeur", "tbl_Vraisemblance_Valeur"
        ]
        
        for range_name in expected_ranges:
            assert range_name in wb.defined_names, f"Plage nommée manquante: {range_name}"
    
    def test_measure_catalog_structure(self, template_path):
        """Test du catalogue des mesures ISO 27001."""
        wb = load_workbook(template_path)
        refs_ws = wb["__REFS"]
        
        # Chercher la table tbl_Measure
        found_measure_table = False
        measure_headers = []
        
        for row in refs_ws.iter_rows():
            for cell in row:
                if cell.value == "Measure_ID":
                    found_measure_table = True
                    # Récupérer les en-têtes de la ligne
                    for col_cell in refs_ws[cell.row]:
                        if col_cell.value:
                            measure_headers.append(col_cell.value)
                    break
            if found_measure_table:
                break
        
        assert found_measure_table, "Table tbl_Measure non trouvée"
        
        expected_headers = ["Measure_ID", "Libelle", "Category", "Cout", "Efficacite_pct", "AnnexA_Control"]
        for header in expected_headers:
            assert header in measure_headers, f"En-tête manquant dans tbl_Measure: {header}"
    
    def test_formulas_protection(self, template_path):
        """Test de protection sélective des formules."""
        wb = load_workbook(template_path)
        
        # Vérifier l'Atelier 5 pour les formules de risque résiduel
        ws = wb["Atelier5_Traitement"]
        
        # Vérifier qu'il y a des formules protégées
        protected_formulas = 0
        for row in ws.iter_rows():
            for cell in row:
                if (cell.value and isinstance(cell.value, str) and 
                    cell.value.startswith('=') and cell.protection.locked):
                    protected_formulas += 1
        
        assert protected_formulas > 0, "Aucune formule protégée trouvée"
    
    def test_data_validations(self, template_path):
        """Test des validations de données avec messages personnalisés."""
        wb = load_workbook(template_path)
        ws = wb["Atelier2_Sources"]
        
        # Vérifier qu'il y a des validations
        validations = ws.data_validations.dataValidation
        assert len(validations) > 0, "Aucune validation de données trouvée"
        
        # Vérifier les messages d'erreur personnalisés pour Pertinence/Exposition
        found_custom_error = False
        for dv in validations:
            if dv.error and "Niveau de pertinence invalide" in dv.error:
                found_custom_error = True
                assert dv.showErrorMessage, "Message d'erreur non activé"
                assert dv.showInputMessage, "Message d'aide non activé"
                break
        
        assert found_custom_error, "Messages d'erreur personnalisés non trouvés"


class TestJSONExport:
    """Tests de l'export JSON avec bloc enumerations."""
    
    @pytest.fixture
    def json_exporter(self, template_path):
        """Crée un exporteur JSON avec un template de test."""
        return EBIOSJSONExporter(template_path)
    
    def test_enumerations_bloc(self, json_exporter, tmp_path):
        """Test du bloc enumerations dans l'export JSON."""
        json_file = tmp_path / "test_export.json"
        json_exporter.export_to_json(json_file)
        
        assert json_file.exists()
        
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Vérifier la structure avec bloc enumerations en racine
        assert "enumerations" in data
        assert "metadata" in data
        assert "measure_catalog" in data
        assert "annexa_controls" in data
        
        # Vérifier les échelles dans enumerations
        enums = data["enumerations"]
        assert "gravity_labels" in enums
        assert "likelihood_labels" in enums
        assert "pertinence_labels" in enums
        assert "exposition_labels" in enums
        
        # Vérifier les libellés français
        assert "Négligeable" in enums["gravity_labels"]
        assert "Minimal" in enums["likelihood_labels"]
        assert "Faible" in enums["pertinence_labels"]
        assert "Limitée" in enums["exposition_labels"]
    
    def test_soa_generation(self, json_exporter, tmp_path):
        """Test de génération du Statement of Applicability."""
        json_file = tmp_path / "test_export.json"
        json_exporter.export_to_json(json_file)
        
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        soa = data["annexa_controls"]
        
        # Vérifier la structure SoA
        assert "iso27001_version" in soa
        assert soa["iso27001_version"] == "2022"
        assert "overall_coverage" in soa
        assert "domains" in soa
        assert "summary" in soa
        
        # Vérifier les métriques
        summary = soa["summary"]
        assert "total_controls" in summary
        assert "implemented" in summary
        assert "coverage_target" in summary
        assert summary["coverage_target"] == 90
        assert "compliance_level" in summary
    
    def test_json_parity_validation(self, json_exporter, tmp_path):
        """Test de cohérence entre Excel et JSON pour les énumérations."""
        json_file = tmp_path / "test_export.json"
        json_exporter.export_to_json(json_file)
        
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        enums = data["enumerations"]
        
        # Vérifier que les échelles correspondent
        assert len(enums["gravity_scale"]) == len(enums["gravity_labels"])
        assert len(enums["likelihood_scale"]) == len(enums["likelihood_labels"])
        assert len(enums["pertinence_scale"]) == len(enums["pertinence_labels"])
        
        # Vérifier les valeurs numériques
        assert enums["gravity_scale"] == [1, 2, 3, 4]
        assert enums["likelihood_scale"] == [1, 2, 3, 4]
        assert enums["pertinence_scale"] == [1, 2, 3]


class TestKPICalculations:
    """Tests des calculs KPI Velocity/Preparedness."""
    
    def test_velocity_formulas(self, template_path):
        """Test des formules Velocity dans Synthèse."""
        wb = load_workbook(template_path, data_only=False)
        ws = wb["Synthese"]
        
        # Chercher les formules de Velocity
        velocity_formulas = []
        velocity_section_found = False
        
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and "VELOCITY" in str(cell.value):
                    velocity_section_found = True
                # **CORRECTION** : Rechercher les formules françaises Excel
                if (cell.value and isinstance(cell.value, str) and 
                    cell.value.startswith('=') and 
                    ("AVERAGE" in cell.value or "MOYENNE" in cell.value or "COUNTIFS" in cell.value or "NB.SI.ENS" in cell.value)):
                    velocity_formulas.append(cell.value)
        
        assert velocity_section_found, "Section Velocity non trouvée"
        assert len(velocity_formulas) > 0, "Formules Velocity non trouvées"
        
        # **CORRECTION** : Vérifier que les formules référencent la table Incidents
        incidents_references = 0
        for formula in velocity_formulas:
            if "Incidents[" in formula:
                incidents_references += 1
        
        assert incidents_references > 0, "Aucune référence à la table Incidents trouvée"
    
    def test_preparedness_indicators(self, template_path):
        """Test des indicateurs Preparedness."""
        wb = load_workbook(template_path)
        ws = wb["Synthese"]
        
        # Vérifier la présence des sections KPI
        found_velocity = False
        found_preparedness = False
        
        for row in ws.iter_rows():
            for cell in row:
                if cell.value:
                    if "VELOCITY" in str(cell.value):
                        found_velocity = True
                    if "PREPAREDNESS" in str(cell.value):
                        found_preparedness = True
        
        assert found_velocity, "Section Velocity non trouvée"
        assert found_preparedness, "Section Preparedness non trouvée"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
