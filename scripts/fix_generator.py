"""
Script de diagnostic et correction pour le générateur EBIOS RM.
Identifie les méthodes manquantes et propose des corrections.
"""

import inspect
from pathlib import Path
import re

def analyze_generator_methods():
    """Analyse le générateur pour identifier les méthodes manquantes."""
    
    print("🔍 DIAGNOSTIC DU GÉNÉRATEUR EBIOS RM")
    print("="*50)
    
    try:
        from generate_template import EBIOSTemplateGenerator
        
        # Instancier le générateur
        generator = EBIOSTemplateGenerator()
        
        # Lister toutes les méthodes
        all_methods = [method for method in dir(generator) if method.startswith('_create')]
        
        print(f"✅ Générateur importé avec succès")
        print(f"📊 Méthodes _create disponibles : {len(all_methods)}")
        
        for method in all_methods:
            print(f"   • {method}")
        
        # Vérifier les méthodes appelées dans generate_template
        print("\n🔍 Analyse du code source...")
        
        generator_file = Path("c:/Users/mushm/Documents/AR/scripts/generate_template.py")
        if generator_file.exists():
            content = generator_file.read_text(encoding='utf-8')
            
            # Rechercher les appels de méthodes _create
            create_calls = re.findall(r'self\.(_create_\w+)\(', content)
            unique_calls = list(set(create_calls))
            
            print(f"📋 Méthodes appelées dans generate_template : {len(unique_calls)}")
            
            missing_methods = []
            for called_method in unique_calls:
                if not hasattr(generator, called_method):
                    missing_methods.append(called_method)
                    print(f"   ❌ {called_method} - MANQUANTE")
                else:
                    print(f"   ✅ {called_method}")
            
            if missing_methods:
                print(f"\n🚨 {len(missing_methods)} méthode(s) manquante(s) détectée(s)")
                print("💡 Création des méthodes manquantes...")
                
                create_missing_methods(missing_methods, generator_file)
            else:
                print("\n✅ Toutes les méthodes requises sont présentes")
        
        else:
            print("❌ Fichier generate_template.py non trouvé")
    
    except Exception as e:
        print(f"❌ Erreur lors de l'analyse : {e}")


def create_missing_methods(missing_methods, generator_file):
    """Crée les méthodes manquantes dans le générateur."""
    
    method_templates = {
        '_create_config_sheet': '''
    def _create_config_sheet(self, pme_profile: bool = False) -> None:
        """Crée l'onglet de configuration EBIOS RM."""
        ws = self.wb.create_sheet("Config_EBIOS", 0)  # Première position
        
        # Titre principal
        ws["A1"] = "🔧 CONFIGURATION EBIOS RISK MANAGER"
        ws["A1"].font = Font(size=14, bold=True, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        ws.merge_cells("A1:F1")
        
        # Description
        ws["A2"] = "Configuration des paramètres EBIOS RM selon profil organisationnel"
        ws["A2"].font = Font(italic=True)
        
        # Section profil
        ws["A4"] = "📋 PROFIL ORGANISATIONNEL"
        ws["A4"].font = Font(size=12, bold=True)
        
        profile_text = "PME/TPE - Configuration simplifiée" if pme_profile else "Grande entreprise - Configuration complète"
        ws["A5"] = f"Type d'organisation : {profile_text}"
        
        # Instructions
        ws["A7"] = "📝 INSTRUCTIONS D'UTILISATION"
        ws["A7"].font = Font(size=12, bold=True)
        
        instructions = [
            "1. Renseignez les actifs dans l'Atelier 1",
            "2. Analysez les sources de risque dans l'Atelier 2", 
            "3. Définissez les scénarios dans l'Atelier 3",
            "4. Évaluez les mesures dans l'Atelier 4",
            "5. Consultez la synthèse pour les résultats"
        ]
        
        for i, instruction in enumerate(instructions, 8):
            ws.cell(row=i, column=1, value=instruction)
        
        logger.info("✅ Onglet de configuration créé")
''',
        
        '_create_synthese': '''
    def _create_synthese(self) -> None:
        """Crée l'onglet Synthèse avec indicateurs clés."""
        ws = self.wb.create_sheet("Synthese")
        
        # Titre
        ws["A1"] = "📊 SYNTHÈSE EXÉCUTIVE - ANALYSE DES RISQUES"
        ws["A1"].font = Font(size=14, bold=True, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
        ws.merge_cells("A1:F1")
        
        # Métriques principales
        ws["A3"] = "🎯 INDICATEURS CLÉS"
        ws["A3"].font = Font(size=12, bold=True)
        
        # En-têtes
        headers = ["Indicateur", "Valeur", "Statut", "Tendance"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col, value=header)
            cell.font = Font(bold=True)
            cell.fill = self.gray_fill
        
        # Données de synthèse
        metrics = [
            ["Nombre d'actifs analysés", "=COUNTA(Atelier1_Socle.A:A)-1", "En cours", "↗️"],
            ["Sources de risque identifiées", "=COUNTA(Atelier2_Sources.A:A)-1", "Complété", "→"],
            ["Scénarios évalués", "=COUNTA(Atelier3_Scenarios.A:A)-1", "En cours", "↗️"],
            ["Mesures planifiées", "=COUNTA(Atelier4_Operationnels.A:A)-1", "Planifié", "↗️"]
        ]
        
        for row_idx, metric_data in enumerate(metrics, 5):
            for col_idx, value in enumerate(metric_data, 1):
                ws.cell(row=row_idx, column=col_idx, value=value)
        
        logger.info("✅ Onglet de synthèse créé")
'''
    }
    
    print("🔧 Génération des méthodes manquantes...")
    
    # Lire le contenu actuel
    content = generator_file.read_text(encoding='utf-8')
    
    # Ajouter les méthodes manquantes avant la méthode main()
    main_pattern = r'(\ndef main\(\):)'
    
    methods_to_add = ""
    for method_name in missing_methods:
        if method_name in method_templates:
            methods_to_add += method_templates[method_name] + "\n"
            print(f"   ✅ {method_name} ajoutée")
        else:
            # Créer une méthode basique
            basic_method = f'''
    def {method_name}(self) -> None:
        """Méthode générée automatiquement - À implémenter."""
        logger.warning(f"Méthode {method_name} appelée mais pas encore implémentée")
        pass
'''
            methods_to_add += basic_method + "\n"
            print(f"   ⚠️  {method_name} créée (template basique)")
    
    # Insérer les nouvelles méthodes
    new_content = re.sub(main_pattern, methods_to_add + r'\1', content)
    
    # Sauvegarder le fichier corrigé
    backup_file = generator_file.with_suffix('.py.backup')
    generator_file.rename(backup_file)
    print(f"💾 Sauvegarde créée : {backup_file}")
    
    generator_file.write_text(new_content, encoding='utf-8')
    print(f"✅ Fichier corrigé : {generator_file}")
    
    print("\n🎯 Correction terminée ! Vous pouvez maintenant exécuter :")
    print("   python generate_template.py")


if __name__ == "__main__":
    analyze_generator_methods()
