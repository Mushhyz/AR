"""Test configuration and shared fixtures."""

from __future__ import annotations

import csv
from pathlib import Path

import pytest
import yaml

from ebiosrm_core.models import Asset, Threat, Settings


@pytest.fixture
def sample_assets():
    """Sample assets for testing."""
    return [
        Asset(id="A001", type="Data", label="Customer Database", criticality="Critical"),
        Asset(id="A002", type="System", label="Web Server", criticality="High"),
        Asset(id="A003", type="Data", label="Application Logs", criticality="Low"),
        Asset(id="A004", type="System", label="Database Server", criticality="Critical"),
        Asset(id="A005", type="Network", label="Internal Network", criticality="Medium"),
    ]


@pytest.fixture
def sample_threats():
    """Sample threats for testing."""
    return [
        Threat(
            sr_id="SR001",
            ov_id="OV001", 
            strategic_path="External Cyber Attack",
            operational_steps="Reconnaissance:Low,Initial Access:Medium,Privilege Escalation:High,Data Exfiltration:Critical"
        ),
        Threat(
            sr_id="SR002",
            ov_id="OV002",
            strategic_path="Insider Threat", 
            operational_steps="Credential Abuse:High,Data Access:Critical,Data Theft:Critical"
        ),
        Threat(
            sr_id="SR003",
            ov_id="OV003",
            strategic_path="Supply Chain Attack",
            operational_steps="Third Party Compromise:Medium,Lateral Movement:High,Persistence:High"
        ),
    ]


@pytest.fixture
def sample_objectives():
    """Sample objectives for testing - matching actual CSV structure."""
    from ebiosrm_core.models import Objective
    return [
        Objective(
            id="OBJ001",
            label="Data Theft",
            target_assets="A001",
            business_impact="Critical",
            attack_scenarios="SR001,SR002"
        ),
        Objective(
            id="OBJ002",
            label="Service Disruption", 
            target_assets="A002",
            business_impact="High",
            attack_scenarios="SR006"
        ),
        Objective(
            id="OBJ003",
            label="Financial Fraud",
            target_assets="A003",
            business_impact="Critical",
            attack_scenarios="SR002,SR005"
        ),
    ]


@pytest.fixture
def sample_settings():
    """Sample settings for testing."""
    return Settings(
        excel_template="ebiosrm_template.xlsx",
        output_dir="output/",
        severity_scale=["Low", "Medium", "High", "Critical"],
        likelihood_scale=["One-shot", "Occasional", "Probable", "Systematic"]
    )


@pytest.fixture
def temp_config_dir(tmp_path, sample_assets, sample_threats, sample_objectives, sample_settings):
    """Create temporary configuration directory with test data."""
    config_dir = tmp_path / "config"
    config_dir.mkdir()
    
    # Create assets.csv with proper string formatting
    assets_file = config_dir / "assets.csv"
    with open(assets_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["id", "type", "label", "criticality"])
        for asset in sample_assets:
            writer.writerow([str(asset.id), str(asset.type), str(asset.label), str(asset.criticality)])
    
    # Create threats.csv with proper string formatting
    threats_file = config_dir / "threats.csv"
    with open(threats_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["sr_id", "ov_id", "strategic_path", "operational_steps"])
        for threat in sample_threats:
            writer.writerow([
                str(threat.sr_id), str(threat.ov_id), 
                str(threat.strategic_path), str(threat.operational_steps)
            ])
    
    # Create objectives.csv with the working structure
    objectives_file = config_dir / "objectives.csv"
    with open(objectives_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["id", "label", "target_assets", "business_impact", "attack_scenarios"])
        for objective in sample_objectives:
            writer.writerow([
                str(objective.id), str(objective.label), 
                str(objective.target_assets), str(objective.business_impact),
                str(objective.attack_scenarios)
            ])

    # Create minimal optional files to prevent FileNotFoundError
    for optional_file in ["risk_sources.csv", "stakeholders.csv", "measures.csv"]:
        file_path = config_dir / optional_file
        with open(file_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if optional_file == "risk_sources.csv":
                writer.writerow(["id", "label", "category", "motivation", "capability_level", "resources"])
            elif optional_file == "stakeholders.csv":
                writer.writerow(["id", "name", "type", "role", "responsibilities", "contact_info"])
            elif optional_file == "measures.csv":
                writer.writerow(["id", "label", "type", "description", "effectiveness", "implementation_cost", "target_threats", "responsible_stakeholder"])
    
    # Create settings.yaml
    settings_file = config_dir / "settings.yaml"
    with open(settings_file, "w", encoding="utf-8") as f:
        yaml.dump(sample_settings.dict(), f)
    
    return config_dir


@pytest.fixture
def cli_runner():
    """Provide a CLI test runner."""
    from typer.testing import CliRunner
    return CliRunner()


@pytest.fixture
def temp_output_dir(tmp_path):
    """Create temporary output directory for testing."""
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    return output_dir


@pytest.fixture
def sample_excel_template(tmp_path):
    """Create a minimal Excel template for testing."""
    from openpyxl import Workbook
    
    template_path = tmp_path / "ebiosrm_template.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "EBIOS_RM"
    
    # Add some basic headers
    ws['A1'] = "Asset ID"
    ws['B1'] = "Asset Type"
    ws['C1'] = "Asset Label"
    ws['D1'] = "Criticality"
    
    wb.save(template_path)
    return template_path


@pytest.fixture
def complete_test_environment(temp_config_dir, temp_output_dir, sample_excel_template):
    """Provide a complete test environment with all necessary files."""
    return {
        "config_dir": temp_config_dir,
        "output_dir": temp_output_dir,
        "template": sample_excel_template
    }


@pytest.fixture
def mock_cli_app():
    """Provide a mock CLI app for testing when the real CLI isn't available."""
    import typer
    
    app = typer.Typer()
    
    @app.command()
    def validate():
        """Mock validate command."""
        typer.echo("Validation successful")
        return True
    
    @app.command()
    def export(fmt: str = "json"):
        """Mock export command."""
        typer.echo(f"Export to {fmt} format successful")
        return True
    
    return app


@pytest.fixture
def project_root():
    """Get the project root directory."""
    return Path(__file__).parent.parent


@pytest.fixture
def debug_module_info(project_root):
    """Debug fixture to help identify the correct module structure."""
    import sys
    import importlib.util
    
    info = {
        "project_root": str(project_root),
        "python_path": sys.path,
        "available_modules": []
    }
    
    # Check for ebiosrm modules
    for module_name in ["ebiosrm", "ebiosrm_core", "ebiosrm_generator"]:
        try:
            spec = importlib.util.find_spec(module_name)
            if spec:
                info["available_modules"].append({
                    "name": module_name,
                    "location": spec.origin,
                    "submodule_search_paths": spec.submodule_search_paths
                })
        except (ImportError, ModuleNotFoundError):
            pass
    
    return info


@pytest.fixture
def debug_csv_data(temp_config_dir):
    """Debug fixture to inspect CSV data structure."""
    import csv
    
    threats_file = temp_config_dir / "threats.csv"
    debug_info = {
        "file_exists": threats_file.exists(),
        "file_content": [],
        "csv_rows": []
    }
    
    if threats_file.exists():
        # Read raw file content
        with open(threats_file, "r", encoding="utf-8") as f:
            debug_info["file_content"] = f.readlines()
        
        # Read CSV structure
        with open(threats_file, "r", newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for i, row in enumerate(reader):
                debug_info["csv_rows"].append({
                    "row_number": i,
                    "keys": list(row.keys()),
                    "key_types": {k: type(k).__name__ for k in row.keys()},
                    "values": dict(row),
                    "value_types": {k: type(v).__name__ for k, v in row.items()}
                })
    
    return debug_info


@pytest.fixture
def csv_debugging_helper():
    """Helper to debug CSV loading issues in real config files."""
    def debug_csv(file_path):
        import csv
        print(f"Debugging {file_path}")
        
        with open(file_path, 'r', encoding='utf-8') as f:
            # Read raw content
            content = f.read()
            print(f"Raw content (first 200 chars): {repr(content[:200])}")
        
        with open(file_path, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            print(f"Headers: {reader.fieldnames}")
            print(f"Header types: {[type(h) for h in reader.fieldnames]}")
            
            for i, row in enumerate(reader):
                if i < 3:  # Show first 3 rows
                    print(f"Row {i}: {dict(row)}")
                    print(f"Key types: {[(k, type(k)) for k in row.keys()]}")
                break  # Only show first row for debugging
    
    return debug_csv


@pytest.fixture
def real_config_debugger():
    """Debug fixture specifically for the real config directory."""
    def debug_real_config():
        from pathlib import Path
        import csv
        
        config_path = Path("config")
        threats_file = config_path / "threats.csv"
        
        if not threats_file.exists():
            print(f"❌ {threats_file} does not exist")
            return
        
        print(f"🔍 Debugging {threats_file}")
        
        # Check raw file content
        with open(threats_file, 'rb') as f:
            raw_bytes = f.read(100)
            print(f"Raw bytes: {raw_bytes}")
        
        # Check text content
        with open(threats_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()[:5]
            print(f"First 5 lines: {lines}")
        
        # Check CSV parsing
        try:
            with open(threats_file, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                print(f"Headers: {reader.fieldnames}")
                print(f"Header types: {[(h, type(h)) for h in reader.fieldnames]}")
                
                for i, row in enumerate(reader):
                    print(f"Row {i} keys: {[(k, type(k)) for k in row.keys()]}")
                    print(f"Row {i} values: {dict(row)}")
                    if i >= 2:  # Only show first 3 rows
                        break
        except Exception as e:
            print(f"❌ CSV parsing error: {e}")
    
    return debug_real_config


@pytest.fixture
def csv_header_cleaner():
    """Fixture to clean CSV headers and ensure they are valid strings."""
    def clean_csv_file(file_path):
        import csv
        import tempfile
        
        # Read original file with different encodings to handle BOM
        encodings_to_try = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']
        content = None
        
        for encoding in encodings_to_try:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    content = f.read()
                break
            except UnicodeDecodeError:
                continue
        
        if content is None:
            raise ValueError(f"Could not decode {file_path} with any supported encoding")
        
        # Remove BOM if present
        if content.startswith('\ufeff'):
            content = content[1:]
            
        # Write to temp file and read back
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv', encoding='utf-8') as temp_f:
            temp_f.write(content)
            temp_path = temp_f.name
            
        try:
            # Read and validate CSV structure
            with open(temp_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                headers = reader.fieldnames
                
                # Check for problematic headers
                if headers is None:
                    raise ValueError("No headers found in CSV")
                    
                clean_headers = []
                for i, h in enumerate(headers):
                    if h is None:
                        clean_headers.append(f"unnamed_column_{i}")
                    elif not isinstance(h, str):
                        clean_headers.append(str(h).strip())
                    else:
                        clean_headers.append(h.strip())
                
                # Remove empty headers
                clean_headers = [h for h in clean_headers if h]
                
                # Read all rows
                rows = []
                for row in reader:
                    clean_row = {}
                    for old_h, new_h in zip(headers, clean_headers):
                        if old_h is None:
                            continue
                        value = row.get(old_h, "")
                        clean_row[new_h] = str(value).strip() if value is not None else ""
                    rows.append(clean_row)
            
            # Write cleaned version back to original file
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                if rows and clean_headers:
                    writer = csv.DictWriter(f, fieldnames=clean_headers)
                    writer.writeheader()
                    writer.writerows(rows)
            
            return clean_headers
            
        finally:
            # Clean up temp file
            import os
            try:
                os.unlink(temp_path)
            except FileNotFoundError:
                pass
    
    return clean_csv_file


@pytest.fixture
def validate_real_config():
    """Validate and fix the real config directory CSV files."""
    def validate_config_dir():
        from pathlib import Path
        config_path = Path("config")
        
        if not config_path.exists():
            print("❌ Config directory does not exist")
            return False
        
        csv_files = ["threats.csv", "assets.csv"]
        issues = []
        
        for csv_file in csv_files:
            file_path = config_path / csv_file
            if file_path.exists():
                try:
                    # Try to validate the CSV
                    import csv
                    with open(file_path, 'r', newline='', encoding='utf-8') as f:
                        reader = csv.DictReader(f)
                        headers = reader.fieldnames
                        
                        if headers is None:
                            issues.append(f"{csv_file}: No headers found")
                            continue
                        
                        # Check for problematic headers
                        for i, header in enumerate(headers):
                            if header is None:
                                issues.append(f"{csv_file}: None header at position {i}")
                            elif not isinstance(header, str):
                                issues.append(f"{csv_file}: Non-string header '{header}' (type: {type(header)})")
                        
                        # Try to read first row
                        try:
                            first_row = next(reader, None)
                            if first_row:
                                for key in first_row.keys():
                                    if not isinstance(key, str):
                                        issues.append(f"{csv_file}: Non-string key '{key}' (type: {type(key)})")
                        except Exception as e:
                            issues.append(f"{csv_file}: Error reading first row: {e}")
                            
                except Exception as e:
                    issues.append(f"{csv_file}: Error opening file: {e}")
        
        if issues:
            print("❌ CSV validation issues found:")
            for issue in issues:
                print(f"  - {issue}")
            return False
        else:
            print("✅ CSV files validated successfully")
            return True
    
    return validate_config_dir


@pytest.fixture
def csv_file_cleaner():
    """Fixture to clean problematic CSV files."""
    def clean_csv_file(file_path):
        """Clean a CSV file to remove common issues."""
        import csv
        import tempfile
        
        # Read with BOM handling
        with open(file_path, 'r', encoding='utf-8-sig', newline='') as f:
            content = f.read()
        
        # Remove any remaining BOM characters
        content = content.replace('\ufeff', '')
        
        # Write to temporary file
        temp_file = file_path.with_suffix('.cleaned.csv')
        with open(temp_file, 'w', encoding='utf-8', newline='') as f:
            f.write(content)
        
        # Validate the cleaned file
        with open(temp_file, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            headers = reader.fieldnames
            
            # Check headers
            if not headers or None in headers:
                raise ValueError("Invalid or missing headers in CSV")
            
            # Ensure all headers are strings
            clean_headers = [str(h).strip() for h in headers if h is not None]
            
            # Re-read and clean data
            f.seek(0)
            reader = csv.DictReader(f)
            rows = []
            for row in reader:
                clean_row = {}
                for old_key, new_key in zip(headers, clean_headers):
                    value = row.get(old_key, "")
                    clean_row[new_key] = str(value).strip() if value is not None else ""
                rows.append(clean_row)
        
        # Write final cleaned version
        with open(temp_file, 'w', newline='', encoding='utf-8') as f:
            if rows:
                writer = csv.DictWriter(f, fieldnames=clean_headers)
                writer.writeheader()
                writer.writerows(rows)
        
        return temp_file
    
    return clean_csv_file


@pytest.fixture
def test_csv_with_embedded_commas(tmp_path):
    """Test fixture for CSV files with embedded commas in fields."""
    csv_content = '''sr_id,ov_id,strategic_path,operational_steps
SR001,OV001,External Attack,"Step1:High,Step2:Medium,Step3:Low"
SR002,OV002,Internal Threat,"Step1:Critical,Step2:High"'''
    
    csv_file = tmp_path / "test_threats.csv"
    with open(csv_file, 'w', encoding='utf-8') as f:
        f.write(csv_content)
    
    return csv_file


@pytest.fixture
def working_config_validator():
    """Fixture to validate that config matches working structure."""
    def validate_working_config(config_dir):
        from pathlib import Path
        import csv
        
        validation_results = {
            "assets": {"exists": False, "valid": False, "count": 0},
            "threats": {"exists": False, "valid": False, "count": 0},
            "objectives": {"exists": False, "valid": False, "count": 0}
        }
        
        # Check assets.csv
        assets_file = Path(config_dir) / "assets.csv"
        if assets_file.exists():
            validation_results["assets"]["exists"] = True
            try:
                with open(assets_file, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    expected_headers = ["id", "type", "label", "criticality"]
                    if reader.fieldnames == expected_headers:
                        validation_results["assets"]["valid"] = True
                        validation_results["assets"]["count"] = sum(1 for _ in reader)
            except Exception:
                pass
        
        # Check threats.csv
        threats_file = Path(config_dir) / "threats.csv"
        if threats_file.exists():
            validation_results["threats"]["exists"] = True
            try:
                with open(threats_file, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    expected_headers = ["sr_id", "ov_id", "strategic_path", "operational_steps"]
                    if reader.fieldnames == expected_headers:
                        validation_results["threats"]["valid"] = True
                        validation_results["threats"]["count"] = sum(1 for _ in reader)
            except Exception:
                pass
        
        # Check objectives.csv
        objectives_file = Path(config_dir) / "objectives.csv"
        if objectives_file.exists():
            validation_results["objectives"]["exists"] = True
            try:
                with open(objectives_file, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    expected_headers = ["id", "label", "target_assets", "business_impact", "attack_scenarios"]
                    if reader.fieldnames == expected_headers:
                        validation_results["objectives"]["valid"] = True
                        validation_results["objectives"]["count"] = sum(1 for _ in reader)
            except Exception:
                pass
        
        return validation_results
    
    return validate_working_config
