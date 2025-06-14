"""Data loading and validation functions."""

import csv
import yaml
from pathlib import Path
from typing import Tuple

from .models import Asset, Threat, RiskSource, TargetedObjective, Stakeholder, SecurityMeasure, Settings


def load_assets(config_dir: Path) -> list[Asset]:
    """Load assets from CSV file."""
    assets_file = config_dir / "assets.csv"
    if not assets_file.exists():
        raise FileNotFoundError(f"Assets file not found: {assets_file}")
    
    assets = []
    with open(assets_file, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        
        # Validate headers
        if reader.fieldnames is None:
            raise ValueError(f"No headers found in {assets_file}")
        
        for row_num, row in enumerate(reader):
            try:
                # Clean the row data - ensure all keys are strings
                clean_row = {}
                for key, value in row.items():
                    if key is None or not isinstance(key, str) or not key.strip():
                        continue  # Skip invalid keys
                    
                    clean_key = str(key).strip()
                    clean_value = str(value).strip() if value is not None else ""
                    clean_row[clean_key] = clean_value
                
                assets.append(Asset(**clean_row))
            except Exception as e:
                raise ValueError(f"Error processing row {row_num + 1} in {assets_file}: {e}")
    
    return assets


def load_threats(config_dir: Path) -> list[Threat]:
    """Load threats from CSV file."""
    threats_file = config_dir / "threats.csv"
    if not threats_file.exists():
        raise FileNotFoundError(f"Threats file not found: {threats_file}")
    
    threats = []
    with open(threats_file, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        
        # Validate headers
        if reader.fieldnames is None:
            raise ValueError(f"No headers found in {threats_file}")
        
        # Check for invalid headers (None or non-string) and filter them out
        valid_headers = [h for h in reader.fieldnames if h is not None and isinstance(h, str) and h.strip()]
        invalid_headers = [h for h in reader.fieldnames if h is None or not isinstance(h, str) or not h.strip()]
        
        if invalid_headers:
            print(f"Warning: Ignoring invalid headers in {threats_file}: {invalid_headers}")
        
        for row_num, row in enumerate(reader):
            try:
                # Clean the row data - only use valid headers and ensure all keys are strings
                clean_row = {}
                for key, value in row.items():
                    if key is None or not isinstance(key, str) or not key.strip():
                        continue  # Skip invalid keys
                    
                    clean_key = str(key).strip()
                    if clean_key in valid_headers:
                        clean_value = str(value).strip() if value is not None else ""
                        clean_row[clean_key] = clean_value
                
                # Ensure we have the required fields
                required_fields = ["sr_id", "ov_id", "strategic_path", "operational_steps"]
                missing_fields = [field for field in required_fields if field not in clean_row]
                if missing_fields:
                    raise ValueError(f"Missing required fields: {missing_fields}")
                
                # Handle optional list fields
                if "risk_sources" in clean_row and clean_row["risk_sources"]:
                    clean_row["risk_sources"] = [x.strip() for x in clean_row["risk_sources"].split(",")]
                if "targeted_objectives" in clean_row and clean_row["targeted_objectives"]:
                    clean_row["targeted_objectives"] = [x.strip() for x in clean_row["targeted_objectives"].split(",")]
                
                threats.append(Threat(**clean_row))
            except Exception as e:
                raise ValueError(f"Error processing row {row_num + 1} in {threats_file}: {e}")
    
    return threats


def load_risk_sources(config_dir: Path) -> list[RiskSource]:
    """Load risk sources from CSV file."""
    sources_file = config_dir / "risk_sources.csv"
    if not sources_file.exists():
        return []  # Optional file
    
    sources = []
    with open(sources_file, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            sources.append(RiskSource(**row))
    
    return sources


def load_objectives(config_dir: Path) -> list[TargetedObjective]:
    """Load targeted objectives from CSV file."""
    objectives_file = config_dir / "objectives.csv"
    if not objectives_file.exists():
        return []  # Optional file
    
    objectives = []
    with open(objectives_file, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        
        # Validate headers
        if reader.fieldnames is None:
            raise ValueError(f"No headers found in {objectives_file}")
        
        for row_num, row in enumerate(reader):
            try:
                # Clean the row data - ensure all keys are strings
                clean_row = {}
                for key, value in row.items():
                    if key is None or not isinstance(key, str) or not key.strip():
                        continue  # Skip invalid keys
                    
                    clean_key = str(key).strip()
                    clean_value = str(value).strip() if value is not None else ""
                    clean_row[clean_key] = clean_value
                
                # Handle list fields
                if "target_assets" in clean_row and clean_row["target_assets"]:
                    clean_row["target_assets"] = [x.strip() for x in clean_row["target_assets"].split(",")]
                if "attack_scenarios" in clean_row and clean_row["attack_scenarios"]:
                    clean_row["attack_scenarios"] = [x.strip() for x in clean_row["attack_scenarios"].split(",")]
                
                objectives.append(TargetedObjective(**clean_row))
            except Exception as e:
                raise ValueError(f"Error processing row {row_num + 1} in {objectives_file}: {e}")
    
    return objectives


def load_stakeholders(config_dir: Path) -> list[Stakeholder]:
    """Load stakeholders from CSV file."""
    stakeholders_file = config_dir / "stakeholders.csv"
    if not stakeholders_file.exists():
        return []  # Optional file
    
    stakeholders = []
    with open(stakeholders_file, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        
        # Validate headers
        if reader.fieldnames is None:
            return []  # Empty file is OK for optional files
        
        for row_num, row in enumerate(reader):
            try:
                # Clean the row data - ensure all keys are strings
                clean_row = {}
                for key, value in row.items():
                    if key is None or not isinstance(key, str) or not key.strip():
                        continue  # Skip invalid keys
                    
                    clean_key = str(key).strip()
                    clean_value = str(value).strip() if value is not None else ""
                    clean_row[clean_key] = clean_value
                
                # Handle list fields
                if "responsibilities" in clean_row and clean_row["responsibilities"]:
                    clean_row["responsibilities"] = [x.strip() for x in clean_row["responsibilities"].split(",")]
                
                stakeholders.append(Stakeholder(**clean_row))
            except Exception as e:
                raise ValueError(f"Error processing row {row_num + 1} in {stakeholders_file}: {e}")
    
    return stakeholders


def load_measures(config_dir: Path) -> list[SecurityMeasure]:
    """Load security measures from CSV file."""
    measures_file = config_dir / "measures.csv"
    if not measures_file.exists():
        return []  # Optional file
    
    measures = []
    with open(measures_file, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        
        # Validate headers
        if reader.fieldnames is None:
            return []  # Empty file is OK for optional files
        
        for row_num, row in enumerate(reader):
            try:
                # Clean the row data - ensure all keys are strings
                clean_row = {}
                for key, value in row.items():
                    if key is None or not isinstance(key, str) or not key.strip():
                        continue  # Skip invalid keys
                    
                    clean_key = str(key).strip()
                    clean_value = str(value).strip() if value is not None else ""
                    clean_row[clean_key] = clean_value
                
                # Handle list fields
                if "target_threats" in clean_row and clean_row["target_threats"]:
                    clean_row["target_threats"] = [x.strip() for x in clean_row["target_threats"].split(",")]
                
                measures.append(SecurityMeasure(**clean_row))
            except Exception as e:
                raise ValueError(f"Error processing row {row_num + 1} in {measures_file}: {e}")
    
    return measures


def load_settings(config_dir: Path) -> Settings:
    """Load settings from YAML file."""
    settings_file = config_dir / "settings.yaml"
    if not settings_file.exists():
        return Settings()  # Use defaults
    
    with open(settings_file, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    
    return Settings(**data)


def load_all(config_dir: Path) -> Tuple[list[Asset], list[Threat], Settings, list[RiskSource], list[TargetedObjective], list[Stakeholder], list[SecurityMeasure]]:
    """Load all configuration data."""
    config_path = Path(config_dir)
    
    assets = load_assets(config_path)
    threats = load_threats(config_path)
    settings = load_settings(config_path)
    risk_sources = load_risk_sources(config_path)
    objectives = load_objectives(config_path)
    stakeholders = load_stakeholders(config_path)
    measures = load_measures(config_path)
    
    return assets, threats, settings, risk_sources, objectives, stakeholders, measures
