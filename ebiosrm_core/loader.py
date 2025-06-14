"""Data loading and validation functions."""

import csv
import yaml
import pandas as pd
from pathlib import Path
from typing import Tuple

from .models import (
    Asset,
    Threat,
    RiskSource,
    TargetedObjective,
    Stakeholder,
    SecurityMeasure,
    Settings,
)


def load_referentials(config_dir: Path) -> pd.DataFrame:
    """Load and consolidate all referential CSV files.

    Handles multiple versions by keeping only the highest version number.
    Supports both filename versioning (_v{version}) and column versioning.

    Args:
        config_dir: Configuration directory path

    Returns:
        Consolidated DataFrame with all referential controls

    References:
        - File naming best practices DATACC: https://datacc.org/resources/file-naming-versioning/
        - Find duplicates in CSV w/ Python: https://stackoverflow.com/questions/find-duplicates-csv
    """
    referentials_dir = config_dir / "referentials"
    if not referentials_dir.exists():
        return pd.DataFrame()

    # Find all CSV files in referentials directory
    csv_files = list(referentials_dir.glob("*.csv"))
    if not csv_files:
        return pd.DataFrame()

    all_dataframes = []

    for csv_file in csv_files:
        if csv_file.name.startswith("_"):  # Skip manifest files
            continue

        try:
            # Load CSV with proper encoding handling
            df = pd.read_csv(csv_file, encoding="utf-8-sig")

            # Extract version from filename if not in column
            if "version" not in df.columns:
                file_name = csv_file.stem
                if "_v" in file_name:
                    _, version = file_name.rsplit("_v", 1)
                    df["version"] = version
                else:
                    df["version"] = "1.0"

            # Add source file for tracking
            df["source_file"] = csv_file.name

            all_dataframes.append(df)

        except Exception as e:
            print(f"Warning: Could not load {csv_file}: {e}")
            continue

    if not all_dataframes:
        return pd.DataFrame()

    # Concatenate all dataframes
    df_combined = pd.concat(all_dataframes, ignore_index=True)

    # Handle duplicates by keeping highest version
    if "id" in df_combined.columns and "version" in df_combined.columns:
        # Sort by version (descending) and drop duplicates keeping first (highest version)
        df_combined = df_combined.sort_values("version", ascending=False)
        df_combined = df_combined.drop_duplicates(subset=["id"], keep="first")

    # Reset index and ensure required columns exist
    df_combined = df_combined.reset_index(drop=True)

    # Ensure required columns exist with defaults
    required_columns = ["id", "label", "category", "description", "criticality"]
    for col in required_columns:
        if col not in df_combined.columns:
            df_combined[col] = ""

    # Ensure cross-reference columns exist
    xref_columns = ["xref_iso", "xref_nist"]
    for col in xref_columns:
        if col not in df_combined.columns:
            df_combined[col] = ""

    return df_combined


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

        # Check required columns before processing rows
        required = {"id", "type", "label", "criticality"}
        if not required.issubset(set(reader.fieldnames)):
            missing = required - set(reader.fieldnames)
            raise ValueError(f"Missing columns: {', '.join(sorted(missing))}")

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
                raise ValueError(
                    f"Error processing row {row_num + 1} in {assets_file}: {e}"
                )

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
        valid_headers = [
            h
            for h in reader.fieldnames
            if h is not None and isinstance(h, str) and h.strip()
        ]
        invalid_headers = [
            h
            for h in reader.fieldnames
            if h is None or not isinstance(h, str) or not h.strip()
        ]

        if invalid_headers:
            print(
                f"Warning: Ignoring invalid headers in {threats_file}: {invalid_headers}"
            )

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
                required_fields = [
                    "sr_id",
                    "ov_id",
                    "strategic_path",
                    "operational_steps",
                ]
                missing_fields = [
                    field for field in required_fields if field not in clean_row
                ]
                if missing_fields:
                    raise ValueError(f"Missing required fields: {missing_fields}")

                # Handle optional list fields
                if "risk_sources" in clean_row and clean_row["risk_sources"]:
                    clean_row["risk_sources"] = [
                        x.strip() for x in clean_row["risk_sources"].split(",")
                    ]
                if (
                    "targeted_objectives" in clean_row
                    and clean_row["targeted_objectives"]
                ):
                    clean_row["targeted_objectives"] = [
                        x.strip() for x in clean_row["targeted_objectives"].split(",")
                    ]

                threats.append(Threat(**clean_row))
            except Exception as e:
                raise ValueError(
                    f"Error processing row {row_num + 1} in {threats_file}: {e}"
                )

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
                    clean_row["target_assets"] = [
                        x.strip() for x in clean_row["target_assets"].split(",")
                    ]
                if "attack_scenarios" in clean_row and clean_row["attack_scenarios"]:
                    clean_row["attack_scenarios"] = [
                        x.strip() for x in clean_row["attack_scenarios"].split(",")
                    ]

                objectives.append(TargetedObjective(**clean_row))
            except Exception as e:
                raise ValueError(
                    f"Error processing row {row_num + 1} in {objectives_file}: {e}"
                )

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
                    clean_row["responsibilities"] = [
                        x.strip() for x in clean_row["responsibilities"].split(",")
                    ]

                stakeholders.append(Stakeholder(**clean_row))
            except Exception as e:
                raise ValueError(
                    f"Error processing row {row_num + 1} in {stakeholders_file}: {e}"
                )

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
                    clean_row["target_threats"] = [
                        x.strip() for x in clean_row["target_threats"].split(",")
                    ]

                measures.append(SecurityMeasure(**clean_row))
            except Exception as e:
                raise ValueError(
                    f"Error processing row {row_num + 1} in {measures_file}: {e}"
                )

    return measures


def load_settings(config_dir: Path) -> Settings:
    """Load settings from YAML file."""
    settings_file = config_dir / "settings.yaml"
    if not settings_file.exists():
        return Settings()  # Use defaults

    with open(settings_file, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    return Settings(**data)


def load_all(
    config_dir: Path,
) -> Tuple[
    list[Asset],
    list[Threat],
    Settings,
    list[RiskSource],
    list[TargetedObjective],
    list[Stakeholder],
    list[SecurityMeasure],
]:
    """Load all configuration data.

    Returns:
        Tuple with all loaded data (7 components)
    """
    config_path = Path(config_dir)

    assets = load_assets(config_path)
    threats = load_threats(config_path)
    settings = load_settings(config_path)
    risk_sources = load_risk_sources(config_path)
    objectives = load_objectives(config_path)
    stakeholders = load_stakeholders(config_path)
    measures = load_measures(config_path)

    return assets, threats, settings, risk_sources, objectives, stakeholders, measures


def load_all_with_referentials(
    config_dir: Path,
) -> Tuple[
    list[Asset],
    list[Threat],
    Settings,
    list[RiskSource],
    list[TargetedObjective],
    list[Stakeholder],
    list[SecurityMeasure],
    pd.DataFrame,
]:
    """Load all configuration data including referentials.

    Returns:
        Tuple with all loaded data plus referentials DataFrame (8 components)
    """
    config_path = Path(config_dir)

    assets = load_assets(config_path)
    threats = load_threats(config_path)
    settings = load_settings(config_path)
    risk_sources = load_risk_sources(config_path)
    objectives = load_objectives(config_path)
    stakeholders = load_stakeholders(config_path)
    measures = load_measures(config_path)
    referentials = load_referentials(config_path)

    return (
        assets,
        threats,
        settings,
        risk_sources,
        objectives,
        stakeholders,
        measures,
        referentials,
    )
