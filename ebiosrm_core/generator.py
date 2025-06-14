"""Risk assessment calculation and report generation."""

import logging
from pathlib import Path
from typing import Dict, Any

from . import loader, exporters
from .models import Asset, Threat

logger = logging.getLogger(__name__)


def calculate_risk_levels(
    assets: list[Asset], threats: list[Threat]
) -> list[Dict[str, Any]]:
    """Calculate risk levels for all threat-asset combinations."""
    results = []

    for threat in threats:
        # For each threat, calculate risk against all assets
        # In a real implementation, you'd filter based on threat-asset relationships
        max_severity = max(asset.severity_score() for asset in assets)

        risk_level = threat.risk_level(max_severity)
        likelihood = threat.likelihood_score()

        result = {
            "threat_id": threat.sr_id,
            "threat_ov": threat.ov_id,
            "strategic_path": threat.strategic_path,
            "operational_steps": threat.operational_steps,
            "likelihood_score": likelihood,
            "max_severity": max_severity,  # Add this field for test compatibility
            "severity_score": max_severity,
            "risk_level": risk_level,
            "affected_assets": [asset.id for asset in assets],  # Simplified
        }

        results.append(result)

    return results


def calculate_risks(*args, **kwargs):  # pragma: no cover
    """Backward-compat wrapper kept for tests."""
    return calculate_risk_levels(*args, **kwargs)


def run(
    cfg_dir: Path,
    out_dir: Path,
    fmt: str = "xlsx",
    pme_profile: bool = False,
    output_filename: str = "ebios_risk_assessment.xlsx",
) -> None:
    """Main generator function."""
    logger.info(f"Loading configuration from {cfg_dir}")

    # Load all data
    assets, threats, settings, risk_sources, objectives, stakeholders, measures = (
        loader.load_all(cfg_dir)
    )

    logger.info(f"Loaded {len(assets)} assets, {len(threats)} threats")
    logger.info(
        f"Additional components: {len(risk_sources)} risk sources, {len(objectives)} objectives"
    )
    logger.info(f"Stakeholders: {len(stakeholders)}, Measures: {len(measures)}")

    # Calculate risk levels
    risk_results = calculate_risk_levels(assets, threats)

    # Prepare export data
    export_data = {
        "assets": [asset.model_dump() for asset in assets],
        "threats": [threat.model_dump() for threat in threats],
        "risk_sources": [source.model_dump() for source in risk_sources],
        "objectives": [obj.model_dump() for obj in objectives],
        "stakeholders": [stakeholder.model_dump() for stakeholder in stakeholders],
        "measures": [measure.model_dump() for measure in measures],
        "risk_results": risk_results,
        "settings": settings.model_dump(),
    }

    # Ensure output directory exists
    out_path = Path(out_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    # Determine output filename if not provided
    if not output_filename or output_filename == "ebios_risk_assessment.xlsx":
        default_name = "ebios_risk_assessment"
        if fmt.lower() in {"xlsx", "excel"}:
            output_filename = f"{default_name}.xlsx"
        elif fmt.lower() == "json":
            output_filename = f"{default_name}.json"
        elif fmt.lower() in {"md", "markdown"}:
            output_filename = f"{default_name}.md"
        else:
            output_filename = f"{default_name}.{fmt.lower()}"

    # Export results
    output_file_path = out_path / output_filename

    if fmt.lower() in {"xlsx", "excel"}:
        exporters.export_excel(export_data, output_file_path, pme_profile=pme_profile)
    elif fmt.lower() == "json":
        exporters.export_json(export_data, output_file_path)
    elif fmt.lower() in {"md", "markdown"}:
        exporters.export_markdown(export_data, output_file_path)
    else:
        raise ValueError(f"Unsupported format: {fmt}")

    logger.info(f"Report generated successfully in {output_file_path}")
