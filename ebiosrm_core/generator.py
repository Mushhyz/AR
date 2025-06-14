"""Risk assessment calculation and report generation."""

import logging
from pathlib import Path
from typing import Dict, Any

from . import loader, exporters
from .models import Asset, Threat

logger = logging.getLogger(__name__)


def calculate_risk_levels(assets: list[Asset], threats: list[Threat]) -> list[Dict[str, Any]]:
    """Calculate risk levels for all threat-asset combinations."""
    results = []
    
    # Create asset lookup for faster access
    asset_lookup = {asset.id: asset for asset in assets}
    
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
            "severity_score": max_severity,
            "risk_level": risk_level,
            "affected_assets": [asset.id for asset in assets],  # Simplified
        }
        
        results.append(result)
    
    return results


def run(cfg_dir: Path, out_dir: Path, fmt: str = "xlsx") -> None:
    """Main generator function."""
    logger.info(f"Loading configuration from {cfg_dir}")
    
    # Load all data
    assets, threats, settings, risk_sources, objectives, stakeholders, measures = loader.load_all(cfg_dir)
    
    logger.info(f"Loaded {len(assets)} assets, {len(threats)} threats")
    logger.info(f"Additional components: {len(risk_sources)} risk sources, {len(objectives)} objectives")
    logger.info(f"Stakeholders: {len(stakeholders)}, Measures: {len(measures)}")
    
    # Calculate risk levels
    risk_results = calculate_risk_levels(assets, threats)
    
    # Prepare export data
    export_data = {
        "assets": [asset.dict() for asset in assets],
        "threats": [threat.dict() for threat in threats],
        "risk_sources": [source.dict() for source in risk_sources],
        "objectives": [obj.dict() for obj in objectives],
        "stakeholders": [stakeholder.dict() for stakeholder in stakeholders],
        "measures": [measure.dict() for measure in measures],
        "risk_results": risk_results,
        "settings": settings.dict(),
    }
    
    # Ensure output directory exists
    out_path = Path(out_dir)
    out_path.mkdir(parents=True, exist_ok=True)
    
    # Export results
    if fmt.lower() == "xlsx":
        exporters.export_excel(export_data, out_path / "ebiosrm_report.xlsx")
    elif fmt.lower() == "json":
        exporters.export_json(export_data, out_path / "ebiosrm_report.json")
    else:
        raise ValueError(f"Unsupported export format: {fmt}")
    
    logger.info(f"Report generated successfully in {out_path}")
