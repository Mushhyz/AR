"""Export functions for different output formats."""

import json
import logging
from pathlib import Path
from typing import Dict, Any

import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.styles import Font, PatternFill

logger = logging.getLogger(__name__)


def export_json(data: Dict[str, Any], output_path: Path) -> None:
    """Export data to JSON format."""
    logger.info(f"Exporting to JSON: {output_path}")
    
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def export_excel(data: Dict[str, Any], output_path: Path) -> None:
    """Export data to Excel format."""
    logger.info(f"Exporting to Excel: {output_path}")
    
    wb = Workbook()
    
    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
    
    # Create worksheets
    _create_assets_sheet(wb, data["assets"])
    _create_threats_sheet(wb, data["threats"])
    _create_risk_results_sheet(wb, data["risk_results"])
    
    # Create optional component sheets if data exists
    if data["risk_sources"]:
        _create_risk_sources_sheet(wb, data["risk_sources"])
    
    if data["objectives"]:
        _create_objectives_sheet(wb, data["objectives"])
    
    if data["stakeholders"]:
        _create_stakeholders_sheet(wb, data["stakeholders"])
    
    if data["measures"]:
        _create_measures_sheet(wb, data["measures"])
    
    wb.save(output_path)


def _create_assets_sheet(wb: Workbook, assets: list[Dict[str, Any]]) -> None:
    """Create assets worksheet."""
    ws = wb.create_sheet("Assets")
    
    # Headers
    headers = ["ID", "Type", "Label", "Criticality", "Severity Score"]
    ws.append(headers)
    
    # Style headers
    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
    
    # Data rows
    for asset in assets:
        ws.append([
            asset["id"],
            asset["type"],
            asset["label"],
            asset["criticality"],
            asset.get("severity_score", "")
        ])


def _create_threats_sheet(wb: Workbook, threats: list[Dict[str, Any]]) -> None:
    """Create threats worksheet."""
    ws = wb.create_sheet("Threats")
    
    headers = ["SR ID", "OV ID", "Strategic Path", "Operational Steps"]
    ws.append(headers)
    
    # Style headers
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    
    for threat in threats:
        ws.append([
            threat["sr_id"],
            threat["ov_id"],
            threat["strategic_path"],
            threat["operational_steps"]
        ])


def _create_risk_results_sheet(wb: Workbook, results: list[Dict[str, Any]]) -> None:
    """Create risk assessment results worksheet."""
    ws = wb.create_sheet("Risk Assessment")
    
    headers = ["Threat ID", "OV ID", "Strategic Path", "Likelihood", "Severity", "Risk Level"]
    ws.append(headers)
    
    # Style headers
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    
    # Risk level colors
    risk_colors = {
        "Low": "92D050",
        "Medium": "FFC000", 
        "High": "FF6600",
        "Critical": "C00000"
    }
    
    for result in results:
        row = [
            result["threat_id"],
            result["threat_ov"],
            result["strategic_path"],
            f"{result['likelihood_score']:.2f}",
            result["severity_score"],
            result["risk_level"]
        ]
        ws.append(row)
        
        # Color code risk level
        risk_level = result["risk_level"]
        if risk_level in risk_colors:
            last_row = ws.max_row
            risk_cell = ws.cell(row=last_row, column=6)
            risk_cell.fill = PatternFill(
                start_color=risk_colors[risk_level],
                end_color=risk_colors[risk_level],
                fill_type="solid"
            )


def _create_risk_sources_sheet(wb: Workbook, sources: list[Dict[str, Any]]) -> None:
    """Create risk sources worksheet."""
    ws = wb.create_sheet("Risk Sources")
    
    headers = ["ID", "Label", "Category", "Motivation", "Capability", "Resources"]
    ws.append(headers)
    
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    
    for source in sources:
        ws.append([
            source["id"],
            source["label"],
            source["category"],
            source["motivation"],
            source["capability_level"],
            source["resources"]
        ])


def _create_objectives_sheet(wb: Workbook, objectives: list[Dict[str, Any]]) -> None:
    """Create objectives worksheet."""
    ws = wb.create_sheet("Objectives")
    
    headers = ["ID", "Label", "Target Assets", "Business Impact"]
    ws.append(headers)
    
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    
    for obj in objectives:
        ws.append([
            obj["id"],
            obj["label"],
            ", ".join(obj["target_assets"]),
            obj["business_impact"]
        ])


def _create_stakeholders_sheet(wb: Workbook, stakeholders: list[Dict[str, Any]]) -> None:
    """Create stakeholders worksheet."""
    ws = wb.create_sheet("Stakeholders")
    
    headers = ["ID", "Name", "Type", "Role", "Responsibilities", "Contact"]
    ws.append(headers)
    
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    
    for stakeholder in stakeholders:
        ws.append([
            stakeholder["id"],
            stakeholder["name"],
            stakeholder["type"],
            stakeholder["role"],
            ", ".join(stakeholder["responsibilities"]),
            stakeholder.get("contact_info", "")
        ])


def _create_measures_sheet(wb: Workbook, measures: list[Dict[str, Any]]) -> None:
    """Create security measures worksheet."""
    ws = wb.create_sheet("Security Measures")
    
    headers = ["ID", "Label", "Type", "Description", "Effectiveness", "Cost", "Target Threats", "Responsible"]
    ws.append(headers)
    
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    
    for measure in measures:
        ws.append([
            measure["id"],
            measure["label"],
            measure["type"],
            measure["description"],
            measure["effectiveness"],
            measure["implementation_cost"],
            ", ".join(measure["target_threats"]),
            measure["responsible_stakeholder"]
        ])
