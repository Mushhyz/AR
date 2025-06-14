"""Data models for EBIOS RM components."""

from __future__ import annotations

from enum import Enum
from pathlib import Path
from pydantic import BaseModel, Field, validator


class CriticalityLevel(str, Enum):
    """Asset criticality levels."""
    LOW = "Low"
    MEDIUM = "Medium"
    HIGH = "High"
    CRITICAL = "Critical"


class LikelihoodLevel(str, Enum):
    """Threat likelihood levels."""
    ONE_SHOT = "One-shot"
    OCCASIONAL = "Occasional"
    PROBABLE = "Probable"
    SYSTEMATIC = "Systematic"


class StakeholderType(str, Enum):
    """Types of stakeholders."""
    INTERNAL = "Internal"
    EXTERNAL = "External"
    REGULATORY = "Regulatory"
    COMMERCIAL = "Commercial"


class MeasureType(str, Enum):
    """Types of security measures."""
    PREVENTIVE = "Preventive"
    DETECTIVE = "Detective"
    CORRECTIVE = "Corrective"
    RECOVERY = "Recovery"


class Asset(BaseModel):
    """Business asset with criticality assessment."""
    
    id: str = Field(..., description="Unique asset identifier")
    type: str = Field(..., description="Asset category")
    label: str = Field(..., description="Human-readable name")
    criticality: CriticalityLevel = Field(..., description="Business criticality")
    
    def severity_score(self) -> int:
        """Convert criticality to numeric score (1-4)."""
        mapping = {
            CriticalityLevel.LOW: 1,
            CriticalityLevel.MEDIUM: 2,
            CriticalityLevel.HIGH: 3,
            CriticalityLevel.CRITICAL: 4,
        }
        return mapping[self.criticality]


class RiskSource(BaseModel):
    """Source de risque - entities that can generate threats."""
    
    id: str = Field(..., description="Risk source identifier")
    label: str = Field(..., description="Risk source name")
    category: str = Field(..., description="Source category (cybercriminal, insider, state, etc.)")
    motivation: str = Field(..., description="Primary motivation")
    capability_level: CriticalityLevel = Field(..., description="Technical capability level")
    resources: str = Field(..., description="Available resources description")


class TargetedObjective(BaseModel):
    """Objectif visé - what attackers aim to achieve."""
    
    id: str = Field(..., description="Objective identifier")
    label: str = Field(..., description="Objective description")
    target_assets: list[str] = Field(..., description="List of targeted asset IDs")
    business_impact: CriticalityLevel = Field(..., description="Potential business impact")
    attack_scenarios: list[str] = Field(default_factory=list, description="Related attack scenarios")


class Stakeholder(BaseModel):
    """Partie prenante - stakeholders in the risk management process."""
    
    id: str = Field(..., description="Stakeholder identifier")
    name: str = Field(..., description="Stakeholder name")
    type: StakeholderType = Field(..., description="Stakeholder type")
    role: str = Field(..., description="Role in the organization")
    responsibilities: list[str] = Field(..., description="Security responsibilities")
    contact_info: str = Field(default="", description="Contact information")


class SecurityMeasure(BaseModel):
    """Mesure de sécurité - security controls and countermeasures."""
    
    id: str = Field(..., description="Measure identifier")
    label: str = Field(..., description="Measure name")
    type: MeasureType = Field(..., description="Type of security measure")
    description: str = Field(..., description="Detailed description")
    effectiveness: CriticalityLevel = Field(..., description="Effectiveness level")
    implementation_cost: CriticalityLevel = Field(..., description="Implementation cost")
    target_threats: list[str] = Field(default_factory=list, description="Threat IDs this measure addresses")
    responsible_stakeholder: str = Field(..., description="Responsible stakeholder ID")


class Threat(BaseModel):
    """Security threat with operational scenario."""
    
    sr_id: str = Field(..., description="Strategic risk identifier")
    ov_id: str = Field(..., description="Operational view identifier")
    strategic_path: str = Field(..., description="High-level attack path")
    operational_steps: str = Field(..., description="Detailed steps with likelihood")
    risk_sources: list[str] = Field(default_factory=list, description="Associated risk source IDs")
    targeted_objectives: list[str] = Field(default_factory=list, description="Associated objective IDs")
    
    @validator("operational_steps")
    def validate_steps_format(cls, v: str) -> str:
        """Ensure steps follow 'StepX:Level' format."""
        if not v or ":" not in v:
            raise ValueError("operational_steps must contain 'Step:Level' pairs")
        return v
    
    def likelihood_score(self) -> float:
        """Calculate weighted likelihood from operational steps."""
        step_mapping = {
            "Low": 1, "Medium": 2, "High": 3, "Critical": 4
        }
        
        steps = [s.strip() for s in self.operational_steps.split(",")]
        scores = []
        
        for step in steps:
            if ":" not in step:
                continue
            _, level = step.split(":", 1)
            level = level.strip()
            if level in step_mapping:
                scores.append(step_mapping[level])
        
        return sum(scores) / len(scores) if scores else 1.0
    
    def risk_level(self, max_asset_severity: int) -> str:
        """Calculate final risk level using 4x4 matrix."""
        likelihood = self.likelihood_score()
        severity = max_asset_severity
        
        # Risk matrix: severity (rows) x likelihood (cols)
        matrix = [
            ["Low", "Low", "Medium", "High"],      # Low severity
            ["Low", "Medium", "Medium", "High"],   # Medium severity  
            ["Medium", "Medium", "High", "Critical"], # High severity
            ["Medium", "High", "Critical", "Critical"], # Critical severity
        ]
        
        sev_idx = min(int(severity) - 1, 3)
        lik_idx = min(int(likelihood) - 1, 3)
        
        return matrix[sev_idx][lik_idx]


class Settings(BaseModel):
    """Global configuration settings."""
    
    excel_template: str = Field(default="templates/ebiosrm_empty.xlsx")
    severity_scale: list[str] = Field(default=["Low", "Medium", "High", "Critical"])
    likelihood_scale: list[str] = Field(default=["One-shot", "Occasional", "Probable", "Systematic"])
    output_dir: str = Field(default="build/")
    
    @validator("output_dir")
    def ensure_trailing_slash(cls, v: str) -> str:
        """Ensure output directory ends with slash."""
        return v if v.endswith("/") else v + "/"
