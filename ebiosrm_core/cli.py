"""Command-line interface for EBIOS RM generator."""

from __future__ import annotations

import logging
from pathlib import Path

import typer
from typing_extensions import Annotated

from . import generator, loader

# Console setup - Rich preferred, fallback to typer
try:
    from rich.console import Console

    console = Console()
except ImportError:
    # Fallback to typer.echo for console operations
    class SimpleConsole:
        def print(self, message, **kwargs):
            typer.echo(message)

    console = SimpleConsole()

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

app = typer.Typer(
    name="ebiosrm",
    help="EBIOS Risk Manager - Modular risk assessment generator",
    add_completion=False,
)


@app.command()
def export(
    cfg: Annotated[Path, typer.Option("--cfg", help="Configuration directory")] = Path(
        "config"
    ),
    out: Annotated[Path, typer.Option("--out", help="Output directory")] = Path(
        "build"
    ),
    fmt: Annotated[
        str, typer.Option("--fmt", help="Export format: xlsx|json|excel|markdown")
    ] = "xlsx",
    pme_profile: Annotated[
        bool, typer.Option("--pme-profile", "-p", help="Use simplified PME/TPE profile")
    ] = False,
    output_file: Annotated[
        str, typer.Option("--output", help="Override default filename")
    ] = "",
) -> None:
    """Generate risk assessment report from configuration data.

    Loads assets and threats from CSV files, calculates risk levels,
    and exports results in the specified format.

    Use --pme-profile flag for simplified PME/TPE configuration.
    """
    try:
        profile_text = "PME/TPE" if pme_profile else "Standard"
        typer.echo(f"🚀 Starting EBIOS RM risk assessment... (Profile: {profile_text})")
        typer.echo(f"   Config: {cfg}")
        typer.echo(f"   Output: {out}")
        typer.echo(f"   Format: {fmt}")

        # Load PME defaults if requested
        if pme_profile:
            pme_config = cfg / "pme_defaults.yaml"
            if pme_config.exists():
                import yaml

                with open(pme_config, "r", encoding="utf-8") as f:
                    yaml.safe_load(f)  # Load but don't store unused settings
                typer.echo(f"   Using PME profile from: {pme_config}")

        # Determine output filename - use default naming if not specified
        if not output_file:
            default_name = "ebios_risk_assessment"
            if fmt.lower() in {"xlsx", "excel"}:
                output_filename = f"{default_name}.xlsx"
            elif fmt.lower() == "json":
                output_filename = f"{default_name}.json"
            elif fmt.lower() in {"md", "markdown"}:
                output_filename = f"{default_name}.md"
            else:
                output_filename = f"{default_name}.{fmt.lower()}"
        else:
            # User provided custom filename
            if fmt.lower() in {"xlsx", "excel"}:
                output_filename = f"{output_file}.xlsx"
            elif fmt.lower() == "json":
                output_filename = f"{output_file}.json"
            elif fmt.lower() in {"md", "markdown"}:
                output_filename = f"{output_file}.md"
            else:
                output_filename = f"{output_file}.{fmt.lower()}"
        console.print(f"output_filename: {output_filename}")
        generator.run(
            cfg_dir=cfg,
            out_dir=out,
            fmt=fmt,
            pme_profile=pme_profile,
            output_filename=output_filename,
        )

        typer.echo("✅ Risk assessment completed successfully!")
        typer.echo(f"   Output file: {out / output_filename}")

    except FileNotFoundError as e:
        typer.echo(f"❌ Error: {e}", err=True)
        raise typer.Exit(1)
    except ValueError as e:
        typer.echo(f"❌ Configuration error: {e}", err=True)
        raise typer.Exit(1)
    except Exception as e:
        typer.echo(f"❌ Unexpected error: {e}", err=True)
        logging.exception("Unexpected error during export")
        raise typer.Exit(1)


@app.command()
def validate(
    cfg: Annotated[Path, typer.Option("--cfg", help="Configuration directory")] = Path(
        "config"
    ),
) -> None:
    """Validate configuration files without generating reports.

    Checks CSV and YAML files for correct format and data consistency.
    """
    try:
        typer.echo(f"🔍 Validating configuration in {cfg}...")

        assets, threats, settings, risk_sources, objectives, stakeholders, measures = (
            loader.load_all(cfg)
        )

        typer.echo("✅ Validation successful!")
        typer.echo(f"   Assets: {len(assets)} loaded")
        typer.echo(f"   Threats: {len(threats)} loaded")
        typer.echo(f"   Risk Sources: {len(risk_sources)} loaded")
        typer.echo(f"   Objectives: {len(objectives)} loaded")
        typer.echo(f"   Stakeholders: {len(stakeholders)} loaded")
        typer.echo(f"   Security Measures: {len(measures)} loaded")
        typer.echo(f"   Settings: {settings.output_dir}")

        # Additional validation checks
        asset_ids = {asset.id for asset in assets}
        if len(asset_ids) != len(assets):
            typer.echo("⚠️  Warning: Duplicate asset IDs detected", err=True)

        threat_ids = {threat.sr_id for threat in threats}
        if len(threat_ids) != len(threats):
            typer.echo("⚠️  Warning: Duplicate threat IDs detected", err=True)

        objective_ids = {obj.id for obj in objectives}
        if len(objective_ids) != len(objectives):
            typer.echo("⚠️  Warning: Duplicate objective IDs detected", err=True)

    except FileNotFoundError as e:
        typer.echo(f"❌ Error: {e}", err=True)
        raise typer.Exit(1)
    except ValueError as e:
        typer.echo(f"❌ Validation error: {e}", err=True)
        raise typer.Exit(1)
    except Exception as e:
        typer.echo(f"❌ Unexpected error: {e}", err=True)
        logging.exception("Unexpected error during validation")
        raise typer.Exit(1)


@app.command()
def version() -> None:
    """Display version information."""
    from . import __version__

    typer.echo(f"EBIOS RM Generator v{__version__}")


@app.command()
def template(
    output: Path = typer.Option(
        Path("templates/"),
        "--output",
        "-o",
        help="Output directory for template files",
    ),
    pme_profile: bool = typer.Option(
        False, "--pme", help="Generate PME-specific template"
    ),
):
    """Generate EBIOS RM Excel template."""
    try:
        # Fix import path to use scripts module
        from scripts.generate_template import EBIOSTemplateGenerator

        output.mkdir(parents=True, exist_ok=True)
        template_file = output / "ebiosrm_template.xlsx"

        typer.echo(f"🔧 Generating EBIOS RM template...")
        typer.echo(f"   Output: {template_file}")
        typer.echo(f"   Profile: {'PME' if pme_profile else 'Standard'}")

        generator = EBIOSTemplateGenerator()
        generator.generate_template(template_file, pme_profile=pme_profile)

        typer.echo("✅ Template generated successfully!")
        typer.echo(f"   File: {template_file}")

    except ImportError as e:
        typer.echo(f"❌ Template generator not available: {e}")
        typer.echo("   Make sure the 'scripts' module is available")
        raise typer.Exit(1)
    except Exception as e:
        typer.echo(f"❌ Template generation failed: {e}")
        logger.error(f"Template generation error: {e}")
        raise typer.Exit(1)


if __name__ == "__main__":
    app()
