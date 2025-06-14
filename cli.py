"""Command-line interface for EBIOS RM generator."""

from __future__ import annotations

import logging
from pathlib import Path

import typer
from typing_extensions import Annotated

from ebiosrm_core import generator, loader

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)

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
    fmt: Annotated[str, typer.Option("--fmt", help="Export format")] = "xlsx",
) -> None:
    """Generate risk assessment report from configuration data.

    Loads assets and threats from CSV files, calculates risk levels,
    and exports results in the specified format.
    """
    try:
        typer.echo("🚀 Starting EBIOS RM risk assessment...")
        typer.echo(f"   Config: {cfg}")
        typer.echo(f"   Output: {out}")
        typer.echo(f"   Format: {fmt}")

        generator.run(cfg_dir=cfg, out_dir=out, fmt=fmt)

        typer.echo("✅ Risk assessment completed successfully!")

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

        assets, threats, settings = loader.load_all(cfg)

        typer.echo("✅ Validation successful!")
        typer.echo(f"   Assets: {len(assets)} loaded")
        typer.echo(f"   Threats: {len(threats)} loaded")
        typer.echo(f"   Settings: {settings.output_dir}")

        # Additional validation checks
        asset_ids = {asset.id for asset in assets}
        if len(asset_ids) != len(assets):
            typer.echo("⚠️  Warning: Duplicate asset IDs detected", err=True)

        threat_ids = {threat.sr_id for threat in threats}
        if len(threat_ids) != len(threats):
            typer.echo("⚠️  Warning: Duplicate threat IDs detected", err=True)

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
    from ebiosrm_core import __version__

    typer.echo(f"EBIOS RM Generator v{__version__}")


if __name__ == "__main__":
    app()
