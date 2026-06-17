from typer.testing import CliRunner

from recipe_automation.main import app

runner = CliRunner()


def test_app_scan_help() -> None:
    result = runner.invoke(app, ["scan", "--help"])
    assert result.exit_code == 0
    assert "Taranıyor" not in result.stdout
