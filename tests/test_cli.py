import json
import subprocess
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CLI = ROOT / "pptxcli"


def run_cli(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [str(CLI), *args],
        check=False,
        capture_output=True,
        text=True,
    )


class CliSmokeTest(unittest.TestCase):
    def test_help_runs(self) -> None:
        result = run_cli("--help")
        self.assertEqual(result.returncode, 0)
        self.assertIn("pptxcli", result.stdout)

    def test_demo_form_outputs_json(self) -> None:
        result = run_cli("demo", "form")
        self.assertEqual(result.returncode, 0)
        payload = json.loads(result.stdout)
        self.assertEqual(payload["template_path"], "template.pptx")


if __name__ == "__main__":
    unittest.main()
