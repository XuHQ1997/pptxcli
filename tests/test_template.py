import json
import os
import subprocess
import tempfile
import unittest
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Inches

from pptx_cli.inspect import inspect_slide

ROOT = Path(__file__).resolve().parents[1]
CLI = ROOT / "pptxcli"


def run_cli(*args: str, env: dict[str, str] | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [str(CLI), *args],
        check=False,
        capture_output=True,
        text=True,
        env=env,
    )


def collect_slide_texts(presentation: Presentation) -> list[str]:
    slide_texts: list[str] = []
    for slide in presentation.slides:
        parts: list[str] = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                text = " ".join(shape.text.split())
                if text:
                    parts.append(text)
        slide_texts.append(" | ".join(parts))
    return slide_texts


class TemplateWorkflowTest(unittest.TestCase):
    def test_template_create_add_slide_save_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir_str:
            tmp_dir = Path(tmp_dir_str)
            state_file = tmp_dir / ".pptxcli-session.json"
            template_root = tmp_dir / "template-store"
            pptx_path = tmp_dir / "demo.pptx"
            image_path = tmp_dir / "hero.png"
            Image.new("RGB", (80, 80), "green").save(image_path)

            presentation = Presentation()
            slide_0 = presentation.slides.add_slide(presentation.slide_layouts[6])
            slide_0.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1)).text_frame.text = (
                "Hero Title"
            )

            slide_1 = presentation.slides.add_slide(presentation.slide_layouts[6])
            slide_1.shapes.add_picture(
                str(image_path),
                Inches(1),
                Inches(1.5),
                Inches(2),
                Inches(2),
            )

            slide_2 = presentation.slides.add_slide(presentation.slide_layouts[6])
            slide_2.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1)).text_frame.text = (
                "Summary Page"
            )
            presentation.save(pptx_path)

            slide_0_candidates = inspect_slide(pptx_path, 0).candidates
            slide_2_candidates = inspect_slide(pptx_path, 2).candidates

            env = os.environ.copy()
            env["PPTXCLI_STATE_FILE"] = str(state_file)
            env["PPTXCLI_TEMPLATE_ROOT"] = str(template_root)

            init_result = run_cli("init", "--origin_file", str(pptx_path), env=env)
            self.assertEqual(init_result.returncode, 0, init_result.stderr)

            create_result = run_cli("template", "create", "--name", "demo_template.pptx", env=env)
            self.assertEqual(create_result.returncode, 0, create_result.stderr)
            create_payload = json.loads(create_result.stdout)
            self.assertEqual(create_payload["template_name"], "demo_template")
            draft_path = Path(create_payload["draft_path"])
            self.assertTrue(draft_path.exists())

            add_slide_2_result = run_cli(
                "template",
                "add_slide",
                "--slide",
                "2",
                "-f",
                f"{slide_2_candidates[0].index}:summary title",
                env=env,
            )
            self.assertEqual(add_slide_2_result.returncode, 0, add_slide_2_result.stderr)

            add_slide_0_result = run_cli(
                "template",
                "add_slide",
                "--slide",
                "0",
                "-f",
                f"{slide_0_candidates[0].index}:hero title",
                env=env,
            )
            self.assertEqual(add_slide_0_result.returncode, 0, add_slide_0_result.stderr)

            save_result = run_cli("template", "save", env=env)
            self.assertEqual(save_result.returncode, 0, save_result.stderr)
            save_payload = json.loads(save_result.stdout)
            manifest_path = Path(save_payload["manifest_path"])
            template_pptx_path = Path(save_payload["template_pptx_path"])

            self.assertTrue(manifest_path.exists())
            self.assertTrue(template_pptx_path.exists())

            draft_payload = json.loads(draft_path.read_text(encoding="utf-8"))
            self.assertEqual(
                [slide["source_slide_index"] for slide in draft_payload["slides"]],
                [2, 0],
            )

            manifest_payload = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(manifest_payload["template_name"], "demo_template")
            self.assertEqual(manifest_payload["slide_count"], 2)
            self.assertEqual(manifest_payload["field_count"], 2)
            self.assertEqual(
                [slide["source_slide_index"] for slide in manifest_payload["slides"]],
                [2, 0],
            )
            self.assertEqual(
                [slide["slide_name"] for slide in manifest_payload["slides"]],
                ["slide_2", "slide_0"],
            )
            self.assertEqual(
                [field["description"] for field in manifest_payload["fields"]],
                ["summary title", "hero title"],
            )
            self.assertEqual(
                [field["index"] for field in manifest_payload["fields"]],
                [slide_2_candidates[0].index, slide_0_candidates[0].index],
            )

            template_presentation = Presentation(str(template_pptx_path))
            self.assertEqual(len(template_presentation.slides), 2)
            self.assertEqual(
                collect_slide_texts(template_presentation),
                ["Summary Page", "Hero Title"],
            )

            finish_result = run_cli("finish", env=env)
            self.assertEqual(finish_result.returncode, 0, finish_result.stderr)


if __name__ == "__main__":
    unittest.main()
