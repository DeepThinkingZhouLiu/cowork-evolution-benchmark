"""使用本地小型 PPTX/PDF 验证渲染、manifest 与 reference 隔离。"""

import json
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from cowork_evo_bench.judging.rendering import render_artifacts
from cowork_evo_bench.judging.visual_config import RenderingSettings
from cowork_evo_bench.judging.visual_models import VisualArtifactSpec


def _make_pptx(path: Path) -> None:
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    textbox.text = "本地视觉评分测试"
    deck.save(path)


class VisualRenderingIntegrationTest(unittest.TestCase):
    def setUp(self) -> None:
        if not shutil.which("soffice") or not shutil.which("pdftoppm"):
            self.skipTest("缺少 LibreOffice 或 pdftoppm")
        self.settings = RenderingSettings(
            dpi=96,
            contact_sheet_columns=2,
            max_pages=10,
            command_timeout_seconds=60,
        )

    def test_renders_pptx_and_pdf_with_auditable_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pptx = root / "sample.pptx"
            _make_pptx(pptx)
            ppt_bundle = render_artifacts(
                (VisualArtifactSpec("sample.pptx", "pptx"),),
                root,
                root / "render-pptx",
                self.settings,
                role="candidate",
            )
            self.assertEqual(len(ppt_bundle.artifacts[0].pages), 1)
            self.assertTrue(ppt_bundle.artifacts[0].contact_sheet_path.is_file())
            manifest = json.loads(ppt_bundle.manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(manifest["artifacts"][0]["source_relative_path"], "sample.pptx")
            self.assertEqual(len(manifest["artifacts"][0]["source_sha256"]), 64)

            profile = root / "lo-profile"
            profile.mkdir()
            result = subprocess.run(
                [
                    shutil.which("soffice"),
                    "--headless",
                    f"-env:UserInstallation={profile.resolve().as_uri()}",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(root),
                    str(pptx),
                ],
                check=False,
                capture_output=True,
                text=True,
                timeout=60,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            pdf_bundle = render_artifacts(
                (VisualArtifactSpec("sample.pdf", "pdf"),),
                root,
                root / "render-pdf",
                self.settings,
                role="candidate",
            )
            self.assertEqual(len(pdf_bundle.artifacts[0].pages), 1)


if __name__ == "__main__":
    unittest.main()
