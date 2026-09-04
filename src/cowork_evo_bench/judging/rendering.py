"""把 PPTX/PDF 稳定渲染为视觉 Judge 可审计的逐页图片。"""

from __future__ import annotations

import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import time
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageOps

from cowork_evo_bench.judging.visual_config import RenderingSettings
from cowork_evo_bench.judging.visual_models import VisualArtifactSpec


class RenderFailure(RuntimeError):
    """渲染失败；task_gate 表示候选核心文件本身无效，应记任务零分。"""

    def __init__(self, code: str, message: str, *, task_gate: bool = False) -> None:
        super().__init__(message)
        self.code = code
        self.task_gate = task_gate


@dataclass(frozen=True, slots=True)
class RenderedPage:
    page_number: int
    image_path: Path
    sha256: str
    width: int
    height: int


@dataclass(frozen=True, slots=True)
class RenderedArtifact:
    role: str
    label: str
    source_relative_path: str
    source_sha256: str
    kind: str
    page_count: int
    pages: tuple[RenderedPage, ...]
    contact_sheet_path: Path
    contact_sheet_sha256: str
    runtime_ms: int


@dataclass(frozen=True, slots=True)
class RenderBundle:
    role: str
    artifacts: tuple[RenderedArtifact, ...]
    manifest_path: Path

    @property
    def attachment_paths(self) -> tuple[Path, ...]:
        """先给全局缩略图，再给逐页原图，兼顾全局一致性与局部细节。"""

        contacts = tuple(artifact.contact_sheet_path for artifact in self.artifacts)
        pages = tuple(page.image_path for artifact in self.artifacts for page in artifact.pages)
        return contacts + pages

    @property
    def image_sha256(self) -> tuple[str, ...]:
        values = [artifact.contact_sheet_sha256 for artifact in self.artifacts]
        values.extend(page.sha256 for artifact in self.artifacts for page in artifact.pages)
        return tuple(dict.fromkeys(values))


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _resolve_safe_file(root: Path, relative_path: str, *, task_gate: bool) -> Path:
    root = root.resolve()
    candidate = root / relative_path
    if candidate.is_symlink():
        raise RenderFailure("unsafe_symlink", "评分输入不能是符号链接。", task_gate=task_gate)
    resolved = candidate.resolve()
    try:
        resolved.relative_to(root)
    except ValueError as error:
        raise RenderFailure("unsafe_path", "评分输入路径越出允许根目录。", task_gate=task_gate) from error
    if not resolved.is_file():
        raise RenderFailure(
            "candidate_artifact_missing",
            f"候选核心产物不存在：{relative_path}",
            task_gate=task_gate,
        )
    return resolved


def _validate_source(path: Path, kind: str, *, candidate: bool) -> None:
    gate = candidate
    if kind == "pptx":
        try:
            with zipfile.ZipFile(path) as archive:
                names = set(archive.namelist())
                if "[Content_Types].xml" not in names or "ppt/presentation.xml" not in names:
                    raise RenderFailure(
                        "invalid_pptx",
                        "PPTX 缺少必要 OOXML 结构。",
                        task_gate=gate,
                    )
                bad_member = archive.testzip()
                if bad_member is not None:
                    raise RenderFailure(
                        "damaged_pptx",
                        f"PPTX ZIP 成员损坏：{bad_member}",
                        task_gate=gate,
                    )
        except (OSError, zipfile.BadZipFile) as error:
            raise RenderFailure("damaged_pptx", "PPTX 不是有效 OOXML 文件。", task_gate=gate) from error
    elif kind == "pdf":
        try:
            header = path.read_bytes()[:5]
        except OSError as error:
            raise RenderFailure("pdf_read_failure", "无法读取 PDF。", task_gate=gate) from error
        if header != b"%PDF-":
            raise RenderFailure("damaged_pdf", "PDF 文件头无效。", task_gate=gate)


def validate_candidate_artifacts(
    specs: Iterable[VisualArtifactSpec],
    candidate_root: Path,
) -> None:
    """在启动渲染前识别“核心产物缺失/损坏”这一唯一任务零分门槛。"""

    for spec in specs:
        source = _resolve_safe_file(candidate_root, spec.path, task_gate=True)
        _validate_source(source, spec.kind, candidate=True)


def _run(command: list[str], timeout_seconds: int, failure_code: str) -> subprocess.CompletedProcess[str]:
    try:
        result = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
        )
    except FileNotFoundError as error:
        raise RenderFailure(f"{failure_code}_binary_missing", "缺少渲染所需系统程序。") from error
    except subprocess.TimeoutExpired as error:
        raise RenderFailure(f"{failure_code}_timeout", "渲染命令超时。") from error
    if result.returncode != 0:
        stderr = (result.stderr or result.stdout or "").strip()[-1000:]
        raise RenderFailure(failure_code, f"渲染命令失败：{stderr}")
    return result


def _convert_pptx_to_pdf(source: Path, directory: Path, timeout_seconds: int) -> Path:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RenderFailure("office_renderer_missing", "系统没有 LibreOffice/soffice。")
    output_dir = directory / "pdf"
    profile_dir = directory / "libreoffice-profile"
    output_dir.mkdir(parents=True)
    profile_dir.mkdir(parents=True)
    profile_uri = profile_dir.resolve().as_uri()
    _run(
        [
            soffice,
            "--headless",
            f"-env:UserInstallation={profile_uri}",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(source),
        ],
        timeout_seconds,
        "pptx_conversion_failure",
    )
    expected = output_dir / f"{source.stem}.pdf"
    if not expected.is_file() or expected.stat().st_size == 0:
        raise RenderFailure("pptx_conversion_failure", "LibreOffice 未生成 PDF。")
    return expected


def _pdf_page_count(pdf: Path, timeout_seconds: int) -> int:
    pdfinfo = shutil.which("pdfinfo")
    if not pdfinfo:
        raise RenderFailure("pdfinfo_binary_missing", "系统没有 pdfinfo。")
    output = _run([pdfinfo, str(pdf)], timeout_seconds, "pdfinfo_failure").stdout
    match = re.search(r"^Pages:\s+([0-9]+)\s*$", output, re.MULTILINE)
    if not match:
        raise RenderFailure("pdfinfo_failure", "无法从 pdfinfo 读取页数。")
    return int(match.group(1))


def _render_pdf_pages(
    pdf: Path,
    output_dir: Path,
    settings: RenderingSettings,
) -> tuple[RenderedPage, ...]:
    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm:
        raise RenderFailure("pdftoppm_binary_missing", "系统没有 pdftoppm。")
    page_count = _pdf_page_count(pdf, settings.command_timeout_seconds)
    if page_count < 1 or page_count > settings.max_pages:
        raise RenderFailure(
            "page_count_out_of_range",
            f"页数 {page_count} 超出允许范围 1..{settings.max_pages}。",
        )
    pages_dir = output_dir / "pages"
    pages_dir.mkdir(parents=True)
    prefix = pages_dir / "page"
    _run(
        [
            pdftoppm,
            "-png",
            "-r",
            str(settings.dpi),
            "-f",
            "1",
            "-l",
            str(page_count),
            str(pdf),
            str(prefix),
        ],
        settings.command_timeout_seconds,
        "pdf_render_failure",
    )
    image_paths = sorted(
        pages_dir.glob("page-*.png"),
        key=lambda path: int(path.stem.rsplit("-", 1)[1]),
    )
    if len(image_paths) != page_count:
        raise RenderFailure(
            "pdf_render_incomplete",
            f"期望 {page_count} 页，实际生成 {len(image_paths)} 张图片。",
        )
    pages: list[RenderedPage] = []
    for index, image_path in enumerate(image_paths, start=1):
        with Image.open(image_path) as image:
            width, height = image.size
            image.verify()
        pages.append(
            RenderedPage(
                page_number=index,
                image_path=image_path.resolve(),
                sha256=sha256_file(image_path),
                width=width,
                height=height,
            )
        )
    return tuple(pages)


def _make_contact_sheet(
    pages: tuple[RenderedPage, ...],
    output_path: Path,
    columns: int,
) -> None:
    thumbnail_width = 480
    margin = 16
    rendered: list[Image.Image] = []
    try:
        for page in pages:
            with Image.open(page.image_path) as image:
                image = image.convert("RGB")
                height = max(1, round(image.height * thumbnail_width / image.width))
                rendered.append(ImageOps.contain(image, (thumbnail_width, height)).copy())
        row_heights = [
            max(image.height for image in rendered[index : index + columns])
            for index in range(0, len(rendered), columns)
        ]
        sheet_width = columns * thumbnail_width + (columns + 1) * margin
        sheet_height = sum(row_heights) + (len(row_heights) + 1) * margin
        sheet = Image.new("RGB", (sheet_width, sheet_height), "white")
        y = margin
        for row_index, row_height in enumerate(row_heights):
            row = rendered[row_index * columns : (row_index + 1) * columns]
            for column_index, image in enumerate(row):
                x = margin + column_index * (thumbnail_width + margin)
                sheet.paste(image, (x, y))
            y += row_height + margin
        sheet.save(output_path, format="PNG", optimize=True)
    finally:
        for image in rendered:
            image.close()


def render_artifacts(
    specs: tuple[VisualArtifactSpec, ...],
    source_root: Path,
    output_root: Path,
    settings: RenderingSettings,
    *,
    role: str,
) -> RenderBundle:
    """渲染一组同角色产物并写 manifest；不复制原始 Office 文件。"""

    if role not in {"candidate", "reference"}:
        raise ValueError("role 必须是 candidate 或 reference。")
    output_root.mkdir(parents=True, exist_ok=False)
    artifacts: list[RenderedArtifact] = []
    for index, spec in enumerate(specs, start=1):
        started = time.monotonic()
        source = _resolve_safe_file(
            source_root,
            spec.path,
            task_gate=role == "candidate",
        )
        _validate_source(source, spec.kind, candidate=role == "candidate")
        artifact_dir = output_root / f"artifact-{index:02d}"
        artifact_dir.mkdir()
        with tempfile.TemporaryDirectory(prefix="cowork-evo-render-") as temp:
            temp_dir = Path(temp)
            pdf = (
                _convert_pptx_to_pdf(source, temp_dir, settings.command_timeout_seconds)
                if spec.kind == "pptx"
                else source
            )
            pages = _render_pdf_pages(pdf, artifact_dir, settings)
        contact_sheet = artifact_dir / "contact-sheet.png"
        _make_contact_sheet(pages, contact_sheet, settings.contact_sheet_columns)
        artifacts.append(
            RenderedArtifact(
                role=role,
                label=spec.label or Path(spec.path).name,
                source_relative_path=spec.path,
                source_sha256=sha256_file(source),
                kind=spec.kind,
                page_count=len(pages),
                pages=pages,
                contact_sheet_path=contact_sheet.resolve(),
                contact_sheet_sha256=sha256_file(contact_sheet),
                runtime_ms=round((time.monotonic() - started) * 1000),
            )
        )
    manifest_path = output_root / "manifest.json"
    payload = {
        "schema_version": "1.0",
        "role": role,
        "renderer": {
            "office": "LibreOffice headless",
            "pdf": "pdftoppm",
            "dpi": settings.dpi,
        },
        "artifacts": [
            {
                "role": artifact.role,
                "label": artifact.label,
                "source_relative_path": artifact.source_relative_path,
                "source_sha256": artifact.source_sha256,
                "kind": artifact.kind,
                "page_count": artifact.page_count,
                "pages": [
                    {
                        **asdict(page),
                        "image_path": str(page.image_path.relative_to(output_root.resolve())),
                    }
                    for page in artifact.pages
                ],
                "contact_sheet_path": str(
                    artifact.contact_sheet_path.relative_to(output_root.resolve())
                ),
                "contact_sheet_sha256": artifact.contact_sheet_sha256,
                "runtime_ms": artifact.runtime_ms,
            }
            for artifact in artifacts
        ],
    }
    manifest_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return RenderBundle(role=role, artifacts=tuple(artifacts), manifest_path=manifest_path)
