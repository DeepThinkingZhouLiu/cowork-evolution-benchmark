"""Codex-VLM 评分配置；本模块只保存变量名，不持有或序列化秘密。"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Mapping
from urllib.parse import urlparse


class VisualConfigError(ValueError):
    """公共视觉评分配置无效。"""


@dataclass(frozen=True, slots=True)
class ProviderSettings:
    name: str
    base_url: str
    api_key_env: str
    wire_api: str


@dataclass(frozen=True, slots=True)
class CodexSettings:
    binary: str
    expected_version: str
    model: str
    reasoning_effort: str
    timeout_seconds: int
    max_retries: int
    concurrency: int


@dataclass(frozen=True, slots=True)
class StabilitySettings:
    temperature_strategy: str
    repetitions: int
    aggregation: str


@dataclass(frozen=True, slots=True)
class RenderingSettings:
    dpi: int
    contact_sheet_columns: int
    max_pages: int
    command_timeout_seconds: int


@dataclass(frozen=True, slots=True)
class RewardWeights:
    deterministic: float
    visual: float


@dataclass(frozen=True, slots=True)
class VisualJudgeSettings:
    config_version: str
    provider: ProviderSettings
    codex: CodexSettings
    stability: StabilitySettings
    rendering: RenderingSettings
    weights: RewardWeights


def _require_mapping(value: object, name: str) -> dict[str, object]:
    if not isinstance(value, dict):
        raise VisualConfigError(f"{name} 必须是 JSON 对象。")
    return value


def _require_string(value: object, name: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise VisualConfigError(f"{name} 必须是非空字符串。")
    return value.strip()


def _require_int(value: object, name: str, minimum: int) -> int:
    if not isinstance(value, int) or isinstance(value, bool) or value < minimum:
        raise VisualConfigError(f"{name} 必须是不小于 {minimum} 的整数。")
    return value


def _require_number(value: object, name: str) -> float:
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        raise VisualConfigError(f"{name} 必须是数值。")
    return float(value)


def load_visual_judge_settings(
    path: Path,
    environ: Mapping[str, str] | None = None,
) -> VisualJudgeSettings:
    """读取非秘密配置，并允许用环境变量覆盖模型和本地二进制。"""

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise VisualConfigError(f"无法读取评分配置 {path}: {error}") from error
    root = _require_mapping(payload, "配置根节点")
    provider = _require_mapping(root.get("provider"), "provider")
    codex = _require_mapping(root.get("codex"), "codex")
    stability = _require_mapping(root.get("stability"), "stability")
    rendering = _require_mapping(root.get("rendering"), "rendering")
    aggregation = _require_mapping(root.get("aggregation"), "aggregation")
    env = os.environ if environ is None else environ

    provider_name = _require_string(provider.get("name"), "provider.name")
    if provider_name != "zcloud":
        raise VisualConfigError("当前可信评分 Provider 固定为 zcloud。")
    base_url = _require_string(
        env.get("COWORK_EVO_VLM_BASE_URL", provider.get("base_url")),
        "provider.base_url",
    ).rstrip("/")
    parsed_url = urlparse(base_url)
    if parsed_url.scheme != "https" or not parsed_url.netloc:
        raise VisualConfigError("provider.base_url 必须是有效 HTTPS 地址。")
    api_key_env = _require_string(provider.get("api_key_env"), "provider.api_key_env")
    if not re.fullmatch(r"[A-Z][A-Z0-9_]*", api_key_env):
        raise VisualConfigError("provider.api_key_env 必须是合法的大写环境变量名。")
    wire_api = _require_string(provider.get("wire_api"), "provider.wire_api")
    if wire_api != "responses":
        raise VisualConfigError("Codex-VLM 评分当前只支持 responses wire API。")

    expected_version = _require_string(
        codex.get("expected_version"), "codex.expected_version"
    )
    if not re.fullmatch(r"[0-9]+\.[0-9]+\.[0-9]+", expected_version):
        raise VisualConfigError("codex.expected_version 必须是 x.y.z。")
    reasoning_effort = _require_string(
        codex.get("reasoning_effort"), "codex.reasoning_effort"
    )
    if reasoning_effort not in {"low", "medium", "high", "xhigh"}:
        raise VisualConfigError("codex.reasoning_effort 不受支持。")

    deterministic_weight = _require_number(
        aggregation.get("deterministic_weight"),
        "aggregation.deterministic_weight",
    )
    visual_weight = _require_number(
        aggregation.get("visual_weight"), "aggregation.visual_weight"
    )
    if deterministic_weight < 0 or visual_weight < 0:
        raise VisualConfigError("聚合权重不能为负数。")
    if abs(deterministic_weight + visual_weight - 1.0) > 1e-9:
        raise VisualConfigError("确定性与视觉权重之和必须为 1。")

    temperature_strategy = _require_string(
        stability.get("temperature_strategy"),
        "stability.temperature_strategy",
    )
    if temperature_strategy != "provider_default":
        raise VisualConfigError(
            "Codex CLI 0.148.0 不暴露温度参数，必须使用 provider_default。"
        )
    stability_aggregation = _require_string(
        stability.get("aggregation"), "stability.aggregation"
    )
    if stability_aggregation != "median":
        raise VisualConfigError("多次视觉评分当前只支持 median 聚合。")

    return VisualJudgeSettings(
        config_version=_require_string(root.get("config_version"), "config_version"),
        provider=ProviderSettings(
            name=provider_name,
            base_url=base_url,
            api_key_env=api_key_env,
            wire_api=wire_api,
        ),
        codex=CodexSettings(
            binary=_require_string(
                env.get("COWORK_EVO_CODEX_BINARY", codex.get("binary")),
                "codex.binary",
            ),
            expected_version=expected_version,
            model=_require_string(
                env.get("COWORK_EVO_VLM_MODEL", codex.get("model")),
                "codex.model",
            ),
            reasoning_effort=reasoning_effort,
            timeout_seconds=_require_int(
                codex.get("timeout_seconds"), "codex.timeout_seconds", 1
            ),
            max_retries=_require_int(codex.get("max_retries"), "codex.max_retries", 0),
            concurrency=_require_int(codex.get("concurrency"), "codex.concurrency", 1),
        ),
        stability=StabilitySettings(
            temperature_strategy=temperature_strategy,
            repetitions=_require_int(
                stability.get("repetitions"), "stability.repetitions", 1
            ),
            aggregation=stability_aggregation,
        ),
        rendering=RenderingSettings(
            dpi=_require_int(rendering.get("dpi"), "rendering.dpi", 72),
            contact_sheet_columns=_require_int(
                rendering.get("contact_sheet_columns"),
                "rendering.contact_sheet_columns",
                1,
            ),
            max_pages=_require_int(rendering.get("max_pages"), "rendering.max_pages", 1),
            command_timeout_seconds=_require_int(
                rendering.get("command_timeout_seconds"),
                "rendering.command_timeout_seconds",
                1,
            ),
        ),
        weights=RewardWeights(
            deterministic=deterministic_weight,
            visual=visual_weight,
        ),
    )


def load_env_file(path: Path) -> dict[str, str]:
    """读取简单 KEY=VALUE 文件；拒绝 shell 展开，避免执行秘密文件内容。"""

    values: dict[str, str] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError as error:
        raise VisualConfigError(f"无法读取本地凭据文件 {path}: {error}") from error
    for line_number, raw_line in enumerate(lines, start=1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[7:].lstrip()
        if "=" not in line:
            raise VisualConfigError(f"{path}:{line_number} 缺少等号。")
        key, value = line.split("=", 1)
        key = key.strip()
        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", key):
            raise VisualConfigError(f"{path}:{line_number} 变量名不合法。")
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
            value = value[1:-1]
        values[key] = value
    return values


def resolve_api_key(
    settings: VisualJudgeSettings,
    dotenv_values: Mapping[str, str],
    environ: Mapping[str, str] | None = None,
) -> str:
    """只在调用边界解析秘密；错误消息绝不包含秘密值。"""

    env = os.environ if environ is None else environ
    api_key = env.get(settings.provider.api_key_env) or dotenv_values.get(
        settings.provider.api_key_env
    )
    if not api_key:
        raise VisualConfigError(
            f"缺少评分凭据环境变量 {settings.provider.api_key_env}。"
        )
    if any(char.isspace() for char in api_key):
        raise VisualConfigError("评分凭据不能包含空白字符。")
    return api_key
