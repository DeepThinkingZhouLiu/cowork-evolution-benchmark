"""Codex-VLM 公共配置与秘密解析测试。"""

import json
import tempfile
import unittest
from pathlib import Path

from cowork_evo_bench.judging.visual_config import (
    VisualConfigError,
    load_env_file,
    load_visual_judge_settings,
    resolve_api_key,
)


def _payload() -> dict:
    return {
        "config_version": "1.0",
        "provider": {
            "name": "zcloud",
            "base_url": "https://example.invalid/v1",
            "api_key_env": "COWORK_EVO_VLM_API_KEY",
            "wire_api": "responses",
        },
        "codex": {
            "binary": "codex",
            "expected_version": "0.148.0",
            "model": "gpt-5.6-terra",
            "reasoning_effort": "medium",
            "timeout_seconds": 30,
            "max_retries": 2,
            "concurrency": 1,
        },
        "stability": {
            "temperature_strategy": "provider_default",
            "repetitions": 1,
            "aggregation": "median",
        },
        "rendering": {
            "dpi": 96,
            "contact_sheet_columns": 2,
            "max_pages": 10,
            "command_timeout_seconds": 30,
        },
        "aggregation": {"deterministic_weight": 0.7, "visual_weight": 0.3},
    }


class VisualConfigTest(unittest.TestCase):
    def test_loads_config_and_runtime_model_override(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "config.json"
            path.write_text(json.dumps(_payload()), encoding="utf-8")
            settings = load_visual_judge_settings(
                path, {"COWORK_EVO_VLM_MODEL": "runtime-model"}
            )
        self.assertEqual(settings.provider.name, "zcloud")
        self.assertEqual(settings.codex.model, "runtime-model")
        self.assertEqual(settings.codex.expected_version, "0.148.0")
        self.assertAlmostEqual(settings.weights.deterministic, 0.7)

    def test_loads_dotenv_without_shell_expansion(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / ".env"
            path.write_text(
                "# local\nexport COWORK_EVO_VLM_API_KEY='literal-$(id)'\n",
                encoding="utf-8",
            )
            values = load_env_file(path)
        self.assertEqual(values["COWORK_EVO_VLM_API_KEY"], "literal-$(id)")

    def test_resolves_secret_without_placing_it_in_settings(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "config.json"
            path.write_text(json.dumps(_payload()), encoding="utf-8")
            settings = load_visual_judge_settings(path, {})
        secret = resolve_api_key(settings, {"COWORK_EVO_VLM_API_KEY": "local-secret"}, {})
        self.assertEqual(secret, "local-secret")
        self.assertNotIn("local-secret", repr(settings))

    def test_rejects_non_zcloud_provider(self) -> None:
        payload = _payload()
        payload["provider"]["name"] = "other"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "config.json"
            path.write_text(json.dumps(payload), encoding="utf-8")
            with self.assertRaises(VisualConfigError):
                load_visual_judge_settings(path, {})


if __name__ == "__main__":
    unittest.main()
