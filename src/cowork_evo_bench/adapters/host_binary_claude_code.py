"""把宿主机固定版本的 Claude Code 可执行文件注入 Harbor 容器。"""

from __future__ import annotations

import os
import re
import shlex
import shutil
import socket
from pathlib import Path

from harbor.agents.installed.claude_code import ClaudeCode
from harbor.environments.base import BaseEnvironment


class HostBinaryClaudeCode(ClaudeCode):
    """复用宿主机 Claude Code 二进制，避免每个 Trial 临时联网安装。"""

    _REMOTE_STAGING_PATH = "/tmp/cowork-evo-claude"
    _REMOTE_BINARY_PATH = "/usr/local/bin/claude"

    def __init__(
        self,
        *args,
        host_binary: str | Path | None = None,
        provider_host: str | None = None,
        **kwargs,
    ) -> None:
        configured = host_binary or os.environ.get("COWORK_EVO_CLAUDE_BINARY")
        discovered = configured or shutil.which("claude")
        if not discovered:
            raise FileNotFoundError(
                "没有找到宿主机 Claude Code；请设置 COWORK_EVO_CLAUDE_BINARY。"
            )

        self.host_binary = Path(discovered).expanduser().resolve()
        if not self.host_binary.is_file():
            raise FileNotFoundError(
                f"Claude Code 宿主机二进制不存在：{self.host_binary}"
            )

        self.provider_host = provider_host
        self.provider_ip: str | None = None
        if provider_host:
            if not re.fullmatch(r"[A-Za-z0-9.-]+", provider_host):
                raise ValueError(f"Provider 主机名不合法：{provider_host}")
            self.provider_ip = socket.gethostbyname(provider_host)

        super().__init__(*args, **kwargs)

    async def install(self, environment: BaseEnvironment) -> None:
        """上传固定二进制，并在容器内核对实际版本。"""

        if not await self._installed_claude_satisfies_version(environment):
            await environment.upload_file(
                self.host_binary,
                self._REMOTE_STAGING_PATH,
            )
            install_result = await environment.exec(
                command=(
                    "mkdir -p /usr/local/bin && "
                    f"cp {self._REMOTE_STAGING_PATH} {self._REMOTE_BINARY_PATH} && "
                    f"chmod 0755 {self._REMOTE_BINARY_PATH} && "
                    f"rm -f {self._REMOTE_STAGING_PATH}"
                ),
                user="root",
            )
            if install_result.return_code != 0:
                raise RuntimeError(
                    "无法把宿主机 Claude Code 安装到 Trial 容器："
                    f"{install_result.stderr or install_result.stdout}"
                )

        if self.provider_host and self.provider_ip:
            mapping = f"{self.provider_ip} {self.provider_host}"
            mapping_result = await environment.exec(
                command=f"printf '%s\\n' {shlex.quote(mapping)} >> /etc/hosts",
                user="root",
            )
            if mapping_result.return_code != 0:
                raise RuntimeError(
                    "无法把 Provider DNS 映射写入 Trial 容器："
                    f"{mapping_result.stderr or mapping_result.stdout}"
                )

        version_result = await environment.exec(
            command=f"{self._REMOTE_BINARY_PATH} --version",
        )
        if version_result.return_code != 0:
            raise RuntimeError(
                "Trial 容器中的 Claude Code 无法启动："
                f"{version_result.stderr or version_result.stdout}"
            )

        actual_version = self.parse_version(version_result.stdout or "")
        if self._version and actual_version != self._version:
            raise RuntimeError(
                "Claude Code 版本不一致："
                f"期望 {self._version}，实际 {actual_version}。"
            )
