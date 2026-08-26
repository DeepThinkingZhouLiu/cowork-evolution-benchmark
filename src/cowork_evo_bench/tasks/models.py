"""单题路径模型，只描述文件位置，不负责运行或评分。"""

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class TaskPackage:
    """一个 Harbor 兼容任务包的稳定路径视图。"""

    root: Path

    @property
    def instruction(self) -> Path:
        return self.root / "instruction.md"

    @property
    def config(self) -> Path:
        return self.root / "task.toml"

    @property
    def rubric(self) -> Path:
        return self.root / "rubric.json"

    @property
    def input_files(self) -> Path:
        return self.root / "data" / "input_files"

    @property
    def solution(self) -> Path:
        return self.root / "solution"

    @property
    def tests(self) -> Path:
        return self.root / "tests"
