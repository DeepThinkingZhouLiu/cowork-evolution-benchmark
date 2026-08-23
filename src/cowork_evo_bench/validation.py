"""跨模块共用的轻量校验结果，不包含任何业务判断。"""

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class ValidationIssue:
    """一条可以稳定展示和测试的校验问题。"""

    code: str
    location: str
    message: str
