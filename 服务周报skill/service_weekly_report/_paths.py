"""路径与凭据解析助手 —— 所有环境相关的位置统一在这里解析。

设计原则：
  1. 项目内部路径一律基于 ``__file__`` 推算，不写绝对路径。
  2. 项目外的工具（smartbi-data-cli、bi_skill）从环境变量读取，
     缺失时给出清晰的"如何设置"提示。
  3. 凭据只在运行时从环境变量读取，绝不出现在源码或文档里。
"""
from __future__ import annotations
import os
import sys
from pathlib import Path

PROJECT_ROOT: Path = Path(__file__).resolve().parent


def resolve_smartbi_cli_dir() -> Path:
    """定位 smartbi-data-cli-internal 工具目录。

    优先级：
      1. 环境变量 ``SMARTBI_CLI_DIR``
      2. 自动在项目父级 / 祖父级目录中查找 ``smartbi-data-cli-internal*``
    """
    env_dir = os.environ.get("SMARTBI_CLI_DIR")
    if env_dir:
        p = Path(env_dir)
        if p.exists():
            inner = p / p.name
            if (inner / "scripts" / "smartbi_cli.py").exists():
                return inner
            return p
    for parent in (PROJECT_ROOT.parent, PROJECT_ROOT.parent.parent):
        if not parent.exists():
            continue
        for child in parent.glob("smartbi-data-cli-internal*"):
            inner = child / child.name
            if (inner / "scripts" / "smartbi_cli.py").exists():
                return inner
            if (child / "scripts" / "smartbi_cli.py").exists():
                return child
    raise SystemExit(
        "[X] 未找到 smartbi-data-cli 工具目录。请设置环境变量：\n"
        '   PowerShell: $env:SMARTBI_CLI_DIR = "<smartbi-data-cli-internal-* 绝对路径>"'
    )


def resolve_bi_skill_path() -> Path:
    """定位 bi_skill.py。优先 ``BI_SKILL_PATH`` 环境变量，回退到 ``~/.workbuddy/skills/bi_skill/bi_skill.py``。"""
    env_path = os.environ.get("BI_SKILL_PATH")
    if env_path and Path(env_path).exists():
        return Path(env_path)
    return Path.home() / ".workbuddy" / "skills" / "bi_skill" / "bi_skill.py"


def ensure_credentials() -> None:
    """确认 SmartBI 凭据已通过环境变量配置。"""
    if os.environ.get("SMARTBI_USERNAME") and os.environ.get("SMARTBI_PASSWORD"):
        return
    sys.stderr.write(
        "[X] 缺少 SmartBI 凭据。请先在当前 PowerShell 会话设置环境变量：\n"
        '     $env:SMARTBI_USERNAME = "<your-username>"\n'
        '     $env:SMARTBI_PASSWORD = "<your-password>"\n'
        "切勿将凭据写入源码 / 配置 / 文档。\n"
    )
    sys.exit(1)
