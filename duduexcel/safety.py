#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
安全层：路径白名单 + 写前自动备份。

设计要点（汲取自调研）：
- 路径白名单：借鉴 haris-musa/excel-mcp-server 的 EXCEL_FILES_PATH 思路，
  设了 DUDU_EXCEL_ROOT 后强制所有路径落在根目录内，拒绝绝对路径与目录穿越。
  未设置时（本地 stdio 模式）放行，由用户自行负责——与 haris-musa 的
  "stdio 模式由客户端传路径、服务端不限制"保持一致。
- 写前自动备份：汲取 Anthropic 官方 xlsx skill 对"不可逆操作"的护栏思想
  （它的 recalc.py 会为外链丢失直接拒绝执行）。Excel 写入没有回收站，
  因此任何写操作前先落一份 .bak，并提供 revert_last() 回滚。
"""

import os
import shutil
from pathlib import Path

# 环境变量：限制本服务可访问的根目录（HTTP 远程模式强烈建议设置）
ROOT_ENV = "DUDU_EXCEL_ROOT"
# 单次写入的备份后缀
BACKUP_SUFFIX = ".bak"


class SafetyError(Exception):
    """路径越界或安全检查失败时抛出，消息面向 Agent 保持可读、可纠正。"""


def get_root() -> Path | None:
    """返回配置的根目录；未设置返回 None（表示不限制）。"""
    raw = os.environ.get(ROOT_ENV, "").strip()
    if not raw:
        return None
    return Path(raw).resolve()


def resolve_path(file_path: str) -> Path:
    """把用户传入的路径解析为安全的绝对路径。

    规则：
    1. 设置了 DUDU_EXCEL_ROOT 时，路径必须是相对路径或位于根目录内；
       拒绝绝对路径与 `..` 目录穿越。
    2. 未设置根目录时直接使用传入路径（本地 stdio 默认模式）。
    """
    if not file_path or not str(file_path).strip():
        raise SafetyError("文件路径为空，请提供 .xlsx 文件路径")

    root = get_root()
    raw = Path(file_path)

    if root is not None:
        # 绝对路径：必须在根目录内
        if raw.is_absolute():
            resolved = raw.resolve()
        else:
            resolved = (root / raw).resolve()
        try:
            resolved.relative_to(root)
        except ValueError:
            raise SafetyError(
                f"路径越界：{file_path} 不在允许的根目录 {root} 内。"
                f"请改用相对路径，或调整环境变量 {ROOT_ENV}。"
            )
        return resolved

    # 无根目录限制：仅做基本展开
    return Path(os.path.expanduser(str(file_path))).resolve()


def require_exists(path: Path) -> None:
    """确认文件存在且是文件，否则给出可读错误。"""
    if not path.exists():
        raise SafetyError(f"文件不存在：{path}")
    if not path.is_file():
        raise SafetyError(f"该路径不是文件：{path}")


def backup_file(path: Path) -> str:
    """写操作前备份，返回备份文件路径字符串。

    采用覆盖式单备份（只保留最近一次），避免 .bak 无限堆积；
    配合 revert_last() 可回滚最近一次写入。
    """
    bak = path.with_suffix(path.suffix + BACKUP_SUFFIX)
    try:
        shutil.copy2(path, bak)
        return str(bak)
    except Exception as e:
        # 备份失败不应静默放过：写操作不可逆，必须让调用方知道没有安全网
        raise SafetyError(f"写入前备份失败，已中止写入以避免不可逆损坏：{e}")


def revert_last(path: Path) -> str:
    """用 .bak 回滚最近一次写入。"""
    bak = path.with_suffix(path.suffix + BACKUP_SUFFIX)
    if not bak.exists():
        raise SafetyError(f"没有找到备份文件：{bak}（该文件尚未被本服务写入过）")
    shutil.copy2(bak, path)
    return str(bak)
