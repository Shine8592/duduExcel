#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
duduExcel 安装自检（面向用户，安装后运行）。

检查项：
1. 包与依赖是否可用
2. MCP 工具是否全部注册
3. MCP stdio 协议握手是否正常
4. LibreOffice 是否就绪（决定 recalculate 可用性）
5. 端到端冒烟：读 → 批量写 → 读 → 回滚

用法：
    python scripts/health_check.py
"""
import json
import subprocess
import sys
import tempfile
import time
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PY = sys.executable
OK = "[ OK ]"
WARN = "[WARN]"
FAIL = "[FAIL]"


def header(t: str) -> None:
    print(f"\n--- {t} ---")


def main() -> int:
    print("=" * 56)
    print(" duduExcel 安装自检")
    print("=" * 56)

    # 1. 包与依赖
    header("1. 包与依赖")
    try:
        import duduexcel  # noqa: F401
        import openpyxl   # noqa: F401
        print(f"{OK} duduexcel 导入成功")
        try:
            import pandas  # noqa: F401
            print(f"{OK} pandas 可用（分析工具完整）")
        except ImportError:
            print(f"{WARN} pandas 未安装 -> filter_count/aggregate/top_n/join/compare 不可用")
            print("       补齐：pip install 'duduexcel[analysis]'")
    except Exception as e:
        print(f"{FAIL} 导入失败: {e}")
        return 1

    # 2. 工具注册 + 协议握手
    header("2. MCP 工具注册与协议握手")
    proc = subprocess.Popen(
        [PY, "-u", "-m", "duduexcel"],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, encoding="utf-8", errors="replace", bufsize=1,
    )
    _id = [0]

    def send(obj: dict):
        _id[0] += 1
        obj["id"] = _id[0]
        proc.stdin.write(json.dumps(obj) + "\n")
        proc.stdin.flush()
        deadline = time.time() + 60
        while time.time() < deadline:
            line = proc.stdout.readline()
            if not line:
                return None
            try:
                j = json.loads(line.strip())
                if j.get("id") == _id[0]:
                    return j
            except Exception:
                continue
        return None

    init = send({"jsonrpc": "2.0", "method": "initialize",
                 "params": {"protocolVersion": "2024-11-05", "capabilities": {},
                            "clientInfo": {"name": "health", "version": "1.0"}}})
    if init and "result" in init:
        print(f"{OK} initialize 握手成功 (serverInfo={init['result'].get('serverInfo')})")
    else:
        print(f"{FAIL} initialize 握手失败")
        proc.kill()
        return 1

    tl = send({"jsonrpc": "2.0", "method": "tools/list", "params": {}})
    tools = [t["name"] for t in (tl or {}).get("result", {}).get("tools", [])]
    if tools:
        print(f"{OK} 注册 {len(tools)} 个工具")
        print("       " + ", ".join(tools))
    else:
        print(f"{FAIL} 未获取到工具列表")

    def call(name, args=None):
        r = send({"jsonrpc": "2.0", "method": "tools/call",
                  "params": {"name": name, "arguments": args or {}}})
        if r is None:
            return {"__error__": "TIMEOUT"}
        if "error" in r:
            return {"__error__": r["error"].get("message", "")}
        txt = r["result"]["content"][0].get("text", "")
        try:
            return json.loads(txt)
        except Exception:
            return {"__raw__": txt[:200]}

    # 3. 端到端冒烟
    header("3. 端到端冒烟")
    with tempfile.TemporaryDirectory() as td:
        xlsx = Path(td) / "health.xlsx"
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "数据"
        ws.append(["产品", "销量"])
        ws.append(["A", 10])
        ws.append(["B", 20])
        wb.save(xlsx)

        r = call("workbook_info", {"file_path": str(xlsx)})
        print(f"{OK if r.get('sheet_count') == 1 else FAIL} workbook_info -> sheets={r.get('sheet_count')}")

        r = call("write_cells", {"file_path": str(xlsx), "sheet": "数据",
                                 "cells": [{"cell": "A4", "value": "C"}, {"cell": "B4", "value": 30}]})
        print(f"{OK if r.get('written_count') == 2 else FAIL} write_cells -> {r.get('written_count')} 格")

        r = call("read_range", {"file_path": str(xlsx), "sheet": "数据", "limit": 10})
        print(f"{OK if len(r.get('grid', [])) == 4 else FAIL} read_range -> {len(r.get('grid', []))} 行")

        r = call("revert_last_write", {"file_path": str(xlsx)})
        print(f"{OK if '已回滚' in str(r.get('status', '')) else FAIL} revert_last_write -> {r.get('status')}")

    proc.kill()

    # 4. LibreOffice
    header("4. 公式重算环境")
    try:
        from duduexcel.recalc import _find_soffice
        so = _find_soffice()
        if so:
            print(f"{OK} LibreOffice 已就绪 -> recalculate 可用")
            print(f"       {so}")
        else:
            print(f"{WARN} 未检测到 LibreOffice -> recalculate 会明确降级（不会静默假装成功）")
            print("       安装：winget install --id TheDocumentFoundation.LibreOffice -e")
            print("       或访问 https://www.libreoffice.org/download/download/")
    except Exception as e:
        print(f"{WARN} 检测失败: {e}")

    print("\n" + "=" * 56)
    print(" 自检完成")
    print("=" * 56)
    return 0


if __name__ == "__main__":
    sys.exit(main())
