#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
安装 LibreOffice（M3 公式重算的依赖）。

背景：先前用 --scope user 静默安装被静默终止（进程消失且未创建任何目录），
改为系统级安装并完整捕获输出，便于定位失败原因。

用法：
    python scripts/install_libreoffice.py          # 前台安装（等待并打印结果）
    python scripts/install_libreoffice.py --wait 0 # 启动后不等待
"""
import subprocess
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def build_cmd(system_wide: bool) -> list[str]:
    cmd = [
        "winget", "install",
        "--id", "TheDocumentFoundation.LibreOffice", "-e",
        "--silent",
        "--accept-package-agreements",
        "--accept-source-agreements",
        "--disable-interactivity",
    ]
    if not system_wide:
        cmd.append("--scope")
        cmd.append("user")
    return cmd


def run(cmd: list[str], timeout: int = 600) -> int:
    print("执行:", " ".join(cmd))
    print("（LibreOffice 约 300MB，首次安装可能需要几分钟）\n")
    try:
        p = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except FileNotFoundError:
        print("❌ 未找到 winget，请手动安装：https://www.libreoffice.org/download/download/")
        return 2

    start = time.time()
    while True:
        line = p.stdout.readline()
        if line:
            print("   ", line.rstrip())
        if p.poll() is not None:
            # 读剩余输出
            for rest in p.stdout:
                print("   ", rest.rstrip())
            break
        if time.time() - start > timeout:
            p.kill()
            print(f"\n⏱ 超过 {timeout}s 仍未完成，已终止。请稍后重试。")
            return 3
    print(f"\n退出码: {p.returncode}（耗时 {time.time()-start:.0f}s）")
    return p.returncode


def main() -> int:
    # 先试系统级（去掉 --scope user）
    rc = run(build_cmd(system_wide=True))
    if rc == 0:
        print("\n✅ 安装完成")
        return 0
    print("\n⚠️ 系统级安装未成功（退出码 %s）。常见原因：需要管理员权限。" % rc)
    print("   可选方案：")
    print("   1) 用管理员身份打开 PowerShell 后重新运行本脚本")
    print("   2) 手动下载安装：https://www.libreoffice.org/download/download/")
    print("   3) 若公司环境禁用 winget，可下载 MSI 后 msiexec /i <file.msi> /quiet")
    return rc


if __name__ == "__main__":
    sys.exit(main())