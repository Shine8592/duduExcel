#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""duduExcel API 推送（git 直连不稳时的备用通道）。"""

import base64
import json
import os
import subprocess
import sys
import urllib.request
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
TOKEN = (os.environ.get("GITHUB_TOKEN") or "").strip()
REPO = "Shine8592/duduExcel"
BRANCH = "main"
ROOT = Path(__file__).resolve().parent.parent

if not TOKEN:
    print("缺少 GITHUB_TOKEN")
    raise SystemExit(2)

H = {"Authorization": f"token {TOKEN}", "User-Agent": "ps",
     "Accept": "application/vnd.github+json"}
OP = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def api(method, path, body=None):
    url = f"https://api.github.com/repos/{REPO}{path}"
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, headers=H, method=method)
    if body is not None:
        req.add_header("Content-Type", "application/json")
    with OP.open(req, timeout=60) as r:
        raw = r.read().decode()
        return json.loads(raw) if raw.strip() else {}


def main() -> int:
    head = subprocess.run(["git", "rev-parse", "HEAD"], cwd=ROOT,
                          capture_output=True, text=True).stdout.strip()
    changed = subprocess.run(
        ["git", "diff", "--name-only", "HEAD~1..HEAD", "--diff-filter=ACMR"],
        cwd=ROOT, capture_output=True, text=True,
    ).stdout.split()
    print(f"local HEAD: {head[:12]}")
    print(f"待推送: {len(changed)} 个文件")

    if not changed:
        print("无文件")
        return 0

    ref = api("GET", f"/git/ref/heads/{BRANCH}")
    base_sha = ref["object"]["sha"]
    commit = api("GET", f"/git/commits/{base_sha}")
    base_tree = commit["tree"]["sha"]
    print(f"remote base: {base_sha[:12]}")

    items = []
    for rel in sorted(changed):
        p = ROOT / rel
        if not p.is_file():
            print(f"  [SKIP] {rel}")
            continue
        raw = p.read_bytes()
        blob = api("POST", "/git/blobs", {
            "content": base64.b64encode(raw).decode(),
            "encoding": "base64",
        })
        items.append({"path": rel, "mode": "100644", "type": "blob",
                      "sha": blob["sha"]})
        print(f"  [BLOB] {rel} ({len(raw)}B)")

    if not items:
        return 1
    tree = api("POST", "/git/trees", {"base_tree": base_tree, "tree": items})
    msg = subprocess.run(["git", "log", "-1", "--pretty=%B"], cwd=ROOT,
                         capture_output=True, text=True, encoding="utf-8"
                         ).stdout.strip()
    nc = api("POST", "/git/commits", {"message": msg, "tree": tree["sha"],
                                      "parents": [base_sha]})
    api("PATCH", f"/git/refs/heads/{BRANCH}", {"sha": nc["sha"]})
    print(f"\n[OK] 已推送，commit {nc['sha'][:12]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
