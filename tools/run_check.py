#!/usr/bin/env python3
"""
手動実行ラッパー
使い方:
  python run_check.py --回 21
  python run_check.py --回 21 --dry-run

ファイルパスはリポジトリのルールに従い自動検出する。
"""
import os
import sys
import glob
import re
import argparse
import subprocess
from pathlib import Path

# ─── 設定：プロジェクトのルートパスに合わせて変更 ───
PROJECT_ROOT = os.environ.get("SEMINAR_PROJECT_ROOT", os.path.expanduser("~/seminar_project"))
AGENT_SCRIPT = os.path.join(os.path.dirname(__file__), "check_agent.py")

# ─── ファイル名パターン（第N回を含む） ───
DOCX_PATTERN  = f"{PROJECT_ROOT}/**/第{{N}}回*台本*.docx"
PPTX_PATTERN  = f"{PROJECT_ROOT}/**/第{{N}}回*資料*.pptx"
XLSX_PATTERNS = [
    f"{PROJECT_ROOT}/**/セミナー企画書.xlsx",
    f"{PROJECT_ROOT}/**/ウェビナー_TODO*.xlsx",
]


def find_file(pattern: str) -> str | None:
    results = glob.glob(pattern, recursive=True)
    # 最新バージョン（v番号が大きいもの）を優先
    if results:
        results.sort(key=lambda p: [int(n) for n in re.findall(r'\d+', p)], reverse=True)
        return results[0]
    return None


def main():
    parser = argparse.ArgumentParser(description="セミナーチェックエージェント 手動実行ラッパー")
    parser.add_argument("--回", type=int, required=True, dest="session_num", help="対象セミナーの回数")
    parser.add_argument("--xlsx", default=None, help="企画書xlsxパスを直接指定（省略時は自動検索）")
    parser.add_argument("--docx", default=None, help="台本docxパスを直接指定（省略時は自動検索）")
    parser.add_argument("--pptx", default=None, help="スライドpptxパスを直接指定（省略時は自動検索）")
    parser.add_argument("--dry-run", action="store_true", help="Slack通知なしで実行")
    parser.add_argument("--slack-webhook", default=os.environ.get("SLACK_WEBHOOK_URL", ""), help="Slack Webhook URL")
    args = parser.parse_args()

    n = args.session_num

    # ファイル自動検索
    xlsx = args.xlsx or next(
        (find_file(p) for p in XLSX_PATTERNS if find_file(p)), None
    )
    docx = args.docx or find_file(DOCX_PATTERN.format(N=n))
    pptx = args.pptx or find_file(PPTX_PATTERN.format(N=n))

    print(f"=== 第{n}回セミナー 乖離チェック ===")
    print(f"  企画書 : {xlsx or '見つかりません'}")
    print(f"  台本   : {docx or '見つかりません'}")
    print(f"  スライド: {pptx or '（省略）'}")
    print()

    if not xlsx:
        print("[エラー] 企画書xlsxが見つかりません。--xlsxで直接指定してください。")
        sys.exit(1)
    if not docx:
        print("[エラー] 台本docxが見つかりません。--docxで直接指定してください。")
        sys.exit(1)

    cmd = [sys.executable, AGENT_SCRIPT, xlsx, docx]
    if pptx:
        cmd.append(pptx)
    cmd += ["--回", str(n)]
    if args.dry_run:
        cmd.append("--dry-run")
    if args.slack_webhook:
        cmd += ["--slack-webhook", args.slack_webhook]

    result = subprocess.run(cmd)
    sys.exit(result.returncode)


if __name__ == "__main__":
    main()
