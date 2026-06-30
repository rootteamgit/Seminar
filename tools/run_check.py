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
import tempfile
import urllib.request
from pathlib import Path

# ─── 設定：プロジェクトのルートパスに合わせて変更 ───
PROJECT_ROOT = os.environ.get("SEMINAR_PROJECT_ROOT", os.path.expanduser("~/seminar_project"))
AGENT_SCRIPT = os.path.join(os.path.dirname(__file__), "check_agent.py")

# ─── Google Sheets URL（企画書） ───
KIKAKUSHO_GDRIVE_URL = os.environ.get(
    "KIKAKUSHO_GDRIVE_URL",
    "https://docs.google.com/spreadsheets/d/1s_7uFGfK2dQYAxMBhgQbH17gJfIe7DXny67TTJygeAc/export?format=xlsx"
)

# ─── ファイル名パターン（第N回を含む） ───
DOCX_PATTERN  = f"{PROJECT_ROOT}/**/第{{N}}回*台本*.docx"
PPTX_PATTERN  = f"{PROJECT_ROOT}/**/第{{N}}回*資料*.pptx"
XLSX_PATTERNS = [
    f"{PROJECT_ROOT}/**/セミナー企画書.xlsx",
    f"{PROJECT_ROOT}/**/ウェビナー_TODO*.xlsx",
]


def download_kikakusho() -> str:
    """Google SheetsからセミナーI企画書をxlsxとしてダウンロードして一時ファイルパスを返す"""
    print("[企画書] Google Driveから最新版を取得中...")
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    try:
        urllib.request.urlretrieve(KIKAKUSHO_GDRIVE_URL, tmp.name)
        print(f"[企画書] ダウンロード完了: {tmp.name}")
        return tmp.name
    except Exception as e:
        print(f"[企画書] ダウンロード失敗: {e}")
        print("[企画書] ローカルファイルを探します...")
        return None


def find_file(pattern: str):
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
    parser.add_argument("--auto-push", action="store_true", help="チェックOK時に自動でgit add/commit/pushする")
    parser.add_argument("--commit-msg", default=None, help="git commitメッセージ（省略時は自動生成）")
    args = parser.parse_args()

    n = args.session_num

    # 企画書：引数指定 → ローカル検索 の順
    # Google Driveからは手動でダウンロードして --xlsx で指定するか
    # SEMINAR_KIKAKUSHO_PATH 環境変数で固定パスを設定してください
    if args.xlsx:
        xlsx = args.xlsx
    else:
        env_path = os.environ.get("SEMINAR_KIKAKUSHO_PATH")
        if env_path and Path(env_path).exists():
            xlsx = env_path
            print(f"[企画書] 環境変数から取得: {xlsx}")
        else:
            xlsx = next((find_file(p) for p in XLSX_PATTERNS if find_file(p)), None)
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

    # ─── 自動git push ───
    if args.auto_push and result.returncode == 0:
        print("\n[auto-push] チェックOK → git push を実行します...")

        # Seminarリポジトリのルートを探す
        repo_root = Path(__file__).parent.parent
        if not (repo_root / ".git").exists():
            repo_root = Path(os.environ.get("SEMINAR_REPO_ROOT", str(Path.home() / "Seminar")))

        if not (repo_root / ".git").exists():
            print(f"[auto-push] エラー: gitリポジトリが見つかりません: {repo_root}")
            sys.exit(result.returncode)

        # outputs/にファイルをコピー
        import shutil
        outputs_dir = repo_root / "outputs"
        outputs_dir.mkdir(exist_ok=True)

        files_to_add = []
        for f in [docx, pptx]:
            if f and Path(f).exists():
                dest = outputs_dir / Path(f).name
                shutil.copy2(f, dest)
                rel = str(dest.relative_to(repo_root))
                files_to_add.append(rel)
                print(f"[auto-push] コピー: {Path(f).name} → outputs/")

        if not files_to_add:
            print("[auto-push] コピーするファイルがありませんでした")
            sys.exit(result.returncode)

        # コミットメッセージ生成
        commit_msg = args.commit_msg or f"Add 第{n}回セミナー台本・スライド（乖離チェック済み）"

        try:
            subprocess.run(["git", "-C", str(repo_root), "add"] + files_to_add, check=True)
            subprocess.run(["git", "-C", str(repo_root), "commit", "-m", commit_msg], check=True)
            subprocess.run(["git", "-C", str(repo_root), "push", "origin", "master"], check=True)
            print(f"[auto-push] ✅ push完了: {commit_msg}")
        except subprocess.CalledProcessError as e:
            print(f"[auto-push] ❌ git操作に失敗しました: {e}")

    sys.exit(result.returncode)


if __name__ == "__main__":
    main()
