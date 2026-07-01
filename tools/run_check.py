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


def _download_gdrive_xlsx(file_id: str, dest_path: str) -> str:
    """
    Google DriveスプレッドシートをMCP経由ではなく
    download_file_content相当の処理でxlsxとしてダウンロードして保存する。

    実際のMCP呼び出しはClaude側で行うため、ここではキャッシュファイルの
    存在確認のみを行い、なければエラーを返す設計。

    将来的にはMCPのCLI統合でこの関数から直接呼び出せるようにする予定。
    現状の運用：
      1. Claudeが download_file_content(file_id) でxlsxを取得
      2. base64デコードして ~/Documents/agent/セミナー企画書.xlsx に保存
      3. run_check.py がそのキャッシュを読む
    """
    cache = Path(dest_path)
    if cache.exists():
        return str(cache)

    # キャッシュがない場合はエクスポートURLで試みる
    export_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    print(f"[企画書] エクスポートURLでダウンロード試行: {export_url}")
    try:
        cache.parent.mkdir(parents=True, exist_ok=True)
        urllib.request.urlretrieve(export_url, str(cache))
        import os
        if os.path.getsize(str(cache)) > 1000:
            print(f"[企画書] ダウンロード成功: {cache}")
            return str(cache)
        else:
            cache.unlink()
            raise RuntimeError("ダウンロードしたファイルが空または不正です（認証が必要な可能性）")
    except Exception as e:
        raise RuntimeError(
            f"Google Driveからの自動取得に失敗しました: {e}\n"
            f"対処法: Claudeに「企画書をダウンロードして ~/Documents/agent/ に保存して」と依頼してください"
        )


def main():
    parser = argparse.ArgumentParser(description="セミナーチェックエージェント 手動実行ラッパー")
    parser.add_argument("--回", type=int, required=True, dest="session_num", help="対象セミナーの回数")
    parser.add_argument("--xlsx", default=None, help="企画書xlsxパスを直接指定（省略時は自動検索）")
    parser.add_argument("--docx", default=None, help="台本docxパスを直接指定（省略時は自動検索）")
    parser.add_argument("--pptx", default=None, help="スライドpptxパスを直接指定（省略時は自動検索）")
    parser.add_argument("--gdrive-file-id", default="1s_7uFGfK2dQYAxMBhgQbH17gJfIe7DXny67TTJygeAc",
                        help="Google DriveのスプレッドシートID（デフォルト：セミナー企画書）")
    parser.add_argument("--dry-run", action="store_true", help="Slack通知なしで実行")
    parser.add_argument("--slack-webhook", default=os.environ.get("SLACK_WEBHOOK_URL", ""), help="Slack Webhook URL")
    parser.add_argument("--auto-push", action="store_true", help="チェックOK時に自動でgit add/commit/pushする")
    parser.add_argument("--skip-slide-check", action="store_true", help="スライドファイルなしでもpre-flight checkをスキップして続行")
    parser.add_argument("--commit-msg", default=None, help="git commitメッセージ（省略時は自動生成）")
    args = parser.parse_args()

    n = args.session_num

    # ─── 企画書の取得（優先順位）───
    # 1. --xlsx 直接指定
    # 2. ローカルキャッシュ（~/Documents/agent/セミナー企画書.xlsx）
    # 3. Google Driveから自動ダウンロード（download_file_contentで全シート取得）
    GDRIVE_LOCAL_CACHE = Path.home() / "Documents" / "agent" / "セミナー企画書.xlsx"

    if args.xlsx:
        xlsx = args.xlsx
    elif GDRIVE_LOCAL_CACHE.exists():
        xlsx = str(GDRIVE_LOCAL_CACHE)
        print(f"[企画書] ローカルキャッシュを使用: {xlsx}")
    else:
        # Google Driveから自動ダウンロード
        file_id = args.gdrive_file_id
        print(f"[企画書] Google Driveからダウンロード中... (file_id: {file_id})")
        try:
            xlsx = _download_gdrive_xlsx(file_id, str(GDRIVE_LOCAL_CACHE))
            print(f"[企画書] ダウンロード完了: {xlsx}")
        except Exception as e:
            xlsx = next((find_file(p) for p in XLSX_PATTERNS if find_file(p)), None)
            if not xlsx:
                print(f"[エラー] 企画書の取得に失敗しました: {e}")
                print("  対処法: ~/Documents/agent/セミナー企画書.xlsx に企画書を配置するか --xlsx で直接指定してください")
                sys.exit(1)

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
