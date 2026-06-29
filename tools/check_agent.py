#!/usr/bin/env python3
"""
セミナー再発防止エージェント
企画書（xlsx）のアジェンダと台本（docx）・スライド（pptx）の乖離をチェックし、
乖離があればSlackに通知する。

使い方:
  python check_agent.py <企画書.xlsx> <台本.docx> [スライド.pptx] [--回 N]

Claude Codeでの自動実行:
  git commit時にpre-commitフックから呼ばれる想定
"""

import sys
import os
import re
import json
import argparse
import textwrap
from pathlib import Path
from datetime import datetime

# ─────────────────────────────────────────────
# 1. テキスト抽出
# ─────────────────────────────────────────────

def extract_agenda_from_xlsx(xlsx_path: str, session_num: int = None) -> dict:
    """
    企画書xlsxからアジェンダを抽出する。
    セミナー企画書.xlsx の構造：
      - シート「企画書」またはシート名に回数を含む
      - 「アジェンダ」行を探して以降の項目を取得
      - 「テーマ」「登壇者」「ターゲット」も取得
    """
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)

    result = {
        "theme": "",
        "speaker": "",
        "target": "",
        "agenda_items": [],
        "composition": "",   # 構成欄（台本生成の詳細指示）
        "lead_text": "",     # リード文
        "solve_issues": "",  # 解決できる課題
        "raw_text": []
    }

    # 対象シートを選択
    sheet = None
    for ws in wb.worksheets:
        name = ws.title
        # 回数指定がある場合は回数を含むシートを優先
        if session_num and (f"第{session_num}回" in name) and "TODO" not in name:
            sheet = ws
            break
    if sheet is None:
        # 「企画書」シートを探す
        for ws in wb.worksheets:
            if "企画書" in ws.title:
                sheet = ws
                break
    if sheet is None:
        sheet = wb.active

    for row in sheet.iter_rows(values_only=True):
        cells = [c for c in row if c is not None and str(c).strip() and str(c) != "None"]
        if not cells:
            continue
        row_text = " ".join(str(c) for c in cells)
        result["raw_text"].append(row_text)

        first = str(cells[0]).strip()
        second = str(cells[1]).strip() if len(cells) > 1 else ""

        # テーマ行
        if any(k in first for k in ["セミナーテーマ", "タイトル"]):
            result["theme"] = second

        # 登壇者行
        if "登壇者名" in first and "ふりがな" not in first:
            if result["speaker"] == "":
                result["speaker"] = second

        # ターゲット行
        if "ターゲット" in first:
            result["target"] = second

        # アジェンダ行：1セルに全アジェンダが改行で入っている
        if first == "アジェンダ" and second:
            for line in second.split("\n"):
                line = line.strip()
                if line and re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩\d]', line):
                    result["agenda_items"].append(line)

        # 構成欄（常に取得・台本生成の詳細指示として使う）
        if "構成" in first and "ページ" in first and second:
            result["composition"] = second
        elif first == "構成" and second:
            result["composition"] = second
            # アジェンダが取れていない場合は構成欄から補完
            if not result["agenda_items"]:
                for line in second.split("\n"):
                    line = line.strip()
                    if line and re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩\d]', line):
                        result["agenda_items"].append(line)

        # リード文
        if "リード文" in first and second:
            result["lead_text"] = second[:300]

        # 解決できる課題
        if "解決できる課題" in first and second:
            result["solve_issues"] = second[:300]

    wb.close()
    return result


def extract_text_from_docx(docx_path: str) -> str:
    """台本docx（またはmdテキスト）から全テキストを抽出"""
    import zipfile
    # ZIPかどうかで本物のdocxか判定
    try:
        with zipfile.ZipFile(docx_path):
            pass
        # 本物のdocx
        import docx as docx_lib
        doc = docx_lib.Document(docx_path)
        lines = []
        for para in doc.paragraphs:
            if para.text.strip():
                lines.append(para.text.strip())
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        lines.append(cell.text.strip())
        return "\n".join(lines)
    except (zipfile.BadZipFile, Exception):
        # テキスト/markdownファイルとして読み込み
        with open(docx_path, "r", encoding="utf-8", errors="replace") as f:
            return f.read()


def extract_text_from_pptx(pptx_path: str) -> list[str]:
    """スライドpptxから各スライドのタイトル・本文を抽出"""
    from pptx import Presentation
    from pptx.util import Pt
    prs = Presentation(pptx_path)
    slides_text = []
    for i, slide in enumerate(prs.slides, 1):
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = para.text.strip()
                    if t:
                        texts.append(t)
        if texts:
            slides_text.append(f"[スライド{i}] " + " / ".join(texts[:5]))
    return slides_text


# ─────────────────────────────────────────────
# 2. Claude API で乖離チェック
# ─────────────────────────────────────────────

def check_deviation_with_claude(
    agenda_data: dict,
    script_text: str,
    slide_texts: list[str],
    session_label: str
) -> dict:
    """
    Claude API（claude-sonnet-4-6）を使って乖離を判定する。
    戻り値: {"ok": bool, "issues": [...], "summary": str}
    """
    import urllib.request

    agenda_str = "\n".join(f"  - {item}" for item in agenda_data["agenda_items"]) or "（取得できませんでした）"
    composition_str = agenda_data.get("composition", "")
    lead_str = agenda_data.get("lead_text", "")
    solve_str = agenda_data.get("solve_issues", "")
    slide_str = "\n".join(slide_texts[:30]) if slide_texts else "（スライドなし）"
    # 台本は長いので最初の3000文字
    script_excerpt = script_text[:3000] if script_text else "（台本なし）"

    prompt = f"""あなたはセミナー制作の品質チェック担当です。
以下の「企画書アジェンダ・構成」と「台本・スライドの内容」を照合し、乖離・不整合を検出してください。

## 対象セミナー
{session_label}
テーマ：{agenda_data['theme']}
登壇者：{agenda_data['speaker']}

## 企画書のアジェンダ（正とする）
{agenda_str}

## 企画書の構成欄（台本の詳細指示）
{composition_str if composition_str else "（記載なし）"}

## 企画書のリード文（ターゲットの課題感）
{lead_str if lead_str else "（記載なし）"}

## 企画書の解決できる課題
{solve_str if solve_str else "（記載なし）"}

## 台本の冒頭抜粋（最初3000文字）
{script_excerpt}

## スライドのテキスト抜粋（各スライドの要素）
{slide_str}

## 判定基準
1. 台本・スライドに、企画書のアジェンダにない章・トピックが含まれていないか？
2. 企画書のアジェンダ項目が台本・スライドで欠落していないか？
3. 企画書の構成欄の指示（登壇形式・各章の内容指示）が台本に反映されているか？
4. 企画書のリード文・解決できる課題が台本の内容にカバーされているか？
5. 登壇者・テーマが台本・スライドと一致しているか？

## 出力形式（JSONのみ・マークダウン不要）
{{
  "ok": true or false,
  "issues": [
    {{"severity": "高/中/低", "description": "問題の説明（具体的に）"}}
  ],
  "summary": "全体の判定サマリー（2〜3文）"
}}

乖離がなければ "ok": true、issues: [] としてください。"""

    payload = json.dumps({
        "model": "claude-sonnet-4-6",
        "max_tokens": 1000,
        "messages": [{"role": "user", "content": prompt}]
    }).encode("utf-8")

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    headers = {
        "Content-Type": "application/json",
        "anthropic-version": "2023-06-01"
    }
    if api_key:
        headers["x-api-key"] = api_key
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers=headers
    )

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = json.loads(resp.read().decode("utf-8"))
            text = body["content"][0]["text"]
            # JSON部分だけ取り出す
            text = re.sub(r"```json|```", "", text).strip()
            return json.loads(text)
    except Exception as e:
        return {
            "ok": False,
            "issues": [{"severity": "高", "description": f"Claude API呼び出しエラー: {e}"}],
            "summary": "チェックAPIへの接続に失敗しました。"
        }


# ─────────────────────────────────────────────
# 3. Slack通知
# ─────────────────────────────────────────────

def notify_slack(webhook_url: str, result: dict, session_label: str, files: dict):
    """乖離チェック結果をSlackに投稿する"""
    import urllib.request

    if result["ok"]:
        icon = "✅"
        color = "#36a64f"
        header = f"{icon} 企画書との乖離なし：{session_label}"
    else:
        high_count = sum(1 for i in result["issues"] if i["severity"] == "高")
        icon = "🚨" if high_count > 0 else "⚠️"
        color = "#ff0000" if high_count > 0 else "#ffa500"
        header = f"{icon} 乖離を検出：{session_label}（{len(result['issues'])}件）"

    # issues を整形
    issues_text = ""
    for issue in result.get("issues", []):
        sev = issue["severity"]
        sev_icon = {"高": "🔴", "中": "🟡", "低": "🔵"}.get(sev, "•")
        issues_text += f"{sev_icon} [{sev}] {issue['description']}\n"

    files_text = "\n".join(f"• {k}: `{v}`" for k, v in files.items())
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

    blocks = [
        {
            "type": "header",
            "text": {"type": "plain_text", "text": header, "emoji": True}
        },
        {
            "type": "section",
            "text": {"type": "mrkdwn", "text": f"*サマリー：* {result['summary']}"}
        }
    ]

    if issues_text:
        blocks.append({
            "type": "section",
            "text": {"type": "mrkdwn", "text": f"*検出された問題：*\n{issues_text.strip()}"}
        })

    blocks += [
        {"type": "divider"},
        {
            "type": "section",
            "text": {"type": "mrkdwn", "text": f"*チェック対象ファイル：*\n{files_text}\n\n_実行日時: {timestamp}_"}
        }
    ]

    payload = json.dumps({"blocks": blocks}).encode("utf-8")
    req = urllib.request.Request(
        webhook_url,
        data=payload,
        headers={"Content-Type": "application/json"}
    )
    try:
        urllib.request.urlopen(req, timeout=10)
        print("[Slack] 通知送信完了")
    except Exception as e:
        print(f"[Slack] 通知失敗: {e}", file=sys.stderr)


# ─────────────────────────────────────────────
# 4. メイン
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="セミナー再発防止エージェント")
    parser.add_argument("xlsx",   help="企画書 .xlsx のパス")
    parser.add_argument("docx",   help="台本 .docx のパス")
    parser.add_argument("pptx",   nargs="?", default=None, help="スライド .pptx のパス（省略可）")
    parser.add_argument("--回",   type=int, default=None, dest="session_num", help="セミナー回数（例: 21）")
    parser.add_argument("--slack-webhook", default=os.environ.get("SLACK_WEBHOOK_URL", ""),
                        help="Slack Incoming Webhook URL（環境変数 SLACK_WEBHOOK_URL でも設定可）")
    parser.add_argument("--dry-run", action="store_true", help="Slack通知せず結果をターミナルに出力")
    args = parser.parse_args()

    # ファイル存在確認
    for path, label in [(args.xlsx, "企画書"), (args.docx, "台本")]:
        if not Path(path).exists():
            print(f"[エラー] {label}ファイルが見つかりません: {path}", file=sys.stderr)
            sys.exit(1)

    session_label = f"第{args.session_num}回" if args.session_num else Path(args.docx).stem

    print(f"[チェック開始] {session_label}")
    print(f"  企画書: {args.xlsx}")
    print(f"  台本  : {args.docx}")
    if args.pptx:
        print(f"  スライド: {args.pptx}")

    # テキスト抽出
    print("[1/3] テキスト抽出中...")
    agenda_data = extract_agenda_from_xlsx(args.xlsx, args.session_num)
    script_text = extract_text_from_docx(args.docx)
    slide_texts = extract_text_from_pptx(args.pptx) if args.pptx and Path(args.pptx).exists() else []

    if not agenda_data["agenda_items"]:
        print("[警告] 企画書からアジェンダ項目を抽出できませんでした。raw_text で代替します。")
        # raw_text をフォールバック用アジェンダとして使用
        agenda_data["agenda_items"] = agenda_data["raw_text"][:20]

    print(f"  → アジェンダ項目: {len(agenda_data['agenda_items'])}件")
    print(f"  → 台本テキスト: {len(script_text)}文字")
    print(f"  → スライド枚数: {len(slide_texts)}枚")

    # Claude API でチェック
    print("[2/3] Claude API でチェック中...")
    result = check_deviation_with_claude(agenda_data, script_text, slide_texts, session_label)

    # 結果表示
    print("\n─── チェック結果 ───")
    if result["ok"]:
        print("✅ 乖離なし")
    else:
        print(f"⚠️  {len(result['issues'])}件の問題を検出")
        for issue in result["issues"]:
            print(f"  [{issue['severity']}] {issue['description']}")
    print(f"サマリー: {result['summary']}")
    print("─────────────────────\n")

    # Slack通知
    files_info = {"企画書": args.xlsx, "台本": args.docx}
    if args.pptx:
        files_info["スライド"] = args.pptx

    if args.dry_run:
        print("[dry-run] Slack通知はスキップしました")
    elif args.slack_webhook:
        print("[3/3] Slack通知中...")
        notify_slack(args.slack_webhook, result, session_label, files_info)
    else:
        print("[3/3] SLACK_WEBHOOK_URL が未設定のためSlack通知をスキップ")
        print("      設定方法: export SLACK_WEBHOOK_URL=https://hooks.slack.com/services/...")

    # 乖離あり → exit code 1（pre-commitフックでcommitブロックに使う）
    if not result["ok"]:
        high_issues = [i for i in result["issues"] if i["severity"] == "高"]
        if high_issues:
            print("[pre-commit] 重大な乖離があるためcommitを中断します。", file=sys.stderr)
            sys.exit(1)

    sys.exit(0)


if __name__ == "__main__":
    main()
