# セミナー再発防止エージェント — セットアップ手順

---

## 概要

企画書（xlsx）のアジェンダを**正とした照合**を自動実行し、台本（docx）・スライド（pptx）との乖離をSlackに通知します。

```
企画書.xlsx
    ↓ 照合
台本.docx   → Claude API で乖離判定 → Slack通知
スライド.pptx
```

---

## ファイル構成

```
tools/
├── check_agent.py   ← 本体（これ単体で動く）
├── run_check.py     ← 手動実行ラッパー
├── pre-commit       ← git pre-commitフック
└── README.md        ← このファイル
```

---

## セットアップ（初回のみ）

### 1. 依存ライブラリのインストール

```bash
pip install python-docx openpyxl python-pptx slack-sdk --break-system-packages
```

### 2. Slack Webhook URLを環境変数に設定

SlackのIncoming Webhookを作成し `.bashrc` または `.zshrc` に追記：

```bash
export SLACK_WEBHOOK_URL=https://hooks.slack.com/services/XXXX/YYYY/ZZZZ
```

Slack Webhookの作成手順：
1. Slack → アプリ管理 → 「Incoming WebHooks」を追加
2. チャンネルを選択（例: `#セミナー制作`）
3. Webhook URLをコピー

### 3. git pre-commitフックを設置

```bash
# リポジトリルートで実行
cp tools/pre-commit .git/hooks/pre-commit
chmod +x .git/hooks/pre-commit
```

### 4. プロジェクトルートの設定（run_check.py 用）

```bash
export SEMINAR_PROJECT_ROOT=/path/to/your/seminar/folder
```

または `run_check.py` の `PROJECT_ROOT` を直接書き換え：
```python
PROJECT_ROOT = "/path/to/your/seminar/folder"
```

---

## 使い方

### A. git commitに自動で付随（推奨）

台本・スライドをステージングしてcommitするだけ：

```bash
git add 第21回セミナー台本_v1.docx
git add 第21回セミナー資料_v1.pptx
git commit -m "第21回：台本・スライド初稿"
# → 自動でチェックが走る
# → 乖離があればSlack通知 + commitブロック
```

乖離を無視して強制commit（緊急時のみ）：
```bash
git commit --no-verify -m "..."
```

### B. 手動で単発実行（Claude Codeのターミナルから）

```bash
# 第21回をチェック（ファイル自動検索）
python tools/run_check.py --回 21

# Slack通知なしで確認だけ
python tools/run_check.py --回 21 --dry-run

# ファイルを直接指定
python tools/check_agent.py セミナー企画書.xlsx 第21回セミナー台本_v1.docx 第21回セミナー資料_v1.pptx --回 21
```

---

## Slack通知の例

### 乖離なしの場合

```
✅ 企画書との乖離なし：第21回
サマリー：アジェンダの全項目が台本・スライドに含まれており問題ありません。
チェック対象ファイル：
• 企画書: セミナー企画書.xlsx
• 台本: 第21回セミナー台本_v1.docx
実行日時: 2026-07-01 14:30
```

### 乖離ありの場合

```
🚨 乖離を検出：第21回（2件）
サマリー：台本に企画書にない章が含まれており、修正が必要です。

🔴 [高] 台本②章のタイトルが「GTM共催の背景」となっているが、
         企画書アジェンダでは「少人数で商談を増やす仕組み」。
🟡 [中] スライドに企画書にない「OpenClaw料金プラン」ページが含まれている。
```

---

## 判定ロジック

| 判定 | 内容 | exitコード |
|---|---|---|
| 乖離なし | 全アジェンダ項目が台本・スライドに対応 | 0（commit続行） |
| 中/低の乖離 | Slack通知のみ | 0（commit続行） |
| 高の乖離あり | Slack通知 + commitブロック | 1（commit中断） |

---

## トラブルシューティング

**企画書からアジェンダが取得できない**
→ `check_agent.py` の `extract_agenda_from_xlsx()` を開き、`in_agenda` を判定するキーワードを企画書の実際のセル値に合わせて調整してください。

**Claude APIが応答しない**
→ `SLACK_WEBHOOK_URL` と同様に `ANTHROPIC_API_KEY` が必要な場合は、`check_agent.py` の `urllib.request.Request` ヘッダーに `"x-api-key"` を追加してください。

**Slack通知だけテストしたい**
```bash
python tools/check_agent.py セミナー企画書.xlsx 第21回台本.docx --dry-run
```
