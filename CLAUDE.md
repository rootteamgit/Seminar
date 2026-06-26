# CLAUDE.md — OneStreamセミナー制作 Claude Code 設定ファイル

このファイルはClaude Codeが起動時に自動で読み込むルール定義です。
**タスク開始前に必ずこのファイルと `rules/` 以下のmdをすべて読むこと。**

---

## このリポジトリの目的

OneStream（BtoB SaaS動画配信プラットフォーム）のウェビナーシリーズの台本・スライドを
品質基準を守りながら制作・管理するためのルールとツールを集約したリポジトリ。

---

## ⚠️ 作業開始前の必須チェックリスト

タスクを受けたら、コードを1行も書く前に以下を実行すること。

```bash
# 1. ルールファイルをすべて読む
cat rules/MASTER_RULES.md
cat rules/slide_design.md
cat rules/script.md
cat rules/kikakusho.md

# 2. 対象セミナーの企画書アジェンダを確認
#    → セミナー企画書.xlsx の該当シートを確認

# 3. 企画書アジェンダと照合してから制作開始
#    → 照合を怠ると第13回素材を第17回として出力するなどの重大ミスが発生する
```

---

## ルールファイル一覧

| ファイル | 内容 |
|---|---|
| `rules/MASTER_RULES.md` | 全ルール集約・最優先参照 |
| `rules/slide_design.md` | スライドデザイン仕様 |
| `rules/script.md` | 台本構造ルール |
| `rules/kikakusho.md` | 企画書読み取りルール |

---

## ツール一覧

| ファイル | 用途 |
|---|---|
| `tools/check_agent.py` | 企画書と台本・スライドの乖離チェック |
| `tools/run_check.py` | 手動実行ラッパー |
| `tools/pre-commit` | git pre-commitフック |
| `tools/update_rules.sh` | ミス発生時のルール自己更新スクリプト |

---

## ミスが発生したときの手順

```bash
# ルールを自己更新する
bash tools/update_rules.sh "<対象ファイル>" "<ミスの内容>" "<追加するルール>"

# 例：
bash tools/update_rules.sh "rules/slide_design.md" \
  "S5のカード背景色が白になっていた" \
  "カード背景色は必ず#1A2272を使うこと。#E8EEFFなどの薄い色は禁止。"
```

---

## 出力ファイルの命名規則

```
第{N}回セミナー台本_v{X}.docx
第{N}回セミナー資料_v{X}.pptx
```

バージョンは元ファイルを上書きせず v2/v3 等のサフィックスで保存すること。

---

## 重要な連絡先・背景情報

- **プロダクト**: OneStream（動画配信SaaS）
- **ターゲット**: SMB 10〜50名規模・営業/マーケ職種
- **主要登壇者**: 森永悠介（COO）、澤居宏紀（CEO）
- **MC**: 羽生みな美（外注アナウンサー・第9回〜）
- **共催**: 上出キャンディ（スクールライズ代表）
- **CTA URL**:
  - 無料トライアル: https://one-stream.io/register
  - 商談予約: https://one-stream.youcanbook.me/
