#!/bin/bash
# =============================================================
# update_rules.sh — ミス発生時のルール自己更新スクリプト
#
# 使い方:
#   bash tools/update_rules.sh <対象ファイル> <ミスの内容> <追加するルール>
#
# 例:
#   bash tools/update_rules.sh "rules/slide_design.md" \
#     "S5のカード背景色が白になっていた" \
#     "カード背景色は必ず#1A2272を使うこと。#E8EEFFなどの薄い色は禁止。"
#
# Claude Codeから呼ぶ場合:
#   emaが「〇〇のミスがあった」と報告
#   → Claude Codeがこのスクリプトを実行してルールを更新
#   → git commit & push で自動保存
# =============================================================

set -e

TARGET_FILE="$1"
MISTAKE="$2"
NEW_RULE="$3"
TIMESTAMP=$(date '+%Y-%m-%d %H:%M')

# 引数チェック
if [[ -z "$TARGET_FILE" || -z "$MISTAKE" || -z "$NEW_RULE" ]]; then
    echo "使い方: bash tools/update_rules.sh <対象ファイル> <ミスの内容> <追加するルール>"
    echo ""
    echo "例:"
    echo '  bash tools/update_rules.sh "rules/slide_design.md" \'
    echo '    "S5のカード背景色が白になっていた" \'
    echo '    "カード背景色は必ず#1A2272を使うこと。"'
    exit 1
fi

# ファイル存在チェック
if [[ ! -f "$TARGET_FILE" ]]; then
    echo "[エラー] ファイルが見つかりません: $TARGET_FILE"
    exit 1
fi

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CHANGELOG="$REPO_ROOT/CHANGELOG.md"

echo ""
echo "╔══════════════════════════════════════════════════╗"
echo "║  ルール自己更新スクリプト 起動                   ║"
echo "╚══════════════════════════════════════════════════╝"
echo ""
echo "対象ファイル : $TARGET_FILE"
echo "ミスの内容   : $MISTAKE"
echo "追加するルール: $NEW_RULE"
echo ""

# ─── 1. 対象ファイルの「よくあるミス」セクションにルールを追記 ───
if grep -q "よくあるミス" "$TARGET_FILE"; then
    # 「よくあるミス」セクションの末尾に追記
    echo "- ❌ ${MISTAKE} → ${NEW_RULE}" >> "$TARGET_FILE"
    echo "[1/4] $TARGET_FILE の「よくあるミス」に追記しました"
else
    # セクションがなければファイル末尾に追記
    cat >> "$TARGET_FILE" << MDEOF

---

## よくあるミス（自動追記）

- ❌ ${MISTAKE} → ${NEW_RULE}
MDEOF
    echo "[1/4] $TARGET_FILE の末尾に「よくあるミス」セクションを追記しました"
fi

# ─── 2. MASTER_RULES.md にも反映（slide_design or script の場合）───
MASTER="$REPO_ROOT/rules/MASTER_RULES.md"
if [[ "$TARGET_FILE" == *"slide_design"* || "$TARGET_FILE" == *"script"* ]]; then
    echo "" >> "$MASTER"
    echo "<!-- auto-added $TIMESTAMP -->" >> "$MASTER"
    echo "<!-- ❌ $MISTAKE → $NEW_RULE -->" >> "$MASTER"
    echo "[2/4] MASTER_RULES.md にも記録しました"
else
    echo "[2/4] MASTER_RULES.md への記録はスキップ（対象外ファイル）"
fi

# ─── 3. CHANGELOG.md に履歴を追記 ───
CHANGELOG_ENTRY="
## $TIMESTAMP — $TARGET_FILE
**ミス内容:** $MISTAKE
**追加ルール:** $NEW_RULE
---"

# CHANGELOGの「## 形式」行の後に挿入
python3 - << PYEOF
import re

with open('$CHANGELOG', 'r') as f:
    content = f.read()

entry = """
## $TIMESTAMP — $TARGET_FILE
**ミス内容:** $MISTAKE
**追加ルール:** $NEW_RULE
---"""

# 最初の "---\n\n## " の後に挿入
insert_marker = "---\n\n## 2026"
if insert_marker in content:
    content = content.replace(insert_marker, f"---\n{entry}\n\n## 2026", 1)
else:
    content += entry

with open('$CHANGELOG', 'w') as f:
    f.write(content)
print("[3/4] CHANGELOG.md に履歴を追記しました")
PYEOF

# ─── 4. git commit & push ───
cd "$REPO_ROOT"

git add "$TARGET_FILE" "$MASTER" "$CHANGELOG" 2>/dev/null || true

COMMIT_MSG="fix(rules): ${MISTAKE:0:50} [auto]"
if git diff --cached --quiet; then
    echo "[4/4] 変更なし（git commitをスキップ）"
else
    git commit -m "$COMMIT_MSG"
    echo "[4/4] git commit 完了: $COMMIT_MSG"

    if git push 2>/dev/null; then
        echo "✅ push 完了"
    else
        echo "⚠️  push に失敗しました。手動で git push を実行してください。"
    fi
fi

echo ""
echo "────────────────────────────────────────────────────"
echo "✅ ルール更新完了"
echo "   $TARGET_FILE に「$NEW_RULE」を追記しました"
echo "────────────────────────────────────────────────────"
