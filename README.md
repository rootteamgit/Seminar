# RootTeam / Seminar

RootTeamのセミナー制作に関するルール・テンプレート・ナレッジを管理するリポジトリ。

---

## 📖 このリポジトリの使い方

### Claudeへの指示（必須）
セミナー関連の作業を開始する前に、必ず以下のURLをfetchすること:

```
https://raw.githubusercontent.com/rootteamgit/Seminar/main/rules/MASTER_RULES.md
```

### ルール更新の手順
1. ユーザーから新しい指示・ルールが出たら即座にMASTER_RULES.mdを更新
2. CHANGELOG.mdに変更を記録
3. GitHub APIでpush（同じ指示を二度させない）

---

## 📁 ディレクトリ構成

```
/
├── README.md               # このファイル
├── CHANGELOG.md            # 全ルール変更履歴
└── rules/
    └── MASTER_RULES.md     # 全ルール集約（作業前に必ず参照）
```

---

## ⚠️ 注意事項

- このリポジトリはPrivateに設定すること（社内ナレッジのため）
- PATはリポジトリに絶対にコミットしない
