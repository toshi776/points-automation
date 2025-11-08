# Work Log - points-automation

## 2025-11-08
- `mail_client.py` を大幅拡張し、メール本文を `◆...◆` 単位で分割するロジックを実装。
- 各ブロックごとにジャンル判定・ポイント抽出・除外理由を判定するように変更。
- `mail_client.py analyze --folder モッピー --days 7 --output logs/moppy_blocks.json --rejected-output logs/moppy_rejected.json` を実行し、抽出ブロック32件・除外ブロック67件を JSON で出力。
- `logs/moppy_summary.txt` に「日時／概要／詳細／還元額」形式で抽出・除外一覧をまとめるスクリプトを作成。

### 次の作業予定
1. Codex を用いた第2段階フィルタリング（候補ブロックに対して LLM 判定）用のスクリプト雛形を作る。
2. 投稿文テンプレート（タイトル＋要約＋還元額＋ハッシュタグ＋PR表記）の仕様を決め、Codex プロンプトに組み込む。
3. `refined_posts.json` のような最終出力を想定し、X 投稿スクリプトに渡せる形へ整形する。

