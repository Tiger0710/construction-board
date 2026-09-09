# 月またぎ修正の仕様と検証

## 操作

工事は開始日～終了日を1つの案件として登録し、その期間のどの月でも同じ行を編集する。月移動は未保存の入力を保持する。別担当への移動は未保存確認、保存中の別担当移動は禁止。

2026-09-09追加: 光洋の担当枠 `id:光洋` は表示名を「佐藤」に変更。IDを維持するため、既存データと旧リンクを引き継ぐ。

案件の「基本の作業時間」は `default_shift:day/night` として保存。省略時は昼。未入力日の表示・日次編集の初期値・サイネージで共通に使用する。既存の明示的な日次データ（部分設定も含む）は従来どおり `day!==false` / `night===true` を維持し、基本設定変更時に上書きしない。土日の自動休みも保持。案件統合は異なる実効デフォルト同士を候補にせず、未設定と昼は互換扱い。

既存の別ID案件は勝手に統合しない。「案件をまとめる」で、同じ基本情報かつ連続・重複する期間を候補表示する。選択した案件の日次設定が同日に異なる場合は、残す設定を選ぶまで統合できない。候補提示・プレビューは変更なし、統合確認はローカル変更、ヘッダーの保存で反映。

## 保存・互換性

- `GET ?scope=all&month=YYMM&user=NAME`: 担当者の全期間 `{projects,daily,_revision,_warnings}`。
- 初回は `_YYMM_NAME.json` の全月を読み込む。初回保存で `_projects_NAME.json` に集約し、旧月別ファイルはバックアップとして残す。GETではファイルを作らない。
- 同じIDの旧月別記録は期間を統合し、基本情報は後の月を採用。同日の日次差分は該当月のデータを優先して警告表示する。実在の旧データ検証ではこの衝突は0件。
- 異なるIDはそのまま保持し、利用者が明示的にまとめる。IDを維持するため、期間を延ばしても再登録は不要。
- `PUT {scope:'all',month,user,revision,data}` は担当者の全期間を保存。応答は `{success,revision}`。画面の月は保存先を分割しない。
- 参照するGitコミットを固定し、対象ファイルSHA群からrevisionを作る。保存直前のrevision照合と、親コミットを固定したnon-force ref更新の両方で競合を検知。
- 初回保存以後、古い画面からの旧月別PUTは409で拒否し、更新後のページへの再読込を案内。旧月別GETは該当月の表示用に互換応答。
- 読込失敗・不正データ・途中失敗は空の成功応答にしない。入力画面は読込成功まで追加・保存を無効にし、保存失敗時は未保存内容を保持する。
- 旧Excel同期は廃止済み。旧月別ファイルを直接編集しても全月共通ファイルの作成後は反映されない。

## サイネージ・上限

全履歴の保存元から有効な案件を取得し、展開範囲だけを前月～翌月に限定する。数か月前に登録した長期工事も含む。同名の別工事を消す名前ベースの重複排除は使わず、担当者・案件IDを保つ。

Netlifyの同期関数は60秒、通常の要求・応答は6MBが上限。[公式設定資料](https://docs.netlify.com/build/functions/configuration/?fn-language=js)を確認し、アプリは要求・応答4MiB、取得全体23秒・1回7秒、同時読込6件、履歴512ファイル/32MiB、サイネージ12,000行/約4MiBで明示的に制限する。超過はエラーとして返し、一部データだけの正常表示を避ける。

Git更新は `force:false` のfast-forward制約を使用。[GitHub公式API資料](https://docs.github.com/en/enterprise-cloud%40latest/rest/git/refs?apiVersion=2026-03-10)参照。

## 検証コマンド

```
node --test tests/data-api.test.mjs tests/project-merge.test.cjs
node tests/company-rotation.cjs
node tests/month-boundary.cjs
node tests/month-merge-ui.cjs
node tests/default-shift.cjs
git diff --check
```

単体・統合45件、会社別表示14項目、月またぎ編集12項目、統合画面24項目、佐藤表示・昼夜デフォルト15項目、合計110項目PASS。ブラウザテストはPlaywright＋Chrome、モックAPIを使用。

実データの形式検証: 担当別26ファイル・7担当の全月統合後・旧集約1ファイルがすべてPASS。元データは未変更。

実際のGit書込と関数配信は、テストsiteのDraftと `DATA_BRANCH=test` を使い、専用の合成担当者で初回移行・翌月保存・再読込・古いrevisionと旧画面の拒否・夜デフォルトの保持など13項目PASS。作成した合成担当者ファイル2つは検証後に同じtestブランチから削除済み。本番反映前はDraftで利用者が操作確認する。

最新Draft: https://6aa143ebc5c96264d31aae3c--hsj-construction-board-test.netlify.app 。配信後の佐藤表示・昼夜選択・マニュアル・API正常応答・ブラウザ例外なしを確認。前Draftではサイネージ・入力（実データ116案件中、9月表示10行）・月移動も確認済み。`python ci_pipeline.py` で生成後に `netlify deploy --no-build --skip-functions-cache --dir deploy --functions netlify/functions --site <test-site-id>` で配信。`--no-build` と `--context` は併用不可。
