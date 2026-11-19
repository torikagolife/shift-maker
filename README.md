# シフト作成 Web（shift-maker）無料プラン版

希望休の入力 → 自動作成 → チェックまでをブラウザでやるためのツールです。
**Cloudflare の無料プランだけで動きます**（Workers ＋ D1 ＋ 静的ファイル）。

- 画面・パスワード・データ保存：Cloudflare Workers ＋ D1
- 自動作成とチェックの計算：**ブラウザの中**（Pyodide で Python を動かし、scipy の HiGHS で解く）
- 勤務種類・職員・必要人数・ルールの設定は Excel（設定.xlsx）で書いて、画面から登録します

職員のデータは計算のために外部のサーバーへ送られません（保存先は自分の D1 だけ）。

## ファイル構成

```
wrangler.jsonc              Worker・静的ファイル・D1 の設定
package.json                wrangler
src/index.js                Worker（パスワード確認・API・D1）
public/index.html           画面
public/engine-worker.js     計算エンジン（ブラウザの裏側で Python を動かす Web Worker）
public/py/shift_engine.py   エンジン本体（PC でコマンドとしても使える）
public/py/solver_highs.py   ソルバー（scipy の HiGHS 版）
public/py/engine_web.py     画面から呼ぶための入口
```

## デプロイ手順（ブラウザだけでできる方法）

1. **GitHub にリポジトリを作る**（Private 推奨）
   このフォルダの中身を「Add file → Upload files」でドラッグ＆ドロップ。
   `node_modules` と `.dev.vars` はアップロードしない。

2. **D1 データベースを作る**
   Cloudflare ダッシュボード → ストレージとデータベース → D1 → 作成。名前は `shift-maker`。
   できたら「データベースID」をコピーして、GitHub 上で `wrangler.jsonc` の
   `"database_id": "ここにD1のIDを貼る"` を書き換えてコミット。
   テーブルは初回アクセス時に自動で作られます。

3. **Worker を作る**
   Workers & Pages → 作成 → 「リポジトリをインポート」→ GitHub のリポジトリを選ぶ。
   プロジェクト名は `shift-maker`（wrangler.jsonc の name と同じ）。
   ビルドコマンドは空欄、デプロイコマンドは `npx wrangler deploy` のままでOK。

4. **パスワードを設定する**
   作った Worker → 設定 → 変数とシークレット → 追加
   種類「シークレット」、名前 `APP_PASSWORD`、値は好きなパスワード。

5. **使う**
   Worker の URL を開く → パスワードを入れる（ユーザー名は何でもよい）
   → 右上が「計算エンジン準備OK」になるのを待つ → 「設定」タブで `設定_医事部.xlsx` を登録。

以後、GitHub のファイルを更新してコミットすると自動でデプロイし直されます。

## 計算エンジンについて

- ページを開くと、裏で Pyodide（ブラウザで Python を動かす仕組み）と numpy・scipy・openpyxl を
  `https://cdn.jsdelivr.net/pyodide/v0.27.2/full/` から読み込みます。
- **初回は数十MBのダウンロードで1分ほど**かかります。2回目からはブラウザのキャッシュが効きます。
- 読み込みの状態は画面右上に出ます。失敗したときはその表示にマウスを乗せると理由が出ます。
- `engine-worker.js` と `py/*.py` はパスワードなしで取得できます（プログラムだけで、データは含みません）。

## 使い方

| タブ | やること |
|---|---|
| 希望・予定 | ペンを選んでマスをクリック／ドラッグ。`休` `有` などは固定、`?休` は希望、`×遅番` はその勤務にしない、`A/AS` はどちらか。自動で保存。Ctrl+Z で元に戻す。右に公休の数、下に日ごとの休み希望の人数。 |
| 特別日 | その日だけ必要人数を変えたい日（早出の日、レセ日、全体MT など） |
| 自動作成 | ボタンを押すと計算。結果の表・守れなかったルール・Excel ダウンロード |
| チェック | 手直しした勤務表や、今までの手作りの勤務表をアップロードしてルール違反を確認。その月の特別日も使う |
| 設定 | 設定.xlsx の登録・ダウンロード、登録内容と設定ミスの確認 |

## PC でコマンドとして使う

`public/py/shift_engine.py` と `solver_highs.py` を同じフォルダに置いて、

```
pip install openpyxl scipy        （OR-Tools があれば ortools でも可。あればそちらを使う）
python shift_engine.py new   設定.xlsx 2026-12
python shift_engine.py solve 設定.xlsx 予定_2026-12.xlsx
python shift_engine.py check 設定.xlsx 勤務表.xlsx シート名
```

## ローカルで動かす（開発用）

```
printf 'APP_PASSWORD=test\n' > .dev.vars
npm install
npx wrangler dev
```

## データの保存先（D1）

| テーブル | 中身 |
|---|---|
| files | 設定.xlsx（name=settings）、自動作成の結果（name=result:YYYY-MM）。どちらも base64 |
| plan | 予定表の書き込み（月・職員・日・値） |
| special | 特別日 |
| month_meta | 月ごとの公休数 |
