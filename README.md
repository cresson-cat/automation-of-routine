# automation-of-routine

## 使い方

1. 設定ファイル `init.json` 及びテンプレ `temp.xlsx` を受け取る
2. テンプレを右記パスに格納する`./template/temp.xlsx`
3. テンプレ1シート目のA9セルに、自分の名前を記入する
4. `init.json`にユーザ／パスワードを記入する
5. 自OS用のchromedriverをダウンロードして上書き（インストール済のchromedriverは、MacOS 64bit版）
6. 以下のいずれかの方法で実行する
   1. nodeコマンドで実行

      ```bash
      # 実行
      node cron.js
      ```

   2. pm2で実行

      ```bash
      # 必要に応じてインストール
      npm i -g pm2

      # 実行
      pm2 start cron.js
      ```

## 補足事項

- "休暇見込＋実績"は「見込み報告フォーマット」シートに直接入力すること
- Windowsだと、chromedriverの格納場所を変えないと動かないらしい
