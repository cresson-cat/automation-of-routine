# automation-of-routine

## 使い方

1. テンプレ1シート目のA9セルに、自分の名前を記入する
2. テンプレを右記パスに格納する「template/temp.xlsx」
3. 必要に応じて「init.json」を修正する
4. 自OS用のchromedriverをダウンロードして上書き（インストール済のchromedriverは、MacOS 64bit版）
5. 右記の引数を付与してキックする（引数1：ユーザ名, 引数2：パスワード）

　 　※ なお「休暇見込＋実績」は、本ツール直下に作成されたファイルの提出用シートに直接入力すること

### memo

1. fs -> fs-extra へ変更するかも
2. （念のため）ESLintの設定は全体的に見直す
