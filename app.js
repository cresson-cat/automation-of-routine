const getOverTime = require('./modules/overtime-acquirer');
const createReport = require('./modules/report-creator');

/* 予期せぬエラーをcatchする */
process.on('uncaughtException', function (err) {
    console.log(err);
});

/* todo：init.jsonから初期データを取得する
 * 引数が無かった場合は、以下変数にセットする */

let userId = process.argv[2]; // 引数.. ユーザ名
let password = process.argv[3]; // 引数.. パスワード

// ユーザIDやパスワードが未設定の場合、処理終了
if ((userId === null || userId === undefined) || (password === null || password === undefined))
    process.exit(0);

(async function () {
    // webサイトにアクセス
    let overTime = await getOverTime(userId, password);
    // 見込み資料を作成
    createReport(overTime);
})();