const fs = require('fs');
const getOverTime = require('./modules/overtime-acquirer');
const createReport = require('./modules/report-creator');

/* 予期せぬエラーをcatchする */
process.on('uncaughtException', function (err) {
    console.log(err);
});

const userId = process.argv[2];   // 引数.. ユーザ名
const password = process.argv[3]; // 引数.. パスワード
// init.jsonから初期データを取得する（基本稼働時間や、週次の想定残業時間等）
const conf = JSON.parse(fs.readFileSync('./init.json', 'utf8')); // 初期設定ファイル取得

// ユーザIDやパスワードが未設定の場合、処理終了
if ((userId === null || userId === undefined) || (password === null || password === undefined))
    process.exit(0);

// メイン処理開始
(async function () {
    // webサイトにアクセス
    let overTime = await getOverTime(userId, password);
    if (overTime === undefined || overTime === null) return;
    // 見込み資料を作成
    createReport(overTime, conf);
})();