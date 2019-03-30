const fs = require('fs');
const path = require('path');
const getOverTime = require('./lib/overtime-capturer');
const createReport = require('./lib/report-creator');

/* 予期せぬエラーをcatchする */
process.on('uncaughtException', function (err) {
    console.log(err);
});

/* メイン処理 */
const main = function () {

    // init.jsonから初期データを取得する（基本稼働時間や、週次の想定残業時間等）
    const conf = JSON.parse(fs.readFileSync('./init.json', 'utf8'));

    // ユーザIDやパスワードが未設定の場合、処理終了
    const { userId, password, isSaved } = conf;
    if ((userId === null || userId === undefined) || (password === null || password === undefined))
        process.exit(0);

    (async () => {
        // webサイトにアクセス
        let overTime = await getOverTime(userId, password);
        if (overTime === undefined || overTime === null) return;
        // 週次の報告資料を作成
        let fileName = await createReport(overTime, conf);
        if (!fileName) return;
        //#region OSによってfs.copyFileSync（node.js v8.9.4）に失敗したため、fs-extraを使う 
        /* ・ubuntu 16.04 LTS：失敗
         * ・macOS High Sierra 10.13.4：成功
         * .. いずれ解決したい */
        //#endregion
        // 提出用のフォルダにコピーする
        try {
            if (isSaved) await require('fs-extra').copy(fileName, path.join(conf.outputDir, fileName));
        } catch (err) {
            console.log(err);
        }
    })();
};

module.exports = main;