/* import */
const fs = require('fs');     // ファイルIO
const xlsx = require('xlsx'); // xlsx操作
require('date-utils');        // Dateの拡張
const webdriver = require('selenium-webdriver'); // Selenium
const by = require('selenium-webdriver').By;     // By

/* 定数 */
const TEMPLATE_PATH = './template/temp.xlsx'; // テンプレのパス

/* 予期せぬエラーをcatchする */
process.on('uncaughtException', function (err) {
    console.log(err);
});

let userId = process.argv[2]; // 引数.. ユーザ名
let password = process.argv[3]; // 引数.. パスワード
let nowDate = new Date(); // 現在日時

(async function () {
    let driver = await new webdriver.Builder().forBrowser('chrome').build();
    try {
        // webdriverで、要素が生成されるまで一律5秒待機
        await driver.manage().setTimeouts({
            implicit: 5000,
            pageLoad: 5000,
            script: 5000
        });

        /* ログイン画面 */
        await driver.get('https://cyncs-mobile.technopro.com/it/login/index');
        await driver.findElement(by.id('loginUserId')).sendKeys(userId);
        await driver.findElement(by.id('loginPassword')).sendKeys(password);
        await driver.findElement(by.id('btn-login')).click();

        /* メニュー画面 */
        await driver.findElement(by.id('btn-attendance')).click();

        /* 勤怠申請画面 */
        /* 残業時間の周辺に特定し易いidやclassが無い
         * また、cssセレクタはjqueryと違って、テキストノードの文字列にマッチングできない
         * .. 勤怠編集用のdivにidが振られているため、少し遠いが、そこから取得する */
        let elem = await driver.findElement(by.css('#attendanceEditDtl + div + div')); // 勤怠編集用のdivの2番目の兄弟要素を取得
        elem = await elem.findElement(by.css('div > div > table > thead > tr:nth-child(1)')); // tableのtrまで取得
        elem = await elem.findElement(by.css('td + td > h3 > strong')); // "36協定"の次のtd配下から、strong要素を取得

        // 残業時間取得
        let overTime = await elem.getText();
        writeMessage('-----');
        writeMessage(`残業時間：${parseInt(overTime)}H`);

        /* excelに書き込む */
        // todo：2ヶ月前のファイルを削除する
        let oldDate = new Date();
        oldDate.setMonth(nowDate.getMonth() - 2);
        let oldName = `${oldDate.toFormat('YYYY')}年${oldDate.toFormat('M')}月度_見込み報告（東京第一支店）Ver1.4.xlsx`;
        fs.unlink(oldName, (err) => {
            if (err) writeMessage(`${oldName} は未削除です`);
        });

        /* 当月のファイルを作成する
         * Node.jsの流儀に従って、まず試してみて、エラーが発生した場合に対処する */
        let newName = `${nowDate.toFormat('YYYY')}年${nowDate.toFormat('M')}月度_見込み報告（東京第一支店）Ver1.4.xlsx`;
        try {
            fs.copyFileSync(TEMPLATE_PATH, newName, fs.constants.COPYFILE_EXCL);
        } catch (err) {
            // エラーを出力しておく
            writeMessage(`テンプレート ${newName} に追記します`);
        }

        // ファイルオープン
        let workbook = xlsx.readFile(newName);
        let worksheet = workbook.Sheets['見込み報告フォーマット'];

        // （初回の場合）D4に現在の月を書き込み
        // ※ D4が現在月と異なる時、初回とみなす
        if (worksheet['D4'].v !== parseInt(nowDate.toFormat('M'))) {
            writeMessage('月初の書込みを開始します');
            // （初回の場合）E9に基本稼働時間を書き込み
            // （初回の場合）U9に次月の基本稼働時間を書き込み
        }

        // 現在が何週目か判定し、書き込み先のセルを取得する

        // 残業時間を書き込み

        // ファイルクローズ

    } catch (err) {
        // エラーを出力しておく
        writeMessage(err);
    } finally {
        driver.quit();
    }
})();

/**
 * ログ出力
 * @param {string} message 
 */
function writeMessage(message) {
    // （node.jsの流儀に従い）フォルダの存在チェックはしない
    fs.mkdir('./logs/', (err) => {
        if (err && err.code === 'EEXIST') {
            fs.appendFile('./logs/app.log', new Date().toFormat('YYYY/MM/DD HH24:MI:SS') + ' ' + message + '\n', (err) => {
                if (err) console.log(err);
            });
        }
    });
}