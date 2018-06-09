<<<<<<< HEAD
/* I/Oや日付関連 */
const fs = require('fs');     // ファイルIO
const xlsx = require('xlsx'); // xlsx操作
const moment = require('moment'); // 日付操作
const koyomi = require('koyomi'); // 日本営業日計算
/* 2018/06時点で、OSによってfs.fileCopy（node.js v8.9.4）の成否が異なる 
 * ・ubuntu 16.04 LTS：失敗
 * ・macOS High Sierra 10.13.4：成功
 * ※ 一貫性のため、fsは全てfs-extraに変更するかも */
const _copy = require('fs-extra').copy;
/* Selenium */
const webdriver = require('selenium-webdriver'); // webdriver
const by = require('selenium-webdriver').By; // By

/* 定数 */
const TEMPLATE_PATH = './template/temp.xlsx'; // テンプレのパス
const tempCells = [{
    'forecast': 'F9',
    'achievement': 'G9',
    'vacation': 'H9'
}, {
    'forecast': 'I9',
    'achievement': 'J9',
    'vacation': 'K9'
}, {
    'forecast': 'L9',
    'achievement': 'M9',
    'vacation': 'N9'
}, {
    'forecast': 'O9',
    'achievement': 'P9',
    'vacation': 'Q9'
}, {
    'forecast': 'R9',
    'achievement': 'S9',
    'vacation': 'T9'
}];
=======
/* import */
const fs = require('fs');     // ファイルIO
const xlsx = require('xlsx'); // xlsx操作
require('date-utils');        // Dateの拡張
const webdriver = require('selenium-webdriver'); // Selenium
const by = require('selenium-webdriver').By;     // By

/* 定数 */
const TEMPLATE_PATH = './template/temp.xlsx'; // テンプレのパス
>>>>>>> 6b250e347c7e8dd36825b5648f3483f0368cb0a5

/* 予期せぬエラーをcatchする */
process.on('uncaughtException', function (err) {
    console.log(err);
});

<<<<<<< HEAD
/* todo：init.jsonから初期データを取得する
 * 引数が無かった場合は、以下変数にセットする */

let userId = process.argv[2]; // 引数.. ユーザ名
let password = process.argv[3]; // 引数.. パスワード
let nowDate = moment(); // 現在日時

// ユーザIDやパスワードが未設定の場合、処理終了
if ((userId === null || userId === undefined) || (password === null || password === undefined))
    process.exit(0);
=======
let userId = process.argv[2]; // 引数.. ユーザ名
let password = process.argv[3]; // 引数.. パスワード
let nowDate = new Date(); // 現在日時
>>>>>>> 6b250e347c7e8dd36825b5648f3483f0368cb0a5

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
        // 2ヶ月前のファイルを削除する
<<<<<<< HEAD
        let oldDate = moment().add(-2, 'months');
        // oldDate.setMonth(nowDate.getMonth() - 2);
        let oldName = `${oldDate.format('YYYY')}年${oldDate.format('M')}月度_見込み報告（東京第一支店）Ver1.4.xlsx`;
=======
        let oldDate = new Date();
        oldDate.setMonth(nowDate.getMonth() - 2);
        let oldName = `${oldDate.toFormat('YYYY')}年${oldDate.toFormat('M')}月度_見込み報告（東京第一支店）Ver1.4.xlsx`;
>>>>>>> 6b250e347c7e8dd36825b5648f3483f0368cb0a5
        fs.unlink(oldName, (err) => {
            if (err) writeMessage(`${oldName} は未削除です`);
        });

        /* 当月のファイルを作成する
         * Node.jsの流儀に従って、まず試してみて、エラーが発生した場合に対処する */
<<<<<<< HEAD
        let newName = `${nowDate.format('YYYY')}年${nowDate.format('M')}月度_見込み報告（東京第一支店）Ver1.4.xlsx`;
        /*
        try {
            // linux環境だと基底のfs.copyFileSyncが失敗するためfs-extraを使用
            fs.copyFileSync(TEMPLATE_PATH, newName, fs.constants.COPYFILE_EXCL);
        } catch (err) {
            writeMessage(`テンプレート ${newName} に追記します`);
        }
        */
        // ファイルコピー
        await _copy(TEMPLATE_PATH, newName);
=======
        let newName = `${nowDate.toFormat('YYYY')}年${nowDate.toFormat('M')}月度_見込み報告（東京第一支店）Ver1.4.xlsx`;
        try {
            fs.copyFileSync(TEMPLATE_PATH, newName, fs.constants.COPYFILE_EXCL);
        } catch (err) {
            // エラーになった場合、既に当月のファイルが作成済
            writeMessage(`テンプレート ${newName} に追記します`);
        }
>>>>>>> 6b250e347c7e8dd36825b5648f3483f0368cb0a5

        // ファイルオープン
        let workbook = xlsx.readFile(newName);
        let worksheet = workbook.Sheets['見込み報告フォーマット'];

        // （初回の場合）D4に現在の月を書き込み
        // ※ D4が現在月と異なる時、初回とみなす
        if (worksheet['D4'].v !== parseInt(nowDate.format('M'))) {
            writeMessage('月初の書込みを開始します');
            // D4に現在月を書込み
            worksheet['D4'].v = parseInt(nowDate.format('M'));

            /* GASで基本稼働時間の取得APIを作成中｡｡そちらも検証
             * https://script.google.com/macros/s/AKfycbyemymH0VeIAqlInNQmuZG1tMGJ0q6zpGbLQJ19ZtKWtMgXz9v3/exec
             * ?tgtMonth=201806&baseTime=6 */

            // 当月の基本稼働時間を書込み
            worksheet['E9'].v = koyomi.biz(nowDate.format('YYYYMM')) * 7.5;
            // 次月の基本稼働時間を書込み
            worksheet['U9'].v = koyomi.biz(nowDate.add(1, 'M').format('YYYYMM')) * 7.5;
            // 現在の翌週から、残業見込を更新

        }

        // 現在が何週目か判定し、書き込み先を特定する
        // let thisWk = parseInt(nowDate.format('D'));

        // 現在週以降の残業見積を更新する

        // 残業実績を書き込み

        // 休暇見込+実績をコピーする

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
<<<<<<< HEAD
        if (err && err.code === 'EEXIST') { // フォルダが存在する時のエラー
            fs.appendFile('./logs/app.log', moment().format('YYYY/MM/DD HH24:MI:SS') + ' ' + message + '\n', (err) => {
=======
        if (err && err.code === 'EEXIST') {
            fs.appendFile('./logs/app.log', new Date().toFormat('YYYY/MM/DD HH24:MI:SS') + ' ' + message + '\n', (err) => {
>>>>>>> 6b250e347c7e8dd36825b5648f3483f0368cb0a5
                if (err) console.log(err);
            });
        }
    });
}