/* I/Oや日付関連 */
const fs = require('fs'); // ファイルIO
const xlsx = require('xlsx'); // xlsx操作
const moment = require('moment'); // 日付操作
const koyomi = require('koyomi'); // 日本営業日計算
/* 2018/06時点で、OSによってfs.fileCopy（node.js v8.9.4）の成否が異なる 
 * ・ubuntu 16.04 LTS：失敗
 * ・macOS High Sierra 10.13.4：成功
 * ※ 一貫性のため、fsは全てfs-extraに変更するかも。というか、ubuntuはfsが入ってないかも.. */
const _copy = require('fs-extra').copy;
/* Selenium */
const webdriver = require('selenium-webdriver'); // webdriver
const by = require('selenium-webdriver').By; // By

/* 定数 */
const TEMPLATE_PATH = './template/temp.xlsx'; // テンプレのパス
const TMP_CELLS = [{
    'forecast': 'F9', // 残業見積
    'achievement': 'G9', // 残業実績
    'vacation': 'H9' // 休暇見込＋実績
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

/* 予期せぬエラーをcatchする */
process.on('uncaughtException', function (err) {
    console.log(err);
});

/* todo：init.jsonから初期データを取得する
 * 引数が無かった場合は、以下変数にセットする */

let userId = process.argv[2]; // 引数.. ユーザ名
let password = process.argv[3]; // 引数.. パスワード
let nowDate = moment(); // 現在日時

// ユーザIDやパスワードが未設定の場合、処理終了
if ((userId === null || userId === undefined) || (password === null || password === undefined))
    process.exit(0);

(async function () {
    let driver = await new webdriver.Builder().forBrowser('chrome').build();
    try {
        // webdriverで、要素が生成されるまで一律5秒待機
        await driver.manage().setTimeouts({
            implicit: 10000,
            pageLoad: 10000,
            script: 10000
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
        // todo：xpath形式でテキストノードを検索できる。追ってxpathに変更する
        let elem = await driver.findElement(by.css('#attendanceEditDtl + div + div')); // 勤怠編集用のdivの2番目の兄弟要素を取得
        elem = await elem.findElement(by.css('div > div > table > thead > tr:nth-child(1)')); // tableのtrまで取得
        elem = await elem.findElement(by.css('td + td > h3 > strong')); // "36協定"の次のtd配下から、strong要素を取得

        // 残業時間取得
        let overTime = await elem.getText();
        writeMessage('-----');
        writeMessage(`残業時間：${parseFloat(overTime)}H`);

        /* excelに書き込む */
        // 2ヶ月前のファイルを削除する
        let oldDate = moment().add(-2, 'months');
        let oldName = `${oldDate.format('YYYY')}年${oldDate.format('M')}月度_見込み報告（東京第一支店）Ver1.4.xlsx`;
        fs.unlink(oldName, (err) => {
            if (err) writeMessage(`${oldName} は未削除です`);
        });

        /* 当月のファイルを作成する
         * Node.jsの流儀に従って、まず試してみて、エラーが発生した場合に対処する */
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
        await _copy(TEMPLATE_PATH, newName, false);

        // ファイルオープン
        let workbook = xlsx.readFile(newName);
        let worksheet = workbook.Sheets['見込み報告フォーマット'];

        // （初回の場合）D4に現在の月を書き込み
        // ※ D4が現在月と異なる時、初回とみなす
        if (worksheet['D4'].v !== parseFloat(nowDate.format('M'))) {
            writeMessage('月初の書込みを開始します');
            // D4に現在月を書込み
            worksheet['D4'].v = parseFloat(nowDate.format('M'));

            /* GASで基本稼働時間の取得APIを作成中｡｡可能であれば、そちらも試したい
             * https://script.google.com/macros/s/AKfycbyemymH0VeIAqlInNQmuZG1tMGJ0q6zpGbLQJ19ZtKWtMgXz9v3/exec
             * ?tgtMonth=201806&baseTime=6 */

            // 当月の基本稼働時間を書込み
            worksheet['E9'].v = koyomi.biz(nowDate.format('YYYYMM')) * 7.5;
            // 次月の基本稼働時間を書込み
            worksheet['U9'].v = koyomi.biz(nowDate.add(1, 'M').format('YYYYMM')) * 7.5;
        }

        // 現在が何週目か判定し、書き込み先を特定する
        let weekNum = Math.floor((nowDate.date() - nowDate.day() + 12) / 7);

        // 今週の残業実績を書き込む
        worksheet[TMP_CELLS[weekNum - 1].achievement].v = parseFloat(overTime);

        // 今週の休暇見込みを取得する
        let prospect = worksheet[TMP_CELLS[weekNum - 1].vacation].v;

        // 現在週以降のセルを更新する
        for (let i = weekNum; i < TMP_CELLS.length; i++) {
            // 残業見積りを更新
            worksheet[TMP_CELLS[i].forecast].v = parseFloat(overTime) + (i - weekNum + 1) * 3;
            // 休暇見込をコピー
            worksheet[TMP_CELLS[i].vacation].v = prospect;
        }

        // 範囲の更新
        worksheet['!ref'] = 'A1:W41';
        workbook.Sheets['見込み報告フォーマット'] = worksheet;
        // ファイルを保存する
        xlsx.writeFile(workbook, newName);

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
        if (err && err.code === 'EEXIST') { // フォルダが存在する時のエラー
            fs.appendFile('./logs/app.log', moment().format('YYYY/MM/DD HH24:MI:SS') + ' ' + message + '\n', (err) => {
                if (err) console.log(err);
            });
        }
    });
}