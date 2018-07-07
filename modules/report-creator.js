/* ファイルIOや日付関連等 */
const fs = require('fs'); // ファイルIO
const Excel = require('exceljs'); // excel操作
const moment = require('moment'); // 日付操作
const koyomi = require('koyomi'); // 日本営業日計算
/* ログ用のモジュール */
const writeMessage = require('./logger');

/* 定数 */
const TEMPLATE_PATH = './template/temp.xlsx'; // テンプレのパス
const TMP_CELLS = [{
    'forecast': 'F9',    // 残業見積
    'achievement': 'G9', // 残業実績
    'vacation': 'H9'     // 休暇見込＋実績
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

/**
 * 見込報告資料作成
 * @param {number} overTime 残業時間（H）
 * @param {object} conf 設定データ
 */
module.exports = async function (overTime, conf) {
    try {
        writeMessage('見込み報告資料の作成を開始します');
        // 定義のチェック
        if (!conf)
            conf = {
                'overTimeAt1stWk': 2, // 初週の残業時間
                'baseTime': 7.5,      // 基本稼働時間/1D
                'expectedOvertime': 3 // 残業時間/1D
            };

        // 各種日時取得
        let nowDate = moment();    // 現在日付処理用
        let oldDate = moment();    // 古ファイル削除用
        let futureDate = moment(); // 翌月の稼働時間計算用

        // 2ヶ月前のファイルを削除する
        oldDate.add(-2, 'months');
        let oldName = `${oldDate.format('YYYY')}年${oldDate.format('M')}月度_見込み報告（東京第一支店）Ver1.4.xlsx`;
        fs.unlink(oldName, (err) => {
            if (err) writeMessage(`  過去ファイル ${oldName} は削除不要です`);
        });

        /* 当月のファイルを作成する
         * Node.jsの流儀に従って、まず試してみて、エラーが発生した場合に対処する */
        let newName = `${nowDate.format('YYYY')}年${nowDate.format('M')}月度_見込み報告（東京第一支店）Ver1.4.xlsx`;

        //#region OSによってfs.copyFileSync（node.js v8.9.4）に失敗するため、fs-extraを使う 
        /* ・ubuntu 16.04 LTS：失敗
         * ・macOS High Sierra 10.13.4：成功
         * .. いずれ解決したい */
        /*
        try {
            fs.copyFileSync(TEMPLATE_PATH, newName, fs.constants.COPYFILE_EXCL);
        } catch (err) {
            writeMessage(`テンプレート ${newName} に追記します`);
        }
        */
        //#endregion
        // ファイルコピー
        await require('fs-extra').copy(TEMPLATE_PATH, newName, {
            overwrite: false
        });

        // ファイルオープン
        let workbook = new Excel.Workbook();
        await workbook.xlsx.readFile(newName);
        /* exceljsで出力した場合、隣接したセルの書式が変わってしまったため、
         * 提出用でない、別シートの「data」に書込みする */
        let workSheet = workbook.getWorksheet('data');

        // （初回の場合）D4に現在の月を書き込み ※ D4が現在月と異なる時、初回とみなす
        if (workSheet.getCell('D4').value !== parseFloat(nowDate.format('M'))) {
            writeMessage('  月初の書込みを開始します');
            // F9に初週の残業時間を書込み
            workSheet.getCell('F9').value = conf.overTimeAt1stWk;
            // D4に現在月を書込み
            workSheet.getCell('D4').value = parseFloat(nowDate.format('M'));

            //#region GASで稼働時間の取得APIを作成中｡｡可能であれば、そちらも試してみたい
            /* https://script.google.com/macros/s/AKfycbyemymH0VeIAqlInNQmuZG1tMGJ0q6zpGbLQJ19ZtKWtMgXz9v3/exec
             * ?tgtMonth=201806&baseTime=6 */
            //#endregion

            // 当月の基本稼働時間を書込み
            workSheet.getCell('E9').value = koyomi.biz(nowDate.format('YYYYMM')) * conf.baseTime;
            // 次月の基本稼働時間を書込み
            workSheet.getCell('U9').value = koyomi.biz(futureDate.add(1, 'months').format('YYYYMM')) * conf.baseTime;
        }

        // 現在が何週目か判定し、書き込み先を特定する
        writeMessage('  現在週の書込みを開始します');
        let weekNum = Math.floor((nowDate.date() - nowDate.day() + 12) / 7);
        // 今日が日曜なら-1しておく
        if (nowDate.day() === 0) weekNum = weekNum - 1;

        // 今週の残業実績を書き込む
        workSheet.getCell(TMP_CELLS[weekNum - 1].achievement).value = overTime;
        // 今週の休暇見込みを取得する
        let prospect = workSheet.getCell(TMP_CELLS[weekNum - 1].vacation).value;

        // 現在週以降のセルを更新する
        for (let i = weekNum; i < TMP_CELLS.length; i++) {
            // 残業見積りを更新
            workSheet.getCell(TMP_CELLS[i].forecast).value = overTime + (i - weekNum + 1) * conf.expectedOvertime;
            // 休暇見込をコピー
            workSheet.getCell(TMP_CELLS[i].vacation).value = prospect;
        }
        // ファイルを保存する
        await workbook.xlsx.writeFile(newName);
    } catch (err) {
        console.error(err); // エラーを出力しておく
        writeMessage(err);
    } finally {
        writeMessage('見込み報告資料の作成が完了しました');
    }
};