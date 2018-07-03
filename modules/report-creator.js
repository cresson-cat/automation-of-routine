/* I/Oや日付関連 */
const fs = require('fs');         // ファイルIO
const Excel = require('exceljs'); // excel操作
const moment = require('moment'); // 日付操作
const koyomi = require('koyomi'); // 日本営業日計算

/* 2018/06時点で、OSによってfs.fileCopy（node.js v8.9.4）の成否が異なる 
 * ・ubuntu 16.04 LTS：失敗
 * ・macOS High Sierra 10.13.4：成功
 * ※ 一貫性のため、fsは全てfs-extraに変更するかも。というか、ubuntuはfsが入ってないかも.. */
const _copy = require('fs-extra').copy;

/* 独自のモジュール */
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
 * @param {string} overTime 残業時間
 */
module.exports = async function (overTime) {
    writeMessage('見込み報告資料の作成を開始します');

    // 現在日時
    let nowDate = moment();

    /* excelに書き込む */
    // 2ヶ月前のファイルを削除する
    let oldDate = moment().add(-2, 'months');
    let oldName = `${oldDate.format('YYYY')}年${oldDate.format('M')}月度_見込み報告（東京第一支店）Ver1.4.xlsx`;
    fs.unlink(oldName, (err) => {
        if (err) writeMessage(`過去ファイル ${oldName} は削除不要です`);
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
    await _copy(TEMPLATE_PATH, newName, {
        overwrite: false
    });

    // ファイルオープン
    let workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(newName);
    /* exceljsで出力した場合、隣接したセルの書式が変わってしまったため、
     * 「見込み報告フォーマット」ではなく、別シートの「data」に書込みする */
    let workSheet = workbook.getWorksheet('data');

    // （初回の場合）D4に現在の月を書き込み ※ D4が現在月と異なる時、初回とみなす
    if (workSheet.getCell('D4').value !== parseFloat(nowDate.format('M'))) {
        writeMessage('月初の書込みを開始します');
        // D4に現在月を書込み
        workSheet.getCell('D4').value = parseFloat(nowDate.format('M'));

        /* GASで基本稼働時間の取得APIを作成中｡｡可能であれば、そちらも試したい
         * https://script.google.com/macros/s/AKfycbyemymH0VeIAqlInNQmuZG1tMGJ0q6zpGbLQJ19ZtKWtMgXz9v3/exec
         * ?tgtMonth=201806&baseTime=6 */

        // 当月の基本稼働時間を書込み
        workSheet.getCell('E9').value = koyomi.biz(nowDate.format('YYYYMM')) * 7.5;
        // 次月の基本稼働時間を書込み
        workSheet.getCell('U9').value = koyomi.biz(nowDate.add(1, 'M').format('YYYYMM')) * 7.5;
    }

    // 現在が何週目か判定し、書き込み先を特定する
    writeMessage('現在週の書込みを開始します');
    let weekNum = Math.floor((nowDate.date() - nowDate.day() + 12) / 7);
    // 今日が日曜なら-1しておく
    if (nowDate.day() === 0) weekNum = weekNum - 1;

    // 今週の残業実績を書き込む
    workSheet.getCell(TMP_CELLS[weekNum - 1].achievement).value = parseFloat(overTime);

    // 今週の休暇見込みを取得する
    let prospect = workSheet.getCell(TMP_CELLS[weekNum - 1].vacation).value;

    // 現在週以降のセルを更新する
    for (let i = weekNum; i < TMP_CELLS.length; i++) {
        // 残業見積りを更新
        workSheet.getCell(TMP_CELLS[i].forecast).value = parseFloat(overTime) + (i - weekNum + 1) * 3;
        // 休暇見込をコピー
        workSheet.getCell(TMP_CELLS[i].vacation).value = prospect;
    }
    // ファイルを保存する
    await workbook.xlsx.writeFile(newName);
    writeMessage('見込み報告資料の作成が完了しました');
};