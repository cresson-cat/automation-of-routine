/* ファイルIOや日付関連等 */
const fs = require('fs');         // ファイルIO
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
 * 週次の報告資料作成
 * @param {number} overTime 残業時間（H）
 * @param {object} conf 設定データ
 * @returns {string} 作成したファイル名
 */
module.exports = async (overTime, conf) => {
    try {
        writeMessage('報告資料の作成を開始します');
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
            if (err) writeMessage(`  過去ファイル ${oldName} は存在しないか、削除できません`);
        });

        /* 当月のファイルを作成する
         * Node.jsの流儀に従って、まず試してみて、エラーが発生した場合に対処する */
        let newName = `${nowDate.format('YYYY')}年${nowDate.format('M')}月度_見込み報告（東京第一支店）Ver1.4.xlsx`;

        //#region ファイルコピー
        /*
        try {
            fs.copyFileSync(TEMPLATE_PATH, newName, fs.constants.COPYFILE_EXCL);
        } catch (err) {
            writeMessage(`テンプレート ${newName} に追記します`);
        }
        */
        //#endregion
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
            // D4に現在月を書込み
            workSheet.getCell('D4').value = parseFloat(nowDate.format('M'));
            // F9に初週の残業時間を書込み
            workSheet.getCell('F9').value = conf.overTimeAt1stWk;

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
        /**
         * 本日が（報告資料上で）何週目かを取得する
         * @returns {number} 何週目かを示す数値
         */
        let weekNum = (() => {
            // 本日が何週目か単純に求める
            let val = Math.floor((nowDate.date() - nowDate.day() + 12) / 7);
            // 当月の1日が土日開始だった場合、報告資料のフォーマットに合わせて-1する
            let firstDay = moment([nowDate.year(), nowDate.month(), 1]);
            if (firstDay.day() === 6 || firstDay.day() === 0) val = val - 1;
            // 今日が日曜 or 月曜なら、1週前にする
            if (nowDate.day() === 0 || nowDate === 1) val = val - 1;
            return val;
        })();

        // 今週の残業実績を書き込む
        workSheet.getCell(TMP_CELLS[weekNum - 1].achievement).value = overTime;
        // 休暇見込みを取得する
        let prospect = 0;
        if (weekNum > 1) {
            // "2週目以降の休暇見込み"に限り報告用のシートから取得する
            prospect = workbook.getWorksheet('見込み報告フォーマット')
                .getCell(TMP_CELLS[weekNum - 2].vacation).value;
            // 現在週に書込み
            workSheet.getCell(TMP_CELLS[weekNum - 1].vacation).value = prospect;
        } else {
            prospect = 0; // 初週の場合0
        }

        // 現在週以降の残業時間は丸めておく
        overTime = parseInt(overTime);

        // 現在週以降のセルを更新する
        for (let i = weekNum; i < TMP_CELLS.length; i++) {
            // 残業見積りを更新
            workSheet.getCell(TMP_CELLS[i].forecast).value = overTime + (i - weekNum + 1) * conf.expectedOvertime;
            // 休暇見込をコピー
            workSheet.getCell(TMP_CELLS[i].vacation).value = prospect;
        }
        // ファイルを保存する
        await workbook.xlsx.writeFile(newName);
        // 作成したファイル名を返す
        return newName;
    } catch (err) {
        console.error(err); // エラーを出力しておく
        writeMessage(err);
    } finally {
        writeMessage('報告資料の作成が完了しました');
    }
};