/* I/Oや日付関連 */
const fs = require('fs'); // ファイルIO
const moment = require('moment'); // 日付操作

/**
 * ログ出力
 * @param {string} message
 */
module.exports = function (message) {
    // （node.jsの流儀に従い）フォルダの存在チェックはしない
    fs.mkdir('./logs/', (err) => {
        if (err && err.code === 'EEXIST') { // フォルダが存在する時のエラー
            fs.appendFile('./logs/app.log', moment().format('YYYY/MM/DD HH:mm:ss') + ' ' + message + '\n', (err) => {
                if (err) console.log(err);
            });
        }
    });
};