/* I/Oや日付関連 */
const fs = require('fs'); // ファイルIO
const moment = require('moment'); // 日付操作
const { promisify } = require('util');

// promise化
const mkdir = promisify(fs.mkdir);
const appendFile = promisify(fs.appendFile);

/**
 * ログ出力
 * @param {string} message
 */
module.exports = (message) => {
    // （node.jsの流儀に従い）フォルダの存在チェックはしない
    mkdir('./logs/')
        .then(
            () => {
                // 初回のみ、mkdirはfulfilledになる
                appendFile('./logs/app.log', moment().format('YYYY/MM/DD HH:mm:ss') + ' ' + message + '\n')
            },
            (err) => {
                // フォルダが存在する時（2回目以降）
                if (err && err.code === 'EEXIST') {
                    appendFile('./logs/app.log', moment().format('YYYY/MM/DD HH:mm:ss') + ' ' + message + '\n');
                }
            })
        .catch((err) => {
            if (err) console.log(err);
        });
};