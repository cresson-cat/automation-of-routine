const fs = require('fs');
require('date-utils');

const webdriver = require('selenium-webdriver');
const by = require('selenium-webdriver').By;

const USER_ID = '316214961';
const PASSWORD = 'xxxxx'; // 引数で渡すか｡｡

(async function () {
    let driver = await new webdriver.Builder().forBrowser('chrome').build();
    try {
        /* ログイン画面 */
        await driver.get('https://cyncs-mobile.technopro.com/it/login/index');
        await driver.findElement(by.id('loginUserId')).sendKeys(USER_ID);
        await driver.findElement(by.id('loginPassword')).sendKeys(PASSWORD);
        await driver.findElement(by.id('btn-login')).click();

        /* メニュー画面 */
        await driver.findElement(by.id('btn-attendance')).click();

        /* 勤怠申請画面 */
        /* 残業時間の周辺に特定し易いidやclassが無い
         * また、cssセレクタはjqueryと違って、テキストノードの文字列にマッチングできない
         * .. 勤怠編集用のdivにidが振られているため、少し遠いが、そこから取得する */
        let elem = await driver.findElement(by.css('#attendanceEditDtl + div + div'));        // 勤怠編集用のdivの2番目の兄弟要素を取得
        elem = await elem.findElement(by.css('div > div > table > thead > tr:nth-child(1)')); // tableのtrまで取得
        elem = await elem.findElement(by.css('td + td > h3 > strong'));                       // "36協定"の次のtd配下から、strong要素を取得

        // 残業時間取得
        await elem.getText().then((overTime) => {
            fs.writeFileSync('./logs/app.log', `残業時間：${overTime}`);

            /* excelに出力する */

        });
    } catch (err) {
        // エラーを出力しておく
        fs.writeFileSync('./logs/err.log', err);
    } finally {
        await driver.quit();
    }
})();
