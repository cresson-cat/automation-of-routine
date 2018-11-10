/* Selenium */
const webdriver = require('selenium-webdriver'); // webdriver
const by = require('selenium-webdriver').By; // By
/* ログ用のモジュール */
const writeMessage = require('./logger');

/**
 * 残業時間取得
 * @param {string} userId ユーザID
 * @param {string} password パスワード
 * @returns {number} 残業時間
 */
module.exports = async (userId, password) => {
    writeMessage('-----');
    writeMessage('Webサイトへのアクセスを開始します');

    let driver = await new webdriver.Builder().forBrowser('chrome').build();
    try {
        // webdriverで、要素が生成されるまで一律10秒待機
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
         * cssセレクタは（jqueryと違って）テキストノードの文字列にマッチングできない
         * .. xpath形式で、"36協定"の隣のセル（td）を取得する */
        let elem = await driver.findElement(by.xpath('//small[text()="36協定"]/../../following-sibling::td'));
        elem = await elem.findElement(by.css('h3 > strong')); // td下の、残業時間取得

        // 残業時間取得
        let overTime = await elem.getText();
        writeMessage(`  残業時間：${parseFloat(overTime)}H`);

        // 残業時間を返却する（数値に変換済）
        return parseFloat(overTime);
    } catch (err) {
        console.error(err); // エラーを出力しておく
        writeMessage(err);
    } finally {
        driver.quit();
        writeMessage('Webサイトへのアクセスが完了しました');
    }
};