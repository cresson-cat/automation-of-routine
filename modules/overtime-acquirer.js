/* Selenium */
const webdriver = require('selenium-webdriver'); // webdriver
const by = require('selenium-webdriver').By; // By
/* 独自のモジュール */
const writeMessage = require('./logger');

/**
 * 残業時間取得
 * @param {string} userId ユーザID
 * @param {string} password パスワード
 * @returns {string} 残業時間
 */
module.exports = async function(userId, password) {
    writeMessage('-----');
    writeMessage('Webサイトへのアクセスを開始します');
    
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
        writeMessage(`残業時間：${parseFloat(overTime)}H`);
        
        // 残業時間を返却する
        return overTime;
    } catch (err) {
        writeMessage(err); // エラー内容を出力
    } finally {
        driver.quit();
        writeMessage('Webサイトへのアクセスが完了しました');
    }
};