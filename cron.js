const cron = require('node-cron');
const main = require('./app');

// ++ DEBUG ++
// main();

// 定期実行（日曜日の朝6時）
cron.schedule('0 0 6 * * 0', () => main());
