module.exports = {
    "env": {
        "es6": true,
        "node": true
    },
    "extends": "eslint:recommended",
    "parserOptions": {
        "sourceType": "module",
        "ecmaVersion": 2017 // async／awaitでエラーにならないようにする
    },
    "rules": {
        "no-console": "off", // consoleを使用可能にする
        "indent": [
            "error",
            4,
            { "ArrayExpression": 1 } // 配列の要素はインデントを付ける
        ],
        "linebreak-style": [
            "error",
            "unix"
        ],
        "quotes": [
            "error",
            "single"
        ],
        "semi": [
            "error",
            "always"
        ]
    }
};