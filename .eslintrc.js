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
<<<<<<< HEAD
            4,
            { "ArrayExpression": 1 } // 配列の要素はインデントを付ける
=======
            4
>>>>>>> 6b250e347c7e8dd36825b5648f3483f0368cb0a5
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