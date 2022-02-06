module.exports = {
    "extends": ["prettier"],
    "env": "es2020",
    "rules": {
        "no-extend-native": "off",
        "no-unused-vars": ["error", { "vars": "local", "args": "after-used", "ignoreRestSiblings": false }],
        "no-var": "off",
        "require-jsdoc": "off",
        "valid-jsdoc": "off"
    }
};