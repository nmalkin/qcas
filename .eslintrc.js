module.exports = {
    "extends": ["prettier"],
    "rules": {
        "no-extend-native": "off",
        "no-unused-vars": ["error", { "vars": "local", "args": "after-used", "ignoreRestSiblings": false }],
        "no-var": "off",
        "require-jsdoc": "off",
        "valid-jsdoc": "off"
    }
};