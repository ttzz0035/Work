import js from "@eslint/js";
import globals from "globals";

export default [
  {
    // 生成物と依存ライブラリは解析対象から除外する。
    ignores: [
      "node_modules/**",
      "coverage/**",
      "docs/**",
    ],
  },
  {
    // FUXAスクリプトを静的解析する。
    files: ["targets/**/*.js"],

    languageOptions: {
      ecmaVersion: "latest",

      // トップレベルawaitを含むFUXAスクリプトを解析する。
      sourceType: "module",

      // FUXA実行環境が提供するAPIを宣言する。
      globals: {
        $getTag: "readonly",
        $setTag: "readonly",
        $getTagId: "readonly",
        axios: "readonly",
        ...globals.es2021,
      },
    },

    // ESLint標準の推奨ルールを使用する。
    rules: {
      ...js.configs.recommended.rules,

      // 未使用の関数引数・変数を検出する。
      "no-unused-vars": [
        "error",
        {
          args: "all",
          vars: "all",
          caughtErrors: "none",
        },
      ],

      // 到達不能コードを検出する。
      "no-unreachable": "error",

      // 再代入しない変数はconstを推奨する。
      "prefer-const": "warn",

      // 意図しない定数条件を検出する。
      "no-constant-condition": "error",
    },
  },
];