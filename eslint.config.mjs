import js from "@eslint/js";
import globals from "globals";
import { defineConfig } from "eslint/config";

const gasGlobals = {
  SpreadsheetApp: "readonly",
  Logger: "readonly",
  Session: "readonly",
  Utilities: "readonly",
  DriveApp: "readonly",
  MailApp: "readonly",
  GmailApp: "readonly",
  Browser: "readonly",
  PropertiesService: "readonly",
  ScriptApp: "readonly",
  HtmlService: "readonly",
  ContentService: "readonly",
  CalendarApp: "readonly",
  FormApp: "readonly",
  console: "readonly",
};

const projectGlobals = {
  SHEET_NAMES: "readonly",
  LAYOUT: "readonly",
  getRenamedVendor: "readonly",
  columnToLetter: "readonly",        // ← add
  isRecurringDateMatch: "readonly",  // ← add (used here + dashboard.js)
  updateDashboard: "readonly",
  populateNextMonth: "readonly",
};

export default defineConfig([
  {
    files: ["**/*.{js,mjs,cjs}"],
    plugins: { js },
    extends: ["js/recommended"],
    languageOptions: {
      globals: {
        ...globals.node,
        ...gasGlobals,
        ...projectGlobals,
      },
      sourceType: "module",
      ecmaVersion: 2020,
    },
    rules: {
      "no-unused-vars": "warn",
      "no-undef": "error",
      "no-console": "off",
      "prefer-const": "error",
      "no-var": "error",
    },
  },
  {
    files: ["utils.js"],
    rules: {
      "no-unused-vars": ["warn", {
        varsIgnorePattern: "^(SHEET_NAMES|LAYOUT|columnToLetter|isRecurringDateMatch)$"
      }],
    },
  },
]);