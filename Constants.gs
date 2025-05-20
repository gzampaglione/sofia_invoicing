// Constants.gs
// This file contains all global constants and configuration settings for the Google Workspace Add-on.

// === CONFIGURATION ===
const API_KEY = PropertiesService.getScriptProperties().getProperty('GL_API_KEY');
const SHEET_ID = '1qlEw5k5K-Tqg0joxXeiKtpgr6x8BzclobI-E-_mORhY'; // Main Spreadsheet ID
const ITEM_LOOKUP_SHEET_NAME = "Item Lookup";
const INVOICE_TEMPLATE_SHEET_NAME = "INVOICE_TEMPLATE";
const KITCHEN_SHEET_TEMPLATE_NAME = "KITCHEN_SHEET_TEMPLATE";

const FALLBACK_CUSTOM_ITEM_SKU = "CUSTOM_SKU"; // Ensure this SKU exists in your Item Lookup sheet

// Client matching rules, sorted by rule length (descending) for more specific matches first.
const CLIENT_RULES_LOOKUP = [
  { rule: "wharton.upenn.edu", clientName: "Penn Wharton" },
  { rule: "law.upenn.edu", clientName: "Penn Law"},
  { rule: "upenn.edu", clientName: "Penn" },
  { rule: "pennmedicine.upenn.edu", clientName: "Penn Medicine" }
].sort((a, b) => b.rule.length - a.rule.length);

// Delivery fee and utensil cost constants
const BASE_DELIVERY_FEE = 25.00;
const AFTER_4PM_DELIVERY_FEE = 40.00; // Example, adjust as needed for after-hours
const DELIVERY_FEE_CUTOFF_HOUR = 16; // 4 PM (16:00 in 24-hour format)
const COST_PER_UTENSIL_SET = 0.25; // Example cost per utensil set

// === INVOICE TEMPLATE CELL MAPPING CONSTANTS ===
// These map to specific cells in your INVOICE_TEMPLATE sheet for data population.
const ORDER_NUM_CELL = "D7";
const CUSTOMER_NAME_CELL = "B12"; // This will be the invoice "bill to" name
const ADDRESS_LINE_1_CELL = "B13";
const ADDRESS_LINE_2_CELL = "B14";
const CITY_STATE_ZIP_CELL = "B15";
const DELIVERY_DATE_CELL_INVOICE = "E15";
const DELIVERY_TIME_CELL_INVOICE = "G15";

// Item details start row and columns in the invoice template
const ITEM_START_ROW_INVOICE = 19;
const ITEM_DESCRIPTION_COL_INVOICE = "B";
const ITEM_QTY_COL_INVOICE = "E";
const ITEM_UNIT_PRICE_COL_INVOICE = "F";
const ITEM_TOTAL_PRICE_COL_INVOICE = "G";

// === KITCHEN SHEET TEMPLATE CELL MAPPING CONSTANTS ===
// These map to specific cells/columns in your KITCHEN_SHEET_TEMPLATE sheet.
const KITCHEN_CUSTOMER_PHONE_CELL = "B1"; // Cell for customer name and phone
const KITCHEN_DELIVERY_DATE_CELL = "C9";
const KITCHEN_DELIVERY_TIME_CELL = "F9";

// Item details start row and columns in the kitchen sheet template
const KITCHEN_ITEM_START_ROW = 12;
const KITCHEN_QTY_COL = "B";
const KITCHEN_SIZE_COL = "C";
const KITCHEN_ITEM_NAME_COL = "D";
const KITCHEN_FILLING_COL = "E";
const KITCHEN_NOTES_COL = "F";