// V13
// === CONFIGURATION ===
const API_KEY = PropertiesService.getScriptProperties().getProperty('GL_API_KEY');
const SHEET_ID = '1qlEw5k5K-Tqg0joxXeiKtpgr6x8BzclobI-E-_mORhY'; // Main Spreadsheet ID
const ITEM_LOOKUP_SHEET_NAME = "Item Lookup";
const INVOICE_TEMPLATE_SHEET_NAME = "INVOICE_TEMPLATE";
const KITCHEN_SHEET_TEMPLATE_NAME = "KITCHEN_SHEET_TEMPLATE";

const FALLBACK_CUSTOM_ITEM_SKU = "CUSTOM_SKU"; // Ensure this SKU exists in your Item Lookup sheet

const CLIENT_RULES_LOOKUP = [
  { rule: "@wharton.upenn.edu", clientName: "University of Pennsylvania - Wharton" },
  { rule: "law.upenn.edu", clientName: "University of Pennsylvania - Law School"},
  { rule: "@upenn.edu", clientName: "University of Pennsylvania" },
  { rule: "@pennmedicine.upenn.edu", clientName: "Penn Medicine" }
].sort((a, b) => b.rule.length - a.rule.length); // Sort by rule length, descending

const BASE_DELIVERY_FEE = 15.00;
const AFTER_4PM_DELIVERY_FEE = 25.00; // Example, adjust as needed
const DELIVERY_FEE_CUTOFF_HOUR = 16; // 4 PM (4 PM is 16:00 in 24-hour format)
const COST_PER_UTENSIL_SET = 0.50; // Example cost per utensil set

// === INVOICE TEMPLATE CELL MAPPING CONSTANTS ===
const ORDER_NUM_CELL = "D7";
const CUSTOMER_NAME_CELL = "B12"; // This will be the invoice "bill to" name
const ADDRESS_LINE_1_CELL = "B13";
const ADDRESS_LINE_2_CELL = "B14";
const CITY_STATE_ZIP_CELL = "B15";
const DELIVERY_DATE_CELL_INVOICE = "E15";
const DELIVERY_TIME_CELL_INVOICE = "G15";

const ITEM_START_ROW_INVOICE = 19;
const ITEM_DESCRIPTION_COL_INVOICE = "B";
const ITEM_QTY_COL_INVOICE = "E";
const ITEM_UNIT_PRICE_COL_INVOICE = "F";
const ITEM_TOTAL_PRICE_COL_INVOICE = "G";

// === KITCHEN SHEET TEMPLATE CELL MAPPING CONSTANTS ===
const KITCHEN_CUSTOMER_PHONE_CELL = "B1";
const KITCHEN_DELIVERY_DATE_CELL = "C9";
const KITCHEN_DELIVERY_TIME_CELL = "F9";

const KITCHEN_ITEM_START_ROW = 12;
const KITCHEN_QTY_COL = "B";
const KITCHEN_SIZE_COL = "C";
const KITCHEN_ITEM_NAME_COL = "D";
const KITCHEN_FILLING_COL = "E";
const KITCHEN_NOTES_COL = "F";