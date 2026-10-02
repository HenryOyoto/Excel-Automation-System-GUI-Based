"""Application constants: the schema the app EXPECTS from input files."""

APP_NAME = "Excel Automation System"
APP_TAGLINE = "Automated Sales Data Processing & Reporting"
APP_VERSION = "1.0.0"

MASTER_FILENAME = "Master_Sales_Report.xlsx"
MIN_ROWS = 50
MIN_COLUMNS = 10

COL_DATE = "Date"
COL_INVOICE = "Invoice No."
COL_CUSTOMER = "Customer"
COL_PRODUCT = "Product"
COL_CATEGORY = "Category"
COL_QUANTITY = "Quantity"
COL_UNIT_PRICE = "Unit Price"
COL_DISCOUNT = "Discount"
COL_TAX_RATE = "Tax Rate"
COL_SALESPERSON = "Salesperson"
COL_GROSS = "Gross Amount"
COL_DISCOUNT_AMT = "Discount Amount"
COL_TAX_AMT = "Tax Amount"
COL_NET = "Net Sales"

REQUIRED_COLUMNS = [
    COL_DATE, COL_INVOICE, COL_CUSTOMER, COL_PRODUCT, COL_CATEGORY,
    COL_QUANTITY, COL_UNIT_PRICE, COL_DISCOUNT, COL_TAX_RATE, COL_SALESPERSON,
    COL_GROSS, COL_DISCOUNT_AMT, COL_TAX_AMT, COL_NET,
]
NUMERIC_COLUMNS = [COL_QUANTITY, COL_UNIT_PRICE, COL_DISCOUNT, COL_TAX_RATE]

# Extra columns the app adds to the consolidated sheet.
COL_SOURCE = "Source File"
COL_MONTH = "Month"


# --- Master workbook sheet names ---------------------------------------------
SHEET_DASHBOARD = "Dashboard"
SHEET_CONSOLIDATED = "Consolidated Data"
SHEET_PRODUCTS = "Product Analysis"
SHEET_SALESPERSONS = "Salesperson Analysis"
SHEET_MONTHLY = "Monthly Summary"

# Columns the user types in (the other four are calculated).
BASE_COLUMNS = REQUIRED_COLUMNS[:10]
TEXT_COLUMNS = [COL_INVOICE, COL_CUSTOMER, COL_PRODUCT, COL_CATEGORY, COL_SALESPERSON]

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]