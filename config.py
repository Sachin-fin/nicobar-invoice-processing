import os
from pathlib import Path

# Base Paths
BASE_DIR = Path(__file__).resolve().parent
EXPENSES_DIR = Path(os.getenv("EXPENSES_DIR", r"D:\Google _invoice Processing"))
FALLBACK_DIR = Path(r"D:\Claude_Sales\Expenses")

REGISTER_PATH = EXPENSES_DIR / "1. Expense_Invoice_Register.xlsx"
VENDOR_CODES_PATH = EXPENSES_DIR / "3.Vendor_GL_Codes.xlsx"
DASHBOARD_HTML_PATH = EXPENSES_DIR / "Expense_Dashboard.html"
GMAIL_STATE_PATH = EXPENSES_DIR / "_gmail_scan_state.json"
TOOLS_DIR = EXPENSES_DIR / "_tools" if (EXPENSES_DIR / "_tools").exists() else FALLBACK_DIR / "_tools"
BUILD_DASHBOARD_SCRIPT = TOOLS_DIR / "build_dashboard.py"

# Known Nicobar GSTINs and States
NICOBAR_GSTINS = {
    "06AAFCN6567R1ZL": {"state": "Haryana", "code": "06", "name": "Nicobar Design Pvt Ltd - Haryana"},
    "07AAFCN6567R1ZJ": {"state": "Delhi", "code": "07", "name": "Nicobar Design Pvt Ltd - Delhi"},
    "07AAFCN6567R2ZI": {"state": "Delhi", "code": "07", "name": "Nicobar Design Pvt Ltd - Delhi Unit 2"},
    "08AAFCN6567R1ZH": {"state": "Rajasthan", "code": "08", "name": "Nicobar Design Pvt Ltd - Rajasthan"},
    "09AAFCN6567R1ZF": {"state": "Uttar Pradesh", "code": "09", "name": "Nicobar Design Pvt Ltd - UP"},
    "22AAFCN6567R1ZM": {"state": "Chhattisgarh", "code": "22", "name": "Nicobar Design Pvt Ltd - Chhattisgarh"},
    "24AAFCN6567R1ZI": {"state": "Gujarat", "code": "24", "name": "Nicobar Design Pvt Ltd - Gujarat"},
    "27AAFCN6567R1ZC": {"state": "Maharashtra", "code": "27", "name": "Nicobar Design Pvt Ltd - Maharashtra"},
    "29AAFCN6567R1ZD": {"state": "Karnataka", "code": "29", "name": "Nicobar Design Pvt Ltd - Karnataka"},
    "36AAFCN6567R1ZU": {"state": "Telangana", "code": "36", "name": "Nicobar Design Pvt Ltd - Telangana"},
}

# Indian State Codes Mapping
STATE_CODES = {
    "01": "Jammu & Kashmir",
    "02": "Himachal Pradesh",
    "03": "Punjab",
    "04": "Chandigarh",
    "05": "Uttarakhand",
    "06": "Haryana",
    "07": "Delhi",
    "08": "Rajasthan",
    "09": "Uttar Pradesh",
    "10": "Bihar",
    "11": "Sikkim",
    "12": "Arunachal Pradesh",
    "13": "Nagaland",
    "14": "Manipur",
    "15": "Mizoram",
    "16": "Tripura",
    "17": "Meghalaya",
    "18": "Assam",
    "19": "West Bengal",
    "20": "Jharkhand",
    "21": "Odisha",
    "22": "Chhattisgarh",
    "23": "Madhya Pradesh",
    "24": "Gujarat",
    "26": "Dadra & Nagar Haveli and Daman & Diu",
    "27": "Maharashtra",
    "28": "Andhra Pradesh (Old)",
    "29": "Karnataka",
    "30": "Goa",
    "31": "Lakshadweep",
    "32": "Kerala",
    "33": "Tamil Nadu",
    "34": "Puducherry",
    "35": "Andaman & Nicobar Islands",
    "36": "Telangana",
    "37": "Andhra Pradesh",
    "38": "Ladakh"
}

# Ignore list as per SKILL.md
IGNORED_FILES = {
    "cowork_invoice_to_dynamics365_process.pdf",
    "expense_dashboard.html",
    "_gmail_scan_state.json",
}
IGNORED_DIRS = {"_tools", "_docs", ".venv", "__pycache__"}
