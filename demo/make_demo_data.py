"""Generate synthetic demo inputs for uar-tool.html (fictional people, example.com domain).

Run:  python demo/make_demo_data.py   -> writes 4 .xlsx files next to this script.
Each user is designed to light up one specific check; see comments.
"""
from datetime import date
from pathlib import Path
from openpyxl import Workbook

OUT = Path(__file__).parent

def book(path, sheets):
    wb = Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        for r in rows:
            ws.append(r)
    wb.save(OUT / path)

# (username, full name "Last, First", email local == username) — None email = no Oracle account
PEOPLE = [
    ("JSMITH",  "Smith, John"),     ("PKUMAR", "Kumar, Priya"),   ("ALEE",    "Lee, Andrew"),
    ("MGARCIA", "Garcia, Maria"),   ("RWONG",  "Wong, Raymond"),  ("SNAIR",   "Nair, Sunita"),
    ("DCHEN",   "Chen, David"),     ("KPATEL", "Patel, Kiran"),   ("TBROWN",  "Brown, Tara"),
    ("LMARTIN", "Martin, Laura"),   ("EWILSON","Wilson, Emma"),   ("VSINGH",  "Singh, Vikram"),
    ("HJONES",  "Jones, Helen"),    ("BMILLER","Miller, Ben"),    ("AROY",    "Roy, Arjun"),
]
NO_ORACLE = ["Gupta, Neha", "Iyer, Meera"]   # in HR mapping, never had ERP access

D = date
OPEN = None
# (user, responsibility, app, start, end)
RESP = [
    # Note: "Manager"/"Consultant" names are not transactional at Level 1, so those scenarios surface at Level 3 (functions).
    # AP end-to-end: supplier + invoice + payment -> 3 HIGH conflicts. Old buyer role end-dated (ignored).
    ("JSMITH", "Payables Manager", "SQLAP", D(2022,4,1), OPEN),
    ("JSMITH", "AP Payment Processing", "SQLAP", D(2022,4,1), OPEN),
    ("JSMITH", "Supplier Maintenance", "SQLAP", D(2023,1,10), OPEN),
    ("JSMITH", "Purchasing Buyer", "PO", D(2021,1,1), D(2025,12,31)),
    ("PKUMAR", "Purchasing Buyer", "PO", D(2021,6,1), OPEN),            # order + receipt
    ("PKUMAR", "Receiving User", "PO", D(2021,6,1), OPEN),
    ("ALEE", "General Ledger Super User", "SQLGL", D(2020,3,1), OPEN),  # journals + period close, super user
    ("ALEE", "GL Period Control", "SQLGL", D(2020,3,1), OPEN),
    ("MGARCIA", "Receivables Manager", "AR", D(2022,2,1), OPEN),        # lapping: AR + cash + bank rec
    ("MGARCIA", "AR Cash Receipts", "AR", D(2022,2,1), OPEN),
    ("MGARCIA", "Bank Reconciliation", "CE", D(2023,7,1), OPEN),
    ("RWONG", "Order Management Super User", "ONT", D(2021,9,1), OPEN), # order + ship
    ("RWONG", "Shipping Execution", "WSH", D(2021,9,1), OPEN),
    ("SNAIR", "HR Manager", "PER", D(2019,5,1), OPEN),                  # ghost employee
    ("SNAIR", "Payroll Manager", "PAY", D(2019,5,1), OPEN),
    ("DCHEN", "Inventory Manager", "INV", D(2022,8,1), OPEN),           # adjust + count
    ("DCHEN", "Cycle Count User", "INV", D(2022,8,1), OPEN),
    ("KPATEL", "System Administrator", "SYSADMIN", D(2024,1,15), OPEN), # sysadmin access is per admin policy -> not reported; Payables Manager alone -> clean
    ("KPATEL", "Payables Manager", "SQLAP", D(2024,1,15), OPEN),
    ("TBROWN", "iProcurement Requisitioner", "ICX", D(2023,3,1), OPEN), # clean: inquiries excluded
    ("TBROWN", "Payables Inquiry", "SQLAP", D(2023,3,1), OPEN),
    ("TBROWN", "General Ledger Inquiry", "SQLGL", D(2023,3,1), OPEN),
    ("LMARTIN", "Payables Manager", "SQLAP", D(2022,11,1), OPEN),       # self-approval exposure
    ("LMARTIN", "Expense Auditor", "SQLAP", D(2022,11,1), OPEN),
    ("EWILSON", "Customer Master Maintenance", "AR", D(2023,5,1), OPEN),# credit limit + orders
    ("EWILSON", "Order Entry Clerk", "ONT", D(2023,5,1), OPEN),
    ("VSINGH", "Payables Manager", "SQLAP", D(2020,1,1), OPEN),         # leaver, still active -> HIGH
    ("VSINGH", "Payables Inquiry", "SQLAP", D(2020,1,1), OPEN),
    ("HJONES", "Receivables Manager", "AR", D(2021,1,1), D(2026,9,30)), # leaver, end-dated late -> MEDIUM
    ("BMILLER", "Inventory Manager", "INV", D(2021,1,1), D(2026,6,30)), # leaver, clean -> OK
    ("AROY", "Purchasing Buyer", "PO", D(2026,9,1), OPEN),              # joiner with conflicting access
    ("AROY", "Supplier Maintenance", "SQLAP", D(2026,9,1), OPEN),
    ("XLIU", "Inventory Manager", "INV", D(2025,2,1), OPEN),            # not in HR mapping -> Unmatched
    ("AP_INTERFACE", "Payables Manager", "SQLAP", D(2019,1,1), OPEN),   # generic account w/ transactional access
    ("SHARED_FINANCE", "General Ledger Inquiry", "SQLGL", D(2019,1,1), OPEN),
    ("SYSADMIN", "System Administrator", "SYSADMIN", D(2015,1,1), OPEN),# seeded admin account -> excluded per admin policy
    ("SYSADMIN", "Payables Manager", "SQLAP", D(2015,1,1), OPEN),
]

book("Demo_Oracle_Responsibilities.xlsx", {"Responsibility_data":
    [("User Name", "Responsibility Name", "Start Date", "End Date", "Responsibility Key", "Application Short Name")]
    + [(u, r, s, e, r.upper().replace(" ", "_"), a) for u, r, a, s, e in RESP]})

mapping = [(1000 + i, full.split(",")[0], f"{u.lower()}@example.com", full) for i, (u, full) in enumerate(PEOPLE)]
mapping += [(2000 + i, full.split(",")[0], f"{full.split(', ')[1][0].lower()}{full.split(',')[0].lower()}@example.com", full)
            for i, full in enumerate(NO_ORACLE)]
book("Demo_Name_Mapping.xlsx", {"Name_Oracle": [("PERSON_ID", "LAST_NAME", "EMAIL_ADDRESS", "FULL_NAME")] + mapping})

upper = lambda full: " ".join(reversed(full.split(", "))).upper()
leavers = {"VSINGH": D(2026,8,31), "HJONES": D(2026,7,15), "BMILLER": D(2026,6,30)}
active = [(f"E{1000+i}", upper(f), "Finance" if i % 2 else "Operations", "Active")
          for i, (u, f) in enumerate(PEOPLE) if u not in leavers and u != "AROY"]
book("Demo_UAR_HR_Lists.xlsx", {
    "Overall": [("Employee ID", "Employee Name", "Function", "Status")] + active,
    "Joiner": [("Employee ID", "Employee Name", "Designation", "Join Date", "Status"),
               ("E1014", "ARJUN ROY", "Procurement Executive", D(2026,9,1), "Active"),
               ("E2001", "MEERA IYER", "Finance Analyst", D(2026,9,15), "Active"),
               ("E3001", "KARAN MEHTA", "Sales Executive", D(2026,10,1), "Active")],      # not in mapping
    "Exit": [("Employee ID", "Employee Name", "Designation", "Resigned On", "Exit Date", "Status"),
             ("E1011", "VIKRAM SINGH", "AP Lead", D(2026,7,31), leavers["VSINGH"], "Exited"),
             ("E1012", "HELEN JONES", "AR Analyst", D(2026,6,15), leavers["HJONES"], "Exited"),
             ("E1013", "BEN MILLER", "Stores Supervisor", D(2026,5,31), leavers["BMILLER"], "Exited"),
             ("E2000", "NEHA GUPTA", "HR Executive", D(2026,8,1), D(2026,8,31), "Exited"),
             ("E3002", "RAVI SHANKAR", "Driver", D(2026,8,1), D(2026,9,1), "Exited")],    # not in mapping
})

# Level 3: functions reachable per responsibility (standard R12 menus, one deliberate customisation)
FUNCS = {
    "Payables Manager": [("AP_APXINWKB", "Invoice Workbench"), ("AP_APXPAWKB", "Payments Workbench")],  # conflict inside ONE resp
    "AP Payment Processing": [("AP_APXPAWKB", "Payments Workbench"), ("IBY_FD_PAYMENT_ADMIN", "Payment Manager")],
    "Supplier Maintenance": [("AP_APXVDMVD", "Suppliers")],
    "Purchasing Buyer": [("PO_POXPOEPO", "Enter Purchase Orders"), ("PO_POXBWVRP", "AutoCreate Documents")],
    "Receiving User": [("RCV_RCVRCVRC", "Receiving Transactions Summary (Inquiry)")],   # customised: view only
    "iProcurement Requisitioner": [("ICX_POR_LAUNCH_IP", "iProcurement Home")],
    "Payables Inquiry": [("AP_APXINWKB_Q", "Invoice Workbench (Inquiry)")],
    "General Ledger Inquiry": [("GL_GLXIQJRN", "Journal Inquiry")],
    "General Ledger Super User": [("GL_GLXJEENT", "Enter Journals"), ("GL_GLXCOOPS", "Open and Close Periods")],
    "GL Period Control": [("GL_GLXCOOPS", "Open and Close Periods")],
    "Receivables Manager": [("AR_ARXTWMAI", "Transactions Workbench"), ("AR_ARXRWMAI", "Receipts Workbench")],
    "AR Cash Receipts": [("AR_ARXRWMAI", "Receipts Workbench")],
    "Bank Reconciliation": [("CE_CEXRECRE", "Reconcile Bank Statements")],
    "Order Management Super User": [("ONT_OEXOEORD", "Sales Orders"), ("WSH_WSHFSTRX", "Shipping Transactions")],
    "Shipping Execution": [("WSH_WSHFSTRX", "Shipping Transactions")],
    "HR Manager": [("PER_PERWSHRG", "Enter and Maintain People")],
    "Payroll Manager": [("PAY_PAYWSQPY", "QuickPay")],
    "Inventory Manager": [("INV_INVTTMTX", "Miscellaneous Transactions"), ("INV_INVADCYC", "Cycle Count Entries")],
    "Cycle Count User": [("INV_INVADCYC", "Cycle Count Entries")],
    "System Administrator": [("FND_FNDSCAUS", "Define Users"), ("FND_FNDSCRSP", "Define Responsibility")],
    "Expense Auditor": [("AP_WEB_AUDIT_QUEUE", "Expense Report Audit")],
    "Customer Master Maintenance": [("AR_ARXCUDCI_STD", "Customers Standard")],
    "Order Entry Clerk": [("ONT_OEXOEORD", "Sales Orders")],
}
# Same extract as above with function columns (client dump layout, one row per function) -> switches on Level 3
EMP_ID = {u: 1000 + i for i, (u, _) in enumerate(PEOPLE)}
book("Demo_Oracle_Access_with_Functions.xlsx", {"Access":
    [("User Name", "EMAIL_ADDRESS", "EMPLOYEE_ID", "Responsibility Name", "Start Date", "End Date",
      "Responsibility Key", "Application Short Name", "Function Code", "Function Name")]
    + [(u, f"{u.lower()}@example.com" if u in EMP_ID else None, EMP_ID.get(u), r, s, e, r.upper().replace(" ", "_"), a, code, name)
       for u, r, a, s, e in RESP for code, name in FUNCS[r]]})

# Optional: client-supplied matrix (overrides built-in conflicts; built-in classifications still apply)
book("Demo_Client_SoD_Matrix.xlsx", {"Conflicts": [("Category A", "Category B", "Risk", "Description"),
    ("SUPPLIER_MAINT", "AP_PAYMENTS", "CRITICAL", "Client policy: vendor master + payment release is zero-tolerance"),
    ("AP_INVOICE_ENTRY", "AP_PAYMENTS", "HIGH", "Enter invoices and release payments"),
    ("PO_CREATE", "PO_RECEIVING", "HIGH", "Order goods and receipt them"),
    ("HR_MAINTAIN", "PAYROLL", "HIGH", "Maintain employees and run payroll"),
    ("GL_JOURNAL", "GL_PERIOD_CONTROL", "HIGH", "Post journals and reopen periods")]})

print("Wrote:", *sorted(p.name for p in OUT.glob("Demo_*.xlsx")), sep="\n  ")
