APP_NAME = "ServeOne ERP Portal"
PAGE_TITLE = "ServeOne ERP Portal"
PAGE_ICON = "🏢"

ENABLE_SUPABASE = False  # tahap 1: masih pakai auth lokal agar transisi aman

DEFAULT_DB_KEYS = [
    "db_attendance",
    "db_tasks",
    "db_logs",
    "db_leave",
    "db_ot",
]

USERS = {
    "andrey": {
        "pass": "admin123",
        "name": "Andrey Mangun Parisyanto",
        "role": "Purchasing Supervisor",
        "lokasi": "Gunung Putri",
    },
    "staff": {
        "pass": "1234",
        "name": "Staff Operasional",
        "role": "Operator",
        "lokasi": "Bogor",
    },
}

ROLE_CAN_VIEW_DASHBOARD = ["Manager", "Supervisor", "Purchasing Supervisor", "Admin"]
PRIORITY_OPTIONS = ["P1 (Urgent & Important)", "P2 (Important, Not Urgent)", "P3 (Low)"]
LEAVE_CODE_OPTIONS = ["CT - Cuti Tahunan", "S - Sakit", "I - Izin"]
