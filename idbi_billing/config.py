"""
Master configuration — all account lists, BoM line values, and fixed rules.
Edit this file (not the pricing code) when accounts or the BoM change.

v3 changes (see RULES.md):
  * Direct Connect BoM line added (point 2).
  * Windows m-family compute priced from m5a BoM lines; m5a used in place of
    m6a; Windows r6 stays non-BoM @8% (point 1).
  * Instance spec table expanded so no real instance resolves to (0,0)
    (Case-2 fix). A parametric fallback in pricing.py backstops anything
    still missing.
"""

# ── Allowed accounts ────────────────────────────────────────────────────────
DBD_ACCOUNTS = {
    "528757802229": "DBD_Sandbox_Account",
    "562559071104": "Kaizen",
    "294539078227": "ImpactSure",
    "689397855052": "Atlas_Sandbox",
    "417651893846": "DBD-Network_Account",
    "919437049652": "SB-Atlas-Account",
}

ITD_ACCOUNTS = {
    "122610479172": "IDBI_Bank",
    "354929337302": "Network_Account",
    "647910641444": "Infra_Dev_Account",
    "435200406974": "Infra_Prod_Account",
    "789665322921": "Common_Services_Account",
    "347887033541": "Audit_Account",
    "497783363028": "Log_Archive_Account",
    "403276262026": "Infra_Sandbox_Account",
    "226202864192": "Infra_UAT_Account",
}

# ── BoM line values (read live from workbook; these are fallbacks) ──────────
# Format: key -> {unit: INR, disc: fraction, basis: quantity the unit covers}

BOM = {
    # Windows compute (disc 40.77%) — anchored on m5a / c6a instances
    "win_4v8g":   {"unit": 12620.3275, "disc": 0.4077, "basis": 1, "line": 1,  "inst": "c6a.xlarge"},
    "win_4v16g":  {"unit": 13088.9752, "disc": 0.4077, "basis": 1, "line": 2,  "inst": "m5a.xlarge"},
    "win_8v16g":  {"unit": 23393.7700, "disc": 0.4077, "basis": 1, "line": 3,  "inst": "c6a.2xlarge"},
    "win_8v32g":  {"unit": 24869.0528, "disc": 0.4077, "basis": 1, "line": 4,  "inst": "m5a.2xlarge"},
    "win_16v32g": {"unit": 61562.5528, "disc": 0.4077, "basis": 1, "line": 5,  "inst": "c6a.4xlarge"},
    "win_16v64g": {"unit": 68374.4893, "disc": 0.4077, "basis": 1, "line": 6,  "inst": "m5a.4xlarge"},
    # RHEL compute (disc 41.28%)
    "rhel_2v4g":  {"unit":  4342.6431, "disc": 0.4128, "basis": 1, "line": 7,  "inst": "c6g.large"},
    "rhel_2v8g":  {"unit":  4768.3924, "disc": 0.4128, "basis": 1, "line": 8,  "inst": "m6g.large"},
    "rhel_4v8g":  {"unit":  6794.9591, "disc": 0.4128, "basis": 1, "line": 9,  "inst": "c6g.xlarge"},
    "rhel_4v16g": {"unit":  7220.7084, "disc": 0.4128, "basis": 1, "line": 10, "inst": "m6g.xlarge"},
    "rhel_4v32g": {"unit":  7987.0572, "disc": 0.4128, "basis": 1, "line": 11, "inst": "r6a.xlarge"},
    "rhel_8v16g": {"unit": 13947.5477, "disc": 0.4128, "basis": 1, "line": 12, "inst": "c6a.2xlarge"},
    "rhel_8v32g": {"unit": 14799.0463, "disc": 0.4128, "basis": 1, "line": 13, "inst": "m6g.2xlarge"},
    "rhel_16v32g":{"unit": 20418.9373, "disc": 0.4128, "basis": 1, "line": 14, "inst": "c6a.4xlarge"},
    "rhel_16v64g":{"unit": 22136.5463, "disc": 0.4128, "basis": 1, "line": 15, "inst": "m6g.4xlarge"},
    # Block SSD — 5 slabs (disc 41.08%)
    "ebs_128":    {"unit":   980.6860, "disc": 0.41078, "basis": 128,  "line": 16},
    "ebs_256":    {"unit":  1961.3889, "disc": 0.41078, "basis": 256,  "line": 17},
    "ebs_512":    {"unit":  3922.7608, "disc": 0.41078, "basis": 512,  "line": 18},
    "ebs_1024":   {"unit":  8860.6938, "disc": 0.41078, "basis": 1024, "line": 19},
    "ebs_2048":   {"unit": 19157.9472, "disc": 0.41078, "basis": 2048, "line": 20},
    # Object Storage (disc 39.18%)
    "s3_hot":     {"unit": 53550.0,   "disc": 0.39176, "basis": 25000, "line": 21},
    # Network / infra
    "net_xfer":   {"unit":   983.7,   "disc": 0.80,    "basis": 100,   "line": 29},
    "nat_data":   {"unit":  4183.2,   "disc": 0.70,    "basis": 100,   "line": 30},
    "static_ip":  {"unit":   328.5,   "disc": 0.99,    "basis": 1,     "line": 31},
    "app_lb":     {"unit":  4150.8,   "disc": 0.75,    "basis": 730,   "line": 36},
    "net_fw":     {"unit": 57893.4,   "disc": 0.60,    "basis": 1,     "line": 37},
    "waf":        {"unit":  1377.0,   "disc": 0.40288, "basis": 1000000,"line": 38},
    "kms":        {"unit":  4502.7,   "disc": 0.99,    "basis": 1,     "line": 39},
    # Direct Connect (point 2) — "Dedicated DC-Cloud connect leased line charges"
    # ₹19,912.5 covers 1 Gbps (= 1000 Mbps). basis = 1000 → per-Mbps pricing.
    # Costing: (((19912.5/1000) * actual_mbps) * quantity) / rate ; discount 75%.
    "direct_connect": {"unit": 19912.5, "disc": 0.75, "basis": 1000, "line": 40,
                       "name": "Dedicated DC-Cloud connect leased line charges"},
    # Site-to-Site VPN (BoM line 35) — per connection per month, disc 78%.
    "vpn": {"unit": 10035.0, "disc": 0.78, "basis": 1, "line": 35,
            "name": "Site-to-Site VPN"},
}

# ── EBS slab helper ─────────────────────────────────────────────────────────
def ebs_key(gb: float) -> str:
    if gb <= 128:   return "ebs_128"
    elif gb <= 256: return "ebs_256"
    elif gb <= 512: return "ebs_512"
    elif gb <= 1024:return "ebs_1024"
    else:           return "ebs_2048"

# ── EC2 substitution table ───────────────────────────────────────────────────
# (os_class, vcpu, mem_gb) -> bom_key  (exact matches only)
EC2_EXACT = {
    ("win",  4,  8): "win_4v8g",   ("win",  4, 16): "win_4v16g",
    ("win",  8, 16): "win_8v16g",  ("win",  8, 32): "win_8v32g",
    ("win", 16, 32): "win_16v32g", ("win", 16, 64): "win_16v64g",
    ("rhel", 2,  4): "rhel_2v4g",  ("rhel", 2,  8): "rhel_2v8g",
    ("rhel", 4,  8): "rhel_4v8g",  ("rhel", 4, 16): "rhel_4v16g",
    ("rhel", 4, 32): "rhel_4v32g", ("rhel", 8, 16): "rhel_8v16g",
    ("rhel", 8, 32): "rhel_8v32g", ("rhel",16, 32): "rhel_16v32g",
    ("rhel",16, 64): "rhel_16v64g",
}
# Scaling anchors for non-exact sizes
EC2_ANCHOR = {
    ("rhel", "r6"): ("rhel_4v32g",  4),   # r6a.xlarge, anchor vCPU=4
    ("rhel", "m6"): ("rhel_16v64g", 16),  # m6g.4xlarge, anchor vCPU=16
    ("win",  "m6"): None,                 # resolved dynamically from m5a table
}
# Windows m5a anchors for scaling (vcpu -> bom_key). Windows m5a AND m6a both
# scale through these — "take m5a in place of m6a" (point 1).
WIN_M5A_ANCHORS = {4: "win_4v16g", 8: "win_8v32g", 16: "win_16v64g"}
# Windows compute families that map onto the m5a BoM lines.
WIN_M_FAMILIES = ("m5", "m6")

# ── Styling constants ────────────────────────────────────────────────────────
ORANGE   = "F4B084"
GREEN    = "FF92D050"
HDR_BLUE = "4472C4"

COL_HEADERS = [
    "Serial No.", "Account ID", "Account Name", "Service",
    "Additional details", "Configuration of Service", "AWS SKU",
    "Quantity (X)", "Monthy cost", "Indicative Unit Cost",
    "Discount %", "Discounted Unit Cost", "Monthly Cost Discounted",
    "Converted Monthly Cost", "Notes",
]
COL_WIDTHS = [7.0, 15.0, 19.0, 22.0, 24.0, 30.0, 31.0,
              9.0, 11.5, 12.5, 8.5, 12.5, 13.5, 14.5, 6.5]

NOTE_TEXT = (
    "Note:- Rows highlighted in orange are priced as per the finalised "
    "Section B AWS BoM (unit price and discount % from BoM, scaled to "
    "actual {month} consumption). Non-highlighted rows are at 8% discount "
    "off public pricing, as per Section D."
)
NON_BOM_DISC = 0.08
ADDR_HRS_PER_MONTH = 730
