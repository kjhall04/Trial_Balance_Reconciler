import re
from typing import Dict, Tuple


# Buckets
CATEGORY_RANGES: Dict[str, Tuple[int, int]] = {
    "assets": (1000, 1999),
    "liabilities": (2000, 2999),
    "equity": (3000, 3999),
    "income": (4000, 4999),
    "cogs": (5000, 5999),
    "admin": (6000, 6999),
    "other": (7000, 7999),
}


NORMALIZATION_RULES: list[tuple[str, str]] = [
    (r"&", " and "),
    (r"\ba/r\b", "accounts receivable"),
    (r"\ba/p\b", "accounts payable"),
    (r"\bc o g s\b", "cogs"),
    (r"\bw i p\b", "wip"),
    (r"\bl/t\b", "long term"),
    (r"\bs/t\b", "short term"),
    (r"\bloc\b", "line of credit"),
    (r"\bn/p\b", "notes payable"),
    (r"\bu/f\b", "undeposited funds"),
    (r"\bacc dep\b", "accumulated depreciation"),
    (r"\baccum depr\b", "accumulated depreciation"),
    (r"\bcip\b", "construction in progress"),
    (r"\bppe\b", "property plant equipment"),
    (r"\bdepr\b", "depreciation"),
    (r"\bamort\b", "amortization"),
]


STRONG_PHRASES = {
    "assets": [
        "accounts receivable",
        "retainage receivable",
        "costs and estimated earnings in excess of billings",
        "costs in excess of billings",
        "deficit billings",
        "prepaid expenses",
        "prepaid insurance",
        "prepaid rent",
        "undeposited funds",
        "cash on hand",
        "cash in bank",
        "deposit in transit",
        "security deposit",
        "vendor deposit",
        "deferred tax asset",
        "due from shareholder",
        "due from related party",
        "right of use asset",
        "rou asset",
        "construction in progress",
        "allowance for doubtful accounts",
    ],
    "liabilities": [
        "accounts payable",
        "retainage payable",
        "sales tax payable",
        "payroll tax payable",
        "accrued expenses",
        "accrued payroll",
        "accrued wages",
        "deferred revenue",
        "unearned revenue",
        "customer deposits",
        "customer deposit",
        "billings in excess of costs and estimated earnings",
        "billings in excess of costs",
        "excess billings",
        "contract liability",
        "lease liability",
        "deferred rent",
        "current maturities of long term debt",
        "line of credit",
        "notes payable",
        "income tax payable",
        "due to shareholder",
        "due to related party",
        "bank overdraft",
    ],
    "equity": [
        "opening balance equity",
        "retained earnings",
        "member equity",
        "owners equity",
        "owner equity",
        "partner capital",
        "capital contribution",
        "owner contribution",
        "common stock",
        "preferred stock",
        "capital stock",
        "additional paid in capital",
        "paid in capital",
        "treasury stock",
        "accumulated other comprehensive income",
        "shareholder distributions",
        "member distributions",
        "stockholders equity",
        "shareholders equity",
        "accumulated deficit",
    ],
    "income": [
        "sales revenue",
        "service revenue",
        "service income",
        "consulting revenue",
        "consulting income",
        "contract revenue",
        "contract income",
        "construction revenue",
        "net sales",
        "gross sales",
        "rental income",
        "interest income",
        "dividend income",
        "management fee income",
        "commission revenue",
        "other income",
        "miscellaneous income",
        "forgiveness income",
    ],
    "cogs": [
        "cost of goods sold",
        "cost of services",
        "cost of contract revenue",
        "cost of construction",
        "job cost",
        "direct costs",
        "direct labor",
        "direct materials",
        "subcontract costs",
        "field payroll",
        "job payroll",
        "job materials",
        "job supplies",
        "labor burden",
    ],
    "admin": [
        "payroll tax expense",
        "employer payroll taxes",
        "office rent",
        "professional fees",
        "legal fees",
        "accounting fees",
        "software subscription",
        "office supplies",
        "bank service charges",
        "depreciation expense",
        "amortization expense",
    ],
    "other": [
        "interest expense",
        "finance charges",
        "gain on sale",
        "loss on sale",
        "gain on disposal",
        "loss on disposal",
        "income tax expense",
        "prior year adjustment",
        "prior period adjustment",
        "unrealized gain",
        "unrealized loss",
    ],
}


STRONG_REGEX = {
    "assets": [
        r"\bcosts?( and estimated earnings?)? in excess of billings?\b",
        r"\bdeficit billings?\b",
        r"\bdue from\b",
        r"\bdeposit in transit\b",
        r"\b(prepaid|prepaids)\b",
    ],
    "liabilities": [
        r"\bbillings? in excess of costs?( and estimated earnings?)?\b",
        r"\bexcess billings?\b",
        r"\bdue to\b",
        r"\bcurrent maturities?( of)? (long term debt|ltd)\b",
        r"\b(accrued|payroll liabilities)\b",
        r"\b(bank overdraft|overdraft)\b",
    ],
    "equity": [
        r"\bowners? draws?\b",
        r"\bmembers? distributions?\b",
        r"\bcapital contributions?\b",
        r"\b(stockholders?|shareholders?) equity\b",
    ],
    "income": [
        r"\b(service|contract|construction|consulting|management fee|commission|rental|interest|dividend) (income|revenue)\b",
        r"\b(net|gross) sales\b",
        r"\bother income\b",
    ],
    "cogs": [
        r"\bcost of (goods|sales|services|contract revenue|construction)\b",
        r"\b(job|project|contract|direct|field) (cost|costs|labor|labour|materials|payroll)\b",
        r"\bsubcontract(or)? (cost|costs)\b",
    ],
    "admin": [
        r"\b(office|administrative) (rent|payroll|salary|salaries|expense|supplies)\b",
        r"\b(payroll tax expense|employer payroll taxes)\b",
    ],
    "other": [
        r"\b(gain|loss) on\b",
        r"\b(prior period|prior year) adjustment\b",
        r"\b(unrealized|realized) (gain|loss)\b",
    ],
}


ASSET_KEYWORDS = [
    # Cash and banking
    "cash", "petty cash", "cash on hand", "cash in bank", "bank", "checking", "savings",
    "money market", "cash sweep", "sweep", "lockbox", "merchant account", "stripe", "square",
    "paypal", "undeposited funds", "clearing", "bank clearing", "cash clearing", "clearing account",
    "deposit in transit", "escrow", "restricted cash", "certificate of deposit", "cd account",

    # Receivables and contract assets
    "accounts receivable", "receivable", "trade receivable", "other receivable", "notes receivable",
    "loan receivable", "retainage receivable", "retention receivable", "unbilled", "unbilled receivable",
    "contract asset", "progress billings receivable", "billings receivable", "due from customer",
    "customer receivable", "employee receivable", "shareholder receivable", "related party receivable",
    "due from", "tax refund receivable", "income tax receivable", "sales tax receivable",
    "vat receivable", "gst receivable",

    # Inventory, WIP, and current assets
    "inventory", "inventory reserve", "allowance for obsolete inventory", "materials",
    "raw materials", "finished goods", "work in process", "wip", "supplies", "preproduction",
    "consumables", "shop supplies", "small tools", "deficit billings", "costs in excess of billings",

    # Prepaids and deposits
    "prepaid", "prepaids", "prepaid expense", "prepaid expenses", "prepaid rent", "prepaid insurance",
    "prepaid taxes", "prepaid license", "prepaid licenses", "prepaid software", "prepaid subscription",
    "security deposit", "utility deposit", "vendor deposit", "insurance deposit", "work comp deposit",
    "advance to vendor", "advances", "deposit",

    # Fixed assets and long-term assets
    "equipment", "office equipment", "machinery", "vehicles", "vehicle", "truck", "trailer",
    "forklift", "computer", "computers", "hardware", "software asset", "furniture", "fixtures",
    "leasehold", "leasehold improvements", "building", "land", "property", "construction in progress",
    "property plant equipment", "tools", "tooling", "capitalized", "fixed asset", "fixed assets",

    # Contra assets and other assets
    "accumulated depreciation", "accumulated amortization", "allowance for doubtful accounts",
    "allowance for bad debts", "intangible", "goodwill", "right of use asset", "rou asset",
    "lease asset", "deferred charge", "deferred charges", "deferred costs", "deferred tax asset",
    "other current asset", "other asset",
]


LIABILITY_KEYWORDS = [
    # Payables and accruals
    "accounts payable", "payable", "trade payable", "vendor payable", "retainage payable",
    "retention payable", "accrued", "accrual", "accrued expenses", "accrued payroll", "accrued wages",
    "accrued compensation", "accrued bonus", "accrued commissions", "accrued vacation", "accrued pto",
    "accrued rent", "accrued interest", "accrued taxes", "payroll liabilities", "benefits payable",
    "sales commissions payable", "payable to owner", "payable to members", "payable to shareholder",

    # Taxes payable
    "sales tax payable", "use tax payable", "payroll tax payable", "withholding", "fica payable",
    "medicare payable", "futa payable", "suta payable", "state withholding", "federal withholding",
    "income tax payable", "property tax payable", "franchise tax payable", "tax payable",

    # Debt and financing
    "loan", "loans payable", "note payable", "notes payable", "debt", "line of credit",
    "credit line", "credit card payable", "merchant payable", "current maturities", "current portion",
    "short term debt", "short term loan", "long term debt", "lt debt", "installment", "bank overdraft",
    "overdraft", "loan from shareholder", "shareholder note", "member loan",

    # Deferred and contract liabilities
    "deferred revenue", "unearned revenue", "customer deposit", "customer deposits", "customer advance",
    "customer advances", "advance from customer", "deferred income", "contract liability",
    "billings in excess of costs", "billings in excess of costs and estimated earnings", "excess billings",

    # Leases and other liabilities
    "lease liability", "rou liability", "operating lease liability", "finance lease liability",
    "deferred rent", "warranty liability", "due to", "due to shareholder", "due to stockholder",
    "due to member", "due to related party", "related party payable", "other current liability",
    "other liability",
]


EQUITY_KEYWORDS = [
    "equity", "owners equity", "owner equity", "member equity", "members equity", "partner equity",
    "partners equity", "stockholders equity", "shareholders equity", "capital", "capital account",
    "member capital", "partner capital", "owners capital", "owner contribution", "owner contributions",
    "member contribution", "member contributions", "capital contribution", "capital contributions",
    "paid in capital", "paid-in capital", "contributed capital", "additional paid in capital", "apic",
    "pic", "common stock", "preferred stock", "capital stock", "treasury stock", "retained earnings",
    "opening balance", "opening balance equity", "prior year retained earnings", "accumulated other comprehensive income",
    "accumulated deficit", "deficit equity", "draw", "draws", "owner draw", "owner draws",
    "member draw", "distribution", "distributions", "member distributions", "shareholder distributions",
    "dividend", "dividends", "current year earnings", "current year profit", "current year loss",
    "net income", "net profit", "net loss",
]


INCOME_KEYWORDS = [
    "revenue", "sales", "sales revenue", "service income", "service revenue", "fees earned", "fee income",
    "earned fees", "consulting income", "consulting revenue", "management fee", "management fee income",
    "commission income", "commission revenue", "contract revenue", "contract income", "job revenue",
    "project revenue", "construction revenue", "rental income", "lease income", "royalty income",
    "interest income", "dividend income", "rebate income", "refund income", "other income",
    "misc income", "miscellaneous income", "grant income", "forgiveness income", "change order revenue",
    "change order income", "pass through income", "rebill income", "markup income", "shop income",
    "labor income", "product sales", "net sales", "gross sales", "sales returns", "sales discounts",
]


COGS_KEYWORDS = [
    "cogs", "cost of goods", "cost of goods sold", "cost of sales", "cost of service", "cost of services",
    "cost of contract revenue", "cost of construction", "direct labor", "direct labour", "field labor",
    "field labour", "field payroll", "job payroll", "project payroll", "crew wages", "direct materials",
    "material cost", "job materials", "subcontract", "subcontractor", "subcontract costs", "subs",
    "1099 labor", "1099 labour", "freight in", "inbound freight", "job cost", "job costs",
    "project costs", "contract costs", "equipment rental", "job equipment rental", "rental equipment",
    "job supplies", "site supplies", "small tools job", "job permits", "permit job", "burden labor",
    "labor burden", "construction costs", "wip adjustment", "inventory adjustment", "cost applied to jobs",
]


ADMIN_KEYWORDS = [
    # Payroll and employee costs
    "wages", "salary", "salaries", "salary expense", "salaries and wages", "office payroll",
    "administrative payroll", "payroll", "payroll taxes", "payroll tax expense", "employer payroll taxes",
    "benefits", "health insurance", "medical insurance", "dental", "vision", "workers comp",
    "workers compensation", "workers compensation insurance", "retirement", "401k", "ira",
    "payroll processing", "recruiting", "training", "education",

    # Occupancy and utilities
    "rent", "office rent", "warehouse rent", "storage", "lease expense", "utilities", "electric",
    "electricity", "water", "gas", "trash", "internet", "phone", "telephone", "cell phone", "mobile",
    "security", "alarm", "janitorial", "cleaning",

    # Insurance and professional services
    "insurance", "general liability", "gl insurance", "auto insurance", "property insurance",
    "liability insurance", "professional fees", "legal", "legal fees", "attorney", "accounting",
    "accounting fees", "audit", "audit fees", "tax prep", "bookkeeping", "consulting", "outside services",

    # Sales, marketing, and office operations
    "advertising", "marketing", "promotion", "website", "seo", "leads", "sponsorship",
    "office supplies", "postage", "shipping", "printing", "software", "software subscription",
    "subscriptions", "dues", "dues and subscriptions", "licenses", "license", "permits", "membership",
    "meals", "travel", "lodging", "airfare", "mileage", "fuel", "repairs", "maintenance",
    "repairs and maintenance", "vehicle expense", "auto expense", "bank fees", "bank service charge",
    "service charge", "merchant fees", "credit card fees", "office expense", "uniforms",
    "telephone expense", "internet expense", "depreciation expense", "amortization expense",
    "bad debt", "bad debts",
]


OTHER_KEYWORDS = [
    "interest expense", "finance charge", "finance charges", "other expense", "misc expense",
    "miscellaneous expense", "gain", "loss", "gain on sale", "loss on sale", "gain on disposal",
    "loss on disposal", "gain on extinguishment", "unrealized gain", "unrealized loss", "realized gain",
    "realized loss", "income tax", "income taxes", "income tax expense", "tax expense",
    "charitable", "donation", "donations", "penalty", "penalties", "late fee", "late fees",
    "rounding", "round off", "prior period", "prior year adjustment", "foreign exchange", "fx gain",
    "fx loss",
]


def _normalize_account_name(account_name: str) -> str:
    s = (account_name or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    for pattern, replacement in NORMALIZATION_RULES:
        s = re.sub(pattern, replacement, s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def guess_category(account_name: str) -> str:
    s = _normalize_account_name(account_name)

    for cat, phrases in STRONG_PHRASES.items():
        for phrase in phrases:
            if phrase in s:
                return cat

    for cat, patterns in STRONG_REGEX.items():
        for pattern in patterns:
            if re.search(pattern, s):
                return cat

    scores = {k: 0 for k in CATEGORY_RANGES.keys()}

    def add_score(cat: str, keywords: list[str], weight: int = 1):
        for kw in keywords:
            if kw in s:
                scores[cat] += weight

    add_score("assets", ASSET_KEYWORDS, 1)
    add_score("liabilities", LIABILITY_KEYWORDS, 1)
    add_score("equity", EQUITY_KEYWORDS, 2)
    add_score("income", INCOME_KEYWORDS, 2)
    add_score("cogs", COGS_KEYWORDS, 2)
    add_score("admin", ADMIN_KEYWORDS, 1)
    add_score("other", OTHER_KEYWORDS, 2)

    # Tie breakers and special overrides
    if "payable" in s or "accrued" in s or "due to" in s or "to pay" in s:
        scores["liabilities"] += 3

    if "receivable" in s or "due from" in s:
        scores["assets"] += 3

    if "deferred tax asset" in s:
        scores["assets"] += 4

    if "deferred tax liability" in s:
        scores["liabilities"] += 4

    if "customer deposit" in s or "unearned revenue" in s or "deferred revenue" in s:
        scores["liabilities"] += 4

    if "payroll tax payable" in s or "sales tax payable" in s or "income tax payable" in s:
        scores["liabilities"] += 4

    if "payroll tax expense" in s or "employer payroll taxes" in s:
        scores["admin"] += 4

    if "bank overdraft" in s or "overdraft" in s:
        scores["liabilities"] += 4

    if "accumulated depreciation" in s or "accumulated amortization" in s:
        scores["assets"] += 4

    if ("depreciation" in s or "amortization" in s) and "expense" in s:
        scores["admin"] += 3

    if "interest income" in s:
        scores["income"] += 4

    if "interest expense" in s or "finance charge" in s:
        scores["other"] += 4

    if "gain on" in s or "loss on" in s or "unrealized" in s or "realized" in s:
        scores["other"] += 3

    if "draw" in s or "distribution" in s or "dividend" in s or "retained earnings" in s:
        scores["equity"] += 4

    if "opening balance equity" in s:
        scores["equity"] += 5

    if "sales return" in s or "sales discount" in s:
        scores["income"] += 3

    if "cost of goods sold" in s or "job cost" in s or "subcontract" in s or "direct labor" in s:
        scores["cogs"] += 4

    if "construction in progress" in s or "wip" in s:
        scores["assets"] += 2

    best_cat = max(scores.items(), key=lambda x: x[1])[0]

    if scores[best_cat] == 0:
        return "assets"

    return best_cat
