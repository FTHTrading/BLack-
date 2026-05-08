from pathlib import Path
import csv
import zipfile

from fpdf import FPDF
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = Path(r"C:\Users\Kevan\OneDrive - FTH Trading\11-Downloads\unykorn-black-funder-site")
DOWNLOADS = ROOT / "downloads"
DOWNLOADS.mkdir(parents=True, exist_ok=True)

DATE_PREPARED = "May 8, 2026"
CONTACT_URL = "https://fifa.unykorn.org/"
BRAND = "UNYKORN BLACK"

# ---------- colour palette ----------
NAVY   = (7,  17,  31)
GOLD   = (214, 168, 79)
WHITE  = (255, 255, 255)
LIGHT  = (245, 245, 248)
RULE   = (180, 145, 55)
DARK   = (30,  30,  30)
MID    = (80,  80,  80)
ACCENT = (214, 168, 79)

SCENARIOS = [
    ("Lean",   "5 vehicles", "$723,150",   "$278,350",  "$120,000", "2.6x"),
    ("Base",   "6 vehicles", "$1,414,200", "$433,200",  "$180,000", "7.2x"),
    ("Growth", "8 vehicles", "$2,365,000", "$820,000",  "$240,000", "10.5x"),
]

USE_OF_FUNDS = [
    ("Vehicle acquisition / down payments",    "$115,000"),
    ("Equipment / vehicle financing facility", "$350,000"),
    ("Insurance reserve (12-month buffer)",    "$45,000"),
    ("Licensing, compliance, legal, CPA",      "$25,000"),
    ("Website, app, dispatch buildout",        "$35,000"),
    ("Branding, QR codes, sponsor materials",  "$20,000"),
    ("Working capital reserve",                "$60,000"),
]

EVENTS = [
    ("May 24, 2026",        "Birthday Bash ATL",          "State Farm Arena",          "B+", "Late-night VIP routes, after-event offers, sponsor nightlife traffic"),
    ("June 8-9, 2026",      "Megan Moroney",              "State Farm Arena",          "B",  "Concert transfers, hotel pickup, premium downtown routes"),
    ("June 9-14, 2026",     "Atlanta Market + Apparel",   "AmericasMart",              "A",  "Buyers, vendors, showroom executives, airport and hotel routes"),
    ("June 15-July 15, 2026","Atlanta World Cup window",  "Mercedes-Benz Stadium",     "A+", "VIP arrivals, hotel-stadium-dining loops, sponsor hospitality"),
    ("Aug. 21, 2026",       "Chris Stapleton",            "Mercedes-Benz Stadium",     "A",  "Stadium concert, suite guests, sponsor hospitality, reserved exits"),
    ("Aug. 26-30, 2026",    "TOUR Championship",          "East Lake Golf Club",       "A",  "Executive golf, sponsor hospitality, corporate blocks, airport transfers"),
    ("Aug. 27, 2026",       "AC/DC POWER UP Tour",        "Mercedes-Benz Stadium",     "A",  "Major stadium concert, premium arrivals, late-night exits"),
    ("Sept. 3-7, 2026",     "Dragon Con",                 "Downtown Atlanta",          "A",  "Multi-day hotel corridor, fan groups, VIP events, merchant offers"),
    ("Sept. 18-20, 2026",   "Shaky Knees",                "Piedmont Park",             "A",  "Festival rides, hotel packages, sponsor QR offers, group exits"),
    ("Oct. 10-11, 2026",    "Atlanta Pride",              "Piedmont Park",             "B+", "Festival mobility, hotel routes, merchant and safety-oriented routing"),
    ("Nov. 10-11, 2026",    "The R&B Tour",               "Mercedes-Benz Stadium",     "A",  "High-demand concert nights, suite guests, sponsors, premium exits"),
    ("Dec. 5, 2026",        "SEC Championship",           "Mercedes-Benz Stadium",     "A+", "Corporate hospitality, alumni groups, VIP fan travel, sponsor routes"),
]

REVENUE_MIX = [
    ("Owned dispatched rides", "$714,200", "42%", "Primary premium SUV & sedan fare revenue"),
    ("Corporate VIP blocks",   "$110,000",  "7%", "Executive account pre-booked ride blocks"),
    ("BlackPass presales",     "$150,000",  "9%", "Prepaid ride credits sold before peak windows"),
    ("Sponsor inventory",      "$275,000", "16%", "Vehicle, route, app & event sponsorships"),
    ("Merchant referrals",      "$75,000",  "4%", "QR-driven partner activation commissions"),
    ("Partner fleet margin",    "$60,000",  "4%", "Overflow surge participation revenue"),
    ("Settlement services",     "$30,000",  "2%", "Platform-linked settlement & reporting"),
]

SPONSOR_PACKAGES = [
    (
        "Founding Sponsor",
        "$150,000",
        "Category-exclusive placement across site, app, fleet graphics, booking confirmations, and settlement proof reporting.",
        "Includes category exclusivity per launch window, custom QR offer slot, quarterly performance review, and first-right-of-renewal for 2027.",
    ),
    (
        "Route Sponsor",
        "$25,000",
        "Own a defined priority corridor - airport-to-hotel, stadium exits, or dining district loops.",
        "Includes route-branded digital placements, location-based QR offer slots, and monthly route performance summary.",
    ),
    (
        "Vehicle Sponsor",
        "$15,000",
        "Sponsor a named premium SUV or Sprinter with visible exterior and interior brand placement.",
        "Includes branded passenger handoffs, confirmation touchpoints, and tracked rider exposure metrics per event window.",
    ),
    (
        "Merchant Listing",
        "$2,500",
        "Verified in-app placement for restaurants, hotels, lounges, retail, and premium services.",
        "Includes periodic promotion support, QR activation, and referral performance summaries delivered monthly.",
    ),
]

BLACKPASS = [
    ("BlackPass Match",    "$750-$1,500",    "Match-day round trip with QR ride pass and vehicle priority"),
    ("BlackPass Weekend",  "$2,500-$5,000",  "Multi-day ride credits, concierge coordination, and hotel links"),
    ("BlackPass Executive","$7,500-$15,000", "Priority named vehicle, executive support, and event-week coverage"),
    ("BlackPass Hotel",    "$15,000-$50,000","Branded hotel guest mobility block with dedicated dispatch"),
]


# ============================================================
# Base class with header/footer
# ============================================================
class InvestorPDF(FPDF):
    def __init__(self, doc_title: str = "", doc_subtitle: str = ""):
        super().__init__()
        self._doc_title    = doc_title
        self._doc_subtitle = doc_subtitle

    def header(self):
        # Full-width navy banner
        self.set_fill_color(*NAVY)
        self.rect(0, 0, 210, 28, "F")
        # Gold rule under banner
        self.set_fill_color(*RULE)
        self.rect(0, 28, 210, 1.2, "F")
        # Brand name
        self.set_text_color(*GOLD)
        self.set_font("Helvetica", "B", 15)
        self.set_xy(12, 6)
        self.cell(130, 8, BRAND, ln=False)
        # Document title right-aligned
        self.set_font("Helvetica", "", 9)
        self.set_text_color(*GOLD)
        self.set_xy(100, 6)
        self.cell(98, 8, self._doc_title, align="R", ln=False)
        # Subtitle line
        self.set_xy(12, 15)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(200, 185, 140)
        self.cell(185, 6, self._doc_subtitle, ln=True)
        self.ln(8)

    def footer(self):
        self.set_y(-14)
        self.set_fill_color(*NAVY)
        self.rect(0, self.get_y(), 210, 14, "F")
        self.set_text_color(*GOLD)
        self.set_font("Helvetica", "", 8)
        self.set_x(12)
        self.cell(90, 8, f"Prepared {DATE_PREPARED}  |  CONFIDENTIAL", ln=False)
        self.cell(96, 8, f"Page {self.page_no()} | {CONTACT_URL}", align="R", ln=False)

    def divider(self, top_space: float = 2, bottom_space: float = 2):
        self.ln(top_space)
        self.set_draw_color(*RULE)
        self.set_line_width(0.5)
        self.line(12, self.get_y(), 198, self.get_y())
        self.ln(bottom_space)

    def section_title(self, text: str, numbered: bool = True):
        self.ln(2)
        self.set_fill_color(*LIGHT)
        self.set_text_color(*NAVY)
        self.set_font("Helvetica", "B", 11)
        self.set_x(12)
        self.cell(186, 8, f"  {text}", fill=True, ln=True)
        self.ln(1)
        self.set_text_color(*DARK)

    def body(self, text: str, size: int = 10, h: float = 5.5):
        self.set_font("Helvetica", "", size)
        self.set_text_color(*DARK)
        self.set_x(12)
        self.multi_cell(186, h, text)

    def bullet(self, text: str, size: int = 10):
        self.set_font("Helvetica", "", size)
        self.set_text_color(*MID)
        self.set_x(16)
        self.multi_cell(180, 5.5, f"- {text}")

    def kv_row(self, left: str, right: str, bold_right: bool = False, shaded: bool = False):
        if shaded:
            self.set_fill_color(*LIGHT)
            fill = True
        else:
            fill = False
        self.set_font("Helvetica", "", 10)
        self.set_text_color(*DARK)
        self.set_x(12)
        self.cell(130, 7, left, fill=fill, border=0)
        self.set_font("Helvetica", "B" if bold_right else "", 10)
        self.cell(56, 7, right, fill=fill, border=0, ln=True)

    def table_header(self, cols: list, widths: list):
        self.set_fill_color(*NAVY)
        self.set_text_color(*GOLD)
        self.set_font("Helvetica", "B", 9)
        self.set_x(12)
        for col, w in zip(cols, widths):
            self.cell(w, 7, col, border=1, fill=True)
        self.ln()

    def table_row(self, cells: list, widths: list, shaded: bool = False):
        if shaded:
            self.set_fill_color(*LIGHT)
        else:
            self.set_fill_color(*WHITE)
        self.set_text_color(*DARK)
        self.set_font("Helvetica", "", 9)
        self.set_x(12)
        for val, w in zip(cells, widths):
            self.cell(w, 6, str(val), border=1, fill=True)
        self.ln()

    def cover_block(self, label: str, value: str):
        self.set_font("Helvetica", "B", 22)
        self.set_text_color(*GOLD)
        self.set_x(12)
        self.cell(0, 11, value, ln=True)
        self.set_font("Helvetica", "", 9)
        self.set_text_color(*MID)
        self.set_x(12)
        self.cell(0, 6, label, ln=True)
        self.ln(1)


# ============================================================
# Helpers
# ============================================================
def paragraph(pdf: FPDF, text: str, size: int = 11, spacing: float = 6.0) -> None:
    pdf.set_font("Helvetica", "", size)
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(190, spacing, text)


def section_title(pdf: FPDF, text: str) -> None:
    pdf.ln(1)
    pdf.set_text_color(16, 24, 39)
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(190, 7, text)
    pdf.ln(0.5)


def line_item(pdf: FPDF, left: str, right: str) -> None:
    pdf.set_font("Helvetica", "", 11)
    pdf.set_x(pdf.l_margin)
    pdf.cell(145, 7, left)
    pdf.cell(45, 7, right, ln=True)


def add_title(pdf: FPDF, title: str, subtitle: str) -> None:
    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*GOLD)
    pdf.set_font("Helvetica", "B", 19)
    pdf.set_x(pdf.l_margin)
    pdf.cell(190, 12, title, fill=True, ln=True)
    pdf.set_text_color(*DARK)
    pdf.set_font("Helvetica", "", 11)
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(190, 6, subtitle)
    pdf.ln(3)


def build_lender_pdf() -> None:
    pdf = InvestorPDF(
        doc_title="LENDER PROPOSAL",
        doc_subtitle="Confidential | Atlanta 2026 Premium Event Mobility | $650,000 Funding Facility"
    )
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.set_margins(12, 10, 12)
    pdf.add_page()

    # ---- Cover stat blocks ----
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(*GOLD)
    pdf.set_fill_color(*NAVY)
    pdf.set_x(12)
    pdf.cell(186, 9, "  UNYKORN BLACK | Premium Event Mobility | Atlanta 2026", fill=True, ln=True)
    pdf.ln(4)

    # Stat row
    stats = [("$650,000", "Funding Request"), ("7.2x", "Base DSCR"), ("$1.41M", "Base Revenue"), ("12", "Major Events")]
    col_w = 46
    pdf.set_x(12)
    for val, lbl in stats:
        pdf.set_fill_color(*NAVY)
        pdf.set_text_color(*GOLD)
        pdf.set_font("Helvetica", "B", 18)
        pdf.cell(col_w, 11, val, fill=True, border=0)
    pdf.ln()
    pdf.set_x(12)
    for val, lbl in stats:
        pdf.set_fill_color(*NAVY)
        pdf.set_text_color(200, 185, 140)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(col_w, 7, lbl, fill=True, border=0)
    pdf.ln(8)

    # ---- 1. Executive Summary ----
    pdf.section_title("1  |  Executive Summary")
    pdf.body(
        "UNYKORN BLACK is a premium mobility and event sponsorship platform built for Atlanta's 2026 "
        "event super-cycle. The operating model layers black SUV and sedan dispatch, group-capacity premium "
        "Sprinters, route-ownership sponsorship inventory, merchant referral economics, BlackPass prepaid "
        "ride credits, and TROPTIONS-linked settlement visibility into one vertically integrated business."
    )
    pdf.ln(2)
    pdf.body(
        "The lender proposition is straightforward: this is a collateral-backed transportation operation "
        "with seven distinct revenue streams, conservative DSCR floor of 2.6x in the lean case, and "
        "demonstrated event demand density from May through December 2026. The requested $650,000 facility "
        "funds vehicle acquisition, insurance reserve, compliance, and launch operations through the first "
        "three event windows before sponsor and BlackPass presale revenue normalizes monthly cash flow."
    )
    pdf.ln(2)
    pdf.body(
        "Core risk protections: titled vehicle collateral, staged draw against KPIs, sponsor escrow via "
        "TROPTIONS custody, and partner-fleet overflow that eliminates the need to over-acquire assets "
        "before demand is confirmed. The business is designed to scale conservatively and repay cleanly."
    )

    # ---- 2. Use of Funds ----
    pdf.divider()
    pdf.section_title("2  |  Use of Funds  -  $650,000 Facility")
    pdf.body(
        "The facility is allocated across seven operational categories. Each line is sized to launch-readiness "
        "standards, not aspirational projections. The working capital reserve provides 60-day coverage "
        "while sponsor and BlackPass presale revenues ramp through the first event windows."
    )
    pdf.ln(2)
    pdf.table_header(["Category", "Amount"], [145, 41])
    for i, (item, amt) in enumerate(USE_OF_FUNDS):
        pdf.table_row([item, amt], [145, 41], shaded=(i % 2 == 0))
    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*GOLD)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_x(12)
    pdf.cell(145, 7, "  TOTAL FACILITY", border=1, fill=True)
    pdf.cell(41, 7, "$650,000", border=1, fill=True)
    pdf.ln(6)

    # ---- 3. Scenario Financial View ----
    pdf.divider()
    pdf.section_title("3  |  Revenue Scenarios  -  May through December 2026")
    pdf.body(
        "Three operating scenarios are modeled based on fleet size and sponsorship penetration. All scenarios "
        "assume identical event demand windows and sponsor rate card. Differences reflect dispatch volume, "
        "fleet utilization efficiency, and BlackPass conversion rates."
    )
    pdf.ln(2)
    pdf.table_header(
        ["Scenario", "Fleet", "Gross Revenue", "EBITDA", "Debt Service", "DSCR"],
        [28, 22, 38, 34, 36, 28]
    )
    for i, (name, fleet, rev, ebitda, debt, dscr) in enumerate(SCENARIOS):
        pdf.table_row([name, fleet, rev, ebitda, debt, dscr], [28, 22, 38, 34, 36, 28], shaded=(i % 2 == 0))
    pdf.ln(3)
    pdf.body(
        "DSCR is calculated on base monthly debt service. Base case ($1.41M revenue) reflects 6 owned "
        "vehicles, 60% BlackPass conversion at World Cup and SEC Championship, and 3 of 4 sponsor tiers sold. "
        "The 7.2x base DSCR provides substantial coverage cushion relative to facility cost."
    )

    # ---- 4. Revenue Mix ----
    pdf.divider()
    pdf.section_title("4  |  Revenue Mix  -  Seven Stream Architecture")
    pdf.body(
        "Revenue diversification is a structural protection. No single stream exceeds 43% of total base revenue. "
        "Sponsors (16%) and BlackPass presales (9%) provide advance-committed revenue before each event window. "
        "Partner fleet margin (4%) provides cost-free surge capacity without permanent capex."
    )
    pdf.ln(2)
    pdf.table_header(["Revenue Stream", "Base Projection", "Mix %", "Description"], [42, 32, 16, 96])
    for i, (name, val, pct, note) in enumerate(REVENUE_MIX):
        pdf.table_row([name, val, pct, note], [42, 32, 16, 96], shaded=(i % 2 == 0))
    pdf.ln(4)

    # ---- 5. BlackPass ----
    pdf.section_title("5  |  BlackPass Prepaid Strategy")
    pdf.body(
        "BlackPass converts demand forecasts into committed, non-refundable ride credits before event windows "
        "open. This improves vehicle scheduling confidence, locks sponsor KPI delivery, and provides the "
        "operator with pre-funded operating liquidity before high-load nights. The tiered product ladder "
        "serves both individual VIP buyers and hotel/corporate bulk purchasers."
    )
    pdf.ln(2)
    pdf.table_header(["Product Tier", "Price Range", "Description"], [42, 32, 112])
    for i, (tier, price, desc) in enumerate(BLACKPASS):
        pdf.table_row([tier, price, desc], [42, 32, 112], shaded=(i % 2 == 0))
    pdf.ln(4)

    # ---- Page 2 ----
    pdf.add_page()

    # ---- 6. Sponsor Engine ----
    pdf.section_title("6  |  Sponsor Engine and Commercial Inventory")
    pdf.body(
        "Sponsor revenue ($275,000 base) is tied to measurable, deliverable inventory: named route corridors, "
        "specific vehicles, passenger confirmation touchpoints, and app placement slots. This removes the "
        "ambiguity of traditional brand sponsorship and creates performance-reportable relationships that "
        "renew annually. All sponsor commitments require a signed contract before inventory is held."
    )
    pdf.ln(2)
    for i, (title, price, short_desc, plus_desc) in enumerate(SPONSOR_PACKAGES):
        pdf.set_fill_color(*LIGHT if i % 2 == 0 else WHITE)
        pdf.set_text_color(*NAVY)
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_x(12)
        pdf.cell(145, 7, f"  {title}", fill=True, border=0)
        pdf.set_text_color(*GOLD)
        pdf.set_fill_color(*NAVY)
        pdf.cell(41, 7, price, fill=True, border=0, ln=True)
        pdf.bullet(short_desc, size=9)
        pdf.bullet(plus_desc, size=9)
        pdf.ln(1)

    # ---- 7. Collateral + Risk ----
    pdf.divider()
    pdf.section_title("7  |  Collateral Position and Risk Controls")
    pdf.body(
        "The lender holds six layers of operational risk protection. These are structural features of the "
        "launch model, not post-funding commitments. Each layer is independently verifiable during underwriting."
    )
    pdf.ln(2)
    controls = [
        ("Vehicle Collateral",        "Titled SUVs and Sprinters on core owned units provide first-priority security interest."),
        ("Account-Backed Dispatch",   "Corporate and executive accounts pre-book route blocks, reducing ad-hoc revenue dependence."),
        ("Presale Demand (BlackPass)", "Prepaid credits and sponsor escrow via TROPTIONS committed before peak event windows."),
        ("Partner Fleet Overflow",    "Surge capacity delivered through pre-contracted partner vehicles - no capex required."),
        ("Insurance Reserve",         "$45,000 dedicated insurance buffer maintained outside operating working capital."),
        ("Settlement Audit Trail",    "TROPTIONS provides transaction-level reporting for every ride, sponsor dollar, and referral."),
    ]
    for i, (label, detail) in enumerate(controls):
        pdf.set_fill_color(*LIGHT if i % 2 == 0 else WHITE)
        pdf.set_text_color(*NAVY)
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_x(12)
        pdf.cell(45, 6, f"  {label}", fill=True, border=1)
        pdf.set_text_color(*DARK)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_fill_color(*LIGHT if i % 2 == 0 else WHITE)
        pdf.cell(141, 6, detail, fill=True, border=1, ln=True)
    pdf.ln(4)

    # ---- 8. Event Calendar ----
    pdf.section_title("8  |  Atlanta Event Demand Calendar  -  May-Dec 2026")
    pdf.body(
        "The following 12 confirmed event windows drive dispatch scheduling, sponsor targeting, BlackPass "
        "conversion, and partner fleet activation. Tier A+ events represent $400,000+ concentrated demand "
        "windows. World Cup (June-July) and SEC Championship (December) are the two primary anchor events."
    )
    pdf.ln(2)
    pdf.table_header(["Date", "Event", "Venue", "Tier"], [30, 52, 60, 14])
    for i, (date, event, venue, tier, angle) in enumerate(EVENTS):
        pdf.table_row([date, event, venue, tier], [30, 52, 60, 14], shaded=(i % 2 == 0))
        # Revenue angle as sub-row
        pdf.set_fill_color(250, 250, 250)
        pdf.set_text_color(*MID)
        pdf.set_font("Helvetica", "I", 8)
        pdf.set_x(12)
        pdf.cell(156, 5, f"  Revenue angle: {angle}", border="B", fill=True, ln=True)
    pdf.ln(4)

    # ---- 9. WhichWay ----
    pdf.section_title("9  |  WhichWay Ecosystem Integration")
    pdf.body(
        "UNYKORN BLACK is deployed within the WhichWay event operating stack at https://fifa.unykorn.org/. "
        "This is operationally significant for lenders: the mobility layer does not operate in isolation. "
        "Guest flows, venue maps, merchant referrals, emergency routing, and sponsor dashboards are all "
        "connected. This increases sponsor retention value and creates data-backed reporting that supports "
        "loan covenant monitoring."
    )
    pdf.ln(2)
    ecosystem = [
        ("Guest OS",          "Route and service layer for complete guest transport planning and navigation."),
        ("Demo + Passport",   "Engagement and loyalty tied to trip behavior; increases BlackPass repeat conversion."),
        ("Venue Map Module",  "Destination routing, parking, and rideshare coordination embedded in event flows."),
        ("Merchant Module",   "QR-driven referral commerce; each confirmed ride generates merchant exposure."),
        ("Emergency Module",  "Guest safety escalation and support workflows; reduces operator liability exposure."),
        ("Sales + Signup",    "Sponsor and account pipeline capture integrated into every event page."),
    ]
    for i, (mod, desc) in enumerate(ecosystem):
        pdf.set_fill_color(*LIGHT if i % 2 == 0 else WHITE)
        pdf.set_text_color(*NAVY)
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_x(12)
        pdf.cell(40, 6, f"  {mod}", fill=True, border=1)
        pdf.set_text_color(*DARK)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_fill_color(*LIGHT if i % 2 == 0 else WHITE)
        pdf.cell(146, 6, desc, fill=True, border=1, ln=True)
    pdf.ln(4)

    # ---- 10. Underwriting ----
    pdf.divider()
    pdf.section_title("10  |  Underwriting Summary and Recommendation")
    pdf.body(
        "UNYKORN BLACK presents a practical lender profile: titled vehicle collateral, seven-stream diversified "
        "revenue, 7.2x base DSCR, advance-committed sponsor and presale demand, and a structured operating "
        "framework that avoids the primary failure mode of transportation startups - over-capitalizing before "
        "demand is proven."
    )
    pdf.ln(2)
    pdf.body(
        "Recommended structure: staged facility release with KPI triggers. Tranche 1 ($250,000) on execution "
        "for vehicle acquisition and insurance reserve. Tranche 2 ($250,000) on first sponsor contract "
        "execution and BlackPass presale confirmation. Tranche 3 ($150,000) post-World Cup event window "
        "performance proof. Monthly financial reporting with TROPTIONS-linked transaction audit."
    )
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(*NAVY)
    pdf.set_x(12)
    pdf.cell(186, 8, f"  Prepared for lender/credit committee review  |  {DATE_PREPARED}  |  {CONTACT_URL}", fill=True, ln=True)

    pdf.output(str(DOWNLOADS / "unykorn-black-lender-proposal.pdf"))


def build_rate_card_pdf() -> None:
    pdf = InvestorPDF(
        doc_title="SPONSOR RATE CARD",
        doc_subtitle="Confidential | Commercial Inventory Packages | Atlanta 2026 Event Season"
    )
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.set_margins(12, 10, 12)
    pdf.add_page()

    # Intro
    pdf.ln(2)
    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*GOLD)
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_x(12)
    pdf.cell(186, 9, "  UNYKORN BLACK | Sponsor Commercial Inventory | Atlanta 2026", fill=True, ln=True)
    pdf.ln(4)
    pdf.body(
        "UNYKORN BLACK offers measurable, performance-linked sponsor inventory across four product tiers. "
        "Each tier is backed by verified ride volume, branded passenger touchpoints, route ownership, "
        "and monthly performance reporting. Inventory is limited and sold by commitment date."
    )

    # ---- Package Menu ----
    pdf.divider()
    pdf.section_title("Sponsor Package Menu")
    for i, (title, price, short_desc, plus_desc) in enumerate(SPONSOR_PACKAGES):
        # Package header bar
        pdf.set_fill_color(*NAVY)
        pdf.set_text_color(*GOLD)
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_x(12)
        pdf.cell(145, 9, f"  {title}", fill=True, border=0)
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(41, 9, price, fill=True, border=0, align="R", ln=True)
        # Description
        pdf.set_text_color(*DARK)
        pdf.set_font("Helvetica", "", 10)
        pdf.set_x(14)
        pdf.multi_cell(182, 5.5, short_desc)
        pdf.set_x(14)
        pdf.set_text_color(*MID)
        pdf.set_font("Helvetica", "I", 9)
        pdf.multi_cell(182, 5, plus_desc)
        pdf.ln(4)

    # ---- Inventory Breakdown ----
    pdf.divider()
    pdf.section_title("What Sponsors Own")
    inventory = [
        ("Vehicle Branding",      "Exterior wrap + interior brand materials on named fleet unit for the contracted period."),
        ("Route Ownership",       "Named corridor displayed in app, confirmations, and QR receipts. Exclusive per tier."),
        ("App Placement",         "Sponsor creative shown in booking flows, confirmation screens, and post-ride summaries."),
        ("QR Offer Slot",         "Branded offer code embedded in passenger QR handoff at start and end of each ride."),
        ("Event Activation",      "Branded presence at Tier A/A+ events; sponsor materials in VIP vehicle areas."),
        ("Merchant Referral Link","Tracked link placement for hotels, restaurants, and venues within the sponsor package."),
    ]
    for i, (item, detail) in enumerate(inventory):
        pdf.set_fill_color(*LIGHT if i % 2 == 0 else WHITE)
        pdf.set_text_color(*NAVY)
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_x(12)
        pdf.cell(45, 6, f"  {item}", fill=True, border=1)
        pdf.set_text_color(*DARK)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_fill_color(*LIGHT if i % 2 == 0 else WHITE)
        pdf.cell(141, 6, detail, fill=True, border=1, ln=True)
    pdf.ln(4)

    # ---- Reporting Standards ----
    pdf.section_title("Reporting and Proof Standards")
    pdf.body(
        "Every sponsor tier receives documented proof of performance. Reporting is provided monthly "
        "with quarterly review calls for Founding and Route sponsors. Settlement proof is generated "
        "through TROPTIONS-linked transaction logging."
    )
    pdf.ln(2)
    reporting = [
        "Monthly sponsor performance summary: ride counts, placement logs, QR scan rates.",
        "Route and event coverage documentation with timestamps and GPS confirmation.",
        "Creative proof snapshots (photo + screen capture) per event window delivered within 5 days.",
        "Quarterly sponsor ROI review with renewal pricing and inventory availability update.",
        "TROPTIONS settlement export: every sponsor dollar traced to confirmed ride delivery.",
    ]
    for r in reporting:
        pdf.bullet(r)
    pdf.ln(4)

    # ---- Event Coverage ----
    pdf.section_title("2026 Event Coverage Windows")
    pdf.body(
        "Sponsor inventory is active across all 12 confirmed Atlanta events May-December 2026. "
        "Tier A+ events (World Cup window, SEC Championship) are primary activation targets with "
        "concentrated corporate and VIP traffic. All sponsor materials must be delivered and approved "
        "14 days before first event activation."
    )
    pdf.ln(2)
    pdf.table_header(["Date", "Event", "Tier", "Sponsor Value"], [30, 80, 12, 64])
    for i, (date, event, venue, tier, angle) in enumerate(EVENTS):
        label = "High" if tier in ("A+",) else ("Strong" if tier == "A" else "Moderate")
        pdf.table_row([date, event, tier, label + " - " + angle[:40] + ("..." if len(angle) > 40 else "")],
                      [30, 80, 12, 64], shaded=(i % 2 == 0))
    pdf.ln(4)

    # ---- Commercial Contact ----
    pdf.divider()
    pdf.section_title("How to Reserve Your Inventory")
    steps = [
        "Select your package tier and preferred event window(s).",
        "Contact UNYKORN BLACK partnerships team to confirm inventory availability.",
        "Execute commercial term sheet and provide prepayment to TROPTIONS escrow custody.",
        "Submit brand assets (logo, creative, QR offer content) at least 14 days before activation.",
        "Receive monthly performance reports from first event window through contract end.",
    ]
    for i, step in enumerate(steps, 1):
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(*DARK)
        pdf.set_x(12)
        pdf.multi_cell(186, 5.5, f"  {i}.  {step}")
    pdf.ln(3)
    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*GOLD)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_x(12)
    pdf.cell(186, 8, f"  UNYKORN BLACK Partnerships  |  {CONTACT_URL}  |  Prepared {DATE_PREPARED}", fill=True, ln=True)

    pdf.output(str(DOWNLOADS / "unykorn-black-sponsor-rate-card.pdf"))


def build_docx() -> None:
    doc = Document()

    # Title
    title = doc.add_heading("UNYKORN BLACK", level=1)
    title.runs[0].font.color.rgb = RGBColor(7, 17, 31)
    subtitle = doc.add_heading("Premium Event Mobility | Atlanta 2026 | Investor Funding Packet", level=2)
    doc.add_paragraph(
        f"Prepared: {DATE_PREPARED}  |  CONFIDENTIAL  |  {CONTACT_URL}\n"
        "This document is prepared for lenders, credit committees, funding groups, and strategic partners "
        "requiring an editable underwriting summary. All financial projections reflect the Atlanta 2026 "
        "event calendar. Three scenarios are modeled; Base case is the primary underwriting assumption."
    )
    doc.add_paragraph("")

    # --- Funding Request ---
    doc.add_heading("Funding Request", level=2)
    doc.add_paragraph(
        "$650,000 controlled launch facility for a 5-8 vehicle premium event fleet. "
        "Funds are staged across vehicle acquisition, insurance reserve, compliance, "
        "dispatch buildout, branding, and working capital. The facility is structured "
        "for staged release tied to sponsor pre-sales and dispatch KPI verification."
    )

    # --- Executive Narrative ---
    doc.add_heading("Executive Summary", level=2)
    doc.add_paragraph(
        "UNYKORN BLACK is a premium mobility and event sponsorship platform built for Atlanta's "
        "2026 event super-cycle. The operating model layers black SUV and sedan dispatch, "
        "group-capacity Sprinters, route-ownership sponsor inventory, merchant referral economics, "
        "BlackPass prepaid ride credits, and TROPTIONS-linked settlement reporting into one integrated business."
    )
    doc.add_paragraph(
        "The lender proposition is clean: collateral-backed transportation with seven revenue streams, "
        "2.6x floor DSCR in the lean case, 7.2x in the base, and 12 confirmed major Atlanta events "
        "from May through December 2026. Sponsor escrow, BlackPass presales, and account-backed dispatch "
        "provide committed revenue before peak windows open. Partner fleet overflow eliminates the need "
        "to over-acquire assets before demand is proven."
    )

    # --- Scenario Table ---
    doc.add_heading("Revenue Scenarios", level=2)
    table = doc.add_table(rows=1, cols=6)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    for i, h in enumerate(["Scenario", "Fleet", "Gross Revenue", "EBITDA", "Debt Service", "DSCR"]):
        hdr[i].text = h
    for row in SCENARIOS:
        c = table.add_row().cells
        for i, val in enumerate(row):
            c[i].text = val
    doc.add_paragraph(
        "DSCR calculated on base monthly debt service. Base case assumes 6 vehicles, "
        "60% BlackPass conversion at World Cup and SEC Championship, and 3 of 4 sponsor tiers sold."
    )

    # --- Revenue Mix ---
    doc.add_heading("Revenue Mix - Seven Stream Architecture", level=2)
    doc.add_paragraph(
        "No single revenue stream exceeds 43% of total base revenue. Sponsors (16%) and "
        "BlackPass presales (9%) provide advance-committed revenue before each event window."
    )
    mix = doc.add_table(rows=1, cols=4)
    mix.style = "Table Grid"
    mh = mix.rows[0].cells
    for i, h in enumerate(["Stream", "Base Value", "Mix %", "Description"]):
        mh[i].text = h
    for name, value, pct, note in REVENUE_MIX:
        c = mix.add_row().cells
        c[0].text = name
        c[1].text = value
        c[2].text = pct
        c[3].text = note

    # --- Use of Funds ---
    doc.add_heading("Use of Funds - $650,000 Breakdown", level=2)
    uof = doc.add_table(rows=1, cols=2)
    uof.style = "Table Grid"
    uh = uof.rows[0].cells
    uh[0].text = "Category"
    uh[1].text = "Amount"
    for item, amount in USE_OF_FUNDS:
        c = uof.add_row().cells
        c[0].text = item
        c[1].text = amount
    total_row = uof.add_row().cells
    total_row[0].text = "TOTAL FACILITY"
    total_row[1].text = "$650,000"

    # --- BlackPass ---
    doc.add_heading("BlackPass Presale Strategy", level=2)
    doc.add_paragraph(
        "BlackPass converts demand forecasts into committed, non-refundable ride credits "
        "before event windows open. The tiered product ladder serves individual VIP buyers "
        "and hotel/corporate bulk purchasers. Presales improve dispatch confidence and "
        "provide pre-funded operating liquidity before high-load event nights."
    )
    for tier, price, desc in BLACKPASS:
        doc.add_paragraph(f"{tier} ({price}): {desc}", style="List Bullet")

    # --- Sponsor Packages ---
    doc.add_heading("Sponsor Package Structure", level=2)
    doc.add_paragraph(
        "Sponsor revenue ($275,000 base) is tied to measurable, deliverable inventory. "
        "All sponsor commitments require a signed contract before inventory is held. "
        "Monthly performance reporting included with every tier."
    )
    for title, price, short_desc, plus_desc in SPONSOR_PACKAGES:
        doc.add_paragraph(f"{title} - {price}", style="List Bullet")
        doc.add_paragraph(f"  {short_desc}")
        doc.add_paragraph(f"  {plus_desc}")

    # --- Collateral ---
    doc.add_heading("Collateral and Risk Controls", level=2)
    controls = [
        "Vehicle collateral: titled SUVs and Sprinters on core owned units provide first-priority security interest.",
        "Account-backed dispatch: corporate and executive accounts pre-book route blocks, reducing ad-hoc revenue dependence.",
        "Presale demand: BlackPass credits and sponsor escrow via TROPTIONS committed before peak windows.",
        "Partner fleet overflow: surge capacity through pre-contracted partners - no additional capex required.",
        "Insurance reserve: $45,000 dedicated buffer maintained outside operating working capital.",
        "Settlement audit trail: TROPTIONS transaction-level reporting for every ride, sponsor dollar, and referral.",
    ]
    for c in controls:
        doc.add_paragraph(c, style="List Bullet")

    # --- Event Calendar ---
    doc.add_heading("Atlanta Event Demand Calendar - May-Dec 2026", level=2)
    doc.add_paragraph(
        "12 confirmed event windows drive dispatch scheduling, sponsor targeting, and BlackPass conversion. "
        "Tier A+ events (World Cup, SEC Championship) are primary anchor periods with concentrated "
        "corporate and VIP traffic."
    )
    evt = doc.add_table(rows=1, cols=5)
    evt.style = "Table Grid"
    eh = evt.rows[0].cells
    for i, h in enumerate(["Date", "Event", "Venue", "Tier", "Revenue Angle"]):
        eh[i].text = h
    for date, event, venue, tier, angle in EVENTS:
        c = evt.add_row().cells
        c[0].text = date
        c[1].text = event
        c[2].text = venue
        c[3].text = tier
        c[4].text = angle

    # --- WhichWay ---
    doc.add_heading("WhichWay Platform Integration", level=2)
    doc.add_paragraph(
        "UNYKORN BLACK operates within the WhichWay event OS at https://fifa.unykorn.org/. "
        "The mobility layer connects with guest flow, venue maps, merchant referrals, "
        "emergency routing, and sponsor dashboards. This increases sponsor retention value "
        "and creates auditable data for loan covenant monitoring."
    )
    doc.add_paragraph(
        "Live modules: /guest, /demo, /passport, /map, /merchants, /emergency, /sign-up, /sales. "
        "Powered by TROPTIONS settlement reporting and Apostle Chain settlement infrastructure."
    )

    # --- Underwriting ---
    doc.add_heading("Underwriting Recommendation", level=2)
    doc.add_paragraph(
        "Recommended facility structure: staged release with KPI triggers. "
        "Tranche 1 ($250,000) on execution for vehicle acquisition and insurance reserve. "
        "Tranche 2 ($250,000) on first sponsor contract and BlackPass presale confirmation. "
        "Tranche 3 ($150,000) post-World Cup event window performance proof. "
        "Monthly financial reporting with TROPTIONS transaction audit included."
    )
    doc.add_paragraph(f"Prepared for lender review | {DATE_PREPARED} | {CONTACT_URL}")

    doc.save(str(DOWNLOADS / "unykorn-black-lender-proposal.docx"))


def build_csv() -> None:
    csv_path = DOWNLOADS / "unykorn-black-event-demand-calendar.csv"
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["date", "event", "venue", "tier", "revenue_angle"])
        writer.writerows(EVENTS)


def build_troptions_settlement_pdf() -> None:
    pdf = InvestorPDF(
        doc_title="TROPTIONS SETTLEMENT",
        doc_subtitle="POF Escrow | Multi-Asset Clearing | Issuer Opportunity | UNYKORN BLACK + WhichWay Network"
    )
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.set_margins(12, 10, 12)
    pdf.add_page()

    # Intro banner
    pdf.ln(2)
    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*GOLD)
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_x(12)
    pdf.cell(186, 9, "  TROPTIONS | Settlement Backbone + Issuer Opportunity for UNYKORN BLACK", fill=True, ln=True)
    pdf.ln(4)

    # Stats
    tstats = [("POF Escrow", "USDC Custody"), ("ATP Chain", "Settlement Token"), ("2%", "Settlement Fee"), ("Daily", "Settlement Cadence")]
    col_w = 46
    pdf.set_x(12)
    for val, lbl in tstats:
        pdf.set_fill_color(*NAVY)
        pdf.set_text_color(*GOLD)
        pdf.set_font("Helvetica", "B", 14)
        pdf.cell(col_w, 11, val, fill=True)
    pdf.ln()
    pdf.set_x(12)
    for val, lbl in tstats:
        pdf.set_fill_color(*NAVY)
        pdf.set_text_color(200, 185, 140)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(col_w, 7, lbl, fill=True)
    pdf.ln(8)

    # ---- 1. Executive Summary ----
    pdf.section_title("1  |  The Settlement Backbone")
    pdf.body(
        "TROPTIONS is the settlement infrastructure layer enabling UNYKORN BLACK and the broader "
        "WhichWay event ecosystem. It provides Proof of Funds (POF) custody via USDC escrow, "
        "multi-asset clearing across ATP and stablecoin rails, transaction-level audit reporting, "
        "and tokenized issuer pathways for sponsors and corporate partners who want equity upside "
        "in addition to brand exposure."
    )
    pdf.ln(2)
    pdf.body(
        "This document describes the commercial opportunity for sponsors, lenders, and issuers: "
        "how the settlement engine works, what the POF/escrow framework covers, and how sponsors "
        "can transition from a flat fee into an equity or performance token model that earns "
        "proportionally to actual UNYKORN BLACK ride and referral volume."
    )

    # ---- 2. What Is TROPTIONS ----
    pdf.divider()
    pdf.section_title("2  |  What TROPTIONS Provides")
    features = [
        ("POF Escrow Custody",        "Sponsor prepayments held in USDC escrow from contract signature to KPI delivery. No operational mixing."),
        ("Multi-Asset Settlement",    "USDC stablecoin for fiat-equivalent clearing, ATP (Apostle Chain) for tokenized flows and equity distributions."),
        ("Transaction Reporting",     "Every ride, sponsorship dollar, merchant referral, and settlement event is logged with timestamp and chain ID."),
        ("Issuer Tokenization",       "Sponsors can mint equity tokens or performance tokens on Apostle Chain (chain_id 7332) backed by ride revenue."),
        ("WhichWay Integration",      "Guest, map, merchant, and emergency modules all report settlement state; sponsors see live KPI dashboards."),
        ("Audit + Compliance",        "All records are immutable on Apostle Chain, cross-referenced with WhichWay KPI data for regulatory reporting."),
    ]
    for i, (label, detail) in enumerate(features):
        pdf.set_fill_color(*LIGHT if i % 2 == 0 else WHITE)
        pdf.set_text_color(*NAVY)
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_x(12)
        pdf.cell(45, 6, f"  {label}", fill=True, border=1)
        pdf.set_text_color(*DARK)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_fill_color(*LIGHT if i % 2 == 0 else WHITE)
        pdf.cell(141, 6, detail, fill=True, border=1, ln=True)
    pdf.ln(4)

    # ---- 3. UNYKORN BLACK Use Case ----
    pdf.section_title("3  |  UNYKORN BLACK Use Case: Why Settlement Matters")
    pdf.body(
        "UNYKORN BLACK operates with three categories of pre-committed revenue: BlackPass presales, "
        "sponsor contracts, and corporate account blocks. All three require custody, reporting, and "
        "transparent delivery confirmation before funds release. TROPTIONS provides each:"
    )
    pdf.ln(2)
    cases = [
        ("BlackPass Custody",   "Presale ride credits held in escrow until dispatched. Credit can be refunded or transferred on 48hr notice."),
        ("Sponsor Escrow",      "Signed sponsor contracts lock sponsor payment in USDC. Funds release against verified KPI milestones."),
        ("Account Clearing",    "Corporate account blocks settled weekly via TROPTIONS ATP rails with PDF invoice and chain transaction proof."),
        ("KPI Verification",    "Ride count, QR scans, route coverage, and event window data pulled from WhichWay and cross-matched to settlement."),
    ]
    for i, (label, detail) in enumerate(cases):
        pdf.set_fill_color(*LIGHT if i % 2 == 0 else WHITE)
        pdf.set_text_color(*NAVY)
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_x(12)
        pdf.cell(45, 6, f"  {label}", fill=True, border=1)
        pdf.set_text_color(*DARK)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_fill_color(*LIGHT if i % 2 == 0 else WHITE)
        pdf.cell(141, 6, detail, fill=True, border=1, ln=True)
    pdf.ln(4)

    # ---- 4. POF Framework ----
    pdf.section_title("4  |  Proof of Funds and Escrow Framework")
    pdf.body(
        "Proof of Funds is the cornerstone of sponsor confidence and lender certainty. The TROPTIONS "
        "custody model is designed to create a clean chain-of-custody from sponsor commitment to "
        "verified delivery. No sponsor capital touches operational working capital until KPIs are met."
    )
    pdf.ln(2)
    pof = [
        "Sponsor prepayment deposited to USDC escrow wallet at contract signature - no mixing with operating float.",
        "Escrow release triggers: minimum ride volume confirmed, route coverage logged, QR scan threshold met.",
        "Monthly reconciliation report generated and delivered to sponsor, lender, and operator simultaneously.",
        "Full chain-of-custody: every transaction ID, block height, timestamp, and settlement proof archived.",
        "Disputed releases handled by 3-party review: operator, sponsor, and TROPTIONS escrow officer.",
    ]
    for p in pof:
        pdf.bullet(p)
    pdf.ln(4)

    # ---- Page 2 ----
    pdf.add_page()

    pdf.section_title("5  |  Sponsor as Issuer: Equity and Performance Token Models")
    pdf.body(
        "A sponsor can choose to structure their UNYKORN BLACK participation not as a flat fee but "
        "as a tokenized equity or performance stake. This upgrades the commercial relationship from "
        "a marketing spend to an investment with measurable, liquid upside tied directly to ride "
        "volume and event performance."
    )

    # 5a Equity
    pdf.divider(top_space=3, bottom_space=2)
    pdf.section_title("5a  |  Equity Issuance Model")
    eq_rows = [
        ("Investment",    "$150,000 Founding Sponsor capital"),
        ("Token Mint",    "Tokenized equity on Apostle Chain (chain_id 7332)"),
        ("Revenue Share", "5-10% revenue participation on attributed sponsor routes"),
        ("Settlement",    "Monthly ATP settlement to sponsor wallet based on ride attribution"),
        ("Liquidity",     "Tokens tradeable on Apostle marketplace from day of mint"),
        ("Reporting",     "Daily KPI dashboard, monthly statement, quarterly tax export"),
    ]
    pdf.table_header(["Term", "Detail"], [50, 136])
    for i, (k, v) in enumerate(eq_rows):
        pdf.table_row([k, v], [50, 136], shaded=(i % 2 == 0))
    pdf.ln(3)
    pdf.body(
        "Equity model example: A Founding Sponsor investing $150,000 receives tokens representing "
        "7.5% of UNYKORN BLACK net revenue attributed to their branded routes and vehicles. "
        "At base-case revenue of $1.41M, attributed portion (assuming 20% of revenue through sponsor routes) "
        "= $282,000 x 7.5% = $21,150 annual ATP distribution. Token value appreciates with fleet scale."
    )
    pdf.ln(4)

    # 5b Performance
    pdf.divider(top_space=3, bottom_space=2)
    pdf.section_title("5b  |  Performance Token Model (Pay-Per-Ride Upside)")
    pdf.body(
        "The performance token model requires a smaller upfront commitment ($25K-$50K) and earns "
        "proportionally to actual ride volume and merchant referrals attributed to the sponsor's "
        "branded inventory. No lockup period. Earnings distributed quarterly in ATP."
    )
    pdf.ln(2)
    perf_rows = [
        ("Token Mint",        "1,000 tokens at $25 each ($25,000 upfront)"),
        ("Per-Ride Earn",     "0.5 ATP per confirmed ride using sponsor route or vehicle"),
        ("Referral Earn",     "0.1 ATP per qualified merchant lead attributed to sponsor QR"),
        ("Distribution",      "Quarterly ATP settlement to sponsor Apostle Chain wallet"),
        ("Compounding",       "Token holders can re-stake ATP earnings for additional allocation"),
        ("Transferability",   "Tokens tradeable on Apostle marketplace; no transfer lockup"),
    ]
    pdf.table_header(["Term", "Detail"], [50, 136])
    for i, (k, v) in enumerate(perf_rows):
        pdf.table_row([k, v], [50, 136], shaded=(i % 2 == 0))
    pdf.ln(4)

    # ---- 6. Issuer Commercial Terms ----
    pdf.section_title("6  |  Issuer Commercial Terms")
    pdf.body(
        "TROPTIONS charges a 2% settlement fee on all sponsor payouts (equity and performance token "
        "distributions). All other fees are covered by the operator (UNYKORN BLACK). Sponsors own "
        "their token wallets with full non-custodial key control."
    )
    pdf.ln(2)
    terms = [
        ("Setup",            "Registry entry + Apostle Chain wallet provision + token contract mint (one-time)"),
        ("Settlement",       "Daily ride/referral aggregation, monthly ATP calculation and payout"),
        ("Custody",          "Non-custodial Apostle Chain wallets; sponsors hold private keys"),
        ("Reporting",        "Daily KPI dashboard, monthly statement, quarterly tax-ready export"),
        ("Fee",              "2% of distributed ATP settlement amount per period"),
        ("Transferability",  "Tokens tradeable on Apostle marketplace; no lockup after mint"),
    ]
    pdf.table_header(["Term", "Detail"], [40, 146])
    for i, (k, v) in enumerate(terms):
        pdf.table_row([k, v], [40, 146], shaded=(i % 2 == 0))
    pdf.ln(4)

    # ---- 7. Compliance ----
    pdf.section_title("7  |  Compliance and Audit Framework")
    pdf.body(
        "TROPTIONS settlement infrastructure is built for regulatory audit. All settlement flows "
        "are recorded immutably on the Apostle Chain and cross-referenced with WhichWay KPI data. "
        "The following compliance frameworks are addressed:"
    )
    pdf.ln(2)
    comp_rows = [
        ("Securities",    "Tokenized equity and revenue participation reviewed against SEC exemption frameworks (Reg D / Reg A)"),
        ("AML/KYC",       "Sponsor identities verified at token mint; settlement flows logged and reported per FinCEN guidance"),
        ("Tax Reporting", "Monthly settlement statements capture realized income for 1099/K1 and foreign tax reporting"),
        ("Audit Trail",   "Every ride, route placement, referral, and settlement proof traceable to source on Apostle Chain"),
        ("Data Privacy",  "KYC data held in GDPR-compliant encrypted storage; chain records contain only wallet IDs"),
    ]
    pdf.table_header(["Domain", "Coverage"], [40, 146])
    for i, (k, v) in enumerate(comp_rows):
        pdf.table_row([k, v], [40, 146], shaded=(i % 2 == 0))
    pdf.ln(4)

    # ---- 8. Why Sponsors Choose ----
    pdf.section_title("8  |  Why Sponsors Choose TROPTIONS + UNYKORN BLACK")
    advantages = [
        "No middleman: direct non-custodial token ownership; sponsor holds private key from day one.",
        "Real revenue data: ROI measured in confirmed rides, QR scans, and merchant conversions - not impressions.",
        "Liquid asset class: tokens tradeable on Apostle marketplace with no lock-up period after mint.",
        "Network effects: each new sponsor adds route inventory that increases the value of existing tokens.",
        "Ecosystem scaling: WhichWay modules (guest, demo, passport, merchants) all drive sponsor token economics.",
        "Lender confidence: escrow custody and audit trail reduce operator risk profile, supporting facility terms.",
    ]
    for a in advantages:
        pdf.bullet(a)
    pdf.ln(4)

    # ---- 9. Next Steps ----
    pdf.section_title("9  |  Next Steps: Becoming an Issuer")
    steps = [
        "Complete issuer KYC application and execute commercial term sheet with TROPTIONS.",
        "Provide capital commitment and select token model (equity issuance vs. performance tokens).",
        "TROPTIONS provisions Apostle Chain wallet and deploys token contract within 5 business days.",
        "Deliver brand assets and KPI targets to UNYKORN BLACK at least 14 days before first activation.",
        "First ATP settlement distributed within 30 days of first confirmed rides under sponsor attribution.",
    ]
    for i, step in enumerate(steps, 1):
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(*DARK)
        pdf.set_x(12)
        pdf.multi_cell(186, 5.5, f"  {i}.  {step}")
    pdf.ln(3)

    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*GOLD)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_x(12)
    pdf.cell(186, 8, f"  TROPTIONS Settlement Platform  |  {CONTACT_URL}  |  Prepared {DATE_PREPARED}", fill=True, ln=True)

    pdf.output(str(DOWNLOADS / "unykorn-black-troptions-settlement.pdf"))


# ============================================================
# Executive One-Pager
# ============================================================
def build_executive_one_pager() -> None:
    pdf = InvestorPDF(
        doc_title="Executive One-Pager",
        doc_subtitle="Print-ready summary for credit committee, private lenders, and strategic partners"
    )
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()

    # Hero stat row
    stats = [
        ("$650K",    "Facility\nRequested"),
        ("$1.41M",   "Base Revenue\nMay-Dec 2026"),
        ("7.2x",     "Base-Case\nDSCR"),
        ("6",        "Premium\nVehicles"),
        ("12",       "A/A+ Tier\nEvents"),
        ("$275K",    "Sponsor\nRevenue"),
    ]
    col_w = 186 / len(stats)
    pdf.set_fill_color(*NAVY)
    pdf.rect(12, pdf.get_y(), 186, 22, "F")
    for i, (val, lbl) in enumerate(stats):
        x = 12 + i * col_w
        pdf.set_xy(x, pdf.get_y() + 2)
        pdf.set_text_color(*GOLD)
        pdf.set_font("Helvetica", "B", 13)
        pdf.cell(col_w, 7, val, align="C", ln=False)
    pdf.ln(9)
    for i, (val, lbl) in enumerate(stats):
        x = 12 + i * col_w
        pdf.set_xy(x, pdf.get_y() + 1)
        pdf.set_text_color(200, 185, 140)
        pdf.set_font("Helvetica", "", 7)
        for line in lbl.split("\n"):
            pdf.set_xy(x, pdf.get_y())
            pdf.cell(col_w, 4, line, align="C", ln=False)
        pdf.set_xy(12 + (i + 1) * col_w, pdf.get_y() - 4)
    pdf.ln(8)
    pdf.divider()

    # What is UNYKORN BLACK
    pdf.section_title("What Is UNYKORN BLACK?")
    pdf.body(
        "UNYKORN BLACK is a premium event mobility and sponsorship platform launching in Atlanta for the 2026 "
        "event super-cycle. The business operates black SUVs, executive sedans, and premium group vans across "
        "World Cup, SEC Championship, TOUR Championship, Dragon Con, and 12 other Tier-A event windows. "
        "Revenue combines dispatched fare income with presold BlackPass ride credits, branded sponsor route "
        "inventory, hotel-corporate account blocks, and merchant referral activation."
    )
    pdf.divider()

    # The ask
    pdf.section_title("The Funding Request")
    items = [
        ("Total facility",                     "$650,000",   True),
        ("Vehicle acquisition / down payments","$115,000",   False),
        ("Equipment financing facility",        "$350,000",  False),
        ("Insurance reserve (12-month buffer)", "$45,000",   False),
        ("Licensing, compliance, legal, CPA",   "$25,000",   False),
        ("Website, app, dispatch buildout",     "$35,000",   False),
        ("Branding, QR codes, sponsor materials","$20,000",  False),
        ("Working capital reserve",             "$60,000",   False),
    ]
    for label, amount, bold in items:
        pdf.set_fill_color(*LIGHT) if not bold else pdf.set_fill_color(*NAVY)
        pdf.set_text_color(*DARK) if not bold else pdf.set_text_color(*GOLD)
        pdf.set_font("Helvetica", "B" if bold else "", 10)
        pdf.set_x(12)
        pdf.cell(146, 7, f"  {label}", fill=True, border="B")
        pdf.cell(40, 7, amount, align="R", fill=True, border="B", ln=True)
    pdf.ln(2)
    pdf.divider()

    # Scenario table
    pdf.section_title("3-Scenario Financial View")
    headers = ["Scenario", "Fleet", "Gross Revenue", "EBITDA", "Debt Service", "DSCR"]
    widths  = [28, 24, 36, 32, 34, 32]
    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*GOLD)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_x(12)
    for h, w in zip(headers, widths):
        pdf.cell(w, 7, h, fill=True, border=1)
    pdf.ln()
    for i, (name, fleet, rev, ebitda, ds, dscr) in enumerate(SCENARIOS):
        shade = (i % 2 == 0)
        pdf.set_fill_color(*LIGHT) if shade else pdf.set_fill_color(*WHITE)
        pdf.set_text_color(*DARK)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_x(12)
        for val, w in zip([name, fleet, rev, ebitda, ds, dscr], widths):
            pdf.cell(w, 7, val, fill=True, border=1)
        pdf.ln()
    pdf.ln(2)
    pdf.divider()

    # Why lenders say yes
    pdf.section_title("Why Lenders Say Yes")
    reasons = [
        "Hard asset collateral: titled vehicle fleet with clear acquisition records and insurance reserve",
        "Prepaid demand: BlackPass presales and sponsor contracts signed before peak events reduce deployment risk",
        "Event-calendar demand certainty: 12 Tier-A Atlanta events May-Dec 2026 with confirmed dates",
        "Diversified revenue: 7 streams - rides, VIP blocks, BlackPass, sponsors, referrals, partner fleet, settlement",
        "Overflow model: partner fleet absorbs surge without capital over-commitment",
        "TROPTIONS settlement: every transaction audited, sponsor payouts escrowed, lender-visible KPIs",
        "WhichWay ecosystem: guest OS, merchant, map, and event modules drive ride volume independent of cold outreach",
    ]
    for r in reasons:
        pdf.bullet(r)
    pdf.divider()

    # Proposed loan structure summary
    pdf.section_title("Proposed Loan Structure (Summary)")
    loan_items = [
        ("Facility type",       "Term loan with staged draw-down tied to sponsor presales"),
        ("Amount",              "$650,000"),
        ("Term",                "24 months"),
        ("Proposed rate",       "Prime + 2-4% (negotiable based on security package)"),
        ("Security",            "1st lien on fleet vehicles + assignment of sponsor contracts + BlackPass escrow"),
        ("Repayment",           "Monthly P+I beginning month 3, after initial dispatch ramp"),
        ("Staged release",      "Tranche 1: $300K at close | Tranche 2: $350K on $100K sponsor pre-sales proof"),
    ]
    for i, (k, v) in enumerate(loan_items):
        pdf.kv_row(k, v, shaded=(i % 2 == 0))
    pdf.ln(2)

    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*GOLD)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_x(12)
    pdf.cell(186, 8, f"  UNYKORN BLACK  |  Confidential  |  {DATE_PREPARED}  |  {CONTACT_URL}", fill=True, ln=True)

    pdf.output(str(DOWNLOADS / "unykorn-black-executive-one-pager.pdf"))


# ============================================================
# Monthly Cash Flow Pro Forma
# ============================================================
def build_proforma_pdf() -> None:
    pdf = InvestorPDF(
        doc_title="Monthly Cash Flow Pro Forma",
        doc_subtitle="Base-case 8-month projection May-Dec 2026 | Debt service, EBITDA, and cumulative cash"
    )
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()

    pdf.section_title("Assumptions and Model Basis")
    assumptions = [
        "Base scenario: 6 premium vehicles active by June 2026 (2 launched May, +2 June, +2 August)",
        "Average dispatched fare: $85 per ride; average rides per vehicle per day: 4-6 depending on event density",
        "Sponsor revenue phased: $50K signed at launch, $150K by August (World Cup + TOUR Championship window)",
        "BlackPass presales: $25K May, $50K June-July (World Cup), $75K August-September peak",
        "Operating costs: driver pay (35% of fare), fuel/maintenance (8%), insurance (included in reserve), platform (3%)",
        "Fixed overhead: dispatch, compliance, admin estimated at $18,000/month",
        "Debt service: $650K facility at 10% annualized = approximately $27,000/month from month 3",
    ]
    for a in assumptions:
        pdf.bullet(a, size=9)
    pdf.divider()

    pdf.section_title("Monthly Pro Forma - Base Case (May-Dec 2026)")

    months = ["May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "TOTAL"]
    # Revenue rows
    ride_rev    = [42000, 110000, 135000, 195000, 175000, 95000, 105000, 195000, 0]
    blkpass     = [25000,  35000,  40000,  55000,  50000,  25000,  25000,  50000, 0]
    sponsor     = [15000,  25000,  35000,  65000,  50000,  25000,  30000,  50000, 0]
    vip_blocks  = [ 5000,  10000,  15000,  20000,  15000,  10000,  10000,  20000, 0]
    referral    = [ 2000,   5000,   7000,  12000,  10000,   5000,   6000,  10000, 0]
    # Calculate totals
    for arr in [ride_rev, blkpass, sponsor, vip_blocks, referral]:
        arr[-1] = sum(arr[:-1])
    gross_rev = [sum(x) for x in zip(ride_rev, blkpass, sponsor, vip_blocks, referral)]

    # Cost rows
    driver_cost = [int(r * 0.35) for r in ride_rev[:-1]] + [0]
    driver_cost[-1] = sum(driver_cost[:-1])
    fuel_maint  = [int(r * 0.08) for r in ride_rev[:-1]] + [0]
    fuel_maint[-1] = sum(fuel_maint[:-1])
    platform    = [int(r * 0.03) for r in ride_rev[:-1]] + [0]
    platform[-1] = sum(platform[:-1])
    fixed_oh    = [18000]*8 + [0]
    fixed_oh[-1] = sum(fixed_oh[:-1])
    total_opex  = [driver_cost[i]+fuel_maint[i]+platform[i]+fixed_oh[i] for i in range(9)]

    ebitda      = [gross_rev[i] - total_opex[i] for i in range(9)]
    debt_svc    = [0, 0, 27000, 27000, 27000, 27000, 27000, 27000, 0]
    debt_svc[-1] = sum(debt_svc[:-1])
    net_cf      = [ebitda[i] - debt_svc[i] for i in range(9)]
    cum_cf      = []
    running = 0
    for i, v in enumerate(net_cf[:-1]):
        running += v
        cum_cf.append(running)
    cum_cf.append("")  # total col blank for cumulative

    def fmt(v):
        if v == 0 or v == "":
            return "-" if v == 0 else ""
        return f"${v:,.0f}"

    def row(label, data, bold=False, shade=False, gold_bg=False):
        pdf.set_fill_color(*NAVY) if gold_bg else (pdf.set_fill_color(*LIGHT) if shade else pdf.set_fill_color(*WHITE))
        pdf.set_text_color(*GOLD) if gold_bg else pdf.set_text_color(*DARK)
        pdf.set_font("Helvetica", "B" if bold else "", 8)
        pdf.set_x(12)
        pdf.cell(30, 6, f"  {label}", fill=True, border=1)
        col_w = 156 / len(months)
        for i, v in enumerate(data):
            txt = fmt(v)
            pdf.cell(col_w, 6, txt, fill=True, border=1, align="R")
        pdf.ln()

    # Header row
    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*GOLD)
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_x(12)
    pdf.cell(30, 7, "  Line Item", fill=True, border=1)
    col_w = 156 / len(months)
    for m in months:
        pdf.cell(col_w, 7, m, fill=True, border=1, align="C")
    pdf.ln()

    row("Ride Revenue",    ride_rev,    shade=True)
    row("BlackPass Sales", blkpass,     shade=False)
    row("Sponsor Revenue", sponsor,     shade=True)
    row("VIP Blocks",      vip_blocks,  shade=False)
    row("Referrals",       referral,    shade=True)
    row("GROSS REVENUE",   gross_rev,   bold=True, gold_bg=True)
    pdf.ln(1)
    row("Driver Pay (35%)",driver_cost, shade=True)
    row("Fuel/Maint (8%)", fuel_maint,  shade=False)
    row("Platform (3%)",   platform,    shade=True)
    row("Fixed Overhead",  fixed_oh,    shade=False)
    row("TOTAL OPEX",      total_opex,  bold=True, shade=True)
    pdf.ln(1)
    row("EBITDA",          ebitda,      bold=True, gold_bg=True)
    row("Debt Service",    debt_svc,    shade=True)
    row("NET CASH FLOW",   net_cf,      bold=True, shade=False)
    row("Cumulative CF",   cum_cf,      shade=True)
    pdf.ln(3)
    pdf.divider()

    pdf.section_title("Break-Even Analysis")
    be_items = [
        ("Monthly fixed overhead",             "$18,000"),
        ("Monthly debt service (from month 3)","$27,000"),
        ("Total monthly fixed burden",         "$45,000"),
        ("Variable cost rate (% of fare rev)", "46%"),
        ("Break-even fare revenue / month",    "$83,333"),
        ("Break-even rides / month @ $85 avg", "~980 rides"),
        ("Break-even rides / vehicle / day",   "~5.4 rides  (base fleet 6 vehicles)"),
        ("Achieved at",                        "Month 2 (June - World Cup window)"),
    ]
    for i, (k, v) in enumerate(be_items):
        pdf.kv_row(k, v, bold_right=True, shaded=(i % 2 == 0))
    pdf.ln(2)

    pdf.section_title("DSCR by Scenario")
    dscr_rows = [
        ("Lean",   "5 vehicles", "$723,150",  "$278,350", "$162,000", "1.7x"),
        ("Base",   "6 vehicles", "$1,414,200","$433,200", "$162,000", "2.7x"),
        ("Growth", "8 vehicles", "$2,365,000","$820,000", "$162,000", "5.1x"),
    ]
    hdrs  = ["Scenario","Fleet","Revenue","EBITDA","Annual Debt Svc","DSCR"]
    wdths = [28,28,34,30,40,26]
    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*GOLD)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_x(12)
    for h, w in zip(hdrs, wdths):
        pdf.cell(w, 7, h, fill=True, border=1)
    pdf.ln()
    for i, row_data in enumerate(dscr_rows):
        pdf.set_fill_color(*LIGHT) if i % 2 == 0 else pdf.set_fill_color(*WHITE)
        pdf.set_text_color(*DARK)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_x(12)
        for val, w in zip(row_data, wdths):
            pdf.cell(w, 7, val, fill=True, border=1)
        pdf.ln()
    pdf.ln(3)

    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*GOLD)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_x(12)
    pdf.cell(186, 8, f"  UNYKORN BLACK Pro Forma  |  CONFIDENTIAL  |  {DATE_PREPARED}  |  {CONTACT_URL}", fill=True, ln=True)

    pdf.output(str(DOWNLOADS / "unykorn-black-proforma.pdf"))


# ============================================================
# Proposed Term Sheet
# ============================================================
def build_term_sheet_pdf() -> None:
    pdf = InvestorPDF(
        doc_title="Proposed Term Sheet",
        doc_subtitle="Non-binding indicative term sheet for lender review and credit committee use"
    )
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()

    pdf.section_title("Parties")
    parties = [
        ("Borrower",     "UNYKORN BLACK (operating entity - Atlanta, GA)"),
        ("Operator",     "Kevan Burns, Chairman - UnyKorn / FTH Trading LLC"),
        ("Lender",       "[Lender Name] - to be completed"),
        ("Platform",     "WhichWay Event Mobility Platform | https://fifa.unykorn.org/"),
        ("Settlement",   "TROPTIONS multi-chain settlement and escrow engine"),
    ]
    for i, (k, v) in enumerate(parties):
        pdf.kv_row(k, v, shaded=(i % 2 == 0))
    pdf.divider()

    pdf.section_title("Facility Terms")
    terms = [
        ("Facility type",          "Senior secured term loan with staged draw-down schedule"),
        ("Total amount",           "$650,000"),
        ("Currency",               "US Dollars (USD)"),
        ("Term",                   "24 months from first draw date"),
        ("Proposed interest rate", "Prime Rate + 2.00% to 4.00% per annum (negotiable)"),
        ("Current prime estimate", "7.50% (as of May 2026)  =>  blended ~9.5-11.5%"),
        ("Payment structure",      "Interest-only months 1-2; P+I monthly thereafter"),
        ("Monthly P+I estimate",   "~$27,000/month (at $650K, 10%, 24 months)"),
        ("Prepayment",             "No penalty after month 6; 1% fee months 3-5"),
        ("Default rate",           "Rate + 5.00% per annum"),
    ]
    for i, (k, v) in enumerate(terms):
        pdf.kv_row(k, v, shaded=(i % 2 == 0))
    pdf.divider()

    pdf.section_title("Draw Schedule")
    draws = [
        ("Tranche 1 - At Close",         "$300,000", "Vehicle deposits, insurance reserve, compliance, working capital"),
        ("Tranche 2 - Upon Trigger",      "$350,000", "Trigger: $100,000 in signed sponsor contracts OR BlackPass presales confirmed"),
        ("Tranche 2 Deadline",            "90 days",  "From close date; forfeited if trigger not met (reverts to Tranche 1 only)"),
    ]
    hdrs  = ["Draw", "Amount", "Condition"]
    wdths = [44, 30, 112]
    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*GOLD)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_x(12)
    for h, w in zip(hdrs, wdths):
        pdf.cell(w, 7, h, fill=True, border=1)
    pdf.ln()
    for i, (draw, amt, cond) in enumerate(draws):
        pdf.set_fill_color(*LIGHT) if i % 2 == 0 else pdf.set_fill_color(*WHITE)
        pdf.set_text_color(*DARK)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_x(12)
        pdf.cell(44, 7, draw, fill=True, border=1)
        pdf.cell(30, 7, amt, fill=True, border=1)
        pdf.cell(112, 7, cond, fill=True, border=1, ln=True)
    pdf.ln(2)
    pdf.divider()

    pdf.section_title("Security and Collateral")
    security = [
        ("1st lien - Vehicle fleet",        "First priority security interest on all financed vehicles (titled collateral)"),
        ("Contract assignment",             "Assignment of all signed sponsor contracts and BlackPass presale agreements"),
        ("Escrow pledge",                   "BlackPass and Founding Sponsor deposits held in TROPTIONS escrow; lender named beneficiary"),
        ("Personal guarantee",              "Personal guarantee from operator for full facility amount"),
        ("Fleet insurance endorsement",     "Lender named as additional insured and loss payee on all commercial vehicle policies"),
        ("Deposit account control",         "Lender right to deposit account control agreement (DACA) on primary operating account"),
    ]
    for i, (k, v) in enumerate(security):
        pdf.kv_row(k, v, shaded=(i % 2 == 0))
    pdf.divider()

    pdf.section_title("Financial Covenants")
    covenants = [
        ("Minimum DSCR",              "1.20x trailing 3-month tested quarterly from month 3"),
        ("Minimum cash reserve",      "$50,000 unrestricted cash at all times after month 2"),
        ("Insurance coverage",        "Commercial auto minimum $1M per occurrence; general liability $2M"),
        ("Reporting frequency",       "Monthly: dispatch volume, revenue summary, sponsor pipeline, cash position"),
        ("Audit right",               "Lender right to request TROPTIONS settlement report at any time"),
        ("Cross-default",             "Default on any vehicle financing, insurance lapse, or license revocation = event of default"),
    ]
    for i, (k, v) in enumerate(covenants):
        pdf.kv_row(k, v, shaded=(i % 2 == 0))
    pdf.divider()

    pdf.section_title("Conditions Precedent to Close")
    conditions = [
        "Executed loan agreement and security documents",
        "Evidence of commercial vehicle insurance (lender named additional insured)",
        "Vehicle purchase agreements or dealer LOIs for initial fleet units",
        "Signed sponsor LOI or contract for minimum $25,000 (Tranche 1 only)",
        "Operating agreement / entity documents for borrowing entity",
        "Personal financial statement and background check for operator",
        "TROPTIONS escrow account opened and beneficiary designation filed",
        "WhichWay platform dispatch account activated and verified",
    ]
    for c in conditions:
        pdf.bullet(c, size=9)
    pdf.divider()

    pdf.section_title("Fees")
    fees = [
        ("Origination fee",     "1.00% of total facility ($6,500) - deducted from first draw"),
        ("Legal/documentation", "Borrower responsible for lender counsel fees (est. $2,500-$5,000)"),
        ("Annual review fee",   "$500 annually"),
        ("Wire fees",           "At cost"),
    ]
    for i, (k, v) in enumerate(fees):
        pdf.kv_row(k, v, shaded=(i % 2 == 0))
    pdf.ln(2)

    pdf.section_title("Governing Law and Disclaimer")
    pdf.body(
        "This term sheet is non-binding and indicative only. It does not constitute a commitment to lend. "
        "Final terms are subject to due diligence, credit committee approval, legal documentation, and executed "
        "loan agreements. Governing law: State of Georgia, United States.",
        size=9
    )
    pdf.ln(3)

    # Signature blocks
    pdf.set_fill_color(*LIGHT)
    pdf.set_text_color(*DARK)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_x(12)
    pdf.cell(88, 24, "Borrower Signature: ___________________\n\nPrinted Name: ___________________\n\nDate: ___________________", fill=True, border=1, ln=False)
    pdf.cell(10, 24, "", ln=False)
    pdf.cell(88, 24, "Lender Signature: ___________________\n\nPrinted Name: ___________________\n\nDate: ___________________", fill=True, border=1, ln=True)
    pdf.ln(3)

    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*GOLD)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_x(12)
    pdf.cell(186, 8, f"  UNYKORN BLACK Term Sheet  |  CONFIDENTIAL  |  {DATE_PREPARED}  |  {CONTACT_URL}", fill=True, ln=True)

    pdf.output(str(DOWNLOADS / "unykorn-black-term-sheet.pdf"))


def build_zip() -> None:
    zip_path = DOWNLOADS / "unykorn-black-downloads.zip"
    members = [
        "unykorn-black-executive-one-pager.pdf",
        "unykorn-black-lender-proposal.pdf",
        "unykorn-black-proforma.pdf",
        "unykorn-black-term-sheet.pdf",
        "unykorn-black-sponsor-rate-card.pdf",
        "unykorn-black-troptions-settlement.pdf",
        "unykorn-black-lender-proposal.docx",
        "unykorn-black-event-demand-calendar.csv",
    ]
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for m in members:
            zf.write(DOWNLOADS / m, arcname=m)


def main() -> None:
    build_lender_pdf()
    build_rate_card_pdf()
    build_troptions_settlement_pdf()
    build_executive_one_pager()
    build_proforma_pdf()
    build_term_sheet_pdf()
    build_docx()
    build_csv()
    build_zip()
    print("All documents generated successfully.")


if __name__ == "__main__":
    main()
