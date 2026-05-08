from pathlib import Path
import csv
import zipfile

from fpdf import FPDF
from docx import Document

ROOT = Path(r"C:\Users\Kevan\OneDrive - FTH Trading\11-Downloads\unykorn-black-funder-site")
DOWNLOADS = ROOT / "downloads"
DOWNLOADS.mkdir(parents=True, exist_ok=True)

SCENARIOS = [
    ("Lean", "5 vehicles", "$723,150", "2.6x"),
    ("Base", "6 vehicles", "$1,414,200", "7.2x"),
    ("Growth", "8 vehicles", "$2,365,000", "10.5x"),
]

USE_OF_FUNDS = [
    ("Vehicle acquisition / down payments", "$115,000"),
    ("Equipment / vehicle financing facility", "$350,000"),
    ("Insurance reserve", "$45,000"),
    ("Licensing, compliance, legal, accounting", "$25,000"),
    ("Website, app, dispatch buildout", "$35,000"),
    ("Branding, QR, sponsor materials", "$20,000"),
    ("Working capital reserve", "$60,000"),
]

EVENTS = [
    ("May 24, 2026", "Birthday Bash ATL", "State Farm Arena", "B+", "Late-night VIP routes, after-event offers, sponsor nightlife traffic"),
    ("June 8-9, 2026", "Megan Moroney", "State Farm Arena", "B", "Concert transfers, hotel pickup, premium downtown routes"),
    ("June 9-14, 2026", "Atlanta Market + Atlanta Apparel", "AmericasMart", "A", "Buyers, vendors, showroom executives, airport and hotel routes"),
    ("June 15-July 15, 2026", "Atlanta World Cup match window", "Atlanta Stadium / Mercedes-Benz Stadium", "A+", "VIP arrivals, hotel-stadium-dining loops, sponsors"),
    ("Aug. 21, 2026", "Chris Stapleton", "Mercedes-Benz Stadium", "A", "Stadium concert, suite guests, sponsor hospitality, reserved exits"),
    ("Aug. 26-30, 2026", "TOUR Championship", "East Lake Golf Club", "A", "Executive golf, sponsor hospitality, corporate blocks, airport transfers"),
    ("Aug. 27, 2026", "AC/DC POWER UP Tour", "Mercedes-Benz Stadium", "A", "Major stadium concert, premium arrivals and late-night exits"),
    ("Sept. 3-7, 2026", "Dragon Con", "Downtown Atlanta", "A", "Multi-day hotel corridor, fan groups, VIP events, merchant offers"),
    ("Sept. 18-20, 2026", "Shaky Knees", "Piedmont Park", "A", "Festival rides, hotel packages, sponsor QR offers, group exits"),
    ("Oct. 10-11, 2026", "Atlanta Pride", "Piedmont Park", "B+", "Festival mobility, hotel routes, merchant and safety-oriented routing"),
    ("Nov. 10-11, 2026", "The R&B Tour", "Mercedes-Benz Stadium", "A", "High-demand concert nights, suite guests, sponsors, premium exits"),
    ("Dec. 5, 2026", "SEC Championship", "Mercedes-Benz Stadium", "A+", "Corporate hospitality, alumni groups, VIP fan travel, sponsor routes"),
]

REVENUE_MIX = [
    ("Owned rides", "$714,200", "Primary dispatched premium ride revenue"),
    ("VIP blocks", "$110,000", "Corporate and executive account ride blocks"),
    ("BlackPass", "$150,000", "Prepaid ride products sold before peak windows"),
    ("Sponsors", "$275,000", "Vehicle, route, and app inventory sponsorships"),
    ("Referrals", "$75,000", "Merchant referral and partner activation revenue"),
    ("Partner fleet", "$60,000", "Overflow and surge margin participation"),
    ("Settlement", "$30,000", "Platform-linked settlement and reporting services"),
]

SPONSOR_PACKAGES = [
    (
        "Founding Sponsor",
        "$150,000",
        "Top launch placement across site, app, fleet, confirmations, and proof reporting.",
        "Includes category exclusivity in defined launch windows and quarterly performance review.",
    ),
    (
        "Route Sponsor",
        "$25,000",
        "Own priority corridor inventory such as airport-hotel-stadium and post-event exits.",
        "Includes route-branded digital placements and location-based QR offer slots.",
    ),
    (
        "Vehicle Sponsor",
        "$15,000",
        "Sponsor a premium SUV or Sprinter with visible placement and confirmation touchpoints.",
        "Includes branded passenger materials and tracked rider exposure metrics.",
    ),
    (
        "Merchant Listing",
        "$2,500",
        "Verified listing for restaurants, hotels, lounges, retail, and premium services.",
        "Includes periodic promotion support and referral performance summaries.",
    ),
]

BLACKPASS = [
    ("BlackPass Match", "$750-$1,500", "Match-day round trip with QR ride pass"),
    ("BlackPass Weekend", "$2,500-$5,000", "Multi-day ride credits and concierge support"),
    ("BlackPass Executive", "$7,500-$15,000", "Priority vehicle blocks and premium support"),
    ("BlackPass Hotel", "$15,000-$50,000", "Hotel guest mobility package"),
]


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
    pdf.set_fill_color(7, 17, 31)
    pdf.set_text_color(214, 168, 79)
    pdf.set_font("Helvetica", "B", 19)
    pdf.cell(0, 12, title, ln=True)
    pdf.set_text_color(40, 40, 40)
    pdf.set_font("Helvetica", "", 11)
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(190, 6, subtitle)
    pdf.ln(3)


def build_lender_pdf() -> None:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    add_title(
        pdf,
        "UNYKORN BLACK - Lender Proposal",
        "Premium Event Mobility for Atlanta 2026. Requested facility: $650,000 for a controlled 5-8 vehicle launch with sponsor-backed upside, disciplined risk controls, and event-driven demand routing.",
    )

    section_title(pdf, "1) Executive Summary")
    paragraph(
        pdf,
        "UNYKORN BLACK is a premium mobility and sponsorship platform designed for Atlanta's 2026 event economy. The operating model combines black SUV and sedan services, group-capacity premium vans, route sponsorship inventory, merchant referral economics, and platform-linked settlement visibility.",
    )
    paragraph(
        pdf,
        "The core lender proposition is straightforward: a collateral-backed transportation operation with multiple monetization layers beyond fare revenue. The business is structured to preserve downside through controlled launch scaling while capturing upside from sponsorship, pre-sold ride credits, account contracts, and high-density event windows.",
    )

    section_title(pdf, "2) Funding Request and Use")
    paragraph(
        pdf,
        "Requested facility: $650,000. The facility supports launch readiness, compliance discipline, insurance reserve posture, and controlled operating liquidity through the first major event cycles.",
    )
    for item, amount in USE_OF_FUNDS:
        line_item(pdf, item, amount)
    pdf.set_font("Helvetica", "B", 11)
    line_item(pdf, "Total listed use of funds", "$650,000")

    section_title(pdf, "3) Revenue Model and Scenario View")
    paragraph(
        pdf,
        "Revenue is built on seven streams. Base-case gross revenue for May-Dec 2026 is projected at $1,414,200, with EBITDA of $433,200 and DSCR of 7.2x.",
    )
    for name, fleet, rev, dscr in SCENARIOS:
        line_item(pdf, f"{name} scenario - {fleet}", f"Revenue {rev} | DSCR {dscr}")

    pdf.ln(1)
    for name, value, note in REVENUE_MIX:
        line_item(pdf, f"{name}: {note}", value)

    section_title(pdf, "4) Prepaid BlackPass Strategy")
    paragraph(
        pdf,
        "BlackPass creates prepaid demand and improves dispatch confidence before high-load periods. The product ladder is intentionally tiered to serve both consumer and institutional buyers.",
    )
    for tier, price, desc in BLACKPASS:
        line_item(pdf, f"{tier} - {desc}", price)

    section_title(pdf, "5) Sponsor Engine and Commercial Inventory")
    paragraph(
        pdf,
        "Sponsor revenue is attached to tangible inventory: routes, vehicles, passenger confirmations, and premium placements. This translates brand budgets into measurable transportation-channel economics.",
    )
    for title, price, short_desc, plus_desc in SPONSOR_PACKAGES:
        line_item(pdf, f"{title}: {short_desc}", price)
        paragraph(pdf, f"   {plus_desc}", size=10, spacing=5)

    section_title(pdf, "6) Collateral and Risk Controls")
    paragraph(pdf, "Risk management is built around six operational controls:")
    controls = [
        "Vehicle collateral and titled asset visibility on core owned fleet units.",
        "Account-backed dispatch and route-level booking discipline.",
        "Prepaid BlackPass credits and sponsor contracts before peak demand windows.",
        "Partner fleet overflow to avoid over-acquiring assets before demand proof.",
        "Insurance reserve and compliance budgeting held in the launch package.",
        "TROPTIONS-linked settlement reporting for transaction-level audit trail visibility.",
    ]
    for idx, c in enumerate(controls, start=1):
        paragraph(pdf, f"{idx}. {c}", size=10, spacing=5.5)

    pdf.add_page()
    add_title(
        pdf,
        "UNYKORN BLACK - Operating Plan Annex",
        "Detailed execution notes for lenders, credit committees, and underwriting support teams.",
    )

    section_title(pdf, "7) Event Demand Calendar and Revenue Angle")
    paragraph(
        pdf,
        "The Atlanta event calendar is the demand backbone for launch sequencing. The table below is used for sponsor targeting, route planning, driver scheduling, and hotel-corporate conversion.",
    )
    for date, event, venue, tier, angle in EVENTS:
        paragraph(pdf, f"{date} | {event} | {venue} | Tier {tier}", size=10, spacing=5)
        paragraph(pdf, f"Revenue angle: {angle}", size=10, spacing=5)

    section_title(pdf, "8) WhichWay Ecosystem Integration")
    paragraph(
        pdf,
        "UNYKORN BLACK sits inside the WhichWay event operating stack (https://fifa.unykorn.org/). This matters to lenders because mobility is not isolated - it is integrated with guest flow, mapping, merchant traffic, emergency support, and sponsor touchpoints.",
    )
    ecosystem_points = [
        "Guest OS: route and service layer for guest transport planning.",
        "Demo and Passport: engagement and loyalty mechanisms tied to trip behavior.",
        "Venue map and merchant modules: destination routing and referral monetization.",
        "Emergency module: support escalation controls and guest safety workflows.",
        "Sales and signup modules: sponsor and account pipeline capture.",
    ]
    for p in ecosystem_points:
        paragraph(pdf, f"- {p}", size=10, spacing=5.5)

    section_title(pdf, "9) Underwriting Notes and Recommendation")
    paragraph(
        pdf,
        "The proposed launch structure offers a practical lender profile: visible collateral, event-driven demand concentration, diversified revenue streams, and operational safeguards that reduce over-expansion risk. Recommendation: structure the facility with staged release triggers linked to sponsor pre-sales and account dispatch volumes.",
    )
    paragraph(
        pdf,
        "Prepared for lender review: UNYKORN BLACK Funding Packet. Date: 2026-05-08.",
        size=10,
        spacing=5.5,
    )

    pdf.output(str(DOWNLOADS / "unykorn-black-lender-proposal.pdf"))


def build_rate_card_pdf() -> None:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    add_title(
        pdf,
        "UNYKORN BLACK - Sponsor Rate Card",
        "Premium inventory across vehicles, routes, app flows, and merchant activations. Structured for measurable sponsor value and clear reporting cadence.",
    )

    section_title(pdf, "Inventory Framework")
    paragraph(
        pdf,
        "Sponsor inventory is sold across five channels: vehicle exterior/interior branding, route ownership, passenger touchpoints, app placements, and merchant activation bundles.",
    )

    section_title(pdf, "Package Menu")
    for title, price, short_desc, plus_desc in SPONSOR_PACKAGES:
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(190, 6, f"{title} - {price}")
        paragraph(pdf, short_desc, size=10, spacing=5.5)
        paragraph(pdf, plus_desc, size=10, spacing=5.5)
        pdf.ln(0.8)

    section_title(pdf, "Reporting and Proof Standards")
    report_points = [
        "Monthly sponsor performance summary with delivered inventory counts.",
        "Route and placement logs with campaign period coverage.",
        "Periodic creative proof snapshots for compliance and partner records.",
        "Renewal recommendations based on ride window and event performance.",
    ]
    for p in report_points:
        paragraph(pdf, f"- {p}", size=10, spacing=5.5)

    section_title(pdf, "Commercial Contact")
    paragraph(
        pdf,
        "UNYKORN BLACK Partnerships\nWebsite: https://fifa.unykorn.org/\nPurpose: route sponsorship, vehicle sponsorship, merchant placement, and strategic launch partnerships.",
    )

    pdf.output(str(DOWNLOADS / "unykorn-black-sponsor-rate-card.pdf"))


def build_docx() -> None:
    doc = Document()
    doc.add_heading("UNYKORN BLACK - Editable Funding Proposal", level=1)
    doc.add_paragraph(
        "Premium event mobility launch package for Atlanta 2026. This editable document is prepared for lenders, funding groups, and strategic partners requiring a fully written underwriting summary."
    )

    doc.add_heading("Funding Request", level=2)
    doc.add_paragraph("$650,000 controlled launch package for a 5-8 vehicle premium event fleet with sponsor-backed upside and collateral-aware operating controls.")

    doc.add_heading("Executive Narrative", level=2)
    doc.add_paragraph(
        "UNYKORN BLACK combines premium transportation, sponsor route inventory, merchant referrals, and platform-linked reporting into one integrated operating model. The launch structure emphasizes disciplined growth through prepaid demand, account-backed dispatch, and partner-fleet overflow before permanent asset expansion."
    )

    doc.add_heading("Scenario Snapshot", level=2)
    table = doc.add_table(rows=1, cols=4)
    hdr = table.rows[0].cells
    hdr[0].text = "Scenario"
    hdr[1].text = "Fleet"
    hdr[2].text = "Gross Revenue"
    hdr[3].text = "DSCR"
    for row in SCENARIOS:
        c = table.add_row().cells
        c[0].text = row[0]
        c[1].text = row[1]
        c[2].text = row[2]
        c[3].text = row[3]

    doc.add_heading("Revenue Mix", level=2)
    mix = doc.add_table(rows=1, cols=3)
    mh = mix.rows[0].cells
    mh[0].text = "Stream"
    mh[1].text = "Value"
    mh[2].text = "Comment"
    for name, value, note in REVENUE_MIX:
        c = mix.add_row().cells
        c[0].text = name
        c[1].text = value
        c[2].text = note

    doc.add_heading("Use of Funds", level=2)
    for item, amount in USE_OF_FUNDS:
        doc.add_paragraph(f"{item}: {amount}", style="List Bullet")

    doc.add_heading("BlackPass Presale Strategy", level=2)
    for tier, price, desc in BLACKPASS:
        doc.add_paragraph(f"{tier} ({price}): {desc}", style="List Bullet")

    doc.add_heading("Sponsor Package Structure", level=2)
    for title, price, short_desc, plus_desc in SPONSOR_PACKAGES:
        doc.add_paragraph(f"{title} - {price}", style="List Bullet")
        doc.add_paragraph(short_desc)
        doc.add_paragraph(plus_desc)

    doc.add_heading("Collateral and Risk Controls", level=2)
    controls = [
        "Vehicle collateral position on core owned units",
        "Prepaid BlackPass and sponsor agreement demand support",
        "Partner-fleet overflow to avoid over-acquisition",
        "Insurance reserve and compliance budgeting",
        "Settlement and reporting visibility through TROPTIONS-linked flows",
    ]
    for c in controls:
        doc.add_paragraph(c, style="List Bullet")

    doc.add_heading("Event Demand Window", level=2)
    for date, event, venue, tier, angle in EVENTS:
        doc.add_paragraph(f"{date} | {event} | {venue} | Tier {tier} | {angle}")

    doc.add_heading("WhichWay Platform Integration", level=2)
    doc.add_paragraph(
        "Live platform references: https://fifa.unykorn.org/, /guest, /demo, /passport, /map, /merchants, /emergency, /sign-up, /sales."
    )
    doc.add_paragraph("Powered by TROPTIONS reporting and WhichWay platform integration.")

    doc.save(str(DOWNLOADS / "unykorn-black-lender-proposal.docx"))


def build_csv() -> None:
    csv_path = DOWNLOADS / "unykorn-black-event-demand-calendar.csv"
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["date", "event", "venue", "tier", "revenue_angle"])
        writer.writerows(EVENTS)


def build_troptions_settlement_pdf() -> None:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    add_title(
        pdf,
        "TROPTIONS Settlement & Escrow Engine",
        "POF, Multi-Asset Clearing, and Issuer Opportunity Framework for UNYKORN BLACK and WhichWay Network",
    )

    section_title(pdf, "1) Executive Summary: The Settlement Backbone")
    paragraph(
        pdf,
        "TROPTIONS is the settlement infrastructure enabling UNYKORN BLACK and the broader WhichWay event ecosystem. The platform provides Proof of Funds (POF) custody, multi-asset escrow, transparent transaction reporting, and issuer tokenization pathways for sponsors and partners.",
    )
    paragraph(
        pdf,
        "This document outlines the commercial opportunity for corporate sponsors, merchants, and issuers looking to participate in event mobility and settlement economics - and how TROPTIONS creates trust, compliance visibility, and equity upside for the entire network.",
    )

    section_title(pdf, "2) What is TROPTIONS Settlement?")
    paragraph(
        pdf,
        "TROPTIONS combines regulated escrow custody, multi-chain settlement (USDC, ATP, and partner assets), real-time transaction visibility, and issuer tokenization to create a single trust layer for event commerce and mobility.",
    )
    features = [
        "Proof of Funds (POF) custody: Sponsor prepayments and prepaid ride credits held in regulated escrow",
        "Multi-asset settlement: USDC stablecoin, ATP settlement tokens, and future partner integrations",
        "Transaction-level reporting: Every ride, sponsorship dollar, and settlement flow is auditable and transparent",
        "Issuer tokenization: Sponsors and corporate partners can mint equity or performance tokens backed by actual ride volume and sponsor revenue",
        "WhichWay ecosystem integration: Guest, map, merchant, and emergency modules all report settlement state",
    ]
    for f in features:
        paragraph(pdf, f"- {f}", size=10, spacing=5.5)

    section_title(pdf, "3) UNYKORN BLACK Use Case: Why Settlement Matters")
    paragraph(
        pdf,
        "UNYKORN BLACK operates with prepaid demand (BlackPass), sponsor commitments (route and vehicle inventory), and account-backed corporate blocks. TROPTIONS provides:",
    )
    black_benefits = [
        "Pre-event sponsor funds held in custody, releasing only on proof of delivery (rides confirmed, QR scans logged)",
        "Real-time KPI reporting: sponsors see their vehicle placements, ride volume, and ROI by date, hour, and venue",
        "Compliance audit trail: every sponsor dollar, every ride, every settlement recorded and auditable",
        "Fast settlement: sponsor and corporate account charges settle daily to USDC or ATP, reducing operational float",
    ]
    for b in black_benefits:
        paragraph(pdf, f"- {b}", size=10, spacing=5.5)

    section_title(pdf, "4) POF and Escrow Framework")
    paragraph(
        pdf,
        "Proof of Funds is the cornerstone of sponsor confidence and lender certainty. TROPTIONS custody model:",
    )
    pof_points = [
        "Sponsor prepayment deposits held in USDC escrow upon contract signature - never mixed with operational float",
        "Escrow release tied to KPI delivery: minimum ride volume, geographic coverage, compliance checks, and event window performance",
        "Monthly reconciliation and reporting to sponsors, lenders, and regulatory oversight",
        "Full chain-of-custody documentation - every transaction ID, timestamp, and settlement proof recorded",
    ]
    for p in pof_points:
        paragraph(pdf, f"- {p}", size=10, spacing=5.5)

    pdf.add_page()
    add_title(
        pdf,
        "TROPTIONS Sponsor Opportunity - Equity & Token Sales",
        "How Sponsors Become Issuers and Capture Upside",
    )

    section_title(pdf, "5) Sponsor as Issuer: Equity and Token Participation")
    paragraph(
        pdf,
        "A Founding Sponsor or Route Sponsor can structure their participation as a tokenized equity or performance share. Instead of a fixed $150K or $25K sponsorship spend, the sponsor can allocate capital toward ownership in one or more event mobility verticals.",
    )

    section_title(pdf, "5a) Equity Issuance Model")
    equity_model = [
        "Sponsor invests $150K in UNYKORN BLACK and receives tokenized equity (e.g., 5-10% revenue share or profit participation)",
        "Tokens are minted on the Apostle Chain (chain_id 7332) and registered with TROPTIONS settlement engine",
        "Monthly settlement: UNYKORN BLACK revenue attributed to that route/sponsor tier is calculated, and sponsor token holders receive proportional ATP settlement",
        "Transferable: sponsor can sell, trade, or hold their equity tokens on the Apostle marketplace",
    ]
    for em in equity_model:
        paragraph(pdf, f"- {em}", size=10, spacing=5.5)

    section_title(pdf, "5b) Performance Token Model (Pay-Per-Ride Upside)")
    paragraph(
        pdf,
        "Alternative structure: sponsor pays a smaller upfront fee ($25-50K) and receives performance tokens that earn based on actual ride volume and merchant referrals triggered by their branded inventory.",
    )
    perf_model = [
        "Sponsor mints 1,000 performance tokens at $25 each",
        "Each ride using sponsor route earns 0.5 ATP per ride (distributed to token holders quarterly)",
        "Merchant referrals attributed to sponsor signage earn 0.1 ATP per qualified lead",
        "Token holders can claim their ATP earnings directly to their Apostle Chain wallet or keep compounding",
    ]
    for pm in perf_model:
        paragraph(pdf, f"- {pm}", size=10, spacing=5.5)

    section_title(pdf, "6) Issuer Commercial Terms")
    paragraph(
        pdf,
        "TROPTIONS charges 2% settlement fee on all sponsor payouts. Equity and token sponsors benefit from full transparency and monthly settlement.",
    )
    issuer_terms = [
        "Setup: TROPTIONS registry entry + Apostle Chain wallet provisioning + smart contract token mint",
        "Settlement: daily ride/referral aggregation, monthly ATP calculation and settlement",
        "Custody: sponsor tokens held in non-custodial Apostle Chain wallets; sponsors own private keys",
        "Reporting: daily KPI dashboard, monthly settlement statement, quarterly tax-ready reporting",
        "Transferability: tokens tradeable on Apostle marketplace, no lockup period after minting",
    ]
    for it in issuer_terms:
        paragraph(pdf, f"- {it}", size=10, spacing=5.5)

    section_title(pdf, "7) Compliance and Audit")
    paragraph(
        pdf,
        "TROPTIONS settlement is built for compliance and audit. All transactions are recorded on the Apostle Chain (transparent, immutable) and cross-referenced with WhichWay KPI systems.",
    )
    compliance = [
        "SEC framework: tokenized equity and revenue participation structures reviewed for securities compliance",
        "AML/KYC: sponsor identities verified at token mint time; settlement flows logged and reported",
        "Tax reporting: monthly settlement statements capture realized gains, losses, and income for 1099/foreign tax reporting",
        "Audit trail: every ride, route placement, merchant referral, and settlement proof auditable back to source transaction",
    ]
    for c in compliance:
        paragraph(pdf, f"- {c}", size=10, spacing=5.5)

    section_title(pdf, "8) Why Sponsors Choose TROPTIONS + UNYKORN BLACK")
    paragraph(
        pdf,
        "Sponsors and issuers benefit from a complete operating stack: real-time event mobility asset inventory, transparent sponsor ROI measurement, and tokenized equity upside backed by actual ride economics.",
    )
    why_troptions = [
        "No middleman: direct token ownership, no escrow lock-in, daily settlement to your wallet",
        "Real revenue data: sponsorship ROI is calculated in rides, conversions, and merchant traffic - not guesses",
        "Liquid asset class: your token is tradeable day one; no 5-7 year lockup",
        "Network effects: more sponsors = more ride inventory = higher token value for early issuers",
        "Ecosystem scaling: WhichWay guest flow, demo, passport, and merchant modules all drive sponsor token value",
    ]
    for w in why_troptions:
        paragraph(pdf, f"- {w}", size=10, spacing=5.5)

    section_title(pdf, "9) Next Steps: Become an Issuer")
    paragraph(
        pdf,
        "Sponsors interested in equity or token participation should contact TROPTIONS directly.",
    )
    next_steps = [
        "Complete issuer KYC application and commercial term sheet",
        "Provide capital commitment and token allocation preference (equity model vs. performance model)",
        "TROPTIONS provisions Apostle Chain wallet and mints token contract",
        "Settlement begins within 30 days of first UNYKORN BLACK rides under your sponsorship",
    ]
    for ns in next_steps:
        paragraph(pdf, f"- {ns}", size=10, spacing=5.5)

    paragraph(
        pdf,
        "\nTROPTIONS Settlement Platform\nhttps://fifa.unykorn.org/\nSupporting UNYKORN BLACK, WhichWay ecosystem, and multi-chain event commerce.",
        size=10,
        spacing=5.5,
    )

    pdf.output(str(DOWNLOADS / "unykorn-black-troptions-settlement.pdf"))


def build_zip() -> None:
    zip_path = DOWNLOADS / "unykorn-black-downloads.zip"
    members = [
        "unykorn-black-lender-proposal.pdf",
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
    build_docx()
    build_csv()
    build_zip()


if __name__ == "__main__":
    main()
