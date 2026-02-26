import os
import re
import json
import requests
import pandas as pd
from bs4 import BeautifulSoup
from openai import OpenAI
from anthropic import Anthropic
from playwright.sync_api import sync_playwright

OUTPUT_DIR = "one_pagers_out"
import os as _os
CSV_PATH = "team_tal.csv" if _os.path.exists("team_tal.csv") else "tal.csv"
N = 1  # lock to 1 while tuning

NETSUITE_LOGO_URL = "https://www.arin-innovation.com/wp-content/uploads/2022/09/Oracle-NetSuite-Portada.png"

REP_CONTACTS = {
    "brian o'neill": {
        "name": "Brian O'Neill",
        "title": "Senior Account Executive",
        "email": "brian.br.oneill@oracle.com",
        "phone": "(702) 306-1527"
    },
    "tommy dambrosio": {
        "name": "Tommy Dambrosio",
        "title": "Senior Account Executive",
        "email": "thomas.dambrosio@oracle.com",
        "phone": "(813) 380-8626"
    },
    "danielle corbett": {
        "name": "Danielle Corbett",
        "title": "Senior Account Executive",
        "email": "danielle.corbett@oracle.com",
        "phone": "(949) 525-6943"
    },
    "reggie traywick": {
        "name": "Reggie Traywick",
        "title": "Account Executive, Products",
        "email": "reginald.traywick@oracle.com",
        "phone": "(805) 403-9084"
    },
    "peter uritis": {
        "name": "Peter Uritis",
        "title": "Account Executive",
        "email": "peter.uritis@oracle.com",
        "phone": "(949) 939-4205"
    },
    "chris dynek": {
        "name": "Chris Dynek",
        "title": "Corporate Account Executive",
        "email": "chris.dynek@oracle.com",
        "phone": "(831) 905-4490"
    },
    "ross zapalac": {
        "name": "Ross Zapalac",
        "title": "Senior Account Executive",
        "email": "ross.b.zapalac@oracle.com",
        "phone": "(832) 229-8103"
    },
}

VERTICAL_INTEL = {
    "manufacturing": {
        "modules": ["NetSuite Manufacturing", "Work Orders & Routings", "NetSuite MRP", "Shop Floor Control", "NetSuite WMS", "Demand Planning", "Quality Management"],
        "pain_points": ["production scheduling across multiple work centers", "raw material procurement tied to demand forecasts", "work order visibility and labor tracking", "lot and serial number traceability", "BOM management across product variants", "shop floor visibility and machine utilization", "quality holds and non-conformance tracking"],
        "metrics": ["on-time delivery rate", "production yield", "inventory turns", "days inventory outstanding", "scrap rate", "manufacturing cycle time", "BOM accuracy"],
        "personas": {"finance": "CFO cares about COGS accuracy, variance reporting, and cash tied up in WIP", "ops": "VP Ops cares about production throughput, downtime, and supplier lead times", "sales": "Sales cares about accurate ATP dates and order-to-ship cycle time"}
    },
    "retail": {
        "modules": ["SuiteCommerce", "NetSuite POS", "Multi-Location Inventory", "Demand Planning", "NetSuite WMS", "CRM", "NetSuite Returns Management"],
        "pain_points": ["inventory visibility across multiple store locations", "seasonal demand spikes creating stockouts or overstock", "unified view of online and in-store sales", "rental fleet tracking and utilization", "end-of-season markdown coordination", "customer data fragmented across channels", "returns and exchanges across locations"],
        "metrics": ["inventory turnover", "sell-through rate", "stockout rate", "gross margin by location", "return rate", "average transaction value", "same-store sales growth"],
        "personas": {"finance": "CFO cares about margin by location, shrinkage, and working capital tied up in inventory", "ops": "VP Ops cares about replenishment speed, stockout rates, and transfer orders", "sales": "Sales cares about loyalty, omnichannel experience, and promotional performance"}
    },
    "apparel, footwear & accessories": {
        "modules": ["NetSuite Matrix Items", "SuiteCommerce", "Demand Planning", "NetSuite WMS", "EDI integration", "NetSuite Returns Management", "Multi-Location Inventory"],
        "pain_points": ["managing SKU explosion across size/color/style matrices", "EDI compliance with wholesale retail partners", "pre-season buying and open-to-buy planning", "landed cost calculation for imported goods", "return rates and reverse logistics", "seasonal collection changeovers", "DTC vs wholesale channel margin visibility"],
        "metrics": ["sell-through rate", "open-to-buy accuracy", "return rate", "gross margin by channel", "SKU productivity", "landed cost accuracy", "weeks of supply"],
        "personas": {"finance": "CFO cares about landed costs, channel margin, and working capital in seasonal inventory", "ops": "VP Ops cares about warehouse throughput, EDI compliance, and return processing speed", "sales": "Sales cares about DTC growth, wholesale fill rates, and new collection launch performance"}
    },
    "food & beverage": {
        "modules": ["NetSuite Lot Tracking", "Recipe & Formula Management", "NetSuite WMS", "Demand Planning", "Quality Management", "Catch Weight", "Regulatory Compliance"],
        "pain_points": ["lot traceability for recalls and compliance", "recipe and formula management across production runs", "catch weight and variable unit pricing", "shelf life and FEFO inventory management", "co-packer and contract manufacturer visibility", "FDA/FSMA compliance documentation", "seasonal and promotional demand volatility"],
        "metrics": ["days inventory outstanding", "waste and spoilage rate", "on-time delivery", "recall response time", "production yield", "cost per unit", "compliance audit pass rate"],
        "personas": {"finance": "CFO cares about waste costs, COGS by SKU, and working capital in perishable inventory", "ops": "VP Ops cares about lot traceability, production yield, and co-packer compliance", "sales": "Sales cares about fill rates, promotional lift, and new product launch timelines"}
    },
    "health & beauty": {
        "modules": ["NetSuite Lot Tracking", "Quality Management", "SuiteCommerce", "NetSuite WMS", "Demand Planning", "Regulatory Compliance", "CRM"],
        "pain_points": ["FDA and GMP compliance documentation", "lot traceability and expiry management", "DTC subscription and recurring order management", "influencer and affiliate revenue attribution", "rapid SKU proliferation across product lines", "contract manufacturer quality controls", "Amazon and retail channel inventory sync"],
        "metrics": ["subscription retention rate", "return rate", "compliance audit results", "DTC vs retail margin", "inventory turnover", "customer acquisition cost", "lot recall response time"],
        "personas": {"finance": "CFO cares about DTC vs retail margin, compliance costs, and working capital", "ops": "VP Ops cares about lot traceability, contract manufacturer compliance, and fulfillment speed", "sales": "Sales cares about subscription growth, DTC conversion, and retail shelf performance"}
    },
    "wholesale/distribution": {
        "modules": ["NetSuite WMS", "Demand Planning", "NetSuite SuiteCommerce", "EDI integration", "Multi-Location Inventory", "Landed Cost", "3PL Management"],
        "pain_points": ["order fulfillment accuracy across multiple warehouses", "EDI compliance with large retail customers", "landed cost visibility on imported goods", "vendor managed inventory programs", "pick pack ship efficiency and labor cost", "freight cost allocation and carrier management", "customer-specific pricing and contract management"],
        "metrics": ["order fill rate", "on-time shipment", "warehouse cost per order", "EDI compliance rate", "inventory accuracy", "days sales outstanding", "freight cost as % of revenue"],
        "personas": {"finance": "CFO cares about landed costs, DSO, and warehouse labor as % of revenue", "ops": "VP Ops cares about fulfillment accuracy, carrier performance, and warehouse utilization", "sales": "Sales cares about fill rates, EDI compliance scores, and customer portal access"}
    },
    "software": {
        "modules": ["NetSuite Revenue Recognition", "SuiteBilling", "NetSuite PSA", "CRM", "Financial Consolidation", "Multi-Currency", "NetSuite OpenAir"],
        "pain_points": ["ASC 606 revenue recognition across multi-element contracts", "subscription billing and renewal management", "professional services project margin tracking", "multi-entity consolidation for VC reporting", "deferred revenue waterfall accuracy", "sales commission calculations", "customer success metrics and churn visibility"],
        "metrics": ["ARR/MRR", "churn rate", "revenue recognition accuracy", "project margin", "DSO", "time-to-revenue", "customer acquisition cost"],
        "personas": {"finance": "CFO cares about ASC 606 compliance, deferred revenue accuracy, and board-ready financials", "ops": "VP Ops cares about project margins, utilization rates, and PS delivery efficiency", "sales": "Sales cares about renewal rates, upsell visibility, and commission accuracy"}
    },
    "services": {
        "modules": ["NetSuite PSA", "NetSuite OpenAir", "CRM", "SuiteBilling", "Revenue Recognition", "Project Management", "Resource Scheduling"],
        "pain_points": ["project profitability visibility in real time", "resource utilization and scheduling across projects", "time and expense capture accuracy", "milestone billing and contract management", "subcontractor cost tracking", "multi-project financial consolidation", "client reporting and project dashboards"],
        "metrics": ["utilization rate", "project margin", "bill rate realization", "DSO", "time-to-invoice", "client satisfaction score", "revenue per employee"],
        "personas": {"finance": "CFO cares about project margin, revenue recognition timing, and utilization-driven forecasting", "ops": "VP Ops cares about resource scheduling, subcontractor compliance, and project delivery", "sales": "Sales cares about pipeline-to-project conversion, upsell timing, and client retention"}
    },
    "general business": {
        "modules": ["NetSuite ERP", "NetSuite CRM", "SuiteAnalytics", "Financial Management", "NetSuite Planning & Budgeting", "Multi-Currency", "Vendor Management"],
        "pain_points": ["manual reporting across disconnected spreadsheets", "lack of real-time financial visibility for decisions", "month-end close taking too long", "no single source of truth across departments", "scaling headcount without scaling costs", "audit preparation and compliance documentation", "cash flow forecasting accuracy"],
        "metrics": ["days to close", "reporting accuracy", "headcount ratio", "DSO", "budget variance", "audit findings", "cash flow forecast accuracy"],
        "personas": {"finance": "CFO cares about close speed, audit readiness, and real-time financial visibility", "ops": "VP Ops cares about cross-department process efficiency and headcount leverage", "sales": "Sales cares about pipeline visibility, quote accuracy, and order-to-cash speed"}
    },
}

def get_vertical_intel(vertical: str) -> str:
    if not vertical or str(vertical).lower() in ("nan", "none", ""):
        return ""
    key = str(vertical).strip().lower()
    intel = VERTICAL_INTEL.get(key)
    if not intel:
        # fuzzy match on first word
        for k, v in VERTICAL_INTEL.items():
            if key.split()[0] in k:
                intel = v
                break
    if not intel:
        return ""
    return f"""Vertical Intelligence for {vertical}:
- Relevant NetSuite modules: {", ".join(intel["modules"])}
- Typical pain points in this vertical: {"; ".join(intel["pain_points"])}
- KPIs this vertical tracks: {", ".join(intel["metrics"])}
- Finance persona focus: {intel["personas"]["finance"]}
- Ops persona focus: {intel["personas"]["ops"]}
- Sales persona focus: {intel["personas"]["sales"]}"""


DEFAULT_REP = {
    "name": "NetSuite Sales Team",
    "title": "Account Executive",
    "email": "netsuite@oracle.com",
    "phone": ""
}

def get_rep(rep_name: str) -> dict:
    if not rep_name or str(rep_name).lower() in ("nan", "none", ""):
        return DEFAULT_REP
    name = str(rep_name).strip()
    # Normalize "Last, First" to "First Last"
    if "," in name:
        parts = [p.strip() for p in name.split(",", 1)]
        name = f"{parts[1]} {parts[0]}"
    # Normalize apostrophes and whitespace
    name = name.replace("’", "'").replace("‘", "'").strip().lower()
    match = REP_CONTACTS.get(name)
    if match:
        return match
    # Fuzzy fallback: match on last name only
    last = name.split()[-1] if name.split() else ""
    for key, val in REP_CONTACTS.items():
        if last and key.split()[-1] == last:
            return val
    return DEFAULT_REP

BANNER_HTML_TEMPLATE = """
<div class="banner">
  <div class="banner-left">
    <div class="company-name">{company_name}</div>
    <div class="headline">{headline}</div>
    <div class="subheadline">{subheadline}</div>
  </div>
  <div class="banner-right">
      <img class="ns-logo" src="{logo_url}" alt="NetSuite logo">
  </div>
</div>
"""

PAGE_CSS = """
@page { size: 8.5in 5.5in; margin: 0.35in; }
html, body { margin: 0; padding: 0; }

.page {
  width: 7.8in;
  height: 4.8in;
  overflow: hidden;
  box-sizing: border-box;
  background: #F4F7F8;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
  display: flex;
  flex-direction: column;
}

.banner {
  height: 1.1in;
  min-height: 1.1in;
  max-height: 1.1in;
  display: grid;
  grid-template-columns: 5fr 1.2fr;
  column-gap: 0.2in;
  background: #2E4759;
  color: #F4F7F8;
  padding: 0.15in 0.25in;
  box-sizing: border-box;
}
.banner-left {
  display: flex;
  flex-direction: column;
  justify-content: center;
  min-width: 0;
  overflow: hidden;
}
.company-name {
  font-size: 13px;
  font-weight: 800;
  font-style: italic;
  letter-spacing: 0.4px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  color: #AFC2D3;
}
.headline {
  margin-top: 4px;
  font-size: 17px;
  font-weight: 800;
  color: #D6B66A;
  line-height: 1.15;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
}
.subheadline {
  margin-top: 5px;
  font-size: 10.5px;
  opacity: 0.9;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
}
.banner-right {
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  align-items: flex-end;
  text-align: right;
  min-width: 0;
}
.ns-logo {
  max-width: 110px;
  max-height: 38px;
  object-fit: contain;
}
.contact {
  font-size: 9.5px;
  line-height: 1.35;
  opacity: 0.9;
}

.hear {
  background: #FFFFFF;
  padding: 0.15in 0.25in;
  box-sizing: border-box;
  border-bottom: 2px solid #AFC2D3;
}
.hear h2 {
  margin: 0 0 9px 0;
  font-size: 10px;
  font-weight: 700;
  letter-spacing: 1px;
  color: #425D73;
  text-transform: uppercase;
}
.hear-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 0.1in;
}
.hear-card {
  background: #F4F7F8;
  border: 1px solid #AFC2D3;
  border-left: 4px solid #2E4759;
  border-radius: 4px;
  padding: 0.1in 0.12in;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
  justify-content: center;
  gap: 4px;
}
.hear-card-title {
  font-size: 12px;
  font-weight: 800;
  color: #2E4759;
  line-height: 1.2;
}
.hear-card-consequence {
  font-size: 9.5px;
  color: #425D73;
  line-height: 1.35;
}

.cso-section {
  flex: 1;
  display: flex;
  flex-direction: column;
  padding: 0.15in 0.25in 0.12in;
  box-sizing: border-box;
  overflow: hidden;
}
.cso-header {
  display: grid;
  grid-template-columns: 1fr 1fr 1fr;
  gap: 0.12in;
  margin-bottom: 0.07in;
}
.col-label {
  font-size: 10px;
  font-weight: 700;
  letter-spacing: 1px;
  text-transform: uppercase;
  padding: 5px 10px;
  border-radius: 4px 4px 0 0;
  text-align: center;
}
.col-label.challenge { background: #2E4759; color: #F4F7F8; }
.col-label.solution  { background: #425D73; color: #F4F7F8; }
.col-label.outcome   { background: #D6B66A; color: #2E4759; }

.cso-grid {
  display: grid;
  grid-template-columns: 1fr 1fr 1fr;
  grid-template-rows: repeat(4, 1fr);
  gap: 0.07in 0.12in;
  flex: 1;
}
.cso-cell {
  border-radius: 5px;
  padding: 0.09in 0.11in;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
  justify-content: center;
  overflow: hidden;
}
.cell-title {
  font-size: 11px;
  font-weight: 700;
  margin-bottom: 4px;
  line-height: 1.2;
}
.cso-cell p {
  margin: 0;
  font-size: 9.5px;
  line-height: 1.35;
}
.cso-cell.challenge {
  background: #EEF2F5;
  border-left: 3px solid #2E4759;
}
.cso-cell.challenge p { color: #2E4759; }
.cso-cell.challenge .cell-title { color: #2E4759; }
.cso-cell.solution {
  background: #FFFFFF;
  border-left: 3px solid #425D73;
}
.cso-cell.solution p { color: #2E4759; }
.cso-cell.solution .cell-title { color: #425D73; }
.cso-cell.outcome {
  background: #FBF7EE;
  border-left: 3px solid #D6B66A;
}
.cso-cell.outcome p { color: #2E4759; }
.cso-cell.outcome .cell-title { color: #8B6914; }

/* ── ROI Section ── */
.roi-section {
  background: #2E4759;
  padding: 0.12in 0.25in;
  box-sizing: border-box;
}
.roi-header {
  display: flex;
  justify-content: space-between;
  align-items: baseline;
  margin-bottom: 0.08in;
}
.roi-title {
  font-size: 10px;
  font-weight: 700;
  letter-spacing: 1px;
  text-transform: uppercase;
  color: #AFC2D3;
}
.roi-disclaimer {
  font-size: 8.5px;
  color: #AFC2D3;
  opacity: 0.7;
  font-style: italic;
}
.roi-grid {
  display: grid;
  grid-template-columns: 1fr 1fr 1fr;
  gap: 0.1in;
}
.roi-card {
  background: #3D5A6E;
  border-radius: 6px;
  padding: 0.1in 0.12in;
  box-sizing: border-box;
}
.roi-range {
  font-size: 16px;
  font-weight: 800;
  color: #D6B66A;
  line-height: 1.1;
  margin-bottom: 3px;
}
.roi-label {
  font-size: 9px;
  color: #AFC2D3;
  line-height: 1.3;
  margin-bottom: 7px;
  opacity: 0.9;
}
.roi-bullets {
  list-style: none;
  margin: 0;
  padding: 0;
}
.roi-bullets li {
  font-size: 9px;
  color: #F4F7F8;
  line-height: 1.3;
  padding-left: 10px;
  position: relative;
  margin-bottom: 3px;
}
.roi-bullets li::before {
  content: "›";
  position: absolute;
  left: 0;
  color: #D6B66A;
  font-weight: 700;
}

/* ── Challenge / Solution Grid ── */
.cs-section {
  height: 2.4in;
  min-height: 2.4in;
  max-height: 2.4in;
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 0;
  overflow: hidden;
}
.cs-column {
  display: flex;
  flex-direction: column;
  padding: 0.08in 0.12in;
  box-sizing: border-box;
  gap: 0.04in;
  overflow: hidden;
}
.cs-column-header {
  font-size: 9px;
  font-weight: 700;
  letter-spacing: 1px;
  text-transform: uppercase;
  padding: 4px 8px;
  border-radius: 3px;
  margin-bottom: 0.04in;
  text-align: center;
}
.cs-column-header.challenge { background: #2E4759; color: #F4F7F8; }
.cs-column-header.solution  { background: #425D73; color: #F4F7F8; }
.cs-card {
  flex: 1;
  border-radius: 4px;
  padding: 0.06in 0.09in;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
  justify-content: center;
  overflow: hidden;
  min-height: 0;
}
.cs-card.challenge {
  background: #EEF2F5;
  border-left: 3px solid #2E4759;
}
.cs-card.solution {
  background: #FFFFFF;
  border-left: 3px solid #425D73;
}
.cs-card.challenge.known {
  background: #FBF6EC;
  border-left: 3px solid #D6B66A;
}
.cs-card.solution.known {
  background: #FFFDF5;
  border-left: 3px solid #D6B66A;
}
.known-label {
  font-size: 7.5px;
  font-weight: 700;
  letter-spacing: 0.8px;
  text-transform: uppercase;
  color: #D6B66A;
  margin-bottom: 3px;
}
.cs-card-title {
  font-size: 9.5px;
  font-weight: 700;
  line-height: 1.2;
  margin-bottom: 2px;
  display: -webkit-box;
  -webkit-line-clamp: 1;
  -webkit-box-orient: vertical;
  overflow: hidden;
}
.cs-card.challenge .cs-card-title { color: #2E4759; }
.cs-card.solution .cs-card-title  { color: #425D73; }
.cs-card p {
  margin: 0;
  font-size: 8px;
  line-height: 1.25;
  color: #2E4759;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
}

/* ── North Star Outcomes ── */
.north-star {
  background: #425D73;
  height: 0.75in;
  min-height: 0.75in;
  max-height: 0.75in;
  padding: 0.08in 0.2in;
  box-sizing: border-box;
  display: grid;
  grid-template-columns: 36px 1fr 1fr;
  gap: 0.15in;
  align-items: center;
  overflow: hidden;
}
.ns-icon {
  display: flex;
  align-items: center;
  justify-content: center;
  opacity: 0.9;
}
.ns-item {
  display: flex;
  flex-direction: column;
  gap: 3px;
}
.ns-outcome {
  font-size: 12px;
  font-weight: 800;
  color: #D6B66A;
  line-height: 1.2;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
}
.ns-detail {
  font-size: 8.5px;
  color: #AFC2D3;
  line-height: 1.25;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
  overflow: hidden;
}

.bottom-band {
  height: 0.55in;
  min-height: 0.55in;
  max-height: 0.55in;
  background: #2E4759;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 0 0.35in;
  box-sizing: border-box;
}
.cta {
  font-size: 13px;
  font-weight: 700;
  color: #D6B66A;
  text-align: center;
  line-height: 1.4;
}
.cta span {
  display: block;
  font-size: 10px;
  font-weight: 400;
  color: #AFC2D3;
  margin-top: 4px;
}
"""

CLAUDE_SYSTEM_PROMPT = """You are a NetSuite sales copywriter. Return exactly 2 sections, no more.

HEADLINE:
[One punchy line, MAXIMUM 7 words. No em dashes. No filler words. Hard limit — if it is over 7 words, cut it.]

SUBHEADLINE:
[One sentence, 12-20 words, connecting their business model to what NetSuite fixes. No em dashes.]

Rules:
- No markdown, no HTML, no asterisks, no bullet points.
- No em dashes anywhere.
- No buzzwords: leverage, synergy, streamline, robust, scalable, cutting-edge.
- Return only HEADLINE and SUBHEADLINE. Nothing else.
"""


def load_instructions(path="gpt_instructions.txt") -> str:
    try:
        return open(path).read()
    except FileNotFoundError:
        return ""


def fetch_page_text(url: str) -> str:
    try:
        r = requests.get(url, timeout=8, headers={"User-Agent": "Mozilla/5.0"})
        soup = BeautifulSoup(r.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer"]):
            tag.decompose()
        return soup.get_text(separator=" ", strip=True)[:1500]
    except Exception:
        return ""


def fetch_site_context(domain: str) -> str:
    if not domain or str(domain).lower() in ("nan", "none", ""):
        return ""
    base = domain if domain.startswith("http") else f"https://{domain}"
    base = base.rstrip("/")

    # Scrape homepage + about + products/services pages
    pages_to_try = [
        base,
        base + "/about",
        base + "/about-us",
        base + "/products",
        base + "/services",
        base + "/what-we-do",
    ]

    chunks = []
    for url in pages_to_try:
        text = fetch_page_text(url)
        if text:
            chunks.append(text)
        if len(chunks) >= 3:
            break

    return "\n\n---\n\n".join(chunks)[:5000]


def safe_filename(name: str) -> str:
    name = re.sub(r"[^\w\s-]", "", str(name)).strip()
    name = re.sub(r"\s+", "_", name)
    return name[:80]


def clean_company_name(raw: str) -> str:
    name = str(raw or "").strip()
    # Strip leading CRM IDs (6+ digit numbers) but keep short brand numbers like "88 Tactical"
    name = re.sub(r"^\d{6,}\s+", "", name)
    # Strip trailing legal suffixes
    name = re.sub(r"\s+(LLC|Inc\.?|Corp\.?|Co\.?|Ltd\.?|L\.L\.C\.?|Incorporated|Corporation|formerly\s+\w+)\s*$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def generate_content(openai_client: OpenAI, instructions: str,
                     lead_name: str, domain: str, site_context: str, row: dict,
                     vertical_intel: str = "") -> dict:

    row_text = "\n".join(f"{k}: {v}" for k, v in row.items())

    vertical_context = ("Vertical Intelligence (use this to write specific, credible content):\n" + vertical_intel) if vertical_intel else ""
    current_system = str(row.get("current system", "")).strip()
    rep_notes = str(row.get("rep notes", "")).strip()
    system_context = f"Current system they are replacing: {current_system}" if current_system else ""
    notes_context = f"REP INTEL — HIGHEST PRIORITY: The rep has provided the following notes from actual discovery or research. This must directly shape the challenges, solutions, and north star outcomes. Do not ignore this. Do not treat it as background. Write as if you know these specific facts about the company:\n{rep_notes}" if rep_notes else ""
    user_msg = f"""Company: {lead_name}
Website: {domain}

CSV row data:
{row_text}

Website context (scraped):
{site_context or '(none available)'}

{vertical_context}
{system_context}
{notes_context}

Instructions:
{instructions}

---
You are building content for a NetSuite sales one-pager. Return a JSON object with exactly these keys:

{{
  "content_brief": "150-200 word brief: what the company does, who they sell to, operational complexity, top pain points for their vertical. Written as a brief to a copywriter.",
  "triplets": [
    {{
      "challenge_title": "2-4 words, names the SPECIFIC operational problem, not a category. Bad: 'Operational Efficiency'. Good: 'Rental Fleet Tracking'",
      "challenge": "One sentence. Must reference something specific to THIS company's business model or vertical. No generic ERP pain points. Do NOT mention the company name.",
      "solution_title": "2-4 words, names the SPECIFIC NetSuite capability that solves it. Bad: 'Automated Workflows'. Good: 'Real-Time Rental Tracking'",
      "solution": "One sentence. Name the specific NetSuite module or feature. No vague statements like 'NetSuite streamlines processes'.",
      "outcome_title": "2-4 words, a concrete result. Bad: 'Better Efficiency'. Good: 'Faster Season Closeout'",
      "outcome": "One sentence. A specific measurable result this type of company would care about.",
      "known": "true if this pain point came directly from rep intel or current system info, false if it is an industry-level assumption"
    }},
    {{"challenge_title": "...", "challenge": "...", "solution_title": "...", "solution": "...", "outcome_title": "...", "outcome": "...", "known": false}},
    {{"challenge_title": "...", "challenge": "...", "solution_title": "...", "solution": "...", "outcome_title": "...", "outcome": "...", "known": false}},
    {{"challenge_title": "...", "challenge": "...", "solution_title": "...", "solution": "...", "outcome_title": "...", "outcome": "...", "known": false}}
  ],
  "hear_bullets": [
    {{"title": "Short bold problem statement, 4-7 words, no verb endings like -ing", "consequence": "One sentence: what goes wrong because of this problem, specific to their vertical"}},
    {{"title": "...", "consequence": "..."}},
    {{"title": "...", "consequence": "..."}},
    {{"title": "...", "consequence": "..."}}
  ],
  "roi": {{
    "time_savings": {{
      "range": "Conservative % or $ range for labor/efficiency savings based on annual revenue",
      "label": "Reduction in manual finance and ops labor",
      "bullets": ["2-3 short bullets, specific outcomes, low-end estimates"]
    }},
    "working_capital": {{
      "range": "Conservative % or $ range for working capital or inventory improvement",
      "label": "Improvement in working capital management",
      "bullets": ["2-3 short bullets, specific outcomes, low-end estimates"]
    }},
    "system_consolidation": {{
      "range": "Conservative % or $ range for IT/system cost reduction",
      "label": "Reduction in IT overhead and system costs",
      "bullets": ["2-3 short bullets, specific outcomes, low-end estimates"]
    }}
  }},
  "north_star": [
    {{"outcome": "Bold 6-10 word business outcome, executive level, specific to their vertical", "detail": "One sentence supporting it, tangible, no buzzwords"}},
    {{"outcome": "Second bold 6-10 word business outcome, different angle from first", "detail": "One sentence supporting it, tangible, no buzzwords"}}
  ],
  "email": "Outbound email: 4-8 sentences, one timely hook, 2 value points, bold meeting ask. No em dashes. No buzzwords."
}}

CRITICAL RULES - violating these makes the output useless:
0. REP INTEL RULE — If "REP INTEL" is provided above, it overrides everything else. At least 2 of the 4 challenge/solution pairs must directly reflect what the rep told you. The north star outcomes must also reflect it. If the rep says the CFO is frustrated with month-end close, that goes in. If they say they just opened a new location, that goes in. Do not dilute it with generic content.

1. BANNED GENERIC TERMS - never use these as challenge or solution titles:
   "Operational Efficiency", "Customer Experience", "Business Growth", "Digital Transformation",
   "Automated Workflows", "Centralized Operations", "Streamlined Processes", "Better Visibility",
   "Scalability", "Data Analytics", "Real-Time Analytics", "Financial Visibility"
   If you are about to write any of these, STOP and think about what specifically makes THIS company hard to run.

2. SPECIFICITY TEST - before writing each challenge, ask: "Could this apply to ANY company in any industry?"
   If yes, it is too generic. Rewrite it to reference something specific to this company's products, customers, or operations.

3. SOLUTION NAMES must reference actual NetSuite modules or capabilities:
   Good examples: "NetSuite WMS", "SuiteCommerce", "NetSuite Manufacturing", "Multi-Location Inventory",
   "NetSuite Demand Planning", "SuitePeople", "NetSuite CRM", "Revenue Recognition module"
   Bad: "NetSuite automates your processes", "NetSuite provides visibility"

4. NORTH STAR outcomes must be executive-level and specific to their vertical.
   Bad: "Achieve seamless omni-channel retail experience"
   Good: "Cut peak-season stockouts by 40% across all resort locations"

5. Use the website context and vertical field to inform every specific detail.

6. CURRENT SYSTEM RULE — If a current system is provided, reference it by name in at least one challenge. Frame it as a specific limitation, not a generic one. Example: if they are on QuickBooks, write "QuickBooks has no native multi-location inventory" not "their current system lacks visibility."

Return only valid JSON, no markdown fences.
"""

    resp = openai_client.chat.completions.create(
        model="gpt-4o",
        temperature=0.5,
        messages=[
            {"role": "system", "content": "You are a B2B sales content expert. Return only valid JSON."},
            {"role": "user", "content": user_msg}
        ]
    )

    raw = resp.choices[0].message.content.strip()
    raw = re.sub(r"^```[a-zA-Z]*\s*", "", raw).strip()
    raw = re.sub(r"\s*```\s*$", "", raw).strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        data = {"content_brief": raw, "triplets": [], "hear_bullets": [], "email": ""}

    triplets = data.get("triplets", [])[:4]
    while len(triplets) < 4:
        triplets.append({"challenge": "Review needed", "solution": "", "outcome": ""})
    data["triplets"] = triplets

    raw_hear = data.get("hear_bullets", [])[:4]
    hear = []
    for b in raw_hear:
        if isinstance(b, dict):
            hear.append(b)
        elif isinstance(b, str) and b:
            hear.append({"title": b, "consequence": ""})
    data["hear_bullets"] = hear

    # Parse ROI
    roi = data.get("roi", {})
    if not roi:
        roi = {
            "time_savings": {"range": "15-25%", "label": "Reduction in manual finance labor", "bullets": ["Faster month-end close", "Automated reporting", "Reduced reconciliation time"]},
            "working_capital": {"range": "5-10%", "label": "Working capital improvement", "bullets": ["Better inventory visibility", "Improved AR/AP cycles", "Reduced carrying costs"]},
            "system_consolidation": {"range": "20-30%", "label": "IT overhead reduction", "bullets": ["Fewer disconnected systems", "Lower maintenance costs", "Unified platform"]},
        }
    data["roi"] = roi

    # Parse north star
    north_star = data.get("north_star", [])
    if not north_star or len(north_star) < 2:
        north_star = [
            {"outcome": "Scale operations without adding overhead", "detail": "NetSuite grows with your business so your team stays lean."},
            {"outcome": "Close the books faster every month", "detail": "Automated reconciliation and real-time reporting cut close time significantly."}
        ]
    data["north_star"] = north_star[:2]
    return data


def generate_headline(anthropic_client: Anthropic, content_brief: str) -> dict:

    msg = anthropic_client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=300,
        temperature=0.5,
        system=CLAUDE_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": f"Content brief:\n\n{content_brief}"}]
    )

    raw = msg.content[0].text
    text = re.sub(r"(?m)^#+\s*(HEADLINE|SUBHEADLINE)\s*:", r"\1:", raw)

    def grab(section):
        key = section + ":"
        start = text.find(key)
        if start == -1:
            return ""
        start += len(key)
        while start < len(text) and text[start] in (" ", "\t", "\r"):
            start += 1
        if start < len(text) and text[start] == "\n":
            start += 1
        other = "SUBHEADLINE" if section == "HEADLINE" else "HEADLINE"
        end = text.find("\n" + other + ":", start)
        return text[start:end].strip() if end != -1 else text[start:].strip()

    return {
        "headline": grab("HEADLINE"),
        "subheadline": grab("SUBHEADLINE"),
    }


def build_hear_html(bullets: list) -> str:
    cards = ""
    for b in bullets:
        title = b.get("title", "")
        consequence = b.get("consequence", "")
        cards += f"""
    <div class="hear-card">
      <div class="hear-card-title">{title}</div>
      <div class="hear-card-consequence">{consequence}</div>
    </div>"""
    return f"""
<div class="hear">
  <h2>What we hear from companies like yours</h2>
  <div class="hear-grid">{cards}
  </div>
</div>
"""


def build_cso_html(triplets: list) -> str:
    header = """
<div class="cso-header">
  <div class="col-label challenge">Challenge</div>
  <div class="col-label solution">NetSuite Solution</div>
  <div class="col-label outcome">Outcome</div>
</div>"""

    cells = ""
    for t in triplets:
        cells += f"""
  <div class="cso-cell challenge"><div class="cell-title">{t.get('challenge_title','')}</div><p>{t.get('challenge','')}</p></div>
  <div class="cso-cell solution"><div class="cell-title">{t.get('solution_title','')}</div><p>{t.get('solution','')}</p></div>
  <div class="cso-cell outcome"><div class="cell-title">{t.get('outcome_title','')}</div><p>{t.get('outcome','')}</p></div>"""

    return f"""
<div class="cso-section">
  {header}
  <div class="cso-grid">{cells}
  </div>
</div>
"""


def build_cs_html(triplets: list) -> str:
    challenges = ""
    solutions = ""
    for t in triplets:
        known = t.get('known', False)
        if isinstance(known, str):
            known = known.lower() == 'true'
        known_class = " known" if known else ""
        known_label_c = '<div class="known-label">● From our conversation</div>' if known else ""
        known_label_s = '<div class="known-label">● NetSuite native</div>' if known else ""
        challenges += f"""
    <div class="cs-card challenge{known_class}">
      {known_label_c}
      <div class="cs-card-title">{t.get('challenge_title','')}</div>
      <p>{t.get('challenge','')}</p>
    </div>"""
        solutions += f"""
    <div class="cs-card solution{known_class}">
      {known_label_s}
      <div class="cs-card-title">{t.get('solution_title','')}</div>
      <p>{t.get('solution','')}</p>
    </div>"""
    return f"""
<div class="cs-section">
  <div class="cs-column">
    <div class="cs-column-header challenge">Challenge</div>
    {challenges}
  </div>
  <div class="cs-column">
    <div class="cs-column-header solution">NetSuite Solution</div>
    {solutions}
  </div>
</div>
"""


COMPASS_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" width="28" height="28">
  <circle cx="50" cy="50" r="48" fill="none" stroke="#D6B66A" stroke-width="2"/>
  <polygon points="50,5 55,45 50,50 45,45" fill="#D6B66A"/>
  <polygon points="50,95 55,55 50,50 45,55" fill="#AFC2D3"/>
  <polygon points="5,50 45,45 50,50 45,55" fill="#AFC2D3"/>
  <polygon points="95,50 55,45 50,50 55,55" fill="#D6B66A"/>
  <circle cx="50" cy="50" r="5" fill="#D6B66A"/>
</svg>"""

def build_north_star_html(north_star: list) -> str:
    if not north_star:
        return ""
    items = ""
    for ns in north_star[:2]:
        items += f"""
  <div class="ns-item">
    <div class="ns-outcome">{ns.get('outcome','')}</div>
    <div class="ns-detail">{ns.get('detail','')}</div>
  </div>"""
    return f"""
<div class="north-star">
  <div class="ns-icon">{COMPASS_SVG}</div>
  {items}
</div>
"""


def build_roi_html(roi: dict) -> str:
    def card(key, title):
        d = roi.get(key, {})
        rng = d.get("range", "")
        label = d.get("label", "")
        bullets = d.get("bullets", [])
        if isinstance(bullets, str):
            bullets = [bullets]
        bullet_html = "".join(f"<li>{b}</li>" for b in bullets[:3])
        return f"""
    <div class="roi-card">
      <div class="roi-range">{rng}</div>
      <div class="roi-label">{label}</div>
      <ul class="roi-bullets">{bullet_html}</ul>
    </div>"""

    cards = (
        card("time_savings", "Time Savings") +
        card("working_capital", "Working Capital") +
        card("system_consolidation", "System Consolidation")
    )
    return f"""
<div class="roi-section">
  <div class="roi-header">
    <div class="roi-title">Financial Impact of Moving to NetSuite</div>
    <div class="roi-disclaimer">Industry average estimates only</div>
  </div>
  <div class="roi-grid">{cards}
  </div>
</div>
"""


def build_html_page(company_name: str, headline: str, subheadline: str,
                    hear_bullets: list, triplets: list, roi: dict, rep: dict = None, north_star: list = None,
                    cta: str = "") -> str:
    rep = rep or DEFAULT_REP
    cta_text = cta if cta else "Open to 15 minutes next week to map this to your current process?"
    company_logo_html = ""
    banner = BANNER_HTML_TEMPLATE.format(
        company_name=company_name,
        headline=headline,
        subheadline=subheadline,
        logo_url=NETSUITE_LOGO_URL,
    )
    cs = build_cs_html(triplets)
    north_star_html = build_north_star_html(north_star or [])

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <style>{PAGE_CSS}</style>
</head>
<body>
<div class="page">
  {banner}
  {cs}
  {north_star_html}
  <div class="bottom-band">
    <div class="cta">
      {cta_text}
      <span>{rep["name"]} &nbsp;|&nbsp; {rep["email"]} &nbsp;|&nbsp; {rep["phone"]} &nbsp;|&nbsp; <a href="https://www.netsuite.com/portal/resource/demo.shtml" style="color:#D6B66A;">Request a Demo</a></span>
    </div>
  </div>
</div>
</body>
</html>"""


def html_to_pdf(html: str, path: str):
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page()
        page.set_content(html, wait_until="networkidle")
        page.pdf(path=path, width="8.5in", height="5.5in", print_background=True)
        browser.close()


def main():
    if not os.environ.get("OPENAI_API_KEY"):
        raise RuntimeError("OPENAI_API_KEY not set")
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise RuntimeError("ANTHROPIC_API_KEY not set")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    df = pd.read_csv(CSV_PATH, encoding="latin1")
    df.columns = [c.strip().lower() for c in df.columns]

    # Filter junk rows
    name_col = df.columns[0]
    domain_col = df.columns[1]
    df = df[df[domain_col].notna() & (df[domain_col].str.strip() != "")]
    df = df[~df[name_col].str.contains("[|:]", na=False, regex=True)]
    df = df.reset_index(drop=True)

    rep_col = next((c for c in df.columns if "sales" in c and "rep" in c), None)
    if rep_col:
        df = df.groupby(rep_col).head(1).head(N).reset_index(drop=True)
    else:
        df = df.head(N)

    instructions = load_instructions()
    openai_client = OpenAI()
    anthropic_client = Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    for _, row in df.iterrows():
        row_dict = row.to_dict()
        lead_name = str(row.iloc[0])
        domain = str(row.iloc[1]) if len(row) > 1 else ""

        try:
            site_context = fetch_site_context(domain)
            vertical = str(row_dict.get("vertical", "")).strip()
            vertical_intel = get_vertical_intel(vertical)
            content = generate_content(openai_client, instructions, lead_name, domain, site_context, row_dict, vertical_intel)
            headline_data = generate_headline(anthropic_client, content["content_brief"])

            company_name = clean_company_name(lead_name)
            rep_field = str(row_dict.get("sales rep", "")).strip()
            rep = get_rep(rep_field)
            base = safe_filename(lead_name)

            with open(os.path.join(OUTPUT_DIR, base + ".email.txt"), "w") as f:
                f.write(content.get("email", ""))

            html = build_html_page(
                company_name=company_name,
                headline=headline_data["headline"],
                subheadline=headline_data["subheadline"],
                hear_bullets=content["hear_bullets"],
                triplets=content["triplets"],
                roi=content["roi"],
                rep=rep,
                north_star=content.get("north_star", []),
            )

            html_to_pdf(html, os.path.join(OUTPUT_DIR, base + ".pdf"))
            print(f"Saved: {base}")

        except Exception as e:
            print(f"ERROR on {lead_name}: {e}")
            import traceback; traceback.print_exc()


if __name__ == "__main__":
    main()
