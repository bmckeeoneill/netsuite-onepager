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
CSV_PATH = "team_tal.csv"
N = 7

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
    <div class="contact">
      {rep_name}<br>
      {rep_title}<br>
      {rep_email}<br>
      {rep_phone}
    </div>
  </div>
</div>
"""

PAGE_CSS = """
@page { size: Letter; margin: 0.5in; }
html, body { margin: 0; padding: 0; }

.page {
  width: 7.5in;
  height: 10in;
  overflow: hidden;
  box-sizing: border-box;
  background: #F4F7F8;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
  display: flex;
  flex-direction: column;
}

.banner {
  height: 1.4in;
  min-height: 1.4in;
  max-height: 1.4in;
  display: grid;
  grid-template-columns: 5fr 1.4fr;
  column-gap: 0.2in;
  background: #2E4759;
  color: #F4F7F8;
  padding: 0.2in 0.25in;
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

.bottom-band {
  height: 0.9in;
  min-height: 0.9in;
  max-height: 0.9in;
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


def fetch_site_context(domain: str) -> str:
    if not domain or str(domain).lower() in ("nan", "none", ""):
        return ""
    url = domain if domain.startswith("http") else f"https://{domain}"
    try:
        r = requests.get(url, timeout=8, headers={"User-Agent": "Mozilla/5.0"})
        soup = BeautifulSoup(r.text, "html.parser")
        return soup.get_text(separator=" ", strip=True)[:3000]
    except Exception:
        return ""


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
                     lead_name: str, domain: str, site_context: str, row: dict) -> dict:

    row_text = "\n".join(f"{k}: {v}" for k, v in row.items())

    user_msg = f"""Company: {lead_name}
Website: {domain}

CSV row data:
{row_text}

Website context (scraped):
{site_context or '(none available)'}

Instructions:
{instructions}

---
You are building content for a NetSuite sales one-pager. Return a JSON object with exactly these keys:

{{
  "content_brief": "150-200 word brief: what the company does, who they sell to, operational complexity, top pain points for their vertical. Written as a brief to a copywriter.",
  "triplets": [
    {{
      "challenge_title": "2-4 word bold label for the challenge",
      "challenge": "One sentence expanding on it, specific to their vertical, no em dashes",
      "solution_title": "2-4 word bold label for the solution",
      "solution": "One sentence, what NetSuite does about it, concrete, no buzzwords",
      "outcome_title": "2-4 word bold result label",
      "outcome": "One sentence, tangible metric or business impact"
    }},
    {{
      "challenge_title": "...", "challenge": "...",
      "solution_title": "...", "solution": "...",
      "outcome_title": "...", "outcome": "..."
    }},
    {{
      "challenge_title": "...", "challenge": "...",
      "solution_title": "...", "solution": "...",
      "outcome_title": "...", "outcome": "..."
    }},
    {{
      "challenge_title": "...", "challenge": "...",
      "solution_title": "...", "solution": "...",
      "outcome_title": "...", "outcome": "..."
    }}
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
  "email": "Outbound email: 4-8 sentences, one timely hook, 2 value points, bold meeting ask. No em dashes. No buzzwords."
}}

Rules for ROI section:
- Use the annual revenue from the CSV to size all ranges conservatively (low end only).
- Time Savings: focus on finance labor, month-end close, reporting hours.
- Working Capital: focus on inventory turns, cash flow, AR/AP cycle time.
- System Consolidation: focus on eliminated software, IT maintenance, audit costs.
- Keep bullet text short, 6-10 words each.
- Express ranges as percentages or dollar ranges, e.g. "15-25%" or "$50K-$100K".

Return only valid JSON, no markdown fences.
"""

    resp = openai_client.chat.completions.create(
        model="gpt-4o",
        temperature=0.3,
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
    return data


def generate_headline(anthropic_client: Anthropic, content_brief: str) -> dict:

    msg = anthropic_client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=300,
        temperature=0.3,
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
                    hear_bullets: list, triplets: list, roi: dict, rep: dict = None) -> str:
    rep = rep or DEFAULT_REP
    banner = BANNER_HTML_TEMPLATE.format(
        company_name=company_name,
        headline=headline,
        subheadline=subheadline,
        logo_url=NETSUITE_LOGO_URL,
        rep_name=rep["name"],
        rep_title=rep["title"],
        rep_email=rep["email"],
        rep_phone=rep["phone"],
    )
    hear = build_hear_html(hear_bullets)
    cso  = build_cso_html(triplets)
    roi_html = build_roi_html(roi)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <style>{PAGE_CSS}</style>
</head>
<body>
<div class="page">
  {banner}
  {hear}
  {cso}
  {roi_html}
  <div class="bottom-band">
    <div class="cta">
      Open to 15 minutes next week to map this to your current process?
      <span>{rep["name"]} &nbsp;|&nbsp; {rep["email"]} &nbsp;|&nbsp; {rep["phone"]}</span>
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
        page.pdf(path=path, format="Letter", print_background=True)
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
    df = df.iloc[50:].reset_index(drop=True)

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
            content = generate_content(openai_client, instructions, lead_name, domain, site_context, row_dict)
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
            )

            html_to_pdf(html, os.path.join(OUTPUT_DIR, base + ".pdf"))
            print(f"Saved: {base}")

        except Exception as e:
            print(f"ERROR on {lead_name}: {e}")
            import traceback; traceback.print_exc()


if __name__ == "__main__":
    main()
