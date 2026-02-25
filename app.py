import os
import sys
import tempfile
import subprocess
import streamlit as st

# Install Playwright browser on first run
subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=False)

# ── Page config ──
st.set_page_config(page_title="NetSuite One-Pager Generator", page_icon="🔷", layout="centered")

# ── Password gate ──
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True
    st.title("NetSuite One-Pager Generator")
    pwd = st.text_input("Enter passcode", type="password")
    if st.button("Enter"):
        if pwd == os.environ.get("APP_PASSCODE", "netsuite2026"):
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect passcode")
    return False

if not check_password():
    st.stop()

# ── Import pipeline ──
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_half_page import (
    get_vertical_intel, get_rep, load_instructions,
    fetch_site_context, safe_filename, clean_company_name,
    generate_content, generate_headline, build_html_page, html_to_pdf
)
from openai import OpenAI
from anthropic import Anthropic

# ── UI ──
st.title("🔷 NetSuite One-Pager Generator")
st.caption("Enter a company name and website to generate a personalized half-page PDF.")

with st.form("generator"):
    company_name = st.text_input("Company Name", placeholder="e.g. Christy Sports")
    website = st.text_input("Website URL", placeholder="e.g. https://www.christysports.com")
    vertical = st.selectbox("Vertical", [
        "Manufacturing", "Retail", "Apparel, Footwear & Accessories",
        "Food & Beverage", "Health & Beauty", "Wholesale/Distribution",
        "Software", "Services", "General Business"
    ])
    rep_name = st.selectbox("Your Name", [
        "O'Neill, Brian", "Dambrosio, Thomas", "Corbett, Danielle",
        "Traywick, Reginald", "Uritis, Peter", "Dynek, Christopher", "Zapalac, Ross"
    ])
    submitted = st.form_submit_button("Generate One-Pager")

if submitted:
    if not company_name or not website:
        st.error("Please enter both a company name and website.")
        st.stop()

    openai_key = os.environ.get("OPENAI_API_KEY")
    anthropic_key = os.environ.get("ANTHROPIC_API_KEY")

    if not openai_key or not anthropic_key:
        st.error("API keys not configured. Please add them in Streamlit secrets.")
        st.stop()

    openai_client = OpenAI(api_key=openai_key)
    anthropic_client = Anthropic(api_key=anthropic_key)

    with st.spinner("Researching company website..."):
        site_context = fetch_site_context(website)

    with st.spinner("Generating content with GPT-4o..."):
        instructions = load_instructions()
        vertical_intel = get_vertical_intel(vertical)
        row_dict = {
            "name": company_name,
            "web address": website,
            "vertical": vertical,
            "sales rep": rep_name,
        }
        content = generate_content(
            openai_client, instructions, company_name,
            website, site_context, row_dict, vertical_intel
        )

    with st.spinner("Writing headline with Claude..."):
        headline_data = generate_headline(anthropic_client, content["content_brief"])

    with st.spinner("Building PDF..."):
        clean_name = clean_company_name(company_name)
        rep = get_rep(rep_name)
        html = build_html_page(
            company_name=clean_name,
            headline=headline_data["headline"],
            subheadline=headline_data["subheadline"],
            hear_bullets=content["hear_bullets"],
            triplets=content["triplets"],
            roi=content["roi"],
            rep=rep,
            north_star=content.get("north_star", []),
        )
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            pdf_path = tmp.name
        html_to_pdf(html, pdf_path)

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    filename = safe_filename(company_name) + "_half.pdf"
    st.success("✅ One-pager ready!")
    st.download_button(
        label="⬇️ Download PDF",
        data=pdf_bytes,
        file_name=filename,
        mime="application/pdf"
    )
