#!/usr/bin/env python3
"""
FS Vector Proposal Generator — AI-Powered Streamlit Web App

A password-protected web interface that uses Claude AI to generate
professional client proposals from simple inputs.

Supports three proposal types:
  - Regulatory Compliance Advisory Services (Parent template)
  - Bank Charter Advisory Services (Onyx template)
  - Licensing Services (Licensing template)

Run with:
    export ANTHROPIC_API_KEY=sk-ant-...
    python3 -m streamlit run app.py
"""

import json
import logging
import os
import re
import sys
import tempfile
from datetime import date

logging.basicConfig(stream=sys.stdout, level=logging.INFO)
logger = logging.getLogger(__name__)
logger.info("Starting app.py imports...")

try:
    import streamlit as st
    logger.info("Streamlit imported OK")
except Exception as e:
    logger.error(f"Failed to import streamlit: {e}")
    raise

try:
    from generate_proposal import generate_proposal
    logger.info("generate_proposal imported OK")
except Exception as e:
    logger.error(f"Failed to import generate_proposal: {e}")
    raise

try:
    from generate_charter_proposal import generate_charter_proposal
    logger.info("generate_charter_proposal imported OK")
except Exception as e:
    logger.error(f"Failed to import generate_charter_proposal: {e}")
    raise

try:
    from generate_licensing_proposal import generate_licensing_proposal
    logger.info("generate_licensing_proposal imported OK")
except Exception as e:
    logger.error(f"Failed to import generate_licensing_proposal: {e}")
    raise

logger.info("All imports successful")

# ── Config ────────────────────────────────────────────────────────────
APP_PASSWORD = "JediMasterGentle2026!"

PROPOSAL_TYPES = [
    "Regulatory Compliance Advisory Services",
    "Bank Charter Advisory Services",
    "Licensing Services",
]

st.set_page_config(
    page_title="Proposal Generator",
    page_icon=":shield:",
    layout="centered",
)

# ── FSV Brand Assets ─────────────────────────────────────────────────
FSV_LOGO_URL = (
    "https://cdn.prod.website-files.com/68754b350a99a7f49cc40a2e/"
    "689f4ef9914755e4ea5cbe08_Logo.svg"
)

# ── Custom CSS (applied conditionally after login) ───────────────────
_LOGIN_CSS = """
<style>
    /* Minimal styling for the non-descript login page */
    .stTextInput > label { font-weight: 500; }
</style>
"""

_APP_CSS = """
<style>
    /* ── FSV Brand Theme ─────────────────────────────────────── */

    /* Global background */
    .stApp {
        background-color: #F2F2FA;
    }

    /* Main content area */
    .block-container {
        background-color: #FFFFFF;
        border-radius: 12px;
        padding: 2rem 3rem 3rem 3rem !important;
        margin-top: 1.5rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }

    /* Typography */
    h1, h2, h3 {
        color: #000839 !important;
        font-family: 'Arial', sans-serif;
    }
    h1 { font-weight: 500 !important; }
    label, .stTextInput label, .stTextArea label, .stSelectbox label,
    .stCheckbox label, .stDateInput label,
    .block-container > div > div > div > div > p {
        color: #39386A !important;
    }

    /* Dividers */
    hr {
        border-color: #DFE2F0 !important;
    }

    /* Primary buttons (Generate, Revise) */
    .stButton > button[kind="primary"],
    button[data-testid="stBaseButton-primary"] {
        background-color: #39386A !important;
        color: #FFFFFF !important;
        border: none !important;
        font-weight: 500;
        border-radius: 6px;
    }
    .stButton > button[kind="primary"]:hover,
    button[data-testid="stBaseButton-primary"]:hover {
        background-color: #000839 !important;
        color: #FFFFFF !important;
    }

    /* Secondary buttons */
    .stButton > button:not([kind="primary"]),
    button[data-testid="stBaseButton-secondary"] {
        background-color: transparent !important;
        color: #39386A !important;
        border: 1px solid #9FA5C3 !important;
        border-radius: 6px;
    }
    .stButton > button:not([kind="primary"]):hover,
    button[data-testid="stBaseButton-secondary"]:hover {
        background-color: #DFE2F0 !important;
        color: #39386A !important;
    }

    /* Download button — green accent */
    .stDownloadButton > button {
        background-color: #55A745 !important;
        color: #FFFFFF !important;
        font-weight: 500;
        border: none !important;
        border-radius: 6px;
    }
    .stDownloadButton > button:hover {
        background-color: #468A38 !important;
        color: #FFFFFF !important;
    }

    /* Text inputs & text areas */
    .stTextInput input, .stTextArea textarea, .stSelectbox > div > div {
        border-color: #DFE2F0 !important;
        color: #000839 !important;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #9FA5C3 !important;
        box-shadow: 0 0 0 1px #9FA5C3 !important;
    }

    /* Expander */
    .streamlit-expanderHeader {
        color: #39386A !important;
        border-color: #DFE2F0 !important;
    }

    /* Success message */
    .stAlert [data-testid="stNotification"] {
        border-left-color: #55A745 !important;
    }

    /* Caption text */
    .stCaption, [data-testid="stCaptionContainer"] {
        color: #9FA5C3 !important;
    }

    /* Green accent bar at top */
    .block-container::before {
        content: '';
        display: block;
        height: 3px;
        background-color: #55A745;
        margin: -2rem -3rem 1.5rem -3rem;
        border-radius: 12px 12px 0 0;
    }
</style>
"""


# ── Session State Defaults ────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "generated_file" not in st.session_state:
    st.session_state.generated_file = None
if "generated_filename" not in st.session_state:
    st.session_state.generated_filename = None
if "ai_content" not in st.session_state:
    st.session_state.ai_content = None
if "form_inputs" not in st.session_state:
    st.session_state.form_inputs = None


# =====================================================================
# AI SYSTEM PROMPTS
# =====================================================================

COMPLIANCE_SYSTEM_PROMPT = """You are a proposal content writer for FS Vector, a regulatory compliance advisory firm specializing in fintech, banking, and financial services compliance.

Your job is to generate structured proposal content based on brief client descriptions. Write in a professional, consultative tone. Be specific to the client's industry and regulatory needs. Do not use filler — every sentence should add value.

You must return ONLY valid JSON (no markdown, no code fences) with this exact structure:

{
  "engagement_background_paragraphs": [
    "First paragraph introducing the client and the context for the engagement...",
    "Second paragraph describing FS Vector's relevant expertise and why they are a good fit..."
  ],
  "engagement_scope_sections": [
    {
      "heading": "Section Heading (e.g., Partnership Engagement)",
      "intro": "An introductory paragraph for this scope area.",
      "bullets": [
        "Specific deliverable or activity",
        "Another deliverable or activity"
      ],
      "body_paragraphs": ["Optional follow-up paragraph after bullets."]
    }
  ],
  "budget_timeline_intro": "A single paragraph describing the proposed fee structure and engagement duration.",
  "budget_timeline_table_data": {
    "columns": ["Month", "1", "2", "3", "4", "5+"],
    "rows": [
      {
        "Month": "Advisory Services",
        "1": "Activity 1\\n\\nActivity 2",
        "2": "Activity 3\\n\\nActivity 4",
        "3": "Activity 5\\n\\nActivity 6",
        "4": "Activity 7\\n\\nActivity 8",
        "5+": "Ongoing Advisory Services"
      },
      {
        "Month": "Monthly Fee",
        "1": "$15,000",
        "2": "$15,000",
        "3": "$17,500",
        "4": "$17,500",
        "5+": "Per Client's needs"
      }
    ]
  },
  "budget_timeline_followup": [
    "Optional follow-up paragraph about next steps or payment terms."
  ],
  "include_compliance_chart": true,
  "compliance_column": "Deposit Compliance",
  "compliance_sub_items": [
    "Relevant compliance sub-item 1",
    "Relevant compliance sub-item 2",
    "Relevant compliance sub-item 3",
    "Relevant compliance sub-item 4",
    "Relevant compliance sub-item 5"
  ]
}

CRITICAL — Timeline table structure:
The timeline table MUST follow this exact format modeled after FS Vector's standard proposal:
- First column header is always "Month"
- Remaining columns are INDIVIDUAL MONTHS (e.g., "1", "2", "3", "4", "5+"). Adjust the number of month columns to match the engagement duration provided.
  - For a 5-month engagement: "1", "2", "3", "4", "5+"
  - For a 3-month engagement: "1", "2", "3"
  - For a 6-month engagement: "1", "2", "3", "4", "5", "6"
- First data row: "Month" cell = "Advisory Services". Each month cell lists the key activities for that month, separated by \\n\\n (double newline). Activities should be specific, actionable items like "Bank Search & Due Diligence", "Product Advisory", "Compliance Policy & Procedure Development", "Marketing Collateral Review", etc.
- Last data row: "Month" cell = "Monthly Fee". Each month cell contains the dollar amount (e.g., "$15,000"). The final period can say "Per [Client]'s needs" if the engagement may extend.
- You may add additional rows between Advisory Services and Monthly Fee if needed (e.g., "Managed Services" row).

Other guidelines:
- engagement_background_paragraphs: Write 2-3 substantial paragraphs. First paragraph should introduce the client and what they're doing. Second should describe the regulatory landscape and challenges. Third (optional) should explain why FS Vector is the right partner.
- engagement_scope_sections: Create one section per scope area requested. Each should have a heading, an intro paragraph, 4-8 specific bullet points, and optionally a closing paragraph.
- budget_timeline_intro: Reference the specific fee and duration provided. Mention what's included.
- budget_timeline_followup: 1-2 paragraphs about payment terms, next steps, or additional notes.
- include_compliance_chart: Set to true ONLY if the engagement includes a compliance program build, compliance gap analysis, or similar compliance-focused scope. Set to false for engagements that are purely strategic, operational, partnership-focused, or do not involve building/assessing compliance programs. When false, the compliance_column and compliance_sub_items fields are ignored.
- compliance_column: (Only when include_compliance_chart is true) Choose the most appropriate compliance type based on the product AND customer type:
  - Commercial: "Commercial Banking Compliance", "Treasury Management Compliance", "Commercial Lending Compliance"
  - Consumer: "Deposit Compliance", "Consumer Lending Compliance", "Payments and Accounts Compliance"
  - Both: Use the compliance type most relevant to the primary product, covering both commercial and consumer aspects.
- compliance_sub_items: (Only when include_compliance_chart is true) List 5 specific regulatory compliance areas relevant to the client's product and customer type.

Customer Type Guidance:
When the user specifies a customer type, tailor ALL content accordingly:
- "Commercial": Focus on BSA/AML, OFAC sanctions, UCC, ACH/Nacha rules, wire transfer regulations, commercial lending regulations, treasury management, commercial due diligence, CDD/EDD for business entities. Do NOT include consumer-specific regulations (TILA, ECOA, FCRA, FDCPA, RESPA, Fair Lending, CRA, Reg Z, Reg B) unless they also apply in a commercial context. Compliance sub-items should reference commercial-relevant regulations only.
- "Consumer": Focus on the full consumer protection suite: BSA/AML, OFAC, TILA, ECOA, FCRA, FDCPA, RESPA, UDAAP, Fair Lending, Reg Z, Reg B, CRA, Reg E, Reg DD, Reg CC, FDIC requirements. Emphasize credit underwriting, consumer complaints, fair lending, servicing/collections where relevant.
- "Both": Include regulations from both commercial and consumer segments. Engagement scope should address both segments explicitly. Compliance sub-items should cover the broadest range of applicable regulations.

IMPORTANT: Never reference PCI (Payment Card Industry), cardholder data security, cardholder operations, or Money Transmission Licensing (MTL) in any proposal content.

Write with authority and specificity. Reference actual regulations (BSA/AML, FDIC, OCC, state regulations, etc.) where appropriate. The proposals should sound like they were written by an experienced compliance consultant, not generic AI."""


CHARTER_SYSTEM_PROMPT = """You are a proposal content writer for FS Vector, a regulatory compliance advisory firm specializing in bank charter applications, regulatory strategy, and financial services compliance.

Your job is to generate structured bank charter advisory proposal content based on brief client descriptions. Write in a professional, consultative tone with deep knowledge of bank chartering processes, regulatory requirements, and the OCC/FDIC/state banking department application processes. Do not use filler — every sentence should add value.

You must return ONLY valid JSON (no markdown, no code fences) with this exact structure:

{
  "engagement_background_paragraphs": [
    "First paragraph introducing the client and their charter ambitions...",
    "Second paragraph describing FS Vector's charter expertise and relevant experience..."
  ],
  "engagement_scope_phases": [
    {
      "heading": "Exploration, Analysis, and Planning",
      "intro": "During this phase, FS Vector will work with the client to...",
      "deliverables": [
        {
          "name": "Choice-of-Charter Analysis",
          "description": "Understand and analyze the key factors relevant to the decision of whether and how to pursue a bank charter, including...",
          "sub_items": [
            "Capabilities provided by various charters and licensing options",
            "Direct regulation of the business to be operated in the new bank",
            "Financial costs and timing of chartering"
          ]
        },
        {
          "name": "New Bank Strategy",
          "description": "Develop a strategy outlining the general nature of the proposed bank...",
          "sub_items": [
            "Product offerings and operating model",
            "Market and customer segments"
          ]
        }
      ],
      "fee_text": "For the services included in Phase 1, we propose a fee of $X. We anticipate that the deliverables will take approximately Y months.",
      "bridge_text": "After the conclusion of Phase 1, we will be prepared to begin Phase 2."
    }
  ]
}

Guidelines:
- engagement_background_paragraphs: Write 2-3 paragraphs. First paragraph introduces the client and their charter exploration/application. Second describes the regulatory landscape for their charter type. Third (optional) explains FS Vector's fit.
- engagement_scope_phases: Typically 2-4 phases. Common charter phases include:
  - Phase 1: Exploration, Analysis, and Planning (choice-of-charter analysis, new bank strategy, exploratory regulatory discussions)
  - Phase 2: Pre-Filing Charter Application Support (application preparation, business plan, GRC framework, pre-filing regulatory discussions)
  - Phase 3: Post-Filing Charter Application Support (post-filing information requests, GRC framework buildout, infrastructure, field investigation prep)
- Each phase has deliverables with bold names, descriptions, and specific sub-items as bullets.
- fee_text: Reference the specific fee amount for that phase. Be specific about duration estimates.
- bridge_text: Optional transition paragraph between phases. Mention retainer options if there's a gap between phases.
- Deliverable names should be concise and specific (e.g., "Choice-of-Charter Analysis", "Application Preparation, including Business Plan", "Governance, Risk, and Compliance Framework", "Pre-Filing Regulatory Discussions").
- Sub-items should be specific regulatory and operational items, not generic bullets.

Write with authority about bank chartering. Reference specific charter types (national bank, state bank, industrial loan company, trust company, SPDI, etc.), regulators (OCC, FDIC, state banking departments, Federal Reserve), and regulatory processes (pre-filing meetings, field investigations, pre-opening exams) where appropriate."""


LICENSING_SYSTEM_PROMPT = """You are a proposal content writer for FS Vector, a regulatory advisory firm specializing in licensing, compliance, and regulatory strategy for fintech, crypto, blockchain, and financial services companies.

Your job is to generate ONLY the engagement background section for a licensing services proposal. All scope, fees, and timeline content is pre-written in the template — you are only generating the introductory section that describes the client and why they need licensing support.

You must return ONLY valid JSON (no markdown, no code fences) with this exact structure:

{
  "engagement_background_paragraphs": [
    "First paragraph introducing the client, their business model, and products/services...",
    "Second paragraph describing the specific licensing needs, regulatory landscape, and jurisdictions involved...",
    "Third paragraph introducing FS Vector's expertise and how they will support the client's licensing objectives...",
    "Optional fourth paragraph with additional context about timeline, urgency, or strategic considerations..."
  ]
}

Guidelines:
- Write 3-5 substantial paragraphs.
- First paragraph: Introduce the client — who they are, what they do, their industry, and their current stage (startup, expanding, entering new markets, etc.).
- Second paragraph: Describe the specific licensing needs — which licenses they need (money transmitter licenses, BitLicense, state lending licenses, etc.), which jurisdictions, and the regulatory landscape they face.
- Third paragraph: Introduce FS Vector's licensing expertise — their track record with similar engagements, their approach to multi-state licensing, and how they will guide the client through the process.
- Optional additional paragraphs: Cover strategic considerations, timeline urgency, or any unique aspects of the engagement.
- Write in a professional, consultative tone with authority about licensing requirements.
- Reference specific regulators (state banking departments, NMLS, DFS, FinCEN, etc.) and licensing frameworks where appropriate.
- Do NOT include scope, fees, deliverables, or timeline information — these are static in the template.
- Do NOT reference PCI, cardholder data security, or compliance program builds — this is a licensing-focused proposal."""


# =====================================================================
# AI FUNCTIONS
# =====================================================================

def get_anthropic_client():
    """Get the Anthropic client, checking for API key."""
    api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()

    # Check Streamlit secrets (used by Streamlit Cloud)
    if not api_key:
        try:
            api_key = st.secrets.get("ANTHROPIC_API_KEY", "").strip()
        except Exception:
            pass

    # Check if entered via the UI
    if not api_key:
        api_key = st.session_state.get("api_key_input", "").strip()

    if not api_key:
        return None

    try:
        import anthropic
        return anthropic.Anthropic(api_key=api_key)
    except ImportError:
        st.error("The `anthropic` package is not installed. Run: `pip3 install anthropic`")
        return None


def _call_ai(system_prompt, user_message):
    """Call Claude API and return parsed JSON."""
    client = get_anthropic_client()
    if client is None:
        raise ValueError(
            "Anthropic API key not found. Set the ANTHROPIC_API_KEY environment "
            "variable or enter it in the app."
        )

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}],
    )

    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)

    return json.loads(raw)


def generate_with_ai(client_name, description, scope_areas, budget_info,
                     proposal_type_key, compliance_col=None, additional_context=None,
                     customer_type=None):
    """Call Claude API to generate structured proposal content."""
    if proposal_type_key == "charter":
        system_prompt = CHARTER_SYSTEM_PROMPT
    elif proposal_type_key == "licensing":
        system_prompt = LICENSING_SYSTEM_PROMPT
    else:
        system_prompt = COMPLIANCE_SYSTEM_PROMPT

    if proposal_type_key == "licensing":
        # Licensing only needs client info and description — no scope/budget
        user_message = f"""Generate the engagement background for a licensing services proposal:

Client: {client_name}
Engagement Description: {description}"""
        if additional_context:
            user_message += f"\nAdditional Context: {additional_context}"
        user_message += "\n\nReturn the JSON now."
    else:
        user_message = f"""Generate proposal content for the following client:

Client: {client_name}
Engagement Description: {description}
Scope Areas: {scope_areas}
Budget: {budget_info}"""

        if customer_type and proposal_type_key == "compliance":
            user_message += f"\nCustomer Type: {customer_type}"

        if compliance_col and proposal_type_key == "compliance":
            user_message += f"\nCompliance Type: {compliance_col}"

        if additional_context:
            user_message += f"\nAdditional Context: {additional_context}"

        user_message += "\n\nReturn the JSON now."

    return _call_ai(system_prompt, user_message)


def revise_with_ai(previous_content, feedback, proposal_type_key):
    """Call Claude API to revise proposal content based on feedback."""
    if proposal_type_key == "charter":
        system_prompt = CHARTER_SYSTEM_PROMPT
    elif proposal_type_key == "licensing":
        system_prompt = LICENSING_SYSTEM_PROMPT
    else:
        system_prompt = COMPLIANCE_SYSTEM_PROMPT

    client = get_anthropic_client()
    if client is None:
        raise ValueError("Anthropic API key not found.")

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        system=system_prompt,
        messages=[
            {"role": "user", "content": "Generate proposal content for a client."},
            {"role": "assistant", "content": json.dumps(previous_content, indent=2)},
            {
                "role": "user",
                "content": (
                    f"Revise the proposal based on this feedback:\n\n{feedback}\n\n"
                    "Return the complete updated JSON with all fields. "
                    "Only change what the feedback asks for — keep everything else the same."
                ),
            },
        ],
    )

    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)

    return json.loads(raw)


# =====================================================================
# DOCUMENT BUILDERS
# =====================================================================

def _build_and_store_doc(ai_content, form_inputs):
    """Build the .docx from AI content and store in session state."""
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False, prefix="proposal_")
    tmp_path = tmp.name
    tmp.close()

    proposal_type_key = form_inputs.get("proposal_type_key", "compliance")

    if proposal_type_key == "charter":
        output = generate_charter_proposal(
            client_name=form_inputs["client_name"],
            proposal_type=form_inputs["proposal_type"],
            date_str=form_inputs["date_str"],
            engagement_background_paragraphs=ai_content["engagement_background_paragraphs"],
            engagement_scope_phases=ai_content["engagement_scope_phases"],
            client_website=form_inputs["client_website"],
            output_path=tmp_path,
        )
    elif proposal_type_key == "licensing":
        output = generate_licensing_proposal(
            client_name=form_inputs["client_name"],
            date_str=form_inputs["date_str"],
            engagement_background_paragraphs=ai_content["engagement_background_paragraphs"],
            client_website=form_inputs["client_website"],
            acquisition_fee=form_inputs.get("acquisition_fee"),
            maintenance_fee=form_inputs.get("maintenance_fee"),
            output_path=tmp_path,
        )
    else:
        skip_chart = not form_inputs.get("include_compliance", False)
        comp_col = ai_content.get("compliance_column")
        comp_items = ai_content.get("compliance_sub_items")

        if form_inputs.get("include_compliance"):
            user_comp = form_inputs.get("compliance_col", "Deposit Compliance")
            if user_comp and user_comp != "Deposit Compliance":
                comp_col = user_comp

        output = generate_proposal(
            client_name=form_inputs["client_name"],
            proposal_type=form_inputs["proposal_type"],
            date_str=form_inputs["date_str"],
            engagement_background_paragraphs=ai_content["engagement_background_paragraphs"],
            engagement_scope_sections=ai_content["engagement_scope_sections"],
            budget_timeline_intro=ai_content["budget_timeline_intro"],
            budget_timeline_table_data=ai_content.get("budget_timeline_table_data"),
            budget_timeline_followup=ai_content.get("budget_timeline_followup", []),
            client_website=form_inputs["client_website"],
            include_managed_services=form_inputs.get("include_managed", False),
            compliance_column=comp_col if comp_col != "Deposit Compliance" else None,
            compliance_sub_items=comp_items,
            skip_compliance_chart=skip_chart,
            customer_type=form_inputs.get("customer_type"),
            output_path=tmp_path,
        )

    with open(output, "rb") as f:
        file_bytes = f.read()
    os.unlink(output)

    st.session_state.generated_file = file_bytes
    safe_name = form_inputs["client_name"].replace(" ", "_")
    st.session_state.generated_filename = (
        f"FS_Vector_Proposal_{safe_name}_{form_inputs['date_str'].replace(' ', '_')}.docx"
    )
    st.rerun()


# =====================================================================
# LOGIN SCREEN
# =====================================================================
def login_screen():
    st.markdown(_LOGIN_CSS, unsafe_allow_html=True)
    st.markdown("---")
    st.markdown(
        "<h1 style='text-align:center; margin-bottom:0;'>Proposal Generator</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='text-align:center; color:#555; font-size:1.1rem;'>"
        "Proposal Generator</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        password = st.text_input("Enter password", type="password", key="pw_input")
        if st.button("Log In", use_container_width=True):
            if password == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password.")


# =====================================================================
# MAIN FORM
# =====================================================================
def main_form():
    st.markdown(_APP_CSS, unsafe_allow_html=True)

    # ── FSV Logo + Title ──────────────────────────────────────────────
    st.markdown(
        f"""
        <div style="display:flex; align-items:center; gap:14px; margin-bottom:4px;">
            <img src="{FSV_LOGO_URL}" alt="FS Vector" style="height:36px;">
            <span style="font-size:1.6rem; font-weight:500; color:#000839;">
                Proposal Generator
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption(
        "Fill in a few details about the client and engagement. "
        "AI will generate the full proposal content for you."
    )

    # ── API Key check ─────────────────────────────────────────────────
    api_key_env = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not api_key_env:
        try:
            api_key_env = st.secrets.get("ANTHROPIC_API_KEY", "").strip()
        except Exception:
            pass
    if not api_key_env:
        st.text_input(
            "Anthropic API Key *",
            type="password",
            key="api_key_input",
            placeholder="sk-ant-...",
            help="Get your key from console.anthropic.com/settings/keys",
        )

    # ── Proposal Type Selector ────────────────────────────────────────
    proposal_type_label = st.selectbox(
        "Proposal Type",
        PROPOSAL_TYPES,
        help="Choose the type of proposal to generate.",
    )
    if proposal_type_label == "Bank Charter Advisory Services":
        proposal_type_key = "charter"
    elif proposal_type_label == "Licensing Services":
        proposal_type_key = "licensing"
    else:
        proposal_type_key = "compliance"
    is_charter = proposal_type_key == "charter"
    is_licensing = proposal_type_key == "licensing"

    # ── 1. Client Details ─────────────────────────────────────────────
    st.header("Client Details")

    col_a, col_b = st.columns(2)
    with col_a:
        if is_licensing:
            name_placeholder = "e.g. ether.fi"
        elif is_charter:
            name_placeholder = "e.g. Onyx"
        else:
            name_placeholder = "e.g. Stripe"
        client_name = st.text_input(
            "Client Name *",
            placeholder=name_placeholder,
        )
    with col_b:
        client_website = st.text_input(
            "Client Website",
            placeholder="e.g. https://stripe.com",
            help="Used to fetch the client logo for the cover page.",
        )

    col_c, col_d = st.columns(2)
    with col_c:
        proposal_title = st.text_input(
            "Proposal Title",
            value=proposal_type_label,
        )
    with col_d:
        proposal_date = st.date_input("Proposal Date", value=date.today())

    # ── 2. Engagement Description ─────────────────────────────────────
    st.header("Engagement Details")

    if is_licensing:
        desc_placeholder = (
            "e.g. They are a crypto/blockchain company looking to obtain money transmitter "
            "licenses across multiple US states. They need help with license acquisition, "
            "NMLS filings, and ongoing license maintenance."
        )
    elif is_charter:
        desc_placeholder = (
            "e.g. They are exploring a bank charter as part of their US market "
            "strategy. They need help with choice-of-charter analysis, developing "
            "a new bank strategy, and navigating the application process."
        )
    else:
        desc_placeholder = (
            "e.g. They are looking to launch a tokenized deposit product and need "
            "help building a compliance program and finding a bank partner to "
            "support it."
        )

    description = st.text_area(
        "Describe the engagement *",
        height=120,
        placeholder=desc_placeholder,
        help=(
            "A brief description of what the client needs. AI will generate the engagement background."
            if is_licensing else
            "A brief description of what the client needs. AI will expand this into full proposal content."
        ),
    )

    # Scope and budget fields are NOT shown for licensing (all static in template)
    scope_areas = ""
    budget_info = ""
    if not is_licensing:
        if is_charter:
            scope_placeholder = (
                "e.g. Exploration and planning, Pre-filing charter application support, "
                "Post-filing charter application support"
            )
            budget_placeholder = (
                "e.g. $60,000 for Phase 1, $60,000/month for Phases 2-3"
            )
        else:
            scope_placeholder = (
                "e.g. Partnership engagement, Compliance gap analysis and program development"
            )
            budget_placeholder = "e.g. $25,000/month for 6 months"

        scope_areas = st.text_input(
            "Scope areas (comma-separated) *",
            placeholder=scope_placeholder,
            help="The main work streams / phases for the engagement.",
        )

        budget_info = st.text_input(
            "Budget & duration *",
            placeholder=budget_placeholder,
            help="The proposed fee structure and engagement length.",
        )

    if is_licensing:
        context_placeholder = (
            "e.g. They already have a BitLicense in NY. They need MTLs in 48 states. "
            "They plan to launch a stablecoin product."
        )
    elif is_charter:
        context_placeholder = (
            "e.g. They are considering both a national bank charter and a state "
            "charter. They already have outside counsel at Sullivan & Cromwell."
        )
    else:
        context_placeholder = (
            "e.g. The client already has a BSA/AML program in place. "
            "They need state-by-state licensing analysis."
        )

    additional_context = st.text_area(
        "Anything else important?",
        height=100,
        placeholder=context_placeholder,
        help="Optional. Any additional context, constraints, or details the proposal should reflect.",
    )

    # ── 3. Licensing Fees (licensing-specific) ──────────────────────────
    acquisition_fee = ""
    maintenance_fee = ""
    if is_licensing:
        st.header("Engagement Fees")
        st.caption(
            "Enter the fees for the core engagement. Leave blank to keep the "
            "template defaults (fee ranges)."
        )
        col_fee1, col_fee2 = st.columns(2)
        with col_fee1:
            acquisition_fee = st.text_input(
                "License Acquisition Monthly Fee",
                placeholder="e.g. $30,000/month",
                help="Monthly fee for the license acquisition phase. Applied to all timeline phases.",
            )
        with col_fee2:
            maintenance_fee = st.text_input(
                "License Maintenance Annual Fee",
                placeholder="e.g. $180,000/year",
                help="Annual contract fee for ongoing license maintenance.",
            )

    # ── 4. Options (compliance-specific) ──────────────────────────────
    include_compliance = False
    include_managed = False
    compliance_col = "Deposit Compliance"
    customer_type = None

    if not is_charter and not is_licensing:
        customer_type = st.selectbox(
            "Customer Type",
            ["Consumer", "Commercial", "Both"],
            index=0,
            help=(
                "The type of customers the client's product serves. This determines "
                "which regulations and compliance areas are emphasized in the proposal."
            ),
        )

        include_compliance = st.checkbox(
            "Include Compliance Programs chart",
            value=False,
            help=(
                "Include the Financial Crimes / Complaint Management / Enterprise Compliance "
                "matrix. Only check this if the engagement involves a compliance program "
                "build, gap analysis, or similar."
            ),
        )

        with st.expander("Advanced Options", expanded=False):
            include_managed = st.checkbox(
                "Include Managed Services as a primary service",
                value=False,
                help="When unchecked, Managed Services appears under Optional Services.",
            )
            if include_compliance:
                compliance_col = st.text_input(
                    "Compliance column override",
                    value="Deposit Compliance",
                    help=(
                        "The third column in the Compliance Programs chart. "
                        "AI will also suggest the best fit based on the engagement description."
                    ),
                )

    # ── Generate Button ───────────────────────────────────────────────
    st.markdown("---")

    if st.button("Generate Proposal", type="primary", use_container_width=True):
        # Validation
        if not client_name.strip():
            st.error("Client Name is required.")
            return
        if not description.strip():
            st.error("Engagement description is required.")
            return
        if not is_licensing and not scope_areas.strip():
            st.error("Scope areas are required.")
            return
        if not is_licensing and not budget_info.strip():
            st.error("Budget & duration is required.")
            return

        # Generate content with AI
        with st.spinner("Generating proposal content with AI... this may take 15-30 seconds."):
            try:
                ai_content = generate_with_ai(
                    client_name=client_name.strip(),
                    description=description.strip(),
                    scope_areas=scope_areas.strip(),
                    budget_info=budget_info.strip(),
                    proposal_type_key=proposal_type_key,
                    compliance_col=compliance_col.strip() if not is_charter else None,
                    additional_context=additional_context.strip() if additional_context.strip() else None,
                    customer_type=customer_type,
                )
            except json.JSONDecodeError as e:
                st.error(f"AI returned invalid JSON. Please try again. Error: {e}")
                return
            except Exception as e:
                st.error(f"AI generation failed: {e}")
                import traceback
                st.code(traceback.format_exc())
                return

        # Save AI content and form inputs for revision
        st.session_state.ai_content = ai_content
        st.session_state.form_inputs = {
            "client_name": client_name.strip(),
            "proposal_type": proposal_title.strip(),
            "proposal_type_key": proposal_type_key,
            "date_str": proposal_date.strftime("%B %Y"),
            "client_website": client_website.strip() or None,
            "include_managed": include_managed,
            "include_compliance": include_compliance,
            "compliance_col": compliance_col.strip(),
            "customer_type": customer_type,
            "acquisition_fee": acquisition_fee.strip() if acquisition_fee else None,
            "maintenance_fee": maintenance_fee.strip() if maintenance_fee else None,
        }

        # Build the document
        _build_and_store_doc(ai_content, st.session_state.form_inputs)

    # ── Post-generation: download + revise ────────────────────────────
    if st.session_state.generated_file is not None:
        st.success("Proposal generated successfully!")
        st.download_button(
            label="Download Proposal (.docx)",
            data=st.session_state.generated_file,
            file_name=st.session_state.generated_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

        # Revision feedback
        st.markdown("---")
        st.subheader("Revise Proposal")
        feedback = st.text_area(
            "What needs to change?",
            height=120,
            placeholder=(
                "e.g. The engagement background is too formal — make it more conversational. "
                "Add a bullet about state licensing requirements. "
                "The fee for Phase 2 should be $50,000/month instead."
            ),
            key="revision_feedback",
        )

        col_revise, col_clear = st.columns(2)
        with col_revise:
            if st.button("Revise & Regenerate", type="primary", use_container_width=True):
                if not feedback.strip():
                    st.error("Please describe what needs to change.")
                else:
                    with st.spinner("Revising proposal with AI..."):
                        try:
                            ptk = st.session_state.form_inputs.get("proposal_type_key", "compliance")
                            revised = revise_with_ai(
                                previous_content=st.session_state.ai_content,
                                feedback=feedback.strip(),
                                proposal_type_key=ptk,
                            )
                            st.session_state.ai_content = revised
                            _build_and_store_doc(revised, st.session_state.form_inputs)
                        except json.JSONDecodeError as e:
                            st.error(f"AI returned invalid JSON during revision. Try again. Error: {e}")
                        except Exception as e:
                            st.error(f"Revision failed: {e}")
                            import traceback
                            st.code(traceback.format_exc())

        with col_clear:
            if st.button("Clear & Start New Proposal", use_container_width=True):
                st.session_state.generated_file = None
                st.session_state.generated_filename = None
                st.session_state.ai_content = None
                st.session_state.form_inputs = None
                st.rerun()


# =====================================================================
# MAIN
# =====================================================================
if not st.session_state.authenticated:
    login_screen()
else:
    main_form()
