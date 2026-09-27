import os
import io
import re
import json
import secrets
import sqlite3
import hashlib
import warnings
import urllib.parse
import pandas as pd
from PIL import Image
from pypdf import PdfReader
from docx import Document
import streamlit as st
from groq import Groq
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import Chroma

warnings.filterwarnings("ignore")

# Kavach AI Configuration with Collapsible Sidebar
st.set_page_config(
    page_title="Kavach AI — Your Shield Against Tax & Compliance Risks", 
    page_icon="🛡️", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Corporate Blue & White Styling
st.markdown("""
<style>
    .stApp {
        background-color: #f8fafc;
        color: #0f172a;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }
    .stChatMessage[data-testid="stChatMessage"]:nth-child(even) {
        background-color: #eff6ff !important;
        border-radius: 12px !important;
        border: 1px solid #bfdbfe !important;
        padding: 14px !important;
    }
    .stChatMessage[data-testid="stChatMessage"]:nth-child(odd) {
        background-color: #ffffff !important;
        border-radius: 12px !important;
        border: 1px solid #e2e8f0 !important;
        padding: 14px !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02) !important;
    }
    .stButton > button {
        border-radius: 8px !important;
        background-color: #2563eb !important;
        color: #ffffff !important;
        border: none !important;
        font-weight: 500 !important;
        transition: all 0.2s ease-in-out !important;
    }
    .stButton > button:hover {
        background-color: #1d4ed8 !important;
        color: #ffffff !important;
        box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2) !important;
    }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: visible;}
</style>
""", unsafe_allow_html=True)

# Helper function to strip any raw HTML tags like <br>
def clean_html_tags(text):
    if not text:
        return ""
    return text.replace("<br>", "\n").replace("<br/>", "\n").replace("<br />", "\n").replace("</br>", "\n")

# --- 1. LOCAL DATABASE SETUP ---
DB_PATH = "tax_system.db"

def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=30.0, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn

with get_db() as init_conn:
    init_conn.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email TEXT UNIQUE,
        password_hash TEXT
    )
    """)
    init_conn.execute("""
    CREATE TABLE IF NOT EXISTS sessions (
        token TEXT PRIMARY KEY,
        user_id INTEGER,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    init_conn.execute("""
    CREATE TABLE IF NOT EXISTS conversations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        title TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    init_conn.execute("""
    CREATE TABLE IF NOT EXISTS messages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        conversation_id INTEGER,
        role TEXT,
        content TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    init_conn.commit()

def hash_val(val):
    return hashlib.sha256(val.encode()).hexdigest()

def generate_docx(title, content):
    doc = Document()
    doc.add_heading(title, level=0)
    cleaned_content = clean_html_tags(content)
    for line in cleaned_content.split("\n"):
        clean_line = line.strip()
        if clean_line.startswith("# "):
            doc.add_heading(clean_line[2:], level=1)
        elif clean_line.startswith("## "):
            doc.add_heading(clean_line[3:], level=2)
        elif clean_line.startswith("### "):
            doc.add_heading(clean_line[4:], level=3)
        elif clean_line:
            doc.add_paragraph(clean_line)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def generate_xlsx(text):
    cleaned_text = clean_html_tags(text)
    lines = [line.strip() for line in cleaned_text.split("\n") if "|" in line and not line.strip().startswith("|---")]
    if len(lines) >= 2:
        headers = [c.strip() for c in lines[0].split("|")[1:-1]]
        rows = []
        for line in lines[1:]:
            row = [c.strip() for c in line.split("|")[1:-1]]
            if len(row) == len(headers):
                rows.append(row)
        df = pd.DataFrame(rows, columns=headers)
    else:
        df = pd.DataFrame({"Kavach AI Advisory Summary": [cleaned_text[:5000]]})
    
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Advisory_Summary")
    bio.seek(0)
    return bio

# --- 2. VECTOR DATABASE ---
@st.cache_resource
def get_vector_db():
    if os.path.exists("tax_db"):
        embeddings = HuggingFaceEmbeddings(model_name="all-MiniLM-L6-v2")
        return Chroma(persist_directory="tax_db", embedding_function=embeddings)
    return None

vector_db = get_vector_db()

# --- 3. SILENT BACKGROUND API KEY (100% White-Label) ---
groq_key = st.secrets.get("GROQ_API_KEY", "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")

# --- 4. PERSISTENT AUTO-LOGIN ---
if "user_id" not in st.session_state:
    st.session_state.user_id = None
if "current_convo_id" not in st.session_state:
    st.session_state.current_convo_id = None

auth_token = st.query_params.get("session_auth")
if not st.session_state.user_id and auth_token:
    with get_db() as c_conn:
        cur = c_conn.cursor()
        cur.execute("SELECT user_id FROM sessions WHERE token=?", (auth_token,))
        row = cur.fetchone()
        if row:
            st.session_state.user_id = row[0]
            cur.execute("SELECT email FROM users WHERE id=?", (row[0],))
            u = cur.fetchone()
            st.session_state.user_email = u[0]

# --- 5. AUTHENTICATION (Login / Sign Up) ---
if not st.session_state.user_id:
    st.title("🛡️ Kavach AI")
    st.markdown("#### *Your Shield Against Tax & Compliance Risks*")
    st.caption("Enterprise Advisory System for Finance, Tax, Payroll, HR Compliance & Market Analysis")
    
    auth_tab1, auth_tab2 = st.tabs(["🔐 Sign In", "📝 Create Account"])
    
    with auth_tab1:
        with st.form("login_form"):
            login_email = st.text_input("Email Address")
            login_pw = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Sign In to Kavach AI")
            if submitted:
                with get_db() as c_conn:
                    cur = c_conn.cursor()
                    cur.execute("SELECT id, email FROM users WHERE email=? AND password_hash=?", (login_email, hash_val(login_pw)))
                    row = cur.fetchone()
                    if row:
                        new_token = secrets.token_hex(24)
                        cur.execute("INSERT OR REPLACE INTO sessions (token, user_id) VALUES (?, ?)", (new_token, row[0]))
                        c_conn.commit()
                        st.query_params["session_auth"] = new_token
                        st.session_state.user_id = row[0]
                        st.session_state.user_email = row
                        st.rerun()
                    else:
                        st.error("Invalid email or password.")

    with auth_tab2:
        with st.form("signup_form"):
            new_email = st.text_input("Email Address")
            new_pw = st.text_input("Create Password", type="password")
            create_btn = st.form_submit_button("Register Account")
            if create_btn:
                try:
                    with get_db() as c_conn:
                        c_conn.execute("INSERT INTO users (email, password_hash) VALUES (?, ?)", (new_email, hash_val(new_pw)))
                        c_conn.commit()
                    st.success("Account created successfully! Please sign in.")
                except sqlite3.IntegrityError:
                    st.error("This email is already registered.")
    st.stop()

# --- 6. SIDEBAR: CHAT HISTORY & BOTTOM ACCOUNTS ---
with st.sidebar:
    st.markdown("### 🛡️ Kavach AI")
    st.caption("Your Compliance Shield")
    
    if st.button("➕ New Consultation", use_container_width=True):
        st.session_state.current_convo_id = None
        st.rerun()

    st.markdown("### 💬 Consultations")
    with get_db() as c_conn:
        cur = c_conn.cursor()
        cur.execute("SELECT id, title FROM conversations WHERE user_id=? ORDER BY id DESC", (st.session_state.user_id,))
        convos = cur.fetchall()
    
    for c_id, c_title in convos:
        if st.button(f"📄 {c_title[:24]}...", key=f"convo_{c_id}", use_container_width=True):
            st.session_state.current_convo_id = c_id
            st.rerun()

    st.markdown("---")
    st.markdown("#### 👤 Account")
    st.caption(f"Logged in: **{st.session_state.user_email}**")
    
    with st.expander("⚙️ Account Settings"):
        new_pass = st.text_input("New Password", type="password")
        if st.button("Update Password"):
            if len(new_pass) >= 4:
                with get_db() as c_conn:
                    c_conn.execute("UPDATE users SET password_hash=? WHERE id=?", (hash_val(new_pass), st.session_state.user_id))
                    c_conn.commit()
                st.success("Password updated successfully!")
            else:
                st.warning("Password must be at least 4 characters.")
                
    if st.button("🚪 Sign Out", use_container_width=True):
        if "session_auth" in st.query_params:
            with get_db() as c_conn:
                c_conn.execute("DELETE FROM sessions WHERE token=?", (st.query_params["session_auth"],))
                c_conn.commit()
        st.query_params.clear()
        st.session_state.clear()
        st.rerun()

# --- 7. MAIN CHAT AREA ---
st.markdown("<h2 style='color:#1e3a8a; margin-bottom:0;'>🛡️ Kavach AI</h2>", unsafe_allow_html=True)
st.markdown("<p style='color:#475569; font-size:16px; margin-top:0; font-weight:500;'>Your Shield Against Tax & Compliance Risks</p>", unsafe_allow_html=True)

def check_is_image_intent(query):
    q = query.lower()
    img_words = ["image", "poster", "banner", "photo", "graphic", "illustration", "design", "render"]
    action_words = ["create", "generate", "make", "design", "draw", "banao", "chahiye"]
    return any(w in q for w in img_words) and any(a in q for a in action_words)

current_messages = []
full_chat_text = ""
if st.session_state.current_convo_id:
    with get_db() as c_conn:
        cur = c_conn.cursor()
        cur.execute("SELECT role, content FROM messages WHERE conversation_id=? ORDER BY id ASC", (st.session_state.current_convo_id,))
        current_messages = [{"role": r, "content": c} for r, c in cur.fetchall()]

for msg in current_messages:
    with st.chat_message(msg["role"]):
        if msg["content"].startswith("[IMAGE_URL]:"):
            img_url = msg["content"].replace("[IMAGE_URL]:", "").strip()
            st.image(img_url, caption="Generated Visual Graphic", use_container_width=True)
        else:
            display_text = clean_html_tags(msg["content"])
            st.markdown(display_text)
            full_chat_text += f"\n\n[{msg['role'].upper()}]:\n" + display_text

# Top Export Bar
if current_messages:
    exp_col1, exp_col2 = st.columns(2)
    with exp_col1:
        docx_file = generate_docx("Kavach AI - Advisory Report", full_chat_text)
        st.download_button("📥 Export Entire Consultation to Word (.docx)", data=docx_file, file_name="Kavach_AI_Advisory_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with exp_col2:
        xlsx_file = generate_xlsx(full_chat_text)
        st.download_button("📊 Export Consultation Data to Excel (.xlsx)", data=xlsx_file, file_name="Kavach_AI_Summary_Data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# --- 8. SLEEK ATTACHMENT SECTION ---
with st.expander("📎 Attach Client Document (PDF Notice, Excel Sheet, GST/ITR JSON, or Chart Image)", expanded=False):
    uploaded_file = st.file_uploader(
        "Upload client file or chart screenshot for analysis", 
        type=["pdf", "png", "jpg", "jpeg", "xlsx", "csv", "json"],
        label_visibility="collapsed"
    )

file_context_str = ""
if uploaded_file:
    file_name = uploaded_file.name
    st.info(f"Attached: **{file_name}** (Ready for consultation)")
    try:
        if file_name.endswith(".pdf"):
            reader = PdfReader(uploaded_file)
            pdf_text = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
            file_context_str = f"\n[User Attached PDF Content ({file_name})]:\n{pdf_text[:15000]}"
        elif file_name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(uploaded_file)
            file_context_str = f"\n[User Attached Excel Data ({file_name})]:\nSummary:\n{df.head(50).to_markdown()}"
        elif file_name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
            file_context_str = f"\n[User Attached CSV Data ({file_name})]:\nSummary:\n{df.head(50).to_markdown()}"
        elif file_name.endswith(".json"):
            json_data = json.load(uploaded_file)
            file_context_str = f"\n[User Attached JSON Data ({file_name})]:\n{json.dumps(json_data, indent=2)[:20000]}"
        elif file_name.lower().endswith((".png", ".jpg", ".jpeg")):
            file_context_str = f"\n[User has attached a visual chart/invoice image named {file_name} for pattern & structural analysis]"
    except Exception as e:
        st.error(f"Error parsing uploaded file: {e}")

# --- 9. CHAT INPUT & EXECUTION ---
user_query = st.chat_input("Ask any question regarding Tax, Finance, HR Compliance, or Stock Market...")

if user_query:
    if not st.session_state.current_convo_id:
        title = user_query[:35]
        with get_db() as c_conn:
            cur = c_conn.cursor()
            cur.execute("INSERT INTO conversations (user_id, title) VALUES (?, ?)", (st.session_state.user_id, title))
            c_conn.commit()
            st.session_state.current_convo_id = cur.lastrowid

    with get_db() as c_conn:
        c_conn.execute("INSERT INTO messages (conversation_id, role, content) VALUES (?, 'user', ?)", (st.session_state.current_convo_id, user_query))
        c_conn.commit()

    with st.chat_message("user"):
        st.markdown(user_query)

    # CASE A: IMAGE GENERATION
    if check_is_image_intent(user_query):
        with st.chat_message("assistant"):
            with st.spinner("Designing professional visual graphic..."):
                encoded_prompt = urllib.parse.quote(f"{user_query}, professional corporate HR, tax and financial poster, high definition, graphic design, modern typography")
                image_url = f"https://image.pollinations.ai/prompt/{encoded_prompt}?width=1024&height=1024&nologo=true&seed={secrets.randbelow(10000)}"
                st.image(image_url, caption="Generated Visual Graphic", use_container_width=True)
                save_msg = f"[IMAGE_URL]: {image_url}"
                with get_db() as c_conn:
                    c_conn.execute("INSERT INTO messages (conversation_id, role, content) VALUES (?, 'assistant', ?)", (st.session_state.current_convo_id, save_msg))
                    c_conn.commit()
        st.stop()

    # CASE B: COMPREHENSIVE EXPERT ENGINE
    search_query = user_query
    if current_messages:
        last_user_turns = [m["content"] for m in current_messages if m["role"] == "user"]
        if last_user_turns:
            search_query = f"{last_user_turns[-1]} {user_query}"

    statutory_context = ""
    if vector_db:
        sec_match = re.search(r'\b(?:section|sec|rule|धारा|नियम)?\s*([0-9]{1,4}[a-z]{0,3})\b', search_query, re.IGNORECASE)
        target_token = sec_match.group(1).upper() if sec_match else None
        
        retrieved_docs = []
        if target_token and len(target_token) >= 2:
            try:
                retrieved_docs = vector_db.similarity_search(search_query, k=8, where_document={"$contains": target_token})
            except Exception:
                retrieved_docs = []
        
        if not retrieved_docs:
            retrieved_docs = vector_db.similarity_search(search_query, k=8)
            
        statutory_context = "\n\n---\n\n".join([d.page_content for d in retrieved_docs])

    # --- ADVANCED SYSTEM INSTRUCTION ---
    system_instruction = f"""
    You are Kavach AI (powered by NextGen FinHR Architecture), a highly advanced PhD-level expert in:
    1. Indian Finance (Income Tax, GST, Corporate Accounts, Audits)
    2. HR Compliance (Payroll, Labour Laws, PF/ESIC, Gratuity, Bonus)
    3. Indian Stock Market & Technical Analysis (NSE/BSE)

    Your tone is professional, authoritative, and helpful. You speak naturally in the exact language used by the user (Hinglish, English, or Hindi).
    Your absolute priority is 100% ACCURACY delivered at TURBO SPEED with zero filler text.

    STRICT FORMATTING RULE:
    - NEVER use raw HTML tags such as <br>, <p>, or <div>.
    - Always use standard Markdown newlines (press Enter/return) for bullet points and spacing.

    STRICT EXECUTION RULES:

    1. WELCOME LOGIC (On Hi / Hello / Greeting):
    - Greet the user warmly and professionally.
    - Immediately present a clear menu of capabilities:
      * (A) Finance & Taxation (Income Tax, GST, Corporate Accounts, Tax Planning)
      * (B) HR & Payroll Compliance (EPF, ESIC, Gratuity, Labour Codes, Salary Structuring)
      * (C) Indian Stock Market & Technical Analysis (NSE/BSE Stocks, Chart Patterns, Support & Resistance)

    2. COMPREHENSIVE 1-SHOT ANSWER RULE (No Latency, No Half Information):
    - Skip generic filler phrases (e.g., "Aaiye samajhte hain", "Let's explore"). Start directly with the answer.
    - Provide complete, end-to-end information in one high-density response: Core Definition/Rule, Eligibility Criteria, Maximum Limits/Slabs, Crucial Deadlines, Exceptions/Conditions, and Required Documents.
    - Use bold keywords, concise bullet points, and neat Markdown Tables.

    3. MANDATORY PRACTICAL EXAMPLES:
    - For every technical, legal, or calculation concept, provide a concise real-world corporate case:
      * **Scenario:** 1 brief sentence (e.g., "Amit's gross salary is ₹75,000/month" or "Company ABC makes contract payment of ₹2,50,000").
      * **Impact / Math:** Direct calculation breakdown in a neat Markdown Table.

    4. FIN-HR INTEGRATION & COMPLIANCE:
    - For overlapping queries (CTC design, Gratuity, Bonus, Resignation payouts), create distinct, crisp sections:
      * **[Finance & Tax Implications]**
      * **[HR & Labour Law Compliances]**
    - Always apply current financial year guidelines, Finance Act 2026, and updated labour codes.
    - **RISK ALERT:** Proactively scan user queries or data for compliance risks, loopholes, or penalty exposure and flag them immediately.

    5. DATA SOURCES & PRIORITIZATION:
    - Priority: Ground your answers on the provided Internal Statutory Context.
    - Live Updates & Stock Data: Use the built-in browser search for live market prices, corporate actions, and latest government circulars.
    - Never cite internal PDF filenames or page numbers. Formulate with statutory authority (e.g., "Under Section 194C of the Income Tax Act...", "As per the Employees' Provident Funds Act, 1952...").

    6. INDIAN STOCK MARKET & TECHNICAL ANALYSIS EXPERT (NSE/BSE):
    - Real-Time News & Tickers: When asked about live stocks, analyze Current Market Price (CMP), Daily High/Low, 52-Week High/Low, and Volume Spikes.
    - Technical Analysis Engine: Identify Primary/Secondary Trends, Support & Resistance levels, 20/50/200 EMAs, RSI (Overbought >70, Oversold <30), and MACD.
    - Visual Chart Decoding: When a user attaches a chart screenshot, identify Japanese Candlestick Patterns (Hammer, Engulfing, Doji) and Price Action Structures.
    - Probabilistic Setup (Not Tips): Formulate data-backed probabilistic setups with entry zones and risk-to-reward ratios.
    - MANDATORY SEBI DISCLAIMER: Every stock market analysis response MUST end with this exact blockquote:
      > ⚠️ **Disclaimer:** *This analysis is for educational and informational purposes only based on algorithmic indicators and live market data. It does not constitute verified financial advice or a direct buy/sell recommendation under SEBI guidelines. Stock investing involves market risks. Please perform your own due diligence or consult a SEBI-registered financial advisor before making any investment decisions.*

    7. STATUTORY TAX / HR DISCLAIMER:
    - Conclude tax/HR consultations with:
      *(Note: This guidance is based on statutory provisions and regulatory updates; verify case-specific facts prior to final filings.)*

    Internal Statutory Context:
    {statutory_context}

    Client Attached Data:
    {file_context_str}
    """

    with st.chat_message("assistant"):
        message_placeholder = st.empty()
        full_response = ""
        
        llm_messages = [{"role": "system", "content": system_instruction}]
        for prev_msg in current_messages[-8:]:
            if not prev_msg["content"].startswith("[IMAGE_URL]:"):
                llm_messages.append({"role": prev_msg["role"], "content": prev_msg["content"]})
        llm_messages.append({"role": "user", "content": user_query})

        client = Groq(api_key=groq_key)
        
        search_models = ["openai/gpt-oss-120b", "openai/gpt-oss-20b", "llama-3.1-8b-instant"]
        success = False
        last_error = None
        
        for m_name in search_models:
            try:
                try:
                    completion = client.chat.completions.create(
                        model=m_name,
                        messages=llm_messages,
                        tools=[{"type": "browser_search"}],
                        temperature=0.2
                    )
                    full_response = completion.choices[0].message.content
                except Exception:
                    completion = client.chat.completions.create(
                        model=m_name,
                        messages=llm_messages,
                        stream=True,
                        temperature=0.2
                    )
                    for chunk in completion:
                        content = chunk.choices[0].delta.content
                        if content:
                            full_response += content
                            # Real-time cleaning during typing
                            message_placeholder.markdown(clean_html_tags(full_response) + "▌")

                if full_response.strip():
                    full_response = clean_html_tags(full_response)
                    message_placeholder.markdown(full_response)
                    success = True
                    break
            except Exception as e:
                last_error = e
                continue

        if not success:
            st.error(f"Error executing query: {last_error}")
            st.stop()

        # Word / Excel Download Buttons
        q_lower = user_query.lower()
        if any(w in q_lower for w in ["word file", "word me", "docx", "doc file", "report", "document"]):
            docx_data = generate_docx("Kavach AI - Advisory Report", full_response)
            st.download_button(
                "📥 Download Advisory as Word File (.docx)", 
                data=docx_data, 
                file_name="Kavach_AI_Advisory_Report.docx", 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        if any(w in q_lower for w in ["excel file", "excel me", "xlsx", "sheet", "table"]):
            xlsx_data = generate_xlsx(full_response)
            st.download_button(
                "📊 Download Table as Excel Sheet (.xlsx)", 
                data=xlsx_data, 
                file_name="Kavach_AI_Summary_Data.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with get_db() as c_conn:
        c_conn.execute("INSERT INTO messages (conversation_id, role, content) VALUES (?, 'assistant', ?)", (st.session_state.current_convo_id, full_response))
        c_conn.commit()