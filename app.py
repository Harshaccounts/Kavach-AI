import os
import io
import re
import json
import base64
import secrets
import sqlite3
import hashlib
import warnings
import datetime
import urllib.parse
import contextlib
import pandas as pd
from PIL import Image
from pypdf import PdfReader
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import streamlit as st
from groq import Groq
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import Chroma

# Live Internet Search Engine
try:
    from duckduckgo_search import DDGS
    HAS_DDG = True
except ImportError:
    HAS_DDG = False

warnings.filterwarnings("ignore")

# --- KAVACH AI CONFIGURATION ---
st.set_page_config(
    page_title="Kavach AI — NextGen FinHR", 
    page_icon="🛡️", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- EXACT CHATGPT MINIMALIST DESIGN SYSTEM ---
st.markdown("""
<style>
    /* 1. Canvas & Clean Typography (ChatGPT Style) */
    html, body, [data-testid="stAppViewContainer"], .stApp {
        background-color: #FFFFFF !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif !important;
        color: #0D0D0D !important;
    }
    
    .main .block-container {
        max-width: 768px !important;
        padding-top: 1.5rem !important;
        padding-bottom: 7.5rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }

    /* 2. Hide Streamlit Avatars (Pure Clean Look) */
    [data-testid="stChatMessage"] [data-testid="chatAvatarIcon-user"],
    [data-testid="stChatMessage"] [data-testid="chatAvatarIcon-assistant"],
    .stChatMessageAvatar {
        display: none !important;
    }
    
    /* 3. User Message: Right-Aligned Soft Grey Bubble */
    [data-testid="stChatMessage"]:has([data-testid="chatAvatarIcon-user"]) {
        background-color: #F4F4F4 !important;
        border-radius: 20px !important;
        padding: 10px 18px !important;
        margin-bottom: 1.25rem !important;
        margin-left: auto !important;
        margin-right: 0 !important;
        width: fit-content !important;
        max-width: 82% !important;
        border: none !important;
        box-shadow: none !important;
    }
    [data-testid="stChatMessage"]:has([data-testid="chatAvatarIcon-user"]) p {
        color: #0D0D0D !important;
        font-size: 15px !important;
        margin: 0 !important;
        line-height: 1.5 !important;
    }

    /* 4. Assistant Message: Clean, Borderless, Left-Aligned Full Width */
    [data-testid="stChatMessage"]:has([data-testid="chatAvatarIcon-assistant"]) {
        background-color: transparent !important;
        border: none !important;
        padding: 0 !important;
        margin-bottom: 2rem !important;
        max-width: 100% !important;
        box-shadow: none !important;
    }
    
    .stChatMessage [data-testid="stMarkdownContainer"] {
        font-size: 15.5px !important;
        line-height: 1.7 !important;
        color: #0D0D0D !important;
    }
    .stChatMessage [data-testid="stMarkdownContainer"] p {
        margin-bottom: 0.9rem !important;
        color: #0D0D0D !important;
    }
    .stChatMessage [data-testid="stMarkdownContainer"] strong {
        color: #000000 !important;
        font-weight: 600 !important;
    }
    .stChatMessage [data-testid="stMarkdownContainer"] h1,
    .stChatMessage [data-testid="stMarkdownContainer"] h2,
    .stChatMessage [data-testid="stMarkdownContainer"] h3 {
        color: #000000 !important;
        font-weight: 600 !important;
        margin-top: 1.2rem !important;
        margin-bottom: 0.5rem !important;
    }
    .stChatMessage [data-testid="stMarkdownContainer"] ul, 
    .stChatMessage [data-testid="stMarkdownContainer"] ol {
        padding-left: 1.4rem !important;
        margin-bottom: 0.9rem !important;
        color: #0D0D0D !important;
    }
    .stChatMessage [data-testid="stMarkdownContainer"] li {
        margin-bottom: 0.4rem !important;
    }

    /* 5. ChatGPT-Style Clean Bordered Tables */
    .stChatMessage [data-testid="stMarkdownContainer"] table {
        width: 100% !important;
        border-collapse: separate !important;
        border-spacing: 0 !important;
        border: 1px solid #E5E5E5 !important;
        border-radius: 8px !important;
        overflow: hidden !important;
        margin: 1.2rem 0 !important;
        display: table !important;
    }
    .stChatMessage [data-testid="stMarkdownContainer"] th {
        background-color: #F9F9F9 !important;
        color: #0D0D0D !important;
        font-weight: 600 !important;
        padding: 10px 14px !important;
        border-bottom: 1px solid #E5E5E5 !important;
        border-right: 1px solid #E5E5E5 !important;
        text-align: left !important;
        font-size: 14px !important;
    }
    .stChatMessage [data-testid="stMarkdownContainer"] td {
        padding: 10px 14px !important;
        border-bottom: 1px solid #F0F0F0 !important;
        border-right: 1px solid #E5E5E5 !important;
        color: #1A1A1A !important;
        font-size: 14px !important;
    }
    .stChatMessage [data-testid="stMarkdownContainer"] tr:last-child td {
        border-bottom: none !important;
    }
    .stChatMessage [data-testid="stMarkdownContainer"] th:last-child,
    .stChatMessage [data-testid="stMarkdownContainer"] td:last-child {
        border-right: none !important;
    }

    /* 6. ChatGPT-Style Bottom Input Pill */
    [data-testid="stChatInput"],
    [data-testid="stChatInput"] > div,
    .stChatInput,
    .stChatInputContainer {
        border-radius: 28px !important;
        border: 1px solid #E5E5E5 !important;
        background-color: #F4F4F4 !important;
        background: #F4F4F4 !important;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04) !important;
        padding: 4px 10px !important;
    }
    
    [data-testid="stChatInput"] textarea,
    [data-testid="stChatInputTextArea"],
    .stChatInput textarea {
        color: #0D0D0D !important;
        -webkit-text-fill-color: #0D0D0D !important;
        background-color: transparent !important;
        background: transparent !important;
        font-size: 15.5px !important;
        font-weight: 400 !important;
    }
    
    [data-testid="stChatInput"] textarea::placeholder,
    .stChatInput textarea::placeholder {
        color: #8E8E8E !important;
        -webkit-text-fill-color: #8E8E8E !important;
    }

    /* 7. Quick Suggestion Cards */
    div[data-testid="stHorizontalBlock"] button {
        background-color: #FFFFFF !important;
        border: 1px solid #E5E5E5 !important;
        border-radius: 16px !important;
        padding: 14px 16px !important;
        height: auto !important;
        min-height: 76px !important;
        text-align: left !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02) !important;
        transition: all 0.15s ease !important;
        margin-bottom: 0.6rem !important;
    }
    div[data-testid="stHorizontalBlock"] button:hover {
        background-color: #F9F9F9 !important;
        border-color: #D1D5DB !important;
    }
    div[data-testid="stHorizontalBlock"] button p {
        color: #0D0D0D !important;
        font-size: 0.95rem !important;
        font-weight: 500 !important;
        margin: 0 !important;
    }

    /* Sidebar Controls */
    header[data-testid="stHeader"] {
        background: transparent !important;
    }
    [data-testid="stSidebarCollapsedControl"] {
        visibility: visible !important;
        display: block !important;
        color: #0D0D0D !important;
        background-color: #F4F4F4 !important;
        border-radius: 8px !important;
        border: 1px solid #E5E5E5 !important;
        margin: 8px !important;
    }
    #MainMenu, footer {
        visibility: hidden !important;
    }
</style>
""", unsafe_allow_html=True)

def clean_html_tags(text):
    if not text:
        return ""
    return text.replace("<br>", "\n").replace("<br/>", "\n").replace("<br />", "\n").replace("</br>", "\n")

# --- LIVE INTERNET SEARCH FUNCTION ---
def perform_live_web_search(query, max_results=4):
    if not HAS_DDG:
        return ""
    try:
        with DDGS() as ddgs:
            search_term = f"{query} Tax Collected at Source GST Income Tax Act India circular" if query.strip().upper() == "TCS" else f"{query} India tax compliance circular 2025 2026"
            results = list(ddgs.text(search_term, max_results=max_results))
            if results:
                formatted = []
                for r in results:
                    title = r.get("title", "")
                    snippet = r.get("body", "")
                    link = r.get("href", "")
                    formatted.append(f"• Source Title: {title}\n  Details: {snippet}\n  Link: {link}")
                return "\n\n".join(formatted)
    except Exception:
        return ""
    return ""

# --- 1. LOCAL DATABASE SETUP (AUTO-CLOSING) ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "tax_system.db")

@contextlib.contextmanager
def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=60.0, check_same_thread=False)
    conn.execute("PRAGMA busy_timeout = 30000;")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()

def init_db():
    with get_db() as conn:
        conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            email TEXT UNIQUE,
            password_hash TEXT
        )
        """)
        try:
            conn.execute("ALTER TABLE users ADD COLUMN name TEXT")
        except Exception:
            pass
            
        conn.execute("""
        CREATE TABLE IF NOT EXISTS sessions (
            token TEXT PRIMARY KEY,
            user_id INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """)
        conn.execute("""
        CREATE TABLE IF NOT EXISTS conversations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            title TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """)
        conn.execute("""
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            conversation_id INTEGER,
            role TEXT,
            content TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """)

init_db()

def hash_val(val):
    return hashlib.sha256(val.encode()).hexdigest()

# --- PROFESSIONAL WORD (.DOCX) EXPORT HELPER ---
def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def generate_docx(title, full_chat_text):
    doc = Document()
    
    # 1. Page Margins (0.8 inch clean executive layout)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    # 2. Executive Cover Banner
    p_badge = doc.add_paragraph()
    r_badge = p_badge.add_run("CONFIDENTIAL & STATUTORY ADVISORY REPORT")
    r_badge.font.name = "Segoe UI"
    r_badge.font.size = Pt(8.5)
    r_badge.font.bold = True
    r_badge.font.color.rgb = RGBColor(79, 70, 229)
    
    p_title = doc.add_paragraph()
    r_title = p_title.add_run("🛡️ Kavach AI — NextGen FinHR")
    r_title.font.name = "Segoe UI"
    r_title.font.size = Pt(20)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(15, 23, 42)

    p_sub = doc.add_paragraph()
    now_str = datetime.datetime.now().strftime("%d %B, %Y | %I:%M %p")
    r_sub = p_sub.add_run(f"Generated on: {now_str} • Domain: Indian Taxation & Compliance")
    r_sub.font.name = "Segoe UI"
    r_sub.font.size = Pt(9.5)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 3. Process Content Block by Block
    cleaned_content = clean_html_tags(full_chat_text)
    lines = cleaned_content.split("\n")
    
    in_table = False
    table_lines = []

    def flush_table(t_lines):
        if not t_lines or len(t_lines) < 2:
            return
        headers = [c.strip() for c in t_lines[0].split("|")[1:-1]]
        data_rows = []
        for l in t_lines[1:]:
            parts = [c.strip() for c in l.split("|")[1:-1]]
            if len(parts) == len(headers):
                data_rows.append(parts)
        
        if not headers or not data_rows:
            return

        table = doc.add_table(rows=len(data_rows) + 1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Header Row Styling (Dark Navy with White Bold Text)
        hdr_cells = table.rows[0].cells
        for i, header_text in enumerate(headers):
            hdr_cells[i].text = header_text
            set_cell_background(hdr_cells[i], "1E293B")
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for r in p.runs:
                r.font.name = "Segoe UI"
                r.font.size = Pt(9.5)
                r.font.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
        
        # Data Rows (Alternating Zebra Striping)
        for r_idx, row_data in enumerate(data_rows):
            row_cells = table.rows[r_idx + 1].cells
            bg_color = "F8FAFC" if r_idx % 2 == 0 else "FFFFFF"
            for c_idx, cell_value in enumerate(row_data):
                row_cells[c_idx].text = cell_value
                set_cell_background(row_cells[c_idx], bg_color)
                p = row_cells[c_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    r.font.name = "Segoe UI"
                    r.font.size = Pt(9)
                    r.font.color.rgb = RGBColor(30, 41, 59)
                    
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_after = Pt(10)

    for line in lines:
        raw_l = line.strip()
        
        if raw_l.startswith("|") and raw_l.endswith("|"):
            if "---" in raw_l:
                continue
            in_table = True
            table_lines.append(raw_l)
            continue
        else:
            if in_table:
                flush_table(table_lines)
                table_lines = []
                in_table = False

        if not raw_l:
            continue

        if raw_l.startswith("[USER]:"):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run("👤 CLIENT INQUIRY:")
            r.font.name = "Segoe UI"
            r.font.size = Pt(11)
            r.font.bold = True
            r.font.color.rgb = RGBColor(14, 116, 144)
            
            p_content = doc.add_paragraph()
            r_c = p_content.add_run(raw_l.replace("[USER]:", "").strip())
            r_c.font.name = "Segoe UI"
            r_c.font.size = Pt(10)
            r_c.font.color.rgb = RGBColor(51, 65, 85)
            continue
            
        elif raw_l.startswith("[ASSISTANT]:"):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run("🛡️ KAVACH STATUTORY OPINION & ANALYSIS:")
            r.font.name = "Segoe UI"
            r.font.size = Pt(11)
            r.font.bold = True
            r.font.color.rgb = RGBColor(30, 58, 138)
            continue

        if raw_l.startswith("# "):
            h = doc.add_paragraph()
            h.paragraph_format.space_before = Pt(14)
            h.paragraph_format.space_after = Pt(4)
            r = h.add_run(raw_l[2:])
            r.font.name = "Segoe UI"
            r.font.size = Pt(14)
            r.font.bold = True
            r.font.color.rgb = RGBColor(15, 23, 42)
        elif raw_l.startswith("## "):
            h = doc.add_paragraph()
            h.paragraph_format.space_before = Pt(12)
            h.paragraph_format.space_after = Pt(3)
            r = h.add_run(raw_l[3:])
            r.font.name = "Segoe UI"
            r.font.size = Pt(12.5)
            r.font.bold = True
            r.font.color.rgb = RGBColor(30, 58, 138)
        elif raw_l.startswith("### "):
            h = doc.add_paragraph()
            h.paragraph_format.space_before = Pt(10)
            h.paragraph_format.space_after = Pt(2)
            r = h.add_run(raw_l[4:])
            r.font.name = "Segoe UI"
            r.font.size = Pt(11)
            r.font.bold = True
            r.font.color.rgb = RGBColor(51, 65, 85)
        elif raw_l.startswith("- ") or raw_l.startswith("• "):
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_after = Pt(2)
            text_to_format = raw_l[2:].strip()
            parts = re.split(r'(\*\*.*?\*\*)', text_to_format)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    r = p.add_run(part[2:-2])
                    r.bold = True
                else:
                    r = p.add_run(part)
                r.font.name = "Segoe UI"
                r.font.size = Pt(9.5)
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.line_spacing = 1.15
            parts = re.split(r'(\*\*.*?\*\*)', raw_l)
            for part in parts:
                if part.startswith("**") and part.endswith("**"):
                    r = p.add_run(part[2:-2])
                    r.bold = True
                else:
                    r = p.add_run(part)
                r.font.name = "Segoe UI"
                r.font.size = Pt(9.5)
                r.font.color.rgb = RGBColor(30, 41, 59)

    if in_table:
        flush_table(table_lines)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- PROFESSIONAL EXCEL (.XLSX) EXPORT HELPER ---
def generate_xlsx(text):
    cleaned_text = clean_html_tags(text)
    lines = [l.strip() for l in cleaned_text.split("\n")]
    
    table_lines = [l for l in lines if l.startswith("|") and l.endswith("|") and "---" not in l]
    
    headers = []
    rows = []
    if len(table_lines) >= 2:
        headers = [c.strip() for c in table_lines[0].split("|")[1:-1]]
        for l in table_lines[1:]:
            parts = [c.strip() for c in l.split("|")[1:-1]]
            if len(parts) == len(headers):
                rows.append(parts)

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        if headers and rows:
            df = pd.DataFrame(rows, columns=headers)
            df.to_excel(writer, index=False, sheet_name="Tax_Data_Summary")
        else:
            advisory_sections = []
            current_section = "General Compliance"
            for l in lines:
                if l.startswith("### ") or l.startswith("## ") or l.startswith("# "):
                    current_section = l.replace("#", "").strip()
                elif l.startswith("- ") or l.startswith("• "):
                    advisory_sections.append({"Section": current_section, "Guidance & Rules": l[2:].replace("**", "").strip()})
                elif len(l) > 20 and not l.startswith("["):
                    advisory_sections.append({"Section": current_section, "Guidance & Rules": l.replace("**", "").strip()})
            
            if not advisory_sections:
                advisory_sections = [{"Section": "Statutory Assessment", "Guidance & Rules": cleaned_text[:3000]}]
                
            df = pd.DataFrame(advisory_sections)
            df.to_excel(writer, index=False, sheet_name="Advisory_Matrix")

        ws = writer.sheets[list(writer.sheets.keys())[0]]
        
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        regular_font = Font(name="Segoe UI", size=10, color="0F172A")
        
        thin_border = Border(
            left=Side(style='thin', color='E2E8F0'),
            right=Side(style='thin', color='E2E8F0'),
            top=Side(style='thin', color='E2E8F0'),
            bottom=Side(style='thin', color='E2E8F0')
        )
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        
        ws.row_dimensions[1].height = 28

        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=2):
            ws.row_dimensions[row_idx].height = 22
            is_even = (row_idx % 2 == 0)
            for cell in row:
                cell.font = regular_font
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                if is_even:
                    cell.fill = zebra_fill
                    
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = min(max(max_len + 5, 18), 65)

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

# --- 3. BACKGROUND API KEY ---
groq_key = st.secrets.get("GROQ_API_KEY", "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")

# --- 4. PERMANENT AUTO-LOGIN & STATE ---
if "user_id" not in st.session_state:
    st.session_state.user_id = None
if "user_name" not in st.session_state:
    st.session_state.user_name = None
if "current_convo_id" not in st.session_state:
    st.session_state.current_convo_id = None
if "chip_query" not in st.session_state:
    st.session_state.chip_query = None
if "active_attachment" not in st.session_state:
    st.session_state.active_attachment = None
if "uploader_id" not in st.session_state:
    st.session_state.uploader_id = 0
if "last_processed_voice" not in st.session_state:
    st.session_state.last_processed_voice = None

auth_token = st.query_params.get("session_auth")
if not st.session_state.user_id and auth_token:
    try:
        with get_db() as c_conn:
            cur = c_conn.cursor()
            cur.execute("SELECT user_id FROM sessions WHERE token=?", (auth_token,))
            row = cur.fetchone()
            if row:
                st.session_state.user_id = row[0]
                cur.execute("SELECT name, email FROM users WHERE id=?", (row[0],))
                u = cur.fetchone()
                if u:
                    st.session_state.user_name = u[0] if u[0] else (u[1].split('@')[0].title() if len(u) > 1 else "User")
                    st.session_state.user_email = u[1] if len(u) > 1 else ""
    except Exception:
        pass

# --- 5. AUTHENTICATION (Login / Sign Up) ---
if not st.session_state.user_id:
    st.markdown("<h2 style='text-align: center; color: #0D0D0D; margin-top: 2rem;'>🛡️ Kavach AI</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #64748B; font-size: 1.05rem; margin-bottom: 2rem;'>NextGen FinHR — Tax & Compliance Shield</p>", unsafe_allow_html=True)
    
    auth_tab1, auth_tab2 = st.tabs(["🔐 Sign In", "📝 Create Account"])
    
    with auth_tab1:
        with st.form("login_form"):
            login_email = st.text_input("Email Address")
            login_pw = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Sign In", use_container_width=True)
            if submitted:
                with get_db() as c_conn:
                    cur = c_conn.cursor()
                    cur.execute("SELECT id, name, email FROM users WHERE email=? AND password_hash=?", (login_email.strip().lower(), hash_val(login_pw)))
                    row = cur.fetchone()
                    if row:
                        new_token = secrets.token_hex(24)
                        cur.execute("INSERT OR REPLACE INTO sessions (token, user_id) VALUES (?, ?)", (new_token, row[0]))
                        c_conn.commit()
                        st.query_params["session_auth"] = new_token
                        st.session_state.user_id = row[0]
                        st.session_state.user_name = row[1] if row[1] else (row[2].split('@')[0].title() if len(row) > 2 else "User")
                        st.session_state.user_email = row[2] if len(row) > 2 else ""
                        st.rerun()
                    else:
                        st.error("Invalid email or password.")

    with auth_tab2:
        with st.form("signup_form"):
            new_name = st.text_input("Full Name", placeholder="e.g. Harsh Mishra")
            new_email = st.text_input("Email Address")
            new_pw = st.text_input("Password", type="password")
            create_btn = st.form_submit_button("Register Account", use_container_width=True)
            if create_btn:
                if not new_name.strip():
                    st.warning("Please enter your name.")
                elif len(new_pw) < 4:
                    st.warning("Password must be at least 4 characters.")
                else:
                    try:
                        with get_db() as c_conn:
                            c_conn.execute("INSERT INTO users (name, email, password_hash) VALUES (?, ?, ?)", 
                                           (new_name.strip(), new_email.strip().lower(), hash_val(new_pw)))
                            c_conn.commit()
                        st.success("Account created successfully! Please Sign In.")
                    except sqlite3.IntegrityError:
                        st.error("This email is already registered.")
    st.stop()

# --- 6. SIDEBAR: CHAT HISTORY & PROFILE ---
with st.sidebar:
    st.markdown("### 🛡️ Kavach AI")
    st.caption("NextGen FinHR Architecture")

    if st.button("➕ New Consultation", use_container_width=True):
        st.session_state.current_convo_id = None
        st.session_state.active_attachment = None
        st.session_state.uploader_id += 1
        st.rerun()

    st.markdown("### 💬 Consultations")
    with get_db() as c_conn:
        cur = c_conn.cursor()
        cur.execute("SELECT id, title FROM conversations WHERE user_id=? ORDER BY id DESC", (st.session_state.user_id,))
        convos = cur.fetchall()
    
    for c_id, c_title in convos:
        if st.button(f"📄 {c_title[:24]}...", key=f"convo_{c_id}", use_container_width=True):
            st.session_state.current_convo_id = c_id
            st.session_state.active_attachment = None
            st.session_state.uploader_id += 1
            st.rerun()

    st.divider()
    st.markdown("#### 👤 Account")
    current_user_name = st.session_state.get("user_name") or "User"
    st.markdown(f"Name: **{current_user_name}**")
    st.caption(f"Email: `{st.session_state.get('user_email', '')}`")
    
    with st.expander("⚙️ Account Settings", expanded=False):
        edit_name = st.text_input("Full Name", value=current_user_name, key="edit_profile_name")
        if st.button("Save Name", use_container_width=True):
            if edit_name.strip():
                with get_db() as c_conn:
                    c_conn.execute("UPDATE users SET name=? WHERE id=?", (edit_name.strip(), st.session_state.user_id))
                    c_conn.commit()
                st.session_state.user_name = edit_name.strip()
                st.success("Name updated!")
                st.rerun()
        
        new_pass = st.text_input("New Password", type="password")
        if st.button("Update Password", use_container_width=True):
            if len(new_pass) >= 4:
                with get_db() as c_conn:
                    c_conn.execute("UPDATE users SET password_hash=? WHERE id=?", (hash_val(new_pass), st.session_state.user_id))
                    c_conn.commit()
                st.success("Password updated!")
                
    if st.button("🚪 Sign Out", use_container_width=True):
        if "session_auth" in st.query_params:
            with get_db() as c_conn:
                c_conn.execute("DELETE FROM sessions WHERE token=?", (st.query_params["session_auth"],))
                c_conn.commit()
        st.query_params.clear()
        st.session_state.clear()
        st.rerun()

# --- 7. INTENT & MESSAGES EXTRACTION ---
def check_is_image_intent(query):
    q = query.lower()
    img_words = ["poster", "banner", "generate image", "make photo", "create graphic", "banao poster", "design photo"]
    return any(w in q for w in img_words)

current_messages = []
full_chat_text = ""
if st.session_state.current_convo_id:
    with get_db() as c_conn:
        cur = c_conn.cursor()
        cur.execute("SELECT role, content FROM messages WHERE conversation_id=? ORDER BY id ASC", (st.session_state.current_convo_id,))
        current_messages = [{"role": r, "content": c} for r, c in cur.fetchall()]

# Greeting Name
user_greeting_name = st.session_state.get("user_name") or "there"

# --- 8. HERO GREETING & CHIPS (EMPTY CHAT STATE) ---
if len(current_messages) == 0:
    st.markdown(f"""
    <div style="text-align: left; margin-top: 2rem; margin-bottom: 2rem;">
        <h1 style="font-size: 2.2rem; font-weight: 600; color: #0D0D0D; margin-bottom: 0.4rem;">Hello, {user_greeting_name}</h1>
        <p style="font-size: 1.2rem; color: #666666; font-weight: 400;">How can Kavach AI shield your business today?</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📋 Check GST ITC Eligibility\n\nConditions & rules to claim credit under GST", use_container_width=True, key="chip_itc"):
            st.session_state.chip_query = "What are the key conditions and statutory rules to claim Input Tax Credit (ITC) under GST?"
            st.rerun()
            
        if st.button("⚖️ Analyze Section 148 Notice\n\nIncome Tax reassessment guidance and timeline", use_container_width=True, key="chip_148"):
            st.session_state.chip_query = "Explain how to respond to an Income Tax notice received under Section 148 and Section 148A."
            st.rerun()

    with col2:
        if st.button("📊 TDS & TCS Rules FY 2025-26\n\nThreshold limits and current deduction rates", use_container_width=True, key="chip_tds"):
            st.session_state.chip_query = "Explain Tax Collected at Source (TCS) and TDS rates under Section 206C and GST Section 52 with limits."
            st.rerun()
            
        if st.button("💼 New vs Old Tax Regime\n\nDirect comparison for business and salaried individuals", use_container_width=True, key="chip_regime"):
            st.session_state.chip_query = "Compare the New Tax Regime vs Old Tax Regime with key deductions, slabs, and break-even points."
            st.rerun()

# --- 9. RENDER MESSAGES (CHATGPT STYLE: DISPLAY USER IMAGES + TEXT) ---
for msg in current_messages:
    with st.chat_message(msg["role"]):
        content = msg["content"]
        if content.startswith("[IMAGE_URL]:"):
            img_url = content.replace("[IMAGE_URL]:", "").strip()
            st.image(img_url, caption="Generated Visual Graphic", use_container_width=True)
        elif "[ATTACHED_IMAGE]:" in content:
            parts = content.split("\n[PROMPT]:")
            b64_part = parts[0].replace("[ATTACHED_IMAGE]:", "").strip()
            prompt_text = parts[1] if len(parts) > 1 else ""
            st.image(f"data:image/jpeg;base64,{b64_part}", width=320)
            if prompt_text.strip():
                st.markdown(prompt_text.strip())
        else:
            display_text = clean_html_tags(content)
            st.markdown(display_text)
            full_chat_text += f"\n\n[{msg['role'].upper()}]:\n" + display_text

# Top Export Bar
if current_messages:
    exp_col1, exp_col2 = st.columns(2)
    with exp_col1:
        docx_file = generate_docx("Kavach AI - Advisory Report", full_chat_text)
        st.download_button("📥 Export Consultation to Word (.docx)", data=docx_file, file_name="Kavach_AI_Advisory_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with exp_col2:
        xlsx_file = generate_xlsx(full_chat_text)
        st.download_button("📊 Export Consultation to Excel (.xlsx)", data=xlsx_file, file_name="Kavach_AI_Summary_Data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# --- 10. CHATGPT-STYLE ATTACHMENT & VOICE WIDGETS ---
c_att_box, c_mic_box = st.columns([1, 1])

with c_att_box:
    with st.popover("📎 Attach Image / Document", use_container_width=True):
        new_upload = st.file_uploader(
            "Upload client invoice, tax notice photo, PDF, or Excel", 
            type=["png", "jpg", "jpeg", "pdf", "xlsx", "csv", "json"],
            label_visibility="collapsed",
            key=f"file_uploader_widget_{st.session_state.uploader_id}"
        )
        if new_upload is not None:
            file_bytes = new_upload.read()
            file_ext = new_upload.name.lower().split(".")[-1]
            st.session_state.active_attachment = {
                "name": new_upload.name,
                "bytes": file_bytes,
                "ext": file_ext
            }

with c_mic_box:
    with st.popover("🎙️ Voice Query (बोलकर पूछें)", use_container_width=True):
        st.caption("माइक पर क्लिक करें, सवाल बोलें और दोबारा क्लिक करके स्टॉप करें:")
        voice_record = st.audio_input("Record voice question", label_visibility="collapsed", key=f"audio_recorder_{st.session_state.uploader_id}")
        if voice_record is not None:
            v_bytes = voice_record.read()
            v_hash = hashlib.md5(v_bytes).hexdigest()
            if st.session_state.last_processed_voice != v_hash:
                st.session_state.last_processed_voice = v_hash
                with st.spinner("🎙️ आवाज़ को टेक्स्ट में बदला जा रहा है..."):
                    try:
                        v_client = Groq(api_key=groq_key)
                        transcription = v_client.audio.transcriptions.create(
                            file=("voice_prompt.wav", v_bytes),
                            model="whisper-large-v3-turbo",
                            language="hi"
                        )
                        spoken_text = transcription.text.strip()
                        if spoken_text:
                            st.session_state.chip_query = spoken_text
                            st.rerun()
                    except Exception as e:
                        st.error(f"Voice Transcription Error: {e}")

# Live Attachment Banner (Shown right above the prompt box like ChatGPT)
if st.session_state.active_attachment:
    att = st.session_state.active_attachment
    c_preview, c_del = st.columns([4, 1])
    with c_preview:
        if att["ext"] in ["png", "jpg", "jpeg"]:
            st.image(att["bytes"], width=80)
            st.caption(f"🖼️ Attached Image: **{att['name']}** (Ready with prompt)")
        else:
            st.info(f"📄 Attached Document: **{att['name']}** (Ready with prompt)")
    with c_del:
        if st.button("❌ Remove", use_container_width=True, key="btn_remove_att"):
            st.session_state.active_attachment = None
            st.session_state.uploader_id += 1
            st.rerun()

# --- 11. CHAT INPUT & EXECUTION PIPELINE ---
chat_input_val = st.chat_input("Ask anything about Tax, GST, Compliance, Payroll, or ask about attached image/doc...")

user_query = None
if st.session_state.chip_query:
    user_query = st.session_state.chip_query
    st.session_state.chip_query = None
elif chat_input_val:
    user_query = chat_input_val

if user_query:
    if not st.session_state.current_convo_id:
        title = user_query[:35]
        with get_db() as c_conn:
            cur = c_conn.cursor()
            cur.execute("INSERT INTO conversations (user_id, title) VALUES (?, ?)", (st.session_state.user_id, title))
            c_conn.commit()
            st.session_state.current_convo_id = cur.lastrowid

    # Process Attachment Data
    attached_data_text = ""
    attached_image_b64 = None
    save_user_content = user_query

    if st.session_state.active_attachment:
        att = st.session_state.active_attachment
        if att["ext"] in ["png", "jpg", "jpeg"]:
            attached_image_b64 = base64.b64encode(att["bytes"]).decode("utf-8")
            save_user_content = f"[ATTACHED_IMAGE]:{attached_image_b64}\n[PROMPT]: {user_query}"
        elif att["ext"] == "pdf":
            try:
                reader = PdfReader(io.BytesIO(att["bytes"]))
                pdf_text = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
                attached_data_text = f"\n[User Attached PDF Document ({att['name']}) Content]:\n{pdf_text[:15000]}"
                save_user_content = f"📎 *Attached: {att['name']}*\n\n{user_query}"
            except Exception as e:
                attached_data_text = f"\n[Error reading PDF {att['name']}: {e}]"
        elif att["ext"] in ["xlsx", "xls"]:
            try:
                df = pd.read_excel(io.BytesIO(att["bytes"]))
                attached_data_text = f"\n[User Attached Excel Sheet ({att['name']}) Data]:\n{df.head(50).to_markdown()}"
                save_user_content = f"📊 *Attached Excel: {att['name']}*\n\n{user_query}"
            except Exception as e:
                attached_data_text = f"\n[Error reading Excel: {e}]"
        elif att["ext"] == "csv":
            try:
                df = pd.read_csv(io.BytesIO(att["bytes"]))
                attached_data_text = f"\n[User Attached CSV File ({att['name']}) Data]:\n{df.head(50).to_markdown()}"
                save_user_content = f"📊 *Attached CSV: {att['name']}*\n\n{user_query}"
            except Exception as e:
                attached_data_text = f"\n[Error reading CSV: {e}]"
        elif att["ext"] == "json":
            try:
                j_obj = json.loads(att["bytes"].decode("utf-8"))
                attached_data_text = f"\n[User Attached JSON ({att['name']}) Data]:\n{json.dumps(j_obj, indent=2)[:15000]}"
                save_user_content = f"📄 *Attached JSON: {att['name']}*\n\n{user_query}"
            except Exception as e:
                attached_data_text = f"\n[Error reading JSON: {e}]"
        
        st.session_state.active_attachment = None
        st.session_state.uploader_id += 1

    # Save to Database
    with get_db() as c_conn:
        c_conn.execute("INSERT INTO messages (conversation_id, role, content) VALUES (?, 'user', ?)", (st.session_state.current_convo_id, save_user_content))
        c_conn.commit()

    with st.chat_message("user"):
        if attached_image_b64:
            st.image(f"data:image/jpeg;base64,{attached_image_b64}", width=320)
        st.markdown(user_query)

    # 1. Statutory Context from Local Vector DB
    statutory_context = ""
    if vector_db:
        lookup_token = "206C" if user_query.strip().upper() == "TCS" else None
        if not lookup_token:
            sec_match = re.search(r'\b(?:section|sec|rule|धारा|नियम)?\s*([0-9]{1,4}[a-z]{0,3})\b', user_query, re.IGNORECASE)
            lookup_token = sec_match.group(1).upper() if sec_match else None
        
        retrieved_docs = []
        if lookup_token and len(lookup_token) >= 2:
            try:
                retrieved_docs = vector_db.similarity_search(user_query, k=6, where_document={"$contains": lookup_token})
            except Exception:
                retrieved_docs = []
        
        if not retrieved_docs:
            retrieved_docs = vector_db.similarity_search(user_query, k=6)
            
        statutory_context = "\n\n---\n\n".join([d.page_content for d in retrieved_docs])

    # 2. Live Web Search
    live_web_context = perform_live_web_search(user_query)

    # --- ADVANCED SYSTEM INSTRUCTION ---
    system_instruction = f"""
    You are Kavach AI (powered by NextGen FinHR Architecture), an expert senior authority in:
    1. Indian Finance & Taxation (Income Tax Act 1961, CGST/SGST/IGST Acts, Corporate Tax, Audits)
    2. HR & Labour Law Compliance (Payroll, EPF/ESIC, Gratuity, Bonus, Labour Codes)

    CRITICAL TAX ACRONYM RULE:
    - Your core primary domain is INDIAN TAXATION and BUSINESS COMPLIANCE.
    - Whenever the user types an acronym like "TCS", "TDS", "ITC", "GST", "MAT", "AMT", "PF", or "ESIC", you MUST ALWAYS interpret it as its TAX / COMPLIANCE definition.
    - Specifically, "TCS" MUST ALWAYS be answered as **TCS — Tax Collected at Source** under the Income Tax Act (Section 206C) and GST Act (Section 52 - E-Commerce Operator). NEVER discuss Tata Consultancy Services or stock markets unless the user specifically asks about "Tata shares" or "TCS stock price".

    VISION & DOCUMENT ANALYSIS RULE:
    - If the user has provided an attached image or document, meticulously read every number, tax calculation, GSTIN, invoice date, notice section, and clause. 
    - Formulate your response around the attached visual data.

    STRUCTURE OF YOUR RESPONSE (EXACT CHATGPT STYLE):
    1. Direct Concept Header (e.g. `### TCS — Tax Collected at Source`)
    2. The Major Contexts & Sections
    3. Practical Real-World Corporate Example (Scenario + Math Calculation Table)
    4. Compliance Deadlines & Action Items

    FORMATTING RULES:
    - Clean Markdown only (No raw HTML tags).
    - Use clean Markdown tables for numbers and calculations.
    - Professional, authoritative, and helpful tone in Hinglish / English.

    [Internal Statutory Context]:
    {statutory_context}

    [Live Internet Data]:
    {live_web_context}

    [Client Attached Document Data]:
    {attached_data_text}
    """

    with st.chat_message("assistant"):
        message_placeholder = st.empty()
        full_response = ""
        client = Groq(api_key=groq_key)
        last_error = None
        
        try:
            available_models = [m.id for m in client.models.list().data if "whisper" not in m.id and "guard" not in m.id]
        except Exception:
            available_models = []

        # --- MULTIMODAL VISION ROUTING (IMAGE QUERIES) ---
        if attached_image_b64:
            preferred_vision = ["qwen/qwen3.8-27b", "qwen/qwen3.6-27b"]
            vision_models = [m for m in preferred_vision if m in available_models] or preferred_vision
            
            vision_system = (
                "You are Kavach AI, an expert in Indian Taxation, Finance, GST, and HR Compliance. "
                "Examine the attached image (invoice, tax notice, receipt, or ledger) thoroughly and provide "
                "structured analysis, tax breakdown, and statutory action items."
            )
            
            img_format = "png" if (att and att.get("ext") == "png") else "jpeg"
            vision_messages = [
                {"role": "system", "content": vision_system},
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": f"{user_query}\n\nPlease inspect all numbers, tax sections, dates, and details from this image."},
                        {"type": "image_url", "image_url": {"url": f"data:image/{img_format};base64,{attached_image_b64}"}}
                    ]
                }
            ]
            
            for v_model in vision_models:
                try:
                    completion = client.chat.completions.create(
                        model=v_model,
                        messages=vision_messages,
                        temperature=0.2
                    )
                    full_response = completion.choices[0].message.content
                    message_placeholder.markdown(full_response)
                    break
                except Exception as e:
                    last_error = e
                    continue
                    
        # --- TEXT / DOCUMENT PIPELINE ---
        else:
            llm_messages = [{"role": "system", "content": system_instruction}]
            for prev_msg in current_messages[-8:]:
                if not prev_msg["content"].startswith("[IMAGE_URL]:") and not prev_msg["content"].startswith("[ATTACHED_IMAGE]:"):
                    llm_messages.append({"role": prev_msg["role"], "content": prev_msg["content"]})
            llm_messages.append({"role": "user", "content": user_query})

            preferred_text = ["openai/gpt-oss-120b", "openai/gpt-oss-20b", "qwen/qwen3.8-27b"]
            text_models = [m for m in preferred_text if m in available_models] or preferred_text
            
            for m_name in text_models:
                try:
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
                            message_placeholder.markdown(clean_html_tags(full_response) + "▌")

                    if full_response.strip():
                        full_response = clean_html_tags(full_response)
                        message_placeholder.markdown(full_response)
                        break
                except Exception as e:
                    last_error = e
                    continue

        if not full_response.strip():
            if last_error:
                st.error(f"Groq API Error: {last_error}")
            else:
                st.error("Error processing query. Please check your Groq API connection.")
            st.stop()

    with get_db() as c_conn:
        c_conn.execute("INSERT INTO messages (conversation_id, role, content) VALUES (?, 'assistant', ?)", (st.session_state.current_convo_id, full_response))
        c_conn.commit()
    
    st.rerun()