import streamlit as st
from groq import Groq
import pandas as pd
from datetime import datetime, date
import os
import io
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor, white, black
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from audio_recorder_streamlit import audio_recorder
import tempfile

# ─── PAGE CONFIG ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="BIOS Medical Intelligence",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── STYLING ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Main background */
    .stApp { background-color: #0D1B2A; }
    section[data-testid="stSidebar"] { background-color: #0A1628; border-right: 1px solid #1B6CA8; }

    /* Hide default streamlit elements */
    #MainMenu, footer, header { visibility: hidden; }

    /* Custom header */
    .bios-header {
        background: linear-gradient(135deg, #0D1B2A 0%, #1B3A5C 100%);
        border: 1px solid #1B6CA8;
        border-radius: 12px;
        padding: 20px 28px;
        margin-bottom: 24px;
    }
    .bios-title { color: #FFFFFF; font-size: 28px; font-weight: 800; margin: 0; }
    .bios-sub { color: #60A5FA; font-size: 13px; margin: 4px 0 0 0; letter-spacing: 2px; text-transform: uppercase; }

    /* Cards */
    .card {
        background: #112240;
        border: 1px solid #1B3A5C;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 16px;
    }
    .card-title { color: #60A5FA; font-size: 13px; font-weight: 600; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 12px; }

    /* Input styling */
    .stTextArea textarea {
        background-color: #0A1628 !important;
        color: #E2E8F0 !important;
        border: 1px solid #1B6CA8 !important;
        border-radius: 8px !important;
        font-size: 14px !important;
    }
    .stTextInput input {
        background-color: #0A1628 !important;
        color: #E2E8F0 !important;
        border: 1px solid #1B6CA8 !important;
        border-radius: 8px !important;
    }

    /* Buttons */
    .stButton button {
        background: linear-gradient(135deg, #1B6CA8, #2563EB) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        padding: 10px 20px !important;
        width: 100% !important;
    }
    .stButton button:hover {
        background: linear-gradient(135deg, #2563EB, #1D4ED8) !important;
        transform: translateY(-1px);
    }

    /* SOAP output */
    .soap-output {
        background: #112240;
        border: 1px solid #1B6CA8;
        border-left: 4px solid #3B82F6;
        border-radius: 8px;
        padding: 24px;
        color: #E2E8F0;
        font-size: 14px;
        line-height: 1.8;
        white-space: pre-wrap;
    }

    /* Stats cards */
    .stat-card {
        background: #112240;
        border: 1px solid #1B3A5C;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
    }
    .stat-number { color: #3B82F6; font-size: 28px; font-weight: 800; }
    .stat-label { color: #94A3B8; font-size: 12px; text-transform: uppercase; letter-spacing: 1px; }

    /* Patient ID badge */
    .patient-badge {
        background: #1E3A5F;
        border: 1px solid #3B82F6;
        border-radius: 20px;
        padding: 4px 14px;
        color: #60A5FA;
        font-size: 12px;
        font-weight: 600;
        display: inline-block;
    }

    /* Alert boxes */
    .alert-red {
        background: #1F0A0A;
        border-left: 4px solid #EF4444;
        border-radius: 6px;
        padding: 12px 16px;
        color: #FCA5A5;
        font-size: 13px;
    }
    .alert-green {
        background: #0A1F0A;
        border-left: 4px solid #22C55E;
        border-radius: 6px;
        padding: 12px 16px;
        color: #86EFAC;
        font-size: 13px;
    }

    /* Sidebar text */
    .sidebar-label { color: #94A3B8; font-size: 11px; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px; }
    .sidebar-name { color: #E2E8F0; font-size: 16px; font-weight: 600; }
    .sidebar-role { color: #3B82F6; font-size: 12px; }

    /* Select box */
    .stSelectbox select { background-color: #0A1628 !important; color: #E2E8F0 !important; }

    /* Labels */
    label { color: #94A3B8 !important; }
    .stMarkdown { color: #E2E8F0; }
</style>
""", unsafe_allow_html=True)

# ─── GROQ CLIENT ────────────────────────────────────────────────────────────
client = Groq(api_key=st.secrets["GROQ_API_KEY"])

# ─── SYSTEM PROMPT ──────────────────────────────────────────────────────────
SYSTEM_PROMPT = """You are BIOS Clinical Assistant — an expert AI for Pakistani doctors.

Convert rough consultation notes (Urdu, English, or mixed) into a clean professional SOAP note.

Format your output EXACTLY like this:

CHIEF COMPLAINT
[Main problem]

HISTORY OF PRESENT ILLNESS  
[Onset, duration, severity, associated symptoms]

PAST MEDICAL HISTORY
[Previous conditions, surgeries, medications, allergies]

EXAMINATION FINDINGS
[Vitals, physical examination findings]

ASSESSMENT
[Clinical impression / working diagnosis]

PLAN
[Investigations, treatment, medications, follow-up]

RED FLAGS ⚠️
[Serious warning signs — write NONE if not present]

Rules:
- Missing info: write "Not mentioned"
- Urdu input: output in English
- Never invent information
- Be concise and clinically precise
- Always check for red flags like chest pain, stroke symptoms, severe breathlessness"""
- "Bazu" means ARM, "pair" means LEG, "ghutna" means KNEE translate Urdu body parts accurately
- "Tota/Tooti" means FRACTURE/BROKEN
- "Bukhaar" means FEVER, "sar dard" means HEADACHE
- Always specify LEFT or RIGHT if mentioned

# ─── SESSION STATE ───────────────────────────────────────────────────────────
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "doctor_name" not in st.session_state:
    st.session_state.doctor_name = ""
if "doctor_id" not in st.session_state:
    st.session_state.doctor_id = ""
if "patient_count" not in st.session_state:
    st.session_state.patient_count = 0
if "records" not in st.session_state:
    st.session_state.records = []

# ─── DOCTOR ACCOUNTS (Simple — expand later with Supabase) ──────────────────
DOCTORS = {
    "dr.ahmed": {"password": "ahmed123", "name": "Dr. Ahmed Khan", "specialty": "General Physician"},
    "dr.fatima": {"password": "fatima123", "name": "Dr. Fatima Ali", "specialty": "Internal Medicine"},
    "dr.hassan": {"password": "hassan123", "name": "Dr. Hassan Malik", "specialty": "Paediatrician"},
    "dr.demo": {"password": "demo123", "name": "Dr. Demo User", "specialty": "General Physician"},
}

# ─── HELPER FUNCTIONS ────────────────────────────────────────────────────────

def generate_patient_id():
    now = datetime.now()
    st.session_state.patient_count += 1
    return f"BIOS-{now.strftime('%Y%m%d')}-{st.session_state.patient_count:04d}"

def get_excel_filename(doctor_id):
    today = date.today().strftime("%Y-%m-%d")
    return f"BIOS_{doctor_id}_{today}_records.xlsx"

def save_to_excel(record, doctor_id):
    filename = get_excel_filename(doctor_id)
    
    # Check if file exists in session
    excel_key = f"excel_{doctor_id}_{date.today()}"
    
    if excel_key not in st.session_state:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Records {date.today()}"
        
        # Header styling
        headers = ["Patient ID", "Time", "Doctor", "Chief Complaint", "Assessment", "Plan", "Red Flags", "Full Note"]
        header_fill = PatternFill(start_color="0D1B2A", end_color="0D1B2A", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        ws.row_dimensions[1].height = 25
        ws.column_dimensions['A'].width = 22
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 25
        ws.column_dimensions['E'].width = 25
        ws.column_dimensions['F'].width = 30
        ws.column_dimensions['G'].width = 20
        ws.column_dimensions['H'].width = 50
        
        st.session_state[excel_key] = wb
    
    wb = st.session_state[excel_key]
    ws = wb.active
    
    # Alternate row colors
    row_num = ws.max_row + 1
    row_fill_color = "EFF6FF" if row_num % 2 == 0 else "FFFFFF"
    row_fill = PatternFill(start_color=row_fill_color, end_color=row_fill_color, fill_type="solid")
    
    row_data = [
        record["patient_id"],
        record["time"],
        record["doctor"],
        record["chief_complaint"],
        record["assessment"],
        record["plan"],
        record["red_flags"],
        record["full_note"]
    ]
    
    for col, value in enumerate(row_data, 1):
        cell = ws.cell(row=row_num, column=col, value=value)
        cell.fill = row_fill
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    
    # Save to bytes
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    st.session_state[excel_key + "_bytes"] = excel_buffer.getvalue()
    
    return filename

def generate_pdf(structured_note, patient_id, doctor_name, patient_name=""):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )
    
    story = []
    
    # Colors
    navy = HexColor('#0D1B2A')
    blue = HexColor('#1B6CA8')
    light_blue = HexColor('#EFF6FF')
    red = HexColor('#EF4444')
    gray = HexColor('#64748B')
    
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle('Title', fontSize=22, textColor=white,
                                  fontName='Helvetica-Bold', alignment=TA_CENTER, spaceAfter=4)
    sub_style = ParagraphStyle('Sub', fontSize=10, textColor=HexColor('#93C5FD'),
                                fontName='Helvetica', alignment=TA_CENTER)
    section_style = ParagraphStyle('Section', fontSize=11, textColor=white,
                                    fontName='Helvetica-Bold')
    body_style = ParagraphStyle('Body', fontSize=10, textColor=HexColor('#1E293B'),
                                 fontName='Helvetica', leading=16, spaceAfter=8)
    
    # ── HEADER BOX ──
    header_data = [[
        Paragraph("🏥 BIOS", title_style),
        Paragraph("Medical Intelligence System", sub_style)
    ]]
    header_table = Table(header_data, colWidths=[17*cm])
    header_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), navy),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 14),
        ('BOTTOMPADDING', (0,0), (-1,-1), 14),
        ('ROUNDEDCORNERS', (0,0), (-1,-1), [8, 8, 8, 8]),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 0.4*cm))
    
    # ── META INFO ──
    now = datetime.now().strftime("%d %B %Y · %I:%M %p")
    meta_data = [
        [Paragraph(f"<b>Patient ID:</b> {patient_id}", body_style),
         Paragraph(f"<b>Date:</b> {now}", body_style)],
        [Paragraph(f"<b>Doctor:</b> {doctor_name}", body_style),
         Paragraph(f"<b>Patient:</b> {patient_name if patient_name else 'Not specified'}", body_style)],
    ]
    meta_table = Table(meta_data, colWidths=[8.5*cm, 8.5*cm])
    meta_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), light_blue),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 12),
        ('RIGHTPADDING', (0,0), (-1,-1), 12),
        ('GRID', (0,0), (-1,-1), 0.5, HexColor('#BFDBFE')),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 0.5*cm))
    
    # ── DIVIDER ──
    story.append(HRFlowable(width="100%", thickness=2, color=blue))
    story.append(Spacer(1, 0.3*cm))
    
    # ── SOAP SECTIONS ──
    sections = [
        ("CHIEF COMPLAINT", "CC"),
        ("HISTORY OF PRESENT ILLNESS", "HPI"),
        ("PAST MEDICAL HISTORY", "PMH"),
        ("EXAMINATION FINDINGS", "EX"),
        ("ASSESSMENT", "A"),
        ("PLAN", "P"),
        ("RED FLAGS ⚠️", "RF"),
    ]
    
    lines = structured_note.split('\n')
    current_section = None
    section_content = {}
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        matched = False
        for section_name, _ in sections:
            if section_name in line.upper():
                current_section = section_name
                section_content[current_section] = []
                matched = True
                break
        if not matched and current_section:
            section_content[current_section].append(line)
    
    for section_name, abbr in sections:
        content = section_content.get(section_name, ["Not mentioned"])
        content_text = ' '.join([c for c in content if c]).strip()
        if not content_text:
            content_text = "Not mentioned"
        
        is_red_flag = "RED FLAG" in section_name
        bg_color = HexColor('#FEF2F2') if is_red_flag else light_blue
        header_color = red if is_red_flag else navy
        
        section_data = [[
            Paragraph(f'<font color="white"><b>{section_name}</b></font>',
                     ParagraphStyle('SH', fontSize=10, fontName='Helvetica-Bold',
                                   textColor=white, alignment=TA_LEFT)),
        ], [
            Paragraph(content_text,
                     ParagraphStyle('SC', fontSize=10, fontName='Helvetica',
                                   textColor=HexColor('#1E293B'), leading=15)),
        ]]
        
        section_table = Table(section_data, colWidths=[17*cm])
        section_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), header_color),
            ('BACKGROUND', (0,1), (-1,1), bg_color),
            ('TOPPADDING', (0,0), (-1,0), 7),
            ('BOTTOMPADDING', (0,0), (-1,0), 7),
            ('LEFTPADDING', (0,0), (-1,-1), 12),
            ('RIGHTPADDING', (0,0), (-1,-1), 12),
            ('TOPPADDING', (0,1), (-1,1), 10),
            ('BOTTOMPADDING', (0,1), (-1,1), 10),
        ]))
        story.append(section_table)
        story.append(Spacer(1, 0.25*cm))
    
    # ── FOOTER ──
    story.append(Spacer(1, 0.3*cm))
    story.append(HRFlowable(width="100%", thickness=1, color=HexColor('#CBD5E1')))
    story.append(Spacer(1, 0.2*cm))
    footer_style = ParagraphStyle('Footer', fontSize=8, textColor=gray,
                                   fontName='Helvetica', alignment=TA_CENTER)
    story.append(Paragraph(
        "Generated by BIOS Medical Intelligence System · Confidential Clinical Document · For authorized use only",
        footer_style
    ))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

def transcribe_audio(audio_bytes):
    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmp:
        tmp.write(audio_bytes)
        tmp_path = tmp.name
    
    with open(tmp_path, "rb") as audio_file:
        transcription = client.audio.transcriptions.create(
            model="whisper-large-v3",
            file=audio_file,
            response_format="text"
        )
    os.unlink(tmp_path)
    return transcription

def extract_sections(note):
    sections = {"chief_complaint": "", "assessment": "", "plan": "", "red_flags": "None"}
    lines = note.split('\n')
    current = None
    for line in lines:
        line = line.strip()
        if 'CHIEF COMPLAINT' in line.upper(): current = 'chief_complaint'
        elif 'ASSESSMENT' in line.upper(): current = 'assessment'
        elif 'PLAN' in line.upper(): current = 'plan'
        elif 'RED FLAG' in line.upper(): current = 'red_flags'
        elif current and line and line not in ['**', '##']:
            clean = line.replace('**', '').replace('##', '').strip()
            if clean:
                sections[current] += clean + " "
    return {k: v.strip() for k, v in sections.items()}

# ─── LOGIN PAGE ──────────────────────────────────────────────────────────────
def show_login():
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("""
        <div style='text-align:center; margin-bottom:32px;'>
            <div style='font-size:48px;'>🏥</div>
            <h1 style='color:#FFFFFF; font-size:32px; font-weight:800; margin:8px 0;'>BIOS</h1>
            <p style='color:#60A5FA; font-size:13px; letter-spacing:3px; text-transform:uppercase;'>Medical Intelligence System</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<div class='card'>", unsafe_allow_html=True)
        st.markdown("<p class='card-title'>Doctor Login</p>", unsafe_allow_html=True)
        
        username = st.text_input("Username", placeholder="dr.ahmed")
        password = st.text_input("Password", type="password", placeholder="••••••••")
        
        if st.button("🔐 Login to BIOS"):
            if username in DOCTORS and DOCTORS[username]["password"] == password:
                st.session_state.logged_in = True
                st.session_state.doctor_name = DOCTORS[username]["name"]
                st.session_state.doctor_id = username
                st.session_state.doctor_specialty = DOCTORS[username]["specialty"]
                st.rerun()
            else:
                st.markdown("<div class='alert-red'>❌ Invalid credentials. Please try again.</div>", unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("""
        <p style='text-align:center; color:#475569; font-size:12px; margin-top:24px;'>
        Demo: username <b style='color:#60A5FA'>dr.demo</b> · password <b style='color:#60A5FA'>demo123</b>
        </p>
        """, unsafe_allow_html=True)

# ─── MAIN APP ────────────────────────────────────────────────────────────────
def show_app():
    
    # ── SIDEBAR ──
    with st.sidebar:
        st.markdown(f"""
        <div style='padding:16px 0 24px 0; border-bottom:1px solid #1B3A5C; margin-bottom:20px;'>
            <div class='sidebar-label'>Logged in as</div>
            <div class='sidebar-name'>{st.session_state.doctor_name}</div>
            <div class='sidebar-role'>{st.session_state.get('doctor_specialty','')}</div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<div class='sidebar-label'>Navigation</div>", unsafe_allow_html=True)
        page = st.selectbox("", ["📝 New Consultation", "📊 Today's Records", "ℹ️ About BIOS"], label_visibility="collapsed")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Stats
        today_records = [r for r in st.session_state.records if r.get("doctor_id") == st.session_state.doctor_id]
        st.markdown(f"""
        <div class='stat-card' style='margin-bottom:12px;'>
            <div class='stat-number'>{len(today_records)}</div>
            <div class='stat-label'>Today's Patients</div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<br><br>", unsafe_allow_html=True)
        if st.button("🚪 Logout"):
            st.session_state.logged_in = False
            st.rerun()
        
        st.markdown("""
        <div style='position:absolute; bottom:20px; left:16px; right:16px;'>
            <p style='color:#334155; font-size:11px; text-align:center;'>
            BIOS Medical Intelligence<br>v2.0 · Confidential
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    # ── NEW CONSULTATION PAGE ──
    if "New Consultation" in page:
        
        # Header
        st.markdown("""
        <div class='bios-header'>
            <p class='bios-title'>🏥 BIOS Clinical Note Structurer</p>
            <p class='bios-sub'>Medical Intelligence System · AI-Powered SOAP Notes</p>
        </div>
        """, unsafe_allow_html=True)
        
        col_left, col_right = st.columns([1, 1], gap="large")
        
        with col_left:
            st.markdown("<div class='card'>", unsafe_allow_html=True)
            st.markdown("<p class='card-title'>📋 Patient Information</p>", unsafe_allow_html=True)
            patient_name = st.text_input("Patient Name (Optional)", placeholder="e.g. Ahmed Khan")
            st.markdown("</div>", unsafe_allow_html=True)
            
            st.markdown("<div class='card'>", unsafe_allow_html=True)
            st.markdown("<p class='card-title'>🎙️ Input Method</p>", unsafe_allow_html=True)
            input_method = st.radio("", ["✍️ Type Notes", "🎙️ Voice Input"], horizontal=True, label_visibility="collapsed")
            
            raw_notes = ""
            
            if "Type" in input_method:
                raw_notes = st.text_area(
                    "Doctor ki rough notes — Urdu ya English mein",
                    height=220,
                    placeholder="Misaal: 45 saal ki khatoon, subah se seene mein dard, baen haath mein phailta hai, paseena aa raha hai, BP 150/95, diabetes ki history hai, metformin le rahi hain...",
                    label_visibility="collapsed"
                )
            else:
                st.markdown("""
                <p style='color:#94A3B8; font-size:13px; margin-bottom:8px;'>
                🎙️ Record button dabao aur patient ke baare mein bolna shuru karo
                </p>
                """, unsafe_allow_html=True)
                
                audio_bytes = audio_recorder(
                    text="Record karo",
                    recording_color="#EF4444",
                    neutral_color="#1B6CA8",
                    icon_size="2x"
                )
                
                if audio_bytes:
                    st.audio(audio_bytes, format="audio/wav")
                    with st.spinner("🎙️ Transcribe ho raha hai..."):
                        try:
                            raw_notes = transcribe_audio(audio_bytes)
                            st.markdown(f"""
                            <div class='alert-green'>
                            ✅ Transcription complete:<br>
                            <i style='color:#D1FAE5;'>"{raw_notes}"</i>
                            </div>
                            """, unsafe_allow_html=True)
                        except Exception as e:
                            st.error(f"Transcription error: {str(e)}")
            
            st.markdown("</div>", unsafe_allow_html=True)
            
            process_btn = st.button("⚡ BIOS AI se Structure Karo")
        
        with col_right:
            st.markdown("<div class='card' style='min-height:500px;'>", unsafe_allow_html=True)
            st.markdown("<p class='card-title'>📄 Structured SOAP Note</p>", unsafe_allow_html=True)
            
            if process_btn:
                if not raw_notes or raw_notes.strip() == "":
                    st.markdown("<div class='alert-red'>⚠️ Pehle notes likhein ya voice record karein.</div>", unsafe_allow_html=True)
                else:
                    with st.spinner("🧠 BIOS AI process kar raha hai..."):
                        try:
                            response = client.chat.completions.create(
                                model="llama-3.1-8b-instant",
                                messages=[
                                    {"role": "system", "content": SYSTEM_PROMPT},
                                    {"role": "user", "content": f"Structure these clinical notes:\n\n{raw_notes}"}
                                ],
                                temperature=0.2,
                                max_tokens=1200
                            )
                            
                            structured_note = response.choices[0].message.content
                            patient_id = generate_patient_id()
                            sections = extract_sections(structured_note)
                            
                            # Save record
                            record = {
                                "patient_id": patient_id,
                                "time": datetime.now().strftime("%H:%M"),
                                "doctor": st.session_state.doctor_name,
                                "doctor_id": st.session_state.doctor_id,
                                "patient_name": patient_name,
                                "chief_complaint": sections["chief_complaint"],
                                "assessment": sections["assessment"],
                                "plan": sections["plan"],
                                "red_flags": sections["red_flags"],
                                "full_note": structured_note,
                                "raw_notes": raw_notes
                            }
                            st.session_state.records.append(record)
                            excel_filename = save_to_excel(record, st.session_state.doctor_id)
                            
                            # Show patient ID
                            st.markdown(f"""
                            <div style='margin-bottom:16px;'>
                                <span class='patient-badge'>🆔 {patient_id}</span>
                            </div>
                            """, unsafe_allow_html=True)
                            
                            # Red flag check
                            if sections["red_flags"] and sections["red_flags"].lower() not in ["none", "none identified", "not mentioned"]:
                                st.markdown(f"""
                                <div class='alert-red'>
                                ⚠️ <b>RED FLAG DETECTED:</b> {sections["red_flags"]}
                                </div>
                                """, unsafe_allow_html=True)
                                st.markdown("<br>", unsafe_allow_html=True)
                            
                            # SOAP output
                            st.markdown(f'<div class="soap-output">{structured_note}</div>', unsafe_allow_html=True)
                            
                            # Download buttons
                            st.markdown("<br>", unsafe_allow_html=True)
                            col_d1, col_d2 = st.columns(2)
                            
                            with col_d1:
                                pdf_buffer = generate_pdf(structured_note, patient_id, st.session_state.doctor_name, patient_name)
                                st.download_button(
                                    label="📥 PDF Download",
                                    data=pdf_buffer,
                                    file_name=f"BIOS_{patient_id}.pdf",
                                    mime="application/pdf"
                                )
                            
                            with col_d2:
                                excel_key = f"excel_{st.session_state.doctor_id}_{date.today()}_bytes"
                                if excel_key in st.session_state:
                                    st.download_button(
                                        label="📊 Excel Download",
                                        data=st.session_state[excel_key],
                                        file_name=excel_filename,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                        
                        except Exception as e:
                            st.error(f"Error: {str(e)}")
            else:
                st.markdown("""
                <div style='text-align:center; padding:80px 20px; color:#334155;'>
                    <div style='font-size:48px; margin-bottom:16px;'>🧠</div>
                    <p style='font-size:14px;'>Notes type karein ya voice record karein<br>phir AI se structure karwayein</p>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown("</div>", unsafe_allow_html=True)
    
    # ── RECORDS PAGE ──
    elif "Records" in page:
        st.markdown("""
        <div class='bios-header'>
            <p class='bios-title'>📊 Today's Patient Records</p>
            <p class='bios-sub'>Daily Record — Automatically tracked</p>
        </div>
        """, unsafe_allow_html=True)
        
        my_records = [r for r in st.session_state.records if r.get("doctor_id") == st.session_state.doctor_id]
        
        if not my_records:
            st.markdown("""
            <div style='text-align:center; padding:60px; color:#475569;'>
                <div style='font-size:40px;'>📋</div>
                <p>Aaj abhi koi patient record nahi — New Consultation se shuru karein</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            # Summary stats
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f"""<div class='stat-card'>
                    <div class='stat-number'>{len(my_records)}</div>
                    <div class='stat-label'>Total Patients</div>
                </div>""", unsafe_allow_html=True)
            with c2:
                red_flags = sum(1 for r in my_records if r["red_flags"] and r["red_flags"].lower() not in ["none", "none identified", "not mentioned"])
                st.markdown(f"""<div class='stat-card'>
                    <div class='stat-number' style='color:#EF4444;'>{red_flags}</div>
                    <div class='stat-label'>Red Flags</div>
                </div>""", unsafe_allow_html=True)
            with c3:
                st.markdown(f"""<div class='stat-card'>
                    <div class='stat-number' style='color:#22C55E;'>{len(my_records) - red_flags}</div>
                    <div class='stat-label'>Normal Cases</div>
                </div>""", unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            # Records table
            for record in reversed(my_records):
                has_flag = record["red_flags"] and record["red_flags"].lower() not in ["none", "none identified", "not mentioned"]
                border_color = "#EF4444" if has_flag else "#1B3A5C"
                
                with st.expander(f"🆔 {record['patient_id']} · {record['time']} · {record['chief_complaint'][:50]}..."):
                    st.markdown(f'<div class="soap-output">{record["full_note"]}</div>', unsafe_allow_html=True)
                    
                    if has_flag:
                        st.markdown(f"<div class='alert-red'>⚠️ {record['red_flags']}</div>", unsafe_allow_html=True)
            
            # Excel download
            st.markdown("<br>", unsafe_allow_html=True)
            excel_key = f"excel_{st.session_state.doctor_id}_{date.today()}_bytes"
            if excel_key in st.session_state:
                st.download_button(
                    label=f"📥 Download Today's Excel — {get_excel_filename(st.session_state.doctor_id)}",
                    data=st.session_state[excel_key],
                    file_name=get_excel_filename(st.session_state.doctor_id),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    # ── ABOUT PAGE ──
    elif "About" in page:
        st.markdown("""
        <div class='bios-header'>
            <p class='bios-title'>ℹ️ About BIOS</p>
            <p class='bios-sub'>Biomedical Innovation & Organization System</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div class='card'>
            <p class='card-title'>Mission</p>
            <p style='color:#E2E8F0; line-height:1.8;'>
            BIOS Medical Intelligence System is built for Pakistani doctors who see 60–80 patients daily. 
            Our AI converts rough consultation notes — in Urdu or English, typed or spoken — into 
            professional structured SOAP notes instantly. We replace the cognitive load, not the doctor.
            </p>
            <br>
            <p style='color:#60A5FA; font-style:italic; font-size:15px;'>
            "Replace the Brains — Not the Arms, Not the Legs"
            </p>
        </div>
        
        <div class='card'>
            <p class='card-title'>Features</p>
            <p style='color:#E2E8F0;'>✅ &nbsp; Voice + Text input — Urdu & English</p>
            <p style='color:#E2E8F0;'>✅ &nbsp; AI-powered SOAP note generation</p>
            <p style='color:#E2E8F0;'>✅ &nbsp; Professional PDF export</p>
            <p style='color:#E2E8F0;'>✅ &nbsp; Daily Excel record system</p>
            <p style='color:#E2E8F0;'>✅ &nbsp; Per-doctor secure accounts</p>
            <p style='color:#E2E8F0;'>✅ &nbsp; Red flag detection</p>
            <p style='color:#E2E8F0;'>✅ &nbsp; Patient ID auto-generation</p>
        </div>
        
        <div class='card'>
            <p class='card-title'>Version</p>
            <p style='color:#94A3B8; font-size:13px;'>BIOS Medical Intelligence System v2.0 · Confidential · 2025</p>
        </div>
        """, unsafe_allow_html=True)

# ─── ROUTER ──────────────────────────────────────────────────────────────────
if not st.session_state.logged_in:
    show_login()
else:
    show_app()
