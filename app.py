import streamlit as st
from groq import Groq
from supabase import create_client
import pandas as pd
from datetime import datetime, date
import os, io, hashlib, uuid, tempfile
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor, white
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from audio_recorder_streamlit import audio_recorder

# ── CONFIG ───────────────────────────────────────────────────────────────────
st.set_page_config(page_title="BIOS Medical Intelligence", page_icon="🏥", layout="wide")

# ── CLIENTS ──────────────────────────────────────────────────────────────────
groq_client = Groq(api_key=st.secrets["GROQ_API_KEY"])
supabase = create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

# ── STYLING ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .stApp{background:#0D1B2A}
  section[data-testid="stSidebar"]{background:#0A1628;border-right:1px solid #1B6CA8}
  #MainMenu,footer,header{visibility:hidden}
  .bios-header{background:linear-gradient(135deg,#0D1B2A,#1B3A5C);border:1px solid #1B6CA8;border-radius:12px;padding:20px 28px;margin-bottom:24px}
  .bios-title{color:#fff;font-size:26px;font-weight:800;margin:0}
  .bios-sub{color:#60A5FA;font-size:12px;margin:4px 0 0;letter-spacing:2px;text-transform:uppercase}
  .card{background:#112240;border:1px solid #1B3A5C;border-radius:12px;padding:20px;margin-bottom:16px}
  .card-title{color:#60A5FA;font-size:12px;font-weight:600;letter-spacing:1px;text-transform:uppercase;margin-bottom:12px}
  .stTextArea textarea{background:#0A1628!important;color:#E2E8F0!important;border:1px solid #1B6CA8!important;border-radius:8px!important}
  .stTextInput input{background:#0A1628!important;color:#E2E8F0!important;border:1px solid #1B6CA8!important;border-radius:8px!important}
  .stButton button{background:linear-gradient(135deg,#1B6CA8,#2563EB)!important;color:#fff!important;border:none!important;border-radius:8px!important;font-weight:600!important;width:100%!important}
  .soap-output{background:#112240;border:1px solid #1B6CA8;border-left:4px solid #3B82F6;border-radius:8px;padding:24px;color:#E2E8F0;font-size:14px;line-height:1.8;white-space:pre-wrap}
  .stat-card{background:#112240;border:1px solid #1B3A5C;border-radius:10px;padding:16px;text-align:center}
  .stat-number{color:#3B82F6;font-size:28px;font-weight:800}
  .stat-label{color:#94A3B8;font-size:11px;text-transform:uppercase;letter-spacing:1px}
  .badge{background:#1E3A5F;border:1px solid #3B82F6;border-radius:20px;padding:4px 14px;color:#60A5FA;font-size:12px;font-weight:600;display:inline-block}
  .alert-red{background:#1F0A0A;border-left:4px solid #EF4444;border-radius:6px;padding:12px 16px;color:#FCA5A5;font-size:13px}
  .alert-green{background:#0A1F0A;border-left:4px solid #22C55E;border-radius:6px;padding:12px 16px;color:#86EFAC;font-size:13px}
  .alert-yellow{background:#1F1A0A;border-left:4px solid #F59E0B;border-radius:6px;padding:12px 16px;color:#FCD34D;font-size:13px}
  label{color:#94A3B8!important}
  .plan-free{color:#94A3B8;font-size:11px;background:#1B2A1B;padding:2px 8px;border-radius:10px}
  .plan-basic{color:#60A5FA;font-size:11px;background:#1B2A3F;padding:2px 8px;border-radius:10px}
  .plan-pro{color:#F59E0B;font-size:11px;background:#2A1F0A;padding:2px 8px;border-radius:10px}
</style>
""", unsafe_allow_html=True)

# ── SYSTEM PROMPT ─────────────────────────────────────────────────────────────
SYSTEM_PROMPT = """You are BIOS Clinical Assistant — an expert AI medical scribe for Pakistani doctors.

Your task: Convert rough doctor notes (Urdu, English, or mixed) into a clean professional SOAP note in English.

=== URDU MEDICAL DICTIONARY ===
Body Parts: bazu/baazu=ARM, pair/paon=LEG, ghutna=KNEE, kamar=LOWER BACK, seena=CHEST, sar/sir=HEAD, gardan=NECK, pet=ABDOMEN, ankh=EYE, kaan=EAR, haath=HAND, ungali=FINGER, kaandha=SHOULDER, peet=UPPER BACK
Symptoms: dard=PAIN, bukhaar=FEVER, khansi=COUGH, saans phoolna=SHORTNESS OF BREATH, ulti=VOMITING, dast=DIARRHEA, chakkar=DIZZINESS, kamzori=WEAKNESS, soojan=SWELLING, khujli=ITCHING, jalan=BURNING, thakan=FATIGUE, sar dard=HEADACHE, paseena=SWEATING
Conditions: tota/tooti=FRACTURE, zakhm=WOUND, sugar=DIABETES, haddi kamzor=BONE WEAKNESS/OSTEOPOROSIS, joron ka dard=JOINT PAIN, saans ki bimari=RESPIRATORY DISEASE, dil ki bimari=CARDIAC DISEASE
Urgency: jaldi/foran=URGENT, zyada dard=SEVERE PAIN, emergency=EMERGENCY

=== OUTPUT FORMAT ===
CHIEF COMPLAINT
[Age, name if given, main problem — one clear sentence]

HISTORY OF PRESENT ILLNESS
[Onset, duration, severity, associated symptoms]

PAST MEDICAL HISTORY
[Previous conditions, surgeries, medications, allergies]

EXAMINATION FINDINGS
[Vitals if mentioned, physical findings]

ASSESSMENT
[Working diagnosis — specific medical terminology]

PLAN
[Investigations, treatment, medications, follow-up — mark URGENT if needed]

RED FLAGS ⚠️
[Serious warning signs — write NONE if not present]

=== RULES ===
1. Missing info: write exactly "Not mentioned"
2. Never invent information
3. Translate Urdu accurately using dictionary above
4. Specify LEFT or RIGHT if mentioned
5. Age in Chief Complaint always
6. Urgency words → URGENT in PLAN
7. Output always in English
8. Red flags: chest pain, stroke symptoms, breathing difficulty, high fever in child, fracture in child, uncontrolled bleeding"""

# ── HELPERS ──────────────────────────────────────────────────────────────────
def hash_password(password):
    return password

def generate_patient_id():
    now = datetime.now()
    uid = str(uuid.uuid4())[:4].upper()
    return f"BIOS-{now.strftime('%Y%m%d')}-{uid}"

def get_plan_limits(plan):
    limits = {"free": 5, "basic": 50, "pro": 999999}
    return limits.get(plan, 5)

def check_usage_limit(doctor_id, plan):
    today = date.today().isoformat()
    result = supabase.table("usage_log").select("id", count="exact").eq("doctor_id", doctor_id).eq("date", today).execute()
    used = result.count or 0
    limit = get_plan_limits(plan)
    return used, limit

def log_usage(doctor_id):
    supabase.table("usage_log").insert({
        "doctor_id": doctor_id,
        "action": "soap_note",
        "date": date.today().isoformat()
    }).execute()

def extract_sections(note):
    s = {"chief_complaint": "", "assessment": "", "plan": "", "red_flags": "NONE"}
    current = None
    for line in note.split('\n'):
        line = line.strip()
        if not line: continue
        if 'CHIEF COMPLAINT' in line.upper(): current = 'chief_complaint'
        elif 'ASSESSMENT' in line.upper(): current = 'assessment'
        elif 'PLAN' in line.upper(): current = 'plan'
        elif 'RED FLAG' in line.upper(): current = 'red_flags'
        elif current:
            clean = line.replace('**','').replace('##','').strip()
            if clean: s[current] += clean + " "
    return {k: v.strip() for k, v in s.items()}

def transcribe_audio(audio_bytes):
    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmp:
        tmp.write(audio_bytes)
        tmp_path = tmp.name
    with open(tmp_path, "rb") as f:
        result = groq_client.audio.transcriptions.create(
            model="whisper-large-v3", file=f, response_format="text")
    os.unlink(tmp_path)
    return result

def generate_pdf(note, patient_id, doctor_name, patient_name=""):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
        rightMargin=2*cm, leftMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm)
    story = []
    navy = HexColor('#0D1B2A')
    blue = HexColor('#1B6CA8')
    lb   = HexColor('#EFF6FF')
    red  = HexColor('#EF4444')
    gray = HexColor('#64748B')

    # Header
    hdr = Table([[Paragraph("<font color='white'><b>🏥 BIOS Medical Intelligence System</b></font>",
        ParagraphStyle('H', fontSize=16, fontName='Helvetica-Bold', alignment=TA_CENTER)),
        Paragraph("<font color='#93C5FD'>AI-Powered Clinical Documentation</font>",
        ParagraphStyle('S', fontSize=9, fontName='Helvetica', alignment=TA_CENTER))
    ]], colWidths=[17*cm])
    hdr.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),navy),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('TOPPADDING',(0,0),(-1,-1),14),
        ('BOTTOMPADDING',(0,0),(-1,-1),14),
    ]))
    story.append(hdr)
    story.append(Spacer(1,0.3*cm))

    now = datetime.now().strftime("%d %B %Y · %I:%M %p")
    meta = Table([
        [Paragraph(f"<b>Patient ID:</b> {patient_id}", ParagraphStyle('M', fontSize=9, fontName='Helvetica')),
         Paragraph(f"<b>Date:</b> {now}", ParagraphStyle('M', fontSize=9, fontName='Helvetica'))],
        [Paragraph(f"<b>Doctor:</b> {doctor_name}", ParagraphStyle('M', fontSize=9, fontName='Helvetica')),
         Paragraph(f"<b>Patient:</b> {patient_name or 'Not specified'}", ParagraphStyle('M', fontSize=9, fontName='Helvetica'))],
    ], colWidths=[8.5*cm, 8.5*cm])
    meta.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),lb),
        ('TOPPADDING',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7),
        ('LEFTPADDING',(0,0),(-1,-1),12),('RIGHTPADDING',(0,0),(-1,-1),12),
        ('GRID',(0,0),(-1,-1),0.5,HexColor('#BFDBFE')),
    ]))
    story.append(meta)
    story.append(Spacer(1,0.3*cm))
    story.append(HRFlowable(width="100%",thickness=2,color=blue))
    story.append(Spacer(1,0.3*cm))

    sections = [("CHIEF COMPLAINT",""),("HISTORY OF PRESENT ILLNESS",""),
                ("PAST MEDICAL HISTORY",""),("EXAMINATION FINDINGS",""),
                ("ASSESSMENT",""),("PLAN",""),("RED FLAGS ⚠️","")]
    lines = note.split('\n')
    cur = None
    sec_content = {}
    for line in lines:
        line = line.strip()
        if not line: continue
        matched = False
        for sn, _ in sections:
            if sn.replace(" ⚠️","") in line.upper():
                cur = sn; sec_content[cur] = []; matched = True; break
        if not matched and cur:
            sec_content[cur].append(line)

    for sn, _ in sections:
        content = ' '.join(sec_content.get(sn, [])).strip() or "Not mentioned"
        is_rf = "RED FLAG" in sn
        bg = HexColor('#FEF2F2') if is_rf else lb
        hc = red if is_rf else navy
        t = Table([
            [Paragraph(f'<font color="white"><b>{sn}</b></font>',
             ParagraphStyle('SH', fontSize=10, fontName='Helvetica-Bold', textColor=white))],
            [Paragraph(content, ParagraphStyle('SC', fontSize=10, fontName='Helvetica',
             textColor=HexColor('#1E293B'), leading=15))]
        ], colWidths=[17*cm])
        t.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0),hc),('BACKGROUND',(0,1),(-1,1),bg),
            ('TOPPADDING',(0,0),(-1,0),7),('BOTTOMPADDING',(0,0),(-1,0),7),
            ('TOPPADDING',(0,1),(-1,1),10),('BOTTOMPADDING',(0,1),(-1,1),10),
            ('LEFTPADDING',(0,0),(-1,-1),12),('RIGHTPADDING',(0,0),(-1,-1),12),
        ]))
        story.append(t)
        story.append(Spacer(1,0.2*cm))

    story.append(Spacer(1,0.3*cm))
    story.append(HRFlowable(width="100%",thickness=1,color=HexColor('#CBD5E1')))
    story.append(Paragraph(
        "Generated by BIOS Medical Intelligence System · Confidential · For authorized use only",
        ParagraphStyle('F', fontSize=8, textColor=gray, fontName='Helvetica', alignment=TA_CENTER)))
    doc.build(story)
    buf.seek(0)
    return buf

# ── SESSION STATE ─────────────────────────────────────────────────────────────
for key, val in [("logged_in",False),("doctor",None),("page","consult")]:
    if key not in st.session_state:
        st.session_state[key] = val

# ═══════════════════════════════════════════════════════════════════════════════
# LOGIN PAGE
# ═══════════════════════════════════════════════════════════════════════════════
def show_login():
    _, col, _ = st.columns([1,2,1])
    with col:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("""
        <div style='text-align:center;margin-bottom:32px'>
          <div style='font-size:52px'>🏥</div>
          <h1 style='color:#fff;font-size:34px;font-weight:800;margin:8px 0'>BIOS</h1>
          <p style='color:#60A5FA;font-size:12px;letter-spacing:3px;text-transform:uppercase'>Medical Intelligence System</p>
        </div>""", unsafe_allow_html=True)

        tab1, tab2 = st.tabs(["🔐 Login", "📝 Register"])

        with tab1:
            st.markdown("<div class='card'>", unsafe_allow_html=True)
            email = st.text_input("Email", placeholder="doctor@clinic.com", key="login_email")
            password = st.text_input("Password", type="password", key="login_pass")
            if st.button("Login to BIOS", key="login_btn"):
                if email and password:
                    result = supabase.table("doctors").select("*").eq("email", email.lower()).execute()
                    if result.data:
                        doc = result.data[0]
                        if doc["password_hash"] == password:
                            if not doc["is_active"]:
                                st.markdown("<div class='alert-red'>❌ Account suspended.</div>", unsafe_allow_html=True)
                            else:
                                supabase.table("doctors").update({"last_login": datetime.now().isoformat()}).eq("id", doc["id"]).execute()
                                st.session_state.logged_in = True
                                st.session_state.doctor = doc
                                st.rerun()
                        else:
                            st.markdown("<div class='alert-red'>❌ Wrong password.</div>", unsafe_allow_html=True)
                    else:
                        st.markdown("<div class='alert-red'>❌ Email not found.</div>", unsafe_allow_html=True)
                else:
                    st.warning("Email aur password bharein.")
            st.markdown("</div>", unsafe_allow_html=True)

        with tab2:
            st.markdown("<div class='card'>", unsafe_allow_html=True)
            st.markdown("<p class='card-title'>Create Your BIOS Account</p>", unsafe_allow_html=True)
            r_name      = st.text_input("Full Name", placeholder="Dr. Ahmed Khan")
            r_email     = st.text_input("Email", placeholder="ahmed@clinic.com", key="reg_email")
            r_phone     = st.text_input("Phone", placeholder="03XX-XXXXXXX")
            r_specialty = st.selectbox("Specialty", ["General Physician","Internal Medicine",
                          "Paediatrician","Gynaecologist","Surgeon","Cardiologist",
                          "Orthopaedic","ENT","Dermatologist","Other"])
            r_clinic    = st.text_input("Clinic Name", placeholder="Al-Shifa Clinic, Lahore")
            r_pass      = st.text_input("Password", type="password", key="reg_pass")
            r_pass2     = st.text_input("Confirm Password", type="password", key="reg_pass2")
            if st.button("Create Account", key="reg_btn"):
                if not all([r_name, r_email, r_phone, r_clinic, r_pass, r_pass2]):
                    st.warning("Sab fields bharein.")
                elif r_pass != r_pass2:
                    st.markdown("<div class='alert-red'>❌ Passwords match nahi karte.</div>", unsafe_allow_html=True)
                elif len(r_pass) < 6:
                    st.markdown("<div class='alert-red'>❌ Password 6 characters ka hona chahiye.</div>", unsafe_allow_html=True)
                else:
                    existing = supabase.table("doctors").select("id").eq("email", r_email.lower()).execute()
                    if existing.data:
                        st.markdown("<div class='alert-red'>❌ Email already registered.</div>", unsafe_allow_html=True)
                    else:
                        from datetime import timedelta
                        trial_expiry = (date.today() + timedelta(days=7)).isoformat()
                        supabase.table("doctors").insert({
                            "name": r_name,
                            "email": r_email.lower(),
                            "password_hash": r_pass,
                            "specialty": r_specialty,
                            "clinic_name": r_clinic,
                            "phone": r_phone,
                            "plan": "free",
                            "plan_expiry": trial_expiry,
                            "is_active": True
                        }).execute()
                        st.markdown("<div class='alert-green'>✅ Account created! Login karein.</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<p style='text-align:center;color:#334155;font-size:11px;margin-top:16px'>BIOS Medical Intelligence System · Confidential · 2025</p>", unsafe_allow_html=True)
# ═══════════════════════════════════════════════════════════════════════════════
# ADMIN PAGE
# ═══════════════════════════════════════════════════════════════════════════════
def show_admin():
    st.markdown("<div class='bios-header'><p class='bios-title'>⚙️ BIOS Admin Panel</p><p class='bios-sub'>Doctor Management · Usage · Payments</p></div>", unsafe_allow_html=True)

    all_docs = supabase.table("doctors").select("*").execute().data or []

    c1,c2,c3,c4 = st.columns(4)
    with c1: st.markdown(f"<div class='stat-card'><div class='stat-number'>{len(all_docs)}</div><div class='stat-label'>Total Doctors</div></div>", unsafe_allow_html=True)
    with c2:
        active = sum(1 for d in all_docs if d["is_active"])
        st.markdown(f"<div class='stat-card'><div class='stat-number' style='color:#22C55E'>{active}</div><div class='stat-label'>Active</div></div>", unsafe_allow_html=True)
    with c3:
        paid = sum(1 for d in all_docs if d["plan"] in ["basic","pro"])
        st.markdown(f"<div class='stat-card'><div class='stat-number' style='color:#F59E0B'>{paid}</div><div class='stat-label'>Paid Plans</div></div>", unsafe_allow_html=True)
    with c4:
        revenue = sum(5000 for d in all_docs if d["plan"]=="basic") + sum(15000 for d in all_docs if d["plan"]=="pro")
        st.markdown(f"<div class='stat-card'><div class='stat-number' style='color:#3B82F6'>Rs {revenue:,}</div><div class='stat-label'>Monthly Revenue</div></div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 👨‍⚕️ All Doctors")

    for doc in all_docs:
        used, limit = check_usage_limit(doc["id"], doc["plan"])
        with st.expander(f"{'🟢' if doc['is_active'] else '🔴'} {doc['name']} · {doc['specialty']} · {doc['clinic_name']}"):
            c1,c2,c3 = st.columns(3)
            with c1:
                st.write(f"**Email:** {doc['email']}")
                st.write(f"**Phone:** {doc.get('phone','—')}")
                st.write(f"**Joined:** {doc['created_at'][:10]}")
            with c2:
                st.write(f"**Plan:** {doc['plan'].upper()}")
                st.write(f"**Expiry:** {doc.get('plan_expiry','—')}")
                st.write(f"**Today:** {used}/{limit} patients")
            with c3:
                new_plan = st.selectbox("Change Plan", ["free","basic","pro"],
                    index=["free","basic","pro"].index(doc["plan"]),
                    key=f"plan_{doc['id']}")
                new_expiry = st.date_input("Plan Expiry", key=f"exp_{doc['id']}")
                if st.button("Update", key=f"upd_{doc['id']}"):
                    supabase.table("doctors").update({
                        "plan": new_plan,
                        "plan_expiry": new_expiry.isoformat()
                    }).eq("id", doc["id"]).execute()
                    st.success("Updated!")
                    st.rerun()

            col_a, col_b = st.columns(2)
            with col_a:
                status_label = "🔴 Suspend" if doc["is_active"] else "🟢 Activate"
                if st.button(status_label, key=f"tog_{doc['id']}"):
                    supabase.table("doctors").update({"is_active": not doc["is_active"]}).eq("id", doc["id"]).execute()
                    st.rerun()
            with col_b:
                if st.button("📋 Add Payment", key=f"pay_{doc['id']}"):
                    st.session_state[f"pay_modal_{doc['id']}"] = True

            if st.session_state.get(f"pay_modal_{doc['id']}"):
                amt = st.number_input("Amount (Rs)", key=f"amt_{doc['id']}")
                method = st.selectbox("Method", ["JazzCash","Easypaisa","Bank Transfer"], key=f"mth_{doc['id']}")
                months = st.number_input("Months", value=1, min_value=1, key=f"mo_{doc['id']}")
                if st.button("Save Payment", key=f"savepay_{doc['id']}"):
                    supabase.table("payments").insert({
                        "doctor_id": doc["id"],
                        "amount": int(amt),
                        "method": method,
                        "plan": doc["plan"],
                        "months": int(months),
                        "verified": True
                    }).execute()
                    st.success("Payment recorded!")
                    st.session_state[f"pay_modal_{doc['id']}"] = False

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN APP
# ═══════════════════════════════════════════════════════════════════════════════
def show_app():
    doc = st.session_state.doctor
    plan_colors = {"free":"plan-free","basic":"plan-basic","pro":"plan-pro"}

    with st.sidebar:
        st.markdown(f"""
        <div style='padding:16px 0 20px;border-bottom:1px solid #1B3A5C;margin-bottom:16px'>
          <div style='color:#94A3B8;font-size:11px;text-transform:uppercase;letter-spacing:1px'>Logged in as</div>
          <div style='color:#E2E8F0;font-size:15px;font-weight:600;margin:4px 0'>{doc['name']}</div>
          <div style='color:#3B82F6;font-size:12px'>{doc['specialty']}</div>
          <div style='color:#64748B;font-size:11px'>{doc['clinic_name']}</div>
          <span class='{plan_colors.get(doc["plan"],"plan-free")}'>{doc["plan"].upper()} PLAN</span>
        </div>""", unsafe_allow_html=True)

        page = st.selectbox("", ["📝 New Consultation","📊 My Records","ℹ️ About"], label_visibility="collapsed")

        used, limit = check_usage_limit(doc["id"], doc["plan"])
        pct = int((used/limit)*100) if limit < 999999 else 0
        st.markdown(f"""
        <div style='margin-top:16px'>
          <div style='color:#94A3B8;font-size:11px;margin-bottom:6px'>Today's Usage</div>
          <div style='background:#1B3A5C;border-radius:6px;height:8px'>
            <div style='background:#3B82F6;width:{min(pct,100)}%;height:8px;border-radius:6px'></div>
          </div>
          <div style='color:#60A5FA;font-size:12px;margin-top:4px'>{used} / {"∞" if limit==999999 else limit} patients</div>
        </div>""", unsafe_allow_html=True)

        if doc["plan"] == "free":
            st.markdown("""
            <div class='alert-yellow' style='margin-top:16px;font-size:12px'>
            ⭐ <b>Upgrade to Basic</b><br>
            Rs 5,000/month · 50 patients/day<br>
            JazzCash: 0300-0000000
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚪 Logout"):
            st.session_state.logged_in = False
            st.session_state.doctor = None
            st.rerun()

    # ── CONSULTATION ──
    if "Consultation" in page:
        st.markdown("""
        <div class='bios-header'>
          <p class='bios-title'>🏥 BIOS Clinical Note Structurer</p>
          <p class='bios-sub'>Medical Intelligence System · AI-Powered SOAP Notes</p>
        </div>""", unsafe_allow_html=True)

        used, limit = check_usage_limit(doc["id"], doc["plan"])
        if used >= limit and limit != 999999:
            st.markdown(f"""
            <div class='alert-red'>
            ❌ <b>Daily limit reached ({used}/{limit} patients)</b><br>
            Upgrade karo ya kal dobara aao. JazzCash: 0300-0000000
            </div>""", unsafe_allow_html=True)
            return

        col_l, col_r = st.columns([1,1], gap="large")

        with col_l:
            st.markdown("<div class='card'>", unsafe_allow_html=True)
            st.markdown("<p class='card-title'>📋 Patient Info</p>", unsafe_allow_html=True)
            patient_name = st.text_input("Patient Name (Optional)", placeholder="Ahmed Khan")
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("<div class='card'>", unsafe_allow_html=True)
            st.markdown("<p class='card-title'>🎙️ Input Method</p>", unsafe_allow_html=True)
            method = st.radio("", ["✍️ Type Notes","🎙️ Voice Input"], horizontal=True, label_visibility="collapsed")
            raw_notes = ""

            if "Type" in method:
                raw_notes = st.text_area("", height=200,
                    placeholder="Urdu ya English mein notes likhein...\nMisaal: 45 saal ki khatoon, seena dard subah se, baen haath mein phailta hai...",
                    label_visibility="collapsed")
            else:
                st.markdown("<p style='color:#94A3B8;font-size:13px'>Record button dabao aur bolna shuru karo</p>", unsafe_allow_html=True)
                audio = audio_recorder(text="Record", recording_color="#EF4444", neutral_color="#1B6CA8", icon_size="2x")
                if audio:
                    st.audio(audio, format="audio/wav")
                    with st.spinner("Transcribe ho raha hai..."):
                        try:
                            raw_notes = transcribe_audio(audio)
                            st.markdown(f"<div class='alert-green'>✅ <i>\"{raw_notes[:100]}...\"</i></div>", unsafe_allow_html=True)
                        except Exception as e:
                            st.error(f"Voice error: {e}")

            st.markdown("</div>", unsafe_allow_html=True)
            process = st.button("⚡ BIOS AI se Structure Karo")

        with col_r:
            st.markdown("<div class='card' style='min-height:520px'>", unsafe_allow_html=True)
            st.markdown("<p class='card-title'>📄 Structured SOAP Note</p>", unsafe_allow_html=True)

            if process:
                if not raw_notes.strip():
                    st.markdown("<div class='alert-red'>⚠️ Notes likhein ya voice record karein.</div>", unsafe_allow_html=True)
                else:
                    with st.spinner("🧠 BIOS AI processing..."):
                        try:
                            resp = groq_client.chat.completions.create(
                                model="llama-3.1-8b-instant",
                                messages=[
                                    {"role":"system","content":SYSTEM_PROMPT},
                                    {"role":"user","content":f"Structure these clinical notes:\n\n{raw_notes}"}
                                ],
                                temperature=0.2, max_tokens=1200
                            )
                            note = resp.choices[0].message.content
                            patient_id = generate_patient_id()
                            sections = extract_sections(note)
                            log_usage(doc["id"])

                            # Save to Supabase
                            supabase.table("patient_records").insert({
                                "patient_id": patient_id,
                                "doctor_id": doc["id"],
                                "patient_name": patient_name,
                                "raw_notes": raw_notes,
                                "structured_note": note,
                                "chief_complaint": sections["chief_complaint"],
                                "assessment": sections["assessment"],
                                "plan": sections["plan"],
                                "red_flags": sections["red_flags"],
                                "input_method": "voice" if "Voice" in method else "text",
                                "record_date": date.today().isoformat()
                            }).execute()

                            st.markdown(f"<div style='margin-bottom:12px'><span class='badge'>🆔 {patient_id}</span></div>", unsafe_allow_html=True)

                            rf = sections["red_flags"]
                            if rf and rf.upper() not in ["NONE","NONE IDENTIFIED","NOT MENTIONED"]:
                                st.markdown(f"<div class='alert-red'>⚠️ <b>RED FLAG:</b> {rf}</div><br>", unsafe_allow_html=True)

                            st.markdown(f'<div class="soap-output">{note}</div>', unsafe_allow_html=True)
                            st.markdown("<br>", unsafe_allow_html=True)

                            c1, c2 = st.columns(2)
                            with c1:
                                pdf = generate_pdf(note, patient_id, doc["name"], patient_name)
                                st.download_button("📥 PDF Download", data=pdf,
                                    file_name=f"BIOS_{patient_id}.pdf", mime="application/pdf")
                            with c2:
                                st.download_button("📋 Text Copy", data=note,
                                    file_name=f"BIOS_{patient_id}.txt", mime="text/plain")

                        except Exception as e:
                            st.error(f"Error: {e}")
            else:
                st.markdown("""
                <div style='text-align:center;padding:80px 20px;color:#334155'>
                  <div style='font-size:48px'>🧠</div>
                  <p style='font-size:13px'>Notes type karein ya voice record karein<br>phir AI se structure karwayein</p>
                </div>""", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

    # ── RECORDS ──
    elif "Records" in page:
        st.markdown("""
        <div class='bios-header'>
          <p class='bios-title'>📊 My Patient Records</p>
          <p class='bios-sub'>Permanently saved — All your consultations</p>
        </div>""", unsafe_allow_html=True)

        records = supabase.table("patient_records").select("*").eq("doctor_id", doc["id"]).order("created_at", desc=True).execute().data or []

        today_recs = [r for r in records if r["record_date"] == date.today().isoformat()]
        red_flags  = [r for r in today_recs if r["red_flags"] and r["red_flags"].upper() not in ["NONE","NONE IDENTIFIED","NOT MENTIONED"]]

        c1,c2,c3,c4 = st.columns(4)
        with c1: st.markdown(f"<div class='stat-card'><div class='stat-number'>{len(today_recs)}</div><div class='stat-label'>Today</div></div>", unsafe_allow_html=True)
        with c2: st.markdown(f"<div class='stat-card'><div class='stat-number'>{len(records)}</div><div class='stat-label'>All Time</div></div>", unsafe_allow_html=True)
        with c3: st.markdown(f"<div class='stat-card'><div class='stat-number' style='color:#EF4444'>{len(red_flags)}</div><div class='stat-label'>Red Flags Today</div></div>", unsafe_allow_html=True)
        with c4: st.markdown(f"<div class='stat-card'><div class='stat-number' style='color:#22C55E'>{len(today_recs)-len(red_flags)}</div><div class='stat-label'>Normal Today</div></div>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Excel export
        if records:
            df = pd.DataFrame([{
                "Patient ID": r["patient_id"],
                "Date": r["record_date"],
                "Time": r["created_at"][11:16],
                "Patient Name": r.get("patient_name",""),
                "Chief Complaint": r["chief_complaint"],
                "Assessment": r["assessment"],
                "Plan": r["plan"],
                "Red Flags": r["red_flags"],
            } for r in records])

            excel_buf = io.BytesIO()
            with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="BIOS Records")
            excel_buf.seek(0)

            today_str = date.today().strftime("%Y-%m-%d")
            doc_id_short = st.session_state.doctor_id if hasattr(st.session_state, 'doctor_id') else "doctor"
            st.download_button(
                f"📥 Download Excel — BIOS_{doc['name'].replace(' ','_')}_{today_str}.xlsx",
                data=excel_buf.getvalue(),
                file_name=f"BIOS_{doc['name'].replace(' ','_')}_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.markdown("<br>", unsafe_allow_html=True)
        for r in records:
            has_rf = r["red_flags"] and r["red_flags"].upper() not in ["NONE","NONE IDENTIFIED","NOT MENTIONED"]
            icon = "🔴" if has_rf else "🟢"
            with st.expander(f"{icon} {r['patient_id']} · {r['record_date']} {r['created_at'][11:16]} · {r['chief_complaint'][:60]}"):
                st.markdown(f'<div class="soap-output">{r["structured_note"]}</div>', unsafe_allow_html=True)
                if has_rf:
                    st.markdown(f"<div class='alert-red'>⚠️ {r['red_flags']}</div>", unsafe_allow_html=True)
                pdf = generate_pdf(r["structured_note"], r["patient_id"], doc["name"], r.get("patient_name",""))
                st.download_button("📥 PDF", data=pdf,
                    file_name=f"BIOS_{r['patient_id']}.pdf",
                    mime="application/pdf", key=f"pdf_{r['id']}")

    elif "About" in page:
        st.markdown("""
        <div class='bios-header'>
          <p class='bios-title'>ℹ️ About BIOS</p>
          <p class='bios-sub'>Biomedical Innovation & Organization System</p>
        </div>
        <div class='card'>
          <p class='card-title'>Mission</p>
          <p style='color:#E2E8F0;line-height:1.8'>BIOS Medical Intelligence System is built for Pakistani doctors who see 60–80 patients daily. AI converts rough notes — Urdu or English, typed or spoken — into professional SOAP notes instantly.</p>
          <p style='color:#60A5FA;font-style:italic;margin-top:12px'>"Replace the Brains — Not the Arms, Not the Legs"</p>
        </div>
        <div class='card'>
          <p class='card-title'>Pricing</p>
          <p style='color:#94A3B8'>🆓 Free Trial — 7 days · 5 patients/day</p>
          <p style='color:#60A5FA'>⭐ Basic — Rs 5,000/month · 50 patients/day</p>
          <p style='color:#F59E0B'>👑 Pro — Rs 15,000/month · Unlimited patients</p>
          <p style='color:#64748B;font-size:12px;margin-top:8px'>Payment: JazzCash / Easypaisa / Bank Transfer</p>
        </div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ROUTER
# ═══════════════════════════════════════════════════════════════════════════════
if not st.session_state.logged_in:
    show_login()
else:
    # Admin check
    if st.session_state.doctor and st.session_state.doctor.get("email") == "hq@bios.pk":
        show_admin()
    else:
        show_app()
