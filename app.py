import streamlit as st
from groq import Groq

# Page config
st.set_page_config(
    page_title="BIOS Clinical Note Structurer",
    page_icon="🏥",
    layout="centered"
)

# Styling
st.markdown("""
<style>
    .main { background-color: #F8FAFC; }
    .stTextArea textarea {
        font-size: 15px;
        border: 2px solid #1B6CA8;
        border-radius: 8px;
    }
    .stButton button {
        background-color: #1B6CA8;
        color: white;
        font-size: 16px;
        font-weight: bold;
        border-radius: 8px;
        width: 100%;
        padding: 12px;
    }
    .output-box {
        background-color: #EFF6FF;
        border-left: 4px solid #1B6CA8;
        padding: 20px;
        border-radius: 8px;
        font-size: 15px;
        line-height: 1.8;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("## 🏥 BIOS — Clinical Note Structurer")
st.markdown("**Medical Intelligence System** · Powered by BIOS AI")
st.divider()

# API Key from Streamlit secrets
client = Groq(api_key=st.secrets["GROQ_API_KEY"])

# System Prompt — BIOS ka core engine
SYSTEM_PROMPT = """You are BIOS Clinical Assistant — an expert AI for Pakistani doctors.

Your job is to take a doctor's rough, unstructured consultation notes (in Urdu, English, or mixed) and convert them into a clean, professional SOAP note format.

Always structure output as:

**CHIEF COMPLAINT**
[Main problem patient came with]

**HISTORY OF PRESENT ILLNESS**
[Detailed history — onset, duration, severity, associated symptoms]

**PAST MEDICAL HISTORY**
[Any relevant previous conditions, surgeries, medications]

**EXAMINATION FINDINGS**
[Any vitals or examination notes mentioned]

**ASSESSMENT**
[Doctor's clinical impression / working diagnosis]

**PLAN**
[Treatment, investigations, follow-up]

**RED FLAGS** ⚠️
[Only if any serious warning signs detected — otherwise write: None identified]

Rules:
- If information is missing for any section, write: "Not mentioned"
- Keep language professional but clear
- If input is in Urdu, output in English
- Always be concise and clinically accurate
- Never invent information not present in the notes"""

# Input
st.markdown("### 📝 Doctor ki Rough Notes")
raw_notes = st.text_area(
    label="Notes yahan paste karein — Urdu ya English mein",
    height=200,
    placeholder="Misaal: 35 saal ka mard, 3 din se bukhaar, sar dard, khansi nahi, khana nahi kha raha, BP 130/90..."
)

# Button
if st.button("⚡ Structure karo — AI se"):
    if raw_notes.strip() == "":
        st.warning("Pehle notes paste karein.")
    else:
        with st.spinner("BIOS AI process kar raha hai..."):
            try:
                response = client.chat.completions.create(
                    model="llama-3.3-70b-versatile",
                    messages=[
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user", "content": f"Structure these clinical notes:\n\n{raw_notes}"}
                    ],
                    temperature=0.3,
                    max_tokens=1000
                )

                structured_note = response.choices[0].message.content

                st.divider()
                st.markdown("### ✅ Structured SOAP Note")
                st.markdown(f'<div class="output-box">{structured_note}</div>',
                           unsafe_allow_html=True)

                # Download button
                st.download_button(
                    label="📄 PDF ke liye Copy karo",
                    data=structured_note,
                    file_name="BIOS_Clinical_Note.txt",
                    mime="text/plain"
                )

            except Exception as e:
                st.error(f"Error aaya: {str(e)}")

# Footer
st.divider()
st.markdown(
    "<p style='text-align:center; color:#94A3B8; font-size:13px'>"
    "BIOS — Biomedical Innovation & Organization System · Medical Intelligence System · Confidential"
    "</p>",
    unsafe_allow_html=True
)
