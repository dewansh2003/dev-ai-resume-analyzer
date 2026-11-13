import os
import re
import time
from io import BytesIO

import streamlit as st
from docx import Document
from pypdf import PdfReader

st.set_page_config(page_title="Dev AI | Resume Analyzer", page_icon="📄", layout="centered")

# ---------------- Config ----------------
SKILLS = [
    "python", "javascript", "react", "redux", "node.js", "express",
    "sql", "mysql", "mongodb", "machine learning", "generative ai",
    "pandas", "numpy", "excel", "power bi", "html", "css",
    "rest api", "rest apis", "api integration", "git", "github",
    "flask", "django", "data structures", "algorithms",
    "problem solving", "web development",
]

GEMINI_MODELS = ["gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-2.0-flash"]


# ---------------- Helpers ----------------
def get_api_key():
    try:
        return st.secrets["GEMINI_API_KEY"]
    except Exception:
        return os.getenv("GEMINI_API_KEY")


def ocr_pdf(data: bytes) -> str:
    try:
        from pdf2image import convert_from_bytes
        import pytesseract

        pages = convert_from_bytes(data)
        return "\n".join(pytesseract.image_to_string(p) for p in pages)
    except Exception:
        return ""


def extract_text(uploaded_file, use_ocr=False) -> str:
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()

    if name.endswith(".pdf"):
        reader = PdfReader(BytesIO(data))
        text = "\n".join((page.extract_text() or "") for page in reader.pages)
        if not text.strip() and use_ocr:
            text = ocr_pdf(data)
        return text

    if name.endswith(".docx"):
        doc = Document(BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs)

    if name.endswith(".txt"):
        return data.decode("utf-8", errors="ignore")

    return ""


def has_skill(skill: str, text: str) -> bool:
    pattern = r"(?<![a-z0-9])" + re.escape(skill) + r"(?![a-z0-9])"
    return re.search(pattern, text) is not None


def get_gemini_feedback(resume_text: str, job_text: str):
    """Returns (feedback_text, error_text). Retries and falls back across models."""
    api_key = get_api_key()
    if not api_key:
        return None, "GEMINI_API_KEY not found in Streamlit Secrets."

    try:
        from google import genai
    except Exception as e:
        return None, f"google-genai package not installed: {e}"

    client = genai.Client(api_key=api_key)
    prompt = (
        "You are an expert resume reviewer. Compare the resume with the job description.\n"
        "Give: 1) a short summary, 2) top 3 strengths, 3) top 3 gaps, "
        "4) 5 concrete suggestions to improve the resume for this job. "
        "Be honest, do not invent experience. Keep it concise.\n\n"
        f"RESUME:\n{resume_text[:8000]}\n\nJOB DESCRIPTION:\n{job_text[:4000]}"
    )

    last_error = "Unknown error"
    for model in GEMINI_MODELS:
        for attempt in range(3):
            try:
                resp = client.models.generate_content(model=model, contents=prompt)
                if resp.text:
                    return resp.text, None
                last_error = "Empty response from Gemini."
                break
            except Exception as e:
                last_error = str(e)
                if any(code in last_error for code in ("503", "429", "UNAVAILABLE")):
                    time.sleep(2 * (attempt + 1))
                    continue
                break
    return None, last_error


def build_recommendations(missing, score):
    recs = []
    if missing:
        recs.append(
            "Add these skills to your resume only if you genuinely have them: "
            + ", ".join(missing) + "."
        )
    if score < 50:
        recs.append("Tailor your resume headline and summary to match this job description.")
    recs.append("Where truthful, quantify achievements with numbers, outcomes or measurable results.")
    return recs


# ---------------- Header ----------------
col_logo, col_title = st.columns([1, 4])
with col_logo:
    if os.path.exists("logo.jpg"):
        st.image("logo.jpg", width=90)
with col_title:
    st.title("Dev AI")
    st.caption("RESUME ANALYZER")
    st.write("Analyze your resume, compare job skills, and get personalized AI feedback.")

st.header("Analyze Your Resume")

left, right = st.columns(2)
with left:
    resume_file = st.file_uploader("Upload Resume", type=["pdf", "docx", "txt"])
with right:
    job_description = st.text_area("Paste Job Description", height=150)

use_gemini = st.checkbox("Enable Gemini AI feedback", value=True)
use_ocr = st.checkbox("Try OCR for scanned PDFs", value=True)
st.caption(
    "When Gemini feedback is enabled, resume text is sent to Google Gemini. "
    "OCR requires Tesseract and Poppler."
)

# ---------------- Analysis ----------------
if st.button("Analyze Resume", use_container_width=True):
    if resume_file is None or not job_description.strip():
        st.warning("Please upload your resume and paste the job description.")
    else:
        with st.spinner("Analyzing..."):
            resume_text = extract_text(resume_file, use_ocr=use_ocr)

            if not resume_text.strip():
                st.error("Could not extract text. Please check your resume file.")
                st.stop()

            resume_lower = resume_text.lower()
            job_lower = job_description.lower()

            required = [s for s in SKILLS if has_skill(s, job_lower)]
            matched = [s for s in required if has_skill(s, resume_lower)]
            missing = [s for s in required if s not in matched]
            detected = [s for s in SKILLS if has_skill(s, resume_lower)]
            score = (len(matched) / len(required) * 100) if required else 0.0

            feedback, feedback_error = (None, None)
            if use_gemini:
                feedback, feedback_error = get_gemini_feedback(resume_text, job_description)

        st.subheader("📊 Analysis Results")
        st.metric("Job Match Score", f"{score:.1f}%")
        st.write(f"**Required skills detected:** {len(required)}")
        st.write(f"**Matching skills:** {len(matched)}")

        st.success("Matched Skills: " + (", ".join(matched) if matched else "None detected"))
        st.warning("Missing Skills: " + (", ".join(missing) if missing else "None detected"))
        st.info("This score measures keyword overlap, not a complete ATS score.")

        st.subheader("Recommendations")
        for i, rec in enumerate(build_recommendations(missing, score), start=1):
            st.write(f"{i}. {rec}")

        if use_gemini:
            st.subheader("Gemini AI Feedback")
            if feedback:
                st.write(feedback)
            else:
                st.warning("Gemini AI feedback could not be generated. The rule-based analysis is still available.")
                with st.expander("View technical details"):
                    st.code(feedback_error or "Unknown error")

        st.subheader("Skills Detected in Resume")
        if detected:
            chips = " ".join(
                f'<span style="display:inline-block;margin:3px;padding:4px 10px;'
                f'border:1px solid #5eead4;border-radius:8px;">{s}</span>'
                for s in detected
            )
            st.markdown(chips, unsafe_allow_html=True)
        else:
            st.write("No known skills detected.")

        report = (
            "Dev AI - Resume Analysis Report\n"
            f"Job Match Score: {score:.1f}%\n"
            f"Matched Skills: {', '.join(matched) or 'None'}\n"
            f"Missing Skills: {', '.join(missing) or 'None'}\n"
            f"Skills Detected: {', '.join(detected) or 'None'}\n"
        )
        if feedback:
            report += "\nGemini Feedback:\n" + feedback + "\n"

        st.download_button(
            "Download Analysis Report",
            data=report,
            file_name="resume_analysis_report.txt",
            mime="text/plain",
            use_container_width=True,
        )

# ---------------- Footer ----------------
st.divider()
st.markdown(
    '<div style="text-align:center;">'
    '<h3>Dev AI</h3>'
    '<p style="letter-spacing:3px;">RESUME ANALYZER</p>'
    '<p>Python · Streamlit · Google Gemini AI · PDF OCR</p>'
    '<p>Built by Dewansh Pandey</p>'
    '</div>',
    unsafe_allow_html=True,
)
