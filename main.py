import streamlit as st
from pypdf import PdfReader
from docx import Document
from pathlib import Path
import re
import json
import urllib.request
# =====================================================
# PAGE CONFIGURATION
# =====================================================
st.set_page_config(
    page_title="Dev AI | Resume Analyzer",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)
# =====================================================
# PROFESSIONAL COLOUR THEME
# Navy + White + Teal
# =====================================================
st.markdown("""
<style>
:root {
    --background: #0b1220;
    --surface: #111c2e;
    --border: #26364b;
    --primary: #14b8a6;
    --text: #f1f5f9;
    --muted: #a6b3c5;
}
.stApp {
    background: var(--background);
    color: var(--text);
}
.block-container {
    max-width: 1180px;
    padding-top: 2rem;
    padding-bottom: 2rem;
}
header[data-testid="stHeader"] {
    background: transparent;
}
h1, h2, h3 {
    color: var(--text) !important;
}
p, label, .stCaption {
    color: var(--muted);
}
/* BRAND HEADER */
.brand-header {
    display: flex;
    align-items: center;
    gap: 24px;
    padding: 30px;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 18px;
    margin-bottom: 30px;
}
.brand-logo {
    width: 112px;
    height: 112px;
    min-width: 112px;
    border-radius: 50%;
    object-fit: cover;
    border: 3px solid var(--primary);
    background: #ffffff;
}
.brand-name {
    color: #ffffff;
    font-size: 42px;
    font-weight: 800;
    letter-spacing: -1.5px;
    line-height: 1.15;
}
.brand-subtitle {
    color: #5eead4;
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 4px;
    margin-top: 7px;
}
.brand-description {
    color: #b6c3d4;
    font-size: 15px;
    line-height: 1.7;
    margin-top: 14px;
    max-width: 700px;
}
/* SECTION CARDS */
div[data-testid="stFileUploader"],
div[data-testid="stTextArea"] {
    border-radius: 12px;
}
div[data-testid="stFileUploader"] {
    background: #0f192a;
    border: 1px dashed #40536b;
    padding: 12px;
}
.stTextArea textarea {
    background: #0f192a !important;
    color: #f1f5f9 !important;
    border: 1px solid #35465d !important;
    border-radius: 10px !important;
}
.stTextArea textarea:focus {
    border-color: #14b8a6 !important;
    box-shadow: 0 0 0 1px #14b8a6 !important;
}
/* BUTTONS */
.stButton > button,
.stDownloadButton > button {
    background: #0f9f91;
    color: #ffffff;
    border: 1px solid #14b8a6;
    border-radius: 10px;
    min-height: 46px;
    font-weight: 700;
    width: 100%;
    transition: 0.2s ease;
}
.stButton > button:hover,
.stDownloadButton > button:hover {
    background: #14b8a6;
    color: #07111f;
    border-color: #5eead4;
}
/* RESULT METRICS */
.metric-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 22px 12px;
    text-align: center;
    min-height: 112px;
}
.metric-number {
    color: #5eead4;
    font-size: 30px;
    font-weight: 800;
}
.metric-label {
    color: #a6b3c5;
    font-size: 13px;
    margin-top: 7px;
}
/* SKILL TAGS */
.skill-tag {
    display: inline-block;
    background: #123a38;
    color: #99f6e4;
    border: 1px solid #22665f;
    border-radius: 7px;
    padding: 7px 11px;
    margin: 4px;
    font-size: 13px;
}
.missing-tag {
    display: inline-block;
    background: #35232b;
    color: #fecdd3;
    border: 1px solid #70404d;
    border-radius: 7px;
    padding: 7px 11px;
    margin: 4px;
    font-size: 13px;
}
/* INFO PANELS */
.info-panel {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 18px;
    margin: 10px 0;
}
.footer {
    text-align: center;
    color: #8b9bb0;
    border-top: 1px solid var(--border);
    padding-top: 22px;
    margin-top: 45px;
    font-size: 13px;
}
hr {
    border-color: var(--border) !important;
}
.stProgress > div > div > div > div {
    background: #14b8a6;
}
div[data-testid="stAlert"] {
    border-radius: 10px;
}
/* MOBILE RESPONSIVENESS */
@media (max-width: 640px) {
    .block-container {
        padding: 1rem;
    }
    .brand-header {
        padding: 18px;
        gap: 14px;
        align-items: center;
    }
    .brand-logo {
        width: 76px;
        height: 76px;
        min-width: 76px;
    }
    .brand-name {
        font-size: 29px;
    }
    .brand-subtitle {
        font-size: 10px;
        letter-spacing: 2px;
    }
    .brand-description {
        font-size: 13px;
    }
}
</style>
""", unsafe_allow_html=True)
# =====================================================
# LOGO
# logo.jpg should be in the same folder as main.py
# =====================================================
def load_logo():
    possible_paths = [
        Path(__file__).parent / "logo.jpg",
        Path(__file__).parent / "logo.png",
        Path(__file__).parent / "static" / "logo.jpg",
        Path(__file__).parent / "static" / "logo.png",
    ]
    for logo_path in possible_paths:
        if logo_path.exists():
            import base64
            try:
                image_data = base64.b64encode(
                    logo_path.read_bytes()
                ).decode("utf-8")
                extension = logo_path.suffix.lower()
                mime_type = (
                    "image/png" if extension == ".png"
                    else "image/jpeg"
                )
                return (
                    f"data:{mime_type};base64,{image_data}"
                )
            except OSError:
                continue
    return None
logo_data = load_logo()
if logo_data:
    logo_html = (
        f'<img src="{logo_data}" class="brand-logo" '
        f'alt="Dev AI Logo">'
    )
else:
    logo_html = """
    <div class="brand-logo"
         style="display:flex;align-items:center;
         justify-content:center;color:#0f766e;
         font-size:27px;font-weight:800;">
        AI
    </div>
    """
# =====================================================
# SKILLS DATABASE
# =====================================================
SKILLS = [
    "python", "java", "javascript", "typescript",
    "html", "css", "react", "node.js", "express",
    "fastapi", "django", "flask", "sql", "mysql",
    "postgresql", "mongodb", "nosql", "git", "github",
    "rest api", "rest apis", "restful api",
    "machine learning", "deep learning",
    "artificial intelligence", "generative ai",
    "llm", "llms", "nlp", "pandas", "numpy",
    "pytorch", "tensorflow", "scikit-learn",
    "streamlit", "data analysis", "data visualization",
    "docker", "aws", "azure", "communication",
    "problem solving", "power bi", "excel", "tableau",
    "c++", "oop", "data structures", "algorithms",
    "transformers", "rag", "langchain",
    "prompt engineering", "statistics", "matplotlib",
    "seaborn", "linux", "agile", "testing",
    "debugging", "next.js", "redux", "tailwind css",
    "bootstrap", "php", "redis", "kubernetes", "ci/cd",
    "leadership", "teamwork", "project management",
    "rest api development", "api integration",
    "responsive design", "web development",
    "communication skills"
]
# =====================================================
# RESUME TEXT EXTRACTION
# =====================================================
def extract_resume_text(uploaded_file):
    try:
        filename = uploaded_file.name.lower()
        if filename.endswith(".pdf"):
            reader = PdfReader(uploaded_file)
            return "\n".join(
                page.extract_text() or ""
                for page in reader.pages
            )
        elif filename.endswith(".docx"):
            document = Document(uploaded_file)
            parts = [
                paragraph.text
                for paragraph in document.paragraphs
                if paragraph.text.strip()
            ]
            for table in document.tables:
                for row in table.rows:
                    parts.append(
                        " | ".join(
                            cell.text for cell in row.cells
                        )
                    )
            return "\n".join(parts)
        elif filename.endswith(".txt"):
            return uploaded_file.getvalue().decode(
                "utf-8", errors="ignore"
            )
        return ""
    except Exception as error:
        st.error(f"Resume read karte waqt error aaya: {error}")
        return ""
# =====================================================
# SKILL DETECTION
# =====================================================
def find_skills(text):
    text = text.lower()
    found = []
    for skill in SKILLS:
        pattern = (
            r"(?<![a-z0-9])"
            + re.escape(skill)
            + r"(?![a-z0-9])"
        )
        if re.search(pattern, text):
            found.append(skill)
    return sorted(set(found))
# =====================================================
# RESUME QUALITY ANALYSIS
# =====================================================
def analyze_resume_quality(text):
    lower = text.lower()
    words = re.findall(r"\b[\w+#.-]+\b", text)
    word_count = len(words)
    sections = {
        "Contact Information": bool(
            re.search(
                r"\b[\w.+-]+@[\w.-]+\.[a-z]{2,}\b",
                lower
            )
        ) or bool(re.search(r"\b\d{10}\b", text)),
        "Professional Summary": any(
            term in lower
            for term in [
                "summary", "objective",
                "professional profile"
            ]
        ),
        "Work Experience": any(
            term in lower
            for term in [
                "experience", "employment",
                "work history"
            ]
        ),
        "Education": any(
            term in lower
            for term in [
                "education", "university",
                "degree", "b.tech", "bachelor"
            ]
        ),
        "Projects": "project" in lower,
        "Skills": "skills" in lower
    }
    sections_found = sum(sections.values())
    section_score = round(
        sections_found / len(sections) * 100
    )
    if 250 <= word_count <= 900:
        length_score = 100
    elif (
        150 <= word_count < 250
        or 900 < word_count <= 1200
    ):
        length_score = 75
    else:
        length_score = 45
    quality_score = round(
        section_score * 0.7
        + length_score * 0.3
    )
    return {
        "word_count": word_count,
        "sections": sections,
        "section_score": section_score,
        "length_score": length_score,
        "quality_score": quality_score
    }
# =====================================================
# LOCAL AI WITH OLLAMA
# =====================================================
def ask_local_ai(prompt):
    payload = {
        "model": "llama3.2:1b",
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0.3,
            "num_predict": 700
        }
    }
    request = urllib.request.Request(
        "http://localhost:11434/api/generate",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(
        request, timeout=180
    ) as response:
        data = json.loads(
            response.read().decode("utf-8")
        )
        return data.get("response", "").strip()
# =====================================================
# RECOMMENDATIONS
# =====================================================
def generate_recommendations(
    resume_text, quality, missing_skills
):
    recommendations = []
    for section, present in quality["sections"].items():
        if not present:
            recommendations.append(
                f"Add a clearly labelled {section} section if relevant."
            )
    if missing_skills:
        recommendations.append(
            "Review these job requirements and include only "
            "the skills you genuinely know: "
            + ", ".join(missing_skills) + "."
        )
    if "project" not in resume_text.lower():
        recommendations.append(
            "Add relevant projects with your contribution, "
            "technology stack, and measurable outcomes."
        )
    if not re.search(r"\b\d+%", resume_text):
        recommendations.append(
            "Where truthful, quantify achievements using numbers, "
            "such as time saved, users supported, or project results."
        )
    if quality["word_count"] < 150:
        recommendations.append(
            "The extracted resume appears short. Add relevant "
            "experience, projects, education, and skills."
        )
    elif quality["word_count"] > 1000:
        recommendations.append(
            "Consider removing repetitive details and prioritizing "
            "content relevant to the target role."
        )
    if not recommendations:
        recommendations.append(
            "Your resume contains the main sections detected by "
            "this tool. Tailor your achievements to the target role."
        )
    return recommendations
# =====================================================
# BRAND HEADER
# =====================================================
st.markdown(
    f"""
    <div class="brand-header">
        {logo_html}
        <div class="brand-content">
            <div class="brand-name">Dev AI</div>
            <div class="brand-subtitle">RESUME ANALYZER</div>
            <div class="brand-description">
                Analyze your resume, compare job requirements,
                identify missing skills, and receive personalized
                recommendations with local AI assistance.
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)
# =====================================================
# INPUT SECTION
# =====================================================
st.header("Analyze Your Resume")
st.write(
    "Upload your resume and add the target job description "
    "to get a personalized analysis."
)
left, right = st.columns(2, gap="large")
with left:
    st.subheader("📄 Upload Resume")
    uploaded_file = st.file_uploader(
        "Choose your resume",
        type=["pdf", "docx", "txt"],
        help="Supported formats: PDF, DOCX and TXT."
    )
with right:
    st.subheader("🎯 Target Job")
    job_description = st.text_area(
        "Paste the job description",
        placeholder=(
            "Paste the job responsibilities, required skills, "
            "qualifications and experience here..."
        ),
        height=180
    )
st.markdown("")
use_local_ai = st.checkbox(
    "Enable local AI feedback (Ollama)",
    value=True
)
st.caption(
    "Privacy: Resume analysis runs on your computer. "
    "Local AI feedback requires Ollama and the llama3.2:1b model."
)
analyze = st.button(
    "Analyze Resume",
    use_container_width=True
)
# =====================================================
# RUN ANALYSIS
# =====================================================
if analyze:
    if not uploaded_file:
        st.warning("Please upload your resume first.")
        st.stop()
    with st.spinner("Reading your resume..."):
        resume_text = extract_resume_text(uploaded_file)
    if not resume_text.strip():
        st.error(
            "No readable text was found. If your PDF is scanned, "
            "you may need OCR to extract its text."
        )
        st.stop()
    resume_skills = find_skills(resume_text)
    job_skills = find_skills(job_description)
    matched_skills = [
        skill for skill in job_skills
        if skill in resume_skills
    ]
    missing_skills = [
        skill for skill in job_skills
        if skill not in resume_skills
    ]
    quality = analyze_resume_quality(resume_text)
    if job_description.strip() and job_skills:
        skill_match_score = round(
            len(matched_skills) / len(job_skills) * 100
        )
    else:
        skill_match_score = None
    if skill_match_score is not None:
        ats_score = round(
            skill_match_score * 0.7
            + quality["quality_score"] * 0.3
        )
    else:
        ats_score = quality["quality_score"]
    recommendations = generate_recommendations(
        resume_text,
        quality,
        missing_skills
    )
    ai_feedback = ""
    ai_status = "Not requested"
    if use_local_ai:
        prompt = f"""
You are a professional resume reviewer.
Analyze the resume against the target job description.
Provide these sections:
1. Resume strengths
2. Three important improvements
3. Job-specific tailoring suggestions
4. Honest example bullet points based only on supplied facts
5. Interview preparation priorities
Never invent experience, skills, qualifications, or achievements.
If evidence is missing, clearly say so.
RESUME:
{resume_text[:9000]}
JOB DESCRIPTION:
{job_description[:4000] if job_description.strip() else "Not provided"}
DETECTED SKILLS:
{", ".join(resume_skills)}
MISSING JOB SKILLS:
{", ".join(missing_skills)}
"""
        try:
            with st.spinner(
                "Local AI is reviewing your resume. Please wait..."
            ):
                ai_feedback = ask_local_ai(prompt)
            ai_status = "Connected — llama3.2:1b"
        except Exception as error:
            ai_status = "Unavailable"
            st.warning(
                "Local AI is unavailable. The resume analysis "
                "will still be shown. Make sure Ollama is running "
                "and llama3.2:1b is installed."
            )
            with st.expander("View connection details"):
                st.code(str(error))
    st.session_state["analysis_result"] = {
        "filename": uploaded_file.name,
        "resume_skills": resume_skills,
        "job_skills": job_skills,
        "matched_skills": matched_skills,
        "missing_skills": missing_skills,
        "quality": quality,
        "skill_match_score": skill_match_score,
        "ats_score": ats_score,
        "recommendations": recommendations,
        "ai_feedback": ai_feedback,
        "ai_status": ai_status
    }
# =====================================================
# DISPLAY RESULTS
# =====================================================
if "analysis_result" in st.session_state:
    result = st.session_state["analysis_result"]
    st.markdown("---")
    st.header("Analysis Results")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-number">
                    {result["ats_score"]}%
                </div>
                <div class="metric-label">
                    ATS-Style Estimate
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with c2:
        match_value = (
            f'{result["skill_match_score"]}%'
            if result["skill_match_score"] is not None
            else "N/A"
        )
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-number">
                    {match_value}
                </div>
                <div class="metric-label">
                    Job Skill Match
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with c3:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-number">
                    {len(result["matched_skills"])}
                </div>
                <div class="metric-label">
                    Matched Skills
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with c4:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-number">
                    {result["quality"]["word_count"]}
                </div>
                <div class="metric-label">
                    Resume Words
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    st.markdown("")
    st.progress(result["ats_score"] / 100)
    st.caption(
        "Note: This is a heuristic ATS-style estimate, not an "
        "official ATS score. Different recruitment systems use "
        "different scoring methods."
    )
    # -----------------------------------------------
    # MATCHED SKILLS
    # -----------------------------------------------
    st.subheader("Matched Skills")
    if result["matched_skills"]:
        tags = "".join(
            f'<span class="skill-tag">{skill}</span>'
            for skill in result["matched_skills"]
        )
        st.markdown(tags, unsafe_allow_html=True)
    else:
        st.info(
            "No matching skills were detected from the supported "
            "skills list."
        )
    # -----------------------------------------------
    # MISSING SKILLS
    # -----------------------------------------------
    st.subheader("Missing Job Skills")
    if result["missing_skills"]:
        tags = "".join(
            f'<span class="missing-tag">{skill}</span>'
            for skill in result["missing_skills"]
        )
        st.markdown(tags, unsafe_allow_html=True)
    elif result["job_skills"]:
        st.success(
            "All detected job skills were found in the resume."
        )
    else:
        st.info(
            "No supported skills were detected in the job description."
        )
    # -----------------------------------------------
    # RESUME SECTION ANALYSIS
    # -----------------------------------------------
    st.subheader("Resume Section Analysis")
    for section, present in result["quality"]["sections"].items():
        if present:
            st.write(f"✅ {section}: Detected")
        else:
            st.write(f"⚠️ {section}: Not detected")
    st.write(
        f'**Section coverage:** '
        f'{result["quality"]["section_score"]}%'
    )
    # -----------------------------------------------
    # RECOMMENDATIONS
    # -----------------------------------------------
    st.subheader("Personalized Recommendations")
    for index, recommendation in enumerate(
        result["recommendations"],
        start=1
    ):
        st.write(f"{index}. {recommendation}")
    # -----------------------------------------------
    # LOCAL AI FEEDBACK
    # -----------------------------------------------
    st.subheader("Local AI Feedback")
    st.caption(f'Model status: {result["ai_status"]}')
    if result["ai_feedback"]:
        st.markdown(result["ai_feedback"])
    else:
        st.info(
            "AI feedback is unavailable. The rule-based analysis "
            "above is still available."
        )
    # -----------------------------------------------
    # ALL RESUME SKILLS
    # -----------------------------------------------
    st.subheader("Skills Detected in Resume")
    if result["resume_skills"]:
        tags = "".join(
            f'<span class="skill-tag">{skill}</span>'
            for skill in result["resume_skills"]
        )
        st.markdown(tags, unsafe_allow_html=True)
    else:
        st.write(
            "No skills from the current skills database were detected."
        )
    # -----------------------------------------------
    # DOWNLOAD REPORT
    # -----------------------------------------------
    section_report = "\n".join(
        f'{section}: {"Detected" if present else "Not detected"}'
        for section, present in result["quality"]["sections"].items()
    )
    match_report = (
        f'{result["skill_match_score"]}%'
        if result["skill_match_score"] is not None
        else "Not calculated"
    )
    recommendation_report = "\n".join(
        "- " + item
        for item in result["recommendations"]
    )
    report = f"""
DEV AI — RESUME ANALYZER
========================================
Resume file: {result["filename"]}
ATS-style estimate: {result["ats_score"]}%
Job skill match: {match_report}
Resume word count: {result["quality"]["word_count"]}
Section coverage: {result["quality"]["section_score"]}%
MATCHED SKILLS
----------------------------------------
{", ".join(result["matched_skills"]) or "None"}
MISSING JOB SKILLS
----------------------------------------
{", ".join(result["missing_skills"]) or "None"}
SKILLS FOUND IN RESUME
----------------------------------------
{", ".join(result["resume_skills"]) or "None"}
SECTION ANALYSIS
----------------------------------------
{section_report}
RECOMMENDATIONS
----------------------------------------
{recommendation_report}
LOCAL AI STATUS
----------------------------------------
{result["ai_status"]}
LOCAL AI FEEDBACK
----------------------------------------
{result["ai_feedback"] or "Not available"}
DISCLAIMER
----------------------------------------
This report is a heuristic resume analysis. It is not an
official ATS score, hiring decision, or guarantee of selection.
"""
    st.download_button(
        label="Download Full Analysis Report",
        data=report,
        file_name="Dev_AI_Resume_Analysis.txt",
        mime="text/plain",
        use_container_width=True
    )
# =====================================================
# FOOTER
# =====================================================
st.markdown("""
<div class="footer">
    <strong style="color:#f1f5f9;font-size:16px;">
        Dev AI
    </strong>
    <div style="color:#5eead4;letter-spacing:3px;
                margin-top:5px;">
        RESUME ANALYZER
    </div>
    <p>Python · Streamlit · Local AI with Ollama</p>
    <p>Built by Dewansh Pandey</p>
</div>
""", unsafe_allow_html=True)