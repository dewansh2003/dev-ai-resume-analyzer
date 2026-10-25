import streamlit as st
from pypdf import PdfReader
from docx import Document
from pathlib import Path
from google import genai
import re
import io
import base64

# Optional OCR libraries
try:
    import pytesseract
    from pdf2image import convert_from_bytes
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


# =====================================================
# PAGE CONFIGURATION
# =====================================================
st.set_page_config(
    page_title="Dev AI | Resume Analyzer",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# =====================================================
# CUSTOM DESIGN
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

h1, h2, h3 {
    color: var(--text) !important;
}

p, label {
    color: var(--muted);
}

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
    width: 105px;
    height: 105px;
    min-width: 105px;
    border-radius: 50%;
    object-fit: cover;
    border: 3px solid var(--primary);
    background: white;
}

.brand-name {
    color: white;
    font-size: 42px;
    font-weight: 800;
}

.brand-subtitle {
    color: #5eead4;
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 4px;
}

.brand-description {
    color: #b6c3d4;
    font-size: 15px;
    line-height: 1.7;
    margin-top: 12px;
}

.stTextArea textarea {
    background: #0f192a !important;
    color: white !important;
    border: 1px solid #35465d !important;
    border-radius: 10px !important;
}

.stButton > button,
.stDownloadButton > button {
    background: #0f9f91;
    color: white;
    border: 1px solid #14b8a6;
    border-radius: 10px;
    min-height: 45px;
    font-weight: 700;
    width: 100%;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background: #14b8a6;
    color: #07111f;
}

.metric-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 22px 10px;
    text-align: center;
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

.skill-tag, .missing-tag {
    display: inline-block;
    border-radius: 7px;
    padding: 7px 11px;
    margin: 4px;
    font-size: 13px;
}

.skill-tag {
    background: #123a38;
    color: #99f6e4;
    border: 1px solid #22665f;
}

.missing-tag {
    background: #35232b;
    color: #fecdd3;
    border: 1px solid #70404d;
}

.footer {
    text-align: center;
    color: #8b9bb0;
    border-top: 1px solid var(--border);
    padding-top: 22px;
    margin-top: 45px;
}

@media(max-width: 640px) {
    .brand-header {
        padding: 18px;
        gap: 14px;
    }

    .brand-logo {
        width: 75px;
        height: 75px;
        min-width: 75px;
    }

    .brand-name {
        font-size: 28px;
    }

    .brand-subtitle {
        font-size: 10px;
        letter-spacing: 2px;
    }
}
</style>
""", unsafe_allow_html=True)


# =====================================================
# LOGO
# =====================================================
def load_logo():
    paths = [
        Path(__file__).parent / "logo.jpg",
        Path(__file__).parent / "logo.png",
        Path(__file__).parent / "static" / "logo.jpg",
        Path(__file__).parent / "static" / "logo.png",
    ]

    for path in paths:
        if path.exists():
            try:
                image_data = base64.b64encode(
                    path.read_bytes()
                ).decode("utf-8")

                mime = (
                    "image/png"
                    if path.suffix.lower() == ".png"
                    else "image/jpeg"
                )

                return f"data:{mime};base64,{image_data}"

            except OSError:
                pass

    return None


logo = load_logo()

if logo:
    logo_html = (
        f'<img src="{logo}" class="brand-logo" alt="Dev AI Logo">'
    )
else:
    logo_html = """
    <div class="brand-logo"
         style="display:flex;align-items:center;
         justify-content:center;color:#0f766e;
         font-size:30px;font-weight:800;">
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
    "rest api", "rest apis", "machine learning",
    "deep learning", "artificial intelligence",
    "generative ai", "llm", "llms", "nlp",
    "pandas", "numpy", "pytorch", "tensorflow",
    "scikit-learn", "streamlit", "data analysis",
    "data visualization", "docker", "aws", "azure",
    "communication", "problem solving", "power bi",
    "excel", "tableau", "c++", "oop",
    "data structures", "algorithms", "transformers",
    "rag", "langchain", "prompt engineering",
    "statistics", "matplotlib", "seaborn", "linux",
    "agile", "testing", "debugging", "next.js",
    "redux", "tailwind css", "bootstrap", "php",
    "redis", "kubernetes", "ci/cd", "leadership",
    "teamwork", "project management",
    "api integration", "responsive design",
    "web development", "communication skills",
]


# =====================================================
# RESUME TEXT EXTRACTION + OCR
# =====================================================
def extract_resume_text(uploaded_file, use_ocr=True):
    try:
        filename = uploaded_file.name.lower()
        raw_bytes = uploaded_file.getvalue()

        # PDF FILE
        if filename.endswith(".pdf"):
            reader = PdfReader(io.BytesIO(raw_bytes))

            pages = []

            for page in reader.pages:
                try:
                    pages.append(page.extract_text() or "")
                except Exception:
                    pages.append("")

            text = "\n".join(pages).strip()

            # Try OCR if PDF has little readable text
            if len(text) < 80 and use_ocr:

                if not OCR_AVAILABLE:
                    st.warning(
                        "OCR libraries are unavailable. "
                        "Install pytesseract, pdf2image, "
                        "Tesseract and Poppler."
                    )
                    return text

                try:
                    with st.spinner(
                        "Scanned PDF detected. Running OCR..."
                    ):
                        images = convert_from_bytes(
                            raw_bytes,
                            dpi=200
                        )

                        ocr_pages = [
                            pytesseract.image_to_string(image)
                            for image in images
                        ]

                        ocr_text = "\n".join(ocr_pages).strip()

                    if len(ocr_text) > len(text):
                        text = ocr_text

                except Exception as error:
                    st.warning(
                        "OCR failed. Check Tesseract and "
                        "Poppler installation."
                    )

                    with st.expander("OCR error details"):
                        st.code(str(error))

            return text

        # WORD DOCUMENT
        elif filename.endswith(".docx"):
            document = Document(io.BytesIO(raw_bytes))

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

            return "\n".join(parts).strip()

        # TEXT FILE
        elif filename.endswith(".txt"):
            return raw_bytes.decode(
                "utf-8",
                errors="ignore"
            ).strip()

        return ""

    except Exception as error:
        st.error(f"Resume reading error: {error}")
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
        "Contact Information": (
            bool(re.search(
                r"\b[\w.+-]+@[\w.-]+\.[a-z]{2,}\b",
                lower
            ))
            or bool(re.search(r"\b\d{10}\b", text))
        ),

        "Professional Summary": any(
            term in lower
            for term in [
                "summary",
                "objective",
                "professional profile"
            ]
        ),

        "Work Experience": any(
            term in lower
            for term in [
                "experience",
                "employment",
                "work history"
            ]
        ),

        "Education": any(
            term in lower
            for term in [
                "education",
                "university",
                "degree",
                "b.tech",
                "bachelor"
            ]
        ),

        "Projects": "project" in lower,

        "Skills": "skills" in lower,
    }

    section_score = round(
        sum(sections.values()) / len(sections) * 100
    )

    if 250 <= word_count <= 900:
        length_score = 100
    elif 150 <= word_count < 250 or 900 < word_count <= 1200:
        length_score = 75
    else:
        length_score = 45

    quality_score = round(
        section_score * 0.7 + length_score * 0.3
    )

    return {
        "word_count": word_count,
        "sections": sections,
        "section_score": section_score,
        "length_score": length_score,
        "quality_score": quality_score,
    }


# =====================================================
# GEMINI AI
# =====================================================
def ask_gemini_ai(prompt):
    api_key = st.secrets.get("GEMINI_API_KEY", "")

    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY is missing. "
            "Add it in Streamlit Cloud Settings > Secrets."
        )

    client = genai.Client(api_key=api_key)

    response = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=prompt
    )

    answer = getattr(response, "text", None)

    if not answer:
        raise RuntimeError(
            "Gemini returned an empty response. "
            "Check API access and quota."
        )

    return answer.strip()


# =====================================================
# RECOMMENDATIONS
# =====================================================
def generate_recommendations(
    resume_text,
    quality,
    missing_skills
):
    recommendations = []

    for section, present in quality["sections"].items():
        if not present:
            recommendations.append(
                f"Add a clearly labelled {section} section "
                "if relevant."
            )

    if missing_skills:
        recommendations.append(
            "Review these job skills and include only "
            "those you genuinely know: "
            + ", ".join(missing_skills)
            + "."
        )

    if "project" not in resume_text.lower():
        recommendations.append(
            "Add relevant projects, your contribution, "
            "technologies used and measurable outcomes."
        )

    if not re.search(r"\b\d+%", resume_text):
        recommendations.append(
            "Where truthful, quantify achievements "
            "with numbers and measurable results."
        )

    if quality["word_count"] < 150:
        recommendations.append(
            "The resume appears short. Add relevant "
            "experience, projects, education and skills."
        )

    elif quality["word_count"] > 1000:
        recommendations.append(
            "Remove repetitive details and prioritize "
            "information relevant to the target role."
        )

    if not recommendations:
        recommendations.append(
            "Your resume contains the main sections "
            "detected by this tool. Tailor it to the job."
        )

    return recommendations


# =====================================================
# HEADER
# =====================================================
st.markdown(f"""
<div class="brand-header">
    {logo_html}
    <div>
        <div class="brand-name">Dev AI</div>
        <div class="brand-subtitle">RESUME ANALYZER</div>
        <div class="brand-description">
            Analyze your resume, compare job requirements,
            identify missing skills, and receive personalized
            recommendations with Gemini AI assistance.
        </div>
    </div>
</div>
""", unsafe_allow_html=True)


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
        type=["pdf", "docx", "txt"]
    )

with right:
    st.subheader("🎯 Target Job")

    job_description = st.text_area(
        "Paste the job description",
        placeholder=(
            "Paste job responsibilities, required skills, "
            "qualifications and experience here..."
        ),
        height=180
    )

use_ai = st.checkbox(
    "Enable Gemini AI feedback",
    value=True
)

use_ocr = st.checkbox(
    "Try OCR for scanned PDFs",
    value=True
)

st.caption(
    "When enabled, resume text is sent to Google Gemini "
    "for AI feedback. OCR attempts to read scanned PDFs "
    "when the required OCR dependencies are installed."
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
        resume_text = extract_resume_text(
            uploaded_file,
            use_ocr=use_ocr
        )

    if not resume_text.strip():
        st.error(
            "No readable text could be extracted. "
            "Try a text-based PDF or upload a DOCX/TXT copy. "
            "For scanned PDFs, check OCR installation."
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

    if use_ai:

        prompt = f"""
You are a professional resume reviewer.

Analyze the resume against the target job description.

Provide:
1. Resume strengths.
2. Three important improvements.
3. Job-specific tailoring suggestions.
4. Example resume bullet points based only on supplied facts.
5. Interview preparation priorities.

Never invent experience, skills, qualifications,
or achievements. Clearly flag missing evidence.

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
                "Gemini AI is reviewing your resume..."
            ):
                ai_feedback = ask_gemini_ai(prompt)

            ai_status = "Connected — Gemini API"

        except Exception as error:
            ai_status = "Unavailable"

            st.warning(
                "Gemini AI feedback failed. "
                "Rule-based analysis is still available."
            )

            with st.expander("View technical details"):
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
        "ai_status": ai_status,
    }


# =====================================================
# RESULTS
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

    st.progress(
        min(max(result["ats_score"] / 100, 0.0), 1.0)
    )

    st.caption(
        "This is a heuristic ATS-style estimate, "
        "not an official ATS score or hiring decision."
    )

    # Matched skills
    st.subheader("Matched Skills")

    if result["matched_skills"]:
        st.markdown(
            "".join(
                f'<span class="skill-tag">{skill}</span>'
                for skill in result["matched_skills"]
            ),
            unsafe_allow_html=True
        )
    else:
        st.info(
            "No matching skills were detected "
            "from the supported skills list."
        )

    # Missing skills
    st.subheader("Missing Job Skills")

    if result["missing_skills"]:
        st.markdown(
            "".join(
                f'<span class="missing-tag">{skill}</span>'
                for skill in result["missing_skills"]
            ),
            unsafe_allow_html=True
        )

    elif result["job_skills"]:
        st.success(
            "All detected job skills were found in the resume."
        )

    else:
        st.info(
            "No supported skills were detected "
            "in the job description."
        )

    # Section analysis
    st.subheader("Resume Section Analysis")

    for section, present in result["quality"]["sections"].items():
        status = "Detected" if present else "Not detected"
        icon = "✅" if present else "⚠️"

        st.write(f"{icon} {section}: {status}")

    st.write(
        f'**Section coverage:** '
        f'{result["quality"]["section_score"]}%'
    )

    # Recommendations
    st.subheader("Personalized Recommendations")

    for i, recommendation in enumerate(
        result["recommendations"],
        start=1
    ):
        st.write(f"{i}. {recommendation}")

    # Gemini feedback
    st.subheader("Gemini AI Feedback")

    st.caption(f'Model status: {result["ai_status"]}')

    if result["ai_feedback"]:
        st.markdown(result["ai_feedback"])
    else:
        st.info(
            "AI feedback is unavailable or was not requested. "
            "Rule-based analysis is still available."
        )

    # All resume skills
    st.subheader("Skills Detected in Resume")

    if result["resume_skills"]:
        st.markdown(
            "".join(
                f'<span class="skill-tag">{skill}</span>'
                for skill in result["resume_skills"]
            ),
            unsafe_allow_html=True
        )
    else:
        st.write(
            "No skills from the current skills database were detected."
        )

    # Report download
    section_report = "\n".join(
        f'{section}: {"Detected" if present else "Not detected"}'
        for section, present
        in result["quality"]["sections"].items()
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

GEMINI AI STATUS
----------------------------------------
{result["ai_status"]}

GEMINI AI FEEDBACK
----------------------------------------
{result["ai_feedback"] or "Not available"}

DISCLAIMER
----------------------------------------
This report is a heuristic resume analysis,
not an official ATS score or hiring decision.
"""

    st.download_button(
        "Download Full Analysis Report",
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

    <div style="color:#5eead4;letter-spacing:3px;margin-top:5px;">
        RESUME ANALYZER
    </div>

    <p>Python · Streamlit · Google Gemini AI · PDF OCR</p>
    <p>Built by Dewansh Pandey</p>
</div>
""", unsafe_allow_html=True)
