import io
import re
import base64
from pathlib import Path

import streamlit as st
from pypdf import PdfReader
from docx import Document
from google import genai
from google.genai import types


# ==================================================
# PAGE CONFIG
# ==================================================
st.set_page_config(
    page_title="Dev AI | Resume Analyzer",
    page_icon="📄",
    layout="wide",
)


# ==================================================
# DESIGN
# ==================================================
st.markdown("""
<style>
.stApp {
    background: #0b1220;
    color: #f1f5f9;
}

.block-container {
    max-width: 1150px;
    padding-top: 2rem;
    padding-bottom: 2rem;
}

h1, h2, h3 {
    color: #f1f5f9 !important;
}

p, label {
    color: #b6c3d4;
}

.stTextArea textarea {
    background: #111c2e !important;
    color: white !important;
    border: 1px solid #35465d !important;
    border-radius: 10px !important;
}

.stButton > button,
.stDownloadButton > button {
    background: #14b8a6;
    color: #07111f;
    border: none;
    border-radius: 10px;
    font-weight: 700;
    min-height: 44px;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background: #5eead4;
    color: #07111f;
}

.header-box {
    display: flex;
    align-items: center;
    gap: 22px;
    background: #111c2e;
    border: 1px solid #26364b;
    border-radius: 18px;
    padding: 25px;
    margin-bottom: 28px;
}

.brand-logo {
    width: 90px;
    height: 90px;
    min-width: 90px;
    object-fit: cover;
    border-radius: 50%;
    border: 3px solid #14b8a6;
    background: white;
}

.brand-title {
    font-size: 40px;
    font-weight: 800;
    color: white;
}

.brand-subtitle {
    font-size: 13px;
    letter-spacing: 4px;
    color: #5eead4;
    font-weight: 700;
}

.metric-card {
    background: #111c2e;
    border: 1px solid #26364b;
    border-radius: 14px;
    padding: 20px 8px;
    text-align: center;
}

.metric-value {
    font-size: 30px;
    font-weight: 800;
    color: #5eead4;
}

.metric-label {
    color: #b6c3d4;
    font-size: 13px;
    margin-top: 5px;
}

.skill-tag {
    display: inline-block;
    background: #123a38;
    color: #99f6e4;
    border: 1px solid #22665f;
    border-radius: 7px;
    padding: 6px 10px;
    margin: 4px;
    font-size: 13px;
}

.missing-tag {
    display: inline-block;
    background: #35232b;
    color: #fecdd3;
    border: 1px solid #70404d;
    border-radius: 7px;
    padding: 6px 10px;
    margin: 4px;
    font-size: 13px;
}

.footer-box {
    text-align: center;
    border-top: 1px solid #26364b;
    padding-top: 22px;
    margin-top: 45px;
    color: #8b9bb0;
}

@media(max-width: 600px) {
    .header-box {
        padding: 16px;
        gap: 12px;
    }

    .brand-logo {
        width: 65px;
        height: 65px;
        min-width: 65px;
    }

    .brand-title {
        font-size: 28px;
    }

    .brand-subtitle {
        font-size: 10px;
        letter-spacing: 2px;
    }
}
</style>
""", unsafe_allow_html=True)


# ==================================================
# HEADER AND LOGO
# ==================================================
def get_logo():
    base = Path(__file__).parent

    for filename in ["logo.jpg", "logo.png"]:
        path = base / filename

        if path.exists():
            try:
                image_data = base64.b64encode(
                    path.read_bytes()
                ).decode("utf-8")

                mime = (
                    "image/png"
                    if filename.endswith(".png")
                    else "image/jpeg"
                )

                return (
                    f'<img class="brand-logo" '
                    f'src="data:{mime};base64,{image_data}">'
                )
            except Exception:
                pass

    return """
    <div class="brand-logo"
    style="display:flex;align-items:center;
    justify-content:center;color:#0f766e;
    font-size:28px;font-weight:800;">
    AI
    </div>
    """


logo_html = get_logo()

st.markdown(
    f"""
    <div class="header-box">
        {logo_html}
        <div>
            <div class="brand-title">Dev AI</div>
            <div class="brand-subtitle">RESUME ANALYZER</div>
            <p>
                Analyze your resume, compare job skills,
                and get personalized AI feedback.
            </p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ==================================================
# SKILLS DATABASE
# ==================================================
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
    "docker", "aws", "azure", "power bi", "excel",
    "tableau", "c++", "oop", "data structures",
    "algorithms", "transformers", "rag", "langchain",
    "prompt engineering", "statistics", "linux",
    "next.js", "redux", "tailwind css", "php",
    "redis", "kubernetes", "ci/cd", "leadership",
    "teamwork", "project management",
    "api integration", "responsive design",
    "web development", "communication skills",
    "problem solving", "matplotlib", "seaborn",
]


# ==================================================
# RESUME EXTRACTION
# ==================================================
def extract_resume_text(uploaded_file, use_ocr=True):
    filename = uploaded_file.name.lower()
    raw = uploaded_file.getvalue()

    if filename.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(raw))

        pages = []

        for page in reader.pages:
            try:
                pages.append(page.extract_text() or "")
            except Exception:
                pages.append("")

        text = "\n".join(pages).strip()

        if len(text) < 80 and use_ocr:
            try:
                import pytesseract
                from pdf2image import convert_from_bytes

                with st.spinner(
                    "Trying OCR on your scanned PDF..."
                ):
                    images = convert_from_bytes(raw, dpi=180)

                    ocr_pages = [
                        pytesseract.image_to_string(image)
                        for image in images
                    ]

                    ocr_text = "\n".join(ocr_pages).strip()

                if len(ocr_text) > len(text):
                    text = ocr_text

            except Exception as error:
                st.warning(
                    "OCR could not run. Check that "
                    "Tesseract and Poppler are installed."
                )

                with st.expander("OCR technical details"):
                    st.code(str(error))

        return text

    if filename.endswith(".docx"):
        doc = Document(io.BytesIO(raw))

        parts = [
            paragraph.text
            for paragraph in doc.paragraphs
            if paragraph.text.strip()
        ]

        for table in doc.tables:
            for row in table.rows:
                parts.append(
                    " | ".join(cell.text for cell in row.cells)
                )

        return "\n".join(parts).strip()

    if filename.endswith(".txt"):
        return raw.decode("utf-8", errors="ignore").strip()

    return ""


# ==================================================
# SKILL MATCHING
# ==================================================
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


# ==================================================
# RESUME QUALITY
# ==================================================
def resume_quality(text):
    lower = text.lower()
    words = re.findall(r"\b[\w+#.-]+\b", text)

    sections = {
        "Contact Information": (
            bool(re.search(
                r"\b[\w.+-]+@[\w.-]+\.[a-z]{2,}\b",
                lower
            ))
            or bool(re.search(r"\b\d{10}\b", text))
        ),
        "Summary": any(
            x in lower for x in ["summary", "objective"]
        ),
        "Experience": any(
            x in lower for x in [
                "experience", "employment", "work history"
            ]
        ),
        "Education": any(
            x in lower for x in [
                "education", "university", "b.tech",
                "bachelor", "degree"
            ]
        ),
        "Projects": "project" in lower,
        "Skills": "skills" in lower,
    }

    section_score = round(
        sum(sections.values()) / len(sections) * 100
    )

    count = len(words)

    if 250 <= count <= 900:
        length_score = 100
    elif 150 <= count < 250 or 900 < count <= 1200:
        length_score = 75
    else:
        length_score = 45

    score = round(
        section_score * 0.7 + length_score * 0.3
    )

    return {
        "words": count,
        "sections": sections,
        "section_score": section_score,
        "score": score,
    }


# ==================================================
# GEMINI AI — SHORTER PROMPT AND RESPONSE
# ==================================================
def ask_gemini_ai(resume_text, job_description,
                  resume_skills, missing_skills):

    api_key = st.secrets.get("GEMINI_API_KEY", "")

    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY is missing. Add it in "
            "Streamlit Cloud > Settings > Secrets."
        )

    client = genai.Client(
        api_key=api_key,
        http_options=types.HttpOptions(
            timeout=30000
        ),
    )

    prompt = f"""
You are a professional resume reviewer.

Give concise, useful feedback in English using these headings:
1. Strengths
2. Three Improvements
3. Job-Specific Advice
4. Interview Preparation

Do not invent skills, experience, qualifications or achievements.
Only use information provided below.
Keep the complete response under 350 words.

RESUME:
{resume_text[:5000]}

JOB DESCRIPTION:
{job_description[:2500] if job_description.strip() else "Not provided"}

DETECTED RESUME SKILLS:
{", ".join(resume_skills[:35])}

POSSIBLY MISSING JOB SKILLS:
{", ".join(missing_skills[:25])}
"""

    response = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0.2,
            max_output_tokens=500,
        ),
    )

    answer = getattr(response, "text", None)

    if not answer or not answer.strip():
        raise RuntimeError(
            "Gemini returned an empty response."
        )

    return answer.strip()


# ==================================================
# INPUTS
# ==================================================
st.header("Analyze Your Resume")

left, right = st.columns(2)

with left:
    uploaded_file = st.file_uploader(
        "Upload Resume",
        type=["pdf", "docx", "txt"]
    )

with right:
    job_description = st.text_area(
        "Paste Job Description",
        height=170,
        placeholder="Paste the target job description here...",
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
    "When Gemini feedback is enabled, resume text is sent "
    "to Google Gemini. OCR requires Tesseract and Poppler."
)

analyze = st.button(
    "Analyze Resume",
    use_container_width=True
)


# ==================================================
# ANALYZE
# ==================================================
if analyze:

    if not uploaded_file:
        st.warning("Please upload your resume.")
        st.stop()

    try:
        with st.spinner("Reading your resume..."):
            resume_text = extract_resume_text(
                uploaded_file,
                use_ocr=use_ocr
            )
    except Exception as error:
        st.error(f"Could not read the resume: {error}")
        st.stop()

    if not resume_text.strip():
        st.error(
            "No readable text found. Try a text-based PDF "
            "or upload a DOCX/TXT copy."
        )
        st.stop()

    resume_skills = find_skills(resume_text)
    job_skills = find_skills(job_description)

    matched = [
        skill for skill in job_skills
        if skill in resume_skills
    ]

    missing = [
        skill for skill in job_skills
        if skill not in resume_skills
    ]

    quality = resume_quality(resume_text)

    if job_description.strip() and job_skills:
        match_score = round(
            len(matched) / len(job_skills) * 100
        )
        ats_score = round(
            match_score * 0.7 + quality["score"] * 0.3
        )
    else:
        match_score = None
        ats_score = quality["score"]

    recommendations = []

    for section, present in quality["sections"].items():
        if not present:
            recommendations.append(
                f"Consider adding a clear {section} section "
                "if relevant to your experience."
            )

    if missing:
        recommendations.append(
            "Review these job skills and include only those "
            "you genuinely know: " + ", ".join(missing)
        )

    if not re.search(r"\b\d+%", resume_text):
        recommendations.append(
            "Where truthful, quantify achievements with "
            "numbers, outcomes or measurable results."
        )

    if not recommendations:
        recommendations.append(
            "Your resume contains the main detected sections. "
            "Tailor its wording to the target job."
        )

    ai_feedback = ""
    ai_error = ""

    if use_ai:
        try:
            with st.spinner(
                "Gemini AI is reviewing your resume..."
            ):
                ai_feedback = ask_gemini_ai(
                    resume_text,
                    job_description,
                    resume_skills,
                    missing,
                )

        except Exception as error:
            ai_error = str(error)

    st.session_state["resume_result"] = {
        "filename": uploaded_file.name,
        "resume_skills": resume_skills,
        "job_skills": job_skills,
        "matched": matched,
        "missing": missing,
        "quality": quality,
        "match_score": match_score,
        "ats_score": ats_score,
        "recommendations": recommendations,
        "ai_feedback": ai_feedback,
        "ai_error": ai_error,
    }


# ==================================================
# RESULTS
# ==================================================
if "resume_result" in st.session_state:

    result = st.session_state["resume_result"]
    quality = result["quality"]

    st.divider()
    st.header("Analysis Results")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-value">
                    {result["ats_score"]}%
                </div>
                <div class="metric-label">
                    ATS-Style Estimate
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:
        value = (
            f'{result["match_score"]}%'
            if result["match_score"] is not None
            else "N/A"
        )

        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-value">{value}</div>
                <div class="metric-label">Job Skill Match</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col3:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-value">
                    {len(result["matched"])}
                </div>
                <div class="metric-label">Matched Skills</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col4:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-value">
                    {quality["words"]}
                </div>
                <div class="metric-label">Resume Words</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.progress(
        max(0.0, min(1.0, result["ats_score"] / 100))
    )

    st.caption(
        "This is a heuristic ATS-style estimate, not an "
        "official ATS score or hiring decision."
    )

    st.subheader("Matched Skills")

    if result["matched"]:
        st.markdown(
            "".join(
                f'<span class="skill-tag">{skill}</span>'
                for skill in result["matched"]
            ),
            unsafe_allow_html=True
        )
    else:
        st.info("No matching skills were detected.")

    st.subheader("Missing Job Skills")

    if result["missing"]:
        st.markdown(
            "".join(
                f'<span class="missing-tag">{skill}</span>'
                for skill in result["missing"]
            ),
            unsafe_allow_html=True
        )
    elif result["job_skills"]:
        st.success("All detected job skills were found.")
    else:
        st.info(
            "Paste a job description to compare required skills."
        )

    st.subheader("Resume Section Analysis")

    for section, present in quality["sections"].items():
        st.write(
            f'{"✅" if present else "⚠️"} {section}: '
            f'{"Detected" if present else "Not detected"}'
        )

    st.write(
        f'**Section coverage:** {quality["section_score"]}%'
    )

    st.subheader("Recommendations")

    for number, recommendation in enumerate(
        result["recommendations"], 1
    ):
        st.write(f"{number}. {recommendation}")

    st.subheader("Gemini AI Feedback")

    if result["ai_feedback"]:
        st.markdown(result["ai_feedback"])

    elif result["ai_error"]:
        st.error(
            "Gemini AI feedback could not be generated. "
            "The rule-based analysis is still available."
        )

        with st.expander("View technical details"):
            st.code(result["ai_error"])

    else:
        st.info(
            "Gemini feedback was not requested. "
            "Enable the checkbox and analyze again to use AI."
        )

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
        st.info("No skills from the current database were detected.")

    # REPORT
    report = f"""
DEV AI - RESUME ANALYZER
====================================

Resume: {result["filename"]}
ATS-style estimate: {result["ats_score"]}%
Job skill match: {result["match_score"]}
Resume word count: {quality["words"]}
Section coverage: {quality["section_score"]}%

MATCHED SKILLS
{", ".join(result["matched"]) or "None"}

MISSING JOB SKILLS
{", ".join(result["missing"]) or "None"}

RESUME SKILLS
{", ".join(result["resume_skills"]) or "None"}

RECOMMENDATIONS
{chr(10).join("- " + item for item in result["recommendations"])}

GEMINI AI FEEDBACK
{result["ai_feedback"] or result["ai_error"] or "Not available"}

DISCLAIMER
This is an automated estimate, not an official ATS score.
"""

    st.download_button(
        "Download Analysis Report",
        data=report,
        file_name="Dev_AI_Resume_Analysis.txt",
        mime="text/plain",
        use_container_width=True,
    )


# ==================================================
# FOOTER - FIXED HTML RENDERING
# ==================================================
st.markdown(
    """
    <div class="footer-box">
        <strong style="color:#f1f5f9;font-size:18px;">
            Dev AI
        </strong>

        <div style="
            color:#5eead4;
            letter-spacing:3px;
            margin-top:5px;
        ">
            RESUME ANALYZER
        </div>

        <p>Python · Streamlit · Google Gemini AI · PDF OCR</p>
        <p>Built by Dewansh Pandey</p>
    </div>
    """,
    unsafe_allow_html=True,
)
