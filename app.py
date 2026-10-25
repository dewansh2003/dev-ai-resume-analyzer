import streamlit as st
import re
from pypdf import PdfReader
from docx import Document

st.set_page_config(page_title="AI Resume Analyzer", page_icon="📄", layout="wide")

st.title("📄 AI Resume Analyzer")
st.write("Analyze your resume against a job description.")

def extract_text(uploaded_file):
    if uploaded_file.name.endswith(".pdf"):
        reader = PdfReader(uploaded_file)
        return "\n".join(page.extract_text() or "" for page in reader.pages)

    elif uploaded_file.name.endswith(".docx"):
        doc = Document(uploaded_file)
        return "\n".join(p.text for p in doc.paragraphs)

    return ""

resume_file = st.file_uploader(
    "Upload your Resume",
    type=["pdf", "docx"]
)

job_description = st.text_area(
    "Paste Job Description",
    height=200,
    placeholder="Paste the job description here..."
)

if st.button("Analyze Resume"):
    if resume_file is None or not job_description.strip():
        st.warning("Please upload your resume and paste the job description.")
    else:
        resume_text = extract_text(resume_file)

        if not resume_text.strip():
            st.error("Could not extract text. Please check your resume file.")
        else:
            skills = [
                "python", "javascript", "react", "node.js",
                "sql", "mongodb", "machine learning",
                "pandas", "numpy", "generative ai",
                "html", "css", "rest api", "git", "github",
                "flask", "django", "express.js"
            ]

            resume_lower = resume_text.lower()
            job_lower = job_description.lower()

            required_skills = [
                skill for skill in skills
                if re.search(re.escape(skill), job_lower)
            ]

            matched_skills = [
                skill for skill in required_skills
                if re.search(re.escape(skill), resume_lower)
            ]

            missing_skills = [
                skill for skill in required_skills
                if skill not in matched_skills
            ]

            score = (
                len(matched_skills) / len(required_skills) * 100
                if required_skills else 0
            )

            st.subheader("📊 Analysis Results")
            st.metric("Job Match Score", f"{score:.1f}%")
            st.write(f"**Required skills detected:** {len(required_skills)}")
            st.write(f"**Matching skills:** {len(matched_skills)}")

            st.success(
                "Matched Skills: " +
                (", ".join(matched_skills) if matched_skills else "None detected")
            )

            st.warning(
                "Missing Skills: " +
                (", ".join(missing_skills) if missing_skills else "None detected")
            )

            st.info(
                "This score measures keyword overlap, not a complete ATS score."
            )