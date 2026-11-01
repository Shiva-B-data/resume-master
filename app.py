import streamlit as st
import spacy
import pypdf
import io
import os
import re
from langchain_groq import ChatGroq
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
from dotenv import load_dotenv
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

load_dotenv()

@st.cache_resource
def load_nlp():
    return spacy.load("en_core_web_sm")

nlp = load_nlp()

def get_llm():
    return ChatGroq(
        model="llama-3.3-70b-versatile",
        temperature=0.2,
        api_key=os.getenv("GROQ_API_KEY")
    )

def extract_text(file):
    file.seek(0)
    filename = file.name.lower()
    if filename.endswith(".pdf"):
        reader = pypdf.PdfReader(io.BytesIO(file.read()))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    elif filename.endswith(".docx"):
        doc = Document(io.BytesIO(file.read()))
        return "\n".join(p.text for p in doc.paragraphs)
    return ""

def extract_keywords(text):
    doc = nlp(text.lower())
    keywords = set()
    for token in doc:
        if token.pos_ in ["NOUN", "PROPN"] and not token.is_stop and len(token.text) > 2:
            keywords.add(token.lemma_.lower())
    for ent in doc.ents:
        if ent.label_ in ["ORG", "PRODUCT", "GPE"]:
            keywords.add(ent.text.lower())
    return keywords

def extract_tech_keywords(text):
    tech_patterns = [
        r'\b(python|java|javascript|sql|scala|kotlin|swift|typescript)\b',
        r'\b(pytorch|tensorflow|scikit.learn|keras|xgboost|hugging.face|langchain|spacy|nltk)\b',
        r'\b(aws|azure|gcp|docker|kubernetes|spark|kafka|airflow|hadoop|databricks|snowflake)\b',
        r'\b(postgresql|mongodb|mysql|redis|elasticsearch|pinecone|chromadb|weaviate|mlflow)\b',
        r'\b(mlops|ci.cd|git|fastapi|flask|streamlit|react)\b',
        r'\b(llm|rag|nlp|bert|gpt|transformers|embeddings|fine.tun|prompt.engineer)\b',
        r'\b(machine learning|deep learning|neural network|computer vision|natural language)\b',
    ]
    found = set()
    text_lower = text.lower()
    for pattern in tech_patterns:
        matches = re.findall(pattern, text_lower)
        found.update(matches)
    return found

def calculate_ats_score(resume_text, jd_text):
    resume_lower = resume_text.lower()

    resume_kw = extract_keywords(resume_text)
    jd_kw = extract_keywords(jd_text)
    kw_match_rate = len(resume_kw.intersection(jd_kw)) / len(jd_kw) if jd_kw else 0

    resume_tech = extract_tech_keywords(resume_text)
    jd_tech = extract_tech_keywords(jd_text)
    tech_match_rate = len(resume_tech.intersection(jd_tech)) / len(jd_tech) if jd_tech else kw_match_rate

    sections = ["experience", "education", "skill"]
    section_score = sum(1 for s in sections if s in resume_lower) / len(sections)

    numbers = re.findall(r'\d+', resume_text)
    quant_score = min(len(numbers) / 10, 1.0)

    raw = (kw_match_rate * 0.40 + tech_match_rate * 0.35 +
           section_score * 0.15 + quant_score * 0.10)
    return round(min(2.0 + (raw * 8.0), 10.0), 1)

def get_keyword_analysis(resume_text, jd_text):
    resume_kw = extract_keywords(resume_text)
    jd_kw = extract_keywords(jd_text)
    resume_tech = extract_tech_keywords(resume_text)
    jd_tech = extract_tech_keywords(jd_text)
    matching = {k for k in resume_kw.intersection(jd_kw) if len(k) > 3}
    missing = {k for k in jd_kw.difference(resume_kw) if len(k) > 3}
    matching_tech = resume_tech.intersection(jd_tech)
    missing_tech = jd_tech.difference(resume_tech)
    return matching, missing, matching_tech, missing_tech

def get_ats_report(resume_text):
    resume_lower = resume_text.lower()
    report = {"passed": [], "failed": [], "warnings": []}

    if "summary" in resume_lower or "objective" in resume_lower:
        report["passed"].append("Professional summary section found")
    else:
        report["failed"].append("Missing professional summary — add one at the top")

    if "experience" in resume_lower or "work history" in resume_lower:
        report["passed"].append("Work experience section found")
    else:
        report["failed"].append("Work experience section not detected")

    if "education" in resume_lower:
        report["passed"].append("Education section found")
    else:
        report["failed"].append("Education section not detected")

    if "skill" in resume_lower:
        report["passed"].append("Skills section found")
    else:
        report["failed"].append("Skills section missing — add a dedicated skills section")

    if "@" in resume_text:
        report["passed"].append("Email address detected")
    else:
        report["failed"].append("No email address found — ATS cannot contact you")

    if "linkedin" in resume_lower:
        report["passed"].append("LinkedIn profile included")
    else:
        report["warnings"].append("No LinkedIn URL — strongly recommended")

    numbers = re.findall(r'\d+', resume_text)
    if len(numbers) >= 5:
        report["passed"].append(f"Quantified achievements found ({len(numbers)} numbers detected)")
    else:
        report["failed"].append("Too few numbers — add metrics like percentages, counts, dollar amounts")

    word_count = len(resume_text.split())
    if 300 <= word_count <= 700:
        report["passed"].append(f"Good resume length ({word_count} words)")
    elif word_count < 300:
        report["warnings"].append(f"Resume may be too short ({word_count} words)")
    else:
        report["warnings"].append(f"Resume may be too long ({word_count} words) — consider trimming")

    if "•" in resume_text or "-" in resume_text:
        report["passed"].append("Bullet points detected — good for ATS parsing")
    else:
        report["warnings"].append("Use bullet points for experience sections")

    return report

def generate_tailored_resume(resume_text, jd_text, llm, custom_instructions=""):
    jd_tech = extract_tech_keywords(jd_text)
    jd_keywords_str = ", ".join(list(jd_tech)[:20]) if jd_tech else "relevant skills from JD"

    custom_section = f"""
CUSTOM INSTRUCTIONS FROM USER — FOLLOW THESE EXACTLY:
{custom_instructions}
""" if custom_instructions.strip() else ""

    prompt = ChatPromptTemplate.from_template("""
You are an expert ATS resume optimizer. Rewrite the candidate's resume to score 9/10 or higher.

EXISTING RESUME:
{resume}

JOB DESCRIPTION:
{jd}

KEY SKILLS FROM JD TO INCORPORATE: {keywords}

{custom_instructions}

STRICT RULES:
1. NEVER invent companies, degrees, or dates not in the original resume
2. KEEP all original companies, job titles, dates and education exactly as they are
3. If user says "add skill X" — add it to the skills section
4. If user says "add certification X" — add it to certifications section
5. If user says "I also have experience with X" — add it naturally to relevant bullets
6. Rewrite every bullet with strong action verbs and exact JD keywords
7. Every bullet must have at least one metric or number
8. Summary must directly reference the target role and key JD requirements
9. Skills must be categorized and include ALL relevant JD skills the candidate has
10. Keep bullets concise — max 2 lines each

OUTPUT — use EXACTLY this format:

NAME: [full name]
EMAIL: [email]
PHONE: [phone if available]
LINKEDIN: [linkedin if available]
LOCATION: [city, state if available]

SUMMARY
[3 sentences tailored to this exact role]

EXPERIENCE
[Job Title] | [Company] | [Start] – [End]
- [action verb + achievement + JD keyword + metric]
- [action verb + achievement + JD keyword + metric]
- [action verb + achievement + JD keyword + metric]

[repeat for ALL roles]

EDUCATION
[Degree] | [University] | [Year][GPA if available]

SKILLS
Languages: [list]
ML/AI: [list]
Cloud & DevOps: [list]
Data Engineering: [list]
Tools & Frameworks: [list]

CERTIFICATIONS
- [each certification]
""")

    chain = prompt | llm | StrOutputParser()
    return chain.invoke({
        "resume": resume_text,
        "jd": jd_text,
        "keywords": jd_keywords_str,
        "custom_instructions": custom_section
    })

def generate_cover_letter(resume_text, jd_text, llm):
    prompt = ChatPromptTemplate.from_template("""
You are an expert cover letter writer. Write a powerful personalized cover letter.

CANDIDATE RESUME:
{resume}

JOB DESCRIPTION:
{jd}

RULES:
- 3 paragraphs only — Opening, Body (3 achievements matching JD), Closing
- Sound confident and human — not robotic
- Reference the specific company and role from the JD
- Only use real experience from the resume
- Under 280 words total
- No subject line, no date, no address

OUTPUT:
Dear Hiring Manager,

[Opening — enthusiasm, specific role and company name, one strong hook]

[Body — 3 specific achievements matching JD requirements with metrics]

[Closing — confidence, request interview, thank them]

Sincerely,
[Candidate Name]
""")
    chain = prompt | llm | StrOutputParser()
    return chain.invoke({"resume": resume_text, "jd": jd_text})

def add_section_header(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(3)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '1a1a2e')
    pBdr.append(bottom)
    pPr.append(pBdr)
    run = p.add_run(title.upper())
    run.font.name = 'Arial'
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)

def create_resume_doc(tailored_text):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(10)

    lines = tailored_text.strip().split("\n")
    name, contact_parts = "", []
    section_headers = ["SUMMARY", "EXPERIENCE", "EDUCATION", "SKILLS", "CERTIFICATIONS", "PROJECTS"]
    skip_prefixes = ["NAME:", "EMAIL:", "PHONE:", "LINKEDIN:", "LOCATION:"]
    skill_categories = ["Languages", "ML/AI", "Cloud & DevOps", "Data Engineering", "Tools & Frameworks", "Tools", "Frameworks"]

    for line in lines:
        line = line.strip()
        if line.startswith("NAME:"):
            name = line.replace("NAME:", "").strip()
        elif line.startswith("EMAIL:"):
            v = line.replace("EMAIL:", "").strip()
            if v: contact_parts.append(v)
        elif line.startswith("PHONE:"):
            v = line.replace("PHONE:", "").strip()
            if v: contact_parts.append(v)
        elif line.startswith("LINKEDIN:"):
            v = line.replace("LINKEDIN:", "").strip()
            if v: contact_parts.append(v)
        elif line.startswith("LOCATION:"):
            v = line.replace("LOCATION:", "").strip()
            if v: contact_parts.append(v)

    if name:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(name)
        run.font.name = 'Arial'
        run.font.size = Pt(18)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)

    if contact_parts:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run("  |  ".join(contact_parts))
        run.font.name = 'Arial'
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    for line in lines:
        line = line.strip()
        if not line or any(line.startswith(px) for px in skip_prefixes):
            continue

        if any(line.upper() == h for h in section_headers):
            add_section_header(doc, line)

        elif line.startswith("•") or line.startswith("-"):
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.left_indent = Inches(0.25)
            run = p.add_run(line.lstrip("•- ").strip())
            run.font.name = 'Arial'
            run.font.size = Pt(10)

        elif "|" in line and not any(line.startswith(px) for px in skip_prefixes):
            parts = line.split("|")
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(7)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(parts[0].strip())
            run.font.name = 'Arial'
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)
            for part in parts[1:]:
                run2 = p.add_run("  |  " + part.strip())
                run2.font.name = 'Arial'
                run2.font.size = Pt(10)
                run2.font.bold = False
                run2.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

        elif any(line.startswith(cat + ":") for cat in skill_categories):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            parts = line.split(":", 1)
            run1 = p.add_run(parts[0].strip() + ": ")
            run1.font.name = 'Arial'
            run1.font.size = Pt(10)
            run1.font.bold = True
            run1.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)
            if len(parts) > 1:
                run2 = p.add_run(parts[1].strip())
                run2.font.name = 'Arial'
                run2.font.size = Pt(10)
                run2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(line)
            run.font.name = 'Arial'
            run.font.size = Pt(10)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def create_cover_letter_doc(cover_text):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.1)
        section.right_margin = Inches(1.1)
    for line in cover_text.strip().split("\n"):
        line = line.strip()
        if not line:
            doc.add_paragraph()
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(line)
        run.font.name = 'Arial'
        run.font.size = Pt(11)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ─── UI ─────────────────────────────────────────────────

st.set_page_config(page_title="Resume Master", page_icon="🏆", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; background: #0d0d1a; }
.hero { background: linear-gradient(135deg, #0d0d1a 0%, #1a1a3e 50%, #0f3460 100%); padding: 3.5rem 2rem; border-radius: 20px; text-align: center; margin-bottom: 2rem; border: 1px solid #2d2d5e; }
.hero-badge { background: #e94560; color: white; border-radius: 20px; padding: 4px 14px; font-size: 0.75rem; font-weight: 700; letter-spacing: 1px; display: inline-block; margin-bottom: 1rem; }
.hero h1 { color: white; font-size: 3rem; font-weight: 800; margin: 0; line-height: 1.1; }
.hero h1 span { color: #e94560; }
.hero p { color: #a8b2d8; font-size: 1.1rem; margin-top: 0.8rem; }
.hero-features { display: flex; justify-content: center; gap: 2rem; margin-top: 1.5rem; flex-wrap: wrap; }
.hero-feature { color: #a8b2d8; font-size: 0.9rem; }
.hero-feature span { color: #4ade80; margin-right: 6px; }
.step-card { background: #1e1e2e; border: 1px solid #2d2d44; border-radius: 14px; padding: 1.5rem; height: 100%; }
.step-num { background: #e94560; color: white; border-radius: 50%; width: 28px; height: 28px; display: inline-flex; align-items: center; justify-content: center; font-weight: 700; font-size: 0.85rem; margin-bottom: 0.5rem; }
.step-title { color: white; font-weight: 600; font-size: 1rem; margin-bottom: 0.3rem; }
.step-desc { color: #a8b2d8; font-size: 0.85rem; }
.score-ring-container { text-align: center; padding: 2rem; background: linear-gradient(135deg, #1a1a2e, #0f3460); border-radius: 16px; border: 2px solid #2d2d5e; }
.score-big { font-size: 4.5rem; font-weight: 800; line-height: 1; }
.meter-wrap { background: #2d2d44; border-radius: 50px; height: 14px; margin: 0.8rem 0; overflow: hidden; }
.kw-pill-green { display: inline-block; background: #1a3a2a; color: #4ade80; border: 1px solid #4ade80; border-radius: 20px; padding: 3px 10px; margin: 3px; font-size: 0.8rem; }
.kw-pill-red { display: inline-block; background: #3a1a1a; color: #f87171; border: 1px solid #f87171; border-radius: 20px; padding: 3px 10px; margin: 3px; font-size: 0.8rem; }
.passed-item { background: #1a3a2a; border-left: 3px solid #4ade80; border-radius: 6px; padding: 0.5rem 1rem; margin: 0.3rem 0; color: #a8d8a8; font-size: 0.88rem; }
.failed-item { background: #3a1a1a; border-left: 3px solid #f87171; border-radius: 6px; padding: 0.5rem 1rem; margin: 0.3rem 0; color: #f8a8a8; font-size: 0.88rem; }
.warn-item { background: #2a2a1a; border-left: 3px solid #fbbf24; border-radius: 6px; padding: 0.5rem 1rem; margin: 0.3rem 0; color: #d8d0a8; font-size: 0.88rem; }
.tip-item { background: #1a2a3a; border-left: 3px solid #60a5fa; border-radius: 6px; padding: 0.6rem 1rem; margin: 0.4rem 0; color: #a8d8f8; font-size: 0.9rem; }
.success-item { background: #1a3a2a; border-left: 3px solid #4ade80; border-radius: 6px; padding: 0.6rem 1rem; margin: 0.4rem 0; color: #a8d8a8; font-size: 0.9rem; }
.section-label { color: #e94560; font-size: 0.85rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; margin-bottom: 0.6rem; }
.card { background: #1e1e2e; border: 1px solid #2d2d44; border-radius: 12px; padding: 1.2rem; margin-bottom: 0.8rem; }
.compare-box { background: #1e1e2e; border: 1px solid #2d2d44; border-radius: 12px; padding: 1.5rem; text-align: center; }
.compare-label { color: #a8b2d8; font-size: 0.8rem; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 0.5rem; }
.compare-score { font-size: 3rem; font-weight: 800; }
.compare-arrow { font-size: 2rem; color: #4ade80; display: flex; align-items: center; justify-content: center; height: 100%; padding-top: 1.5rem; }
div[data-testid="stButton"] button { background: linear-gradient(135deg, #e94560, #c73652); color: white; border: none; border-radius: 10px; padding: 0.7rem 2rem; font-weight: 700; font-size: 1rem; width: 100%; }
.stTextArea textarea { background: #1e1e2e !important; color: #e0e0e0 !important; border: 1px solid #2d2d44 !important; border-radius: 10px !important; }
div[data-testid="stFileUploader"] { background: #1e1e2e; border: 2px dashed #2d2d44; border-radius: 12px; padding: 1rem; }
</style>
""", unsafe_allow_html=True)

# ── Hero ─────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <div class="hero-badge">100% FREE — NO SIGN UP REQUIRED</div>
    <h1>Resume <span>Master</span></h1>
    <p>The AI-powered resume tailor that gets you more interviews</p>
    <div class="hero-features">
        <div class="hero-feature"><span>✓</span> ATS Score Analysis</div>
        <div class="hero-feature"><span>✓</span> AI Resume Tailoring</div>
        <div class="hero-feature"><span>✓</span> Cover Letter Generator</div>
        <div class="hero-feature"><span>✓</span> Full ATS Audit Report</div>
        <div class="hero-feature"><span>✓</span> Before vs After Score</div>
        <div class="hero-feature"><span>✓</span> Custom Instructions</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Steps ────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
for col, (num, title, desc) in zip([c1,c2,c3,c4], [
    ("1","Upload Resume","PDF or Word document"),
    ("2","Paste Job Description","Any role from any job board"),
    ("3","Analyze & Tailor","AI rewrites your resume to match"),
    ("4","Download & Apply","Professional Word doc ready to send"),
]):
    with col:
        st.markdown(f'<div class="step-card"><div class="step-num">{num}</div><div class="step-title">{title}</div><div class="step-desc">{desc}</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("---")
st.markdown("<br>", unsafe_allow_html=True)

# ── Inputs ───────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="section-label">📄 Your Resume</div>', unsafe_allow_html=True)
    resume_file = st.file_uploader("Upload", type=["pdf","docx"], label_visibility="collapsed")
    if resume_file:
        st.markdown(f'<div class="success-item">✅ <strong>{resume_file.name}</strong> uploaded successfully</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="tip-item">📎 Upload your resume in PDF or Word (.docx) format</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="section-label">💼 Job Description</div>', unsafe_allow_html=True)
    jd_text = st.text_area("JD", height=220,
        placeholder="Paste the full job description here — include all requirements, responsibilities and qualifications...",
        label_visibility="collapsed")
    if jd_text:
        st.markdown(f'<div class="success-item">✅ Job description loaded — {len(jd_text.split())} words</div>', unsafe_allow_html=True)

# ── Custom Instructions ───────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown('<div class="section-label">⚙️ Custom Instructions (Optional)</div>', unsafe_allow_html=True)
st.markdown('<div class="tip-item">💡 Tell the AI exactly what to add, change, or emphasize in your resume</div>', unsafe_allow_html=True)
custom_instructions = st.text_area(
    "custom",
    height=120,
    placeholder="""Examples of what you can type here:
- Add skill: MLflow, Pinecone, Terraform
- I also have experience with FastAPI — add it to my skills and mention it in my current role bullets
- Add certification: AWS Certified Solutions Architect
- Make the summary sound more senior-level and leadership-focused
- Emphasize my experience with real-time data pipelines more strongly
- Add: I have 2 years of experience with Kubernetes in production""",
    label_visibility="collapsed"
)
if custom_instructions.strip():
    st.markdown(f'<div class="success-item">✅ Custom instructions added — AI will follow these when building your resume</div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Buttons ──────────────────────────────────────────────
b1, b2, b3 = st.columns(3)
with b1:
    analyze_btn = st.button("🔍 Analyze ATS Score")
with b2:
    build_btn = st.button("✨ Build Tailored Resume")
with b3:
    cover_btn = st.button("📝 Generate Cover Letter")

# ── Analyze ──────────────────────────────────────────────
if analyze_btn:
    if not resume_file or not jd_text:
        st.error("Please upload your resume and paste a job description first")
    else:
        with st.spinner("Running ATS analysis..."):
            resume_text = extract_text(resume_file)
            score = calculate_ats_score(resume_text, jd_text)
            matching, missing, matching_tech, missing_tech = get_keyword_analysis(resume_text, jd_text)
            ats_report = get_ats_report(resume_text)

        score_pct = score * 10
        if score >= 7.5:
            color, verdict, emoji = "#4ade80", "Strong ATS Match", "🟢"
        elif score >= 5.5:
            color, verdict, emoji = "#fbbf24", "Moderate ATS Match", "🟡"
        else:
            color, verdict, emoji = "#f87171", "Low ATS Match", "🔴"

        st.markdown("---")
        st.markdown("### 📊 ATS Analysis Results")
        st.markdown("<br>", unsafe_allow_html=True)

        r1, r2, r3 = st.columns([1, 1.5, 1])
        with r2:
            st.markdown(f"""
            <div class="score-ring-container">
                <div style="color:#a8b2d8;font-size:0.8rem;text-transform:uppercase;letter-spacing:2px;">ATS Match Score</div>
                <div class="score-big" style="color:{color};">{score}<span style="font-size:2rem;color:#a8b2d8;">/10</span></div>
                <div style="color:{color};font-weight:600;margin-top:0.3rem;">{emoji} {verdict}</div>
                <div class="meter-wrap">
                    <div style="height:100%;width:{score_pct}%;background:linear-gradient(90deg,{color},{color}99);border-radius:50px;"></div>
                </div>
                <div style="display:flex;justify-content:space-between;color:#a8b2d8;font-size:0.75rem;"><span>0</span><span>5</span><span>10</span></div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Smart verdict banner
        if score >= 7.5:
            st.markdown(f"""
            <div style="background:linear-gradient(135deg,#1a3a2a,#0d2a1a);border:2px solid #4ade80;border-radius:14px;padding:1.5rem;text-align:center;margin-bottom:1rem;">
                <div style="font-size:2rem;">🎉</div>
                <div style="color:#4ade80;font-size:1.3rem;font-weight:700;margin-top:0.5rem;">Your Resume Looks Great — No Major Changes Needed!</div>
                <div style="color:#a8d8a8;margin-top:0.5rem;">Your resume scores <strong>{score}/10</strong> against this job description. You are a strong candidate — apply with confidence. You can still use Build Tailored Resume for minor fine-tuning if you want.</div>
            </div>
            """, unsafe_allow_html=True)
        elif score >= 5.5:
            st.markdown(f"""
            <div style="background:linear-gradient(135deg,#2a2a1a,#1a1a0d);border:2px solid #fbbf24;border-radius:14px;padding:1.5rem;text-align:center;margin-bottom:1rem;">
                <div style="font-size:2rem;">📝</div>
                <div style="color:#fbbf24;font-size:1.3rem;font-weight:700;margin-top:0.5rem;">Decent Match — Small Tweaks Will Boost Your Chances</div>
                <div style="color:#d8d0a8;margin-top:0.5rem;">Your resume scores <strong>{score}/10</strong>. You have a solid foundation. Adding the missing keywords below will significantly increase your interview chances. Use <strong>Build Tailored Resume</strong> to do this automatically in 30 seconds.</div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div style="background:linear-gradient(135deg,#2a1a1a,#1a0d0d);border:2px solid #f87171;border-radius:14px;padding:1.5rem;text-align:center;margin-bottom:1rem;">
                <div style="font-size:2rem;">⚠️</div>
                <div style="color:#f87171;font-size:1.3rem;font-weight:700;margin-top:0.5rem;">Low Match — Your Resume Needs Tailoring for This Role</div>
                <div style="color:#d8a8a8;margin-top:0.5rem;">Your resume scores <strong>{score}/10</strong> against this specific role. Your experience may be relevant but the keywords don't align with the JD. Click <strong>Build Tailored Resume</strong> below to automatically fix this — it takes 30 seconds.</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Stats row
        s1, s2, s3, s4 = st.columns(4)
        with s1:
            st.markdown(f'<div class="compare-box"><div class="compare-label">ATS Score</div><div class="compare-score" style="color:{color};">{score}/10</div></div>', unsafe_allow_html=True)
        with s2:
            st.markdown(f'<div class="compare-box"><div class="compare-label">Keywords Matched</div><div class="compare-score" style="color:#4ade80;">{len(matching)}</div></div>', unsafe_allow_html=True)
        with s3:
            st.markdown(f'<div class="compare-box"><div class="compare-label">Tech Skills Matched</div><div class="compare-score" style="color:#60a5fa;">{len(matching_tech)}</div></div>', unsafe_allow_html=True)
        with s4:
            st.markdown(f'<div class="compare-box"><div class="compare-label">Keywords Missing</div><div class="compare-score" style="color:#f87171;">{len(missing)}</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ATS Audit
        st.markdown('<div class="section-label">🎯 Full ATS Audit Report</div>', unsafe_allow_html=True)
        a1, a2 = st.columns(2)
        with a1:
            st.markdown("**✅ Passed**")
            for item in ats_report["passed"]:
                st.markdown(f'<div class="passed-item">✓ {item}</div>', unsafe_allow_html=True)
        with a2:
            st.markdown("**❌ Failed / ⚠️ Warnings**")
            for item in ats_report["failed"]:
                st.markdown(f'<div class="failed-item">✗ {item}</div>', unsafe_allow_html=True)
            for item in ats_report["warnings"]:
                st.markdown(f'<div class="warn-item">⚠ {item}</div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Keywords
        kw1, kw2 = st.columns(2)
        with kw1:
            st.markdown('<div class="section-label">✅ Keywords You Have</div>', unsafe_allow_html=True)
            pills = "".join(f'<span class="kw-pill-green">{k}</span>' for k in sorted(matching)[:30])
            st.markdown(f'<div class="card">{pills or "No matching keywords found"}</div>', unsafe_allow_html=True)
            if matching_tech:
                st.markdown('<div class="section-label" style="margin-top:1rem;">🔧 Tech Skills Matched</div>', unsafe_allow_html=True)
                pills2 = "".join(f'<span class="kw-pill-green">{k}</span>' for k in sorted(matching_tech))
                st.markdown(f'<div class="card">{pills2}</div>', unsafe_allow_html=True)
        with kw2:
            st.markdown('<div class="section-label">⚡ Keywords to Add</div>', unsafe_allow_html=True)
            pills = "".join(f'<span class="kw-pill-red">{k}</span>' for k in sorted(missing)[:30])
            st.markdown(f'<div class="card">{pills or "Great — no missing keywords!"}</div>', unsafe_allow_html=True)
            if missing_tech:
                st.markdown('<div class="section-label" style="margin-top:1rem;">🔧 Missing Tech Skills</div>', unsafe_allow_html=True)
                pills2 = "".join(f'<span class="kw-pill-red">{k}</span>' for k in sorted(missing_tech))
                st.markdown(f'<div class="card">{pills2}</div>', unsafe_allow_html=True)

# ── Build Resume ─────────────────────────────────────────
if build_btn:
    if not resume_file or not jd_text:
        st.error("Please upload your resume and paste a job description first")
    else:
        with st.spinner("🤖 AI is tailoring your resume for maximum ATS score — 20-30 seconds..."):
            try:
                resume_text = extract_text(resume_file)
                original_score = calculate_ats_score(resume_text, jd_text)
                llm = get_llm()
                tailored_text = generate_tailored_resume(resume_text, jd_text, llm, custom_instructions)
                tailored_score = calculate_ats_score(tailored_text, jd_text)
                improvement = round(tailored_score - original_score, 1)
                word_file = create_resume_doc(tailored_text)

                st.markdown("---")
                st.markdown("### ✨ Your Tailored Resume is Ready!")
                st.markdown("<br>", unsafe_allow_html=True)

                c1, c2, c3 = st.columns([2, 1, 2])
                orig_color = "#f87171" if original_score < 5 else "#fbbf24" if original_score < 7 else "#4ade80"
                tail_color = "#f87171" if tailored_score < 5 else "#fbbf24" if tailored_score < 7 else "#4ade80"
                with c1:
                    st.markdown(f'<div class="compare-box"><div class="compare-label">Before — Original Resume</div><div class="compare-score" style="color:{orig_color};">{original_score}/10</div></div>', unsafe_allow_html=True)
                with c2:
                    st.markdown(f'<div class="compare-arrow">→</div>', unsafe_allow_html=True)
                with c3:
                    st.markdown(f'<div class="compare-box"><div class="compare-label">After — Tailored Resume</div><div class="compare-score" style="color:{tail_color};">{tailored_score}/10</div><div style="color:#4ade80;font-size:0.9rem;margin-top:0.3rem;">▲ +{improvement} improvement</div></div>', unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button(
                    label="📥 Download Tailored Resume (.docx)",
                    data=word_file,
                    file_name="Resume_Master_Tailored.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown('<div class="section-label">📋 Preview</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="card" style="color:#e0e0e0;white-space:pre-wrap;font-family:monospace;font-size:0.82rem;max-height:500px;overflow-y:auto;">{tailored_text}</div>', unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Error: {str(e)}")

# ── Cover Letter ─────────────────────────────────────────
if cover_btn:
    if not resume_file or not jd_text:
        st.error("Please upload your resume and paste a job description first")
    else:
        with st.spinner("✍️ Writing your personalized cover letter..."):
            try:
                resume_text = extract_text(resume_file)
                llm = get_llm()
                cover_text = generate_cover_letter(resume_text, jd_text, llm)
                cover_file = create_cover_letter_doc(cover_text)

                st.markdown("---")
                st.markdown("### 📝 Your Cover Letter is Ready!")
                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button(
                    label="📥 Download Cover Letter (.docx)",
                    data=cover_file,
                    file_name="Resume_Master_CoverLetter.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown('<div class="section-label">📋 Preview</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="card" style="color:#e0e0e0;white-space:pre-wrap;font-family:monospace;font-size:0.9rem;">{cover_text}</div>', unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Error: {str(e)}")

# ── Footer ───────────────────────────────────────────────
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("""
<div style="text-align:center;color:#555577;font-size:0.85rem;border-top:1px solid #2d2d44;padding-top:1.5rem;">
    <strong style="color:#e94560;">Resume Master</strong> — Free AI Resume Tailor for Every Job Seeker<br>
    Built with ❤️ using Groq · LangChain · spaCy · Scikit-learn · Streamlit
</div>
""", unsafe_allow_html=True)
