import streamlit as st
import os
import tempfile
import docx2txt
import PyPDF2
from io import BytesIO
from docx import Document
import re

# --- Helpers ---
def extract_text(file):
    text = ""
    if file.name.endswith('.docx'):
        temp_path = os.path.join(tempfile.gettempdir(), file.name)
        with open(temp_path, 'wb') as f:
            f.write(file.read())
        text = docx2txt.process(temp_path)

    elif file.name.endswith('.pdf'):
        pdf_reader = PyPDF2.PdfReader(file)
        for page in pdf_reader.pages:
            text += page.extract_text()

    elif file.name.endswith('.txt'):
        text = file.read().decode("utf-8")

    return text

def generate_standards_from_reference(text):
    pattern = r'\n?\d+(\.\d+)*\.?\s+[^\n]+'
    matches = re.findall(pattern, text)
    sections = [match.strip() for match in matches]
    return list(set(sections))

def evaluate_proposal(text, required_sections):
    lower_text = text.lower()
    section_results = {sec: sec.lower() in lower_text for sec in required_sections}

    # Budget Clarity Check
    budget_keywords = ['budget', 'cost', 'expenditure', 'financial plan', 'funds', 'cost breakdown']
    has_budget_section = any(bk in lower_text for bk in budget_keywords)
    has_budget_numbers = bool(re.search(r"\$\d+|\d{1,3}(,\d{3})*(\.\d+)?", text))
    budget_check = has_budget_section and has_budget_numbers

    # Timeline Check
    timeline_keywords = ['timeline', 'schedule', 'work plan', 'gantt chart']
    has_timeline_section = any(tk in lower_text for tk in timeline_keywords)
    has_dates_or_months = bool(re.search(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|\d{4})\b", lower_text))
    timeline_check = has_timeline_section or has_dates_or_months

    # Total score calculation
    section_score = sum(section_results.values())
    extra_checks = [budget_check, timeline_check]
    max_score = len(required_sections) + len(extra_checks)
    percentage_score = round((section_score + sum(extra_checks)) / max_score * 100, 2)

    # Recommendations
    missing_sections = [sec for sec, present in section_results.items() if not present]
    recommendations = []
    if not budget_check:
        recommendations.append("Include a clear budget section with numeric breakdown.")
    if not timeline_check:
        recommendations.append("Include a detailed timeline or schedule with months or dates.")
    if missing_sections:
        recommendations.append(f"Missing sections: {', '.join(missing_sections)}")

    return {
        'sections': section_results,
        'score': percentage_score,
        'budget_check': budget_check,
        'timeline_check': timeline_check,
        'recommendations': recommendations
    }

def create_word_report(evaluation):
    doc = Document()
    doc.add_heading("Proposal Evaluation Report", level=1)

    doc.add_heading("Section Check", level=2)
    for section, found in evaluation['sections'].items():
        doc.add_paragraph(f"{section}: {'Present' if found else 'Missing'}")

    doc.add_heading("Budget Clarity", level=2)
    doc.add_paragraph("Present and Clear" if evaluation['budget_check'] else "Missing or Unclear")

    doc.add_heading("Timeline Inclusion", level=2)
    doc.add_paragraph("Timeline Included" if evaluation['timeline_check'] else "Missing")

    doc.add_heading("Overall Score", level=2)
    doc.add_paragraph(f"{evaluation['score']}%")

    doc.add_heading("Recommendations", level=2)
    if evaluation['recommendations']:
        for rec in evaluation['recommendations']:
            doc.add_paragraph(f"- {rec}")
    else:
        doc.add_paragraph("All criteria met. Great job!")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- Streamlit UI ---
st.title("Proposal Evaluator Using Reference Document")

uploaded_format = st.file_uploader("Upload Organization's Standard Proposal", type=["pdf", "docx", "txt"])
uploaded_proposal = st.file_uploader("Upload Proposal to Evaluate", type=["pdf", "docx", "txt"])

evaluate = st.button("Evaluate Proposal")

if evaluate and uploaded_format and uploaded_proposal:
    st.success("Both documents uploaded successfully.")

    ref_text = extract_text(uploaded_format)
    prop_text = extract_text(uploaded_proposal)

    with st.spinner("Generating standards from reference and evaluating proposal..."):
        ref_sections = generate_standards_from_reference(ref_text)
        evaluation = evaluate_proposal(prop_text, ref_sections)

    st.subheader("Evaluation Results")

    st.write("### Section Check")
    for section, found in evaluation['sections'].items():
        st.write(f"- **{section}**: {'✅' if found else '❌'}")

    st.write("### Budget Clarity")
    st.write("✅ Present and Clear" if evaluation['budget_check'] else "❌ Missing or Unclear")

    st.write("### Timeline Inclusion")
    st.write("✅ Timeline Included" if evaluation['timeline_check'] else "❌ Missing")

    st.write(f"### Overall Score: **{evaluation['score']}%**")

    st.write("### Recommendations")
    if evaluation['recommendations']:
        for rec in evaluation['recommendations']:
            st.warning(rec)
    else:
        st.success("Your proposal aligns well with the uploaded reference!")

    word_buffer = create_word_report(evaluation)
    st.download_button(
        label="Download Evaluation Report (.docx)",
        data=word_buffer,
        file_name="proposal_evaluation.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
