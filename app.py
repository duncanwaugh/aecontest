import os
import streamlit as st
from pptx import Presentation
from docx import Document
from dotenv import load_dotenv
from openai import OpenAI

# Load environment variables
load_dotenv()

# Set up OpenAI client
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# Extract text from PPTX slides
def extract_text_from_pptx(file_path):
    prs = Presentation(file_path)
    text = ''
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text += shape.text_frame.text + "\n"
    return text

# Generate summary and extracted sections using OpenAI GPT
def summarize_and_extract(text):
    prompt = f"""
    From the following incident investigation presentation text, create a concise Serious Event Lessons Learned document containing ONLY:
    
    - Brief, clear title/headline
    - Event Summary (brief, readable paragraph)
    - Clearly listed contributing factors
    - Clearly listed lessons learned

    DO NOT INCLUDE sensitive operational details or any unnecessary internal information.

    Here is the presentation text:
    {text}
    
    Format:
    Title:
    Event Summary:
    Contributing Factors:
    - factor 1
    - factor 2
    Lessons Learned:
    - lesson 1
    - lesson 2
    """

    response = client.chat.completions.create(
        model="gpt-4-turbo",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        max_tokens=700,
    )

    return response.choices[0].message.content.strip()

# Generate Word Document from the summarized content
def create_lessons_learned_doc(content, template_path, output_path):
    doc = Document(template_path)

    sections = {}
    current_section = None
    for line in content.split('\n'):
        line = line.strip()
        if line.endswith(':'):
            current_section = line[:-1]
            sections[current_section] = []
        elif current_section:
            sections[current_section].append(line)

    # Replace placeholders with extracted content
    for para in doc.paragraphs:
        text = para.text
        if '{{Title}}' in text:
            para.text = sections.get('Title', [''])[0]
        elif '{{Event_Summary}}' in text:
            para.text = ' '.join(sections.get('Event Summary', []))
        elif '{{Contributing_Factors}}' in text:
            para.text = '\n'.join(sections.get('Contributing Factors', []))
        elif '{{Lessons_Learned}}' in text:
            para.text = '\n'.join(sections.get('Lessons Learned', []))

    doc.save(output_path)

# Streamlit UI
st.title('🦺 Serious Event Lessons Learned Generator')

uploaded_file = st.file_uploader("Upload Executive Review PPTX", type=['pptx'])

if uploaded_file:
    input_filepath = f'input/{uploaded_file.name}'
    with open(input_filepath, 'wb') as f:
        f.write(uploaded_file.getbuffer())

    with st.spinner("Extracting content from presentation..."):
        pptx_text = extract_text_from_pptx(input_filepath)

    with st.spinner("Generating Lessons Learned summary..."):
        generated_content = summarize_and_extract(pptx_text)

    st.success("✅ Generation Complete!")

    st.text_area("📝 Generated Content:", generated_content, height=300)

    template_path = 'templates/lessons_learned_template.docx'
    output_path = 'generated_lessons_learned.docx'

    create_lessons_learned_doc(generated_content, template_path, output_path)

    with open(output_path, "rb") as file:
        st.download_button(
            label="📥 Download Generated Document",
            data=file,
            file_name=output_path,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )