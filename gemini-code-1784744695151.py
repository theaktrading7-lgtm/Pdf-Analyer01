import streamlit as st
import PyPDF2
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
import json
import io

# Initialize OpenAI Client (Requires OPENAI_API_KEY in environment variables)
client = OpenAI()

def extract_text_from_pdfs(pdf_files):
    text = ""
    for pdf in pdf_files:
        reader = PyPDF2.PdfReader(pdf)
        for page in reader.pages:
            if page.extract_text():
                text += page.extract_text() + "\n"
    return text

def generate_insights(text):
    prompt = f"""
    Analyze the following text extracted from multiple documents. 
    1. Identify common themes, patterns, and key data points.
    2. Generate a highly concise briefing note and a list of action items.
    Use simple, clear, and direct language.
    
    Text: {text[:75000]} # Truncating to avoid token limits, adjust as needed
    """
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3
    )
    return response.choices[0].message.content

def generate_slide_content(text):
    prompt = f"""
    Based on the following text, create an outline for a 15-slide presentation.
    Return ONLY a JSON array of 15 objects. Each object must have a "title" and a "content" (a list of 3-4 short, concise bullet points highlighting key data and action items).
    
    Text: {text[:75000]}
    """
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"}, # Ensure valid JSON
        temperature=0.3
    )
    # Parse the JSON response
    return json.loads(response.choices[0].message.content)

def create_ppt(slide_data):
    prs = Presentation()
    for slide_info in slide_data.get('slides', slide_data):
        slide = prs.slides.add_slide(prs.slide_layouts[1]) # Title and Content layout
        
        # Add Title
        title_shape = slide.shapes.title
        title_shape.text = slide_info['title']
        
        # Add Content
        body_shape = slide.placeholders[1]
        tf = body_shape.text_frame
        for point in slide_info['content']:
            p = tf.add_paragraph()
            p.text = point
            p.font.size = Pt(18)
            
    # Save to memory buffer
    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- Streamlit UI ---
st.set_page_config(page_title="PDF Analyzer & PPT Generator", layout="wide")
st.title("📄 Cross-PDF Analyzer & Slide Generator")

uploaded_files = st.file_uploader("Upload 5-6 PDF files", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("Analyze & Generate"):
        with st.spinner("Extracting text from PDFs..."):
            extracted_text = extract_text_from_pdfs(uploaded_files)
            
        with st.spinner("Analyzing themes and generating briefing note..."):
            insights = generate_insights(extracted_text)
            st.subheader("Briefing Note & Action Items")
            st.write(insights)
            
        with st.spinner("Generating 15-page Presentation..."):
            try:
                slide_json = generate_slide_content(extracted_text)
                ppt_file = create_ppt(slide_json)
                
                st.success("Analysis and Presentation Generation Complete!")
                st.download_button(
                    label="📥 Download 15-Slide Presentation (.pptx)",
                    data=ppt_file,
                    file_name="Executive_Summary.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"An error occurred while generating the presentation: {e}")