import streamlit as st
import pandas as pd
from pypdf import PdfReader
from pptx import Presentation
import json
import io
import openai

st.set_page_config(page_title="present.id - AI Presentation Generator", page_icon="📊", layout="wide")

st.title("📊 present.id")
st.caption("Ubah Dokumen (PDF/Excel) Menjadi Presentasi PPTX dalam Sekejap")

# --- AMBIL API KEY DARI STREAMLIT SECRETS ---
# Jika secrets belum diset, fallback ke input manual di sidebar
api_key = st.secrets.get("OPENAI_API_KEY", "")

if not api_key:
    with st.sidebar:
        st.header("Pengaturan")
        api_key = st.text_input("Masukkan OpenAI API Key:", type="password")

def extract_text_from_pdf(pdf_file):
    reader = PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    return text

def extract_data_from_excel(excel_file):
    xls = pd.ExcelFile(excel_file)
    summary_text = ""
    for sheet in xls.sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet)
        summary_text += f"\n--- Sheet: {sheet} ---\n"
        summary_text += df.head(10).to_string()
    return summary_text

def generate_slides_json(content, key):
    client = openai.OpenAI(api_key=key)
    
    prompt = f"""
    Kamu adalah pembuat presentasi profesional.
    Analisis dokumen/data berikut dan buatlah struktur presentasi 5-7 slide.
    
    Format output HANYA JSON bertipe object dengan key 'slides':
    {{
        "slides": [
            {{
                "title": "Judul Slide 1",
                "subtitle": "Ringkasan singkat",
                "bullets": ["Poin 1", "Poin 2", "Poin 3"]
            }}
        ]
    }}

    Dokumen/Data:
    {content[:4000]}
    """
    
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are a helpful assistant that outputs JSON."},
            {"role": "user", "content": prompt}
        ],
        response_format={"type": "json_object"}
    )
    
    return json.loads(response.choices[0].message.content)

def create_pptx(slides_data):
    prs = Presentation()
    
    slides_list = slides_data.get("slides", []) if isinstance(slides_data, dict) else slides_data
    
    for slide_info in slides_list:
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        
        slide.shapes.title.text = slide_info.get("title", "Slide Tanpa Judul")
        
        tf = slide.placeholders[1].text_frame
        tf.text = slide_info.get("subtitle", "")
        
        for bullet in slide_info.get("bullets", []):
            p = tf.add_paragraph()
            p.text = bullet
            p.level = 1
            
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

# UI Utama
uploaded_file = st.file_uploader("Upload file PDF atau Excel (.xlsx)", type=["pdf", "xlsx", "xls"])

if uploaded_file and api_key:
    if st.button("🚀 Buat Presentasi Now"):
        with st.spinner("AI sedang membaca dokumen dan mengolah slide..."):
            try:
                if uploaded_file.name.endswith('.pdf'):
                    extracted_text = extract_text_from_pdf(uploaded_file)
                else:
                    extracted_text = extract_data_from_excel(uploaded_file)
                    
                slides_data = generate_slides_json(extracted_text, api_key)
                pptx_file = create_pptx(slides_data)
                
                st.success("Presentasi berhasil dibuat!")
                st.download_button(
                    label="📥 Download File PPTX",
                    data=pptx_file,
                    file_name=f"{uploaded_file.name.split('.')[0]}_present_id.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"Terjadi kesalahan: {str(e)}")
elif uploaded_file and not api_key:
    st.warning("Silakan masukkan API Key di sidebar atau atur di Streamlit Secrets.")
