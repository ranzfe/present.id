import streamlit as st
import pandas as pd
from pypdf import PdfReader
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import json
import io
import openai

# --- 1. PAGE CONFIGURATION & CUSTOM CSS (PRESENTATIONS.AI STYLE) ---
st.set_page_config(
    page_title="present.id — AI Presentation Generator",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS untuk gaya UI Dark Modern & Futuristic ala Presentations.ai
st.markdown("""
<style>
    /* Global Styling */
    .stApp {
        background-color: #0d0f17;
        color: #e2e8f0;
    }
    
    /* Header Gradient Banner */
    .main-header {
        background: linear-gradient(135deg, #1e1b4b 0%, #312e81 50%, #4c1d95 100%);
        padding: 2.5rem;
        border-radius: 20px;
        border: 1px solid rgba(255, 255, 255, 0.1);
        text-align: center;
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.5);
        margin-bottom: 2rem;
    }
    
    .main-header h1 {
        font-size: 2.8rem;
        font-weight: 800;
        background: linear-gradient(to right, #38bdf8, #818cf8, #c084fc);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }
    
    .main-header p {
        color: #94a3b8;
        font-size: 1.1rem;
    }

    /* Slide Card Canvas Preview */
    .slide-card {
        background: #1e293b;
        border: 1px solid #334155;
        border-radius: 16px;
        padding: 2rem;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.4);
        margin-bottom: 1.5rem;
    }
    
    .slide-number {
        font-size: 0.85rem;
        font-weight: 700;
        color: #38bdf8;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        margin-bottom: 0.5rem;
    }
    
    .slide-title-display {
        font-size: 1.8rem;
        font-weight: 700;
        color: #f8fafc;
        margin-bottom: 0.5rem;
    }
    
    .slide-subtitle-display {
        font-size: 1.1rem;
        color: #cbd5e1;
        margin-bottom: 1.2rem;
        font-style: italic;
    }
    
    .bullet-point {
        background: rgba(255, 255, 255, 0.03);
        border-left: 3px solid #6366f1;
        padding: 0.6rem 1rem;
        border-radius: 0 8px 8px 0;
        margin-bottom: 0.5rem;
        color: #e2e8f0;
    }

    /* Primary Accent Buttons */
    .stButton>button {
        background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%);
        color: white;
        border: none;
        font-weight: 600;
        border-radius: 10px;
        padding: 0.6rem 1.2rem;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #4f46e5 0%, #4338ca 100%);
        box-shadow: 0 4px 12px rgba(99, 102, 241, 0.4);
        transform: translateY(-1px);
    }
</style>
""", unsafe_allow_html=True)

# --- 2. SESSION STATE MANAGEMENT ---
if "slides_data" not in st.session_state:
    st.session_state.slides_data = None
if "current_slide_idx" not in st.session_state:
    st.session_state.current_slide_idx = 0

# API Key Handler
api_key = st.secrets.get("OPENAI_API_KEY", "")

# --- 3. SIDEBAR CONTROLS ---
with st.sidebar:
    st.title("⚡ present.id Studio")
    st.caption("AI Presentation Engine v2.0")
    st.divider()
    
    if not api_key:
        api_key = st.text_input("🔑 OpenAI API Key:", type="password", help="Masukkan API key OpenAI kamu di sini jika tidak menggunakan Streamlit Secrets.")
    
    st.subheader("🎨 Pilihan Tema & Style")
    theme_choice = st.selectbox(
        "Pilih Palette Warna Slide:",
        ["Modern Dark", "Corporate Blue", "Emerald Executive", "Minimal Light"]
    )
    
    num_slides = st.slider("Jumlah Slide Target:", min_value=3, max_value=10, value=5)
    
    st.divider()
    st.markdown("### 💡 Tips Penggunaan")
    st.info("Upload PDF ringkasan bisnis atau data Excel penjualan kamu, lalu klik **Generate Presentation**.")

# --- 4. HEADER HERO SECTION ---
st.markdown("""
<div class="main-header">
    <h1>Presentations.ai Powered by present.id</h1>
    <p>Ubah Dokumen PDF & Excel Menjadi Deck Presentasi Interaktif Dalam Sekejap</p>
</div>
""", unsafe_allow_html=True)

# --- 5. HELPER FUNCTIONS ---
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
        summary_text += f"Kolom: {', '.join(df.columns.astype(str))}\n"
        summary_text += df.head(10).to_string() + "\n"
    return summary_text

def generate_slides_with_ai(content, key, slide_count):
    client = openai.OpenAI(api_key=key)
    
    prompt = f"""
    Kamu adalah Head of Presentation Design profesional ala McKinsey dan Presentations.ai.
    Analisis dokumen/data berikut dan rancang presentasi sebanyak {slide_count} slide.
    
    Setiap slide harus memiliki:
    - title: Judul singkat, padat, berpengaruh
    - subtitle: Subjudul konteks
    - bullets: 3-4 poin fakta/data penting
    - speaker_notes: Catatan penjelasan untuk presenter ketika membawakan slide ini.
    
    Kembalikan respon HANYA JSON bertipe object dengan key 'slides':
    {{
        "slides": [
            {{
                "title": "...",
                "subtitle": "...",
                "bullets": ["...", "..."],
                "speaker_notes": "..."
            }}
        ]
    }}

    Data Input:
    {content[:5000]}
    """
    
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are a world-class presentation AI that outputs strict JSON."},
            {"role": "user", "content": prompt}
        ],
        response_format={"type": "json_object"}
    )
    
    return json.loads(response.choices[0].message.content)

def create_custom_pptx(slides_list, theme):
    prs = Presentation()
    # Aspect Ratio 16:9
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    # Theme Color Schemes
    colors = {
        "Modern Dark": {"bg": RGBColor(15, 23, 42), "title": RGBColor(248, 250, 252), "sub": RGBColor(148, 163, 184), "bullet": RGBColor(226, 232, 240)},
        "Corporate Blue": {"bg": RGBColor(241, 245, 249), "title": RGBColor(30, 58, 138), "sub": RGBColor(71, 85, 105), "bullet": RGBColor(51, 65, 85)},
        "Emerald Executive": {"bg": RGBColor(6, 78, 59), "title": RGBColor(236, 253, 245), "sub": RGBColor(110, 231, 183), "bullet": RGBColor(209, 250, 229)},
        "Minimal Light": {"bg": RGBColor(255, 255, 255), "title": RGBColor(15, 23, 42), "sub": RGBColor(100, 116, 139), "bullet": RGBColor(51, 65, 85)}
    }
    palette = colors.get(theme, colors["Modern Dark"])

    for item in slides_list:
        slide = prs.slides.add_slide(blank_layout)
        
        # Background Fill
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = palette["bg"]
        
        # Add Title Box
        title_box = slide.shapes.add_textbox(Inches(1), Inches(0.8), Inches(11.333), Inches(1.2))
        tf = title_box.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = item.get("title", "Slide Title")
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = palette["title"]
        
        p_sub = tf.add_paragraph()
        p_sub.text = item.get("subtitle", "")
        p_sub.font.size = Pt(18)
        p_sub.font.color.rgb = palette["sub"]
        
        # Add Bullets Box
        bullet_box = slide.shapes.add_textbox(Inches(1), Inches(2.3), Inches(11.333), Inches(4.5))
        btf = bullet_box.text_frame
        btf.word_wrap = True
        
        for idx, bullet in enumerate(item.get("bullets", [])):
            bp = btf.add_paragraph() if idx > 0 else btf.paragraphs[0]
            bp.text = f"•  {bullet}"
            bp.font.size = Pt(20)
            bp.font.color.rgb = palette["bullet"]
            bp.space_after = Pt(14)
            
        # Add Speaker Notes
        if item.get("speaker_notes"):
            notes_slide = slide.notes_slide
            text_frame = notes_slide.notes_text_frame
            text_frame.text = item.get("speaker_notes")
            
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

# --- 6. WORKFLOW & FILE UPLOADER ---
col_upload, col_prompt = st.columns([1, 1])

with col_upload:
    uploaded_file = st.file_uploader(
        "📁 Upload Dokumen Sumber (PDF / Excel)",
        type=["pdf", "xlsx", "xls"],
        help="Sistem akan otomatis mengekstrak konten teks atau tabel untuk dijadikan materi slide."
    )

with col_prompt:
    custom_prompt = st.text_area(
        "📝 Instruksi Tambahan untuk AI (Opsional):",
        placeholder="Contoh: Fokuskan pada target pendapatan kuartal ketiga dan berikan nada bicara resmi.",
        height=100
    )

if st.button("✨ Generate Presentation with AI", use_container_width=True):
    if not uploaded_file:
        st.warning("Silakan upload file PDF atau Excel terlebih dahulu!")
    elif not api_key:
        st.error("Masukkan OpenAI API Key di sidebar atau Streamlit Secrets!")
    else:
        with st.spinner("AI sedang membaca dokumen & menyusun tata letak slide..."):
            try:
                # 1. Ekstraksi Dokumen
                if uploaded_file.name.endswith('.pdf'):
                    extracted_content = extract_text_from_pdf(uploaded_file)
                else:
                    extracted_content = extract_data_from_excel(uploaded_file)
                
                if custom_prompt:
                    extracted_content += f"\n\nInstruksi Pengguna Tambahan: {custom_prompt}"
                    
                # 2. Olah AI
                ai_result = generate_slides_with_ai(extracted_content, api_key, num_slides)
                st.session_state.slides_data = ai_result.get("slides", [])
                st.session_state.current_slide_idx = 0
                st.success("Presentasi berhasil dibuat! Silakan tinjau dan edit di bawah ini.")
            except Exception as e:
                st.error(f"Gagal memproses presentasi: {str(e)}")

# --- 7. INTERACTIVE CANVAS & INLINE EDITOR ---
if st.session_state.slides_data:
    st.divider()
    
    slides = st.session_state.slides_data
    total_slides = len(slides)
    
    st.markdown("### 🖥️ Interactive Slide Editor & Preview")
    
    # Navigation Bar Preview
    nav_col1, nav_col2, nav_col3 = st.columns([1, 3, 1])
    
    with nav_col1:
        if st.button("⬅️ Slide Sebelumnya", disabled=(st.session_state.current_slide_idx == 0)):
            st.session_state.current_slide_idx -= 1
            st.rerun()
            
    with nav_col2:
        st.markdown(
            f"<h4 style='text-align: center; color: #38bdf8;'>Slide {st.session_state.current_slide_idx + 1} / {total_slides}</h4>", 
            unsafe_allow_html=True
        )
        
    with nav_col3:
        if st.button("Slide Selanjutnya ➡️", disabled=(st.session_state.current_slide_idx == total_slides - 1)):
            st.session_state.current_slide_idx += 1
            st.rerun()

    # Get Active Slide Data
    curr_idx = st.session_state.current_slide_idx
    slide_item = slides[curr_idx]

    tab_preview, tab_edit = st.tabs(["👁️ Canvas Live Preview", "✏️ Inline Editor (Ubah Teks)"])

    with tab_preview:
        # Visual Slide Render
        bullets_html = "".join([f'<div class="bullet-point">• {b}</div>' for b in slide_item.get("bullets", [])])
        
        st.markdown(f"""
        <div class="slide-card">
            <div class="slide-number">Slide {curr_idx + 1}</div>
            <div class="slide-title-display">{slide_item.get('title', '')}</div>
            <div class="slide-subtitle-display">{slide_item.get('subtitle', '')}</div>
            <div style="margin-top: 1.5rem;">
                {bullets_html}
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Speaker Notes Box
        with st.expander("🎙️ AI Speaker Notes (Catatan Presenter)"):
            st.write(slide_item.get("speaker_notes", "Tidak ada catatan untuk slide ini."))

    with tab_edit:
        st.markdown("#### Edit Isi Slide Secara Lintas Teks")
        new_title = st.text_input("Judul Slide:", value=slide_item.get("title", ""), key=f"title_{curr_idx}")
        new_subtitle = st.text_input("Subjudul Slide:", value=slide_item.get("subtitle", ""), key=f"sub_{curr_idx}")
        
        bullets_str = "\n".join(slide_item.get("bullets", []))
        new_bullets_raw = st.text_area("Poin-Poin Bullets (Satu poin per baris):", value=bullets_str, height=120, key=f"bullets_{curr_idx}")
        new_notes = st.text_area("Catatan Pembicara:", value=slide_item.get("speaker_notes", ""), key=f"notes_{curr_idx}")
        
        if st.button("💾 Simpan Perubahan Slide Ini"):
            st.session_state.slides_data[curr_idx]["title"] = new_title
            st.session_state.slides_data[curr_idx]["subtitle"] = new_subtitle
            st.session_state.slides_data[curr_idx]["bullets"] = [b.strip() for b in new_bullets_raw.split("\n") if b.strip()]
            st.session_state.slides_data[curr_idx]["speaker_notes"] = new_notes
            st.success("Perubahan slide berhasil disimpan!")
            st.rerun()

    # --- 8. EXPORT SECTION ---
    st.divider()
    st.markdown("### 📥 Unduh Hasil Presentasi")
    
    pptx_bytes = create_custom_pptx(st.session_state.slides_data, theme_choice)
    
    st.download_button(
        label=f"🚀 Download File PPTX ({theme_choice})",
        data=pptx_bytes,
        file_name=f"present_id_{uploaded_file.name.split('.')[0] if uploaded_file else 'export'}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )
