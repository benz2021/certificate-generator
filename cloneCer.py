import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from io import BytesIO
import zipfile
import re

# --- นำเข้าไลบรารีสำหรับคลิกหาพิกัด ---
try:
    from streamlit_image_coordinates import streamlit_image_coordinates
except ImportError:
    st.error("⚠️ ไม่พบไลบรารี streamlit-image-coordinates")
    st.info("กรุณาเปิด Terminal แล้วพิมพ์: pip install streamlit-image-coordinates")
    st.stop()

# ==========================================
# 🛠️ HELPER FUNCTIONS
# ==========================================
def get_font(size):
    """โหลดฟอนต์จากไฟล์ที่ผู้ใช้อัปโหลด"""
    if 'font_bytes' in st.session_state and st.session_state.font_bytes:
        try:
            return ImageFont.truetype(BytesIO(st.session_state.font_bytes), size)
        except Exception as e:
            st.error(f"ฟอนต์มีปัญหา: {e}")
            return ImageFont.load_default()
    else:
        return ImageFont.load_default()

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', str(name)).strip() or "certificate"

def render_certificate(template_img, texts, row_data=None):
    img = template_img.copy()
    if img.mode != 'RGB':
        img = img.convert('RGB')
    
    draw = ImageDraw.Draw(img)
    
    for txt in texts:
        if txt['type'] == 'static':
            content = txt['text']
        else:
            if row_data and txt['column'] in row_data:
                val = row_data[txt['column']]
                content = str(val) if pd.notna(val) else ""
            else:
                content = "ตัวอย่างข้อมูล"
        
        if not content: continue
            
        font = get_font(txt['size'])
        draw.text((txt['x'], txt['y']), content, fill=txt['color'], font=font, anchor="mm")
    
    return img

# ==========================================
# 🎨 UI - STREAMLIT APP
# ==========================================
st.set_page_config(page_title="Auto Cert Pro", layout="wide")

# ตั้งค่า Session State
if "click_x" not in st.session_state: st.session_state.click_x = 0
if "click_y" not in st.session_state: st.session_state.click_y = 0
if 'texts' not in st.session_state: st.session_state.texts = []

st.title(" Certificate Generator")

# --- SIDEBAR ---
with st.sidebar:
    st.header("1️⃣ อัปโหลดไฟล์")
    
    # 1. Template
    template_file = st.file_uploader("1. พื้นหลังเกียรติบัตร (JPG/PNG)", type=['jpg', 'jpeg', 'png'])
    if template_file:
        st.session_state.template = Image.open(template_file)

    # 2. Font (แก้ไขให้โหลดจากผู้ใช้)
    font_file = st.file_uploader("2. ฟอนต์ภาษาไทย (.ttf)", type=['ttf'])
    if font_file:
        st.session_state.font_bytes = font_file.getvalue()
        st.success("✅ โหลดฟอนต์สำเร็จ")
    else:
        st.warning("⚠️ แนะนำให้อัปโหลดไฟล์ .ttf เพื่อให้ปรับขนาด/แสดงภาษาไทยได้")

    # 3. Data
    data_file = st.file_uploader("3. รายชื่อ (ไม่เกิน 150 รายชื่อ ไฟล์ Excel/CSV)  ", type=['xlsx', 'xls', 'csv'])
    if data_file:
        if data_file.name.endswith('.csv'):
            st.session_state.data = pd.read_csv(data_file)
        else:
            st.session_state.data = pd.read_excel(data_file)

if 'template' not in st.session_state:
    st.info("👈 กรุณาอัปโหลด 'ต้นแบบเกียรติบัติ' ที่เมนูด้านซ้ายเพื่อเริ่มต้น")
    st.stop()

# --- MAIN AREA ---
st.header("2️⃣ กำหนดตำแหน่ง ")

col_img, col_form = st.columns([1.5, 1])

with col_img:
    st.markdown("**🖱️ คลิกลงบนรูปภาพเพื่อดึงพิกัด(คลิกบนรูป หรือ ระบุ X และ Y)")
    
    # คำนวณการย่อภาพ เพื่อไม่ให้ภาพใหญ่ล้นจอ
    original_w, original_h = st.session_state.template.size
    display_w = 700 # ความกว้างสูงสุดที่แสดงบนหน้าเว็บ (ปรับได้)
    
    if original_w > display_w:
        ratio = original_w / display_w
        display_img = st.session_state.template.resize((display_w, int(original_h / ratio)))
    else:
        ratio = 1.0
        display_img = st.session_state.template

    # แสดงรูปภาพและจับพิกัดการคลิก
    coords = streamlit_image_coordinates(display_img, key="target_clicker")
    
    # คำนวณพิกัดกลับไปเป็นขนาดรูปจริง
    if coords is not None:
        st.session_state.click_x = int(coords['x'] * ratio)
        st.session_state.click_y = int(coords['y'] * ratio)

with col_form:
    st.markdown("**📝 ตั้งค่าข้อความ**")
    with st.form("add_text_form", clear_on_submit=False):
        t_type = st.radio("ชนิดข้อมูล", ["พิมพ์เอง", "ดึงจาก Excel"], horizontal=True)
        
        t_val, t_col = "", ""
        if "พิมพ์เอง" in t_type:
            t_val = st.text_input("ระบุข้อความ")
        else:
            if 'data' in st.session_state:
                t_col = st.selectbox("เลือกหัวข้อ (Column)", st.session_state.data.columns)
            else:
                st.warning("อัปโหลด Excel ก่อนครับ")

        c1, c2 = st.columns(2)
        # ดึงพิกัดที่คลิกมาใส่ให้อัตโนมัติ
        x_pos = c1.number_input("แกน X (คลิกรูปเพื่อเปลี่ยน)", value=st.session_state.click_x)
        y_pos = c2.number_input("แกน Y (คลิกรูปเพื่อเปลี่ยน)", value=st.session_state.click_y)
        
        f_size = st.slider("ขนาดฟอนต์", 10, 500, value=60)
        f_color = st.color_picker("เลือกสี", value="#000000")
        
        if st.form_submit_button("➕ แทรกข้อความลงเกียรติบัตร"):
            # เช็คว่าผู้ใช้อัปโหลดฟอนต์หรือยัง ถ้ายังให้แจ้งเตือน
            if 'font_bytes' not in st.session_state:
                st.warning("อย่าลืมอัปโหลดฟอนต์ก่อนนะครับ ไม่งั้นจะปรับขนาดไม่ได้!")
            
            st.session_state.texts.append({
                'type': 'static' if "พิมพ์เอง" in t_type else 'excel',
                'text': t_val, 'column': t_col,
                'x': x_pos, 'y': y_pos,
                'size': f_size, 'color': f_color
            })
            st.rerun()

st.markdown("---")

# --- พรีวิวและจัดการข้อความ ---
st.header("3️⃣ ดูตัวอย่าง (Preview)")

if st.session_state.texts:
    for i, t in enumerate(st.session_state.texts):
        lbl = t['text'] if t['type'] == 'static' else f"จาก: {t['column']}"
        cols = st.columns([4, 1])
        cols[0].write(f"**{i+1}. {lbl}** | ขนาด: {t['size']} | พิกัด: ({t['x']}, {t['y']})")
        if cols[1].button("🗑️ ลบ", key=f"del_{i}"):
            st.session_state.texts.pop(i)
            st.rerun()

    # ดูตัวอย่างจาก Excel
    preview_row = None
    if 'data' in st.session_state:
        row_idx = st.number_input("ดูตัวอย่างแถวที่:", 0, max(0, len(st.session_state.data)-1), 0)
        preview_row = st.session_state.data.iloc[row_idx].to_dict()
    
    # สร้างรูปพรีวิว (ย่อให้หน้ากว้าง 700px เพื่อไม่ให้ล้นจอ)
    preview_img = render_certificate(st.session_state.template, st.session_state.texts, preview_row)
    st.image(preview_img, width=700)
else:
    st.info("ตั้งค่าข้อความด้านบนก่อนครับ")

# --- Export ---
if 'data' in st.session_state and st.session_state.texts:
    st.markdown("---")
    st.header("4️⃣ สร้างไฟล์ทั้งหมด")
    filename_col = st.selectbox("เลือกคอลัมน์ชื่อไฟล์", st.session_state.data.columns)
    
    if st.button("🚀 ดาวน์โหลด (ZIP)", type="primary"):
        zip_buffer = BytesIO()
        with st.spinner("กำลังปั่นเกียรติบัตร..."):
            with zipfile.ZipFile(zip_buffer, 'w') as zf:
                for idx, row in st.session_state.data.iterrows():
                    final_img = render_certificate(st.session_state.template, st.session_state.texts, row.to_dict())
                    img_io = BytesIO()
                    # บันทึกภาพขนาดเต็ม!
                    final_img.save(img_io, format="PNG")
                    clean_name = sanitize_filename(row[filename_col])
                    zf.writestr(f"{clean_name}.png", img_io.getvalue())
            
            st.success("เรียบร้อย!")
            st.download_button("📥 ดาวน์โหลดไฟล์ ZIP", zip_buffer.getvalue(), "certificates.zip", "application/zip")