import io
import streamlit as st
import barcode
from barcode.writer import ImageWriter

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# PDF 생성을 위한 reportlab 라이브러리 (이외의 추가 폰트 설정 없이 깔끔하게 그리기 위함)
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import code128
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 페이지 기본 설정
st.set_page_config(
    page_title="신라면세점 라벨 원스톱 생성기",
    page_icon="🏷️",
    layout="centered"
)

# 세션 상태 초기화
if "barcode_bytes" not in st.session_state:
    st.session_state.barcode_bytes = None
if "barcode_val" not in st.session_state:
    st.session_state.barcode_val = ""

# --- [기능 1] 내부 메모리에서 바코드(Code 128) 이미지 파일 생성 (PPTX용) ---
def generate_code128_image(value: str) -> io.BytesIO:
    Code128 = barcode.get_barcode_class("code128")
    fp = io.BytesIO()
    barcode_image = Code128(value, writer=ImageWriter())
    barcode_image.write(
        fp,
        options={
            "module_width": 0.5,
            "module_height": 24,
            "font_size": 14,
            "text_distance": 10,
            "quiet_zone": 6,
            "dpi": 300,
            "write_text": True,
        }
    )
    fp.seek(0)
    return fp

# --- [기능 2] PPTX 라벨 일괄 생성 ---
def create_shilla_pptx(brand_name: str, label_type: str, total_qty: int, barcode_bytes: io.BytesIO) -> io.BytesIO:
    prs = Presentation()
    prs.slide_width = Inches(11.69)   # A4 가로
    prs.slide_height = Inches(8.27)
    blank_layout = prs.slide_layouts[6]

    import tempfile
    import os
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        tmp.write(barcode_bytes.getvalue())
        tmp_barcode_path = tmp.name

    try:
        for i in range(1, total_qty + 1):
            slide = prs.slides.add_slide(blank_layout)
            
            # 1. 브랜드명
            brand_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.8), Inches(10.69), Inches(1.5))
            tf_brand = brand_box.text_frame
            tf_brand.word_wrap = True
            p_brand = tf_brand.paragraphs[0]
            p_brand.text = brand_name
            p_brand.alignment = PP_ALIGN.CENTER
            p_brand.font.name = "Arial"
            p_brand.font.size = Pt(64)
            p_brand.font.bold = True
            
            # 2. PLT/BOX 번호 (예: PLT NO. 5-1)
            plt_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.2), Inches(10.69), Inches(1.2))
            tf_plt = plt_box.text_frame
            tf_plt.word_wrap = True
            p_plt = tf_plt.paragraphs[0]
            p_plt.text = f"{label_type} NO. {total_qty}-{i}"
            p_plt.alignment = PP_ALIGN.CENTER
            p_plt.font.name = "Arial"
            p_plt.font.size = Pt(54)
            p_plt.font.bold = True
            
            # 3. 바코드 이미지
            img_width = Inches(8.5)
            img_height = Inches(4.0)
            img_left = (prs.slide_width - img_width) / 2
            img_top = Inches(3.7)
            slide.shapes.add_picture(tmp_barcode_path, img_left, img_top, width=img_width, height=img_height)
            
    finally:
        if os.path.exists(tmp_barcode_path):
            os.unlink(tmp_barcode_path)

    out_pptx = io.BytesIO()
    prs.save(out_pptx)
    out_pptx.seek(0)
    return out_pptx

# --- [기능 3] PDF 라벨 일괄 생성 ---
def create_shilla_pdf(brand_name: str, label_type: str, total_qty: int, barcode_value: str) -> io.BytesIO:
    # PDF를 메모리 내에 생성
    pdf_buffer = io.BytesIO()
    
    # 가로형 A4 크기 (841.89 * 595.27 포인트)
    pagesize = (841.89, 595.27)
    c = canvas.Canvas(pdf_buffer, pagesize=pagesize)
    
    for i in range(1, total_qty + 1):
        # 1. 브랜드명 작성
        c.setFont("Helvetica-Bold", 60)
        c.drawCentredString(pagesize[0]/2.0, 480, brand_name)
        
        # 2. PLT / BOX 번호 작성 (예: PLT NO. 5-1)
        c.setFont("Helvetica-Bold", 50)
        c.drawCentredString(pagesize[0]/2.0, 390, f"{label_type} NO. {total_qty}-{i}")
        
        # 3. Code 128 바코드 그리기
        # Reportlab 자체 바코드 위젯 사용
        barcode_obj = code128.Code128(barcode_value, barWidth=1.8, barHeight=140, humanReadable=True)
        # 바코드 밑에 표시되는 글자 폰트 설정
        barcode_obj.fontName = "Helvetica"
        barcode_obj.fontSize = 15
        
        # 바코드를 정중앙에 배치하기 위한 위치 계산
        barcode_width = barcode_obj.width
        x_pos = (pagesize[0] - barcode_width) / 2.0
        y_pos = 120
        
        barcode_obj.drawOn(c, x_pos, y_pos)
        
        # 다음 페이지 추가 (마지막 페이지가 아니면)
        if i < total_qty:
            c.showPage()
            
    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer


# --- 🖥️ 더욱 스마트해진 초간편 웹 화면 (Streamlit UI) ---
st.title("🏷️ 신라면세점 부착라벨 원스톱 생성기")
st.markdown("입력하신 수량과 동일하게 파일 페이지가 일괄 구성됩니다.")
st.write("---")

# 레이아웃 구성
col1, col2 = st.columns(2)

with col1:
    # 1. 라벨 종류 선택 (PLT 또는 BOX)
    label_type = st.radio("1. 라벨 형태 선택", ["PLT", "BOX"], horizontal=True)
    brand_name = st.text_input("2. 브랜드명 입력", value="아비브(ABIB)")

with col2:
    # 2. 수량 하나만 입력 받기
    total_qty = st.number_input(f"3. 총 {label_type} 수량 (N)", min_value=1, value=5, step=1)
    barcode_val = st.text_input("4. 바코드 번호 입력", placeholder="예: 7851260079")

st.write(" ")
st.write(" ")

# 메인 연산 버튼
btn_generate_all = st.button("🚀 신라면세점 규격 라벨 파일 생성", type="primary", use_container_width=True)

if btn_generate_all:
    if not barcode_val.strip():
        st.error("바코드 번호를 입력해 주세요!")
    else:
        with st.spinner("라벨 내부 디자인 및 바코드 생성 처리 중..."):
            try:
                # 1. 공통 바코드 생성
                barcode_bytes = generate_code128_image(barcode_val.strip())
                st.session_state.barcode_bytes = barcode_bytes
                st.session_state.barcode_val = barcode_val.strip()
                
                # 2. PPTX 파일 생성
                pptx_output = create_shilla_pptx(
                    brand_name=brand_name.strip(),
                    label_type=label_type,
                    total_qty=total_qty,
                    barcode_bytes=barcode_bytes
                )
                
                # 3. PDF 파일 생성
                pdf_output = create_shilla_pdf(
                    brand_name=brand_name.strip(),
                    label_type=label_type,
                    total_qty=total_qty,
                    barcode_value=barcode_val.strip()
                )
                
                st.success("🎉 라벨 파일들이 성공적으로 디자인되었습니다!")
                
                # 생성된 바코드 미리보기
                st.image(barcode_bytes, caption=f"자동 생성된 Code 128 바코드 ({barcode_val})", width=320)
                
                st.write("---")
                st.markdown("### 📥 원하시는 포맷의 버튼을 눌러 다운로드하세요")
                
                dl_col1, dl_col2 = st.columns(2)
                
                with dl_col1:
                    # PPTX 다운로드 버튼
                    st.download_button(
                        label="📥 PPTX 파일 다운로드",
                        data=pptx_output,
                        file_name=f"{brand_name}_라벨_{label_type}{total_qty}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )
                    
                with dl_col2:
                    # PDF 다운로드 버튼
                    st.download_button(
                        label="📥 PDF 파일 다운로드",
                        data=pdf_output,
                        file_name=f"{brand_name}_라벨_{label_type}{total_qty}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                
            except Exception as e:
                st.error(f"라벨 자동 제작 실패: {e}")