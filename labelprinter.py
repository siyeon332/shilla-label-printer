import io
import os
import urllib.request
import streamlit as st
import barcode
from barcode.writer import ImageWriter

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# PDF 생성을 위한 reportlab 라이브러리
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import code128
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- 폰트 다운로드 안전 시스템 (Noto Sans KR Bold) ---
FONT_NAME = "NotoSansKR-Bold"
FONT_FILE = "NotoSansKR-Bold.ttf"

@st.cache_resource
def load_korean_font():
    """인터넷에서 무조건 굵은 본고딕 Bold 폰트를 긁어옵니다. (타임아웃 및 예비 링크 추가)"""
    if os.path.exists(FONT_FILE):
        try:
            pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_FILE))
            return True
        except Exception:
            pass

    # 1차 구글 깃허브 다운로드 시도
    urls = [
        "https://github.com/google/fonts/raw/main/ofl/notosanskr/NotoSansKR%5Bwght%5D.ttf",
        "https://fonts.gstatic.com/s/notosanskr/v36/PbykFmXiEBPT4ITbgNA5CgmsclYdxg.ttf" # 예비 구글 폰트 주소
    ]
    
    for url in urls:
        try:
            # 10초 대기 설정으로 안정적으로 가져옵니다
            response = urllib.request.urlopen(url, timeout=10)
            with open(FONT_FILE, "wb") as f:
                f.write(response.read())
            pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_FILE))
            return True
        except Exception:
            continue
    return False

FONT_LOADED = load_korean_font()

# 실패 시 한글 깨짐 방지 예비
if not FONT_LOADED:
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    pdfmetrics.registerFont(UnicodeCIDFont("HYGothic-Medium"))
    FONT_NAME = "HYGothic-Medium"


# 페이지 기본 설정
st.set_page_config(
    page_title="신라면세점 라벨 원스톱 생성기",
    page_icon="🏷️",
    layout="centered"
)

if "barcode_bytes" not in st.session_state:
    st.session_state.barcode_bytes = None
if "barcode_val" not in st.session_state:
    st.session_state.barcode_val = ""

# --- [기능 1] 바코드 이미지 생성 (PPTX용) ---
def generate_code128_image(value: str) -> io.BytesIO:
    Code128 = barcode.get_barcode_class("code128")
    fp = io.BytesIO()
    barcode_image = Code128(value, writer=ImageWriter())
    barcode_image.write(
        fp,
        options={
            "module_width": 0.6, # PPTX용 바코드도 조금 더 크게 키움
            "module_height": 28,
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
def create_sh
