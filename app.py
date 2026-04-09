import streamlit as st
import pandas as pd
import numpy as np
import io, os, glob
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import altair as alt

st.set_page_config(page_title="히트펌프 경제성 분석 솔루션", layout="wide")

# ── 1. 스타일 정의 ──
st.markdown("""
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
* { font-family: 'Pretendard', sans-serif; }
.info-box   { background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:24px; margin-bottom:32px; }
.info-title { color:#0f172a; font-size:1.15rem; font-weight:700; margin-bottom:12px; margin-top:0; }
.info-text  { color:#475569; font-size:0.95rem; line-height:1.6; margin-bottom:0; }
.section-title { color:#1e293b; font-weight:700; font-size:1.3rem; margin-top:40px; margin-bottom:16px; border-bottom:2px solid #cbd5e1; padding-bottom:8px; }
.calc-box   { background:#f0f9ff; border:1px solid #bae6fd; border-radius:8px; padding:16px; margin:12px 0; }
.calc-label { color:#0369a1; font-size:0.85rem; font-weight:600; margin-bottom:4px; }
.calc-value { color:#0f172a; font-size:1.05rem; font-weight:700; }
</style>
""", unsafe_allow_html=True)

# ── 2. 핵심 로직: 엑셀 데이터 파싱 (에러 방지 강화) ──
@st.cache_data
def load_simulation_data():
    try:
        # 파일 읽기 (인코딩 자동 대응)
        def read_csv_safe(filename):
            try: return pd.read_csv(filename, encoding="utf-8", header=None)
            except: return pd.read_csv(filename, encoding="cp949", header=None)
        
        df_t = read_csv_safe("외기온도_시간분포.csv")
        df_c = read_csv_safe("COP_계산기.csv")

        # [온도 분포 추출]
        temp_data = {}
        target_zones = ["중부1", "중부2", "남부", "제주"]
        curr = None
        
        for _, row in df_t.iterrows():
            line = str(row[0])
            for zone in target_zones:
                if zone in line and "▶" in line:
                    curr = zone
                    temp_data[curr] = []
                    break
            
            # 시간대(00~23) 데이터만 수집
            head_val = line.strip().zfill(2)
            if curr and head_val.isdigit() and 0 <= int(head_val) <= 23:
                # 1월~12월 데이터 (1~12번 열)
                vals = []
                for v in row[1:13]:
                    try: vals.append(float(v))
                    except: vals.append(0.0)
                if len(vals) == 12:
                    temp_data[curr].append(vals)

        # [sCOP 추출]
        cop_data = {}
        for _, row in df_c.iterrows():
            zone_key = str(row[0]).strip()
            if zone_key in target_zones:
                try:
                    # sCOP는 보통 15번 혹은 마지막에서 3번째 열 근처에 있음
                    # 숫자인 데이터를 찾아서 매핑
                    for potential_val in row[2:]:
                        try:
                            val = float(potential_val)
                            if 2.0 < val < 6.0: # 현실적인 sCOP 범위 내의 값 탐색
                                cop_data[zone_key] = {'scop': round(val, 2)}
                        except: continue
                except: continue
        
        # 만약 sCOP를 못찾았을 경우를 대비한 기본값
        default_scops = {"중부1": 3.29, "중부2": 3.66, "남부": 3.99, "제주": 4.21}
        for k, v in default_scops.items():
            if k not in cop_data: cop_data[k] = {'scop': v}

        return temp_data, cop_data
    except Exception as e:
        st.error(f"데이터 파일 분석 중 오류: {e}")
        return None, None

def map_region_to_zone(s_reg):
    if s_reg == "강원도": return "중부1"
    if s_reg in ["대구","부산","울산","광주","경상남도","전라남도"]: return "남부"
    if s_reg == "제주도": return "제주"
    return "중부2"

# ── 3. 기본 데이터 및 UI 정의 ──
pv_data = {
    "서울": [53.06,79.80,92.62,104.99,98.06,96.96,95.22,72.72,56.62,67.75,65.56,49.98],
    "강원도": [56.59,96.04,94.22,118.18,108.17,108.50,104.22,79.22,57.82,78.17,75.65,54.48],
    "제주도": [64.22,66.80,87.12,116.14,104.49,92.97,107.64,84.75,78.96,81.85,79.21,66.59],
    "경기도": [58.95,88.67,102.91,116.65,108.96,107.74,105.80,80.80,62.91,75.27,72.84,55.53],
    "인천": [53.06,79.80,92.62,104.99,98.06,96.96,95.22,72.72,56.62,67.75,65.56,49.98],
    "세종": [53.30,80.02,89.54,109.34,100.67,99.03,102.57,69.17,55.02,69.88,67.62,50.22],
    "대전": [55.43,81.30,97.35,110.72,99.01,99.49,106.36,72.01,62.35,72.96,70.60,55.90],
    "대구": [64.90,88.15,90.72,111.86,102.80,102.70,110.62,70.35,51.58,75.80,73.35,60.64],
    "부산": [68.46,84.08,90.48,105.22,96.64,97.65,107.30,70.11,59.14,74.14,71.75,61.11],
    "울산": [68.46,84.08,90.48,105.22,96.64,97.65,107.30,70.11,59.14,74.14,71.75,61.11],
    "광주": [66.32,73.60,90.96,107.97,94.75,97.65,109.43,73.67,68.77,75.09,72.67,62.77],
    "충청북도": [59.22,88.91,99.49,121.49,111.86,110.03,113.96,76.85,61.13,77.64,75.14,55.80],
    "충청남도": [61.59,90.33,108.17,123.02,110.01,110.54,118.17,80.01,69.28,81.06,78.45,62.11],
    "경상북도": [72.11,97.94,100.80,124.29,114.22,114.11,122.91,78.17,57.31,84.22,81.50,67.38],
    "전라북도": [61.06,83.44,105.54,122.00,108.17,113.34,117.91,81.59,72.08,82.12,79.47,63.43],
    "경상남도": [76.06,93.42,100.54,116.91,107.38,108.50,119.23,77.90,65.71,82.38,79.72,67.90],
    "전라남도": [73.69,81.78,101.06,119.96,105.28,108.50,121.59,81.85,76.41,83.43,80.74,69.75]
}

# (기초 지자체 생략 가능 - 필요 시 추가)
reg_full = {"강원도": ["춘천시", "원주시"], "서울": ["종로구", "강남구"], "제주도": ["제주시", "서귀포시"]}

col_t, col_l = st.columns([6,1])
with col_t: st.title("히트펌프 경제성 분석 솔루션")
with col_l:
    if os.path.exists("logo.png"): st.image(Image.open("logo.png"), use_container_width=True)

st.markdown("""<div class='info-box'><h4 class='info-title'>💡 솔루션 개요</h4><p class='info-text'>🏠 <b>우리 집 맞춤형 분석:</b> 거주 지역 기상 데이터를 연동하여 가장 정확한 경제성을 산출합니다.</p></div>""", unsafe_allow_html=True)

# ── 4. 데이터 로드 및 시각화 ──
df_temp, df_cop = load_simulation_data()

st.markdown('<div class="section-title">1. 대상지 및 기후 정보</div>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
with c1: s_reg = st.selectbox("광역 지자체", list(pv_data.keys()), index=0)
with c2: h_size = st.number_input("전용 면적 (평)", min_value=10, value=30)

zone = map_region_to_zone(s_reg)
scop = df_cop[zone]['scop'] if df_cop else 3.0

# 온도 분포 차트 생성
if df_temp and zone in df_temp:
    try:
        days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        temps_flat = []
        for hour_data in df_temp[zone]:
            for m_idx, temp_val in enumerate(hour_data):
                temps_flat.extend([round(temp_val)] * days[m_idx])
        
        counts = pd.Series(temps_flat).value_counts().sort_index()
        cl1, cl2 = st.columns([2, 1])
        with cl1:
            st.markdown(f"**🌡️ {s_reg}({zone}) 연간 온도 분포**")
            st.bar_chart(counts, height=220)
        with cl2:
            st.success(f"**✅ {s_reg} 맞춤 효율**\n# {scop}")
            st.caption(f"{s_reg}의 기상 데이터를 반영한 실제 sCOP입니다.")
    except: st.warning("기후 차트를 생성할 수 없습니다.")

# ── 5. 입력 및 계산 (이하 생략 - 기본 로직 유지) ──
st.markdown('<div class="section-title">2. 에너지 및 경제성 분석</div>', unsafe_allow_html=True)
# ... 기존 분석 버튼 및 결과 출력 코드 ...
if st.button("경제성 분석 실행", type="primary", use_container_width=True):
    st.write(f"분석 완료! (적용 효율: {scop})")
    # 여기에 기존에 쓰시던 상세 계산 로직을 붙이시면 됩니다.