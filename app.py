import streamlit as st
import pandas as pd
import numpy as np
import io, os
from PIL import Image
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import altair as alt

st.set_page_config(page_title="히트펌프 경제성 분석 솔루션", layout="wide")

# ══════════════════════════════════════════════════════════
# 1. 스타일 정의
# ══════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
* { font-family: 'Pretendard', sans-serif; }
.info-box      { background:#f8fafc; border:1px solid #e2e8f0; border-radius:12px; padding:28px; margin-bottom:35px; }
.info-title    { color:#0f172a; font-size:1.25rem; font-weight:700; margin-bottom:15px; margin-top:0; }
.info-text     { color:#475569; font-size:1.0rem; line-height:1.7; margin-bottom:0; }
.section-title { color:#1e293b; font-weight:700; font-size:1.3rem; margin-top:40px; margin-bottom:16px; border-bottom:2px solid #cbd5e1; padding-bottom:8px; }
.help-text     { color:#64748b; font-size:0.85rem; margin-bottom:12px; line-height:1.4; }
.tariff-badge  { display:inline-block; background:#dbeafe; color:#1e40af; padding:4px 10px; border-radius:6px; font-size:0.85rem; font-weight:600; margin-right:6px; }
.solar-badge-x { display:inline-block; background:#fef3c7; color:#92400e; padding:4px 10px; border-radius:6px; font-size:0.85rem; font-weight:600; margin-right:6px; }
.solar-badge-o { display:inline-block; background:#dcfce7; color:#15803d; padding:4px 10px; border-radius:6px; font-size:0.85rem; font-weight:600; margin-right:6px; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# 2. 정적 데이터 정의
# ══════════════════════════════════════════════════════════

regions_full = {
    "서울": ["종로구","중구","용산구","성동구","광진구","동대문구","중랑구","성북구","강북구","도봉구","노원구","은평구","서대문구","마포구","양천구","강서구","구로구","금천구","영등포구","동작구","관악구","서초구","강남구","송파구","강동구"],
    "강원도": ["춘천시","원주시","강릉시","동해시","태백시","속초시","삼척시","홍천군","횡성군","영월군","평창군","정선군","철원군","화천군","양구군","인제군","고성군","양양군"],
    "경기도": ["수원시","고양시","용인시","성남시","부천시","화성시","안산시","남양주시","안양시","평택시","시흥시","파주시","의정부시","김포시","광주시","광명시","군포시","하남시","오산시","양주시","이천시","구리시","안성시","포천시","의왕시","여주시","동두천시","과천시","가평군","양평군","연천군"],
    "제주도": ["제주시","서귀포시"],
    "인천": ["중구","동구","미추홀구","연수구","남동구","부평구","계양구","서구","강화군","옹진군"],
    "부산": ["중구","서구","동구","영도구","부산진구","동래구","남구","북구","해운대구","사하구","금정구","강서구","연제구","수영구","사상구","기장군"],
    "대구": ["중구","동구","서구","남구","북구","수성구","달서구","달성군"],
    "세종": ["세종특별자치시"],
    "대전": ["동구","중구","서구","유성구","대덕구"],
    "울산": ["중구","남구","동구","북구","울주군"],
    "광주": ["동구","서구","남구","북구","광산구"],
    "충청북도": ["청주시","충주시","제천시","보은군","옥천군","영동군","증평군","진천군","괴산군","음성군","단양군"],
    "충청남도": ["천안시","공주시","보령시","아산시","서산시","논산시","계룡시","당진시","금산군","부여군","서천군","청양군","홍성군","예산군","태안군"],
    "전라북도": ["전주시","군산시","익산시","정읍시","남원시","김제시","완주군","진안군","무주군","장수군","임실군","순창군","고창군","부안군"],
    "전라남도": ["목포시","여수시","순천시","나주시","광양시","담양군","곡성군","구례군","고흥군","보성군","화순군","장흥군","강진군","해남군","영암군","무안군","함평군","영광군","장성군","완도군","진도군","신안군"],
    "경상북도": ["포항시","경주시","김천시","안동시","구미시","영주시","영천시","상주시","문경시","경산시","의성군","청송군","영양군","영덕군","청도군","고령군","성주군","칠곡군","예천군","봉화군","울진군","울릉군"],
    "경상남도": ["창원시","진주시","통영시","사천시","김해시","밀양시","거제시","양산시","의령군","함안군","창녕군","고성군","남해군","하동군","산청군","함양군","거창군","합천군"],
}

# 태양광 월별 발전량 (kWh/kW)
pv_monthly_data = {
    "서울":    [53.06,79.80,92.62,104.99,98.06,96.96,95.22,72.72,56.62,67.75,65.56,49.98],
    "인천":    [53.06,79.80,92.62,104.99,98.06,96.96,95.22,72.72,56.62,67.75,65.56,49.98],
    "세종":    [53.30,80.02,89.54,109.34,100.67,99.03,102.57,69.17,55.02,69.88,67.62,50.22],
    "대전":    [55.43,81.30,97.35,110.72,99.01,99.49,106.36,72.01,62.35,72.96,70.60,55.90],
    "대구":    [64.90,88.15,90.72,111.86,102.80,102.70,110.62,70.35,51.58,75.80,73.35,60.64],
    "부산":    [68.46,84.08,90.48,105.22,96.64,97.65,107.30,70.11,59.14,74.14,71.75,61.11],
    "울산":    [68.46,84.08,90.48,105.22,96.64,97.65,107.30,70.11,59.14,74.14,71.75,61.11],
    "광주":    [66.32,73.60,90.96,107.97,94.75,97.65,109.43,73.67,68.77,75.09,72.67,62.77],
    "경기도":  [58.95,88.67,102.91,116.65,108.96,107.74,105.80,80.80,62.91,75.27,72.84,55.53],
    "강원도":  [56.59,96.04,94.22,118.18,108.17,108.50,104.22,79.22,57.82,78.17,75.65,54.48],
    "충청북도":[59.22,88.91,99.49,121.49,111.86,110.03,113.96,76.85,61.13,77.64,75.14,55.80],
    "충청남도":[61.59,90.33,108.17,123.02,110.01,110.54,118.17,80.01,69.28,81.06,78.45,62.11],
    "경상북도":[72.11,97.94,100.80,124.29,114.22,114.11,122.91,78.17,57.31,84.22,81.50,67.38],
    "전라북도":[61.06,83.44,105.54,122.00,108.17,113.34,117.91,81.59,72.08,82.12,79.47,63.43],
    "경상남도":[76.06,93.42,100.54,116.91,107.38,108.50,119.23,77.90,65.71,82.38,79.72,67.90],
    "전라남도":[73.69,81.78,101.06,119.96,105.28,108.50,121.59,81.85,76.41,83.43,80.74,69.75],
    "제주도":  [64.22,66.80,87.12,116.14,104.49,92.97,107.64,84.75,78.96,81.85,79.21,66.59],
}

# HDD(난방도일), Tbase=15°C
hdd_monthly = {
    "중부1": [750, 596, 456, 198,   0,   0,   0,   0,   0, 161, 405, 676],
    "중부2": [521, 398, 264,  54,   0,   0,   0,   0,   0,   0, 225, 450],
    "남부":  [347, 269, 171,   0,   0,   0,   0,   0,   0,   0, 114, 295],
    "제주":  [273, 224, 140,   0,   0,   0,   0,   0,   0,   0,  45, 211],
}

# 지역별 sCOP
SCOP_BY_ZONE = {"중부1": 3.29, "중부2": 3.66, "남부": 3.99, "제주": 4.21}

# ── 엑셀 블록 헤더 위치 (전기요금 시트) ──
# (요금제, 태양광유무, 난방유형) → 헤더 행 번호
# 일반용은 별도 미터(HP 전용 계약)이라 태양광 영향 없음 → 태X만 존재
EXCEL_BLOCK_HEADERS = {
    ("누진제", "태X", "도시가스(콘덴싱)"): 5,
    ("누진제", "태X", "도시가스(일반)"):  18,
    ("누진제", "태X", "등유"):           31,
    ("누진제", "태X", "LPG"):            44,
    ("누진제", "태O", "도시가스(콘덴싱)"): 57,
    ("누진제", "태O", "도시가스(일반)"):  70,
    ("누진제", "태O", "등유"):           83,
    ("누진제", "태O", "LPG"):            96,
    ("일반용", "태X", "도시가스(콘덴싱)"): 109,
    ("일반용", "태X", "도시가스(일반)"):  122,
    ("일반용", "태X", "등유"):           135,
    ("일반용", "태X", "LPG"):            148,
    ("계시별", "태X", "도시가스(콘덴싱)"): 161,
    ("계시별", "태X", "도시가스(일반)"):  174,
    ("계시별", "태X", "등유"):           187,
    ("계시별", "태X", "LPG"):            200,
    ("계시별", "태O", "도시가스(콘덴싱)"): 213,
    ("계시별", "태O", "도시가스(일반)"):  226,
    ("계시별", "태O", "등유"):           239,
    ("계시별", "태O", "LPG"):            252,
}

# UI 옵션 → 엑셀 블록 키 매핑
HEATING_TYPE_MAP = {
    "가스 콘덴싱 보일러": "도시가스(콘덴싱)",
    "가스 일반 보일러":   "도시가스(일반)",
    "등유 보일러":        "등유",
    "LPG 보일러":         "LPG",
}


# ══════════════════════════════════════════════════════════
# 3. 엑셀 로더
# ══════════════════════════════════════════════════════════

@st.cache_data
def load_tariff_xlsx():
    """
    전기요금완료본.xlsx에서 모든 블록(20개) 데이터를 로드합니다.

    [엑셀 시트: "전기요금"]
    - 3개 요금제(누진제/일반용/계시별) × 4개 난방유형 × 2개 태양광 옵션(태X/태O)
    - 일반용은 태X만 존재 (HP 전용 별도 미터)

    [블록당 데이터]
    - 헤더 행: col 19=HP 연합계(원), col 20=기존난방비 연합계(원), col 21=Saving 비율
    - 1~12월 행: col 9~18 = 월/기본/사용량/기후환경/연료비조정/전기요금계/VAT/기금/청구합계/HP전기요금
    """
    fname = "전기요금완료본.xlsx"
    candidates = [
        fname,
        os.path.join(os.path.dirname(os.path.abspath(__file__)), fname),
        os.path.join(os.getcwd(), fname),
    ]
    fp = next((p for p in candidates if os.path.exists(p)), None)
    if fp is None:
        return None, f"파일을 찾을 수 없습니다: {fname}"

    try:
        wb = load_workbook(fp, data_only=True)
        ws = wb["전기요금"]
        blocks = {}
        for key, hr in EXCEL_BLOCK_HEADERS.items():
            hp_ann = ws.cell(row=hr, column=19).value
            ex_ann = ws.cell(row=hr, column=20).value
            sr     = ws.cell(row=hr, column=21).value
            monthly = []
            for m in range(12):
                r = hr + 1 + m
                monthly.append({
                    "월":           ws.cell(row=r, column=9).value,
                    "기본요금":     ws.cell(row=r, column=10).value or 0,
                    "사용량요금":   ws.cell(row=r, column=11).value or 0,
                    "기후환경요금": ws.cell(row=r, column=12).value or 0,
                    "연료비조정":   ws.cell(row=r, column=13).value or 0,
                    "전기요금계":   ws.cell(row=r, column=14).value or 0,
                    "부가가치세":   ws.cell(row=r, column=15).value or 0,
                    "기반기금":     ws.cell(row=r, column=16).value or 0,
                    "청구요금합계": ws.cell(row=r, column=17).value or 0,
                    "HP전기요금":   ws.cell(row=r, column=18).value or 0,
                })
            blocks[key] = {
                "hp_annual_won":       float(hp_ann) if hp_ann else 0,
                "existing_annual_won": float(ex_ann) if ex_ann else 0,
                "saving_ratio":        float(sr)     if sr     else 0,
                "monthly":             monthly,
            }
        return blocks, None
    except Exception as e:
        return None, str(e)


# ══════════════════════════════════════════════════════════
# 4. 계산 함수
# ══════════════════════════════════════════════════════════

def map_region_to_zone(s):
    """광역 지자체명 → 기후 존(zone) 매핑"""
    if s == "강원도": return "중부1"
    if s in ["대구","부산","울산","광주","경상남도","전라남도"]: return "남부"
    if s == "제주도": return "제주"
    return "중부2"


def calc_capex(h_type, h_size):
    """히트펌프 설치 총비용(만원). 국내 기업 자료 기반 1,000만원 고정."""
    return 1000


def get_block_key(tariff_choice, s_capa, heating_ui):
    """
    UI 입력으로부터 엑셀 블록 키 결정.
    - 태양광 0 → 태X, 0보다 크면 → 태O
    - 일반용은 항상 태X (별도 미터)
    """
    heating = HEATING_TYPE_MAP.get(heating_ui, "도시가스(콘덴싱)")
    if tariff_choice == "일반용":
        solar = "태X"
    else:
        solar = "태O" if s_capa > 0 else "태X"
    return (tariff_choice, solar, heating), solar


def calc_csv_jan_heat_man(zone):
    """
    엑셀 기준 가구의 1월 난방비(만원).
    - 모든 블록의 기존 난방비 연합계 = 650,516원 (서울/중부2 기후 가정)
    - 1월 비중은 사용자 지역의 HDD 비율로 안분
    """
    EXCEL_BASE_ANNUAL_WON = 650516
    hdd = hdd_monthly[zone]
    if sum(hdd) == 0: return 0
    return EXCEL_BASE_ANNUAL_WON * hdd[0] / sum(hdd) / 10000


def apply_block_with_scale(block, scale):
    """엑셀 블록 데이터에 사용자 가구 규모(scale) 적용"""
    monthly_won = [round(m["청구요금합계"] * scale) for m in block["monthly"]]
    monthly_man = [round(v / 10000, 2) for v in monthly_won]
    hp_ann_won  = block["hp_annual_won"]   * scale
    ex_ann_won  = block["existing_annual_won"] * scale
    saving_won  = ex_ann_won - hp_ann_won
    return {
        "monthly_won":   monthly_won,
        "monthly_man":   monthly_man,
        "hp_annual_man": round(hp_ann_won / 10000, 1),
        "hp_annual_won": round(hp_ann_won),
        "ex_annual_man": round(ex_ann_won / 10000, 1),
        "ex_annual_won": round(ex_ann_won),
        "saving_man":    round(saving_won / 10000, 1),
        "saving_ratio":  block["saving_ratio"],
    }


# ══════════════════════════════════════════════════════════
# 5. UI 메인
# ══════════════════════════════════════════════════════════

tariff_blocks, load_err = load_tariff_xlsx()

col_t, col_l = st.columns([6, 1])
with col_t: st.title("히트펌프 경제성 분석 솔루션")
with col_l:
    if os.path.exists("logo.png"): st.image(Image.open("logo.png"), use_container_width=True)

if load_err:
    st.error(f"⚠️ 전기요금완료본.xlsx 로드 실패: {load_err}")
    st.info("repo 루트(또는 app.py 옆)에 `전기요금완료본.xlsx` 파일이 있는지 확인해 주세요.")
    st.stop()
else:
    st.success(f"✅ 전기요금완료본.xlsx 로드 완료 — 총 {len(tariff_blocks)}개 블록 사용 가능")

st.markdown("""
<div class='info-box'>
  <h4 class='info-title'>💡 솔루션 개요</h4>
  <p class='info-text'>
    🏠 <b>시민이 직접 해보는 탄소중립 계산기:</b> 거주 환경과 평소 에너지 사용량만 입력하면,
    친환경 히트펌프(AWHP) 전환 시 <b>얼마나 경제적 이득인지</b> 바로 확인하실 수 있습니다.<br><br>
    ⚡ <b>요금제·태양광 자동 매칭:</b> 선택하신 전기 요금제와 태양광 용량(kW)에 따라
    엑셀 데이터의 정확한 블록(태O / 태X)을 자동으로 적용합니다.
  </p>
</div>
""", unsafe_allow_html=True)

# ── 섹션 1: 대상지 ──
st.markdown('<div class="section-title">1. 대상지 기본 정보</div>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
with c1: s_reg  = st.selectbox("광역 지자체", list(regions_full.keys()), index=0)
with c2: s_sub  = st.selectbox("기초 지자체", regions_full.get(s_reg, ["전체"]))
c3, c4 = st.columns(2)
with c3: h_type = st.selectbox("주거 형태", ["단독 주택 / 다가구 주택", "아파트", "연립 / 빌라 / 다세대 주택"])
with c4: h_size = st.number_input("전용 면적 (평)", min_value=10, value=30)

zone        = map_region_to_zone(s_reg)
dynamic_cop = SCOP_BY_ZONE[zone]

# ── 섹션 2: 에너지 소비 ──
st.markdown('<div class="section-title">2. 에너지 소비 현황</div>', unsafe_allow_html=True)
cv1, cv2 = st.columns(2)
with cv1: w_heat = st.number_input("동절기(1월) 평균 난방비 (만원)", value=20)
with cv2: w_elec = st.number_input("동절기(1월) 전기요금 (만원)", value=6)

ce1, ce2 = st.columns(2)
with ce1:
    heating_type = st.selectbox(
        "현재 주택의 난방 방식",
        list(HEATING_TYPE_MAP.keys()),
        help="현재 사용 중인 난방 연료 방식을 선택해 주세요."
    )
with ce2:
    cooking_type = st.selectbox(
        "사용하는 취사 기기",
        ["인덕션 (전기)", "도시가스", "LPG"],
    )

# ── 섹션 3: 시뮬레이션 변수 ──
st.markdown('<div class="section-title">3. 시뮬레이션 상수 변수</div>', unsafe_allow_html=True)
cs1, cs2 = st.columns(2)
with cs1:
    f_inf  = st.slider("가스/등유요금 인상률 (%)", 0.0, 15.0, 5.0)
    e_inf  = st.slider("전기요금 인상률 (%)", 0.0, 15.0, 3.0)
    s_capa = st.number_input(
        "태양광 용량 (kW)",
        value=0.0,
        help="0kW 입력 시 엑셀 '태X' 블록, 0kW 초과 시 '태O' 블록이 자동 적용됩니다."
    )

with cs2:
    sub_nat  = st.checkbox("정부 보조금 적용 (560만원)", value=True)
    is_south = s_reg in ["제주도","경상남도","전라남도"]
    sub_loc  = st.checkbox("지자체 매칭 보조금 적용 (280만원)", value=is_south)
    st.caption("*2026년 현재 제주, 경남, 전남은 보조금 신청이 가능합니다.")

    st.markdown("---")
    st.markdown("**전기 요금제 선택**")
    st.markdown("""
<div class='help-text'>
선택한 요금제 + 태양광 용량 입력에 따라 엑셀 블록이 자동 적용됩니다.
</div>""", unsafe_allow_html=True)

    tariff_choice = st.radio(
        "요금제",
        ["누진제", "일반용", "계시별"],
        horizontal=True,
        label_visibility="collapsed",
        help="누진제: 주택용 저압 누진제 / 일반용: HP 전용 별도 미터 / 계시별: 시간대별 요금제",
    )
    if tariff_choice == "일반용":
        st.caption("ℹ️ 일반용은 HP 전용 별도 미터라 태양광 영향 없음 (항상 태X 블록 사용).")

# 입력 변경 감지
_input_key = (w_heat, w_elec, s_reg, h_type, h_size, heating_type, cooking_type,
              f_inf, e_inf, s_capa, sub_nat, sub_loc, tariff_choice)
if st.session_state.get("_last_input_key") != _input_key:
    st.session_state.analyzed = False
    st.session_state["_last_input_key"] = _input_key

if "analyzed" not in st.session_state: st.session_state.analyzed = False
if st.button("경제성 분석 실행", type="primary", use_container_width=True):
    st.session_state.analyzed = True


# ══════════════════════════════════════════════════════════
# 6. 분석 결과
# ══════════════════════════════════════════════════════════
if st.session_state.analyzed:

    # ── ① 엑셀 블록 결정 ──
    block_key, solar_flag = get_block_key(tariff_choice, s_capa, heating_type)
    block = tariff_blocks[block_key]

    # ── ② 가구 규모 보정 ──
    csv_jan_man = calc_csv_jan_heat_man(zone)
    scale       = (w_heat / csv_jan_man) if csv_jan_man > 0 else 1.0
    result      = apply_block_with_scale(block, scale)

    # ── ③ 보조금/투자비 ──
    total_sub = (560 if sub_nat else 0) + (280 if sub_loc else 0)
    capex     = calc_capex(h_type, h_size)
    net_cap   = max(0, capex - total_sub)

    # ── ④ 18년 시뮬레이션 (엑셀 데이터 기반) ──
    ann_heat_base = result["ex_annual_man"]
    ann_hp_op     = result["hp_annual_man"]

    years, gas_cum, hp_cum, net_p = list(range(1, 19)), [], [], []
    g_s, h_s, pb = 0.0, float(net_cap), "18년 초과"
    for y in years:
        cg = ann_heat_base * ((1 + f_inf / 100) ** y)
        ch = ann_hp_op    * ((1 + e_inf / 100) ** y)
        g_s += cg; h_s += ch
        p = int(g_s - h_s)
        gas_cum.append(int(g_s)); hp_cum.append(int(h_s)); net_p.append(p)
        if pb == "18년 초과" and p > 0: pb = f"{y}년차"

    # ══════════════════════════════════════════════════════════
    # 결과 출력
    # ══════════════════════════════════════════════════════════
    st.markdown('<div class="section-title">📊 분석 결과 요약</div>', unsafe_allow_html=True)

    st.markdown(f"""
<div style='margin-bottom:16px;'>
  <span class='tariff-badge'>요금제: {tariff_choice}</span>
  <span class='{"solar-badge-o" if solar_flag == "태O" else "solar-badge-x"}'>태양광: {solar_flag} ({s_capa}kW)</span>
  <span class='tariff-badge'>난방: {HEATING_TYPE_MAP[heating_type]}</span>
  <span class='tariff-badge'>규모 보정: ×{round(scale, 2)}</span>
</div>
""", unsafe_allow_html=True)

    # 평수별 HP 정보
    if h_size < 20:
        hp_space, hp_space_mm, hp_capacity = "소형 냉장고 크기", "595 × 625 mm", "6 kW"
    elif h_size <= 28:
        hp_space, hp_space_mm, hp_capacity = "워시타워 1대 크기", "800 × 1,115 mm", "10 kW"
    elif h_size <= 35:
        hp_space, hp_space_mm, hp_capacity = "워시타워 1대 크기", "800 × 1,115 mm", "12 kW"
    else:
        hp_space, hp_space_mm, hp_capacity = "보일러실 크기", "1,120 × 1,666 mm", "16 kW"

    ca, cb, cc, cd = st.columns(4)
    ca.metric("투자 회수 시점", pb)
    cb.metric("18년 순이익", f"{net_p[-1]:,} 만원")
    cc.metric("히트펌프 설치 공간", hp_space)
    cc.markdown(f"<div style='font-size:0.78rem; color:#64748b; margin-top:-10px;'>{hp_space_mm}</div>", unsafe_allow_html=True)
    cd.metric("적정 히트펌프 용량", hp_capacity)

    # ── 핵심 Saving 지표 ──
    st.markdown('<div class="section-title">💰 전기요금 분석 (엑셀 데이터 기반)</div>', unsafe_allow_html=True)
    s1, s2, s3 = st.columns(3)
    s1.metric(
        "HP 연간 전기요금",
        f"{result['hp_annual_man']:,.1f} 만원",
        help=f"엑셀 [{tariff_choice}/{solar_flag}/{HEATING_TYPE_MAP[heating_type]}] HP 연합계 × 규모보정"
    )
    s2.metric(
        f"기존 연간 난방비 ({HEATING_TYPE_MAP[heating_type]})",
        f"{result['ex_annual_man']:,.1f} 만원",
        help="엑셀 기준 가구 기존 난방비 × 규모 보정"
    )
    s3.metric(
        "연간 절감액",
        f"{result['saving_man']:,.1f} 만원",
        delta=f"{round(result['saving_ratio'] * 100)}% 절감",
    )

    # ── 월별 차트 ──
    months = list(range(1, 13))
    hdd_zone = hdd_monthly[zone]
    hdd_jan  = hdd_zone[0] if hdd_zone[0] > 0 else 1
    monthly_ex_man = [round(w_heat * hdd_zone[m-1] / hdd_jan, 2) for m in months]

    df_chart = pd.DataFrame({
        "월":               [f"{m}월" for m in months],
        "기존 난방비(만원)": monthly_ex_man,
        "HP 청구요금(만원)": result["monthly_man"],
    }).melt("월", var_name="구분", value_name="금액(만원)")

    chart = alt.Chart(df_chart).mark_bar().encode(
        x=alt.X("월:O", sort=[f"{m}월" for m in months]),
        y=alt.Y("금액(만원):Q"),
        color=alt.Color("구분:N", scale=alt.Scale(
            domain=["기존 난방비(만원)", "HP 청구요금(만원)"],
            range=["#f87171", "#60a5fa"]
        ), legend=alt.Legend(orient="top", title=None)),
        xOffset="구분:N",
        tooltip=["월", "구분", "금액(만원)"],
    ).properties(height=280, title="월별 기존 난방비 vs HP 청구요금")
    st.altair_chart(chart, use_container_width=True)

    # ── 월별 상세 ──
    with st.expander("📋 월별 상세 데이터 (엑셀 원본 기반)"):
        df_detail = pd.DataFrame({
            "월":                [f"{m}월" for m in months],
            "기존 난방비(만원)": monthly_ex_man,
            "HP 청구요금(만원)": result["monthly_man"],
            "월별 절감액(만원)": [round(monthly_ex_man[m-1] - result["monthly_man"][m-1], 2)
                                  for m in months],
            "HP 청구요금(원)":   result["monthly_won"],
        })
        st.dataframe(df_detail, use_container_width=True, hide_index=True)

    # ── 18년 차트 ──
    st.markdown('<div class="section-title">📈 18년 장기 시뮬레이션</div>', unsafe_allow_html=True)
    g1, g2 = st.columns(2)
    with g1:
        st.write("**18년 누적 비용 흐름**")
        df_a = pd.DataFrame({"연도": years, "기존": gas_cum, "HP": hp_cum}).melt(
            "연도", var_name="시나리오", value_name="비용")
        st.altair_chart(alt.Chart(df_a).mark_area(opacity=0.5).encode(
            x="연도:O", y="비용:Q", color="시나리오:N"
        ), use_container_width=True)
    with g2:
        st.write("**연도별 순수익(Cash Flow)**")
        df_c2 = pd.DataFrame({"연도": years, "순수익": net_p,
                               "상태": ["수익" if p > 0 else "회수" for p in net_p]})
        st.altair_chart(alt.Chart(df_c2).mark_bar().encode(
            x="연도:O", y="순수익:Q", color="상태:N"
        ), use_container_width=True)

    # ── 가정값 상세 ──
    with st.expander("📋 적용된 계산 가정값 및 출처"):
        st.markdown(f"""
| 항목 | 적용값 | 근거 |
|------|--------|------|
| 적용 블록 | **{tariff_choice} / {solar_flag} / {HEATING_TYPE_MAP[heating_type]}** | 전기요금완료본.xlsx |
| 규모 보정 계수 | **×{round(scale, 2)}** | 사용자 1월 난방비({w_heat}만원) ÷ 엑셀 기준({csv_jan_man:.2f}만원) |
| 설비 CAPEX | **{capex}만원** | 국내 기업 자료 (설치비 포함) |
| 정부+지방 보조금 | **{total_sub}만원** | {"정부 560 + 지방 280" if sub_nat and sub_loc else ("정부 560만원" if sub_nat else "지방 280만원" if sub_loc else "없음")} |
| 순 투자비 | **{net_cap}만원** | CAPEX − 보조금 |
| 기존 연간 난방비 | **{ann_heat_base}만원** | 엑셀 기존난방비 × 규모보정 |
| HP 연간 전기요금 | **{ann_hp_op}만원** | 엑셀 HP연합계 × 규모보정 |
| 지역 sCOP (참고) | **{dynamic_cop}** | 기후 존 ({zone}) 추정값 |
        """)

    # ── 엑셀 다운로드 ──
    wb  = Workbook()
    hf  = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    sf  = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    pf  = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
    gf  = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")
    fw  = Font(color="FFFFFF", bold=True)
    fb  = Font(bold=True)
    fi  = Font(color="0000FF", bold=True)
    fg  = Font(color="166534", bold=True)
    thin   = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"),  bottom=Side(style="thin"))
    center = Alignment(horizontal="center")
    right  = Alignment(horizontal="right")

    # ① 입력·가정 시트
    ws1 = wb.active; ws1.title = "①입력_가정"
    ws1.merge_cells("A1:D1")
    ws1["A1"] = f"히트펌프 경제성 분석 ({s_reg} / {tariff_choice} / {solar_flag})"
    ws1["A1"].fill = hf; ws1["A1"].font = fw; ws1["A1"].alignment = center
    rows1 = [
        ("항목",            "값",                                         "단위",   "산출 근거"),
        ("1월 난방비",       w_heat,                                       "만원",   "사용자 입력"),
        ("1월 전기요금",     w_elec,                                       "만원",   "사용자 입력"),
        ("난방 방식",        HEATING_TYPE_MAP[heating_type],               "-",      "사용자 선택"),
        ("전기 요금제",      tariff_choice,                                "-",      "사용자 선택"),
        ("태양광 용량",      s_capa,                                       "kW",     f"입력 → 블록: {solar_flag}"),
        ("적용 엑셀 블록",   f"{tariff_choice}/{solar_flag}/{HEATING_TYPE_MAP[heating_type]}", "-", "전기요금완료본.xlsx"),
        ("규모 보정 계수",   round(scale, 2),                              "배",     f"={w_heat}÷{round(csv_jan_man,2)}"),
        ("지역 sCOP",        dynamic_cop,                                  "-",      f"기후 존 ({zone})"),
        ("설비 CAPEX",       capex,                                        "만원",   "국내 기업 자료"),
        ("정부 보조금",      560 if sub_nat else 0,                        "만원",   "기후에너지환경부 2026"),
        ("지방 보조금",      280 if sub_loc else 0,                        "만원",   "정부 50% 매칭"),
        ("순 투자비",        net_cap,                                      "만원",   "=CAPEX-보조금"),
        ("기존 연간 난방비", ann_heat_base,                                "만원",   "엑셀 기존난방비×규모보정"),
        ("HP 연간 전기요금", ann_hp_op,                                    "만원",   "엑셀 HP연합계×규모보정"),
        ("연간 절감액",      result["saving_man"],                         "만원",   "=기존-HP"),
        ("Saving 비율",      f"{round(result['saving_ratio']*100,1)}%",    "-",      "엑셀 원본"),
    ]
    for ri, rdata in enumerate(rows1, 3):
        for ci, val in enumerate(rdata, 1):
            c = ws1.cell(row=ri, column=ci, value=val); c.border = thin
            if ri == 3: c.fill = sf; c.font = fb
            elif ci == 2 and ri != 3: c.font = fi; c.alignment = right
    ws1.column_dimensions["A"].width = 22
    ws1.column_dimensions["D"].width = 45

    # ② 월별 청구요금 시트
    ws2 = wb.create_sheet("②월별_청구요금")
    ws2.merge_cells("A1:G1")
    ws2["A1"] = f"월별 청구요금 [{tariff_choice} / {solar_flag} / {HEATING_TYPE_MAP[heating_type]}]"
    ws2["A1"].fill = hf; ws2["A1"].font = fw; ws2["A1"].alignment = center
    headers2 = ["월","기존 난방비(만원)","HP 청구요금(만원)","HP 청구요금(원)","월 절감액(만원)","Saving %","비고"]
    for ci, h in enumerate(headers2, 1):
        c = ws2.cell(row=2, column=ci, value=h)
        c.fill = sf; c.font = fb; c.border = thin; c.alignment = center

    for m in range(1, 13):
        r = m + 2
        ex  = monthly_ex_man[m-1]
        hp  = result["monthly_man"][m-1]
        won = result["monthly_won"][m-1]
        sav = round(ex - hp, 2)
        sav_pct = round((sav / ex * 100) if ex > 0 else 0, 1)
        note = "난방월" if hdd_zone[m-1] > 0 else "비난방월"
        for ci, val in enumerate([f"{m}월", ex, hp, won, sav, f"{sav_pct}%", note], 1):
            c = ws2.cell(row=r, column=ci, value=val); c.border = thin
        ws2.cell(row=r, column=5).font = fg
        if m % 2 == 0:
            for ci in range(1, 8): ws2.cell(row=r, column=ci).fill = gf

    r_sum = 15
    ws2.cell(row=r_sum, column=1, value="연간 합계").font = fb
    ws2.cell(row=r_sum, column=2, value=ann_heat_base).font = fb
    ws2.cell(row=r_sum, column=3, value=ann_hp_op).font = fb
    ws2.cell(row=r_sum, column=4, value=result["hp_annual_won"]).font = fb
    ws2.cell(row=r_sum, column=5, value=result["saving_man"]).font = fg
    ws2.cell(row=r_sum, column=6, value=f"{round(result['saving_ratio']*100,1)}%").font = fg
    for ci in range(1, 8): ws2.cell(row=r_sum, column=ci).border = thin
    for col in "ABCDEFG": ws2.column_dimensions[col].width = 18

    # ③ 18년 재무 분석
    ws3 = wb.create_sheet("③18년_재무_분석")
    ws3.merge_cells("A1:H1")
    ws3["A1"] = "18년 장기 투자 회수 시뮬레이션"
    ws3["A1"].fill = hf; ws3["A1"].font = fw; ws3["A1"].alignment = center
    for ci, h in enumerate(["연도","물가지수(4%)","기존 OPEX(만)","HP OPEX(만)","연간 순이익(만)","누적 순이익(만)","ROI","상태"], 1):
        c = ws3.cell(row=2, column=ci, value=h)
        c.fill = sf; c.font = fb; c.border = thin; c.alignment = center
    ref_cap = "'①입력_가정'!$B$15"
    for y in range(1, 19):
        r = y + 2
        ws3.cell(row=r, column=1, value=f"{y}년차").border = thin
        ws3.cell(row=r, column=2, value=f"=(1+0.04)^{y-1}").border = thin
        ws3.cell(row=r, column=3, value=ann_heat_base).border = thin
        ws3.cell(row=r, column=4, value=ann_hp_op).border = thin
        ws3.cell(row=r, column=5, value=f"=C{r}-D{r}").border = thin
        if y == 1: ws3.cell(row=r, column=6, value=f"=E{r}-{ref_cap}").border = thin
        else:      ws3.cell(row=r, column=6, value=f"=F{r-1}+E{r}").border = thin
        ws3.cell(row=r, column=7, value=f"=F{r}/{ref_cap}").border = thin
        ws3.cell(row=r, column=7).number_format = "0%"
        ws3.cell(row=r, column=8, value=f'=IF(F{r}>0,"수익","회수중")').border = thin
        if y % 2 == 0:
            for ci in range(1, 9): ws3.cell(row=r, column=ci).fill = pf
    for col in "ABCDEFGH": ws3.column_dimensions[col].width = 16

    buf = io.BytesIO(); wb.save(buf)
    st.markdown("---")
    st.download_button(
        label="🚀 전문가용 수식 연동 정밀 엑셀 다운로드",
        data=buf.getvalue(),
        file_name=f"Expert_Report_{s_reg}_{tariff_choice}_{solar_flag}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )