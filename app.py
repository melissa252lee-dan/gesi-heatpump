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

# ── 2. 기후 데이터 파싱 엔진 (에러 방지 강화 버전) ──
@st.cache_data
def load_simulation_data():
    try:
        def read_csv_safe(filename):
            try: return pd.read_csv(filename, encoding="utf-8", header=None)
            except: return pd.read_csv(filename, encoding="cp949", header=None)
        
        df_t = read_csv_safe("외기온도_시간분포.csv")
        df_c = read_csv_safe("COP_계산기.csv")

        temp_data = {}
        target_zones = ["중부1", "중부2", "남부", "제주"]
        curr = None
        
        for _, row in df_t.iterrows():
            line = str(row[0]).strip()
            for zone in target_zones:
                if zone in line and "▶" in line:
                    curr = zone
                    temp_data[curr] = []
                    break
            head_val = line.strip().zfill(2)
            if curr and head_val.isdigit() and 0 <= int(head_val) <= 23:
                vals = []
                for v in row[1:13]:
                    try: vals.append(float(v))
                    except: vals.append(0.0)
                if len(vals) == 12: temp_data[curr].append(vals)

        cop_data = {}
        for _, row in df_c.iterrows():
            zone_key = str(row[0]).strip()
            if zone_key in target_zones:
                for potential_val in row[2:]:
                    try:
                        val = float(potential_val)
                        if 2.0 < val < 6.0: cop_data[zone_key] = {'scop': round(val, 2)}
                    except: continue
        
        default_scops = {"중부1": 3.29, "중부2": 3.66, "남부": 3.99, "제주": 4.21}
        for k, v in default_scops.items():
            if k not in cop_data: cop_data[k] = {'scop': v}
        return temp_data, cop_data
    except Exception as e:
        return None, None

def map_region_to_zone(s_reg):
    if s_reg == "강원도": return "중부1"
    if s_reg in ["대구","부산","울산","광주","경상남도","전라남도"]: return "남부"
    if s_reg == "제주도": return "제주"
    return "중부2"

# ── 3. 보일러 및 요금 엔진 ──
BOILER_EFF = { "가스 컨덴싱 보일러": 0.92, "일반 가스 보일러": 0.82, "등유 보일러": 0.85, "LPG 보일러": 0.82 }
FUEL_PRICE_PER_MJ = { "가스 컨덴싱 보일러": 68.0, "일반 가스 보일러": 68.0, "등유 보일러": 95.0, "LPG 보일러": 105.0 }
HP_MONTHLY_LOAD = [1.0,0.9,0.4,0.15,0.05,0.05,0.05,0.05,0.05,0.15,0.4,0.85]
MONTH_SEASON = {1:"other",2:"other",3:"other",4:"other",5:"other",6:"other", 7:"summer",8:"summer",9:"other",10:"other",11:"other",12:"other"}
MONTH_TOU = {1:"ws",2:"ws",3:"sf",4:"sf",5:"sf",6:"ws", 7:"ws",8:"ws",9:"sf",10:"sf",11:"ws",12:"ws"}

def calc_elec_bill(kwh, tariff, season="other", contract_kw=5.0):
    kwh = max(kwh, 0)
    if tariff == "주택용 누진제 (저압)":
        if kwh <= 300: base, energy = 910, kwh * 120
        elif kwh <= 450: base, energy = 1_600, 300*120 + (kwh-300)*214.6
        else: base, energy = 7_300, 300*120 + 150*214.6 + (kwh-450)*307.3
    elif tariff == "주택용 누진제 (고압)":
        if season == "summer":
            if kwh <= 300: base, energy = 730, kwh*105
            elif kwh <= 450: base, energy = 1_260, 300*105+(kwh-300)*174
            else: base, energy = 6_060, 300*105+150*174+(kwh-450)*242.3
        else:
            if kwh <= 200: base, energy = 730, kwh*105
            elif kwh <= 400: base, energy = 1_260, 200*105+(kwh-200)*174
            else: base, energy = 6_060, 200*105+200*174+(kwh-400)*242.3
    elif tariff == "계시별 요금제 TOU (제주)":
        base = 4_310 * contract_kw
        r = {"경":138.7,"중":184.7,"최":220.5} if season=="ws" else {"경":125.8,"중":153.8,"최":172.4}
        energy = kwh * (0.4*r["경"] + 0.4*r["중"] + 0.2*r["최"])
    else: return calc_elec_bill(kwh, "주택용 누진제 (저압)", season, contract_kw)
    fee = base + energy + (14.0) * kwh 
    return round((fee * 1.127) / 10_000, 4)

def reverse_kwh(bill_man, tariff, season="other", contract_kw=5.0):
    lo, hi = 0.0, 3000.0
    for _ in range(60):
        mid = (lo + hi) / 2
        if calc_elec_bill(mid, tariff, season, contract_kw) < bill_man: lo = mid
        else: hi = mid
    return round(mid, 1)

def heat_to_hp_kwh(heat_man, boiler_type, cop=3.0):
    eff = BOILER_EFF.get(boiler_type, 0.85); price = FUEL_PRICE_PER_MJ.get(boiler_type, 68.0)
    return round(((heat_man * 10000) / price) / 3.6 * eff / cop, 1)

def estimate_hp_kw(size, htype):
    rate = 0.15 if "아파트" in htype else (0.17 if "연립" in htype else 0.20)
    return round(size * rate, 1)

# ── 4. 메인 UI ──
col_t, col_l = st.columns([6,1])
with col_t: st.title("히트펌프 경제성 분석 솔루션")
with col_l:
    if os.path.exists("logo.png"): st.image(Image.open("logo.png"), use_container_width=True)

st.markdown("""<div class='info-box'><h4 class='info-title'>💡 솔루션 개요</h4><p class='info-text'>🏠 <b>우리 집 맞춤형 분석:</b> 거주 지역 기상 데이터를 연동하여 가장 정확한 경제성을 산출합니다. 🌱💰</p></div>""", unsafe_allow_html=True)

df_temp, df_cop = load_simulation_data()

st.markdown('<div class="section-title">1. 대상지 기본 정보</div>', unsafe_allow_html=True)
c1,c2 = st.columns(2)
# 광역 지자체 리스트 (기존 코드 유지)
pv_monthly_data = {
    "서울": [53.06,79.80,92.62,104.99,98.06,96.96,95.22,72.72,56.62,67.75,65.56,49.98], "인천": [53.06,79.80,92.62,104.99,98.06,96.96,95.22,72.72,56.62,67.75,65.56,49.98],
    "세종": [53.30,80.02,89.54,109.34,100.67,99.03,102.57,69.17,55.02,69.88,67.62,50.22], "대전": [55.43,81.30,97.35,110.72,99.01,99.49,106.36,72.01,62.35,72.96,70.60,55.90],
    "대구": [64.90,88.15,90.72,111.86,102.80,102.70,110.62,70.35,51.58,75.80,73.35,60.64], "부산": [68.46,84.08,90.48,105.22,96.64,97.65,107.30,70.11,59.14,74.14,71.75,61.11],
    "울산": [68.46,84.08,90.48,105.22,96.64,97.65,107.30,70.11,59.14,74.14,71.75,61.11], "광주": [66.32,73.60,90.96,107.97,94.75,97.65,109.43,73.67,68.77,75.09,72.67,62.77],
    "경기도": [58.95,88.67,102.91,116.65,108.96,107.74,105.80,80.80,62.91,75.27,72.84,55.53], "강원도": [56.59,96.04,94.22,118.18,108.17,108.50,104.22,79.22,57.82,78.17,75.65,54.48],
    "충청북도": [59.22,88.91,99.49,121.49,111.86,110.03,113.96,76.85,61.13,77.64,75.14,55.80], "충청남도": [61.59,90.33,108.17,123.02,110.01,110.54,118.17,80.01,69.28,81.06,78.45,62.11],
    "경상북도": [72.11,97.94,100.80,124.29,114.22,114.11,122.91,78.17,57.31,84.22,81.50,67.38], "전라북도": [61.06,83.44,105.54,122.00,108.17,113.34,117.91,81.59,72.08,82.12,79.47,63.43],
    "경상남도": [76.06,93.42,100.54,116.91,107.38,108.50,119.23,77.90,65.71,82.38,79.72,67.90], "전라남도": [73.69,81.78,101.06,119.96,105.28,108.50,121.59,81.85,76.41,83.43,80.74,69.75],
    "제주도": [64.22,66.80,87.12,116.14,104.49,92.97,107.64,84.75,78.96,81.85,79.21,66.59]
}

with c1: s_reg = st.selectbox("광역 지자체", list(pv_monthly_data.keys()), index=0)
with c2: s_sub = st.selectbox("기초 지자체", ["전체"])

c3,c4 = st.columns(2)
with c3: h_type = st.selectbox("주거 형태", ["단독 주택 / 다가구 주택","아파트","연립 / 빌라 / 다세대 주택"])
with c4: h_size = st.number_input("전용 면적 (평)", min_value=10, value=30)

# ── 평수 기반 축열조 크기 직관적 안내 ──
if h_size < 20: tank_size, tank_ref = "약 350L", "소형 냉장고 1대 크기 🧊"
elif h_size < 34: tank_size, tank_ref = "약 550L", "워시타워 1대 설치 공간 🧺"
else: tank_size, tank_ref = "약 800L 이상", "약 0.9평의 여유 공간 룸 🚪"

st.markdown(f"<div style='background-color: #f8fafc; border-left: 4px solid #3b82f6; padding: 16px; margin-top: 12px; border-radius: 4px;'><div style='color: #1e293b; font-weight: 700; margin-bottom: 4px; font-size: 1.05rem;'>📐 우리 집 맞춤 설치 공간 안내</div><div style='color: #475569; font-size: 0.95rem;'>입력하신 <b>{h_size}평</b> 기준, <b>{tank_size}</b> 용량의 축열조가 필요합니다. 👉 체감상 <b>{tank_ref}</b>가 필요합니다!</div></div>", unsafe_allow_html=True)

# ── 기후 시각화 ──
st.markdown('<div class="section-title">📊 우리 동네 기후 및 히트펌프 효율 분석</div>', unsafe_allow_html=True)
zone_name = map_region_to_zone(s_reg); dynamic_cop = 3.0
if df_temp and df_cop:
    try:
        dynamic_cop = df_cop[zone_name]['scop']
        days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        all_temps = []
        for h in range(24):
            if h < len(df_temp[zone_name]):
                for m in range(12): all_temps.extend([round(df_temp[zone_name][h][m])] * days[m])
        counts = pd.Series(all_temps).value_counts().sort_index()
        cl1, cl2 = st.columns([2, 1])
        with cl1: st.bar_chart(counts, height=250)
        with cl2: st.success(f"**✅ {s_reg} 맞춤 효율**\n# {dynamic_cop}"); st.caption(f"{s_reg}의 기상 데이터를 반영한 실제 sCOP입니다.")
    except Exception as e: st.error(f"차트 표시 오류: {e}")
else: st.warning("🚨 데이터 파일을 찾을 수 없습니다.")

st.markdown('<div class="section-title">2. 에너지 소비 현황</div>', unsafe_allow_html=True)
heating_sys = st.selectbox("현재 난방 설비", ["가스 컨덴싱 보일러","일반 가스 보일러","등유 보일러","LPG 보일러"])
cv1,cv2 = st.columns(2)
with cv1: w_heat = st.number_input("동절기(1월) 평균 난방비 (만원)", value=20)
with cv2: w_elec = st.number_input("동절기(1월) 전기요금 (만원)", value=6)
ct1,ct2 = st.columns(2)
with ct1: s_capa = st.number_input("태양광 용량 (kW)", value=3.0)
with ct2: elec_tariff = st.selectbox("전기 요금제", ["주택용 누진제 (저압)", "주택용 누진제 (고압)", "계시별 요금제 TOU (제주)"])

st.markdown('<div class="section-title">3. 정책 및 시뮬레이션 변수</div>', unsafe_allow_html=True)
cs1, cs2 = st.columns(2)
with cs1:
    f_inf = st.slider("화석연료 인상률 (%)", 0.0, 15.0, 5.0)
    e_inf = st.slider("전기요금 인상률 (%)", 0.0, 15.0, 3.0)
with cs2:
    sub_nat = st.checkbox("정부 보조금 (320만)", value=True)
    sub_loc = st.checkbox("지자체 보조금 (240만)", value=(s_reg in ["제주도","경상남도","전라남도","부산","울산","광주"]))
    rep_cost = st.number_input("설비 교체비 (만원)", value=100)

if "analyzed" not in st.session_state: st.session_state.analyzed = False
if st.button("경제성 분석 실행", type="primary", use_container_width=True): st.session_state.analyzed = True

# ── 분석 결과 출력 ──
if st.session_state.analyzed:
    hp_kw = estimate_hp_kw(h_size, h_type)
    cur_kwh = reverse_kwh(w_elec, elec_tariff, "ws" if "제주" in elec_tariff else "other", hp_kw)
    hp_jan_kwh = heat_to_hp_kwh(w_heat, heating_sys, cop=dynamic_cop)

    months = list(range(1,13))
    base_kwh = [cur_kwh * (1.15 if m in [7,8] else (1.05 if m in [6,9] else 1.0)) for m in months]
    hp_add_m  = [hp_jan_kwh * HP_MONTHLY_LOAD[m-1] for m in months]
    pv_kwh_m  = [pv_monthly_data[s_reg][m-1]*s_capa for m in months]

    ann_elec = sum(calc_elec_bill(base_kwh[m-1], elec_tariff, MONTH_TOU[m] if "TOU" in elec_tariff else MONTH_SEASON[m], hp_kw) for m in months)
    ann_hp_add = sum((calc_elec_bill(base_kwh[m-1]+hp_add_m[m-1], elec_tariff, MONTH_TOU[m] if "TOU" in elec_tariff else MONTH_SEASON[m], hp_kw) - calc_elec_bill(base_kwh[m-1], elec_tariff, MONTH_TOU[m] if "TOU" in elec_tariff else MONTH_SEASON[m], hp_kw)) for m in months)
    ann_pv = sum(min(pv_kwh_m[m-1], hp_add_m[m-1]) * (214.6/10_000) for m in months)

    ann_heat_base = (w_heat*3) + (w_heat*0.2*9)
    net_cap = max(0, (600+h_size*10)-( (320 if sub_nat else 0) + (240 if sub_loc else 0) ))
    ann_hp_net = max(ann_hp_add - ann_pv, 0)

    years, gas_cum, hp_cum, net_p = list(range(1,16)), [], [], []
    g, h, payback = 0.0, float(net_cap), "15년 초과"
    for y in range(15):
        fi, ei = (1+f_inf/100)**y, (1+e_inf/100)**y
        cg = ann_heat_base*fi + ann_elec*ei + (rep_cost if y==9 else 0)
        ch = ann_elec*ei + ann_hp_net*ei
        g += cg; h += ch
        gas_cum.append(int(g)); hp_cum.append(int(h)); p = int(g-h); net_p.append(p)
        if payback=="15년 초과" and p>0: payback=f"{y+1}년차"

    st.markdown('<div class="section-title">계산 근거 및 요약</div>', unsafe_allow_html=True)
    r1,r2,r3 = st.columns(3)
    r1.metric("초기 투자비", f"{int(net_cap):,} 만원"); r2.metric("투자 회수", payback); r3.metric("15년 순이익", f"{net_p[-1]:,} 만원")

    # 그래프 출력
    cg1,cg2 = st.columns(2)
    with cg1:
        df_a = pd.DataFrame({"연도":years,"기존 유지":gas_cum,"HP 전환":hp_cum}).melt("연도", var_name="시나리오", value_name="비용")
        st.altair_chart(alt.Chart(df_a).mark_area(opacity=0.5).encode(x="연도:O", y="비용:Q", color="시나리오:N"), use_container_width=True)
    with cg2:
        df_c = pd.DataFrame({"연도":years,"순수익":net_p, "상태":["흑자" if p>0 else "회수" for p in net_p]})
        st.altair_chart(alt.Chart(df_c).mark_bar().encode(x="연도:O", y="순수익:Q", color="상태:N"), use_container_width=True)