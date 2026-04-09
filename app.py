import streamlit as st
import pandas as pd
import numpy as np
import io, os, glob
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import altair as alt

st.set_page_config(page_title="히트펌프 경제성 분석 솔루션", layout="wide")

# ── 1. 스타일 및 전체 지자체 데이터 정의 ──
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

# 기초 지자체 전체 리스트 복구
regions_full = {
    "서울": ["종로구","중구","용산구","성동구","광진구","동대문구","중랑구","성북구","강북구","도봉구","노원구","은평구","서대문구","마포구","양천구","강서구","구로구","금천구","영등포구","동작구","관악구","서초구","강남구","송파구","강동구"],
    "인천": ["중구","동구","미추홀구","연수구","남동구","부평구","계양구","서구","강화군","옹진군"],
    "세종": ["세종특별자치시"],
    "대전": ["동구","중구","서구","유성구","대덕구"],
    "대구": ["중구","동구","서구","남구","북구","수성구","달서구","달성군"],
    "부산": ["중구","서구","동구","영도구","부산진구","동래구","남구","북구","해운대구","사하구","금정구","강서구","연제구","수영구","사상구","기장군"],
    "울산": ["중구","남구","동구","북구","울주군"],
    "광주": ["동구","서구","남구","북구","광산구"],
    "경기도": ["수원시","고양시","용인시","성남시","부천시","화성시","안산시","남양주시","안양시","평택시","시흥시","파주시","의정부시","김포시","광주시","광명시","군포시","하남시","오산시","양주시","이천시","구리시","안성시","포천시","의왕시","여주시","동두천시","과천시","가평군","양평군","연천군"],
    "강원도": ["춘천시","원주시","강릉시","동해시","태백시","속초시","삼척시","홍천군","횡성군","영월군","평창군","정선군","철원군","화천군","양구군","인제군","고성군","양양군"],
    "충청북도":["청주시","충주시","제천시","보은군","옥천군","영동군","증평군","진천군","괴산군","음성군","단양군"],
    "충청남도":["천안시","공주시","보령시","아산시","서산시","논산시","계룡시","당진시","금산군","부여군","서천군","청양군","홍성군","예산군","태안군"],
    "전라북도":["전주시","군산시","익산시","정읍시","남원시","김제시","완주군","진안군","무주군","장수군","임실군","순창군","고창군","부안군"],
    "전라남도":["목포시","여수시","순천시","나주시","광양시","담양군","곡성군","구례군","고흥군","보성군","화순군","장흥군","강진군","해남군","영암군","무안군","함평군","영광군","장성군","완도군","진도군","신안군"],
    "경상북도":["포항시","경주시","김천시","안동시","구미시","영주시","영천시","상주시","문경시","경산시","의성군","청송군","영양군","영덕군","청도군","고령군","성주군","칠곡군","예천군","봉화군","울진군","울릉군"],
    "경상남도":["창원시","진주시","통영시","사천시","김해시","밀양시","거제시","양산시","의령군","함안군","창녕군","고성군","남해군","하동군","산청군","함양군","거창군","합천군"],
    "제주도": ["제주시","서귀포시"]
}

pv_monthly_data = {
    "서울": [53.06,79.80,92.62,104.99,98.06,96.96,95.22,72.72,56.62,67.75,65.56,49.98],
    "인천": [53.06,79.80,92.62,104.99,98.06,96.96,95.22,72.72,56.62,67.75,65.56,49.98],
    "세종": [53.30,80.02,89.54,109.34,100.67,99.03,102.57,69.17,55.02,69.88,67.62,50.22],
    "대전": [55.43,81.30,97.35,110.72,99.01,99.49,106.36,72.01,62.35,72.96,70.60,55.90],
    "대구": [64.90,88.15,90.72,111.86,102.80,102.70,110.62,70.35,51.58,75.80,73.35,60.64],
    "부산": [68.46,84.08,90.48,105.22,96.64,97.65,107.30,70.11,59.14,74.14,71.75,61.11],
    "울산": [68.46,84.08,90.48,105.22,96.64,97.65,107.30,70.11,59.14,74.14,71.75,61.11],
    "광주": [66.32,73.60,90.96,107.97,94.75,97.65,109.43,73.67,68.77,75.09,72.67,62.77],
    "경기도": [58.95,88.67,102.91,116.65,108.96,107.74,105.80,80.80,62.91,75.27,72.84,55.53],
    "강원도": [56.59,96.04,94.22,118.18,108.17,108.50,104.22,79.22,57.82,78.17,75.65,54.48],
    "충청북도": [59.22,88.91,99.49,121.49,111.86,110.03,113.96,76.85,61.13,77.64,75.14,55.80],
    "충청남도": [61.59,90.33,108.17,123.02,110.01,110.54,118.17,80.01,69.28,81.06,78.45,62.11],
    "경상북도": [72.11,97.94,100.80,124.29,114.22,114.11,122.91,78.17,57.31,84.22,81.50,67.38],
    "전라북도": [61.06,83.44,105.54,122.00,108.17,113.34,117.91,81.59,72.08,82.12,79.47,63.43],
    "경상남도": [76.06,93.42,100.54,116.91,107.38,108.50,119.23,77.90,65.71,82.38,79.72,67.90],
    "전라남도": [73.69,81.78,101.06,119.96,105.28,108.50,121.59,81.85,76.41,83.43,80.74,69.75],
    "제주도": [64.22,66.80,87.12,116.14,104.49,92.97,107.64,84.75,78.96,81.85,79.21,66.59]
}

# ── 2. 기후 데이터 파싱 및 연동 함수 ──
def map_region_to_zone(s_reg):
    if s_reg == "강원도": return "중부1"
    if s_reg in ["대구","부산","울산","광주","경상남도","전라남도"]: return "남부"
    if s_reg == "제주도": return "제주"
    return "중부2"

@st.cache_data
def load_simulation_data():
    try:
        def read_csv_safe(filename):
            try: return pd.read_csv(filename, encoding="utf-8", header=None)
            except: return pd.read_csv(filename, encoding="cp949", header=None)
        
        df_t = read_csv_safe("외기온도_시간분포.csv")
        df_c = read_csv_safe("COP_계산기.csv")

        # 온도 분포 파싱 (시간별 온도를 24개 행 x 12개 열 리스트로 추출)
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
                vals = [float(v) if pd.notna(v) else 0.0 for v in row[1:13]]
                if len(vals) == 12: temp_data[curr].append(vals)

        # sCOP 파싱 (엑셀 시트에서 '지역' 행 이후의 sCOP 값 추출)
        cop_data = {}
        header_idx = -1
        for idx, row in df_c.iterrows():
            if str(row[0]).strip() == "지역":
                header_idx = idx; break
        if header_idx != -1:
            for _, row in df_c.iloc[header_idx+1:].iterrows():
                z_name = str(row[0]).strip()
                if z_name in target_zones:
                    try: cop_data[z_name] = {'scop': float(row[15])}
                    except: continue

        # 로드 실패 대비 기본값
        defaults = {"중부1": 3.29, "중부2": 3.66, "남부": 3.99, "제주": 4.21}
        for k, v in defaults.items():
            if k not in cop_data: cop_data[k] = {'scop': v}
        return temp_data, cop_data
    except: return None, None

# ── 3. 요금 및 분석 엔진 ──
BOILER_EFF = { "가스 컨덴싱 보일러": 0.92, "일반 가스 보일러": 0.82, "등유 보일러": 0.85, "LPG 보일러": 0.82 }
FUEL_PRICE_PER_MJ = { "가스 컨덴싱 보일러": 68.0, "일반 가스 보일러": 68.0, "등유 보일러": 95.0, "LPG 보일러": 105.0 }
HP_MONTHLY_LOAD = [1.0, 0.9, 0.4, 0.15, 0.05, 0.05, 0.05, 0.05, 0.05, 0.15, 0.4, 0.85]

def calc_elec_bill(kwh, tariff, season="other"):
    kwh = max(kwh, 0)
    if tariff == "주택용 누진제 (저압)":
        if kwh <= 300: base, energy = 910, kwh * 120
        elif kwh <= 450: base, energy = 1_600, 300*120 + (kwh-300)*214.6
        else: base, energy = 7_300, 300*120 + 150*214.6 + (kwh-450)*307.3
    else: base, energy = 910, kwh * 120
    fee = (base + energy + 14.0 * kwh) * 1.127
    return round(fee / 10000, 4)

def reverse_kwh(bill_man, tariff):
    lo, hi = 0.0, 3000.0
    for _ in range(40):
        mid = (lo + hi) / 2
        if calc_elec_bill(mid, tariff) < bill_man: lo = mid
        else: hi = mid
    return round(mid, 1)

def heat_to_hp_kwh(heat_man, boiler_type, cop=3.0):
    eff = BOILER_EFF.get(boiler_type, 0.85); price = FUEL_PRICE_PER_MJ.get(boiler_type, 68.0)
    return round(((heat_man * 10000) / price) / 3.6 * eff / cop, 1)

# ── 4. UI ──
col_t, col_l = st.columns([6,1])
with col_t: st.title("히트펌프 경제성 분석 솔루션")
with col_l:
    if os.path.exists("logo.png"): st.image(Image.open("logo.png"), use_container_width=True)

df_temp, df_cop = load_simulation_data()

st.markdown('<div class="section-title">1. 대상지 기본 정보</div>', unsafe_allow_html=True)
c1,c2 = st.columns(2)
with c1: s_reg = st.selectbox("광역 지자체", list(regions_full.keys()), index=0)
with c2: s_sub = st.selectbox("기초 지자체", regions_full.get(s_reg, ["전체"]))

c3,c4 = st.columns(2)
with c3: h_type = st.selectbox("주거 형태", ["단독 주택 / 다가구 주택","아파트","연립 / 빌라 / 다세대 주택"])
with c4: h_size = st.number_input("전용 면적 (평)", min_value=10, value=30)

# 기후 시각화
zone = map_region_to_zone(s_reg); dynamic_cop = 3.0
st.markdown('<div class="section-title">📊 우리 동네 기후 및 히트펌프 효율 분석</div>', unsafe_allow_html=True)
if df_temp and df_cop:
    dynamic_cop = df_cop[zone]['scop']
    all_temps = []
    days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    for h in range(len(df_temp[zone])):
        for m in range(12): all_temps.extend([round(df_temp[zone][h][m])] * days[m])
    
    counts = pd.Series(all_temps).value_counts().sort_index().reset_index()
    counts.columns = ['온도', '시간']
    
    cl1, cl2 = st.columns([2, 1])
    with cl1:
        st.markdown(f"**🌡️ {s_reg} 연간 온도별 발생 시간 분포**")
        c = alt.Chart(counts).mark_bar(color='#3b82f6').encode(
            x=alt.X('온도:Q', title='외기 온도 (°C)'),
            y=alt.Y('시간:Q', title='연간 누적 시간 (hours)'),
            tooltip=['온도', '시간']
        ).properties(height=250)
        st.altair_chart(c, use_container_width=True)
    with cl2:
        st.success(f"**✅ [{s_reg}] 맞춤 효율(sCOP)**\n# {dynamic_cop}")
        st.caption(f"{s_reg}의 기상 데이터를 반영한 실제 체감 효율입니다.")

st.markdown('<div class="section-title">2. 에너지 소비 현황</div>', unsafe_allow_html=True)
heating_sys = st.selectbox("현재 난방 설비", list(BOILER_EFF.keys()))
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
    sub_nat = st.checkbox("정부 보조금 적용", value=True)
    sub_loc = st.checkbox("지자체 보조금 적용", value=(s_reg in ["제주도","경상남도","전라남도","부산","울산","광주"]))
    rep_cost = st.number_input("설비 교체비 (만원)", value=100)

if "analyzed" not in st.session_state: st.session_state.analyzed = False
if st.button("경제성 분석 실행", type="primary", use_container_width=True): st.session_state.analyzed = True

# ── 5. 결과 분석 및 엑셀 다운로드 ──
if st.session_state.analyzed:
    hp_jan_kwh = heat_to_hp_kwh(w_heat, heating_sys, cop=dynamic_cop)
    cur_kwh = reverse_kwh(w_elec, "주택용 누진제 (저압)")
    
    # 15년 시뮬레이션
    years, gas_cum, hp_cum, net_p = list(range(1,16)), [], [], []
    ann_heat_base = w_heat * 4.5 # 연간 추정 난방비
    ann_elec_base = calc_elec_bill(cur_kwh, "주택용 누진제 (저압)") * 12
    net_cap = 600 + h_size * 10 - ( (320 if sub_nat else 0) + (240 if sub_loc else 0) )
    g, h, payback = 0.0, float(net_cap), "15년 초과"
    
    for y in years:
        cg = ann_heat_base * ((1 + f_inf/100)**y) + ann_elec_base
        ch = (ann_heat_base/dynamic_cop)*3.2 * ((1 + e_inf/100)**y) + ann_elec_base
        g += cg; h += ch; p = int(g-h)
        gas_cum.append(int(g)); hp_cum.append(int(h)); net_p.append(p)
        if payback=="15년 초과" and p>0: payback=f"{y}년차"

    st.markdown('<div class="section-title">분석 결과 요약</div>', unsafe_allow_html=True)
    ca, cb, cc = st.columns(3)
    ca.metric("투자 회수 시점", payback); cb.metric("15년 누적 순이익", f"{net_p[-1]:,} 만원"); cc.metric("적용 sCOP", f"{dynamic_cop}")

    # 그래프 출력
    g1, g2 = st.columns(2)
    with g1:
        df_a = pd.DataFrame({"연도":years, "기존":gas_cum, "HP":hp_cum}).melt("연도", var_name="시나리오", value_name="비용")
        st.altair_chart(alt.Chart(df_a).mark_area(opacity=0.5).encode(x="연도:O", y="비용:Q", color="시나리오:N"), use_container_width=True)
    with g2:
        df_c = pd.DataFrame({"연도":years, "순수익":net_p, "상태":["수익" if p>0 else "회수" for p in net_p]})
        st.altair_chart(alt.Chart(df_c).mark_bar().encode(x="연도:O", y="순수익:Q", color="상태:N"), use_container_width=True)

    # ════ 인터랙티브 엑셀 생성 ════
    wb = Workbook()
    ws = wb.active; ws.title = "①입력_가정"
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    
    ws.merge_cells("A1:D1"); ws["A1"] = f"히트펌프 경제성 분석 가정 ({s_reg})"; ws["A1"].fill = header_fill; ws["A1"].font = white_font
    
    inputs = [("항목", "값", "단위", "비고"), ("1월 난방비", w_heat, "만원", "사용자 입력"), ("1월 전기요금", w_elec, "만원", "사용자 입력"),
              ("맞춤형 sCOP", dynamic_cop, "-", "기상데이터 연동"), ("순 설치비(CAPEX)", net_cap, "만원", "보조금 차감후")]
    for r, row in enumerate(inputs, 3):
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=val)
            if r == 3: cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid"); cell.font = Font(bold=True)
    
    ws2 = wb.create_sheet("②15년_시뮬레이션")
    ws2.append(["경과연도", "기존설비 누적비용", "히트펌프 누적비용", "누적 이익(NPV)"])
    for y in years:
        ws2.append([y, gas_cum[y-1], hp_cum[y-1], f"=B{y+1}-C{y+1}"]) # 엑셀 내 수식 연동
    
    buf = io.BytesIO(); wb.save(buf)
    st.markdown("---")
    st.download_button(
        label="📊 수식 연동 인터랙티브 엑셀 파일 다운로드",
        data=buf.getvalue(),
        file_name=f"HeatPump_Analysis_{s_reg}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )