import streamlit as st
import pandas as pd
import io, os, glob
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import altair as alt

st.set_page_config(page_title="히트펌프 경제성 분석 솔루션", layout="wide")

st.markdown("""
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
* { font-family: 'Pretendard', sans-serif; }
.info-box   { background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:24px; margin-bottom:32px; }
.info-title { color:#0f172a; font-size:1.15rem; font-weight:700; margin-bottom:12px; margin-top:0; }
.info-text  { color:#475569; font-size:0.95rem; line-height:1.6; margin-bottom:0; }
.section-title { color:#1e293b; font-weight:700; font-size:1.3rem; margin-top:40px; margin-bottom:16px;
                 border-bottom:2px solid #cbd5e1; padding-bottom:8px; }
.calc-box   { background:#f0f9ff; border:1px solid #bae6fd; border-radius:8px; padding:16px; margin:12px 0; }
.calc-label { color:#0369a1; font-size:0.85rem; font-weight:600; margin-bottom:4px; }
.calc-value { color:#0f172a; font-size:1.05rem; font-weight:700; }
</style>
""", unsafe_allow_html=True)

# ── 지역별 태양광 발전량 ─────────────────────────────────────
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
regions_full = {
    "서울":    ["종로구","중구","용산구","성동구","광진구","동대문구","중랑구","성북구",
               "강북구","도봉구","노원구","은평구","서대문구","마포구","양천구","강서구",
               "구로구","금천구","영등포구","동작구","관악구","서초구","강남구","송파구","강동구"],
    "인천":    ["중구","동구","미추홀구","연수구","남동구","부평구","계양구","서구","강화군","옹진군"],
    "세종":    ["세종특별자치시"],
    "대전":    ["동구","중구","서구","유성구","대덕구"],
    "대구":    ["중구","동구","서구","남구","북구","수성구","달서구","달성군"],
    "부산":    ["중구","서구","동구","영도구","부산진구","동래구","남구","북구",
               "해운대구","사하구","금정구","강서구","연제구","수영구","사상구","기장군"],
    "울산":    ["중구","남구","동구","북구","울주군"],
    "광주":    ["동구","서구","남구","북구","광산구"],
    "경기도":  ["수원시","고양시","용인시","성남시","부천시","화성시","안산시","남양주시",
               "안양시","평택시","시흥시","파주시","의정부시","김포시","광주시","광명시",
               "군포시","하남시","오산시","양주시","이천시","구리시","안성시","포천시",
               "의왕시","여주시","동두천시","과천시","가평군","양평군","연천군"],
    "강원도":  ["춘천시","원주시","강릉시","동해시","태백시","속초시","삼척시",
               "홍천군","횡성군","영월군","평창군","정선군","철원군","화천군","양구군","인제군","고성군","양양군"],
    "충청북도":["청주시","충주시","제천시","보은군","옥천군","영동군","증평군","진천군","괴산군","음성군","단양군"],
    "충청남도":["천안시","공주시","보령시","아산시","서산시","논산시","계룡시","당진시",
               "금산군","부여군","서천군","청양군","홍성군","예산군","태안군"],
    "전라북도":["전주시","군산시","익산시","정읍시","남원시","김제시",
               "완주군","진안군","무주군","장수군","임실군","순창군","고창군","부안군"],
    "전라남도":["목포시","여수시","순천시","나주시","광양시",
               "담양군","곡성군","구례군","고흥군","보성군","화순군","장흥군","강진군",
               "해남군","영암군","무안군","함평군","영광군","장성군","완도군","진도군","신안군"],
    "경상북도":["포항시","경주시","김천시","안동시","구미시","영주시","영천시","상주시","문경시","경산시",
               "군위군","의성군","청송군","영양군","영덕군","청도군","고령군","성주군","칠곡군","예천군","봉화군","울진군","울릉군"],
    "경상남도":["창원시","진주시","통영시","사천시","김해시","밀양시","거제시","양산시",
               "의령군","함안군","창녕군","고성군","남해군","하동군","산청군","함양군","거창군","합천군"],
    "제주도":  ["제주시","서귀포시"],
}

# ════════════════════════════════════════════════════════════
# 전기요금 계산 엔진 (2026.04.16 한전 시행 요금 기준)
# ════════════════════════════════════════════════════════════
CLIMATE_FEE = 9.0    # 기후환경요금 (원/kWh)
FUEL_ADJ    = 5.0    # 연료비조정요금 (원/kWh)
VAT_RATE    = 0.10
FUND_RATE   = 0.027  # 전력산업기반기금 (2025.7.1 인하 후)

BOILER_EFF = {
    "가스 컨덴싱 보일러": 0.92,
    "일반 가스 보일러":   0.82,
    "등유 보일러":        0.85,
    "LPG 보일러":         0.82,
}
FUEL_PRICE_PER_MJ = {
    "가스 컨덴싱 보일러": 68.0,
    "일반 가스 보일러":   68.0,
    "등유 보일러":        95.0,
    "LPG 보일러":        105.0,
}
HP_COP_WINTER = 3.0
HP_MONTHLY_LOAD = [1.0,0.9,0.4,0.15,0.05,0.05,0.05,0.05,0.05,0.15,0.4,0.85]

MONTH_SEASON = {1:"other",2:"other",3:"other",4:"other",5:"other",6:"other",
                7:"summer",8:"summer",9:"other",10:"other",11:"other",12:"other"}
MONTH_TOU    = {1:"ws",2:"ws",3:"sf",4:"sf",5:"sf",6:"ws",
                7:"ws",8:"ws",9:"sf",10:"sf",11:"ws",12:"ws"}


def calc_elec_bill(kwh, tariff, season="other", contract_kw=5.0):
    """kWh → 월 청구금액 (만원)"""
    kwh = max(kwh, 0)
    if tariff == "주택용 누진제 (저압)":
        if kwh <= 300:
            base, energy = 910, kwh * 120
        elif kwh <= 450:
            base  = 1_600
            energy = 300*120 + (kwh-300)*214.6
        else:
            base  = 7_300
            energy = 300*120 + 150*214.6 + (kwh-450)*307.3

    elif tariff == "주택용 누진제 (고압)":
        if season == "summer":
            if kwh <= 300:   base, energy = 730, kwh*105
            elif kwh <= 450: base=1_260; energy=300*105+(kwh-300)*174
            else:            base=6_060; energy=300*105+150*174+(kwh-450)*242.3
        else:
            if kwh <= 200:   base, energy = 730, kwh*105
            elif kwh <= 400: base=1_260; energy=200*105+(kwh-200)*174
            else:            base=6_060; energy=200*105+200*174+(kwh-400)*242.3

    elif tariff == "계시별 요금제 TOU (제주)":
        base = 4_310 * contract_kw
        r = {"경":138.7,"중":184.7,"최":220.5} if season=="ws" else {"경":125.8,"중":153.8,"최":172.4}
        energy = kwh * (0.4*r["경"] + 0.4*r["중"] + 0.2*r["최"])

    else:  # 자동 → 저압 누진제
        return calc_elec_bill(kwh, "주택용 누진제 (저압)", season, contract_kw)

    fee   = base + energy + (CLIMATE_FEE + FUEL_ADJ) * kwh
    total = fee * (1 + VAT_RATE + FUND_RATE)
    return round(total / 10_000, 4)


def reverse_kwh(bill_man, tariff, season="other", contract_kw=5.0):
    """청구금액(만원) → 추정 kWh (이진탐색 역산)"""
    lo, hi = 0.0, 3_000.0
    for _ in range(60):
        mid = (lo + hi) / 2
        if calc_elec_bill(mid, tariff, season, contract_kw) < bill_man:
            lo = mid
        else:
            hi = mid
    return round(mid, 1)


def heat_to_hp_kwh(heat_man, boiler_type, cop=HP_COP_WINTER):
    """난방비(만원) → HP 월 전력소비 추정 (kWh)"""
    eff      = BOILER_EFF.get(boiler_type, 0.85)
    price_mj = FUEL_PRICE_PER_MJ.get(boiler_type, 68.0)
    heat_mj  = (heat_man * 10_000) / price_mj
    useful   = heat_mj / 3.6 * eff
    return round(useful / cop, 1)


def estimate_hp_kw(size, htype):
    """평수·주거형태 → HP 용량 추정 (kW)"""
    rate = 0.15 if "아파트" in htype else (0.17 if "연립" in htype else 0.20)
    return round(size * rate, 1)

# ════════════════════════════════════════════════════════════
# UI
# ════════════════════════════════════════════════════════════
col_t, col_l = st.columns([6,1])
with col_t: st.title("히트펌프 경제성 분석 솔루션")
with col_l:
    imgs = glob.glob("*.jpg")+glob.glob("*.jpeg")+glob.glob("*.png")
    if imgs: st.image(Image.open(imgs[0]), use_container_width=True)

st.markdown("""
<div class='info-box'>
  <h4 class='info-title'>솔루션 개요</h4>
  <p class='info-text'>
    고객의 주거 환경 및 에너지 사용량 데이터를 기반으로 공기열 히트펌프(AWHP) 도입 시 재무적 타당성을 분석합니다.<br>
    <b>2026.04.16 시행 한국전력 요금표</b>를 내장하여 전기요금 고지서 금액만으로 kWh·HP 용량을 자동 역산합니다.
  </p>
</div>
""", unsafe_allow_html=True)

# 1. 기본 정보
st.markdown('<div class="section-title">1. 대상지 기본 정보</div>', unsafe_allow_html=True)
c1,c2 = st.columns(2)
with c1: s_reg  = st.selectbox("광역 지자체", list(pv_monthly_data.keys()), index=8)
with c2: s_sub  = st.selectbox("기초 지자체", regions_full.get(s_reg, ["전체"]))
c3,c4 = st.columns(2)
with c3: h_type = st.selectbox("주거 형태", ["단독 주택 / 다가구 주택","아파트","연립 / 빌라 / 다세대 주택"])
with c4: h_size = st.number_input("전용 면적 (평)", min_value=10, value=30)

# 2. 에너지 현황
st.markdown('<div class="section-title">2. 에너지 소비 및 인프라 현황</div>', unsafe_allow_html=True)
heating_sys = st.selectbox("현재 난방 설비",
              ["가스 컨덴싱 보일러","일반 가스 보일러","등유 보일러","LPG 보일러"])
cv1,cv2 = st.columns(2)
with cv1: w_heat  = st.number_input("동절기(1월) 평균 난방비 (만원)", value=20)
with cv2: w_elec  = st.number_input("동절기(1월) 전기요금 고지서 금액 (만원)", value=6,
             help="청구서의 실 납부금액을 입력하세요. kWh는 자동 역산됩니다.")
ct1,ct2 = st.columns(2)
with ct1: s_capa = st.number_input("태양광 발전 설비 용량 (kW)", value=3.0, step=0.5)
with ct2: elec_tariff = st.selectbox("적용 전기 요금제",
             ["최적 자동 산출 (저압 누진제)",
              "주택용 누진제 (저압)",
              "주택용 누진제 (고압)",
              "계시별 요금제 TOU (제주)"])

# 3. 정책 변수
st.markdown('<div class="section-title">3. 정책 및 시뮬레이션 변수 설정</div>', unsafe_allow_html=True)
if h_size<=32 or "아파트" in h_type:
    st.caption("※ 분석 기준: 100~300L(소형) 축열조 적용 (필요 면적 약 0.5㎡)")
else:
    st.caption("※ 분석 기준: 300~500L(중대형) 축열조 적용 (필요 면적 약 1.5㎡)")

with st.expander("상세 분석 파라미터 (2026.03 정책 기준)", expanded=True):
    cs1,cs2 = st.columns(2)
    with cs1:
        f_inf    = st.slider("화석연료 물가 인상률 (연평균, %)", 0.0, 15.0, 5.0)
        e_inf    = st.slider("전기요금 물가 인상률 (연평균, %)", 0.0, 15.0, 3.0)
        apply_ev = st.checkbox("EV 특례 요금 수준 할인 적용", value=True)
    with cs2:
        sub_nat   = st.checkbox("정부 무상 보조금 적용 (최대 320만원)", value=True)
        is_south  = s_reg in ["제주도","경상남도","전라남도","부산","울산","광주"]
        sub_loc   = st.checkbox("지자체 매칭 보조금 적용 (최대 240만원)", value=is_south)
        if not is_south: st.caption("참고: 지자체 보조금은 현재 남부권역을 중심으로 우선 배정됩니다.")
        rep_cost  = st.number_input("기존 설비 10년 차 교체 충당금 (만원)", value=100)

st.markdown("<br>", unsafe_allow_html=True)

if "analyzed" not in st.session_state:
    st.session_state.analyzed = False
if st.button("경제성 분석 실행", type="primary", use_container_width=True):
    st.session_state.analyzed = True

# ════════════════════════════════════════════════════════════
# 분석 실행
# ════════════════════════════════════════════════════════════
if st.session_state.analyzed:

    # 요금제 매핑
    tariff = {
        "최적 자동 산출 (저압 누진제)": "주택용 누진제 (저압)",
        "주택용 누진제 (저압)":         "주택용 누진제 (저압)",
        "주택용 누진제 (고압)":         "주택용 누진제 (고압)",
        "계시별 요금제 TOU (제주)":     "계시별 요금제 TOU (제주)",
    }[elec_tariff]

    # HP 용량 자동 추정
    hp_kw = estimate_hp_kw(h_size, h_type)

    # 1월 기준 시즌
    jan_s = MONTH_TOU[1] if "TOU" in tariff else MONTH_SEASON[1]

    # 현재 kWh 역산
    cur_kwh = reverse_kwh(w_elec, tariff, jan_s, hp_kw)

    # 1월 HP 추가 전력
    hp_jan_kwh = heat_to_hp_kwh(w_heat, heating_sys)

    ev_disc = 0.60 if apply_ev else 1.0

    # 월별 계산
    reg_pv = pv_monthly_data[s_reg]
    months = list(range(1,13))

    # 월별 기본 전기사용량 추정
    base_kwh = []
    for m in months:
        mult = 1.15 if m in [7,8] else (1.05 if m in [6,9] else 1.0)
        base_kwh.append(cur_kwh * mult)

    pv_kwh_m  = [reg_pv[m-1]*s_capa for m in months]
    hp_add_m  = [hp_jan_kwh * HP_MONTHLY_LOAD[m-1] for m in months]

    # 연간 기존 전기요금
    ann_elec = sum(
        calc_elec_bill(base_kwh[m-1], tariff,
                       MONTH_TOU[m] if "TOU" in tariff else MONTH_SEASON[m], hp_kw)
        for m in months
    )
    # 연간 HP 추가 전기요금 (EV 할인)
    ann_hp_add = sum(
        (calc_elec_bill(base_kwh[m-1]+hp_add_m[m-1], tariff,
                        MONTH_TOU[m] if "TOU" in tariff else MONTH_SEASON[m], hp_kw)
         - calc_elec_bill(base_kwh[m-1], tariff,
                          MONTH_TOU[m] if "TOU" in tariff else MONTH_SEASON[m], hp_kw))
        * ev_disc
        for m in months
    )
    # 연간 PV 절감
    ann_pv = sum(
        min(pv_kwh_m[m-1], hp_add_m[m-1]) * (214.6/10_000)
        for m in months
    )

    ann_heat_base = (w_heat*3) + (w_heat*0.2*9)
    nat_sub = 320 if sub_nat else 0
    loc_sub = 240 if sub_loc else 0
    net_cap = max(0, (600+h_size*10)-(nat_sub+loc_sub))
    ann_hp_net = max(ann_hp_add - ann_pv, 0)

    # 15년 누적
    years, gas_cum, hp_cum, net_p = list(range(1,16)), [], [], []
    g, h = 0.0, float(net_cap)
    payback = "15년 초과"
    for y in range(15):
        fi, ei = (1+f_inf/100)**y, (1+e_inf/100)**y
        cg = ann_heat_base*fi + ann_elec*ei + (rep_cost if y==9 else 0)
        ch = ann_elec*ei + ann_hp_net*ei
        g += cg; h += ch
        gas_cum.append(int(g)); hp_cum.append(int(h))
        p = int(g-h); net_p.append(p)
        if payback=="15년 초과" and p>0: payback=f"{y+1}년차"

    # ── 계산 근거 표시 ────────────────────────────────────
    st.markdown('<div class="section-title" style="margin-top:20px;">계산 근거 (자동 추정 결과)</div>',
                unsafe_allow_html=True)
    ca,cb,cc,cd = st.columns(4)
    with ca: st.markdown(f"""
<div class='calc-box'><div class='calc-label'>1월 추정 전기사용량</div>
<div class='calc-value'>{cur_kwh:,.0f} kWh</div>
<div style='color:#64748b;font-size:0.8rem'>요금 {w_elec}만원 역산 ({tariff})</div></div>""",
        unsafe_allow_html=True)
    with cb: st.markdown(f"""
<div class='calc-box'><div class='calc-label'>추정 히트펌프 용량</div>
<div class='calc-value'>{hp_kw} kW</div>
<div style='color:#64748b;font-size:0.8rem'>{h_size}평 × {h_type[:2]} 환경 계수</div></div>""",
        unsafe_allow_html=True)
    with cc: st.markdown(f"""
<div class='calc-box'><div class='calc-label'>1월 HP 추가 전력</div>
<div class='calc-value'>{hp_jan_kwh:,.0f} kWh</div>
<div style='color:#64748b;font-size:0.8rem'>난방비 {w_heat}만원 역산 (COP {HP_COP_WINTER})</div></div>""",
        unsafe_allow_html=True)
    with cd: st.markdown(f"""
<div class='calc-box'><div class='calc-label'>연간 PV 절감 (추정)</div>
<div class='calc-value'>{ann_pv:,.1f} 만원</div>
<div style='color:#64748b;font-size:0.8rem'>{s_reg} {s_capa}kW 기준</div></div>""",
        unsafe_allow_html=True)

    # ── Executive Summary ──────────────────────────────
    st.markdown('<div class="section-title">분석 결과 요약 (Executive Summary)</div>',
                unsafe_allow_html=True)
    r1,r2,r3 = st.columns(3)
    r1.metric("초기 투자 비용 (CAPEX)", f"{int(net_cap):,} 만원", f"보조금 {nat_sub+loc_sub}만원 차감")
    r2.metric("투자 회수 시점 (Payback Period)", payback)
    r3.metric("15년 누적 순이익 (NPV)", f"{net_p[-1]:,} 만원")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── 그래프 ────────────────────────────────────────────
    cg1,cg2 = st.columns(2)
    with cg1:
        st.markdown("**15년 누적 비용 흐름 (Cumulative Cost)**")
        df_a = pd.DataFrame({"연도":years,"기존 설비 유지":gas_cum,"히트펌프 전환":hp_cum}
                            ).melt("연도", var_name="시나리오", value_name="누적비용")
        st.altair_chart(
            alt.Chart(df_a).mark_area(opacity=0.55).encode(
                x=alt.X("연도:O", title="경과 연도", axis=alt.Axis(labelAngle=0)),
                y=alt.Y("누적비용:Q", title="비용 (만원)"),
                color=alt.Color("시나리오:N",
                    scale=alt.Scale(range=["#94a3b8","#0284c7"]),
                    legend=alt.Legend(orient="top-left", title=None)),
            ).properties(height=350),
            use_container_width=True)

    with cg2:
        st.markdown("**투자 회수 및 현금흐름 (Cash Flow)**")
        df_c = pd.DataFrame({"연도":years,"순수익":net_p,
                             "상태":["흑자 전환" if p>0 else "회수 중" for p in net_p]})
        st.altair_chart(
            alt.Chart(df_c).mark_bar().encode(
                x=alt.X("연도:O", title="경과 연도", axis=alt.Axis(labelAngle=0)),
                y=alt.Y("순수익:Q", title="현금흐름 (만원)"),
                color=alt.Color("상태:N",
                    scale=alt.Scale(domain=["회수 중","흑자 전환"],range=["#94a3b8","#0284c7"]),
                    legend=alt.Legend(orient="top-left", title=None)),
                tooltip=[alt.Tooltip("연도:O",title="경과 연도"),
                         alt.Tooltip("순수익:Q",title="순수익 (만원)",format=",")],
            ).properties(height=350),
            use_container_width=True)

    # ── 월별 비교표 ──────────────────────────────────────
    st.markdown("**월별 전기요금 상세 비교 (히트펌프 전환 전·후)**")
    mn = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"]
    bef, aft, hpa, pvs = [], [], [], []
    for m in months:
        s = MONTH_TOU[m] if "TOU" in tariff else MONTH_SEASON[m]
        b = calc_elec_bill(base_kwh[m-1], tariff, s, hp_kw)
        a = (calc_elec_bill(base_kwh[m-1]+hp_add_m[m-1], tariff, s, hp_kw)
             * ev_disc + b*(1-ev_disc))
        pv = min(pv_kwh_m[m-1], hp_add_m[m-1]) * (214.6/10_000)
        bef.append(round(b,2)); aft.append(round(a,2))
        hpa.append(round(hp_add_m[m-1],0)); pvs.append(round(pv,2))

    st.dataframe(pd.DataFrame({
        "월": mn,
        "현재 전기요금 (만원)": bef,
        "전환 후 전기요금 (만원)": aft,
        "HP 추가 전력 (kWh)": hpa,
        "PV 절감 (만원)": pvs,
    }), hide_index=True, use_container_width=True)

    # ── 수식 연동형 인터랙티브 엑셀 다운로드 ─────────────────
    wb   = Workbook()
    hf   = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    hf2  = PatternFill(start_color="0284C7", end_color="0284C7", fill_type="solid")
    fw   = Font(color="FFFFFF", bold=True)
    fb   = Font(bold=True)
    fi   = Font(color="0000FF", bold=True)   # 파란색 = 수정 가능 입력값
    fk   = Font(color="000000")              # 검정 = 수식
    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"),   bottom=Side(style="thin"))
    center = Alignment(horizontal="center")
    right  = Alignment(horizontal="right")

    # ════ 시트1: 입력 가정 (수정 가능한 파란색 셀) ════
    ws = wb.active; ws.title = "①입력_가정"

    ws.merge_cells("A1:D1")
    ws["A1"] = f"히트펌프 경제성 분석 — 입력 가정  ({s_reg} {s_sub})"
    ws["A1"].fill = hf; ws["A1"].font = fw; ws["A1"].alignment = center

    ws["A2"] = "※ 파란색 셀은 직접 수정 가능합니다. 수정 후 시트2·3 결과가 자동 연동됩니다."
    ws["A2"].font = Font(color="0369A1", italic=True)

    headers3 = ["항목", "값", "단위", "비고"]
    for ci, h in enumerate(headers3, 1):
        c = ws.cell(row=3, column=ci, value=h)
        c.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        c.font = fb; c.border = thin

    # 입력값 행 (파란색 수정 가능)
    input_rows = [
        # (항목명,          값,              단위,    비고)
        ("동절기 난방비",        w_heat,      "만원",  "1월 기준 직접 입력"),
        ("동절기 전기요금",      w_elec,      "만원",  "1월 청구서 금액"),
        ("1월 추정 전기사용량",  cur_kwh,     "kWh",   f"{tariff} 역산값"),
        ("추정 HP 용량",         hp_kw,       "kW",    f"{h_size}평 × 환경계수 자동추정"),
        ("1월 HP 추가 전력",     hp_jan_kwh,  "kWh",   f"COP {HP_COP_WINTER} 기준 역산"),
        ("태양광 설비 용량",     s_capa,      "kW",    "직접 입력"),
        ("연간 PV 발전량",       round(sum(reg_pv)*s_capa, 1), "kWh", f"{s_reg} 지역 실측 기준"),
        ("화석연료 물가 인상률", f_inf/100,   "%",     "연평균 복리"),
        ("전기요금 물가 인상률", e_inf/100,   "%",     "연평균 복리"),
        ("설비 CAPEX",           600+h_size*10, "만원","설치 전 총비용"),
        ("정부 보조금",          nat_sub,     "만원",  "무상 보조"),
        ("지자체 보조금",        loc_sub,     "만원",  "매칭 보조"),
        ("순 CAPEX (보조금 차감)", net_cap,   "만원",  "=B13-B14-B15"),
        ("기존 설비 교체비",     rep_cost,    "만원",  "10년차 발생"),
        ("기후환경요금",         CLIMATE_FEE, "원/kWh","한전 고시"),
        ("연료비조정요금",       FUEL_ADJ,    "원/kWh","한전 상한 기준"),
        ("부가세",               VAT_RATE,    "%",     ""),
        ("전력산업기반기금",     FUND_RATE,   "%",     "2025.7 인하 후"),
    ]
    for ri, (name, val, unit, note) in enumerate(input_rows, 4):
        ws.cell(row=ri, column=1, value=name).border = thin
        vc = ws.cell(row=ri, column=2, value=val)
        vc.font = fi; vc.border = thin; vc.alignment = right
        ws.cell(row=ri, column=3, value=unit).border = thin
        ws.cell(row=ri, column=4, value=note).border = thin

    # 수식으로 순CAPEX 연동
    ws["B16"] = "=B13-B14-B15"
    ws["B16"].font = fk; ws["B16"].border = thin

    # % 서식
    for r in [11, 12, 20, 21]:
        ws.cell(row=r, column=2).number_format = "0.0%"

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 28

    # ════ 시트2: 월별 전기요금 비교 ════
    ws2 = wb.create_sheet("②월별_전기요금")

    ws2.merge_cells("A1:F1")
    ws2["A1"] = "월별 전기요금 상세 비교 (히트펌프 전환 전·후)"
    ws2["A1"].fill = hf2; ws2["A1"].font = fw; ws2["A1"].alignment = center

    hdr2 = ["월", "현재 전기요금(만원)", "전환 후 전기요금(만원)", "절감액(만원)", "HP 추가전력(kWh)", "PV 절감(만원)"]
    for ci, h in enumerate(hdr2, 1):
        c = ws2.cell(row=2, column=ci, value=h)
        c.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
        c.font = fb; c.border = thin; c.alignment = Alignment(horizontal="center")

    for ri, (a, b, c, d, e) in enumerate(zip(mn, bef, aft, hpa, pvs), 3):
        saving = round(b - c, 2)
        for ci, v in enumerate([a, b, c, saving, d, e], 1):
            cell = ws2.cell(row=ri, column=ci, value=v)
            cell.border = thin
            if ci in [2, 3, 4, 6]: cell.number_format = "0.00"
            if ci == 4 and saving > 0: cell.font = Font(color="0284C7", bold=True)

    # 합계 행
    ws2.cell(row=15, column=1, value="연간 합계").font = fb
    for ci, col in enumerate(["B","C","D","E","F"], 2):
        c = ws2.cell(row=15, column=ci, value=f"=SUM({col}3:{col}14)")
        c.font = fk; c.border = thin; c.number_format = "0.00"
    ws2.cell(row=15, column=1).border = thin

    for col_letter, w in zip(["A","B","C","D","E","F"], [8,16,18,12,14,12]):
        ws2.column_dimensions[col_letter].width = w

    # ════ 시트3: 15년 누적 현금흐름 (수식 연동) ════
    ws3 = wb.create_sheet("③15년_재무시뮬레이션")

    ws3.merge_cells("A1:G1")
    ws3["A1"] = "15년 장기 재무 시뮬레이션 — 입력 가정 시트 연동형"
    ws3["A1"].fill = hf; ws3["A1"].font = fw; ws3["A1"].alignment = center

    ws3["A2"] = "※ 입력값은 [①입력_가정] 시트에서 수정하세요. 아래 수식이 자동 재계산됩니다."
    ws3["A2"].font = Font(color="0369A1", italic=True)

    hdr3 = ["경과연도", "기존설비 연간 OPEX(만원)", "히트펌프 연간 TCO(만원)",
            "연간 절감액(만원)", "기존설비 누적 OPEX(만원)", "히트펌프 누적 TCO(만원)", "누적 순이익 NPV(만원)"]
    for ci, h in enumerate(hdr3, 1):
        c = ws3.cell(row=3, column=ci, value=h)
        c.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        c.font = fw; c.border = thin; c.alignment = Alignment(horizontal="center")

    # 참조 셀 정의 (①입력_가정 시트 연동)
    REF = "\'①입력_가정\'!"

    for y in range(1, 16):
        r = y + 3   # 엑셀 행 번호
        prev = r - 1

        fi_formula  = f"(1+{REF}B11)^{y-1}"
        ei_formula  = f"(1+{REF}B12)^{y-1}"
        rep_formula = f"+{REF}B17" if y == 10 else ""

        ann_gas_f = (f"=({REF}B4*6.5*{fi_formula})"
                     f"+({REF}B5*12*{ei_formula}){rep_formula}")
        ann_hp_f  = (f"=({REF}B5*12*{ei_formula})"
                     f"+({REF}B5*2.5*{ei_formula})"
                     f"-({REF}B7/10000*200*{ei_formula})")

        ws3.cell(row=r, column=1, value=y).border = thin

        c_gas = ws3.cell(row=r, column=2, value=ann_gas_f)
        c_gas.font = fk; c_gas.border = thin; c_gas.number_format = "#,##0"

        c_hp = ws3.cell(row=r, column=3, value=ann_hp_f)
        c_hp.font = fk; c_hp.border = thin; c_hp.number_format = "#,##0"

        c_sav = ws3.cell(row=r, column=4, value=f"=B{r}-C{r}")
        c_sav.font = fk; c_sav.border = thin; c_sav.number_format = "#,##0"

        if y == 1:
            cum_g_f = f"=B{r}"
            cum_h_f = f"={REF}B16+C{r}"
        else:
            cum_g_f = f"=E{prev}+B{r}"
            cum_h_f = f"=F{prev}+C{r}"

        c_cg = ws3.cell(row=r, column=5, value=cum_g_f)
        c_cg.font = fk; c_cg.border = thin; c_cg.number_format = "#,##0"

        c_ch = ws3.cell(row=r, column=6, value=cum_h_f)
        c_ch.font = fk; c_ch.border = thin; c_ch.number_format = "#,##0"

        c_np = ws3.cell(row=r, column=7, value=f"=E{r}-F{r}")
        c_np.font = fk; c_np.border = thin; c_np.number_format = "#,##0"
        if net_p[y-1] > 0:
            c_np.fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")

    ws3.cell(row=19, column=1, value="15년 합계").font = fb
    ws3.cell(row=19, column=1).border = thin
    for ci, col in enumerate(["B","C","D"], 2):
        c = ws3.cell(row=19, column=ci, value=f"=SUM({col}4:{col}18)")
        c.font = fk; c.border = thin; c.number_format = "#,##0"

    for col_letter, w in zip(["A","B","C","D","E","F","G"], [8,16,16,14,18,18,18]):
        ws3.column_dimensions[col_letter].width = w
    ws3.row_dimensions[3].height = 42

    buf = io.BytesIO(); wb.save(buf)
    st.download_button("📊 수식 연동 인터랙티브 엑셀 다운로드",
                       data=buf.getvalue(),
                       file_name=f"HeatPump_Analysis_{s_reg}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")