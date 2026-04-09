import streamlit as st
import pandas as pd
import numpy as np
import io, os
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
.info-box   { background:#f8fafc; border:1px solid #e2e8f0; border-radius:12px; padding:28px; margin-bottom:35px; }
.info-title { color:#0f172a; font-size:1.25rem; font-weight:700; margin-bottom:15px; margin-top:0; }
.info-text  { color:#475569; font-size:1.0rem; line-height:1.7; margin-bottom:0; }
.section-title { color:#1e293b; font-weight:700; font-size:1.3rem; margin-top:40px; margin-bottom:16px; border-bottom:2px solid #cbd5e1; padding-bottom:8px; }
.help-text { color: #64748b; font-size: 0.85rem; margin-bottom: 12px; line-height: 1.4; }
</style>
""", unsafe_allow_html=True)

# 지자체 및 태양광 데이터
regions_full = {
    "서울": ["종로구","중구","용산구","성동구","광진구","동대문구","중랑구","성북구","강북구","도봉구","노원구","은평구","서대문구","마포구","양천구","강서구","구로구","금천구","영등포구","동작구","관악구","서초구","강남구","송파구","강동구"],
    "강원도": ["춘천시","원주시","강릉시","동해시","태백시","속초시","삼척시","홍천군","횡성군","영월군","평창군","정선군","철원군","화천군","양구군","인제군","고성군","양양군"],
    "경기도": ["수원시","고양시","용인시","성남시","부천시","화성시","안산시","남양주시","안양시","평택시","시흥시","파주시","의정부시","김포시","광주시","광명시","군포시","하남시","오산시","양주시","이천시","구리시","안성시","포천시","의왕시","여주시","동두천시","과천시","가평군","양평군","연천군"],
    "제주도": ["제주시","서귀포시"],
    "인천": ["중구","동구","미추홀구","연수구","남동구","부평구","계양구","서구","강화군","옹진군"],
    "부산": ["중구","서구","동구","영도구","부산진구","동래구","남구","북구","해운대구","사하구","금정구","강서구","연제구","수영구","사상구","기장군"],
    "대구": ["중구","동구","서구","남구","북구","수성구","달서구","달성군"], "세종": ["세종특별자치시"], "대전": ["동구","중구","서구","유성구","대덕구"], "울산": ["중구","남구","동구","북구","울주군"], "광주": ["동구","서구","남구","북구","광산구"],
    "충청북도":["청주시","충주시","제천시","보은군","옥천군","영동군","증평군","진천군","괴산군","음성군","단양군"],
    "충청남도":["천안시","공주시","보령시","아산시","서산시","논산시","계룡시","당진시","금산군","부여군","서천군","청양군","홍성군","예산군","태안군"],
    "전라북도":["전주시","군산시","익산시","정읍시","남원시","김제시","완주군","진안군","무주군","장수군","임실군","순창군","고창군","부안군"],
    "전라남도":["목포시","여수시","순천시","나주시","광양시","담양군","곡성군","구례군","고흥군","보성군","화순군","장흥군","강진군","해남군","영암군","무안군","함평군","영광군","장성군","완도군","진도군","신안군"],
    "경상북도":["포항시","경주시","김천시","안동시","구미시","영주시","영천시","상주시","문경시","경산시","의성군","청송군","영양군","영덕군","청도군","고령군","성주군","칠곡군","예천군","봉화군","울진군","울릉군"],
    "경상남도":["창원시","진주시","통영시","사천시","김해시","밀양시","거제시","양산시","의령군","함안군","창녕군","고성군","남해군","하동군","산청군","함양군","거창군","합천군"]
}
pv_monthly_data = { "서울": [53.06,79.80,92.62,104.99,98.06,96.96,95.22,72.72,56.62,67.75,65.56,49.98], "인천": [53.06,79.80,92.62,104.99,98.06,96.96,95.22,72.72,56.62,67.75,65.56,49.98], "세종": [53.30,80.02,89.54,109.34,100.67,99.03,102.57,69.17,55.02,69.88,67.62,50.22], "대전": [55.43,81.30,97.35,110.72,99.01,99.49,106.36,72.01,62.35,72.96,70.60,55.90], "대구": [64.90,88.15,90.72,111.86,102.80,102.70,110.62,70.35,51.58,75.80,73.35,60.64], "부산": [68.46,84.08,90.48,105.22,96.64,97.65,107.30,70.11,59.14,74.14,71.75,61.11], "울산": [68.46,84.08,90.48,105.22,96.64,97.65,107.30,70.11,59.14,74.14,71.75,61.11], "광주": [66.32,73.60,90.96,107.97,94.75,97.65,109.43,73.67,68.77,75.09,72.67,62.77], "경기도": [58.95,88.67,102.91,116.65,108.96,107.74,105.80,80.80,62.91,75.27,72.84,55.53], "강원도": [56.59,96.04,94.22,118.18,108.17,108.50,104.22,79.22,57.82,78.17,75.65,54.48], "충청북도": [59.22,88.91,99.49,121.49,111.86,110.03,113.96,76.85,61.13,77.64,75.14,55.80], "충청남도": [61.59,90.33,108.17,123.02,110.01,110.54,118.17,80.01,69.28,81.06,78.45,62.11], "경상북도": [72.11,97.94,100.80,124.29,114.22,114.11,122.91,78.17,57.31,84.22,81.50,67.38], "전라북도": [61.06,83.44,105.54,122.00,108.17,113.34,117.91,81.59,72.08,82.12,79.47,63.43], "경상남도": [76.06,93.42,100.54,116.91,107.38,108.50,119.23,77.90,65.71,82.38,79.72,67.90], "전라남도": [73.69,81.78,101.06,119.96,105.28,108.50,121.59,81.85,76.41,83.43,80.74,69.75], "제주도": [64.22,66.80,87.12,116.14,104.49,92.97,107.64,84.75,78.96,81.85,79.21,66.59] }

# ── 2. 유틸리티 및 엔진 ──
@st.cache_data
def load_simulation_data():
    try:
        def read_csv_safe(f):
            try: return pd.read_csv(f, encoding="utf-8", header=None)
            except: return pd.read_csv(f, encoding="cp949", header=None)
        df_t, df_c = read_csv_safe("외기온도_시간분포.csv"), read_csv_safe("COP_계산기.csv")
        temp_data, cop_data = {}, {}
        zones = ["중부1", "중부2", "남부", "제주"]
        curr = None
        for _, r in df_t.iterrows():
            line = str(r[0]).strip()
            for z in zones:
                if z in line and "▶" in line: curr = z; temp_data[curr] = []; break
            h = line.strip().zfill(2)
            if curr and h.isdigit() and 0 <= int(h) <= 23:
                vals = [float(v) if pd.notna(v) else 0.0 for v in r[1:13]]
                if len(vals) == 12: temp_data[curr].append(vals)
        h_idx = -1
        for i, r in df_c.iterrows():
            if str(r[0]).strip() == "지역": h_idx = i; break
        if h_idx != -1:
            for _, r in df_c.iloc[h_idx+1:].iterrows():
                z = str(r[0]).strip()
                if z in zones:
                    try: cop_data[z] = {'scop': float(r[15])}
                    except: continue
        d = {"중부1": 3.29, "중부2": 3.66, "남부": 3.99, "제주": 4.21}
        for k, v in d.items():
            if k not in cop_data: cop_data[k] = {'scop': v}
        return temp_data, cop_data
    except: return None, None

def map_region_to_zone(s):
    if s == "강원도": return "중부1"
    if s in ["대구","부산","울산","광주","경상남도","전라남도"]: return "남부"
    if s == "제주도": return "제주"
    return "중부2"

def calc_elec_bill(kwh, tariff):
    kwh = max(kwh, 0)
    if tariff == "누진제(가전), 일반용(히트펌프)": b, e = 910, 110 # 단순화 모델
    elif tariff == "주택용 계시별 요금제 (제주)": b, e = 4300, 160
    else: b, e = 910, 120 if kwh<=300 else 214
    fee = (b + (kwh * e) + 14.0 * kwh) * 1.127
    return round(fee / 10000, 4)

def reverse_kwh(bill, tariff):
    lo, hi = 0.0, 3000.0
    for _ in range(40):
        mid = (lo + hi) / 2
        if calc_elec_bill(mid, tariff) < bill: lo = mid
        else: hi = mid
    return round(mid, 1)

def heat_to_hp_kwh(heat_man, cop=3.0):
    # MJ당 요금 68원 기준 역산
    return round(((heat_man * 10000) / 68.0) / 3.6 * 0.92 / cop, 1)

# ── 3. UI 메인 ──
col_t, col_l = st.columns([6,1])
with col_t: st.title("히트펌프 경제성 분석 솔루션")
with col_l:
    if os.path.exists("logo.png"): st.image(Image.open("logo.png"), use_container_width=True)

st.markdown("""
<div class='info-box'>
  <h4 class='info-title'>💡 솔루션 개요</h4>
  <p class='info-text'>
    🏠 <b>시민이 직접 해보는 탄소중립 계산기:</b> 거주 환경과 평소 에너지 사용량만 입력하면, 친환경 히트펌프(AWHP) 전환 시 <b>얼마나 경제적 이득인지</b> 바로 확인하실 수 있습니다.<br><br>
    ⚡ <b>전문적인 데이터 연동:</b> 기상청의 지역별 1년치 외기온도 데이터를 분석하여, 우리 동네 날씨에 맞춘 실제 효율(sCOP)을 반영한 <b>정밀 시뮬레이션 결과</b>를 제공합니다.
  </p>
</div>
""", unsafe_allow_html=True)

df_temp, df_cop = load_simulation_data()

st.markdown('<div class="section-title">1. 대상지 기본 정보</div>', unsafe_allow_html=True)
c1,c2 = st.columns(2)
with c1: s_reg = st.selectbox("광역 지자체", list(regions_full.keys()), index=0)
with c2: s_sub = st.selectbox("기초 지자체", regions_full.get(s_reg, ["전체"]))
c3,c4 = st.columns(2)
with c3: h_type = st.selectbox("주거 형태", ["단독 주택 / 다가구 주택","아파트","연립 / 빌라 / 다세대 주택"])
with c4: h_size = st.number_input("전용 면적 (평)", min_value=10, value=30)

zone = map_region_to_zone(s_reg); dynamic_cop = 3.0
st.markdown('<div class="section-title">📊 우리 동네 기후 및 히트펌프 효율 분석</div>', unsafe_allow_html=True)
if df_temp and df_cop:
    dynamic_cop = df_cop[zone]['scop']
    all_t = []
    days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    for h in range(len(df_temp[zone])):
        for m in range(12): all_t.extend([round(df_temp[zone][h][m])] * days[m])
    counts = pd.Series(all_t).value_counts().sort_index().reset_index()
    counts.columns = ['온도', '시간']
    cl1, cl2 = st.columns([2, 1])
    with cl1:
        c = alt.Chart(counts).mark_bar(color='#3b82f6').encode(x=alt.X('온도:Q', title='외기 온도 (°C)'), y=alt.Y('시간:Q', title='연간 누적 시간 (hours)'), tooltip=['온도', '시간']).properties(height=230)
        st.altair_chart(c, use_container_width=True)
    with cl2:
        st.success(f"**✅ [{s_reg}] 맞춤 효율(sCOP)**\n# {dynamic_cop}"); st.caption("기상청 공식 데이터 기반 연평균 효율")

st.markdown('<div class="section-title">2. 에너지 소비 현황</div>', unsafe_allow_html=True)
cv1,cv2 = st.columns(2)
with cv1: w_heat = st.number_input("동절기(1월) 평균 난방비 (만원)", value=20)
with cv2: w_elec = st.number_input("동절기(1월) 전기요금 (만원)", value=6)

st.markdown('<div class="section-title">3. 시뮬레이션 상수 변수</div>', unsafe_allow_html=True)
cs1, cs2 = st.columns(2)
with cs1:
    f_inf = st.slider("가스/등유요금 인상률 (%)", 0.0, 15.0, 5.0)
    e_inf = st.slider("전기요금 인상률 (%)", 0.0, 15.0, 3.0)
    s_capa = st.number_input("태양광 용량 (kW)", value=3.0)

with cs2:
    # ── 보조금 상하 배치 (요청 반영) ──
    sub_nat = st.checkbox("정부 보조금 적용 (560만원)", value=True)
    is_south = s_reg in ["제주도","경상남도","전라남도"]
    sub_loc = st.checkbox("지자체 매칭 보조금 적용 (280만원)", value=is_south)
    st.caption("*2026년 현재 제주, 경남, 전남은 보조금 신청이 가능합니다.")
    
    st.markdown("---")
    # ── 전기 요금제 선택 (이미지 기반 체크박스형) ──
    st.markdown("**전기 요금제 선택**")
    st.markdown('<div class="help-text">고지서 금액을 바탕으로 사용량(kWh)을 역산하기 위해 적용 중인 요금제를 선택해 주세요.</div>', unsafe_allow_html=True)
    
    if 'tariff' not in st.session_state: st.session_state.tariff = "누진제"
    def set_t1(): st.session_state.tariff = "누진제"
    def set_t2(): st.session_state.tariff = "누진제(가전), 일반용(히트펌프)"
    def set_t3(): st.session_state.tariff = "주택용 계시별 요금제 (제주)"

    st.checkbox("누진제", value=(st.session_state.tariff == "누진제"), on_change=set_t1)
    st.checkbox("누진제(가전), 일반용(히트펌프)", value=(st.session_state.tariff == "누진제(가전), 일반용(히트펌프)"), on_change=set_t2)
    st.checkbox("주택용 계시별 요금제 (제주)", value=(st.session_state.tariff == "주택용 계시별 요금제 (제주)"), on_change=set_t3)
    elec_tariff = st.session_state.tariff

if "analyzed" not in st.session_state: st.session_state.analyzed = False
if st.button("경제성 분석 실행", type="primary", use_container_width=True): st.session_state.analyzed = True

# ── 4. 분석 결과 출력 ──
if st.session_state.analyzed:
    cur_k = reverse_kwh(w_elec, elec_tariff)
    total_sub = (560 if sub_nat else 0) + (280 if sub_loc else 0)
    net_cap = max(0, 600 + h_size * 10 - total_sub)
    ann_heat_base = w_heat * 4.8; ann_elec_base = calc_elec_bill(cur_k, elec_tariff) * 12
    ann_hp_net_op = (ann_heat_base / dynamic_cop) * 0.85

    years, gas_cum, hp_cum, net_p = list(range(1,16)), [], [], []
    g_s, h_s, pb = 0.0, float(net_cap), "15년 초과"
    for y in years:
        cg = ann_heat_base * ((1+f_inf/100)**y) + ann_elec_base
        ch = ann_hp_net_op * ((1+e_inf/100)**y) + ann_elec_base
        g_s += cg; h_s += ch; p = int(g_s - h_s)
        gas_cum.append(int(g_s)); hp_cum.append(int(h_s)); net_p.append(p)
        if pb == "15년 초과" and p > 0: pb = f"{y}년차"

    st.markdown('<div class="section-title">분석 결과 요약</div>', unsafe_allow_html=True)
    ca, cb, cc = st.columns(3)
    ca.metric("투자 회수 시점", pb); cb.metric("15년 순이익", f"{net_p[-1]:,} 만원"); cc.metric("적용 sCOP", f"{dynamic_cop}")

    g1, g2 = st.columns(2)
    with g1:
        st.write("**15년 누적 비용 흐름**")
        df_a = pd.DataFrame({"연도":years, "기존":gas_cum, "HP":hp_cum}).melt("연도", var_name="시나리오", value_name="비용")
        st.altair_chart(alt.Chart(df_a).mark_area(opacity=0.5).encode(x="연도:O", y="비용:Q", color="시나리오:N"), use_container_width=True)
    with g2:
        st.write("**연도별 순수익(Cash Flow)**")
        df_c = pd.DataFrame({"연도":years, "순수익":net_p, "상태":["수익" if p>0 else "회수" for p in net_p]})
        st.altair_chart(alt.Chart(df_c).mark_bar().encode(x="연도:O", y="순수익:Q", color="상태:N"), use_container_width=True)

    # ════ 전문가용 3탭 인터랙티브 엑셀 ════
    wb = Workbook()
    hf, sf, pf = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid"), PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid"), PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
    fw, fb, fi = Font(color="FFFFFF", bold=True), Font(bold=True), Font(color="0000FF", bold=True)
    thin = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    center, right = Alignment(horizontal="center"), Alignment(horizontal="right")

    ws1 = wb.active; ws1.title = "①입력_가정"
    ws1.merge_cells("A1:D1"); ws1["A1"] = f"히트펌프 경제성 분석 마스터 ({s_reg})"; ws1["A1"].fill = hf; ws1["A1"].font = fw; ws1["A1"].alignment = center
    rows1 = [("항목", "값", "단위", "비고"), ("1월 난방비", w_heat, "만원", "입력"), ("1월 전기요금", w_elec, "만원", "입력"), ("지역 sCOP", dynamic_cop, "-", "기상연동"), ("설비 CAPEX", 600+h_size*10, "만원", ""), ("정부보조금", 560, "만원", ""), ("지방보조금", 280, "만원", ""), ("순 투자비", net_cap, "만원", "=B7-B8-B9")]
    for ri, rdata in enumerate(rows1, 3):
        for ci, val in enumerate(rdata, 1):
            c = ws1.cell(row=ri, column=ci, value=val); c.border=thin
            if ri==3: c.fill=sf; c.font=fb
            elif ci==2 and ri!=10: c.font=fi; c.alignment=right
    ws1["B10"]="=B7-B8-B9"; ws1.column_dimensions["A"].width=20

    ws2 = wb.create_sheet("②월별_운영_상세")
    ws2.merge_cells("A1:G1"); ws2["A1"] = "월별 에너지 비용 및 절감 시뮬레이션"; ws2["A1"].fill = hf; ws2["A1"].font = fw; ws2["A1"].alignment = center
    for ci, h in enumerate(["월", "기존 난방비", "HP전력(kWh)", "HP 운영비(만)", "PV 절감(만)", "Net 운영비", "월 절감액"], 1):
        c = ws2.cell(row=2, column=ci, value=h); c.fill=sf; c.font=fb; c.border=thin; c.alignment=center
    for m in range(1, 13):
        r = m+2; ws2.cell(row=r, column=1, value=f"{m}월").border=thin
        ws2.cell(row=r, column=2, value=w_heat if m in [1,2,12] else w_heat*0.2).border=thin
        ws2.cell(row=r, column=4, value=f"=C{r}*0.019").border=thin
        ws2.cell(row=r, column=6, value=f"=D{r}-E{r}").border=thin
        ws2.cell(row=r, column=7, value=f"=B{r}-F{r}").border=thin
        if m % 2 == 0:
            for ci in range(1, 8): ws2.cell(row=r, column=ci).fill = pf
    for col in "ABCDEFG": ws2.column_dimensions[col].width=15

    ws3 = wb.create_sheet("③15년_재무_분석")
    ws3.merge_cells("A1:H1"); ws3["A1"] = "15년 장기 투자 회수 및 NPV 시뮬레이션"; ws3["A1"].fill = hf; ws3["A1"].font = fw; ws3["A1"].alignment = center
    for ci, h in enumerate(["연도", "물가지수", "기존 OPEX", "HP OPEX", "연간 순이익", "누적 NPV", "ROI", "상태"], 1):
        c = ws3.cell(row=2, column=ci, value=h); c.fill=sf; c.font=fb; c.border=thin; c.alignment=center
    for y in range(1, 16):
        r = y+2; ws3.cell(row=r, column=1, value=f"{y}년차").border=thin
        ws3.cell(row=r, column=2, value=f"=(1+0.04)^{y-1}").border=thin
        ws3.cell(row=r, column=3, value=ann_heat_base).border=thin
        ws3.cell(row=r, column=4, value=round(ann_hp_net_op, 2)).border=thin
        ws3.cell(row=r, column=5, value=f"=C{r}-D{r}").border=thin
        if y == 1: ws3.cell(row=r, column=6, value=f"=E{r}-'①입력_가정'!$B$10").border=thin
        else: ws3.cell(row=r, column=6, value=f"=F{r-1}+E{r}").border=thin
        ws3.cell(row=r, column=7, value=f"=F{r}/'①입력_가정'!$B$10").border=thin; ws3.cell(row=r, column=7).number_format='0%'
        ws3.cell(row=r, column=8, value=f"=IF(F{r}>0, \"수익\", \"회수중\")").border=thin
        if y % 2 == 0:
            for ci in range(1, 9): ws3.cell(row=r, column=ci).fill = pf
    for col in "ABCDEFGH": ws3.column_dimensions[col].width=15

    buf = io.BytesIO(); wb.save(buf)
    st.markdown("---")
    st.download_button(label="🚀 전문가용 수식 연동 정밀 엑셀 다운로드", data=buf.getvalue(), file_name=f"Expert_Report_{s_reg}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)