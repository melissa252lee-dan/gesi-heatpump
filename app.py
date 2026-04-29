"""
히트펌프 경제성 분석 솔루션
─────────────────────────────────────────────────────────────
시민이 거주 환경과 에너지 사용량을 입력하면, 친환경 히트펌프(AWHP)
전환 시 경제적 이득과 환경 기여도를 계산해주는 Streamlit 앱입니다.

데이터 출처: 전기요금완료본.xlsx (요금제 × 태양광 × 난방유형 20개 블록)
"""
import io
import os
import pandas as pd
import streamlit as st
import altair as alt
from PIL import Image
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


st.set_page_config(page_title="히트펌프 경제성 분석 솔루션", layout="wide")


# ══════════════════════════════════════════════════════════════════════
# 1. 스타일 정의
# ══════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
* { font-family: 'Pretendard', sans-serif; }

/* 박스·타이틀 */
.info-box      { background:#f8fafc; border:1px solid #e2e8f0; border-radius:12px; padding:28px; margin-bottom:35px; }
.info-title    { color:#0f172a; font-size:1.25rem; font-weight:700; margin-bottom:15px; margin-top:0; }
.info-text     { color:#475569; font-size:1.0rem; line-height:1.7; margin-bottom:0; }
.section-title { color:#1e293b; font-weight:700; font-size:1.3rem; margin-top:40px; margin-bottom:16px; border-bottom:2px solid #cbd5e1; padding-bottom:8px; }
.help-text     { color:#64748b; font-size:0.85rem; margin-bottom:12px; line-height:1.4; }

/* 환경 기여 박스 */
.saving-box    { background:#f0fdf4; border:2px solid #86efac; border-radius:12px; padding:20px 24px; margin:16px 0; }
.saving-title  { color:#15803d; font-size:1.1rem; font-weight:700; margin-bottom:4px; }
.saving-sub    { color:#166534; font-size:0.95rem; line-height:1.65; }

/* 배지 (요금제·난방·규모 표시) */
.tariff-badge  { display:inline-block; background:#dbeafe; color:#1e40af; padding:4px 10px; border-radius:6px; font-size:0.85rem; font-weight:600; margin-right:6px; }
.solar-badge-x { display:inline-block; background:#fef3c7; color:#92400e; padding:4px 10px; border-radius:6px; font-size:0.85rem; font-weight:600; margin-right:6px; }
.solar-badge-o { display:inline-block; background:#dcfce7; color:#15803d; padding:4px 10px; border-radius:6px; font-size:0.85rem; font-weight:600; margin-right:6px; }

/* 호버 툴팁 — 배지 위에 마우스를 올리면 상세 설명 표시 */
.has-tooltip { position:relative; cursor:help; }
.has-tooltip::after {
    content: attr(data-tooltip);
    position: absolute;
    bottom: calc(100% + 10px); left: 50%;
    transform: translateX(-50%);
    background: #1e293b; color: #ffffff;
    padding: 10px 14px; border-radius: 8px;
    font-size: 0.82rem; font-weight: 400;
    width: 300px; white-space: normal;
    line-height: 1.55; text-align: left;
    opacity: 0; visibility: hidden;
    pointer-events: none;
    transition: opacity 0.2s, visibility 0.2s;
    z-index: 1000;
    box-shadow: 0 4px 14px rgba(0,0,0,0.18);
}
.has-tooltip::before {
    content: '';
    position: absolute;
    bottom: calc(100% + 4px); left: 50%;
    transform: translateX(-50%);
    border: 6px solid transparent;
    border-top-color: #1e293b;
    opacity: 0; visibility: hidden;
    pointer-events: none;
    transition: opacity 0.2s, visibility 0.2s;
    z-index: 1000;
}
.has-tooltip:hover::after,
.has-tooltip:hover::before { opacity: 1; visibility: visible; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# 2. 정적 데이터 및 상수
# ══════════════════════════════════════════════════════════════════════

# ── 광역·기초 지자체 목록 ──
REGIONS_FULL = {
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

# ── HDD(난방도일) 및 기후존별 sCOP는 엑셀의 'COP_계산기' 시트에서 동적 로드됩니다. ──
# Tbase=18°C 기준, 각 지역 월별 데이터는 load_tariff_xlsx()를 통해 가져옵니다.

# ── 엑셀 블록 헤더 위치 ──
# (요금제, 태양광플래그, 난방유형) → 헤더 행 번호
# 일반용은 별도 미터(HP 전용)이라 태양광 영향 없음 → 태X만 존재
EXCEL_BLOCK_HEADERS = {
    ("누진제", "태X", "도시가스(콘덴싱)"):   5, ("누진제", "태X", "도시가스(일반)"):  18,
    ("누진제", "태X", "등유"):              31, ("누진제", "태X", "LPG"):            44,
    ("누진제", "태O", "도시가스(콘덴싱)"):  57, ("누진제", "태O", "도시가스(일반)"):  70,
    ("누진제", "태O", "등유"):              83, ("누진제", "태O", "LPG"):            96,
    ("일반용", "태X", "도시가스(콘덴싱)"): 109, ("일반용", "태X", "도시가스(일반)"): 122,
    ("일반용", "태X", "등유"):             135, ("일반용", "태X", "LPG"):           148,
    ("계시별", "태X", "도시가스(콘덴싱)"): 161, ("계시별", "태X", "도시가스(일반)"): 174,
    ("계시별", "태X", "등유"):             187, ("계시별", "태X", "LPG"):           200,
    ("계시별", "태O", "도시가스(콘덴싱)"): 213, ("계시별", "태O", "도시가스(일반)"): 226,
    ("계시별", "태O", "등유"):             239, ("계시별", "태O", "LPG"):           252,
}

# ── UI 옵션 → 엑셀 키 매핑 ──
HEATING_TYPE_MAP = {
    "가스 콘덴싱 보일러": "도시가스(콘덴싱)",
    "가스 일반 보일러":   "도시가스(일반)",
    "등유 보일러":        "등유",
    "LPG 보일러":         "LPG",
}

# ── 5개 요금제 탭 → (엑셀 요금제, 엑셀 태양광 플래그) ──
# 사용자에게는 "태양광 설치/미설치"로 표기, 엑셀 내부 키는 태O/태X 그대로 유지
TARIFF_LABEL_MAP = {
    "누진제 (태양광 미설치)": ("누진제", "태X"),
    "누진제 (태양광 설치)":   ("누진제", "태O"),
    "일반용 (HP 전용 미터)":  ("일반용", "태X"),
    "계시별 (태양광 미설치)": ("계시별", "태X"),
    "계시별 (태양광 설치)":   ("계시별", "태O"),
}

# ── CO₂ 배출 환산 계수 (1만원당 kgCO₂) ──
# 출처: 환경부 온실가스 종합정보센터(GIR) 배출계수 + 2024년 평균 단가
# - 도시가스: 8.0m³/만원 × 2.243 kgCO₂/m³ ≈ 18.0
# - 등유:     6.7L/만원 × 2.690 kgCO₂/L ≈ 18.0
# - LPG:      5.0kg/만원 × 3.000 kgCO₂/kg = 15.0
# - 전기:     약 40kWh/만원 × 0.4242 kgCO₂/kWh ≈ 17.0 (한전 2023 그리드 평균)
CO2_PER_MAN_FUEL = {
    "도시가스(콘덴싱)": 18.0, "도시가스(일반)": 18.0,
    "등유":            18.0, "LPG":           15.0,
}
CO2_PER_MAN_ELEC = 17.0

# ── 환경 비유 환산 ──
# 30년생 소나무 1그루: 연 약 6.6 kgCO₂ 흡수 (산림청)
# 승용차 1km 운행: 약 0.15 kgCO₂ 배출 (환경부 평균)
TREE_KG_PER_YEAR = 6.6
CAR_KG_PER_KM    = 0.15

# ── Sheet2 (kWh 기반 물리 데이터) 행 매핑 ──
# 표준 가구(중부2, 거창군 32평, 연 65만원) 기준의 월별 에너지 흐름.
# 행 45-48: 월별 난방 에너지 수요(kWh) — 가스/등유/LPG 기존 보일러 사용 시
# 행 53-56: 월별 HP 전력 사용량(kWh) — HP로 전환했을 때 (= 난방수요 ÷ 월별 COP)
SHEET2_DEMAND_ROWS = {
    "도시가스(콘덴싱)": 45, "도시가스(일반)": 46, "등유": 47, "LPG": 48,
}
SHEET2_HP_KWH_ROWS = {
    "도시가스(콘덴싱)": 53, "도시가스(일반)": 54, "등유": 55, "LPG": 56,
}

# ── Sheet3 광역 시도별 월별 난방 비중 매핑 ──
# 사이트의 약식 시도명 → Sheet3의 정식 명칭
REGION_NAME_MAP = {
    "서울": "서울특별시", "부산": "부산광역시", "대구": "대구광역시",
    "인천": "인천광역시", "광주": "광주광역시", "대전": "대전광역시",
    "울산": "울산광역시", "세종": "세종특별자치시",
    "경기도":   "경기도",       "강원도":   "강원특별자치도",
    "충청북도": "충청북도",     "충청남도": "충청남도",
    "전라북도": "전북특별자치도", "전라남도": "전라남도",
    "경상북도": "경상북도",     "경상남도": "경상남도",
    "제주도":   "제주특별자치도",
}

# ── 기초지자체별 zone 오버라이드 (Sheet2 B49 수식 로직) ──
# 광역 시도만으로는 정확하지 않은 지역 — 기초지자체까지 봐야 함
COASTAL_GANGWON   = {"고성군", "속초시", "양양군", "강릉시", "동해시", "삼척시"}  # 강원 해안 → 중부2
NORTHERN_GYEONGGI = {"연천군", "포천시", "가평군", "남양주시", "의정부시", "양주시", "동두천시", "파주시"}  # 경기 북부 → 중부1
SOUTHERN_GYEONGBUK = {"울진군", "영덕군", "포항시", "경주시", "청도군", "경산시"}  # 경북 남쪽 → 남부
NORTHERN_GYEONGBUK = {"봉화군", "청송군"}  # 경북 북쪽 → 중부1
INNER_GYEONGNAM = {"거창군", "함양군"}  # 경남 내륙 → 중부2 (다른 경남은 남부)

# ── 엑셀 표준 가구 연 난방비(원) — 모든 블록 동일 ──
EXCEL_BASE_ANNUAL_WON = 650516

# ── 보조금 (만원) ──
SUBSIDY_NATIONAL = 560
SUBSIDY_LOCAL    = 280
SOUTHERN_REGIONS = {"제주도", "경상남도", "전라남도"}


# ══════════════════════════════════════════════════════════════════════
# 3. 데이터 로더
# ══════════════════════════════════════════════════════════════════════

@st.cache_data
def load_tariff_xlsx():
    """전기요금완료본.xlsx에서 모든 기준 데이터를 로드.

    [전기요금 시트 — 20개 블록]
    각 블록 = (요금제, 태양광유무, 난방유형) 조합.
    블록 헤더 행: col 19=HP 연합계(원), col 20=기존난방비 연합계(원), col 21=Saving 비율
    헤더+1 ~ 헤더+12 행: 1~12월별 청구 내역 (col 9~18)

    [COP_계산기 시트 — 4개 기후존]
    각 존 3행씩: 월평균기온 / 월 HDD(Tbase=18°C) / 월 COP(난방시간 가중)
    sCOP(HDD가중)는 각 존의 첫 행 col 16

    [Sheet2 — 표준 가구의 물리적 에너지 흐름 (kWh)]
    행 45-48: 월별 난방 에너지 수요(연료별, 보일러 사용 시)
    행 53-56: 월별 HP 전력 사용량(연료→HP 전환 후, = 난방수요 ÷ COP)

    Returns:
        (data, error_msg) — 성공 시 (dict, None), 실패 시 (None, str)
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

        # ── 전기요금 시트 — 20개 블록 ──
        ws = wb["전기요금"]
        blocks = {}
        for key, hr in EXCEL_BLOCK_HEADERS.items():
            blocks[key] = {
                "hp_annual_won":       float(ws.cell(row=hr, column=19).value or 0),
                "existing_annual_won": float(ws.cell(row=hr, column=20).value or 0),
                "saving_ratio":        float(ws.cell(row=hr, column=21).value or 0),
                "monthly": [
                    {
                        "청구요금합계": ws.cell(row=hr+1+m, column=17).value or 0,
                        "HP전기요금":   ws.cell(row=hr+1+m, column=18).value or 0,
                    }
                    for m in range(12)
                ],
            }

        # ── COP_계산기 시트 — 기후존별 sCOP / HDD / 월별 COP / 월평균기온 ──
        ws_cop = wb["COP_계산기"]
        zone_rows = {"중부1": 12, "중부2": 16, "남부": 20, "제주": 24}
        scop, hdd, monthly_cop, monthly_temp = {}, {}, {}, {}
        for zone, r_temp in zone_rows.items():
            r_hdd, r_cop = r_temp + 1, r_temp + 2
            scop[zone] = float(ws_cop.cell(row=r_temp, column=16).value or 0)
            monthly_temp[zone] = [float(ws_cop.cell(row=r_temp, column=3+m).value or 0) for m in range(12)]
            hdd[zone]          = [float(ws_cop.cell(row=r_hdd,  column=3+m).value or 0) for m in range(12)]
            monthly_cop[zone]  = [float(ws_cop.cell(row=r_cop,  column=3+m).value or 0) for m in range(12)]

        # ── Sheet2 — 연료별 월별 난방 에너지 수요 / HP 전력 사용량 (kWh) ──
        ws_s2 = wb["Sheet2"]
        kwh_demand = {fuel: [float(ws_s2.cell(row=r, column=3+m).value or 0) for m in range(12)]
                      for fuel, r in SHEET2_DEMAND_ROWS.items()}
        kwh_hp     = {fuel: [float(ws_s2.cell(row=r, column=3+m).value or 0) for m in range(12)]
                      for fuel, r in SHEET2_HP_KWH_ROWS.items()}

        # ── Sheet2 핵심 파라미터 (행 43 수식 복제용) ──
        # 행 9 col 3-6: 연료별 효율  /  행 40: 기본요금  /  행 42: 단가(원/kWh)  /  행 14: 난방비중
        fuel_order = ["도시가스(콘덴싱)", "도시가스(일반)", "등유", "LPG"]
        sheet2_params = {
            "efficiency":    {f: float(ws_s2.cell(row=9,  column=3+i).value or 0) for i, f in enumerate(fuel_order)},
            "base_fee":      {f: float(ws_s2.cell(row=40, column=3+i).value or 0) for i, f in enumerate(fuel_order)},
            "rate":          {f: float(ws_s2.cell(row=42, column=3+i).value or 0) for i, f in enumerate(fuel_order)},
            "heating_share": float(ws_s2.cell(row=14, column=3).value or 0.85),
        }

        # ── Sheet3 — 광역시도별 월별 난방 비중 (17개 시도 + 전국) ──
        ws_s3 = wb["Sheet3"]
        region_ratios = {}
        for r in range(2, 20):
            name = ws_s3.cell(row=r, column=1).value
            if not name: continue
            ratios = [float(ws_s3.cell(row=r, column=3+m).value or 0) for m in range(12)]
            region_ratios[name] = ratios

        return {
            "blocks":         blocks,
            "scop":           scop,
            "hdd":            hdd,
            "monthly_cop":    monthly_cop,
            "monthly_temp":   monthly_temp,
            "kwh_demand":     kwh_demand,
            "kwh_hp":         kwh_hp,
            "sheet2_params":  sheet2_params,
            "region_ratios":  region_ratios,
        }, None
    except Exception as e:
        return None, str(e)


# ══════════════════════════════════════════════════════════════════════
# 4. 계산 함수
# ══════════════════════════════════════════════════════════════════════

def map_region_to_zone(region, sub_region=""):
    """광역+기초지자체 조합으로 기후존 결정 (엑셀 Sheet2 B49 수식 로직 복제).

    더 정밀한 분류: 강원 해안(따뜻함) / 경기 북부(추움) / 경북 남쪽(따뜻함) 등 반영.
    """
    # 제주
    if region == "제주도":
        return "제주"
    # 남부 (해안 따뜻한 광역)
    if region in {"부산", "대구", "울산", "광주", "전라남도"}:
        return "남부"
    # 남부 (경북 남쪽)
    if region == "경상북도" and sub_region in SOUTHERN_GYEONGBUK:
        return "남부"
    # 남부 (경남, 단 거창·함양 제외)
    if region == "경상남도" and sub_region not in INNER_GYEONGNAM:
        return "남부"
    # 중부1 (강원 산악)
    if region == "강원도" and sub_region not in COASTAL_GANGWON:
        return "중부1"
    # 중부1 (경기 북부)
    if region == "경기도" and sub_region in NORTHERN_GYEONGGI:
        return "중부1"
    # 중부1 (충북 제천)
    if region == "충청북도" and sub_region == "제천시":
        return "중부1"
    # 중부1 (경북 북쪽 봉화·청송)
    if region == "경상북도" and sub_region in NORTHERN_GYEONGBUK:
        return "중부1"
    # 나머지 = 중부2 (서울/대전/세종/인천/충남/전북/강원해안/경기 외/충북 외/경남 거창함양/경북 외)
    return "중부2"


def get_block_key(tariff_label, heating_ui):
    """UI 라벨로부터 엑셀 블록 키 결정.

    Returns: (block_key, tariff, solar_flag)
    """
    tariff, solar = TARIFF_LABEL_MAP[tariff_label]
    heating = HEATING_TYPE_MAP[heating_ui]
    return (tariff, solar, heating), tariff, solar


def calc_standard_jan_heat_man(monthly_distribution):
    """엑셀 표준 가구의 1월 난방비(만원) 추정.

    표준 연간 난방비(650,516원)를 월별 사용 비중으로 안분하여 1월 비중만 추출.

    Args:
        monthly_distribution: 12개월 사용량 분포 (kWh 또는 비중) — 1월 시작
    """
    if sum(monthly_distribution) == 0: return 0
    return EXCEL_BASE_ANNUAL_WON * monthly_distribution[0] / sum(monthly_distribution) / 10000


def apply_block_with_scale(block, scale):
    """엑셀 블록 데이터에 사용자 가구 규모(scale) 적용.

    월별 비용은 'HP전기요금'(col 18, 가전 분리 후 순수 HP만) 기준 — col 19 연합계와 일관.
    """
    monthly_won = [round(m["HP전기요금"] * scale) for m in block["monthly"]]
    monthly_man = [round(v / 10000, 2) for v in monthly_won]
    hp_ann_won  = block["hp_annual_won"]       * scale
    ex_ann_won  = block["existing_annual_won"] * scale
    return {
        "monthly_won":   monthly_won,
        "monthly_man":   monthly_man,
        "hp_annual_man": round(hp_ann_won / 10000, 1),
        "hp_annual_won": round(hp_ann_won),
        "ex_annual_man": round(ex_ann_won / 10000, 1),
        "ex_annual_won": round(ex_ann_won),
        "saving_man":    round((ex_ann_won - hp_ann_won) / 10000, 1),
        "saving_ratio":  block["saving_ratio"],
    }


def get_hp_specs(h_size_pyung):
    """전용면적(평) → (설치 공간 비유, 치수, HP 용량)."""
    if h_size_pyung < 20:
        return ("소형 냉장고 크기",   "595 × 625 mm",   "6 kW")
    if h_size_pyung <= 28:
        return ("워시타워 1대 크기", "800 × 1,115 mm", "10 kW")
    if h_size_pyung <= 35:
        return ("워시타워 1대 크기", "800 × 1,115 mm", "12 kW")
    return     ("보일러실 크기",     "1,120 × 1,666 mm", "16 kW")


def calc_monthly_stats(monthly_ex_man, monthly_hp_man, fuel_key):
    """월별 절감액·누적·절감률·CO₂ 한 번에 계산."""
    co2_factor_fuel = CO2_PER_MAN_FUEL[fuel_key]
    savings, cumulative, savings_pct, co2 = [], [], [], []
    cum = 0.0
    for ex, hp in zip(monthly_ex_man, monthly_hp_man):
        sav = round(ex - hp, 2)
        cum += sav
        savings.append(sav)
        cumulative.append(round(cum, 2))
        # 비난방월(기존 난방비 0)은 비율 의미 없음 → "-"
        savings_pct.append(f"{round(sav/ex*100, 1)}%" if ex > 0 else "-")
        co2.append(round(ex * co2_factor_fuel - hp * CO2_PER_MAN_ELEC, 1))
    return {
        "savings":    savings,
        "cumulative": cumulative,
        "pct":        savings_pct,
        "co2":        co2,
    }


def calc_annual_co2_savings(ex_annual_man, hp_annual_man, fuel_key):
    """연간 CO₂ 절감량(kg) + 시민 친화 비유 환산.

    Returns: (co2_kg, trees, car_km)
    """
    co2_ex = ex_annual_man * CO2_PER_MAN_FUEL[fuel_key]
    co2_hp = hp_annual_man * CO2_PER_MAN_ELEC
    co2_saving = max(0, co2_ex - co2_hp)
    trees  = round(co2_saving / TREE_KG_PER_YEAR)
    car_km = round(co2_saving / CAR_KG_PER_KM)
    return co2_saving, trees, car_km


def calc_demand_kwh_sheet2(annual_cost_won, fuel_key, sheet2_params):
    """Sheet2 행 43 수식 그대로 복제 — 연간 난방 에너지 수요 (kWh) 계산.

    수식 (도시가스):  ((연간요금 - 기본요금×12) ÷ 단가) × 난방비중 × 효율
    수식 (등유/LPG):  (연간요금 ÷ 단가) × 효율  (기본요금 0, 난방비중 미적용)
    """
    base_fee = sheet2_params["base_fee"][fuel_key]
    rate     = sheet2_params["rate"][fuel_key]
    eff      = sheet2_params["efficiency"][fuel_key]
    share    = sheet2_params["heating_share"]
    if rate <= 0: return 0
    fuel_input_kwh = (annual_cost_won - base_fee * 12) / rate
    if "도시가스" in fuel_key:                # 도시가스만 난방비중 적용
        fuel_input_kwh *= share
    return max(0, fuel_input_kwh * eff)


def calc_kwh_data(annual_demand_kwh, monthly_ratios, monthly_cop_zone):
    """Sheet2 행 45-48, 53-56 수식 복제 — 월별 kWh 계산.

    Args:
        annual_demand_kwh: 연간 난방 에너지 수요 (calc_demand_kwh_sheet2 결과)
        monthly_ratios:    사용자 광역시도의 월별 난방 비중 (Sheet3, 합=1)
        monthly_cop_zone:  사용자 기후존의 월별 COP

    Returns: dict with 월별/연간 kWh 및 효율
    """
    monthly_demand = [round(annual_demand_kwh * r, 1) for r in monthly_ratios]
    monthly_hp = [round(d / cop, 1) if cop > 0 else 0.0
                  for d, cop in zip(monthly_demand, monthly_cop_zone)]
    annual_demand = round(sum(monthly_demand), 0)
    annual_hp     = round(sum(monthly_hp), 0)
    efficiency    = round(annual_demand / annual_hp, 1) if annual_hp > 0 else 0.0
    return {
        "monthly_demand": monthly_demand,
        "monthly_hp":     monthly_hp,
        "annual_demand":  annual_demand,
        "annual_hp":      annual_hp,
        "efficiency":     efficiency,
    }


def simulate_18yr(net_capex_man, ann_heat_man, ann_hp_man, fuel_inflation_pct, elec_inflation_pct):
    """18년 누적 비용·순이익 시뮬레이션 (인플레이션 복리 적용).

    Returns: (years, gas_cum, hp_cum, net_profit, payback_year)
    """
    years = list(range(1, 19))
    gas_cum, hp_cum, net_profit = [], [], []
    gas_total, hp_total = 0.0, float(net_capex_man)
    payback = "18년 초과"

    for y in years:
        gas_total += ann_heat_man * ((1 + fuel_inflation_pct / 100) ** y)
        hp_total  += ann_hp_man   * ((1 + elec_inflation_pct / 100) ** y)
        profit = int(gas_total - hp_total)
        gas_cum.append(int(gas_total))
        hp_cum.append(int(hp_total))
        net_profit.append(profit)
        if payback == "18년 초과" and profit > 0:
            payback = f"{y}년차"

    return years, gas_cum, hp_cum, net_profit, payback


def safe_filename(label):
    """라벨에서 공백·괄호 제거 (Windows 파일명 호환)."""
    return label.replace(" ", "").replace("(", "_").replace(")", "")


# ══════════════════════════════════════════════════════════════════════
# 5. 엑셀 리포트 빌더
# ══════════════════════════════════════════════════════════════════════

def build_excel_report(*, region, tariff_label, fuel_key,
                       winter_heat_man, winter_elec_man, solar_capa_kw,
                       scale, csv_jan_man, dynamic_cop, zone,
                       capex_man, use_subsidy_nat, use_subsidy_loc, net_capex_man,
                       ann_heat_base, ann_hp_op, result,
                       monthly_ex_man, hdd_zone):
    """3개 시트로 구성된 엑셀 리포트 생성 후 BytesIO 반환."""
    wb = Workbook()

    # 공통 스타일
    fill_header  = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fill_subhead = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_blue    = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
    fill_green   = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")
    font_white   = Font(color="FFFFFF", bold=True)
    font_bold    = Font(bold=True)
    font_input   = Font(color="0000FF", bold=True)
    font_saving  = Font(color="166534", bold=True)
    border_thin  = Border(left=Side(style="thin"), right=Side(style="thin"),
                          top=Side(style="thin"),  bottom=Side(style="thin"))
    align_center = Alignment(horizontal="center")
    align_right  = Alignment(horizontal="right")

    # ── 시트 ① 입력·가정 ─────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "①입력_가정"
    ws1.merge_cells("A1:D1")
    ws1["A1"] = f"히트펌프 경제성 분석 ({region} / {tariff_label})"
    ws1["A1"].fill = fill_header; ws1["A1"].font = font_white; ws1["A1"].alignment = align_center

    rows1 = [
        ("항목",            "값",                                        "단위",  "산출 근거"),
        ("1월 난방비",       winter_heat_man,                            "만원",  "사용자 입력"),
        ("1월 전기요금",     winter_elec_man,                            "만원",  "사용자 입력"),
        ("난방 방식",        fuel_key,                                   "-",     "사용자 선택"),
        ("전기 요금제",      tariff_label,                               "-",     "사용자 선택"),
        ("태양광 용량",      solar_capa_kw,                              "kW",    "사용자 입력 (참고용)"),
        ("적용 데이터",      f"{tariff_label} / {fuel_key}",             "-",     "전기요금완료본.xlsx"),
        ("규모 보정 계수",   round(scale, 2),                            "배",    f"={winter_heat_man}÷{round(csv_jan_man, 2)}"),
        ("지역 sCOP",        dynamic_cop,                                "-",     f"기후 존 ({zone})"),
        ("설비 CAPEX",       capex_man,                                  "만원",  "국내 기업 자료"),
        ("정부 보조금",      SUBSIDY_NATIONAL if use_subsidy_nat else 0, "만원",  "기후에너지환경부 2026"),
        ("지방 보조금",      SUBSIDY_LOCAL    if use_subsidy_loc else 0, "만원",  "정부 50% 매칭"),
        ("순 투자비",        net_capex_man,                              "만원",  "=CAPEX-보조금"),
        ("기존 연간 난방비", ann_heat_base,                              "만원",  "엑셀 기존난방비×규모보정"),
        ("HP 연간 전기요금", ann_hp_op,                                  "만원",  "엑셀 HP연합계×규모보정"),
        ("연간 절감액",      result["saving_man"],                       "만원",  "=기존-HP"),
        ("Saving 비율",      f"{round(result['saving_ratio']*100, 1)}%", "-",     "엑셀 원본"),
    ]
    for ri, row_data in enumerate(rows1, 3):
        for ci, val in enumerate(row_data, 1):
            cell = ws1.cell(row=ri, column=ci, value=val)
            cell.border = border_thin
            if ri == 3:
                cell.fill = fill_subhead; cell.font = font_bold
            elif ci == 2:
                cell.font = font_input; cell.alignment = align_right
    ws1.column_dimensions["A"].width = 22
    ws1.column_dimensions["D"].width = 45

    # ── 시트 ② 월별 청구요금 ─────────────────────────────────────
    ws2 = wb.create_sheet("②월별_청구요금")
    ws2.merge_cells("A1:H1")
    ws2["A1"] = f"월별 청구요금·CO₂ [{tariff_label} / {fuel_key}]"
    ws2["A1"].fill = fill_header; ws2["A1"].font = font_white; ws2["A1"].alignment = align_center

    headers2 = ["월", "기존 난방비(만원)", "HP 난방요금(만원)", "월별 절감액(만원)",
                "절감률", "누적 절감액(만원)", "CO₂ 절감(kg)", "비고"]
    for ci, h in enumerate(headers2, 1):
        cell = ws2.cell(row=2, column=ci, value=h)
        cell.fill = fill_subhead; cell.font = font_bold
        cell.border = border_thin; cell.alignment = align_center

    co2_factor_fuel = CO2_PER_MAN_FUEL[fuel_key]
    cum = 0.0
    for m in range(1, 13):
        r = m + 2
        ex      = monthly_ex_man[m-1]
        hp      = result["monthly_man"][m-1]
        sav     = round(ex - hp, 2)
        cum    += sav
        sav_pct = f"{round(sav/ex*100, 1)}%" if ex > 0 else "-"
        co2     = round(ex * co2_factor_fuel - hp * CO2_PER_MAN_ELEC, 1)
        note    = "난방월" if hdd_zone[m-1] > 0 else "비난방월"

        for ci, val in enumerate([f"{m}월", ex, hp, sav, sav_pct, round(cum, 2), co2, note], 1):
            ws2.cell(row=r, column=ci, value=val).border = border_thin
        # 절감 관련 셀 초록색 강조
        for ci in (4, 6, 7):
            ws2.cell(row=r, column=ci).font = font_saving
        if m % 2 == 0:
            for ci in range(1, 9):
                ws2.cell(row=r, column=ci).fill = fill_green

    # 합계 행
    r_sum = 15
    co2_total = round(result["ex_annual_man"] * co2_factor_fuel
                      - result["hp_annual_man"] * CO2_PER_MAN_ELEC, 1)
    summary_vals = [
        "연간 합계", ann_heat_base, ann_hp_op, result["saving_man"],
        f"{round(result['saving_ratio']*100, 1)}%", round(cum, 2), co2_total, "-"
    ]
    for ci, val in enumerate(summary_vals, 1):
        cell = ws2.cell(row=r_sum, column=ci, value=val)
        cell.border = border_thin
        cell.font = font_saving if ci in (4, 5, 6, 7) else font_bold

    for col in "ABCDEFGH":
        ws2.column_dimensions[col].width = 18

    # ── 시트 ③ 18년 재무 분석 ────────────────────────────────────
    ws3 = wb.create_sheet("③18년_재무_분석")
    ws3.merge_cells("A1:H1")
    ws3["A1"] = "18년 장기 투자 회수 시뮬레이션"
    ws3["A1"].fill = fill_header; ws3["A1"].font = font_white; ws3["A1"].alignment = align_center

    headers3 = ["연도", "물가지수(4%)", "기존 OPEX(만)", "HP OPEX(만)",
                "연간 순이익(만)", "누적 순이익(만)", "ROI", "상태"]
    for ci, h in enumerate(headers3, 1):
        cell = ws3.cell(row=2, column=ci, value=h)
        cell.fill = fill_subhead; cell.font = font_bold
        cell.border = border_thin; cell.alignment = align_center

    ref_capex = "'①입력_가정'!$B$15"  # 순 투자비 셀 참조
    for y in range(1, 19):
        r = y + 2
        ws3.cell(row=r, column=1, value=f"{y}년차").border = border_thin
        ws3.cell(row=r, column=2, value=f"=(1+0.04)^{y-1}").border = border_thin
        ws3.cell(row=r, column=3, value=ann_heat_base).border = border_thin
        ws3.cell(row=r, column=4, value=ann_hp_op).border = border_thin
        ws3.cell(row=r, column=5, value=f"=C{r}-D{r}").border = border_thin
        ws3.cell(row=r, column=6,
                 value=f"=E{r}-{ref_capex}" if y == 1 else f"=F{r-1}+E{r}").border = border_thin
        ws3.cell(row=r, column=7, value=f"=F{r}/{ref_capex}").border = border_thin
        ws3.cell(row=r, column=7).number_format = "0%"
        ws3.cell(row=r, column=8, value=f'=IF(F{r}>0,"수익","회수중")').border = border_thin
        if y % 2 == 0:
            for ci in range(1, 9):
                ws3.cell(row=r, column=ci).fill = fill_blue

    for col in "ABCDEFGH":
        ws3.column_dimensions[col].width = 16

    buf = io.BytesIO()
    wb.save(buf)
    return buf


# ══════════════════════════════════════════════════════════════════════
# 6. UI — 헤더 및 솔루션 개요
# ══════════════════════════════════════════════════════════════════════

excel_data, load_err = load_tariff_xlsx()
if excel_data:
    tariff_blocks = excel_data["blocks"]
    SCOP_BY_ZONE  = excel_data["scop"]          # 동적 로드된 sCOP
    HDD_MONTHLY   = excel_data["hdd"]           # 동적 로드된 HDD (Tbase=18°C)
    MONTHLY_COP   = excel_data["monthly_cop"]   # 월별 COP (참고용)
    MONTHLY_TEMP  = excel_data["monthly_temp"]  # 월평균 기온 (참고용)
    KWH_DEMAND    = excel_data["kwh_demand"]    # Sheet2: 연료별 월별 난방 수요 (참고용, 거창군 기준)
    KWH_HP        = excel_data["kwh_hp"]        # Sheet2: 연료별 월별 HP 전력 사용량 (참고용)
    SHEET2_PARAMS = excel_data["sheet2_params"] # Sheet2: 연료별 효율/단가/기본요금/난방비중
    REGION_RATIOS = excel_data["region_ratios"] # Sheet3: 광역시도별 월별 난방 비중

col_title, col_logo = st.columns([6, 1])
with col_title:
    st.title("히트펌프 경제성 분석 솔루션")
with col_logo:
    if os.path.exists("logo.png"):
        st.image(Image.open("logo.png"), use_container_width=True)

# 로드 실패 시 중단
if load_err:
    st.error(f"⚠️ 전기요금완료본.xlsx 로드 실패: {load_err}")
    st.info("repo 루트(또는 app.py 옆)에 `전기요금완료본.xlsx` 파일이 있는지 확인해 주세요.")
    st.stop()

st.markdown("""
<div class='info-box'>
  <h4 class='info-title'>💡 이 계산기가 왜 필요한가요?</h4>

  <div style='margin-bottom:20px;'>
    <p style='color:#0f172a; font-size:1.05rem; font-weight:600; margin-bottom:6px;'>🌍 왜 지금 히트펌프인가요?</p>
    <p class='info-text'>
      우리나라 가정의 난방은 대부분 <b>가스·등유 보일러</b>로 이루어지고 있고, 이는 가정에서 발생하는
      탄소 배출의 가장 큰 원인입니다. 정부는 <b>2050 탄소중립</b> 목표 달성을 위해 친환경 히트펌프
      전환 시 <b>최대 840만원의 보조금</b>(정부 560 + 지자체 280)을 지원하고 있어요. 가스요금이
      해마다 오르는 만큼, 지금이 우리 집 난방을 바꾸는 게 정말 이득인지 미리 따져볼 좋은 시점입니다.
    </p>
  </div>

  <div style='margin-bottom:20px;'>
    <p style='color:#0f172a; font-size:1.05rem; font-weight:600; margin-bottom:6px;'>🏠 히트펌프(AWHP)가 뭔가요?</p>
    <p class='info-text'>
      <b>공기 중에 있는 열을 모아 난방에 쓰는 친환경 기기</b>입니다. 에어컨이 실내 열을 밖으로 빼내는
      것과 정반대로 작동하는 원리예요. 가스를 태우는 게 아니라 전기로 움직이고, 같은 에너지로
      <b>가스보일러보다 약 3~4배 효율</b>이 나옵니다. 그래서 태양광 설비를 함께 쓰면
      난방비를 거의 0에 가깝게 만들 수도 있습니다.
    </p>
  </div>

  <div>
    <p style='color:#0f172a; font-size:1.05rem; font-weight:600; margin-bottom:6px;'>📝 어떻게 사용하나요?</p>
    <ol style='color:#475569; font-size:1.0rem; line-height:1.85; margin:0; padding-left:22px;'>
      <li>우리 집 <b>지역과 평수</b>를 골라주세요</li>
      <li>지난겨울 <b>1월 난방비와 전기요금</b>, 현재 쓰는 <b>난방 방식</b>을 입력해 주세요
        <span style='color:#94a3b8; font-size:0.9rem;'>(고지서를 참고하시면 가장 정확합니다)</span></li>
      <li>사용 중인 <b>전기 요금제</b>를 5가지 중 하나 골라주세요
        <span style='color:#94a3b8; font-size:0.9rem;'>(태양광 설치 여부 포함)</span></li>
      <li>받을 수 있는 <b>보조금</b>을 체크하세요</li>
      <li><b>'경제성 분석 실행'</b> 버튼을 누르면 끝!
        <span style='color:#94a3b8; font-size:0.9rem;'>절감액·투자 회수 기간이 한 눈에 보이고, 상세 결과는 엑셀로 다운받을 수 있습니다.</span></li>
    </ol>
  </div>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# 7. UI — 입력 섹션
# ══════════════════════════════════════════════════════════════════════

# ── 섹션 1: 대상지 ──
st.markdown('<div class="section-title">1. 대상지 기본 정보</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    region = st.selectbox("광역 지자체", list(REGIONS_FULL.keys()), index=0)
with col2:
    sub_region = st.selectbox("기초 지자체", REGIONS_FULL.get(region, ["전체"]))

col3, col4 = st.columns(2)
with col3:
    house_type = st.selectbox("주거 형태",
                              ["단독 주택 / 다가구 주택", "아파트", "연립 / 빌라 / 다세대 주택"])
with col4:
    house_size = st.number_input("전용 면적 (평)", min_value=10, value=30)

zone        = map_region_to_zone(region, sub_region)
dynamic_cop = SCOP_BY_ZONE[zone]

# ── 섹션 2: 에너지 소비 ──
st.markdown('<div class="section-title">2. 에너지 소비 현황</div>', unsafe_allow_html=True)
col_h, col_e = st.columns(2)
with col_h:
    winter_heat_man = st.number_input("동절기(1월) 평균 난방비 (만원)", value=20)
with col_e:
    winter_elec_man = st.number_input("동절기(1월) 전기요금 (만원)", value=6)

col_ht, col_ck = st.columns(2)
with col_ht:
    heating_type = st.selectbox(
        "현재 주택의 난방 방식",
        list(HEATING_TYPE_MAP.keys()),
        help="현재 사용 중인 난방 연료 방식을 선택해 주세요."
    )
with col_ck:
    cooking_type = st.selectbox(
        "사용하는 취사 기기",
        ["인덕션 (전기)", "도시가스", "LPG"],
    )

# ── 섹션 3: 시뮬레이션 변수 ──
st.markdown('<div class="section-title">3. 시뮬레이션 상수 변수</div>', unsafe_allow_html=True)
col_sim, col_opt = st.columns(2)
with col_sim:
    fuel_inflation = st.slider("가스/등유요금 인상률 (%)", 0.0, 15.0, 5.0)
    elec_inflation = st.slider("전기요금 인상률 (%)",    0.0, 15.0, 3.0)
    solar_capa_kw  = st.number_input(
        "태양광 용량 (kW)",
        value=0.0,
        help="태양광 설치 시 발전 용량을 입력해 주세요. (요금제 선택에서 '태양광 설치' 옵션을 함께 골라야 적용됩니다)"
    )

with col_opt:
    use_subsidy_nat = st.checkbox(f"정부 보조금 적용 ({SUBSIDY_NATIONAL}만원)", value=True)
    is_southern    = region in SOUTHERN_REGIONS
    use_subsidy_loc = st.checkbox(f"지자체 매칭 보조금 적용 ({SUBSIDY_LOCAL}만원)", value=is_southern)
    st.caption("*2026년 현재 제주, 경남, 전남은 보조금 신청이 가능합니다.")

    st.markdown("---")
    st.markdown("**전기 요금제 선택**")
    st.markdown("""
<div class='help-text'>
사용 중인 요금제와 태양광 설치 여부에 맞춰 선택해 주세요.
</div>""", unsafe_allow_html=True)

    tariff_label = st.radio(
        "요금제",
        list(TARIFF_LABEL_MAP.keys()),
        label_visibility="collapsed",
        help="누진제: 주택용 저압 누진제 / 일반용: HP 전용 별도 미터 (태양광 영향 없음) / 계시별: 시간대별 요금제",
    )

# 입력 변경 감지 → 분석 결과 초기화
_input_signature = (winter_heat_man, winter_elec_man, region, house_type, house_size,
                    heating_type, cooking_type, fuel_inflation, elec_inflation,
                    solar_capa_kw, use_subsidy_nat, use_subsidy_loc, tariff_label)
if st.session_state.get("_last_input_key") != _input_signature:
    st.session_state.analyzed = False
    st.session_state["_last_input_key"] = _input_signature

if "analyzed" not in st.session_state:
    st.session_state.analyzed = False

if st.button("경제성 분석 실행", type="primary", use_container_width=True):
    st.session_state.analyzed = True


# ══════════════════════════════════════════════════════════════════════
# 8. 분석 결과
# ══════════════════════════════════════════════════════════════════════

if st.session_state.analyzed:

    # ─── 8-1. 핵심 계산 ───────────────────────────────────────────────
    block_key, tariff_choice, solar_flag = get_block_key(tariff_label, heating_type)
    block    = tariff_blocks[block_key]
    fuel_key = HEATING_TYPE_MAP[heating_type]   # 이후 모든 곳에서 재사용

    # 광역시도별 월별 난방 비중 가져오기 (Sheet3) — 매핑 실패 시 전국 평균 사용
    sheet3_region_name = REGION_NAME_MAP.get(region, "전국")
    monthly_ratios = REGION_RATIOS.get(sheet3_region_name, REGION_RATIOS.get("전국"))

    # 가구 규모 보정 — 사용자 지역의 월별 비중으로 표준 1월 난방비 산정
    csv_jan_man = calc_standard_jan_heat_man(monthly_ratios)
    scale       = (winter_heat_man / csv_jan_man) if csv_jan_man > 0 else 1.0
    result      = apply_block_with_scale(block, scale)

    # 보조금 및 투자비
    total_subsidy = (SUBSIDY_NATIONAL if use_subsidy_nat else 0) + (SUBSIDY_LOCAL if use_subsidy_loc else 0)
    capex_man     = 1000  # 국내 기업 자료 (설치비 포함)
    net_capex_man = max(0, capex_man - total_subsidy)

    # 18년 시뮬레이션
    ann_heat_base = result["ex_annual_man"]
    ann_hp_op     = result["hp_annual_man"]
    years, gas_cum, hp_cum, net_profit, payback_year = simulate_18yr(
        net_capex_man, ann_heat_base, ann_hp_op, fuel_inflation, elec_inflation
    )

    # 월별 기존 난방비 — 사용자 광역시도의 월별 비중으로 안분 (Sheet3 데이터)
    # 비난방월에도 온수 사용분이 있어 0이 되지 않음
    months = list(range(1, 13))
    hdd_zone = HDD_MONTHLY[zone]            # 비고 표시용 (난방월/비난방월 판단)
    jan_ratio = monthly_ratios[0] if monthly_ratios[0] > 0 else 1
    monthly_ex_man = [round(winter_heat_man * monthly_ratios[m-1] / jan_ratio, 2) for m in months]
    monthly_stats  = calc_monthly_stats(monthly_ex_man, result["monthly_man"], fuel_key)

    # 연간 CO₂ 절감 + 환경 비유
    co2_saving_kg, trees, car_km = calc_annual_co2_savings(
        result["ex_annual_man"], result["hp_annual_man"], fuel_key
    )

    # 에너지 사용량 (kWh) — Sheet2 행 43 수식 그대로 복제
    # 사용자 연간 난방비 = 1월 입력 ÷ 1월 비중 (지역별)
    user_annual_cost_won = winter_heat_man * 10000 / monthly_ratios[0] if monthly_ratios[0] > 0 else 0
    # Sheet2 행 43 수식으로 연간 난방 에너지 수요 산정
    user_annual_demand_kwh = calc_demand_kwh_sheet2(user_annual_cost_won, fuel_key, SHEET2_PARAMS)
    # 월별 분배 + HP 변환
    kwh = calc_kwh_data(user_annual_demand_kwh, monthly_ratios, MONTHLY_COP[zone])

    # ─── 8-2. 결과 요약 헤더 ──────────────────────────────────────────
    st.markdown('<div class="section-title">📊 분석 결과 요약</div>', unsafe_allow_html=True)

    # 배지 색상: 태양광 설치=초록, 미설치=노랑, 일반용=파랑
    if solar_flag == "태O":
        badge_cls = "solar-badge-o"
    elif tariff_choice == "일반용":
        badge_cls = "tariff-badge"
    else:
        badge_cls = "solar-badge-x"

    scale_tooltip = (
        f"표준 가구(연 65만원, 1월 약 {csv_jan_man:.1f}만원) 대비 우리 집 난방 규모입니다. "
        f"입력하신 1월 난방비 {winter_heat_man}만원 ÷ 표준 {csv_jan_man:.1f}만원 = ×{round(scale, 2)}배. "
        f"이 비율이 모든 절감액·HP 요금 계산에 자동 반영됩니다."
    )

    st.markdown(f"""
<div style='margin-bottom:16px;'>
  <span class='{badge_cls}'>{tariff_label}</span>
  <span class='tariff-badge'>난방: {fuel_key}</span>
  <span class='tariff-badge has-tooltip' data-tooltip="{scale_tooltip}">규모 보정: ×{round(scale, 2)} ⓘ</span>
</div>
""", unsafe_allow_html=True)

    # 핵심 지표 4개
    hp_space, hp_space_mm, hp_capacity = get_hp_specs(house_size)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("투자 회수 시점", payback_year)
    m2.metric("18년 순이익", f"{net_profit[-1]:,} 만원")
    m3.metric("히트펌프 설치 공간", hp_space)
    m3.markdown(f"<div style='font-size:0.78rem; color:#64748b; margin-top:-10px;'>{hp_space_mm}</div>",
                unsafe_allow_html=True)
    m4.metric("적정 히트펌프 용량", hp_capacity)

    # ─── 8-3. 전기요금 분석 ──────────────────────────────────────────
    st.markdown('<div class="section-title">💰 전기요금 분석 (엑셀 데이터 기반)</div>', unsafe_allow_html=True)
    s1, s2, s3 = st.columns(3)
    s1.metric(
        "HP 연간 전기요금",
        f"{result['hp_annual_man']:,.1f} 만원",
        help=f"[{tariff_label} / {fuel_key}] 블록 HP 연합계 × 규모 보정"
    )
    s2.metric(
        f"기존 연간 난방비 ({fuel_key})",
        f"{result['ex_annual_man']:,.1f} 만원",
        help="엑셀 표준 가구 기존 난방비 × 규모 보정"
    )
    s3.metric(
        "연간 절감액",
        f"{result['saving_man']:,.1f} 만원",
        delta=f"{round(result['saving_ratio'] * 100)}% 절감",
    )

    # ─── 8-4. 환경 기여 박스 ──────────────────────────────────────────
    st.markdown(f"""
<div class='saving-box'>
  <p class='saving-title'>🌳 우리 가족의 1년 환경 기여</p>
  <p class='saving-sub'>
    히트펌프 전환 시 연간 약 <b>{co2_saving_kg:,.0f} kgCO₂</b>를 줄일 수 있어요.
    <br>
    이는 🌲 <b>30년생 소나무 약 {trees:,}그루</b>가 1년 동안 흡수하는 양이거나,
    🚗 <b>자동차로 약 {car_km:,} km</b>를 운행할 때 발생하는 배출량과 같습니다.
  </p>
</div>
""", unsafe_allow_html=True)

    # ─── 8-4-2. 에너지 효율 비교 박스 (Sheet2 kWh 기반) ───────────────
    fuel_unit_label = "도시가스" if "도시가스" in fuel_key else fuel_key
    st.markdown(f"""
<div class='saving-box' style='background:#eff6ff; border-color:#93c5fd;'>
  <p class='saving-title' style='color:#1e40af;'>⚡ 우리 가족의 1년 에너지 사용량</p>
  <p class='saving-sub' style='color:#1e3a8a;'>
    기존 보일러로는 연간 <b>{kwh['annual_demand']:,.0f} kWh</b>의 {fuel_unit_label} 에너지가 필요해요.
    <br>
    히트펌프로 바꾸면 같은 따뜻함을 만드는 데 <b>전기 {kwh['annual_hp']:,.0f} kWh</b>면 충분합니다 —
    🔥 <b>약 {kwh['efficiency']}배 효율</b>로 같은 난방을 1/{kwh['efficiency']:.0f} 에너지로!
  </p>
</div>
""", unsafe_allow_html=True)

    # ─── 8-5. 월별 차트 ──────────────────────────────────────────────
    df_chart = pd.DataFrame({
        "월":               [f"{m}월" for m in months],
        "기존 난방비(만원)": monthly_ex_man,
        "HP 난방요금(만원)": result["monthly_man"],
    }).melt("월", var_name="구분", value_name="금액(만원)")

    chart = alt.Chart(df_chart).mark_bar().encode(
        x=alt.X("월:O", sort=[f"{m}월" for m in months]),
        y=alt.Y("금액(만원):Q"),
        color=alt.Color("구분:N", scale=alt.Scale(
            domain=["기존 난방비(만원)", "HP 난방요금(만원)"],
            range=["#f87171", "#60a5fa"]
        ), legend=alt.Legend(orient="top", title=None)),
        xOffset="구분:N",
        tooltip=["월", "구분", "금액(만원)"],
    ).properties(height=280, title="월별 기존 난방비 vs HP 난방요금")
    st.altair_chart(chart, use_container_width=True)

    # ─── 8-6. 월별 상세 테이블 ────────────────────────────────────────
    with st.expander("📋 월별 상세 데이터"):
        df_detail = pd.DataFrame({
            "월":                [f"{m}월" for m in months],
            "기존 난방비(만원)": monthly_ex_man,
            "HP 난방요금(만원)": result["monthly_man"],
            "월별 절감액(만원)": monthly_stats["savings"],
            "절감률":            monthly_stats["pct"],
            "누적 절감액(만원)": monthly_stats["cumulative"],
            "CO₂ 절감(kg)":     monthly_stats["co2"],
            "HP 전력 사용(kWh)": kwh["monthly_hp"],
        })
        st.dataframe(df_detail, use_container_width=True, hide_index=True)
        st.caption(
            "💡 **HP 난방요금**은 히트펌프의 난방·온수·냉방 전기 비용입니다 "
            "(가전·취사 사용분은 제외 — 기존 난방비와 동일 기준 비교)."
        )
        st.caption(
            f"📅 **월별 분배 기준**: 입력하신 광역시도({region})의 평균 월별 난방 비중(엑셀 Sheet3)을 사용자의 1월 입력값에 비례하여 분배한 추정치입니다. "
            "여름철에도 온수·취사 등으로 일부 가스 비용이 발생하는 것이 반영되어 있습니다."
        )
        st.caption(
            f"⚡ **HP 전력 사용(kWh)**은 누진제 사용자가 자기 누진 구간을 예측하는 데 도움이 됩니다 "
            f"(예: 1월 {kwh['monthly_hp'][0]:.0f} kWh 사용 → 전체 가전과 합쳐 누진 단계 추정 가능)."
        )
        st.caption(
            f"📊 CO₂는 환경부 GIR 배출계수와 평균 단가({fuel_key}: {CO2_PER_MAN_FUEL[fuel_key]}kg/만원, "
            f"전기: {CO2_PER_MAN_ELEC}kg/만원) 기반 추정치입니다."
        )

    # ─── 8-7. 18년 차트 ──────────────────────────────────────────────
    st.markdown('<div class="section-title">📈 18년 장기 시뮬레이션</div>', unsafe_allow_html=True)
    g1, g2 = st.columns(2)
    with g1:
        st.write("**18년 누적 비용 흐름**")
        df_cum = pd.DataFrame({"연도": years, "기존": gas_cum, "HP": hp_cum}) \
                   .melt("연도", var_name="시나리오", value_name="비용")
        st.altair_chart(
            alt.Chart(df_cum).mark_area(opacity=0.5).encode(
                x="연도:O", y="비용:Q", color="시나리오:N"
            ),
            use_container_width=True
        )
    with g2:
        st.write("**연도별 순수익(Cash Flow)**")
        df_profit = pd.DataFrame({
            "연도":   years,
            "순수익": net_profit,
            "상태":   ["수익" if p > 0 else "회수" for p in net_profit],
        })
        st.altair_chart(
            alt.Chart(df_profit).mark_bar().encode(
                x="연도:O", y="순수익:Q", color="상태:N"
            ),
            use_container_width=True
        )

    # ─── 8-8. 가정값 상세 ────────────────────────────────────────────
    with st.expander("📋 적용된 계산 가정값 및 출처"):
        subsidy_text = ("정부 560 + 지방 280" if use_subsidy_nat and use_subsidy_loc
                        else "정부 560만원" if use_subsidy_nat
                        else "지방 280만원" if use_subsidy_loc
                        else "없음")
        st.markdown(f"""
| 항목 | 적용값 | 근거 |
|------|--------|------|
| 적용 요금제 | **{tariff_label}** | 전기요금완료본.xlsx |
| 적용 난방 유형 | **{fuel_key}** | 사용자 선택 |
| 규모 보정 계수 | **×{round(scale, 2)}** | 사용자 1월 난방비({winter_heat_man}만원) ÷ 엑셀 기준({csv_jan_man:.2f}만원) |
| 설비 CAPEX | **{capex_man}만원** | 국내 기업 자료 (설치비 포함) |
| 정부+지방 보조금 | **{total_subsidy}만원** | {subsidy_text} |
| 순 투자비 | **{net_capex_man}만원** | CAPEX − 보조금 |
| 기존 연간 난방비 | **{ann_heat_base}만원** | 엑셀 기존난방비 × 규모 보정 |
| HP 연간 전기요금 | **{ann_hp_op}만원** | 엑셀 HP연합계 × 규모 보정 |
| 지역 sCOP (참고) | **{dynamic_cop}** | 기후 존 ({zone}) 추정값 |
| 월별 비중 출처 | **{REGION_NAME_MAP.get(region, '전국')}** | 엑셀 Sheet3 광역시도별 데이터 |
        """)

    # ─── 8-9. 엑셀 다운로드 ──────────────────────────────────────────
    excel_buffer = build_excel_report(
        region=region, tariff_label=tariff_label, fuel_key=fuel_key,
        winter_heat_man=winter_heat_man, winter_elec_man=winter_elec_man,
        solar_capa_kw=solar_capa_kw, scale=scale, csv_jan_man=csv_jan_man,
        dynamic_cop=dynamic_cop, zone=zone,
        capex_man=capex_man, use_subsidy_nat=use_subsidy_nat, use_subsidy_loc=use_subsidy_loc,
        net_capex_man=net_capex_man, ann_heat_base=ann_heat_base, ann_hp_op=ann_hp_op,
        result=result, monthly_ex_man=monthly_ex_man, hdd_zone=hdd_zone,
    )

    st.markdown("---")
    st.download_button(
        label="🚀 전문가용 수식 연동 정밀 엑셀 다운로드",
        data=excel_buffer.getvalue(),
        file_name=f"Expert_Report_{region}_{safe_filename(tariff_label)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )