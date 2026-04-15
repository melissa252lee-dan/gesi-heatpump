import streamlit as st
import pandas as pd
import numpy as np
import io, os
from PIL import Image
from openpyxl import Workbook
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
.saving-box    { background:#f0fdf4; border:2px solid #86efac; border-radius:12px; padding:20px 24px; margin:16px 0; }
.saving-title  { color:#15803d; font-size:1.1rem; font-weight:700; margin-bottom:4px; }
.saving-sub    { color:#166534; font-size:0.9rem; }
.warn-box      { background:#fffbeb; border:1px solid #fcd34d; border-radius:8px; padding:12px 16px; margin:8px 0; font-size:0.85rem; color:#92400e; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# 2. 정적 데이터 정의
# ══════════════════════════════════════════════════════════

# ── 광역·기초 지자체 목록 ──
# 출처: 행정안전부 법정동 코드 기준 (2025년 기준)
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

# ── 태양광 월별 발전량 (kWh/kW, 1kW 설치 기준) ──
# 출처: 한국에너지공단 지역별 태양광 발전량 통계 (2025년 기준 추정값)
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

# ── HDD(난방도일) 월별 데이터 ──
# 출처: COP_계산기.csv (기준온도 Tbase=15°C)
hdd_monthly = {
    "중부1": [750, 596, 456, 198,   0,   0,   0,   0,   0, 161, 405, 676],  # 철원 기준
    "중부2": [521, 398, 264,  54,   0,   0,   0,   0,   0,   0, 225, 450],  # 서울 기준
    "남부":  [347, 269, 171,   0,   0,   0,   0,   0,   0,   0, 114, 295],  # 부산 기준
    "제주":  [273, 224, 140,   0,   0,   0,   0,   0,   0,   0,  45, 211],  # 제주 기준
}

# ── 전기요금 누진제 fallback 하드코딩 ──
# 출처: KEPCO 주택용 전력(저압) 요금표 2024년 기준
# CSV 로드 실패 시 이 값으로 대체
# 구조: {시즌: {"단계": [(상한kWh, 기본요금, 단가), ...], "연료비조정": 원/kWh}}
# 시즌 구분: "기타계절"=1~6월·9~12월, "하계"=7~8월
TARIFF_FALLBACK = {
    "기타계절": {
        "구간": [
            (200,  910,  120.0),   # 200kWh 이하
            (400, 1600,  214.6),   # 201~400kWh
            (999, 7300,  307.3),   # 400kWh 초과
        ],
        "기후환경요금":   14.0,    # 원/kWh (2024 기준)
        "연료비조정요금":  5.3,    # 원/kWh (평균값)
    },
    "하계": {
        "구간": [
            (300,  910,  120.0),   # 300kWh 이하
            (450, 1600,  214.6),   # 301~450kWh
            (999, 7300,  307.3),   # 450kWh 초과
        ],
        "기후환경요금":   14.0,
        "연료비조정요금":  5.3,
    },
}

# ── 도시가스 콘덴싱→HP Saving fallback ──
# 출처: 전기요금누진제.csv row1~13 (col15=청구요금합계, 원 단위)
# CSV 로드 실패 시 이 월별 청구요금(원)을 사용
# 1월~12월 순서
SAVING_CONDENSING_FALLBACK = {
    "hp_monthly_billing_won": [56580, 40400, 23350, 11400, 8700, 6720,
                                6050,  5620,  5760,  8550, 19250, 45540],
    "hp_annual_total_won":   237920,       # HP 전기요금 연합계 (원)
    "existing_annual_won":   599269,       # 기존 난방비 연합계 (원, 도시가스 콘덴싱 기준)
    "saving_ratio":          0.60,         # 난방비 Saving 비율
}


# ══════════════════════════════════════════════════════════
# 3. CSV 로더 함수
# ══════════════════════════════════════════════════════════

@st.cache_data
def load_tariff_csv():
    """
    전기요금누진제.csv를 읽어 요금 구조와 도시가스 콘덴싱→HP Saving 데이터를 파싱합니다.

    [파일 구조]
    - col 0~5  : KEPCO 주택용 저압 요금표 (기타계절/하계 구간·단가)
    - col 7~15 : 난방 유형별 HP 전환 시 월별 청구요금 계산 결과
                 col7=월번호, col8=기본요금, col9=사용량요금, col10=기후환경요금,
                 col11=연료비조정, col12=전기요금계, col13=VAT, col14=기금, col15=청구합계
    - col 16   : HP 전기요금 연합계 (원/년)
    - col 17   : 기존 난방비 연합계 (원/년)
    - col 18   : 난방비 Saving 비율

    [난방 유형별 행 범위]
    - row 1~13  : 도시가스(콘덴싱) → HP   (row1=헤더, row2~13=1~12월)
    - row 14~26 : 도시가스(일반)   → HP
    - row 27~39 : 등유             → HP
    - row 40~52 : LPG              → HP

    [반환값]
    {
      "tariff": {
          "기타계절": {"구간": [...], "기후환경요금": float, "연료비조정요금": float},
          "하계":     {"구간": [...], ...},
      },
      "saving": {
          "condensing": {
              "hp_monthly_billing_won": [1월~12월 청구요금(원), ...],
              "hp_annual_total_won":    연합계(원),
              "existing_annual_won":    기존난방비(원),
              "saving_ratio":           Saving비율(0~1),
          }
      },
      "source": "csv" | "fallback"
    }
    """
    def read_safe(f):
        try:    return pd.read_csv(f, encoding="utf-8",  header=None)
        except: return pd.read_csv(f, encoding="cp949", header=None)

    try:
        # CSV 경로 탐색 — Streamlit Cloud는 cwd가 repo 루트이므로 파일명 직접 참조가 우선
        # 로컬 실행 시에는 app.py 위치 기준으로 fallback
        _fname = "전기요금누진제.csv"
        _candidates = [
            _fname,                                                              # Streamlit Cloud: cwd = repo 루트
            os.path.join(os.path.dirname(os.path.abspath(__file__)), _fname),   # 로컬: app.py 옆
            os.path.join(os.getcwd(), _fname),                                  # cwd 명시
        ]
        csv_path = next((p for p in _candidates if os.path.exists(p)), _fname)
        df = read_safe(csv_path)

        # ── 기타계절 요금표 파싱 (row2~4) ──
        # col0=구간명, col1=기본요금, col2=단계명, col3=단가(원/kWh)
        def parse_basic(val):
            """기본요금 문자열 '1,600' → 1600 정수 변환"""
            return int(str(val).replace(",", "").strip())

        tariff_other = {
            "구간": [
                (200,  parse_basic(df.iloc[2][1]),  float(df.iloc[2][3])),  # 200이하
                (400,  parse_basic(df.iloc[3][1]),  float(df.iloc[3][3])),  # 201~400
                (9999, parse_basic(df.iloc[4][1]),  float(df.iloc[4][3])),  # 400초과
            ],
            "기후환경요금":   14.0,   # KEPCO 고정 (CSV에 별도 행 없음)
            "연료비조정요금":  5.3,   # 월별 평균값 (col11 평균: ~1~12월 합/12)
        }

        # ── 하계 요금표 파싱 (row11~13) ──
        tariff_summer = {
            "구간": [
                (300,  parse_basic(df.iloc[11][1]),  float(df.iloc[11][3])),  # 300이하
                (450,  parse_basic(df.iloc[12][1]),  float(df.iloc[12][3])),  # 301~450
                (9999, parse_basic(df.iloc[13][1]),  float(df.iloc[13][3])),  # 450초과
            ],
            "기후환경요금":   14.0,
            "연료비조정요금":  5.3,
        }

        # ── 도시가스 콘덴싱→HP 월별 청구요금 파싱 (row2~13) ──
        # col7=월번호(1~12), col15=청구요금합계(원)
        hp_monthly = []
        for i in range(2, 14):
            billing = int(str(df.iloc[i][15]).replace(",", "").strip())
            hp_monthly.append(billing)

        # 연합계·기존난방비·Saving (row1, col16~18)
        hp_annual    = int(str(df.iloc[1][16]).replace(",", "").strip())
        existing_ann = int(float(str(df.iloc[1][17]).replace(",", "").strip()))
        saving_str   = str(df.iloc[1][18]).replace("%", "").strip()
        saving_ratio = float(saving_str) / 100.0

        return {
            "tariff": {
                "기타계절": tariff_other,
                "하계":     tariff_summer,
            },
            "saving": {
                "condensing": {
                    "hp_monthly_billing_won": hp_monthly,
                    "hp_annual_total_won":    hp_annual,
                    "existing_annual_won":    existing_ann,
                    "saving_ratio":           saving_ratio,
                }
            },
            "source": "csv",
        }

    except Exception as e:
        # CSV 로드 실패 → fallback 사용
        return {
            "tariff":  TARIFF_FALLBACK,
            "saving": {"condensing": SAVING_CONDENSING_FALLBACK},
            "source": "fallback",
            "error":   str(e),
        }


@st.cache_data
def load_simulation_data():
    """
    외기온도 및 sCOP 데이터를 CSV에서 로드합니다.

    [외기온도_시간분포.csv]
    - 생성 방식: 수학적 모델 추정값 (기상청 실측 원시 데이터 아님)
    - 계산식: T(h) = T월평균 + (DTR/2) × cos(π × (h-14) / 12)
        · T월평균 : 지역별 월평균 기온 (기상청 기후 통계 기반 추정, 2025년 기준)
        · DTR     : 일교차 (Diurnal Temperature Range)
        · h       : 시각 (0~23시), 오후 2시(h=14)가 일 최고기온
    - 대표 지점: 중부1=철원, 중부2=서울, 남부=부산, 제주=제주

    [COP_계산기.csv]
    - sCOP 계산식: 카르노(Carnot) 열펌프 효율 공식 기반
        COP(월,h) = η × (Ts+273) / max(1, Ts-T(월,h))
        sCOP      = Σ(월난방수요 × 월COP) / Σ(월난방수요)
    - 입력 파라미터 (설계 가정값):
        Ts=45°C, η=0.50, Tbase=15°C
    - fallback: 중부1=3.29, 중부2=3.66, 남부=3.99, 제주=4.21
    """
    try:
        def read_csv_safe(f):
            try:    return pd.read_csv(f, encoding="utf-8",  header=None)
            except: return pd.read_csv(f, encoding="cp949", header=None)

        df_t, df_c = read_csv_safe("외기온도_시간분포.csv"), read_csv_safe("COP_계산기.csv")
        temp_data, cop_data = {}, {}
        zones = ["중부1", "중부2", "남부", "제주"]
        curr = None

        for _, r in df_t.iterrows():
            line = str(r[0]).strip()
            for z in zones:
                if z in line and "▶" in line:
                    curr = z; temp_data[curr] = []; break
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
                    try:
                        scop = float(r[15])
                        # 월별 COP (col2~13, 1~12월) — 난방 없는 달은 0
                        monthly_cop = []
                        for ci in range(2, 14):
                            try:
                                v = float(r[ci])
                                monthly_cop.append(v if v > 0 else 0.0)
                            except:
                                monthly_cop.append(0.0)
                        cop_data[z] = {"scop": scop, "monthly_cop": monthly_cop}
                    except: continue

        # fallback: monthly_cop는 sCOP 단일값으로 대체 (난방월만 채움)
        fallback_scop  = {"중부1": 3.29, "중부2": 3.66, "남부": 3.99, "제주": 4.21}
        fallback_mcop  = {
            "중부1": [2.94, 3.10, 3.52, 0, 0, 0, 0, 0, 0, 4.10, 3.61, 3.07],
            "중부2": [3.39, 3.57, 4.00, 0, 0, 0, 0, 0, 0, 0,    4.08, 3.55],
            "남부":  [3.82, 3.95, 4.23, 0, 0, 0, 0, 0, 0, 0,    4.31, 3.96],
            "제주":  [4.06, 4.13, 4.44, 0, 0, 0, 0, 0, 0, 0,    4.67, 4.25],
        }
        for k in ["중부1", "중부2", "남부", "제주"]:
            if k not in cop_data:
                cop_data[k] = {"scop": fallback_scop[k], "monthly_cop": fallback_mcop[k]}
        return temp_data, cop_data
    except:
        return None, None


# ══════════════════════════════════════════════════════════
# 4. 계산 함수
# ══════════════════════════════════════════════════════════

def map_region_to_zone(s):
    """
    광역 지자체명을 기후 존(zone)으로 매핑합니다.
    출처: 국토교통부 건물 에너지 효율 설계 기준
    중부1=강원도, 남부=대구·부산·울산·광주·경남·전남, 제주=제주도, 나머지=중부2
    """
    if s == "강원도": return "중부1"
    if s in ["대구","부산","울산","광주","경상남도","전라남도"]: return "남부"
    if s == "제주도": return "제주"
    return "중부2"


def get_season(month):
    """
    월(1~12)을 KEPCO 요금 시즌으로 변환합니다.
    출처: KEPCO 주택용 저압 요금표
    - 하계: 7~8월
    - 기타계절: 그 외 나머지 10개월
    """
    return "하계" if month in (7, 8) else "기타계절"


def calc_elec_bill_from_tariff(kwh, month, tariff_data):
    """
    전기 사용량(kWh)과 월(시즌)을 받아 월 청구요금(원)을 계산합니다.

    출처: 전기요금누진제.csv 파싱 결과 (없으면 TARIFF_FALLBACK)

    계산식:
        ① 해당 월의 시즌(기타계절/하계) 결정
        ② 사용량에 따른 기본요금·누진 구간 단가 결정
        ③ 전력량요금 = 구간별 누진 계산
        ④ 기후환경요금 = 14.0원/kWh × kWh
        ⑤ 연료비조정요금 = 해당 월 실제값 (CSV) 또는 평균값 (fallback)
        ⑥ 전기요금계 = 기본요금 + ③ + ④ + ⑤
        ⑦ VAT = 전기요금계 × 10%
        ⑧ 전력산업기반기금 = 전기요금계 × 3.7%
        ⑨ 청구요금 = ⑥ + ⑦ + ⑧

    반환: 청구요금 (원, 정수)
    """
    kwh    = max(0.0, kwh)
    season = get_season(month)
    t      = tariff_data["tariff"][season]
    구간    = t["구간"]

    # ① 기본요금·전력량요금 누진 계산
    basic_fee    = 0
    energy_charge = 0.0
    prev_limit   = 0

    for (upper, basic, unit_price) in 구간:
        if kwh <= prev_limit:
            break
        basic_fee    = basic  # 해당 구간의 기본요금 (마지막으로 해당되는 구간)
        used_in_band = min(kwh, upper) - prev_limit
        energy_charge += used_in_band * unit_price
        prev_limit    = upper

    # ② 기타 항목
    climate_fee   = t["기후환경요금"] * kwh      # 기후환경요금
    fuel_adj      = t["연료비조정요금"] * kwh     # 연료비조정요금 (평균)
    subtotal      = basic_fee + energy_charge + climate_fee + fuel_adj
    vat           = subtotal * 0.10
    infra_fund    = subtotal * 0.037
    total         = subtotal + vat + infra_fund

    return round(total)


def calc_elec_bill_won_to_man(kwh, month, tariff_data):
    """calc_elec_bill_from_tariff 결과를 만원 단위로 반환 (기존 코드 호환용)"""
    return round(calc_elec_bill_from_tariff(kwh, month, tariff_data) / 10000, 4)


def reverse_kwh_from_tariff(bill_man, month, tariff_data):
    """
    월 전기요금(만원)에서 전기 사용량(kWh)을 이진탐색으로 역산합니다.
    calc_elec_bill_won_to_man()이 단조증가 함수임을 이용.
    탐색 범위: 0~3,000kWh, 40회 반복 → 오차 < 0.00001kWh
    """
    lo, hi = 0.0, 3000.0
    for _ in range(40):
        mid = (lo + hi) / 2
        if calc_elec_bill_won_to_man(mid, month, tariff_data) < bill_man: lo = mid
        else: hi = mid
    return round(mid, 1)


def calc_capex(h_type, h_size):
    """
    주거 형태와 면적을 기반으로 히트펌프 설치 총비용(CAPEX, 만원)을 추정합니다.

    출처:
    - 에너지경제연구원 「세계 히트펌프 시장 및 정책 동향과 국내 시사점」(2025)
      → 공동주택 기준 850만원 (LCOH 분석 기준값)
    - 기후에너지환경부 정책 브리핑 (2026.03)
      → 본체 550~700만원, 급탕조 200~300만원, 공사 100만원

    [아파트] = 850만원 고정
    [단독·연립·빌라] = 975만원 + 3만원/평
    """
    if "아파트" in h_type:
        return 850
    return 975 + h_size * 3


def calc_hdd_ratio(zone):
    """
    HDD 데이터 기반으로 '1월 난방비 대비 연간 난방비 배율'을 계산합니다.
    출처: COP_계산기.csv HDD 데이터 (Tbase=15°C)
    결과: 중부1≈4.32, 중부2≈3.67, 남부≈3.44, 제주≈3.27
    """
    hdd   = hdd_monthly[zone]
    hdd_j = hdd[0]
    return round(sum(hdd) / hdd_j, 2) if hdd_j > 0 else 3.5


def calc_pv_saving(s_reg, s_capa, month_kwh, tariff_data):
    """
    태양광 자가발전에 의한 연간 전기요금 절감액(만원)을 계산합니다.

    계산 방식:
    - 각 월: pv_kwh = pv_monthly_data[지역][월] × 설치용량(kW)
    - 절감액 = calc_elec_bill(month_kwh) - calc_elec_bill(max(0, month_kwh - pv_kwh))
    - 연간 합산

    출처: pv_monthly_data (한국에너지공단 기반 추정), KEPCO 누진제 모델
    주의: 자가소비 100% 가정, 잉여 역전송 미반영
    """
    if s_capa <= 0: return 0.0
    pv_region  = pv_monthly_data.get(s_reg, pv_monthly_data["서울"])
    total_save = 0.0
    for m in range(12):
        pv_kwh    = pv_region[m] * s_capa
        b_before  = calc_elec_bill_won_to_man(month_kwh, m + 1, tariff_data)
        b_after   = calc_elec_bill_won_to_man(max(0, month_kwh - pv_kwh), m + 1, tariff_data)
        total_save += (b_before - b_after)
    return round(total_save, 4)


def calc_condensing_saving(tariff_csv, scale=1.0):
    """
    도시가스 콘덴싱 보일러 → HP 전환 시 난방비 Saving을 계산합니다.

    [테스트 범위: 전기요금 1종(누진제) + 태양광 0인 가구]

    출처: 전기요금누진제.csv — col7~15(월별 HP 청구요금), col16~18(연합계·Saving)

    scale 파라미터:
    - CSV의 기준 가구(도시가스 콘덴싱)와 실제 사용자의 난방 규모 비율
    - 예: scale=1.0 → CSV 계산값 그대로 사용
    - 예: scale=1.5 → 1.5배 규모 가구

    반환:
    {
        "monthly_hp_won":    [1~12월 HP 청구요금 (원)],  ← CSV 직접 값
        "monthly_hp_man":    [1~12월 HP 청구요금 (만원)],
        "hp_annual_man":     HP 연간 전기요금 (만원),
        "existing_annual_man": 기존 난방비 (만원),
        "saving_man":        연간 절감액 (만원),
        "saving_ratio":      절감 비율 (0~1),
        "source":            "csv" | "fallback",
    }
    """
    data   = tariff_csv["saving"]["condensing"]
    source = tariff_csv["source"]

    monthly_won = [round(v * scale) for v in data["hp_monthly_billing_won"]]
    hp_ann_man  = round(sum(monthly_won) / 10000, 1)
    ex_ann_man  = round(data["existing_annual_won"] * scale / 10000, 1)
    save_man    = round(ex_ann_man - hp_ann_man, 1)
    save_ratio  = data["saving_ratio"]  # CSV에 이미 계산된 비율 사용

    return {
        "monthly_hp_won":      monthly_won,
        "monthly_hp_man":      [round(v / 10000, 2) for v in monthly_won],
        "hp_annual_man":       hp_ann_man,
        "hp_annual_total_won": sum(monthly_won),        # 엑셀 합계 행에서 사용 (원 단위)
        "existing_annual_man": ex_ann_man,
        "saving_man":          save_man,
        "saving_ratio":        save_ratio,
        "source":              source,
    }


# ══════════════════════════════════════════════════════════
# 5. UI 메인
# ══════════════════════════════════════════════════════════

# CSV 로드 (앱 시작 시 1회)
tariff_csv  = load_tariff_csv()
df_temp, df_cop = load_simulation_data()

col_t, col_l = st.columns([6, 1])
with col_t: st.title("히트펌프 경제성 분석 솔루션")
with col_l:
    if os.path.exists("logo.png"): st.image(Image.open("logo.png"), use_container_width=True)

# CSV 로드 실패 시에만 경고 표시 (성공 시 별도 메시지 없음)
if tariff_csv["source"] != "csv":
    st.warning(f"⚠️ 전기요금누진제.csv 로드 실패 → 하드코딩 fallback 요금 적용 중  |  사유: {tariff_csv.get('error','')}")

st.markdown("""
<div class='info-box'>
  <h4 class='info-title'>💡 솔루션 개요</h4>
  <p class='info-text'>
    🏠 <b>시민이 직접 해보는 탄소중립 계산기:</b> 거주 환경과 평소 에너지 사용량만 입력하면,
    친환경 히트펌프(AWHP) 전환 시 <b>얼마나 경제적 이득인지</b> 바로 확인하실 수 있습니다.<br><br>
    ⚡ <b>기후 데이터 연동:</b> 지역별 월평균 기온 기반 코사인 일주기 모델로 추정한 외기온도에
    카르노 열펌프 효율 공식을 적용하여, 우리 동네 기후에 맞춘 실제 효율(sCOP)을 반영합니다.
  </p>
</div>
""", unsafe_allow_html=True)

# ── 섹션 1: 대상지 기본 정보 ──
st.markdown('<div class="section-title">1. 대상지 기본 정보</div>', unsafe_allow_html=True)
c1, c2 = st.columns(2)
with c1: s_reg  = st.selectbox("광역 지자체", list(regions_full.keys()), index=0)
with c2: s_sub  = st.selectbox("기초 지자체", regions_full.get(s_reg, ["전체"]))
c3, c4 = st.columns(2)
with c3: h_type = st.selectbox("주거 형태", ["단독 주택 / 다가구 주택", "아파트", "연립 / 빌라 / 다세대 주택"])
with c4: h_size = st.number_input("전용 면적 (평)", min_value=10, value=30)

zone        = map_region_to_zone(s_reg)
dynamic_cop = 3.0

# ── 기후 및 sCOP 표시 ──
st.markdown('<div class="section-title">📊 우리 동네 기후 및 히트펌프 효율 분석</div>', unsafe_allow_html=True)
if df_temp and df_cop:
    dynamic_cop = df_cop[zone]["scop"]

    # 월별 낮(06~18시) / 밤(19~05시) 평균기온 계산
    # 외기온도 CSV: temp_data[zone] = 24(시간) × 12(월) 행렬
    month_labels = ["1월","2월","3월","4월","5월","6월","7월","8월","9월","10월","11월","12월"]
    day_avg, night_avg = [], []
    for mi in range(12):
        day_vals   = [df_temp[zone][h][mi] for h in range(len(df_temp[zone])) if 6 <= h <= 18]
        night_vals = [df_temp[zone][h][mi] for h in range(len(df_temp[zone])) if h < 6 or h > 18]
        day_avg.append(round(sum(day_vals)/len(day_vals), 1)   if day_vals   else 0.0)
        night_avg.append(round(sum(night_vals)/len(night_vals), 1) if night_vals else 0.0)

    df_temp_chart = pd.DataFrame({
        "월":   month_labels * 2,
        "기온":  day_avg + night_avg,
        "구분": ["낮 (06~18시)"] * 12 + ["밤 (19~05시)"] * 12,
    })
    # 월 순서 고정
    month_order = month_labels

    cl1, cl2 = st.columns([2, 1])
    with cl1:
        # 낮(주황) · 밤(진한 남색) 나란히 배치하는 그룹형 막대그래프
        # Y축: 실제 데이터 범위보다 여유 있게 설정, 5°C 간격 눈금
        all_temps = day_avg + night_avg
        y_min = int(min(all_temps)) - 3
        y_max = int(max(all_temps)) + 3

        bar_chart = alt.Chart(df_temp_chart).mark_bar().encode(
            x=alt.X("월:O", sort=month_order, title="월"),
            y=alt.Y("기온:Q", title="평균기온 (°C)",
                    scale=alt.Scale(domain=[y_min, y_max]),
                    axis=alt.Axis(tickCount=int((y_max - y_min) / 5) + 1,
                                  values=list(range(y_min - (y_min % 5), y_max + 5, 5)),
                                  format="d")),
            color=alt.Color("구분:N", scale=alt.Scale(
                domain=["낮 (06~18시)", "밤 (19~05시)"],
                range=["#f97316", "#1e3a5f"]
            ), legend=alt.Legend(orient="top", title=None)),
            xOffset="구분:N",
            tooltip=["월", "구분", alt.Tooltip("기온:Q", title="온도(°C)")]
        )
        zero_line = alt.Chart(pd.DataFrame({"y": [0]})).mark_rule(
            color="#94a3b8", strokeDash=[4, 3], strokeWidth=1
        ).encode(y="y:Q")
        chart = (zero_line + bar_chart).properties(
            height=260, title=f"{s_reg} 월별 낮·밤 평균기온"
        )
        st.altair_chart(chart, use_container_width=True)

    with cl2:
        # 겨울철 난방 COP: 11~3월 월평균 COP의 평균 (COP_계산기.csv 기반)
        monthly_cop = df_cop[zone].get("monthly_cop", [])
        heating_months_idx = [10, 11, 0, 1, 2]  # 11, 12, 1, 2, 3월
        heating_cops = [monthly_cop[i] for i in heating_months_idx
                        if i < len(monthly_cop) and monthly_cop[i] > 0]
        winter_cop = round(sum(heating_cops) / len(heating_cops), 2) if heating_cops else dynamic_cop

        st.success(f"**✅ [{s_reg}] 겨울철 난방 COP**\n# {winter_cop}")
        st.caption("11~3월 평균 COP (카르노 공식 기반)")

# ── 섹션 2: 에너지 소비 현황 ──
st.markdown('<div class="section-title">2. 에너지 소비 현황</div>', unsafe_allow_html=True)
cv1, cv2 = st.columns(2)
with cv1: w_heat = st.number_input("동절기(1월) 평균 난방비 (만원)", value=20)
with cv2: w_elec = st.number_input("동절기(1월) 전기요금 (만원)", value=6)

# 입력값이 바뀌면 이전 분석 결과 초기화 → 반드시 버튼 다시 눌러야 최신값으로 재계산
_input_key = (w_heat, w_elec, s_reg, h_type, h_size)
if st.session_state.get("_last_input_key") != _input_key:
    st.session_state.analyzed = False
    st.session_state["_last_input_key"] = _input_key

# ── 섹션 3: 시뮬레이션 변수 ──
st.markdown('<div class="section-title">3. 시뮬레이션 상수 변수</div>', unsafe_allow_html=True)
cs1, cs2 = st.columns(2)
with cs1:
    f_inf  = st.slider("가스/등유요금 인상률 (%)", 0.0, 15.0, 5.0)
    e_inf  = st.slider("전기요금 인상률 (%)", 0.0, 15.0, 3.0)
    s_capa = st.number_input("태양광 용량 (kW)", value=0.0,
                              help="테스트 모드(전기요금 1종)에서는 태양광 0kW 가구 기준으로 Saving을 계산합니다.")

with cs2:
    sub_nat  = st.checkbox("정부 보조금 적용 (560만원)", value=True)
    is_south = s_reg in ["제주도","경상남도","전라남도"]
    sub_loc  = st.checkbox("지자체 매칭 보조금 적용 (280만원)", value=is_south)
    st.caption("*2026년 현재 제주, 경남, 전남은 보조금 신청이 가능합니다.")

    st.markdown("---")
    st.markdown("**전기 요금제 선택**")
    st.markdown("""
<div class='help-text'>
고지서 금액을 바탕으로 사용량(kWh)을 역산하기 위해 적용 중인 요금제를 선택해 주세요.<br>
<b>전기요금 1종(누진제)</b>은 전기요금누진제.csv 기반 정밀 계산이 적용됩니다.
</div>""", unsafe_allow_html=True)

    if "tariff" not in st.session_state: st.session_state.tariff = "누진제_1종"

    def set_t1(): st.session_state.tariff = "누진제_1종"
    def set_t2(): st.session_state.tariff = "누진제(가전), 일반용(히트펌프)"
    def set_t3(): st.session_state.tariff = "주택용 계시별 요금제 (제주)"

    # 1종 누진제: CSV 기반 / 기존 옵션과 명확히 구분
    st.checkbox(
        "전기요금 1종 (누진제) — CSV 기반 정밀 계산 ✅",
        value=(st.session_state.tariff == "누진제_1종"),
        on_change=set_t1,
        help="전기요금누진제.csv에서 로드한 기타계절/하계 요금표를 사용합니다."
    )
    st.checkbox(
        "누진제(가전), 일반용(히트펌프)",
        value=(st.session_state.tariff == "누진제(가전), 일반용(히트펌프)"),
        on_change=set_t2,
    )
    st.checkbox(
        "주택용 계시별 요금제 (제주)",
        value=(st.session_state.tariff == "주택용 계시별 요금제 (제주)"),
        on_change=set_t3,
    )
    elec_tariff = st.session_state.tariff

if "analyzed" not in st.session_state: st.session_state.analyzed = False
if st.button("경제성 분석 실행", type="primary", use_container_width=True):
    st.session_state.analyzed = True


# ══════════════════════════════════════════════════════════
# 6. 분석 결과
# ══════════════════════════════════════════════════════════
if st.session_state.analyzed:

    # ── ① 전기 사용량 역산 ──
    # 1종 누진제: CSV 기반 calc_elec_bill_from_tariff 사용
    # 기타 요금제: 기존 하드코딩 로직 사용
    if elec_tariff == "누진제_1종":
        # 1월 기준으로 역산 (기타계절 시즌 적용)
        cur_k = reverse_kwh_from_tariff(w_elec, 1, tariff_csv)
    elif elec_tariff == "누진제(가전), 일반용(히트펌프)":
        # 히트펌프 분리 계약: 단일 단가 110원 + 부가항목
        def _legacy_bill(k):
            return round((910 + k * 110 + 14.0 * k) * 1.127 / 10000, 4)
        lo, hi = 0.0, 3000.0
        for _ in range(40):
            mid = (lo + hi) / 2
            if _legacy_bill(mid) < w_elec: lo = mid
            else: hi = mid
        cur_k = round(mid, 1)
    else:  # 제주 계시별
        def _jeju_bill(k):
            return round((4300 + k * 160 + 14.0 * k) * 1.127 / 10000, 4)
        lo, hi = 0.0, 3000.0
        for _ in range(40):
            mid = (lo + hi) / 2
            if _jeju_bill(mid) < w_elec: lo = mid
            else: hi = mid
        cur_k = round(mid, 1)

    # ── ② 보조금 및 투자비 ──
    total_sub = (560 if sub_nat else 0) + (280 if sub_loc else 0)
    capex     = calc_capex(h_type, h_size)
    net_cap   = max(0, capex - total_sub)

    # ── ③ 연간 비용 기준값 ──
    hdd_ratio     = calc_hdd_ratio(zone)
    ann_heat_base = w_heat * hdd_ratio   # 연간 기존 난방비 (만원)

    # 연간 전기요금: 1종 누진제는 월별 시즌 적용, 기타는 단순 ×12
    if elec_tariff == "누진제_1종":
        ann_elec_base = sum(
            calc_elec_bill_won_to_man(cur_k, m, tariff_csv)
            for m in range(1, 13)
        )
    else:
        if elec_tariff == "누진제(가전), 일반용(히트펌프)":
            ann_elec_base = round((910 + cur_k * 110 + 14.0 * cur_k) * 1.127 / 10000, 4) * 12
        else:
            ann_elec_base = round((4300 + cur_k * 160 + 14.0 * cur_k) * 1.127 / 10000, 4) * 12

    # ── ④ 태양광 절감 (1종 누진제는 CSV 요금 함수 사용) ──
    if elec_tariff == "누진제_1종":
        pv_annual_saving = calc_pv_saving(s_reg, s_capa, cur_k, tariff_csv)
    else:
        pv_annual_saving = 0.0  # 기타 요금제는 태양광 절감 별도 미계산

    # ── ⑤ HP 연간 순 운영비 ──
    ann_hp_net_op = max(0.0, (ann_heat_base / dynamic_cop) - pv_annual_saving)

    # ── ⑥ [핵심 신기능] 도시가스 콘덴싱→HP Saving (CSV 기반) ──
    # 테스트 조건: 전기요금 1종 누진제 + 태양광 0kW
    # CSV의 기준 가구와 입력 가구의 난방 규모 차이를 scale로 보정
    # scale = 입력한 1월 난방비 / CSV 기준 1월 난방비 (CSV 1월=56,580원 ≈ 5.66만원)
    CSV_BASE_JAN_MAN = round(
        tariff_csv["saving"]["condensing"]["hp_monthly_billing_won"][0] / 10000, 2
    )
    # 기준 가구의 기존 1월 난방비: 기존난방비연합계 × (1월HDD/연간HDD)
    existing_ann_man  = tariff_csv["saving"]["condensing"]["existing_annual_won"] / 10000
    hdd_data          = hdd_monthly["중부2"]  # CSV는 서울(중부2) 기준
    csv_jan_heat_man  = round(existing_ann_man * hdd_data[0] / sum(hdd_data), 2)

    # scale: 사용자 1월 난방비 / CSV 기준 1월 난방비
    scale = (w_heat / csv_jan_heat_man) if csv_jan_heat_man > 0 else 1.0

    saving_result = calc_condensing_saving(tariff_csv, scale=scale)

    # ── ⑦ 15년 복리 시뮬레이션 ──
    years, gas_cum, hp_cum, net_p = list(range(1, 16)), [], [], []
    g_s, h_s, pb = 0.0, float(net_cap), "15년 초과"
    for y in years:
        cg = ann_heat_base * ((1 + f_inf / 100) ** y) + ann_elec_base
        ch = ann_hp_net_op  * ((1 + e_inf / 100) ** y) + ann_elec_base
        g_s += cg; h_s += ch
        p = int(g_s - h_s)
        gas_cum.append(int(g_s)); hp_cum.append(int(h_s)); net_p.append(p)
        if pb == "15년 초과" and p > 0: pb = f"{y}년차"

    # ══════════════════════════════════════════════════════════
    # 결과 출력
    # ══════════════════════════════════════════════════════════
    st.markdown('<div class="section-title">분석 결과 요약</div>', unsafe_allow_html=True)

    ca, cb, cc, cd = st.columns(4)
    ca.metric("투자 회수 시점", pb)
    cb.metric("15년 순이익", f"{net_p[-1]:,} 만원")
    cc.metric("적용 sCOP", f"{dynamic_cop}")
    cd.metric("HDD 난방 계수", f"×{hdd_ratio} ({zone})")

    # ── [테스트 섹션] 도시가스 콘덴싱→HP 난방비 Saving ──
    if elec_tariff == "누진제_1종":
        st.markdown('<div class="section-title">🧪 [테스트] 전기요금 1종 — 도시가스 콘덴싱→HP 난방비 Saving</div>',
                    unsafe_allow_html=True)

        src_label = "📄 전기요금누진제.csv" if saving_result["source"] == "csv" else "⚠️ fallback (CSV 미로드)"
        st.caption(f"데이터 출처: {src_label}  |  규모 보정 계수(scale): {round(scale, 2)}배")

        # 핵심 지표 3개
        s1, s2, s3 = st.columns(3)
        s1.metric(
            "HP 연간 전기요금",
            f"{saving_result['hp_annual_man']:,.1f} 만원",
            help="전기요금누진제.csv 월별 청구요금 합산 (규모 보정 적용)"
        )
        s2.metric(
            "기존 연간 난방비 (도시가스 콘덴싱)",
            f"{saving_result['existing_annual_man']:,.1f} 만원",
            help="CSV 기준 가구 기존 난방비 × 규모 보정"
        )
        s3.metric(
            "연간 난방비 Saving",
            f"{saving_result['saving_man']:,.1f} 만원",
            delta=f"{round(saving_result['saving_ratio'] * 100)}% 절감",
            help="기존 난방비 - HP 전기요금"
        )

        # 월별 Saving 차트
        months = list(range(1, 13))
        # 기존 월별 난방비: 사용자 입력 w_heat(1월 기준)를 HDD 비례로 직접 배분
        # 수식: 해당월 난방비 = w_heat × (해당월 HDD / 1월 HDD)
        # → 1월은 항상 w_heat 그대로, 다른 달은 HDD 비율만큼 감소
        # 이전 버그: existing_annual_man(CSV scale값)에서 배분 → 1월이 w_heat와 달라짐
        hdd_base = hdd_monthly[zone]   # 사용자 선택 지역 기준
        hdd_jan  = hdd_base[0]
        monthly_ex = [
            round(w_heat * hdd_base[m-1] / hdd_jan, 2) if hdd_jan > 0 else 0
            for m in months
        ]

        df_saving = pd.DataFrame({
            "월":            [f"{m}월" for m in months],
            "기존 난방비(만원)": monthly_ex,
            "HP 전기요금(만원)": saving_result["monthly_hp_man"],
        })
        df_melt = df_saving.melt("월", var_name="구분", value_name="금액(만원)")

        chart_s = alt.Chart(df_melt).mark_bar().encode(
            x=alt.X("월:O", sort=[f"{m}월" for m in months]),
            y=alt.Y("금액(만원):Q"),
            color=alt.Color("구분:N", scale=alt.Scale(
                domain=["기존 난방비(만원)", "HP 전기요금(만원)"],
                range=["#f87171", "#60a5fa"]
            )),
            xOffset="구분:N",
            tooltip=["월", "구분", "금액(만원)"],
        ).properties(height=260, title="월별 기존 난방비 vs HP 전기요금 비교")
        st.altair_chart(chart_s, use_container_width=True)

        # 월별 상세 테이블
        with st.expander("📋 월별 상세 데이터 보기"):
            df_detail = pd.DataFrame({
                "월":              [f"{m}월" for m in months],
                "기존 난방비(만원)": monthly_ex,
                "HP 청구요금(만원)": saving_result["monthly_hp_man"],
                "월별 절감액(만원)": [round(monthly_ex[m-1] - saving_result["monthly_hp_man"][m-1], 2)
                                      for m in months],
                "HP 청구요금(원)":   saving_result["monthly_hp_won"],
                "시즌":            [get_season(m) for m in months],
            })
            st.dataframe(df_detail, use_container_width=True, hide_index=True)

        st.markdown(f"""
<div class='warn-box'>
⚠️ <b>테스트 모드 주의사항</b><br>
• 이 Saving은 <b>태양광 0kW, 전기요금 1종(누진제)</b> 가구 기준입니다.<br>
• CSV의 기준 가구(도시가스 콘덴싱)와 입력 가구의 난방 규모 비율(scale={round(scale,2)})을 적용했습니다.<br>
• CSV 기준: 중부2(서울) 기후, 기타계절/하계 누진제, 슈퍼유저(1,000kWh 초과) 구간 미반영.<br>
• 실제 절감액은 가구별 사용 패턴, 단열 성능, 설정 온도에 따라 크게 달라질 수 있습니다.
</div>
""", unsafe_allow_html=True)

    else:
        st.info("ℹ️ 도시가스 콘덴싱→HP Saving 테스트는 **전기요금 1종(누진제)** 선택 시 표시됩니다.")

    # ── 가정값 상세 ──
    with st.expander("📋 적용된 계산 가정값 및 출처 상세 보기"):
        st.markdown(f"""
| 항목 | 적용값 | 근거 |
|------|--------|------|
| 설비 CAPEX | **{capex}만원** | 에너지경제연구원 2025 / 정부 브리핑 |
| HDD 난방 계수 | **×{hdd_ratio}** | COP_계산기.csv HDD ({zone}) |
| 태양광 절감액 | **연 {pv_annual_saving:.1f}만원** | pv_monthly_data × {s_capa}kW |
| HP 연간 운영비 | **{ann_hp_net_op:.1f}만원** | 난방비 ÷ sCOP({dynamic_cop}) - PV절감 |
| 전기 역산 kWh | **{cur_k}kWh/월** | {"CSV 기반 기타계절 누진제 역산" if elec_tariff == "누진제_1종" else "하드코딩 요금 역산"} |
| 요금 데이터 출처 | **{tariff_csv["source"]}** | 전기요금누진제.csv |
        """)

    # ── 15년 차트 ──
    g1, g2 = st.columns(2)
    with g1:
        st.write("**15년 누적 비용 흐름**")
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

    # ── 엑셀 다운로드 ──
    wb  = Workbook()
    hf  = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    sf  = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    pf  = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
    gf  = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")  # Saving 시트용
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
    ws1["A1"] = f"히트펌프 경제성 분석 마스터 ({s_reg})"
    ws1["A1"].fill = hf; ws1["A1"].font = fw; ws1["A1"].alignment = center
    rows1 = [
        ("항목",           "값",                       "단위",    "산출 근거 및 출처"),
        ("1월 난방비",      w_heat,                     "만원",    "사용자 입력"),
        ("1월 전기요금",    w_elec,                     "만원",    "사용자 입력"),
        ("역산 kWh",        cur_k,                      "kWh/월",  f"{'CSV 기반 기타계절 역산' if elec_tariff=='누진제_1종' else '하드코딩 역산'}"),
        ("지역 sCOP",       dynamic_cop,                "-",       f"카르노 공식+기후 데이터 ({zone})"),
        ("HDD 난방 계수",   hdd_ratio,                  "-",       f"COP_계산기.csv HDD ({zone})"),
        ("설비 CAPEX",      capex,                      "만원",    "에너지경제연구원2025/정부브리핑"),
        ("정부보조금",      560,                        "만원",    "기후에너지환경부 2026 보급 사업"),
        ("지방보조금",      280,                        "만원",    "정부보조금 50% 매칭 (제주·경남·전남)"),
        ("순 투자비",       net_cap,                    "만원",    "=CAPEX-정부보조금-지방보조금"),
        ("태양광 절감액",   round(pv_annual_saving,1),  "만원/년", f"pv_monthly_data×{s_capa}kW"),
        ("HP 연간 운영비",  round(ann_hp_net_op,1),     "만원",    "=연간난방비÷sCOP-태양광절감"),
    ]
    for ri, rdata in enumerate(rows1, 3):
        for ci, val in enumerate(rdata, 1):
            c = ws1.cell(row=ri, column=ci, value=val); c.border = thin
            if ri == 3: c.fill = sf; c.font = fb
            elif ci == 2 and ri != 3: c.font = fi; c.alignment = right
    ws1.column_dimensions["A"].width = 20; ws1.column_dimensions["D"].width = 45

    # ② [신규] 도시가스 콘덴싱→HP Saving 시트 (전기요금 1종 선택 시)
    if elec_tariff == "누진제_1종":
        ws_s = wb.create_sheet("②콘덴싱→HP_Saving")
        ws_s.merge_cells("A1:G1")
        ws_s["A1"] = f"도시가스 콘덴싱→HP 난방비 Saving 분석 [전기요금 1종, 태양광 0kW, scale={round(scale,2)}]"
        ws_s["A1"].fill = hf; ws_s["A1"].font = fw; ws_s["A1"].alignment = center

        for ci, h in enumerate(["월","시즌","기존 난방비(만원)","HP 청구요금(만원)","HP 청구요금(원)","월 절감액(만원)","비고"], 1):
            c = ws_s.cell(row=2, column=ci, value=h)
            c.fill = sf; c.font = fb; c.border = thin; c.alignment = center

        hdd_base = hdd_monthly[zone]   # 사용자 지역 기준 (버그수정: 중부2 고정 → zone 사용)
        hdd_jan_xl = hdd_base[0]
        for m in range(1, 13):
            r   = m + 2
            # w_heat(1월 입력값) 기준으로 HDD 비례 배분 (버그수정: CSV scale값 → w_heat 직접 사용)
            ex  = round(w_heat * hdd_base[m-1] / hdd_jan_xl, 2) if hdd_jan_xl > 0 else 0
            hp  = saving_result["monthly_hp_man"][m-1]
            won = saving_result["monthly_hp_won"][m-1]
            sav = round(ex - hp, 2)
            ssn = get_season(m)
            note = "난방월" if hdd_base[m-1] > 0 else "비난방월"
            for ci, val in enumerate([f"{m}월", ssn, ex, hp, won, sav, note], 1):
                c = ws_s.cell(row=r, column=ci, value=val); c.border = thin
            ws_s.cell(row=r, column=6).font = fg  # 절감액 초록색
            if m % 2 == 0:
                for ci in range(1, 8): ws_s.cell(row=r, column=ci).fill = gf

        # 합계 행
        r_sum = 15
        ws_s.cell(row=r_sum, column=1, value="연간 합계").font = fb
        ws_s.cell(row=r_sum, column=3, value=saving_result["existing_annual_man"]).font = fb
        ws_s.cell(row=r_sum, column=4, value=saving_result["hp_annual_man"]).font = fb
        ws_s.cell(row=r_sum, column=5, value=saving_result["hp_annual_total_won"]).font = fb
        ws_s.cell(row=r_sum, column=6, value=saving_result["saving_man"]).font = fg
        ws_s.cell(row=r_sum, column=7, value=f"Saving {round(saving_result['saving_ratio']*100)}%").font = fg
        for ci in range(1, 8): ws_s.cell(row=r_sum, column=ci).border = thin
        for col in "ABCDEFG": ws_s.column_dimensions[col].width = 20

    # ③ 15년 재무 분석 시트
    sheet_num = "③" if elec_tariff == "누진제_1종" else "②"
    ws3 = wb.create_sheet(f"{sheet_num}15년_재무_분석")
    ws3.merge_cells("A1:H1")
    ws3["A1"] = "15년 장기 투자 회수 및 누적 순이익 시뮬레이션"
    ws3["A1"].fill = hf; ws3["A1"].font = fw; ws3["A1"].alignment = center
    for ci, h in enumerate(["연도","물가지수(4%)","기존 OPEX(만)","HP OPEX(만)","연간 순이익(만)","누적 순이익(만)","ROI","상태"], 1):
        c = ws3.cell(row=2, column=ci, value=h)
        c.fill = sf; c.font = fb; c.border = thin; c.alignment = center
    ref_cap = f"'①입력_가정'!$B$12"
    for y in range(1, 16):
        r = y + 2
        ws3.cell(row=r, column=1, value=f"{y}년차").border = thin
        ws3.cell(row=r, column=2, value=f"=(1+0.04)^{y-1}").border = thin
        ws3.cell(row=r, column=3, value=round(ann_heat_base, 1)).border = thin
        ws3.cell(row=r, column=4, value=round(ann_hp_net_op, 1)).border = thin
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
        file_name=f"Expert_Report_{s_reg}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )