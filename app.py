import os
from pathlib import Path
import pandas as pd
import streamlit as st


# ============================================================
# 1. 페이지 기본 설정
# ============================================================

st.set_page_config(
    page_title="롤 내전 전적 프로그램",
    page_icon="⚔️",
    layout="wide"
)

st.title("⚔️ 롤 내전 전적 프로그램")
st.caption("웹 브라우저에서 편리하게 전적을 조회하고 분석하세요.")

DEFAULT_FILE = "롤내전.xlsx"


# ============================================================
# 2. 엑셀 데이터 로드 함수
# ============================================================

@st.cache_data
def load_data(file_source):
    try:
        df = pd.read_excel(file_source, sheet_name=0)

        # A~K열
        # 날짜, 세트, 이름, 결과, 라인, 챔피언,
        # 킬, 데스, 어시스트, 딜량, 골드
        cols = [
            "날짜",
            "세트",
            "이름",
            "결과",
            "라인",
            "챔피언",
            "킬",
            "데스",
            "어시스트",
            "딜량",
            "골드",
        ]

        # 실제 엑셀 열 개수에 따라 처리
        if len(df.columns) >= 11:
            df = df.iloc[:, :11]
            df.columns = cols
        else:
            df.columns = cols[:len(df.columns)]

        # ----------------------------------------------------
        # 이름이 빈 행 제거
        # ----------------------------------------------------
        df = df.dropna(subset=["이름"])

        df["이름"] = (
            df["이름"]
            .astype(str)
            .str.strip()
        )

        df = df[df["이름"] != ""]

        # ----------------------------------------------------
        # 텍스트 컬럼 공백 제거
        # ----------------------------------------------------
        for col in ["결과", "라인", "챔피언"]:
            if col in df.columns:
                df[col] = (
                    df[col]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

        # ----------------------------------------------------
        # 결과값 정리
        # ----------------------------------------------------
        if "결과" in df.columns:
            df["결과"] = df["결과"].replace({
                "Win": "승",
                "win": "승",
                "W": "승",
                "1": "승",

                "Lose": "패",
                "loss": "패",
                "L": "패",
                "0": "패",

                "Draw": "무승부",
                "draw": "무승부",
                "D": "무승부",
                "무": "무승부",
            })

        # ----------------------------------------------------
        # 수치형 변환
        # ----------------------------------------------------
        num_cols = [
            "킬",
            "데스",
            "어시스트",
            "딜량",
            "골드"
        ]

        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(
                    df[col],
                    errors="coerce"
                ).fillna(0)

        # ----------------------------------------------------
        # 날짜 포맷팅
        # ----------------------------------------------------
        if "날짜" in df.columns:
            df["날짜"] = pd.to_datetime(
                df["날짜"],
                errors="coerce"
            ).dt.strftime("%Y-%m-%d")

        return df

    except Exception as e:
        st.error(
            f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}"
        )
        return None


# ============================================================
# 3. 사이드바 - 파일 업로드 및 필터 설정
# ============================================================

st.sidebar.header("📁 파일 및 필터")

uploaded_file = st.sidebar.file_uploader(
    "엑셀 파일(.xlsx) 선택",
    type=["xlsx", "xlsm"]
)

df = None

if uploaded_file is not None:

    df = load_data(uploaded_file)

elif os.path.exists(DEFAULT_FILE):

    df = load_data(DEFAULT_FILE)

    st.sidebar.success(
        f"기본 파일(`{DEFAULT_FILE}`)을 불러왔습니다."
    )

else:

    st.sidebar.info(
        "엑셀 파일을 업로드하거나, "
        "폴더 내에 '롤내전.xlsx'를 배치하세요."
    )


# ============================================================
# 4. 데이터가 로드된 경우 UI 구성
# ============================================================

if df is not None and not df.empty:

    st.sidebar.subheader("🔍 검색 필터")

    # --------------------------------------------------------
    # [필터 1] 이름
    # --------------------------------------------------------

    people = ["전체"] + sorted(
        [
            p
            for p in df["이름"].unique()
            if p
        ]
    )

    selected_person = st.sidebar.selectbox(
        "이름",
        people
    )

    # 이름 선택 시
    # 해당 소환사가 플레이한 라인/챔피언만 표시
    person_subset = (
        df
        if selected_person == "전체"
        else df[df["이름"] == selected_person]
    )

    # --------------------------------------------------------
    # [필터 2] 라인
    # --------------------------------------------------------

    lines = ["전체"] + sorted(
        [
            l
            for l in person_subset["라인"].unique()
            if l
        ]
    )

    selected_line = st.sidebar.selectbox(
        "라인",
        lines
    )

    # --------------------------------------------------------
    # [필터 3] 챔피언
    # --------------------------------------------------------

    champions = ["전체"] + sorted(
        [
            c
            for c in person_subset["챔피언"].unique()
            if c
        ]
    )

    selected_champion = st.sidebar.selectbox(
        "챔피언",
        champions
    )

    # --------------------------------------------------------
    # 최종 필터링 데이터
    # --------------------------------------------------------

    filtered_df = person_subset.copy()

    if selected_line != "전체":
        filtered_df = filtered_df[
            filtered_df["라인"] == selected_line
        ]

    if selected_champion != "전체":
        filtered_df = filtered_df[
            filtered_df["챔피언"] == selected_champion
        ]


    # ========================================================
    # 5. 전적 요약 Dashboard
    # ========================================================

    st.subheader("📊 전적 요약")

    # 전체 게임 수
    games = len(filtered_df)

    # --------------------------------------------------------
    # 승 / 패 / 무승부 계산
    # --------------------------------------------------------

    wins = sum(
        filtered_df["결과"].isin([
            "승",
            "W",
            "Win",
            "win",
            "1"
        ])
    )

    losses = sum(
        filtered_df["결과"].isin([
            "패",
            "L",
            "Lose",
            "loss",
            "0"
        ])
    )

    draws = sum(
        filtered_df["결과"].isin([
            "무승부",
            "무",
            "D",
            "Draw",
            "draw"
        ])
    )

    # --------------------------------------------------------
    # 승률 계산
    #
    # 무승부는 분모에서 제외
    #
    # 승률 = 승 / (승 + 패) × 100
    # --------------------------------------------------------

    completed_games = wins + losses

    if completed_games > 0:
        winrate = wins / completed_games * 100
    else:
        winrate = 0.0


    # ========================================================
    # 6. KDA 및 평균 수치
    # ========================================================

    kills = filtered_df["킬"].sum()
    deaths = filtered_df["데스"].sum()
    assists = filtered_df["어시스트"].sum()

    if deaths > 0:

        kda = (
            (kills + assists) / deaths
        )

    else:

        kda = (
            (kills + assists)
            if games > 0
            else 0.0
        )

    avg_damage = (
        filtered_df["딜량"].mean()
        if games > 0
        else 0.0
    )

    avg_gold = (
        filtered_df["골드"].mean()
        if games > 0
        else 0.0
    )


    # ========================================================
    # 7. 전적 요약 표시
    # ========================================================

    m1, m2, m3, m4, m5 = st.columns(5)

    m1.metric(
        "총 게임",
        f"{games:,}판"
    )

    m2.metric(
        "승 / 패 / 무",
        f"{wins}승 {losses}패 {draws}무"
    )

    m3.metric(
        "승률",
        f"{winrate:.1f}%"
    )

    m4.metric(
        "K / D / A",
        f"{int(kills)} / {int(deaths)} / {int(assists)}"
    )

    m5.metric(
        "평균 KDA",
        f"{kda:.2f}:1"
    )


    # --------------------------------------------------------
    # 평균 딜량 / 평균 골드
    # --------------------------------------------------------

    m6, m7, _, _, _ = st.columns(5)

    m6.metric(
        "평균 딜량",
        f"{avg_damage:,.0f}"
    )

    m7.metric(
        "평균 골드",
        f"{avg_gold:,.0f}"
    )


    st.divider()


    # ========================================================
    # 8. 상세 경기 기록
    # ========================================================

    st.subheader("📜 상세 경기 기록")

    # 최신 경기부터 표시
    display_df = filtered_df.iloc[::-1].copy()

    # K/D/A 정수 표시
    for col in ["킬", "데스", "어시스트"]:

        if col in display_df.columns:
            display_df[col] = (
                display_df[col]
                .astype(int)
            )


    # --------------------------------------------------------
    # 결과값을 보기 좋게 표시
    # --------------------------------------------------------

    st.dataframe(
        display_df,
        width="stretch",
        hide_index=True,

        column_config={

            "딜량": st.column_config.NumberColumn(
                format="%d"
            ),

            "골드": st.column_config.NumberColumn(
                format="%d"
            ),
        },
    )
