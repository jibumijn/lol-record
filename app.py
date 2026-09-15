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

        df = pd.read_excel(
            file_source,
            sheet_name=0
        )

        # ----------------------------------------------------
        # A~K열
        #
        # A 날짜
        # B 세트
        # C 이름
        # D 매치결과
        # E 세트결과
        # F 라인
        # G 챔피언
        # H 킬
        # I 데스
        # J 어시스트
        # K 딜량
        # L 골드
        #
        # ※ 실제 파일에서 A~L까지 존재하므로 12개 컬럼 사용
        # ----------------------------------------------------

        cols = [
            "날짜",
            "세트",
            "이름",
            "매치결과",
            "세트결과",
            "라인",
            "챔피언",
            "킬",
            "데스",
            "어시스트",
            "딜량",
            "골드",
        ]

        # 최소 12개 컬럼이 존재하는 경우
        if len(df.columns) >= 12:

            df = df.iloc[:, :12]
            df.columns = cols

        else:

            df.columns = cols[:len(df.columns)]


        # ----------------------------------------------------
        # 이름이 빈 행 제거
        # ----------------------------------------------------

        df = df.dropna(
            subset=["이름"]
        )

        df["이름"] = (
            df["이름"]
            .astype(str)
            .str.strip()
        )

        df = df[
            df["이름"] != ""
        ]


        # ----------------------------------------------------
        # 텍스트 컬럼 공백 제거
        # ----------------------------------------------------

        for col in [
            "매치결과",
            "세트결과",
            "라인",
            "챔피언"
        ]:

            if col in df.columns:

                df[col] = (
                    df[col]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )


        # ----------------------------------------------------
        # 결과값 통일
        # ----------------------------------------------------

        if "매치결과" in df.columns:

            df["매치결과"] = df["매치결과"].replace({

                "Win": "승",
                "win": "승",
                "W": "승",
                "1": "승",

                "Lose": "패",
                "lose": "패",
                "loss": "패",
                "L": "패",
                "0": "패",

                "Draw": "무승부",
                "draw": "무승부",
                "D": "무승부",
                "무": "무승부",
            })


        if "세트결과" in df.columns:

            df["세트결과"] = df["세트결과"].replace({

                "Win": "승",
                "win": "승",
                "W": "승",
                "1": "승",

                "Lose": "패",
                "lose": "패",
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


        # ----------------------------------------------------
        # ★ 매치결과 보완용 컬럼 생성
        #
        # D열 매치결과는 첫 번째 세트에만 존재하므로
        # 같은 날짜 + 같은 이름의 모든 세트에
        # 해당 매치 결과를 연결한다.
        #
        # 예:
        #
        # 2025-12-21 / 임성빈
        #
        # 1세트 → 매치결과 패
        # 2세트 → 빈칸
        # 3세트 → 빈칸
        #
        # ↓
        #
        # 매치결과_전체
        #
        # 1세트 → 패
        # 2세트 → 패
        # 3세트 → 패
        # ----------------------------------------------------

        match_result_map = (
            df[
                df["매치결과"] != ""
            ]
            .groupby(
                ["날짜", "이름"]
            )["매치결과"]
            .first()
            .to_dict()
        )


        df["매치결과_전체"] = [
            match_result_map.get(
                (row["날짜"], row["이름"]),
                ""
            )
            for _, row in df.iterrows()
        ]


        return df


    except Exception as e:

        st.error(
            f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}"
        )

        return None


# ============================================================
# 3. 사이드바 - 파일 업로드
# ============================================================

st.sidebar.header("📁 파일 및 필터")


uploaded_file = st.sidebar.file_uploader(
    "엑셀 파일(.xlsx) 선택",
    type=["xlsx", "xlsm"]
)


df = None


if uploaded_file is not None:

    df = load_data(
        uploaded_file
    )


elif os.path.exists(DEFAULT_FILE):

    df = load_data(
        DEFAULT_FILE
    )

    st.sidebar.success(
        f"기본 파일(`{DEFAULT_FILE}`)을 불러왔습니다."
    )


else:

    st.sidebar.info(
        "엑셀 파일을 업로드하거나, "
        "폴더 내에 '롤내전.xlsx'를 배치하세요."
    )


# ============================================================
# 4. 데이터가 로드된 경우
# ============================================================

if df is not None and not df.empty:

    st.sidebar.subheader(
        "🔍 검색 필터"
    )


    # ========================================================
    # 이름 필터
    # ========================================================

    people = [
        "전체"
    ] + sorted(
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


    # 선택한 사람의 데이터
    person_subset = (

        df

        if selected_person == "전체"

        else df[
            df["이름"] == selected_person
        ]
    )


    # ========================================================
    # 라인 필터
    # ========================================================

    lines = [
        "전체"
    ] + sorted(
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


    # ========================================================
    # 챔피언 필터
    # ========================================================

    champions = [
        "전체"
    ] + sorted(
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


    # ========================================================
    # 최종 필터링
    # ========================================================

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
    # 5. 매치 전적 계산
    # ========================================================

    # --------------------------------------------------------
    # 매치 단위로 중복 제거
    #
    # 한 날짜 = 한 매치라는 현재 엑셀 구조를 기준으로 함.
    #
    # 같은 날짜에 3세트가 있다면
    # → 매치 1개
    #
    # 같은 날짜에 5세트가 있다면
    # → 매치 1개
    # --------------------------------------------------------

    match_df = (
        filtered_df[
            filtered_df["매치결과_전체"] != ""
        ]
        .drop_duplicates(
            subset=[
                "날짜",
                "이름"
            ]
        )
        .copy()
    )


    # 매치 승
    match_wins = sum(
        match_df["매치결과_전체"].isin([
            "승",
            "W",
            "Win",
            "win",
            "1"
        ])
    )


    # 매치 패
    match_losses = sum(
        match_df["매치결과_전체"].isin([
            "패",
            "L",
            "Lose",
            "lose",
            "loss",
            "0"
        ])
    )


    # 매치 무승부
    match_draws = sum(
        match_df["매치결과_전체"].isin([
            "무승부",
            "무",
            "D",
            "Draw",
            "draw"
        ])
    )


    # 매치 수
    match_games = len(
        match_df
    )


    # --------------------------------------------------------
    # 매치 승률
    #
    # 무승부는 분모에서 제외
    #
    # 승 / (승 + 패)
    # --------------------------------------------------------

    match_completed = (
        match_wins
        + match_losses
    )


    if match_completed > 0:

        match_winrate = (
            match_wins
            / match_completed
            * 100
        )

    else:

        match_winrate = 0.0


    # ========================================================
    # 6. 세트 전적 계산
    # ========================================================

    set_df = filtered_df[
        filtered_df["세트결과"] != ""
    ].copy()


    # 세트 승
    set_wins = sum(
        set_df["세트결과"].isin([
            "승",
            "W",
            "Win",
            "win",
            "1"
        ])
    )


    # 세트 패
    set_losses = sum(
        set_df["세트결과"].isin([
            "패",
            "L",
            "Lose",
            "lose",
            "loss",
            "0"
        ])
    )


    # 세트 무승부
    set_draws = sum(
        set_df["세트결과"].isin([
            "무승부",
            "무",
            "D",
            "Draw",
            "draw"
        ])
    )


    # 세트 수
    set_games = len(
        set_df
    )


    # --------------------------------------------------------
    # 세트 승률
    #
    # 무승부는 분모에서 제외
    # --------------------------------------------------------

    set_completed = (
        set_wins
        + set_losses
    )


    if set_completed > 0:

        set_winrate = (
            set_wins
            / set_completed
            * 100
        )

    else:

        set_winrate = 0.0


    # ========================================================
    # 7. KDA 계산
    #
    # KDA는 세트 기준으로 계산
    # ========================================================

    games = len(
        filtered_df
    )


    kills = filtered_df["킬"].sum()
    deaths = filtered_df["데스"].sum()
    assists = filtered_df["어시스트"].sum()


    if deaths > 0:

        kda = (
            (kills + assists)
            / deaths
        )

    else:

        kda = (
            (kills + assists)
            if games > 0
            else 0.0
        )


    # ========================================================
    # 8. 평균 딜량 / 골드
    # ========================================================

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
    # 9. 매치 전적 Dashboard
    # ========================================================

    st.subheader(
        "🏆 매치 전적"
    )


    m1, m2, m3, m4 = st.columns(4)


    m1.metric(
        "매치 수",
        f"{match_games:,}매치"
    )


    m2.metric(
        "매치 승 / 패 / 무",
        f"{match_wins}승 {match_losses}패 {match_draws}무"
    )


    m3.metric(
        "매치 승률",
        f"{match_winrate:.1f}%"
    )


    m4.metric(
        "매치 전적",
        f"{match_wins}승 {match_losses}패"
    )


    st.divider()


    # ========================================================
    # 10. 세트 전적 Dashboard
    # ========================================================

    st.subheader(
        "🎮 세트 전적"
    )


    s1, s2, s3, s4 = st.columns(4)


    s1.metric(
        "세트 수",
        f"{set_games:,}세트"
    )


    s2.metric(
        "세트 승 / 패 / 무",
        f"{set_wins}승 {set_losses}패 {set_draws}무"
    )


    s3.metric(
        "세트 승률",
        f"{set_winrate:.1f}%"
    )


    s4.metric(
        "세트 전적",
        f"{set_wins}승 {set_losses}패"
    )


    st.divider()


    # ========================================================
    # 11. 개인 플레이 통계
    # ========================================================

    st.subheader(
        "📊 개인 플레이 통계"
    )


    p1, p2, p3, p4, p5 = st.columns(5)


    p1.metric(
        "K / D / A",
        f"{int(kills)} / {int(deaths)} / {int(assists)}"
    )


    p2.metric(
        "평균 KDA",
        f"{kda:.2f}:1"
    )


    p3.metric(
        "평균 딜량",
        f"{avg_damage:,.0f}"
    )


    p4.metric(
        "평균 골드",
        f"{avg_gold:,.0f}"
    )


    p5.metric(
        "플레이 세트",
        f"{games:,}세트"
    )


    st.divider()


    # ========================================================
    # 12. 상세 경기 기록
    # ========================================================

    st.subheader(
        "📜 상세 경기 기록"
    )


    # 최신 경기부터 표시
    display_df = (
        filtered_df
        .iloc[::-1]
        .copy()
    )


    # --------------------------------------------------------
    # 표시용 매치결과
    #
    # 여기서는 원본 D열과 동일하게
    # 첫 번째 세트에만 매치결과가 표시되도록 함.
    #
    # 따라서:
    #
    # 1세트 → 매치결과 패 / 세트결과 승
    # 2세트 → 빈칸     / 세트결과 패
    # 3세트 → 빈칸     / 세트결과 패
    #
    # 형태로 보여줌.
    # --------------------------------------------------------

    # 원본 D열의 값을 그대로 사용
    display_df["매치결과"] = display_df["매치결과"]


    # --------------------------------------------------------
    # 컬럼 순서
    # --------------------------------------------------------

    display_columns = [
        "날짜",
        "세트",
        "이름",
        "매치결과",
        "세트결과",
        "라인",
        "챔피언",
        "킬",
        "데스",
        "어시스트",
        "딜량",
        "골드"
    ]


    # 실제 존재하는 컬럼만 선택
    display_columns = [
        col
        for col in display_columns
        if col in display_df.columns
    ]


    display_df = display_df[
        display_columns
    ]


    # K/D/A 정수 표시
    for col in [
        "킬",
        "데스",
        "어시스트"
    ]:

        if col in display_df.columns:

            display_df[col] = (
                display_df[col]
                .astype(int)
            )


    # --------------------------------------------------------
    # 데이터 테이블
    # --------------------------------------------------------

    st.dataframe(
        display_df,
        width="stretch",
        hide_index=True,

        column_config={

            "날짜": st.column_config.TextColumn(
                "날짜"
            ),

            "세트": st.column_config.NumberColumn(
                "세트",
                format="%d"
            ),

            "이름": st.column_config.TextColumn(
                "이름"
            ),

            "매치결과": st.column_config.TextColumn(
                "매치결과"
            ),

            "세트결과": st.column_config.TextColumn(
                "세트결과"
            ),

            "킬": st.column_config.NumberColumn(
                "킬",
                format="%d"
            ),

            "데스": st.column_config.NumberColumn(
                "데스",
                format="%d"
            ),

            "어시스트": st.column_config.NumberColumn(
                "어시스트",
                format="%d"
            ),

            "딜량": st.column_config.NumberColumn(
                "딜량",
                format="%d"
            ),

            "골드": st.column_config.NumberColumn(
                "골드",
                format="%d"
            ),
        },
    )
