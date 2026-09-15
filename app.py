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
        # 컬럼 자동 맞춤 및 열 밀림 방지
        # 엑셀의 첫 번째 열이 Index(Unnamed)인 경우 제외 처리
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

        # 'Unnamed' 또는 인덱스 열이 A열에 있는 경우 제거
        first_col_name = str(df.columns[0]).strip()
        if "Unnamed" in first_col_name or first_col_name.isdigit() or first_col_name == "":
            df = df.iloc[:, 1:]

        # 컬럼 매핑
        if len(df.columns) >= 12:
            df = df.iloc[:, :12]
            df.columns = cols
        else:
            df.columns = cols[:len(df.columns)]


        # ----------------------------------------------------
        # 이름이 빈 행 제거
        # ----------------------------------------------------
        df = df.dropna(subset=["이름"])
        df["이름"] = df["이름"].astype(str).str.strip()
        df = df[df["이름"] != ""]


        # ----------------------------------------------------
        # 텍스트 컬럼 공백 제거 및 문자열 변환
        # ----------------------------------------------------
        for col in ["매치결과", "세트결과", "라인", "챔피언"]:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.strip()


        # ----------------------------------------------------
        # 결과값 통일 (승/패/무)
        # ----------------------------------------------------
        result_mapping = {
            "Win": "승", "win": "승", "W": "승", "1": "승", "1.0": "승",
            "Lose": "패", "lose": "패", "loss": "패", "L": "패", "0": "패", "0.0": "패",
            "Draw": "무승부", "draw": "무승부", "D": "무승부", "무": "무승부"
        }

        if "매치결과" in df.columns:
            df["매치결과"] = df["매치결과"].replace(result_mapping)

        if "세트결과" in df.columns:
            df["세트결과"] = df["세트결과"].replace(result_mapping)


        # ----------------------------------------------------
        # 수치형 변환
        # ----------------------------------------------------
        num_cols = ["세트", "킬", "데스", "어시스트", "딜량", "골드"]
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)


        # ----------------------------------------------------
        # 날짜 포맷팅
        # ----------------------------------------------------
        if "날짜" in df.columns:
            df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce").dt.strftime("%Y-%m-%d")


        # ----------------------------------------------------
        # 매치결과 보완용 컬럼 생성
        # ----------------------------------------------------
        match_result_map = (
            df[df["매치결과"] != ""]
            .groupby(["날짜", "이름"])["매치결과"]
            .first()
            .to_dict()
        )

        df["매치결과_전체"] = [
            match_result_map.get((row["날짜"], row["이름"]), "")
            for _, row in df.iterrows()
        ]

        return df

    except Exception as e:
        st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
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
    df = load_data(uploaded_file)
elif os.path.exists(DEFAULT_FILE):
    df = load_data(DEFAULT_FILE)
    st.sidebar.success(f"기본 파일(`{DEFAULT_FILE}`)을 불러왔습니다.")
else:
    st.sidebar.info("엑셀 파일을 업로드하거나, 폴더 내에 '롤내전.xlsx'를 배치하세요.")


# ============================================================
# 4. 데이터가 로드된 경우
# ============================================================

if df is not None and not df.empty:

    st.sidebar.subheader("🔍 검색 필터")

    # 이름 필터
    people = ["전체"] + sorted([p for p in df["이름"].unique() if p])
    selected_person = st.sidebar.selectbox("이름", people)

    person_subset = df if selected_person == "전체" else df[df["이름"] == selected_person]

    # 라인 필터
    lines = ["전체"] + sorted([l for l in person_subset["라인"].unique() if l])
    selected_line = st.sidebar.selectbox("라인", lines)

    # 챔피언 필터
    champions = ["전체"] + sorted([c for c in person_subset["챔피언"].unique() if c])
    selected_champion = st.sidebar.selectbox("챔피언", champions)

    # 최종 필터링
    filtered_df = person_subset.copy()

    if selected_line != "전체":
        filtered_df = filtered_df[filtered_df["라인"] == selected_line]

    if selected_champion != "전체":
        filtered_df = filtered_df[filtered_df["챔피언"] == selected_champion]


    # ========================================================
    # 5. 매치 전적 계산
    # ========================================================
    match_df = (
        filtered_df[filtered_df["매치결과_전체"] != ""]
        .drop_duplicates(subset=["날짜", "이름"])
        .copy()
    )

    match_wins = sum(match_df["매치결과_전체"] == "승")
    match_losses = sum(match_df["매치결과_전체"] == "패")
    match_draws = sum(match_df["매치결과_전체"] == "무승부")
    match_games = len(match_df)

    match_completed = match_wins + match_losses
    match_winrate = (match_wins / match_completed * 100) if match_completed > 0 else 0.0


    # ========================================================
    # 6. 세트 전적 계산
    # ========================================================
    set_df = filtered_df[filtered_df["세트결과"] != ""].copy()

    set_wins = sum(set_df["세트결과"] == "승")
    set_losses = sum(set_df["세트결과"] == "패")
    set_draws = sum(set_df["세트결과"] == "무승부")
    set_games = len(set_df)

    set_completed = set_wins + set_losses
    set_winrate = (set_wins / set_completed * 100) if set_completed > 0 else 0.0


    # ========================================================
    # 7. KDA 및 통계 계산
    # ========================================================
    games = len(filtered_df)

    kills = filtered_df["킬"].sum()
    deaths = filtered_df["데스"].sum()
    assists = filtered_df["어시스트"].sum()

    kda = ((kills + assists) / deaths) if deaths > 0 else ((kills + assists) if games > 0 else 0.0)

    avg_damage = filtered_df["딜량"].mean() if games > 0 else 0.0
    avg_gold = filtered_df["골드"].mean() if games > 0 else 0.0


    # ========================================================
    # 8. 대시보드 출력
    # ========================================================
    st.subheader("🏆 매치 전적")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("매치 수", f"{match_games:,}매치")
    m2.metric("매치 승 / 패 / 무", f"{match_wins}승 {match_losses}패 {match_draws}무")
    m3.metric("매치 승률", f"{match_winrate:.1f}%")
    m4.metric("매치 전적", f"{match_wins}승 {match_losses}패")

    st.divider()

    st.subheader("🎮 세트 전적")
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("세트 수", f"{set_games:,}세트")
    s2.metric("세트 승 / 패 / 무", f"{set_wins}승 {set_losses}패 {set_draws}무")
    s3.metric("세트 승률", f"{set_winrate:.1f}%")
    s4.metric("세트 전적", f"{set_wins}승 {set_losses}패")

    st.divider()

    st.subheader("📊 개인 플레이 통계")
    p1, p2, p3, p4, p5 = st.columns(5)
    p1.metric("K / D / A", f"{int(kills)} / {int(deaths)} / {int(assists)}")
    p2.metric("평균 KDA", f"{kda:.2f}:1")
    p3.metric("평균 딜량", f"{avg_damage:,.0f}")
    p4.metric("평균 골드", f"{avg_gold:,.0f}")
    p5.metric("플레이 세트", f"{games:,}세트")

    st.divider()


    # ========================================================
    # 9. 상세 경기 기록
    # ========================================================
    st.subheader("📜 상세 경기 기록")

    display_df = filtered_df.iloc[::-1].copy()

    display_columns = [
        "날짜", "세트", "이름", "매치결과", "세트결과", "라인", "챔피언",
        "킬", "데스", "어시스트", "딜량", "골드"
    ]
    display_columns = [col for col in display_columns if col in display_df.columns]

    display_df = display_df[display_columns]

    for col in ["세트", "킬", "데스", "어시스트"]:
        if col in display_df.columns:
            display_df[col] = display_df[col].astype(int)

    st.dataframe(
        display_df,
        width="stretch",
        hide_index=True,
        column_config={
            "날짜": st.column_config.TextColumn("날짜"),
            "세트": st.column_config.NumberColumn("세트", format="%d"),
            "이름": st.column_config.TextColumn("이름"),
            "매치결과": st.column_config.TextColumn("매치결과"),
            "세트결과": st.column_config.TextColumn("세트결과"),
            "라인": st.column_config.TextColumn("라인"),
            "챔피언": st.column_config.TextColumn("챔피언"),
            "킬": st.column_config.NumberColumn("킬", format="%d"),
            "데스": st.column_config.NumberColumn("데스", format="%d"),
            "어시스트": st.column_config.NumberColumn("어시스트", format="%d"),
            "딜량": st.column_config.NumberColumn("딜량", format="%d"),
            "골드": st.column_config.NumberColumn("골드", format="%d"),
        },
    )
