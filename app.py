import streamlit as st
import pandas as pd
import plotly.express as px
import OpenDartReader
from datetime import date, timedelta

st.set_page_config(page_title="DART 기업정보 조회", page_icon="📊", layout="wide")

# ---------- 인증키 로드 ----------
try:
    API_KEY = st.secrets["DART_KEY"]
except Exception:
    st.error("⚠️ `.streamlit/secrets.toml`에 DART_KEY를 설정해주세요.")
    st.stop()

if not API_KEY or API_KEY.startswith("여기에"):
    st.error("⚠️ `.streamlit/secrets.toml`에 실제 DART 인증키를 입력해주세요.")
    st.stop()


@st.cache_resource
def get_dart():
    return OpenDartReader(API_KEY)


dart = get_dart()


# ---------- 캐시된 조회 함수 ----------
@st.cache_data(ttl=3600)
def fetch_company(corp):
    return dart.company(corp)


@st.cache_data(ttl=600)
def fetch_list(corp, start, end):
    return dart.list(corp, start=start.isoformat(), end=end.isoformat())


@st.cache_data(ttl=3600)
def fetch_finstate(corp, year, reprt_code):
    return dart.finstate(corp, year, reprt_code=reprt_code)


@st.cache_data(ttl=3600)
def fetch_major_shareholders(corp):
    return dart.major_shareholders(corp)


@st.cache_data(ttl=3600)
def fetch_major_shareholders_exec(corp):
    return dart.major_shareholders_exec(corp)


# ---------- 사이드바 ----------
st.sidebar.title("📊 DART 조회")
corp_input = st.sidebar.text_input("기업명 또는 종목코드", value="삼성전자")

st.sidebar.markdown("---")
st.sidebar.caption("데이터 출처: DART OpenAPI")
st.sidebar.caption("https://opendart.fss.or.kr/")


# ---------- 메인 ----------
st.title(f"📊 {corp_input}")

if not corp_input:
    st.info("좌측에 기업명 또는 종목코드를 입력하세요.")
    st.stop()

# 기업개황으로 유효성 확인
try:
    company = fetch_company(corp_input)
except Exception as e:
    st.error(f"조회 실패: {e}")
    st.stop()

if not company or "corp_name" not in company:
    st.warning("해당 기업을 찾을 수 없습니다. 정확한 기업명 또는 종목코드를 입력해주세요.")
    st.stop()

# ---------- 탭 ----------
tab_overview, tab_disclosures, tab_finance, tab_shareholders = st.tabs(
    ["🏢 기업개황", "📰 공시목록", "💰 재무제표", "👥 지분공시"]
)

# ===== 1. 기업개황 =====
with tab_overview:
    col1, col2, col3 = st.columns(3)
    col1.metric("기업명", company.get("corp_name", "-"))
    col2.metric("종목코드", company.get("stock_code") or "비상장")
    col3.metric("대표이사", company.get("ceo_nm", "-"))

    with st.expander("상세정보", expanded=True):
        info = {
            "영문명": company.get("corp_name_eng"),
            "법인구분": company.get("corp_cls"),
            "법인등록번호": company.get("jurir_no"),
            "사업자등록번호": company.get("bizr_no"),
            "주소": company.get("adres"),
            "홈페이지": company.get("hm_url"),
            "전화번호": company.get("phn_no"),
            "팩스": company.get("fax_no"),
            "업종": company.get("induty_code"),
            "결산월": company.get("acc_mt"),
            "설립일": company.get("est_dt"),
        }
        st.dataframe(
            pd.DataFrame(info.items(), columns=["항목", "값"]),
            use_container_width=True,
            hide_index=True,
        )

# ===== 2. 공시목록 =====
with tab_disclosures:
    c1, c2 = st.columns(2)
    start = c1.date_input("시작일", value=date.today() - timedelta(days=180))
    end = c2.date_input("종료일", value=date.today())

    try:
        df = fetch_list(corp_input, start, end)
        if df is None or len(df) == 0:
            st.info("해당 기간에 공시가 없습니다.")
        else:
            st.caption(f"총 {len(df)}건")
            show_cols = [c for c in ["rcept_dt", "report_nm", "flr_nm", "rm", "rcept_no"] if c in df.columns]
            view = df[show_cols].copy()
            if "rcept_no" in view.columns:
                view["원문링크"] = view["rcept_no"].apply(
                    lambda x: f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={x}"
                )
            st.dataframe(
                view,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "원문링크": st.column_config.LinkColumn("원문", display_text="📄 보기"),
                    "rcept_dt": "접수일",
                    "report_nm": "보고서명",
                    "flr_nm": "제출인",
                    "rm": "비고",
                },
            )
    except Exception as e:
        st.error(f"공시 조회 실패: {e}")

# ===== 3. 재무제표 =====
with tab_finance:
    c1, c2 = st.columns(2)
    year = c1.selectbox("사업연도", list(range(date.today().year, 2014, -1)))
    reprt_map = {
        "사업보고서 (연간)": "11011",
        "1분기보고서": "11013",
        "반기보고서": "11012",
        "3분기보고서": "11014",
    }
    reprt_label = c2.selectbox("보고서 종류", list(reprt_map.keys()))
    reprt_code = reprt_map[reprt_label]

    try:
        fs = fetch_finstate(corp_input, year, reprt_code)
        if fs is None or len(fs) == 0:
            st.info("해당 보고서가 없습니다. (아직 공시되지 않았거나 다른 보고서를 선택해 보세요)")
        else:
            sj_options = fs["sj_nm"].unique().tolist() if "sj_nm" in fs.columns else []
            if sj_options:
                sj = st.radio("재무제표 종류", sj_options, horizontal=True)
                view = fs[fs["sj_nm"] == sj].copy()
            else:
                view = fs

            cols_keep = [c for c in [
                "account_nm", "thstrm_amount", "frmtrm_amount", "bfefrmtrm_amount"
            ] if c in view.columns]
            rename = {
                "account_nm": "계정",
                "thstrm_amount": "당기",
                "frmtrm_amount": "전기",
                "bfefrmtrm_amount": "전전기",
            }
            disp = view[cols_keep].rename(columns=rename)
            st.dataframe(disp, use_container_width=True, hide_index=True)

            # 주요 계정 차트
            if "계정" in disp.columns and "당기" in disp.columns:
                key_accounts = ["매출액", "영업이익", "당기순이익", "자산총계", "부채총계", "자본총계"]
                chart_df = disp[disp["계정"].isin(key_accounts)].copy()
                if not chart_df.empty:
                    for col in ["당기", "전기", "전전기"]:
                        if col in chart_df.columns:
                            chart_df[col] = pd.to_numeric(
                                chart_df[col].astype(str).str.replace(",", ""), errors="coerce"
                            )
                    melt = chart_df.melt(id_vars="계정", var_name="기간", value_name="금액").dropna()
                    fig = px.bar(
                        melt, x="계정", y="금액", color="기간", barmode="group",
                        title="주요 계정 추이",
                    )
                    st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.error(f"재무제표 조회 실패: {e}")

# ===== 4. 지분공시 =====
with tab_shareholders:
    sub1, sub2 = st.tabs(["대량보유(5%)", "임원·주요주주"])

    with sub1:
        try:
            ms = fetch_major_shareholders(corp_input)
            if ms is None or len(ms) == 0:
                st.info("대량보유 보고 내역이 없습니다.")
            else:
                st.dataframe(ms, use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(f"조회 실패: {e}")

    with sub2:
        try:
            mse = fetch_major_shareholders_exec(corp_input)
            if mse is None or len(mse) == 0:
                st.info("임원·주요주주 보고 내역이 없습니다.")
            else:
                st.dataframe(mse, use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(f"조회 실패: {e}")
