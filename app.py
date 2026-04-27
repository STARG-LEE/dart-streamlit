import streamlit as st
import pandas as pd
import plotly.express as px
import OpenDartReader
from datetime import date, timedelta
from io import BytesIO

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
@st.cache_data(ttl=86400)
def load_corp_list(listed_only: bool = True) -> pd.DataFrame:
    """전체 기업 목록 DataFrame. 하루 1회 캐싱."""
    df = dart.corp_codes.copy()
    df["corp_name"] = df["corp_name"].astype(str).str.strip()
    df["stock_code"] = df["stock_code"].astype(str).str.strip()
    df = df[df["corp_name"] != ""]
    if listed_only:
        df = df[df["stock_code"].str.match(r"^\d{6}$", na=False)]
    df = df.drop_duplicates(subset=["corp_name", "stock_code"]).reset_index(drop=True)
    return df


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

listed_only = st.sidebar.toggle("상장사만 보기", value=True, help="끄면 비상장사 포함 전체 기업")

with st.spinner("기업 목록 로딩 중..."):
    corp_df = load_corp_list(listed_only=listed_only)

# selectbox에 표시할 라벨 ("삼성전자 (005930)")
def make_label(row):
    code = row["stock_code"]
    return f"{row['corp_name']} ({code})" if code and code.strip() else row["corp_name"]


corp_df = corp_df.assign(_label=corp_df.apply(make_label, axis=1))
labels = corp_df["_label"].tolist()

# 기본값: 삼성전자 (있으면)
default_idx = 0
samsung_match = corp_df.index[corp_df["corp_name"] == "삼성전자"].tolist()
if samsung_match:
    default_idx = labels.index(corp_df.loc[samsung_match[0], "_label"])

selected_label = st.sidebar.selectbox(
    f"기업 검색 ({len(labels):,}개)",
    options=labels,
    index=default_idx,
    placeholder="기업명 또는 종목코드 입력...",
    help="입력하면 자동으로 후보가 좁혀집니다.",
)

selected_row = corp_df[corp_df["_label"] == selected_label].iloc[0]
corp_input = selected_row["stock_code"] if selected_row["stock_code"] else selected_row["corp_name"]
display_name = selected_row["corp_name"]

st.sidebar.markdown("---")
st.sidebar.caption("데이터 출처: DART OpenAPI")
st.sidebar.caption("https://opendart.fss.or.kr/")


# ---------- 메인 ----------
st.title(f"📊 {display_name}")

# 기업개황으로 유효성 확인
try:
    company = fetch_company(corp_input)
except Exception as e:
    st.error(f"조회 실패: {e}")
    st.stop()

if not company or "corp_name" not in company:
    st.warning("기업 정보를 가져올 수 없습니다.")
    st.stop()

# ---------- 탭 ----------
tab_overview, tab_disclosures, tab_finance, tab_shareholders, tab_export = st.tabs(
    ["🏢 기업개황", "📰 공시목록", "💰 재무제표", "👥 지분공시", "📥 엑셀 다운로드"]
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

# ===== 5. 엑셀 다운로드 =====
with tab_export:
    st.markdown("**여러 기업**과 **원하는 항목**을 선택하면, 하나의 엑셀 파일에 통합되어 저장됩니다.")

    # ---- 일괄 입력 (붙여넣기) ----
    with st.expander("📋 기업명 일괄 붙여넣기 (여러 줄/쉼표 구분)"):
        bulk_text = st.text_area(
            "기업명 또는 종목코드를 줄바꿈 또는 쉼표로 구분해서 붙여넣으세요",
            height=150,
            placeholder="세니젠\n코이즈\n씨엑스아이\n...",
            key="bulk_input",
        )
        cols_bulk = st.columns([1, 1, 3])
        if cols_bulk[0].button("➕ 추가", use_container_width=True):
            tokens = [t.strip() for t in bulk_text.replace(",", "\n").splitlines() if t.strip()]
            current_list: list[str] = list(st.session_state.get("multi_corp", [selected_label]))
            new_labels: list[str] = []
            matched: list[str] = []
            unmatched: list[str] = []
            for tok in tokens:
                # 1) 종목코드(6자리) 정확매치
                if tok.isdigit() and len(tok) == 6:
                    hit = corp_df[corp_df["stock_code"] == tok]
                # 2) 기업명 정확매치
                else:
                    hit = corp_df[corp_df["corp_name"] == tok]
                    if hit.empty:
                        # 3) 기업명 부분일치 (1건만 매칭될 때 채택)
                        partial = corp_df[corp_df["corp_name"].str.contains(tok, regex=False, na=False)]
                        if len(partial) == 1:
                            hit = partial
                if hit.empty:
                    unmatched.append(tok)
                else:
                    lab = hit.iloc[0]["_label"]
                    matched.append(hit.iloc[0]["corp_name"])
                    if lab not in current_list and lab not in new_labels:
                        new_labels.append(lab)
            # 기존 선택은 위치 유지, 새로 매칭된 기업은 붙여넣기 순서대로 뒤에 추가
            st.session_state["multi_corp"] = current_list + new_labels
            st.session_state["bulk_matched"] = matched
            st.session_state["bulk_unmatched"] = unmatched
            st.rerun()
        if cols_bulk[1].button("🗑️ 비우기", use_container_width=True):
            st.session_state["multi_corp"] = []
            st.rerun()

        if st.session_state.get("bulk_matched"):
            st.success(f"✅ {len(st.session_state['bulk_matched'])}개 매칭: " + ", ".join(st.session_state["bulk_matched"]))
        if st.session_state.get("bulk_unmatched"):
            st.warning(
                f"⚠️ {len(st.session_state['bulk_unmatched'])}개 매칭 실패: "
                + ", ".join(st.session_state["bulk_unmatched"])
                + "  (정확한 기업명 또는 종목코드로 다시 시도해보세요)"
            )

    # ---- 기업 다중선택 ----
    if "multi_corp" not in st.session_state:
        st.session_state["multi_corp"] = [selected_label]

    selected_companies = st.multiselect(
        "조회할 기업 (복수 선택 가능)",
        options=labels,
        key="multi_corp",
        placeholder="기업명 또는 종목코드 입력...",
        help="위의 '일괄 붙여넣기'로 여러 기업을 한 번에 추가할 수 있습니다.",
    )

    st.markdown("##### 받고 싶은 항목")
    col_a, col_b = st.columns(2)
    with col_a:
        opt_overview = st.checkbox("🏢 기업개황", value=True)
        opt_disclosures = st.checkbox("📰 공시목록", value=True)
        opt_finance = st.checkbox("💰 재무제표", value=True)
    with col_b:
        opt_major = st.checkbox("👥 대량보유(5%) 보고", value=False)
        opt_exec = st.checkbox("👥 임원·주요주주 보고", value=False)

    st.caption("📅 공시목록 기간 / 재무제표 연도·보고서는 각 탭에서 설정한 값이 그대로 적용됩니다.")

    selected_any = any([opt_overview, opt_disclosures, opt_finance, opt_major, opt_exec])

    if not selected_companies:
        st.info("기업을 1개 이상 선택해주세요.")
    elif not selected_any:
        st.info("받고 싶은 항목을 1개 이상 선택해주세요.")
    else:
        st.caption(f"➡️ {len(selected_companies)}개 기업 × 선택 항목 = 통합 엑셀 1파일 생성")

        if st.button("📥 엑셀 파일 생성", type="primary", use_container_width=True):
            errors: list[str] = []
            overview_per_company: list[tuple[str, dict]] = []
            disclosures_dfs: list[pd.DataFrame] = []
            finance_groups: dict[str, list[pd.DataFrame]] = {}
            major_dfs: list[pd.DataFrame] = []
            exec_dfs: list[pd.DataFrame] = []

            progress = st.progress(0.0, text="시작...")
            n_total = len(selected_companies)

            def with_company_col(df: pd.DataFrame, cname: str, ccode: str) -> pd.DataFrame:
                """DataFrame 앞에 기업명/종목코드 컬럼 추가."""
                out = df.copy()
                out.insert(0, "기업명", cname)
                out.insert(1, "종목코드", ccode)
                return out

            for i, label in enumerate(selected_companies):
                row = corp_df[corp_df["_label"] == label].iloc[0]
                ccode = row["stock_code"]
                cname = row["corp_name"]
                corp_key = ccode if ccode else cname

                progress.progress((i + 1) / n_total, text=f"{cname} 조회 중... ({i + 1}/{n_total})")

                if opt_overview:
                    try:
                        comp = fetch_company(corp_key)
                        if comp:
                            overview_per_company.append((cname, comp))
                    except Exception as e:
                        errors.append(f"[{cname}] 기업개황: {e}")

                if opt_disclosures:
                    try:
                        df_d = fetch_list(corp_key, start, end)
                        if df_d is not None and len(df_d) > 0:
                            disclosures_dfs.append(with_company_col(df_d, cname, ccode))
                    except Exception as e:
                        errors.append(f"[{cname}] 공시목록: {e}")

                if opt_finance:
                    try:
                        fs_all = fetch_finstate(corp_key, year, reprt_code)
                        if fs_all is not None and len(fs_all) > 0:
                            tagged = with_company_col(fs_all, cname, ccode)
                            if "sj_nm" in tagged.columns:
                                for sj_name, sub in tagged.groupby("sj_nm"):
                                    finance_groups.setdefault(sj_name, []).append(
                                        sub.reset_index(drop=True)
                                    )
                            else:
                                finance_groups.setdefault("재무제표", []).append(tagged)
                    except Exception as e:
                        errors.append(f"[{cname}] 재무제표: {e}")

                if opt_major:
                    try:
                        df_m = fetch_major_shareholders(corp_key)
                        if df_m is not None and len(df_m) > 0:
                            major_dfs.append(with_company_col(df_m, cname, ccode))
                    except Exception as e:
                        errors.append(f"[{cname}] 대량보유: {e}")

                if opt_exec:
                    try:
                        df_e = fetch_major_shareholders_exec(corp_key)
                        if df_e is not None and len(df_e) > 0:
                            exec_dfs.append(with_company_col(df_e, cname, ccode))
                    except Exception as e:
                        errors.append(f"[{cname}] 임원주요주주: {e}")

            progress.empty()

            # ---- 시트 빌드 ----
            sheets: dict[str, pd.DataFrame] = {}

            if overview_per_company:
                # 기업개황은 비교가 쉽도록 wide 포맷: 항목 | 회사1 | 회사2 | ...
                all_keys: list[str] = []
                for _, comp in overview_per_company:
                    for k in comp.keys():
                        if k not in all_keys:
                            all_keys.append(k)
                wide = {"항목": all_keys}
                for cname, comp in overview_per_company:
                    col_name = cname
                    suffix = 2
                    while col_name in wide:
                        col_name = f"{cname}_{suffix}"
                        suffix += 1
                    wide[col_name] = [comp.get(k, "") for k in all_keys]
                sheets["기업개황"] = pd.DataFrame(wide)

            if disclosures_dfs:
                sheets["공시목록"] = pd.concat(disclosures_dfs, ignore_index=True)

            for sj_name, dfs in finance_groups.items():
                sheet_name = f"재무_{sj_name}"[:31]
                sheets[sheet_name] = pd.concat(dfs, ignore_index=True)

            if major_dfs:
                sheets["대량보유"] = pd.concat(major_dfs, ignore_index=True)

            if exec_dfs:
                sheets["임원주요주주"] = pd.concat(exec_dfs, ignore_index=True)

            if errors:
                with st.expander(f"⚠️ 일부 항목 누락 ({len(errors)}건)"):
                    for msg in errors:
                        st.write(f"- {msg}")

            if not sheets:
                st.error("저장할 데이터가 없습니다.")
            else:
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    workbook = writer.book
                    header_fmt = workbook.add_format({
                        "bold": True, "bg_color": "#D9E1F2", "border": 1,
                    })
                    for sheet_name, df_sheet in sheets.items():
                        safe_name = sheet_name[:31]
                        df_sheet.to_excel(writer, sheet_name=safe_name, index=False)
                        ws = writer.sheets[safe_name]
                        for col_idx, col_name in enumerate(df_sheet.columns):
                            ws.write(0, col_idx, str(col_name), header_fmt)
                            max_len = max(
                                df_sheet[col_name].astype(str).map(len).max() if len(df_sheet) else 0,
                                len(str(col_name)),
                            )
                            ws.set_column(col_idx, col_idx, min(max_len + 2, 50))
                        ws.freeze_panes(1, 0)

                buffer.seek(0)
                today_str = date.today().strftime("%Y%m%d")
                if len(selected_companies) == 1:
                    filename = f"DART_{overview_per_company[0][0] if overview_per_company else display_name}_{today_str}.xlsx"
                else:
                    filename = f"DART_{len(selected_companies)}개기업_{today_str}.xlsx"

                st.success(f"✅ {len(selected_companies)}개 기업 × {len(sheets)}개 시트 생성 완료")
                st.download_button(
                    label=f"⬇️ {filename} 다운로드",
                    data=buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
