import streamlit as st
import pandas as pd
from io import BytesIO
import xlsxwriter

IFRS18_MAPPING = {
    "매출(컨텐츠제공)": "영업(Operating)",
    "매출(라이센스)": "영업(Operating)",
    "매출(광고수익)": "영업(Operating)",
    "매출(캐릭터상품)": "영업(Operating)",
    "조합관리보수": "영업(Operating)",
    "조합성과보수": "영업(Operating)",
    "매출(임대수익)": "영업(Operating)",
    "매출(임대료)": "영업(Operating)",
    "기타매출": "영업(Operating)",
    "서버수수료": "영업(Operating)",
    "지급수수료(컨텐츠제공)": "영업(Operating)",
    "상품매출원가": "영업(Operating)",
    "급여": "영업(Operating)",
    "상여금": "영업(Operating)",
    "퇴직급여": "영업(Operating)",
    "복리후생비": "영업(Operating)",
    "여비교통비": "영업(Operating)",
    "접대비": "영업(Operating)",
    "통신비": "영업(Operating)",
    "수도광열비": "영업(Operating)",
    "세금과공과": "영업(Operating)",
    "감가상각비": "영업(Operating)",
    "지급임차료": "영업(Operating)",
    "보험료": "영업(Operating)",
    "차량유지비": "영업(Operating)",
    "경상연구개발비": "영업(Operating)",
    "교육훈련비": "영업(Operating)",
    "도서인쇄비": "영업(Operating)",
    "사무용품비": "영업(Operating)",
    "소모품비": "영업(Operating)",
    "지급수수료(기타)": "영업(Operating)",
    "광고선전비": "영업(Operating)",
    "리스료": "영업(Operating)",
    "관리비": "영업(Operating)",
    "무형고정자산상각": "영업(Operating)",
    "외주용역비": "영업(Operating)",
    "판매촉진비": "영업(Operating)",
    "로얄티": "영업(Operating)",
    "주식보상비용": "영업(Operating)",
    "판매수수료(카카오)": "영업(Operating)",
    "대손상각비": "영업(Operating)",
    "지급수수료(로열티)": "영업(Operating)",
    "리스감가상각비": "영업(Operating)",
    "수선비": "영업(Operating)",
    "협회비": "영업(Operating)",
    "관리보수": "영업(Operating)",
    "성과보수비용": "영업(Operating)",
    "재고자산평가손실": "영업(Operating)",
    "잡이익": "투자(Investing)",
    "지원금수익": "투자(Investing)",
    "유형자산처분이익": "투자(Investing)",
    "무형자산처분이익": "투자(Investing)",
    "유형자산손상차손환입": "투자(Investing)",
    "무형자산손상차손환입": "투자(Investing)",
    "리스해지이익": "투자(Investing)",
    "종속기업투자주식처분이익": "투자(Investing)",
    "기타의대손충당금환입": "투자(Investing)",
    "기부금": "투자(Investing)",
    "기타의대손상각비": "투자(Investing)",
    "유형자산처분손실": "투자(Investing)",
    "유형자산폐기손실": "투자(Investing)",
    "유형자산손상차손": "투자(Investing)",
    "잡손실": "투자(Investing)",
    "관계기업투자손상차손": "투자(Investing)",
    "무형자산손상차손": "투자(Investing)",
    "무형자산처분손실": "투자(Investing)",
    "무형자산폐기손실": "투자(Investing)",
    "리스해지손실": "투자(Investing)",
    "종속기업투자주식처분손실": "투자(Investing)",
    "재고자산감모손실": "투자(Investing)",
    "지분법손익": "투자(Investing)",
    "이자수익": "재무(Financing)",
    "배당수익": "재무(Financing)",
    "외환차익": "재무(Financing)",
    "외화환산이익": "재무(Financing)",
    "당기손익인식금융자산평가이익": "재무(Financing)",
    "기타금융수익": "재무(Financing)",
    "당기손익-공정가치금융자산처분이익": "재무(Financing)",
    "이자비용": "재무(Financing)",
    "외환차손": "재무(Financing)",
    "외화환산손실": "재무(Financing)",
    "당기손익인식금융자산평가손실": "재무(Financing)",
    "당기손익-공정가치금융자산평가손실": "재무(Financing)",
    "지분법투자주식처분손실": "재무(Financing)",
    "기타금융비용": "재무(Financing)",
    "법인세비용": "법인세(Income Tax)",
    "비지배지분순손익": "비지배지분",
}

REVENUE_ACCOUNTS = {
    "매출(컨텐츠제공)", "매출(라이센스)", "매출(광고수익)", "매출(캐릭터상품)",
    "조합관리보수", "조합성과보수", "매출(임대수익)", "매출(임대료)", "기타매출",
    "잡이익", "지원금수익", "유형자산처분이익", "무형자산처분이익",
    "유형자산손상차손환입", "무형자산손상차손환입", "리스해지이익",
    "종속기업투자주식처분이익", "기타의대손충당금환입",
    "이자수익", "배당수익", "외환차익", "외화환산이익",
    "당기손익인식금융자산평가이익", "기타금융수익",
    "당기손익-공정가치금융자산처분이익", "지분법손익",
}

SKIP_KEYWORDS = [
    "영업이익", "법인세차감전", "비지배지분 차감전",
    "당기순", "소계", "합계", "순이익", "순손익", "세전",
]

SECTION_ORDER = [
    "영업(Operating)",
    "투자(Investing)",
    "재무(Financing)",
    "법인세(Income Tax)",
    "비지배지분",
]

SECTION_COLOR = {
    "영업(Operating)": "#1a3a5c",
    "투자(Investing)": "#2d1a5c",
    "재무(Financing)": "#1a4a2e",
    "법인세(Income Tax)": "#3a2a10",
    "비지배지분": "#555555",
}


def parse_num(val):
    try:
        if val is None:
            return 0
        s = str(val).strip()
        s = s.replace(",", "").replace(" ", "")
        s = s.replace("(", "-").replace(")", "")
        if s in ["", "nan", "NaN", "None", "-", "—"]:
            return 0
        return float(s)
    except Exception:
        return 0


def safe_round(v, unit_div):
    try:
        if v != v:
            return 0
        return round(float(v) / unit_div)
    except Exception:
        return 0


def preview_excel(file, skip_row):
    df = pd.read_excel(file, header=None, skiprows=int(skip_row))
    df.columns = [str(i) + "Å´" for i in range(len(df.columns))]
    return df


def load_income_statement(file, col_acct, col_amt, skip_row):
    df = pd.read_excel(file, header=None, skiprows=int(skip_row))
    df.columns = [str(i) + "Å´" for i in range(len(df.columns))]
    all_cols = df.columns.tolist()
    others = [c for c in all_cols if c not in [col_acct, col_amt]]
    filter_col = others[0] if others else None
    result = {}
    for _, row in df.iterrows():
        acct = str(row.get(col_acct, "")).strip()
        if acct in ["nan", "", "None", "NaN"]:
            continue
        if any(k in acct for k in SKIP_KEYWORDS):
            continue
        if filter_col:
            fv = str(row.get(filter_col, "")).strip()
            if fv in ["nan", "", "None", "NaN"]:
                continue
        amt = parse_num(row.get(col_amt, 0))
        if amt != 0 and amt == amt:
            result[acct] = amt
    return result


def convert_to_ifrs18(data):
    sections = {s: [] for s in SECTION_ORDER}
    unmapped = []
    for acct, amt in data.items():
        section = IFRS18_MAPPING.get(acct)
        if section:
            sections[section].append({"계정과목": acct, "금액": amt})
        else:
            unmapped.append({"계정과목": acct, "금액": amt})
    return sections, unmapped


def calc_section_total(items, section):
    total = 0
    for item in items:
        amt = item["금액"]
        if amt != amt:
            continue
        if section in ["법인세(Income Tax)", "비지배지분"]:
            total -= float(amt)
        elif item["계정과목"] in REVENUE_ACCOUNTS:
            total += float(amt)
        else:
            total -= float(amt)
    return total


def to_excel_file(sections, unmapped, unit_div, unit):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})
    ws = wb.add_worksheet("IFRS18")
    ft = wb.add_format({"bold": True, "font_size": 14, "align": "center", "bg_color": "#0a0a2e", "font_color": "#fff", "border": 1})
    fs = wb.add_format({"bold": True, "bg_color": "#1a3a5c", "font_color": "#fff", "border": 1})
    fd = wb.add_format({"indent": 2, "border": 1, "num_format": "#,##0"})
    fb = wb.add_format({"bold": True, "bg_color": "#dde8f0", "border": 1, "num_format": "#,##0"})
    fn = wb.add_format({"bold": True, "font_size": 13, "bg_color": "#1a4a2e", "font_color": "#fff", "border": 1, "num_format": "#,##0"})
    fu = wb.add_format({"bg_color": "#fff3cd", "border": 1, "num_format": "#,##0"})
    ws.set_column("A:A", 36)
    ws.set_column("B:B", 20)
    ws.merge_range("A1:B1", "손익계산서 (IFRS 18)", ft)
    r = 1
    grand = 0
    for sec in SECTION_ORDER:
        items = sections.get(sec, [])
        if not items:
            continue
        ws.write(r, 0, "▌ " + sec, fs)
        ws.write(r, 1, "", fs)
        r += 1
        for item in items:
            ws.write(r, 0, item["계정과목"], fd)
            ws.write(r, 1, safe_round(item["금액"], unit_div), fd)
            r += 1
        st_ = calc_section_total(items, sec)
        grand += st_
        ws.write(r, 0, sec + " 소계", fb)
        ws.write(r, 1, safe_round(st_, unit_div), fb)
        r += 1
    ws.write(r, 0, "★ 당기순손익", fn)
    ws.write(r, 1, safe_round(grand, unit_div), fn)
    r += 1
    if unmapped:
        r += 1
        ws.write(r, 0, "미매핑 계정", fs)
        ws.write(r, 1, "", fs)
        r += 1
        for item in unmapped:
            ws.write(r, 0, item["계정과목"], fu)
            ws.write(r, 1, safe_round(item["금액"], unit_div), fu)
            r += 1
    wb.close()
    return output.getvalue()


st.set_page_config(page_title="IFRS 18 변환기", page_icon="📊", layout="wide")
st.title("📊 IFRS 18 손익계산서 변환기")
st.caption("기존 손익계산서 Excel → IFRS 18 형식으로 자동 변환")

with st.sidebar:
    st.header("📁 파일 업로드")
    uploaded = st.file_uploader("손익계산서 Excel 업로드", type=["xlsx", "xls"])
    st.divider()
    st.header("⚙️ 옵션")
    unit = st.selectbox("금액 단위", ["원", "천원", "백만원"])
    unit_div = {"원": 1, "천원": 1000, "백만원": 1000000}[unit]
    show_zero = st.toggle("금액 0인 계정 표시", value=False)
    if uploaded:
        st.divider()
        st.header("🔧 컬럼 설정")
        skip_row = st.number_input("상단 건너뛸 행 수", min_value=0, max_value=20, value=0)
        try:
            df_pre = preview_excel(uploaded, skip_row)
            cols = df_pre.columns.tolist()
            col_acct = st.selectbox("계정과목 컬럼 (A열=0열)", cols, index=0)
            col_amt = st.selectbox("금액 컬럼 (C열=2열)", cols, index=min(2, len(cols) - 1))
            with st.expander("📄 Excel 미리보기"):
                st.dataframe(df_pre.head(10), use_container_width=True)
        except Exception as ex:
            st.error("파일 오류: " + str(ex))
            st.stop()
    else:
        skip_row = 0
        col_acct = "0열"
        col_amt = "2열"

if uploaded:
    with st.spinner("변환 중..."):
        try:
            data = load_income_statement(uploaded, col_acct, col_amt, skip_row)
            sections, unmapped = convert_to_ifrs18(data)
        except Exception as ex:
            st.error("처리 오류: " + str(ex))
            st.stop()
    if not data:
        st.warning("⚠️ 인식된 계정과목이 없습니다. 컬럼을 다시 선택해보세요.")
        st.stop()
    op = calc_section_total(sections.get("영업(Operating)", []), "영업(Operating)")
    inv = calc_section_total(sections.get("투자(Investing)", []), "투자(Investing)")
    fin = calc_section_total(sections.get("재무(Financing)", []), "재무(Financing)")
    tax = calc_section_total(sections.get("법인세(Income Tax)", []), "법인세(Income Tax)")
    nci = calc_section_total(sections.get("비지배지분", []), "비지배지분")
    net = op + inv + fin + tax + nci
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("영업손익", "{:,.0f} {}".format(op / unit_div, unit))
    with c2:
        st.metric("투자손익", "{:,.0f} {}".format(inv / unit_div, unit))
    with c3:
        st.metric("재무손익", "{:,.0f} {}".format(fin / unit_div, unit))
    with c4:
        st.metric("당기순손익", "{:,.0f} {}".format(net / unit_div, unit))
    st.divider()
    cl, cr = st.columns([3, 2])
    with cl:
        st.subheader("📋 IFRS 18 손익계산서")
        grand = 0
        for sec in SECTION_ORDER:
            items = sections.get(sec, [])
            if not items:
                continue
            bg = SECTION_COLOR.get(sec, "#333")
            st.markdown(
                "<div style='background:" + bg + ";color:white;padding:9px 14px;"
                "border-radius:6px;font-weight:bold;margin:10px 0 2px 0'>"
                "▌ " + sec + "</div>",
                unsafe_allow_html=True,
            )
            for item in items:
                a = item["금액"]
                if a != a:
                    continue
                if not show_zero and a == 0:
                    continue
                display = "{:,.0f} {}".format(float(a) / unit_div, unit)
                st.markdown(
                    "<div style='display:flex;justify-content:space-between;"
                    "padding:6px 14px 6px 24px;color:#111;background:#f5f5f5;"
                    "border-left:3px solid " + bg + ";margin:1px 0;"
                    "border-radius:0 4px 4px 0;font-size:14px'>"
                    "<span>" + item["계정과목"] + "</span>"
                    "<span style='font-weight:600'>" + display + "</span></div>",
                    unsafe_allow_html=True,
                )
            st_ = calc_section_total(items, sec)
            grand += st_
            st.markdown(
                "<div style='display:flex;justify-content:space-between;"
                "padding:9px 14px;background:" + bg + ";border-radius:6px;"
                "font-weight:bold;color:white;margin:2px 0 8px 0'>"
                "<span>" + sec + " 소계</span>"
                "<span>{:,.0f} {}</span></div>".format(st_ / unit_div, unit),
                unsafe_allow_html=True,
            )
        nb = "#e8f5e9" if grand >= 0 else "#ffebee"
        nc = "#1b5e20" if grand >= 0 else "#b71c1c"
        st.markdown(
            "<div style='display:flex;justify-content:space-between;"
            "padding:14px 18px;background:" + nb + ";border-radius:8px;"
            "font-weight:bold;font-size:17px;color:" + nc + ";"
            "border:2px solid " + nc + ";margin-top:12px'>"
            "<span>★ 당기순손익</span>"
            "<span>{:,.0f} {}</span></div>".format(grand / unit_div, unit),
            unsafe_allow_html=True,
        )
    with cr:
        if unmapped:
            st.subheader("⚠️ 미매핑 계정 (" + str(len(unmapped)) + "개)")
            st.caption("IFRS 18 분류가 없어요. 직접 섹션을 지정해주세요.")
            for item in unmapped:
                a = item["금액"] if item["금액"] == item["금액"] else 0
                u1, u2 = st.columns([2, 1])
                with u1:
                    st.markdown(
                        "<div style='padding:6px 10px;background:#fff8e1;"
                        "border:1px solid #ffca28;border-radius:6px;"
                        "font-size:13px;color:#333'>"
                        "<b>" + item["계정과목"] + "</b><br>"
                        "<small>{:,.0f} {}</small></div>".format(float(a) / unit_div, unit),
                        unsafe_allow_html=True,
                    )
                with u2:
                    st.selectbox(
                        "섹션",
                        SECTION_ORDER,
                        key="map_" + item["계정과목"],
                        label_visibility="collapsed",
                    )
        else:
            st.success("✅ 모든 계정이 IFRS 18에 매핑되었습니다!")
        st.divider()
        with st.expander("📄 원본 데이터 확인"):
            rows = []
            for k, v in data.items():
                rows.append({"계정과목": k, "금액(" + unit + ")": safe_round(v, unit_div)})
            if rows:
                st.dataframe(pd.DataFrame(rows), use_container_width=True)
    st.divider()
    st.download_button(
        "📥 IFRS 18 손익계산서 Excel 다운로드",
        data=to_excel_file(sections, unmapped, unit_div, unit),
        file_name="손익계산서_IFRS18.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
else:
    st.info("👈 왼쪽 사이드바에서 손익계산서 Excel 파일을 업로드하세요!")
    st.dataframe(
        pd.DataFrame({
            "계정과목": ["매출(컨텐츠제공)", "급여", "이자수익", "법인세비용"],
            "분류": ["영업수익", "영업비용", "금융수익", "법인세"],
            "금액": [100000000, 50000000, 5000000, 3000000],
        }),
        use_container_width=True,
    )
