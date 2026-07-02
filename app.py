# -*- coding: utf-8 -*-
"""원천세 대사표 자동생성 웹앱 — 원천데이터 zip 업로드 → 대사표 xlsx 생성."""
import streamlit as st
import tempfile, zipfile, os, io
import engine, builder

st.set_page_config(page_title="원천세 대사표 자동생성", layout="wide")

# ---- 비밀번호 잠금 ----
PASSWORD = "beyond"
if "auth" not in st.session_state:
    st.session_state.auth = False
if not st.session_state.auth:
    st.title("🔒 접속 인증")
    pw = st.text_input("비밀번호", type="password")
    if st.button("입장") or pw:
        if pw == PASSWORD:
            st.session_state.auth = True
            st.rerun()
        elif pw:
            st.error("비밀번호가 올바르지 않습니다.")
    st.stop()

st.title("📄 원천세 대사표 자동생성")
st.caption("원천자료를 담은 zip을 올리면 대사표를 자동 생성합니다. 폴더 구조는 상관없습니다 — 파일명·내용으로 월을 자동 인식합니다.")

with st.sidebar:
    st.header("입력 정보")
    company = st.text_input("회사명 / 사업자번호", value="", placeholder="예: (주)회사명 | 000-00-00000")
    year = st.number_input("귀속연도", min_value=2020, max_value=2100, value=2026, step=1)
    st.markdown("---")
    st.markdown("**업로드 방법**\n\n그 해 원천자료 파일들을 **zip으로 압축**해 올리세요. 월별 폴더로 나눠도, 한 폴더에 다 담아도 됩니다. 파일명·내용으로 **월을 자동 인식**합니다.\n\n대상: 급여대장·이행상황신고서·급여명세서·근로간이·사업간이지급조서")

up = st.file_uploader("원천데이터 zip 업로드", type=["zip"])

def find_root(base):
    """MM월 하위폴더를 가진 디렉터리를 찾는다."""
    import re
    for cur, dirs, _ in os.walk(base):
        if any(re.match(r"\d{1,2}월$", d) for d in dirs):
            return cur
    return base

if up is not None:
    with st.spinner("압축 해제 및 원천자료 파싱 중… (급여명세서가 많으면 1~2분 걸릴 수 있습니다)"):
        tmp = tempfile.mkdtemp()
        zpath = os.path.join(tmp, "up.zip")
        with open(zpath, "wb") as f:
            f.write(up.getbuffer())
        try:
            with zipfile.ZipFile(zpath) as z:
                z.extractall(tmp)
        except Exception as e:
            st.error(f"압축 해제 실패: {e}")
            st.stop()
        root = find_root(tmp)
        parsed = engine.scan_year(root, int(year))

    months = sorted([m for m, r in parsed.items() if r])
    if not months:
        st.warning("월별 폴더(01월, 02월 …)나 인식 가능한 파일을 찾지 못했습니다. zip 구조를 확인해 주세요.")
        st.stop()

    st.success(f"{len(months)}개 월 자료 인식: {', '.join(str(m)+'월' for m in months)}")

    # 파싱 요약 미리보기
    rows = []
    for m in range(1, 13):
        rec = parsed.get(m, {})
        if not rec: continue
        pr = rec.get("payroll") or {}
        gn = (rec.get("ihaeng") or {}).get("근로") or {}
        sa = (rec.get("ihaeng") or {}).get("사업") or {}
        ps = rec.get("payslips") or {}
        rows.append({
            "월": m,
            "급여(과세)": pr.get("급여"), "제출비과세": pr.get("제출비과세"),
            "신고": gn.get("지급액"), "소득세": gn.get("소득세"),
            "국민연금": ps.get("국민연금"), "건강보험": ps.get("건강보험"),
            "사업지급액": sa.get("지급액"), "사업간이제출": rec.get("saeop_gani"),
        })
    st.subheader("파싱 결과 요약")
    st.dataframe(rows, use_container_width=True)

    # 대사표 생성
    xbytes, miss = builder.build_bytes(parsed, int(year), company)
    fname = (company.split("|")[0].strip() or "원천세") + f"_대사표_{year}.xlsx"
    st.download_button("⬇️ 대사표 엑셀 다운로드", data=xbytes, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if miss:
        st.warning(f"미확보 자료 {len(miss)}건 — 엑셀의 '미확보자료' 시트에 목록 정리됨.")
        with st.expander("미확보 자료 목록 보기"):
            st.dataframe([{"항목": a, "월": b, "사유": c} for a, b, c in miss], use_container_width=True)
    st.info("파란 글자=외부자료 값복사(셀 메모에 출처파일), 검정=계산식, 주황=자료 미확보, 일치/불일치=단일 수식(엑셀에서 열면 색 자동표시).")
else:
    st.info("왼쪽에서 회사명·연도를 확인하고, 원천데이터 zip을 업로드하세요.")
