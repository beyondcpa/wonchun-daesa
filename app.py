# -*- coding: utf-8 -*-
"""원천세 대사표 자동생성 웹앱 — zip 업로드 또는 구글드라이브 폴더에서 자동 생성."""
import streamlit as st
import tempfile, zipfile, os, re
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
st.caption("원천자료를 zip으로 올리거나 구글드라이브 폴더를 지정하면 대사표를 자동 생성합니다. 폴더 구조는 상관없습니다 — 파일명·내용으로 월을 자동 인식합니다.")

with st.sidebar:
    st.header("입력 정보")
    company = st.text_input("회사명 / 사업자번호", value="", placeholder="예: (주)회사명 | 000-00-00000")
    year = st.number_input("귀속연도", min_value=2020, max_value=2100, value=2026, step=1)
    st.markdown("---")
    st.caption("대상 자료: 급여대장 · 이행상황신고서 · 급여명세서 · 근로간이 · 사업간이지급조서")


def find_root(base):
    for cur, dirs, _ in os.walk(base):
        if any(re.match(r"\d{1,2}월$", d) for d in dirs):
            return cur
    return base


def process(root):
    """폴더(root)를 스캔→대사표 생성→미리보기·다운로드."""
    parsed = engine.scan_year(root, int(year))
    months = sorted([m for m, r in parsed.items() if r])
    if not months:
        st.warning("인식 가능한 파일을 찾지 못했습니다. (급여대장/신고서/급여명세/근로간이/사업간이 라는 단어가 파일명에 있어야 합니다.)")
        return
    st.success(f"{len(months)}개 월 자료 인식: {', '.join(str(m)+'월' for m in months)}")

    rows = []
    for m in range(1, 13):
        rec = parsed.get(m, {})
        if not rec: continue
        pr = rec.get("payroll") or {}
        gn = (rec.get("ihaeng") or {}).get("근로") or {}
        sa = (rec.get("ihaeng") or {}).get("사업") or {}
        ps = rec.get("payslips") or {}
        rows.append({"월": m, "급여(과세)": pr.get("급여"), "제출비과세": pr.get("제출비과세"),
                     "신고": gn.get("지급액"), "소득세": gn.get("소득세"),
                     "국민연금": ps.get("국민연금"), "건강보험": ps.get("건강보험"),
                     "사업지급액": sa.get("지급액"), "사업간이제출": rec.get("saeop_gani")})
    st.subheader("파싱 결과 요약")
    st.dataframe(rows, use_container_width=True)

    xbytes, miss = builder.build_bytes(parsed, int(year), company)
    fname = (company.split("|")[0].strip() or "원천세") + f"_대사표_{year}.xlsx"
    st.download_button("⬇️ 대사표 엑셀 다운로드", data=xbytes, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if miss:
        st.warning(f"미확보 자료 {len(miss)}건 — 엑셀의 '미확보자료' 시트에 목록 정리됨.")
        with st.expander("미확보 자료 목록 보기"):
            st.dataframe([{"항목": a, "월": b, "사유": c} for a, b, c in miss], use_container_width=True)
    st.info("파란 글자=외부자료 값복사(셀 메모에 출처파일), 검정=계산식, 주황=자료 미확보, 일치/불일치=단일 수식(엑셀에서 열면 색 자동표시).")


mode = st.radio("자료 불러오기 방식", ["📁 zip 업로드", "☁️ 구글드라이브에서 읽기"], horizontal=True)

if mode == "📁 zip 업로드":
    up = st.file_uploader("원천데이터 zip 업로드", type=["zip"])
    if up is not None:
        with st.spinner("압축 해제 및 파싱 중…"):
            tmp = tempfile.mkdtemp()
            zpath = os.path.join(tmp, "up.zip")
            with open(zpath, "wb") as f:
                f.write(up.getbuffer())
            try:
                with zipfile.ZipFile(zpath) as z:
                    z.extractall(tmp)
            except Exception as e:
                st.error(f"압축 해제 실패: {e}"); st.stop()
            root = find_root(tmp)
        process(root)
    else:
        st.info("연도 폴더(또는 파일들)를 zip으로 압축해 올리세요.")

else:  # 구글드라이브
    st.markdown("해당 **연도 폴더(예: 2026)** 의 구글드라이브 링크 또는 폴더 ID를 넣으세요. 그 아래 월별 자료를 자동으로 받아옵니다.")
    link = st.text_input("구글드라이브 폴더 링크 / ID", placeholder="https://drive.google.com/drive/folders/XXXXXXXX")
    if st.button("드라이브에서 불러오기", type="primary") and link:
        try:
            import gdrive
        except Exception:
            st.error("gdrive 모듈/구글 라이브러리가 없습니다. requirements.txt 확인."); st.stop()
        cred = None
        if "gcp_service_account" in st.secrets:
            cred = st.secrets["gcp_service_account"]
        elif "gcp_service_account_json" in st.secrets:
            cred = st.secrets["gcp_service_account_json"]
        if cred is None:
            st.error("드라이브 인증이 설정되지 않았습니다. (Streamlit 앱 설정 → Secrets 에 서비스계정 등록 — GDRIVE_SETUP.md 참고)")
            st.stop()
        try:
            svc = gdrive.get_service(cred)
        except Exception as e:
            st.error(f"드라이브 인증 실패: {e}"); st.stop()
        tmp = tempfile.mkdtemp()
        with st.spinner("구글드라이브에서 파일 받는 중…"):
            try:
                n = gdrive.download_folder(svc, gdrive.extract_folder_id(link), tmp)
            except Exception as e:
                st.error(f"드라이브 읽기 실패: {e} (서비스계정 이메일에 해당 폴더가 공유돼 있는지 확인)"); st.stop()
        if n == 0:
            st.warning("폴더에서 파일을 받지 못했습니다. 폴더 공유·링크를 확인하세요."); st.stop()
        st.success(f"드라이브에서 {n}개 파일 수신")
        process(find_root(tmp))
