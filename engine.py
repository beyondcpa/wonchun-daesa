# -*- coding: utf-8 -*-
"""원천데이터 → 대사표 값 자동 추출 엔진.
   폴더 구조에 무관하게(평평하게 담든, 하위폴더든, 압축 속 압축이든) 파일명·내용으로
   월(귀속월)과 자료 종류를 자동 판별해 파싱한다."""
import re, io, os, zipfile, warnings, tempfile
from collections import defaultdict
warnings.filterwarnings("ignore")
import fitz, xlrd

def _num(s):
    m = re.search(r"-?[\d,]+", str(s))
    return int(m.group().replace(",", "")) if m else None

def pdf_text(src):
    if isinstance(src, (bytes, bytearray)):
        doc = fitz.open(stream=src, filetype="pdf")
    else:
        doc = fitz.open(src)
    t = "\n".join(pg.get_text("text", sort=True) for pg in doc)
    doc.close()
    return t

# ---------- 월 판별 (파일명 → 내용) ----------
def month_from_name(name):
    m = re.search(r"20\d{2}[.\-_년\s]*(0[1-9]|1[0-2])(?:\D|$)", name)   # 202601 / 2026.01 / 2026년 01
    if m: return int(m.group(1))
    m = re.search(r"(1[0-2]|[1-9])\s*월", name)                          # 3월 / 03월 / 12월
    if m: return int(m.group(1))
    return None

def month_from_ihaeng(t):
    m = re.search(r"귀속연[월왈][^\d]*20\d{2}\s*년?\s*(0[1-9]|1[0-2])", t)
    return int(m.group(1)) if m else None

def month_from_payslip(t):
    m = re.search(r"20\d{2}\s*년?\s*(0[1-9]|1[0-2])\s*월분", t)
    return int(m.group(1)) if m else None

def month_from_saeop_gani(t):
    m = re.search(r"\[[○√ⓞ]\]\s*(1[0-2]|[1-9])\s*월", t)                 # 지급월 체크박스
    if m: return int(m.group(1))
    m = re.search(r"지급월[^\d]*(1[0-2]|[1-9])\s*월", t)
    return int(m.group(1)) if m else None

# ---------- 개별 파서 ----------
def parse_payroll(src):
    wb = xlrd.open_workbook(file_contents=src) if isinstance(src, (bytes, bytearray)) else xlrd.open_workbook(src)
    sh = wb.sheet_by_index(0)
    for r in range(sh.nrows):
        row = [sh.cell_value(r, c) for c in range(sh.ncols)]
        if any(str(x).replace(" ", "").startswith("급여계") for x in row):
            nums = [x for x in row if isinstance(x, (int, float)) and x]
            if len(nums) >= 3:
                return {"급여": int(nums[0]), "제출비과세": int(nums[1]), "급여계": int(nums[2])}
    return None

def parse_ihaeng_text(t):
    out = {"근로": None, "사업": None}
    for ln in t.split("\n"):
        if "A01" in ln:
            n = re.findall(r"[\d,]+", ln)
            if len(n) >= 3: out["근로"] = {"인원": _num(n[-3]), "지급액": _num(n[-2]), "소득세": _num(n[-1])}
        if re.search(r"매월징수\s*A25", ln):
            n = re.findall(r"[\d,]+", ln)
            if len(n) >= 3: out["사업"] = {"인원": _num(n[-3]), "지급액": _num(n[-2]), "소득세": _num(n[-1])}
    return out

_FIELDS = ["국민연금","건강보험","장기요양보험료","고용보험","연말정산소득세","연말정산지방소득세","소득세","지방소득세"]
def parse_payslip_text(t):
    d = {}
    for ln in t.split("\n"):
        for f in _FIELDS:
            m = re.search(re.escape(f) + r"\s*(-?[\d,]+)", ln)
            if m and f not in d: d[f] = int(m.group(1).replace(",", ""))
    return d

def parse_geun_gani_text(t):
    tt = t.replace(" ", "")
    half = "상반기" if ("[√]상반기" in tt or "[v]상반기" in tt.lower()) else ("하반기" if ("[√]하반기" in tt or "[v]하반기" in tt.lower()) else None)
    m = re.search(r"과세소득[^\d]*([\d,]{6,})", t)
    m2 = re.search(r"근로자총인원\s*([\d,]+)", t)
    return {"반기": half, "과세소득": _num(m.group(1)) if m else None, "인원": _num(m2.group(1)) if m2 else None}

def parse_saeop_gani_text(t):
    m = re.search(r"(\d+)\s*명\s+([\d,]{5,})", t)          # ⑤총지급액계 (N명 뒤 총액)
    if m: return _num(m.group(2))
    m = re.search(r"총\s*지급액[^\d]*([\d,]{5,})", t) or re.search(r"총지급액계[^\d]*([\d,]{5,})", t)
    if m: return _num(m.group(1))
    nums = [int(x.replace(",", "")) for x in re.findall(r"[\d,]{7,}", t)]
    return max(nums) if nums else None

# ---------- 전체 파일 수집(압축 속 압축까지) ----------
def _collect(root):
    """(파일명, 경로 또는 bytes) 리스트. 중첩 zip은 풀어서 포함."""
    out = []
    for cur, _, files in os.walk(root):
        for f in files:
            p = os.path.join(cur, f)
            if f.lower().endswith(".zip"):
                try:
                    z = zipfile.ZipFile(p)
                    for zn in z.namelist():
                        if zn.endswith("/"): continue
                        out.append((os.path.basename(zn), z.read(zn)))
                except Exception: pass
            else:
                out.append((f, p))
    return out

def scan_year(root, year):
    """폴더 구조 무관 스캔. 반환: {month: {payroll, ihaeng, payslips, geun_gani, saeop_gani}}"""
    result = defaultdict(dict)
    payslips = defaultdict(list)
    for name, src in _collect(root):
        low = name.lower()
        try:
            if "급여대장" in name and low.endswith((".xls", ".xlsx")):
                mm = month_from_name(name)
                if mm:
                    pr = parse_payroll(src)
                    if pr: result[mm]["payroll"] = pr
            elif "급여명세" in name and low.endswith(".pdf"):
                t = pdf_text(src)
                mm = month_from_payslip(t) or month_from_name(name)
                if mm: payslips[mm].append(t)
            elif ("이행상황" in name or "원천세" in name or "신고" in name) and low.endswith(".pdf") and "급여명세" not in name:
                t = pdf_text(src)
                if "A01" in t or "매월징수" in t:                       # 이행상황신고서만
                    mm = month_from_ihaeng(t) or month_from_name(name)
                    if mm: result[mm]["ihaeng"] = parse_ihaeng_text(t)
            elif "근로간이" in name and low.endswith(".pdf"):
                t = pdf_text(src); g = parse_geun_gani_text(t)
                # 반기 조서는 대표월에 저장(상=6, 하=12) — build 단계에서 반기로 사용
                key = 6 if g.get("반기") == "상반기" else (12 if g.get("반기") == "하반기" else (month_from_name(name) or 6))
                result[key]["geun_gani"] = g
            elif "사업간이" in name and low.endswith(".pdf"):
                t = pdf_text(src)
                mm = month_from_saeop_gani(t) or month_from_name(name)
                if mm: result[mm]["saeop_gani"] = parse_saeop_gani_text(t)
        except Exception:
            pass
    for mm, texts in payslips.items():
        agg = {}
        for t in texts:
            for k, v in parse_payslip_text(t).items(): agg[k] = agg.get(k, 0) + v
        agg["_files"] = len(texts)
        result[mm]["payslips"] = agg
    return dict(result)

if __name__ == "__main__":
    import sys
    root = sys.argv[1] if len(sys.argv) > 1 else "."
    for m, rec in sorted(scan_year(root, 2026).items()):
        print(f"=== {m}월 ===")
        for k, v in rec.items():
            print("  ", k, ":", v)
