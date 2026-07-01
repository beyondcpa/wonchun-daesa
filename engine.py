# -*- coding: utf-8 -*-
"""원천데이터(급여대장·이행상황신고서·급여명세서·간이지급조서) → 대사표 값 자동 추출 엔진."""
import re, io, glob, os, zipfile, warnings
warnings.filterwarnings("ignore")
import pdfplumber, xlrd

def _num(s):
    m = re.search(r"-?[\d,]+", str(s))
    return int(m.group().replace(",", "")) if m else None

def pdf_text(path_or_bytes):
    src = io.BytesIO(path_or_bytes) if isinstance(path_or_bytes, (bytes, bytearray)) else path_or_bytes
    with pdfplumber.open(src) as pdf:
        return "\n".join((p.extract_text() or "") for p in pdf.pages)

# ---------- 급여대장(.xls) : 기본급(과세) / 식대(제출비과세) / 급여계 ----------
def parse_payroll(path):
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    for r in range(sh.nrows):
        row = [sh.cell_value(r, c) for c in range(sh.ncols)]
        if any(str(x).replace(" ", "").startswith("급여계") for x in row):
            nums = [x for x in row if isinstance(x, (int, float)) and x]
            if len(nums) >= 3:
                return {"급여": int(nums[0]), "제출비과세": int(nums[1]), "급여계": int(nums[2])}
    return None

# ---------- 이행상황신고서(.pdf) : 근로 A01 / 사업 A25·A30 ----------
def parse_ihaeng(path):
    t = pdf_text(path)
    out = {"근로": None, "사업": None}
    for ln in t.split("\n"):
        if "A01" in ln:
            n = re.findall(r"[\d,]+", ln)
            if len(n) >= 3:
                out["근로"] = {"인원": _num(n[-3]), "지급액": _num(n[-2]), "소득세": _num(n[-1])}
        if re.search(r"매월징수 A25", ln):
            n = re.findall(r"[\d,]+", ln)
            if len(n) >= 3:
                out["사업"] = {"인원": _num(n[-3]), "지급액": _num(n[-2]), "소득세": _num(n[-1])}
    return out

# ---------- 급여명세서(.pdf) 집계 : 4대보험·소득세·연말정산 ----------
_FIELDS = ["국민연금", "건강보험", "장기요양보험료", "고용보험",
           "연말정산소득세", "연말정산지방소득세", "소득세", "지방소득세"]
def _parse_payslip(t):
    d = {}
    for ln in t.split("\n"):
        for f in _FIELDS:
            m = re.search(re.escape(f) + r"\s*(-?[\d,]+)", ln)
            if m and f not in d:
                d[f] = int(m.group(1).replace(",", ""))
    return d
def parse_payslips(paths):
    agg, n = {}, 0
    for p in paths:
        try:
            t = pdf_text(p)
            if "급여명세" not in t:
                continue
            n += 1
            for k, v in _parse_payslip(t).items():
                agg[k] = agg.get(k, 0) + v
        except Exception:
            pass
    agg["_files"] = n
    return agg

# ---------- 근로 간이지급조서(.pdf) : ⑪과세소득 + 상/하반기 ----------
def parse_geun_gani(path):
    t = pdf_text(path)
    half = "상반기" if "상반기" in t and "√" in t.split("상반기")[0][-4:] else None
    if "[√ ]상반기" in t or "[√]상반기" in t.replace(" ", ""):
        half = "상반기"
    elif "[√ ]하반기" in t or "[√]하반기" in t.replace(" ", ""):
        half = "하반기"
    m = re.search(r"과세소득[^\d]*([\d,]{6,})", t)
    amt = _num(m.group(1)) if m else None
    m2 = re.search(r"근로자총인원\s*([\d,]+)", t)
    return {"반기": half, "과세소득": amt, "인원": _num(m2.group(1)) if m2 else None}

# ---------- 사업 간이지급조서(.pdf) : 총지급액계 ----------
def parse_saeop_gani(path):
    t = pdf_text(path)
    m = re.search(r"총지급액[^\d]*([\d,]{5,})", t) or re.search(r"총 지급액\s*([\d,]{5,})", t)
    if m:
        return _num(m.group(1))
    # fallback: 합계 컬럼 최대값
    nums = [int(x.replace(",", "")) for x in re.findall(r"[\d,]{7,}", t)]
    return max(nums) if nums else None

# ---------- 폴더 스캔 → 월별 통합 ----------
def _find(d, *keys):
    for f in os.listdir(d):
        if all(k in f for k in keys):
            return os.path.join(d, f)
    return None

def scan_year(root, year):
    """root 아래 'MM월' 폴더들을 순회하며 월별 파싱 결과 dict 반환."""
    result = {}
    for m in range(1, 13):
        md = None
        for cand in (f"{m:02d}월", f"{m}월"):
            p = os.path.join(root, cand)
            if os.path.isdir(p):
                md = p; break
        if not md:
            continue
        rec = {}
        # 급여대장
        pf = _find(md, "급여대장", ".xls") or _find(md, "급여대장")
        if pf and pf.endswith((".xls", ".xlsx")):
            try: rec["payroll"] = parse_payroll(pf)
            except Exception: pass
        # 이행상황신고서
        sf = _find(md, "신고") or _find(md, "원천세")
        if sf and sf.endswith(".pdf"):
            try: rec["ihaeng"] = parse_ihaeng(sf)
            except Exception: pass
        # 급여명세서(폴더/개별/zip)
        slips = []
        for r2, _, fs in os.walk(md):
            for f in fs:
                if "급여명세" in f and f.endswith(".pdf"):
                    slips.append(os.path.join(r2, f))
                if f.lower().endswith(".zip"):
                    try:
                        z = zipfile.ZipFile(os.path.join(r2, f))
                        for zn in z.namelist():
                            if zn.endswith(".pdf"):
                                slips.append(io.BytesIO(z.read(zn)))
                    except Exception: pass
        if slips:
            rec["payslips"] = parse_payslips(slips)
        # 간이조서
        gg = _find(md, "근로간이")
        if gg and gg.endswith(".pdf"):
            try: rec["geun_gani"] = parse_geun_gani(gg)
            except Exception: pass
        sg = _find(md, "사업간이")
        if sg and sg.endswith(".pdf"):
            try: rec["saeop_gani"] = parse_saeop_gani(sg)
            except Exception: pass
        result[m] = rec
    return result

if __name__ == "__main__":
    import json, sys
    root = sys.argv[1] if len(sys.argv) > 1 else "/sessions/sharp-gifted-allen/mnt/이노파인더스 원천세대사표 데이터/2026"
    res = scan_year(root, 2026)
    for m, rec in res.items():
        print(f"=== {m}월 ===")
        for k, v in rec.items():
            print("  ", k, ":", v)
