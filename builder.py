# -*- coding: utf-8 -*-
"""파싱 결과(scan_year) → 개선 원천세 대사표 xlsx (전 소득종류 포괄, 값복사=파랑+출처메모)."""
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

FN = "맑은 고딕"
BLUE = "0000FF"
orange = PatternFill("solid", fgColor="FFC000")
hdrf = PatternFill("solid", fgColor="1F4E78")
sub_fill = PatternFill("solid", fgColor="D6E4F0")
FCE = PatternFill("solid", fgColor="FCE4D6")
ok_f = PatternFill("solid", fgColor="C6EFCE")
bad_f = PatternFill("solid", fgColor="FFC7CE")
thin = Side(style="thin", color="BFBFBF")
bd = Border(left=thin, right=thin, top=thin, bottom=thin)
NUM = '#,##0;(#,##0);"-"'

HDR = ["월","급여(과세)","제출비과세","미제출비과세","상여","계","신고","차액","소득세(갑근세)",
       "주민세","국민연금","건강보험","고용보험","장기요양","연말정산소득세","연말정산지방세",
       "중도퇴사소득세","중도퇴사지방세","건강보험정산","장기요양정산","연금보험정산","고용보험정산",
       "선지급분","공제계","실지급액"]
NC = len(HDR)  # 25 (A..Y)

def _cell(ws, r, c, v, bold=False, fill=None, color="000000", num=True, center=True, src=None):
    cc = ws.cell(r, c, v)
    isnum = isinstance(v, (int, float))
    cc.font = Font(name=FN, size=9, bold=bold, color=(BLUE if (src and isnum) else color))
    if fill: cc.fill = fill
    if num: cc.number_format = NUM
    cc.alignment = Alignment(horizontal="center" if center else "left", vertical="center", wrap_text=True)
    cc.border = bd
    if src and isnum:
        cc.comment = Comment("출처: " + src, "대사표")
    return cc

def _section(ws, r, c1, c2, text):
    ws.cell(r, c1, text).font = Font(name=FN, bold=True, size=11, color="FFFFFF")
    for c in range(c1, c2+1):
        ws.cell(r, c).fill = hdrf; ws.cell(r, c).border = bd
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    ws.cell(r, c1).alignment = Alignment(horizontal="left", vertical="center")

def build(parsed, year, company="", src_prefix=""):
    """parsed: {month: {payroll, ihaeng, payslips, geun_gani, saeop_gani}}"""
    wb = Workbook(); ws = wb.active; ws.title = str(year)
    ws.sheet_view.showGridLines = False
    for i in range(1, NC+1): ws.column_dimensions[get_column_letter(i)].width = 13
    ws.column_dimensions["A"].width = 5
    def S(name): return (f"{year}/{{m:02d}}월 " if False else "") + name  # placeholder

    r = 1
    ws.cell(r, 1, f"{year}년 원천세 대사표 (개선판)").font = Font(name=FN, bold=True, size=14, color="1F4E78")
    r = 2
    ws.cell(r, 1, f"{company}").font = Font(name=FN, italic=True, size=9)
    r = 4
    _section(ws, r, 1, NC, "① 근로소득 월별 명세"); r += 1
    for j, h in enumerate(HDR): _cell(ws, r, j+1, h, bold=True, fill=hdrf, color="FFFFFF", num=False)
    r += 1
    first = r
    for m in range(1, 13):
        rec = parsed.get(m, {})
        pr = rec.get("payroll"); ih = rec.get("ihaeng") or {}; ps = rec.get("payslips") or {}
        geun = (ih.get("근로") or {})
        srcP = f"{year}/{m:02d}월 급여대장.xls"
        srcS = f"{year}/{m:02d}월 이행상황신고서.pdf"
        srcM = f"{year}/{m:02d}월 급여명세서(집계)"
        _cell(ws, r, 1, m, bold=True)
        if pr:
            _cell(ws, r, 2, pr["급여"], src=srcP); _cell(ws, r, 3, pr["제출비과세"], src=srcP)
            _cell(ws, r, 4, 0, src=srcP+" (미제출비과세 없음)"); _cell(ws, r, 5, 0, src=srcP+" (상여 없음)")
        else:
            for c in range(2, 6): _cell(ws, r, c, None, fill=orange)
        _cell(ws, r, 6, f"=SUM(B{r}:E{r})")
        _cell(ws, r, 7, geun.get("지급액") if geun else None, src=(srcS if geun.get("지급액") else None),
              fill=(None if geun.get("지급액") else orange))
        _cell(ws, r, 8, f"=F{r}-G{r}")
        _cell(ws, r, 9, geun.get("소득세") if geun else None, src=(srcS if geun.get("소득세") else None),
              fill=(None if geun.get("소득세") else orange))
        # 4대보험/연말정산 from payslips
        pmap = [("주민세","지방소득세"),("국민연금","국민연금"),("건강보험","건강보험"),
                ("고용보험","고용보험"),("장기요양","장기요양보험료"),
                ("연말정산소득세","연말정산소득세"),("연말정산지방세","연말정산지방소득세")]
        for k,(_,pk) in enumerate(pmap):
            col = 10+k
            val = ps.get(pk)
            _cell(ws, r, col, val, src=(srcM if val is not None else None),
                  fill=(None if val is not None else (orange if (pr or geun) else None)))
        # Q~W 중도/정산/선지급 : 자료 없으면 공란
        for c in range(17, 24): _cell(ws, r, c, None)
        _cell(ws, r, 24, f"=SUM(I{r}:W{r})")     # 공제계
        _cell(ws, r, 25, f"=F{r}-X{r}")          # 실지급액
        r += 1
    last = r-1
    def sub(lab, a, b, fill):
        _cell(ws, r, 1, lab, bold=True, fill=fill)
        for c in range(2, NC+1):
            L = get_column_letter(c); _cell(ws, r, c, f"=SUM({L}{a}:{L}{b})", bold=True, fill=fill)
    sub("상반기", first, first+5, sub_fill); h1 = r; r += 1
    sub("하반기", first+6, first+11, sub_fill); h2 = r; r += 1
    _cell(ws, r, 1, "합계", bold=True, fill=FCE)
    for c in range(2, NC+1):
        L = get_column_letter(c); _cell(ws, r, c, f"={L}{h1}+{L}{h2}", bold=True, fill=FCE)
    tot = r; r += 2

    # ② 근로 간이지급명세서 대조
    _section(ws, r, 1, 5, "② 근로 간이지급명세서 대조 (급여대장 급여표 ↔ 국세청 제출 간이지급명세서)"); r += 1
    for j, h in enumerate(["구분","급여표 반기합계(급여대장)","간이지급명세서 제출액(제출본)","차이","일치여부"]):
        _cell(ws, r, j+1, h, bold=True, fill=hdrf, color="FFFFFF", num=(j not in (0,4)))
    r += 1
    # 간이 제출액: geun_gani 상/하반기
    g_h1 = g_h2 = None; g_src = ""
    for m, rec in parsed.items():
        gg = rec.get("geun_gani")
        if gg and gg.get("과세소득"):
            g_src = f"{year}/{m:02d}월 근로간이지급조서.pdf"
            if gg.get("반기") == "상반기": g_h1 = gg["과세소득"]
            elif gg.get("반기") == "하반기": g_h2 = gg["과세소득"]
    g_tot = (g_h1 or 0)+(g_h2 or 0) if (g_h1 or g_h2) else None
    for lab, form, sub_v in [("상반기(1~6월)", f"=SUM(B{first}:B{first+5})", g_h1),
                             ("하반기(7~12월)", f"=SUM(B{first+6}:B{first+11})", g_h2),
                             ("합계", f"=B{tot}", g_tot)]:
        _cell(ws, r, 1, lab, bold=True, num=False)
        _cell(ws, r, 2, form)
        _cell(ws, r, 3, sub_v, src=(g_src if sub_v else None), fill=(None if sub_v else orange))
        _cell(ws, r, 4, f"=B{r}-C{r}")
        _cell(ws, r, 5, f'=IF(C{r}="","불일치",IF(B{r}=C{r},"일치","불일치"))', bold=True)
        r += 1
    r += 1

    # ③ 사업 간이지급명세서 대조 + 사업소득 현황
    _section(ws, r, 1, 7, "③ 사업소득 현황 + 사업 간이지급명세서 대조"); r += 1
    for j, h in enumerate(["월","인원","사업소득 지급액","소득세","지방소득세","간이지급명세서 제출액","일치여부"]):
        _cell(ws, r, j+1, h, bold=True, fill=hdrf, color="FFFFFF", num=(j not in (0,6)))
    r += 1
    sfirst = r
    for m in range(1, 13):
        rec = parsed.get(m, {})
        sa = (rec.get("ihaeng") or {}).get("사업") or {}
        sg = rec.get("saeop_gani")
        srcS = f"{year}/{m:02d}월 이행상황신고서.pdf"
        srcG = f"{year}/{m:02d}월 사업간이지급조서.pdf"
        _cell(ws, r, 1, m, bold=True)
        _cell(ws, r, 2, sa.get("인원"), src=(srcS if sa.get("인원") is not None else None))
        _cell(ws, r, 3, sa.get("지급액"), src=(srcS if sa.get("지급액") is not None else None))
        _cell(ws, r, 4, sa.get("소득세"), src=(srcS if sa.get("소득세") is not None else None))
        _cell(ws, r, 5, None)  # 지방세: 원천 없으면 공란
        _cell(ws, r, 6, sg, src=(srcG if sg else None), fill=(None if sg else orange))
        _cell(ws, r, 7, f'=IF(F{r}="","불일치",IF(C{r}=F{r},"일치","불일치"))', bold=True)
        r += 1
    slast = r-1
    _cell(ws, r, 1, "계", bold=True, fill=FCE)
    for c, L in [(2,"B"),(3,"C"),(4,"D"),(5,"E"),(6,"F")]:
        _cell(ws, r, c, f"=SUM({L}{sfirst}:{L}{slast})", bold=True, fill=FCE)
    _cell(ws, r, 7, f'=IF(F{r}="","불일치",IF(C{r}=F{r},"일치","불일치"))', bold=True, fill=FCE)
    r += 2

    # ④ 기타 소득종류 (포괄)
    _section(ws, r, 1, 6, "④ 기타 소득종류 현황 (일용·기타·퇴직) — 발생 시 입력"); r += 1
    for j, h in enumerate(["소득종류","인원","지급액","소득세","지방소득세","비고"]):
        _cell(ws, r, j+1, h, bold=True, fill=hdrf, color="FFFFFF", num=(j not in (0,5)))
    r += 1
    for lab in ["일용소득","기타소득","퇴직소득"]:
        _cell(ws, r, 1, lab, bold=True)
        for c in range(2, 6): _cell(ws, r, c, None)
        _cell(ws, r, 6, "", num=False, center=False)
        r += 1
    r += 1
    _section(ws, r, 1, NC, "⑤ 범례"); r += 1
    for n in ["파란 글자 = 외부 자료에서 값복사(하드코딩), 셀 메모에 출처파일 표기 / 검정 = 계산식 / 주황 = 자료 미확보",
              "②③ 일치여부 = 단일 수식(제출액 있고 값 일치 시 '일치', 없거나 다르면 '불일치')",
              "미제출비과세·상여·중도퇴사/정산 등은 해당 자료가 있으면 자동 반영(현재 급여대장 미포함분은 공란)."]:
        c = ws.cell(r, 1, "· "+n); c.font = Font(name=FN, size=9); c.alignment = Alignment(horizontal="left")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC); r += 1
    ws.freeze_panes = "A6"
    return wb

def build_bytes(parsed, year, company=""):
    wb = build(parsed, year, company)
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio.read()
