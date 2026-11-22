#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
مولد نموذج دراسة الجدوى العقارية — أكاديمية العقار السعودية
يبني ملف xlsx (5 أوراق RTL بصيغ مترابطة) بالمكتبة القياسية فقط،
ثم يتحقق: سلامة Zip، صلاحية كل XML، وإعادة حساب الأرقام المرجعية يدوياً.
الاستخدام: python3 مولد-النموذج-المالي.py
"""
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from pathlib import Path

OUT = Path(__file__).resolve().parent / "نموذج-دراسة-الجدوى-العقارية.xlsx"

def cell(ref, text=None, num=None, formula=None):
    if text is not None:
        return f'<c r="{ref}" t="inlineStr"><is><t>{escape(text)}</t></is></c>'
    if formula is not None:
        body = formula[1:] if formula.startswith('=') else formula  # علامة = لا تكتب داخل <f> (مواصفة OOXML)
        return f'<c r="{ref}" s="1"><f>{escape(body)}</f></c>'
    return f'<c r="{ref}" s="1"><v>{num}</v></c>'

def row(n, cells):
    return f'<row r="{n}">' + "".join(cells) + "</row>"

def sheet(rows):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
            '<sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="38"/>'
            '<cols><col min="1" max="1" width="48" customWidth="1"/>'
            '<col min="2" max="2" width="24" customWidth="1"/>'
            '<col min="3" max="3" width="34" customWidth="1"/>'
            '<col min="4" max="4" width="68" customWidth="1"/>'
            '<col min="5" max="7" width="24" customWidth="1"/></cols>'
            f'<sheetData>{"".join(rows)}</sheetData>'
            '<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>'
            '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="1"/></worksheet>')

# ---------- ورقة 1: الافتراضات ----------
a = [
    [1, [("A1", "البند"), ("B1", "القيمة"), ("C1", "الوحدة"), ("D1", "ملاحظة")]],
    [2, [("A2", "مساحة الأرض"), ("B2", 1500), ("C2", "م²"), ("D2", "توضيحي: محاكاة 1500م")]],
    [3, [("A3", "سعر الأرض"), ("B3", 2400), ("C3", "ريال/م²"), ("D3", "توضيحي")]],
    [4, [("A4", "نسبة الإشغال"), ("B4", 0.6), ("C4", "-"), ("D4", "تقدير أولي — يُثبّت رسمياً")]],
    [5, [("A5", "عدد الطوابق"), ("B5", 5), ("C5", "-"), ("D5", "")]],
    [6, [("A6", "نسبة المساحة القابلة للبيع"), ("B6", 0.8), ("C6", "-"), ("D6", "يدعم المرافق والجدران")]],
    [7, [("A7", "متوسط مساحة الوحدة"), ("B7", 120), ("C7", "م² صافي"), ("D7", "")]],
    [8, [("A8", "متوسط سعر الوحدة"), ("B8", 500000), ("C8", "ريال"), ("D8", "توضيحي")]],
    [9, [("A9", "تكلفة الإنشاء"), ("B9", 1700), ("C9", "ريال/م² مبني"), ("D9", "تقديري — يُحدَّث")]],
    [10, [("A10", "الأتعاب الهندسية والتراخيص"), ("B10", 250000), ("C10", "ريال"), ("D10", "")]],
    [11, [("A11", "المصاريف الإدارية"), ("B11", 150000), ("C11", "ريال"), ("D11", "")]],
    [12, [("A12", "نسبة التسويق والبيع"), ("B12", 0.015), ("C12", "من الإيراد"), ("D12", "")]],
    [13, [("A13", "احتياطي الطوارئ"), ("B13", 0.05), ("C13", "من الإنشاء"), ("D13", "")]],
    [14, [("A14", "ضريبة شراء الأرض المفترضة"), ("B14", 0.05), ("C14", "من قيمة الأرض"), ("D14", "توضيحي: يتحملها المشروع؛ المسؤولية النظامية في الدليل")]],
    [15, [("A15", "معدل الخصم السنوي"), ("B15", 0.12), ("C15", "-"), ("D15", "")]],
    [16, [("A16", "ضريبة بيع الوحدات إن تحملها المشروع"), ("B16", 0), ("C16", "من إجمالي البيع"), ("D16", "افتراضي: يتحملها المشتري باتفاق مكتوب؛ للاختبار اضبطها 0.05")]],
    [17, [("A17", "نطاق الافتراض أعلاه"), ("C17", "توضيحي"), ("D17", "يؤثر على أثر الضريبة والربح الأساسي بعد الضريبة فقط")]],
    [18, [("A18", ""), ("D18", "لا يغيّر التدفقات أو العوائد أو السيناريوهات")]],
    [19, [("A19", "المسؤولية النظامية"), ("D19", "على المتصرف؛ اتفاق التحمل لا ينقلها وحده")]],
]
s1 = sheet([row(n, [cell(ref, text=v) if isinstance(v, str) else cell(ref, num=v) for ref, v in cells]) for n, cells in a])

# ---------- ورقة 2: التكاليف ----------
c = [
    [1, [("A1", "البند"), ("B1", "القيمة"), ("C1", "الصيغة/المصدر")]],
    [2, [("A2", "قيمة الأرض"), ("B2", "='الافتراضات'!B2*'الافتراضات'!B3"), ("C2", "المساحة × السعر")]],
    [3, [("A3", "ضريبة شراء الأرض المفترضة"), ("B3", "=B2*'الافتراضات'!B14"), ("C3", "")]],
    [4, [("A4", "المساحة المبنية الإجمالية (م²)"), ("B4", "='الافتراضات'!B2*'الافتراضات'!B4*'الافتراضات'!B5"), ("C4", "تقدير أولي — يُثبّت رسمياً")]],
    [5, [("A5", "تكلفة الإنشاء"), ("B5", "=B4*'الافتراضات'!B9"), ("C5", "")]],
    [6, [("A6", "الأتعاب الهندسية والتراخيص"), ("B6", "='الافتراضات'!B10"), ("C6", "")]],
    [7, [("A7", "المصاريف الإدارية"), ("B7", "='الافتراضات'!B11"), ("C7", "")]],
    [8, [("A8", "احتياطي الطوارئ"), ("B8", "=B5*'الافتراضات'!B13"), ("C8", "")]],
]
s2 = sheet([row(n, [cell(ref, text=v) if isinstance(v, str) and not v.startswith("=") else cell(ref, formula=v) for ref, v in cells]) for n, cells in c])

# ---------- ورقة 3: الإيرادات ----------
r3 = [
    [1, [("A1", "البند"), ("B1", "القيمة"), ("C1", "الصيغة/المصدر")]],
    [2, [("A2", "مساحة البيع (م²)"), ("B2", "='التكاليف'!B4*'الافتراضات'!B6"), ("C2", "")]],
    [3, [("A3", "عدد الوحدات"), ("B3", "=IFERROR(ROUNDDOWN(B2/'الافتراضات'!B7,0),0)"), ("C3", "محمي من القسمة على صفر")]],
    [4, [("A4", "متوسط سعر الوحدة"), ("B4", "='الافتراضات'!B8"), ("C4", "")]],
    [5, [("A5", "الإيراد الإجمالي"), ("B5", "=B3*B4"), ("C5", "إجمالي")]],
    [6, [("A6", "مصاريف التسويق والبيع"), ("B6", "=B5*'الافتراضات'!B12"), ("C6", "")]],
    [7, [("A7", "فاقد التحصيل (3%)"), ("B7", "=B5*0.03"), ("C7", "فاقد نهائي، لا تأخير فقط")]],
    [8, [("A8", "الإيراد الصافي"), ("B8", "=B5-B7"), ("C8", "الأساس لكل الحسابات")]],
]
s3 = sheet([row(n, [cell(ref, text=v) if isinstance(v, str) and not v.startswith("=") else cell(ref, formula=v) for ref, v in cells]) for n, cells in r3])

# ---------- ورقة 4: التدفقات ----------
rows4 = [[1, [cell("A1", "الربع"), cell("B1", "تدفق خارج"), cell("C1", "تدفق داخل"), cell("D1", "صافي"), cell("E1", "معامل الخصم"), cell("F1", "صافي مخصوم"), cell("G1", "تراكمي نقدي غير مخصوم")]]]
outs = {
    2: "='التكاليف'!B2+'التكاليف'!B3",
    3: "='التكاليف'!B5*0.3",
    4: "='التكاليف'!B5*0.3+'التكاليف'!B6",
    5: "='التكاليف'!B5*0.4+'التكاليف'!B7+'التكاليف'!B8+'الإيرادات'!B6",
}
for q in range(1, 9):
    r = q + 1
    t = q
    b = outs.get(r, 0)
    ci = "='الإيرادات'!B8*0.25" if r >= 6 else 0
    e = f"=1/POWER(1+(POWER(1+'الافتراضات'!B15,0.25)-1),{t})"
    cells = [
        ("A" + str(r), f"الربع {q}"),
        ("B" + str(r), b if isinstance(b, str) else None, b if not isinstance(b, str) else None),
        ("C" + str(r), ci if isinstance(ci, str) else None, ci if not isinstance(ci, str) else None),
        ("D" + str(r), None, None, f"=C{r}-B{r}"),
        ("E" + str(r), None, None, e),
        ("F" + str(r), None, None, f"=D{r}*E{r}"),
        ("G" + str(r), None, None, (f"=D{r}" if q == 1 else f"=G{r-1}+D{r}")),
    ]
    fmt = []
    for item in cells:
        ref = item[0]
        if len(item) == 2:
            fmt.append(cell(ref, text=item[1]))
        elif len(item) == 3:
            fmt.append(cell(ref, formula=item[1]) if item[1] is not None else cell(ref, num=item[2]))
        else:
            fmt.append(cell(ref, formula=item[3]))
    rows4.append([r, fmt])
rows4.append([10, [cell("A10", "صافي القيمة الحالية قبل ضريبة البيع"), cell("B10", formula="=SUM(F2:F9)")]])
rows4.append([11, [cell("A11", "العائد الداخلي الربع سنوي قبل ضريبة البيع"), cell("B11", formula="=IRR(D2:D9)"), cell("C11", "يتطلب تدفقاً سالباً ثم إيجابياً")]])
rows4.append([12, [cell("A12", "العائد الداخلي السنوي قبل ضريبة البيع"), cell("B12", formula="=(1+B11)^4-1"), cell("C12", "المؤشر المعروض في النتائج")]])
rows4.append([14, [cell("A14", "ملاحظة: كل نتائج هذه الورقة قبل ضريبة بيع الوحدات")]])
rows4.append([15, [cell("A15", "ولا تتأثر بافتراض ضريبة البيع في ورقة الافتراضات")]])
rows4.append([16, [cell("A16", "عند تبنّي السيناريو تُعاد حسابات التوقيت والذروة والخصم منفصلة")]])
s4 = sheet([row(n, cells) for n, cells in rows4])

# ---------- ورقة 5: النتائج والسيناريوهات ----------
r5 = [
    [2, [("A2", "الإيراد الصافي"), ("B2", "='الإيرادات'!B8")]],
    [3, [("A3", "إجمالي التكاليف"), ("B3", "='التكاليف'!B2+'التكاليف'!B3+'التكاليف'!B5+'التكاليف'!B6+'التكاليف'!B7+'التكاليف'!B8+'الإيرادات'!B6")]],
    [4, [("A4", "الربح قبل ضريبة بيع الوحدات (الإيراد الصافي − التكاليف)"), ("B4", "=B2-B3")]],
    [5, [("A5", "الهامش على التكاليف"), ("B5", '=IFERROR(B4/B3,"—")')]],
    [6, [("A6", "صافي القيمة الحالية قبل ضريبة البيع"), ("B6", "='التدفقات'!B10")]],
    [7, [("A7", "العائد الداخلي السنوي قبل ضريبة البيع"), ("B7", "='التدفقات'!B12")]],
    [8, [("A8", "السيناريوهات"), ("B8", "متحفظ"), ("C8", "أساسي"), ("D8", "متفائل")]],
    [9, [("A9", "تعديل سعر البيع"), ("B9", -0.08), ("C9", 0), ("D9", 0.06)]],
    [10, [("A10", "تعديل تكلفة الإنشاء"), ("B10", 0.08), ("C10", 0), ("D10", -0.05)]],
    [11, [("A11", "الإيراد الصافي المعدل"), ("B11", "='الإيرادات'!B8*(1+B9)"), ("C11", "='الإيرادات'!B8*(1+C9)"), ("D11", "='الإيرادات'!B8*(1+D9)")]],
    [12, [("A12", "الإنشاء المعدل"), ("B12", "='التكاليف'!B5*(1+B10)"), ("C12", "='التكاليف'!B5*(1+C10)"), ("D12", "='التكاليف'!B5*(1+D10)")]],
    [13, [("A13", "التسويق المعدل"), ("B13", "='الإيرادات'!B6*(1+B9)"), ("C13", "='الإيرادات'!B6*(1+C9)"), ("D13", "='الإيرادات'!B6*(1+D9)")]],
    [14, [("A14", "الاحتياطي المعدل"), ("B14", "=B12*'الافتراضات'!B13"), ("C14", "=C12*'الافتراضات'!B13"), ("D14", "=D12*'الافتراضات'!B13")]],
    [15, [("A15", "ثابتات (أرض+رسم+هندسة+إداري)"), ("B15", "='التكاليف'!B2+'التكاليف'!B3+'التكاليف'!B6+'التكاليف'!B7"), ("C15", "=B15"), ("D15", "=B15")]],
    [16, [("A16", "إجمالي التكاليف المعدل"), ("B16", "=B15+B12+B14+B13"), ("C16", "=C15+C12+C14+C13"), ("D16", "=D15+D12+D14+D13")]],
    [17, [("A17", "الربح المعدل (قبل ضريبة بيع الوحدات)"), ("B17", "=B11-B16"), ("C17", "=C11-C16"), ("D17", "=D11-D16")]],
    [18, [("A18", "الهامش المعدل (قبل ضريبة بيع الوحدات)"), ("B18", '=IFERROR(B17/B16,"—")'), ("C18", '=IFERROR(C17/C16,"—")'), ("D18", '=IFERROR(D17/D16,"—")')]],
    [19, [("A19", "أثر توضيحي: ضريبة البيع إن تحملها المشروع"), ("B19", "=-'الإيرادات'!B5*'الافتراضات'!B16"), ("C19", "صفر ما دام الافتراض أن المشتري يتحملها")]],
    [20, [("A20", "الربح بعد ضريبة بيع الوحدات"), ("B20", "=B4+B19"), ("C20", "يطابق الربح قبل الضريبة ما دام الافتراض صفراً")]],
    [21, [("A21", "تنبيه: أثر الضريبة توضيحي على الربح الأساسي فقط")]],
    [22, [("A22", "التراكم والعوائد والسيناريوهات كلها قبل ضريبة البيع")]],
    [23, [("A23", "عند تبنّي السيناريو تُعاد حسابات التوقيت منفصلة")]],
]
s5 = sheet([row(n, [cell(ref, text=v) if isinstance(v, str) and not v.startswith("=") else cell(ref, formula=v) if isinstance(v, str) else cell(ref, num=v) for ref, v in cells]) for n, cells in r5])

# Numeric formats remain presentation only; formulas and assumptions are unchanged.
import re

def percent_style(xml, refs):
    for ref in refs:
        xml = xml.replace(f'<c r="{ref}" s="1">', f'<c r="{ref}" s="2">')
    return xml

s1 = percent_style(s1, ['B4', 'B6', 'B12', 'B13', 'B14', 'B15', 'B16'])
s4 = percent_style(s4, ['B11', 'B12'])
s4 = s4.replace('width="48"', 'width="65"').replace('width="68"', 'width="24"')
s5 = percent_style(s5, ['B5', 'B7'] + [f'{c}{r}' for c in 'BCD' for r in [9, 10, 18]])

# ---------- تجميع الحزمة ----------
SHEETS = [("الافتراضات", s1), ("التكاليف", s2), ("الإيرادات", s3), ("التدفقات", s4), ("النتائج والسيناريوهات", s5)]
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, 6))
      + '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
WB = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
      + "".join(f'<sheet name="{escape(n)}" sheetId="{i}" r:id="rId{i}"/>' for i, (n, _) in enumerate(SHEETS, 1))
      + '</sheets></workbook>')
WBRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
          + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, 6))
          + '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
          '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
          '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
          '<borders count="1"><border/></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          '<cellXfs count="3">'
          '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center" wrapText="1"/></xf>'
          '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="center" wrapText="1"/></xf>'
          '<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="center" wrapText="1"/></xf>'
          '</cellXfs></styleSheet>')

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", RELS)
    z.writestr("xl/workbook.xml", WB)
    z.writestr("xl/_rels/workbook.xml.rels", WBRELS)
    z.writestr("xl/styles.xml", STYLES)
    for i, (_, sx) in enumerate(SHEETS, 1):
        z.writestr(f"xl/worksheets/sheet{i}.xml", sx)

# ---------- التحقق ----------
with zipfile.ZipFile(OUT) as z:
    assert z.testzip() is None, "تلف في الحزمة"
    for name in z.namelist():
        ET.fromstring(z.read(name))
print(f"[ok] {OUT} — {len(SHEETS)} أوراق، XML سليم")

# الأرقام المرجعية المحسوبة يدوياً (للمقابلة مع دليل النموذج)
land, transfer = 1500 * 2400, 0.05 * 1500 * 2400
built, con = 1500 * 0.6 * 5, 1500 * 0.6 * 5 * 1700
eng, adm, reserve = 250000, 150000, 1500 * 0.6 * 5 * 1700 * 0.05
fixed = land + transfer + eng + adm
units = int((built * 0.8) // 120)
rev = units * 500000
mkt, loss = rev * 0.015, rev * 0.03
net = rev - loss
total = land + transfer + con + eng + adm + reserve + mkt
profit, margin = net - total, (net - total) / total
rq = (1.12) ** 0.25 - 1
outs_ = [land + transfer, con * 0.3, con * 0.3 + eng, con * 0.4 + adm + reserve + mkt]
ins_ = [net / 4] * 4
flows = [-o for o in outs_] + ins_
npv = sum(cf / (1 + rq) ** (i + 1) for i, cf in enumerate(flows))
def _irr(vals):
    lo, hi = -0.9999, 10.0
    f = lambda r: sum(v / (1 + r) ** i for i, v in enumerate(vals))
    flo = f(lo)
    for _ in range(200):
        mid = (lo + hi) / 2
        fm = f(mid)
        if flo * fm <= 0:
            hi = mid
        else:
            lo, flo = mid, fm
    return (lo + hi) / 2
irr_q = _irr(flows)
irr_a = (1 + irr_q) ** 4 - 1
def scen(dp, dc):
    rr, cc = rev * (1 + dp) * 0.97, con * (1 + dc)
    mk, rs = mkt * (1 + dp), cc * 0.05
    tot = fixed + cc + rs + mk
    pr = rr - tot
    return pr, pr / tot
print(f"إجمالي التكاليف = {total:,.0f} | الربح = {profit:,.0f} | الهامش = {margin:.1%} | NPV = {npv:,.0f}")
print(f"IRR ربع سنوي = {irr_q:.2%} | IRR سنوي = {irr_a:.2%}")
for name, dp, dc in [("متحفظ", -0.08, 0.08), ("متفائل", 0.06, -0.05)]:
    p, m = scen(dp, dc)
    print(f"{name}: ربح = {p:,.0f} | هامش = {m:.1%}")
sale_tax = rev * 0.05
print(f"أثر توضيحي لو تحمل المشروع ضريبة بيع الوحدات 5%: أساسي {profit - sale_tax:,.0f} | متحفظ {scen(-0.08, 0.08)[0] - rev * 0.92 * 0.05:,.0f}")
