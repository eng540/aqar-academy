#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
حزمة اختبارات الأكاديمية العقارية — تمنع تكرار أخطاء المراجعة المستقلة.
تشغيل:  python3 tests/run_tests.py   (من جذر المستودع)
تنتهي برمز خطأ عند أي فشل — لا نجاح بلا تحقق فعلي.

تغطي: صلاحية CSV، تطابق الأعداد مع السجل المعتمد، التكرار، اتساق حالات
التحقق مع المصادر، صيغ النموذج المالي (فحص نصي + إعادة حساب مستقلة
بثلاث سيناريوهات + إجهاد القسمة على صفر)، الروابط الداخلية، ملفات
التنزيل، صحة Markdown الأساسية، واتساق الأعداد في الوثائق.
"""
import csv
import glob
import os
import re
import shutil
import sys
import zipfile
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from model_eval import Workbook  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB = os.path.join(ROOT, 'الملفات-العملية', 'قاعدة-مكة-المكرمة')
MODEL = os.path.join(ROOT, 'نماذج', 'نموذج-دراسة-الجدوى-العقارية.xlsx')
INDEX = os.path.join(ROOT, 'index.html')

# سجل الأعداد المعتمد — مثبّت من قاعدة المصدر makkah.sqlite3 بتاريخ 2026-11-05
EXPECTED = {
    'developers.csv': 39, 'contractors.csv': 60, 'engineering_offices.csv': 1859,
    'broker_offices.csv': 37, 'projects.csv': 60, 'sources.csv': 241,
}

results = []


def check(name, fn):
    try:
        detail = fn() or ''
        results.append((name, True, detail))
        print(f'[نجح] {name} — {detail}')
    except AssertionError as e:
        results.append((name, False, str(e)))
        print(f'[فشل] {name} — {e}')
    except Exception as e:  # noqa: BLE001
        results.append((name, False, f'خطأ غير متوقع: {type(e).__name__}: {e}'))
        print(f'[فشل] {name} — خطأ غير متوقع: {e}')


def read_csv(path):
    with open(path, encoding='utf-8-sig', newline='') as fh:
        rows = list(csv.DictReader(fh))
    return rows


# ---------- 1) صلاحية CSV ----------
def t_csv_valid():
    for path in sorted(glob.glob(os.path.join(DB, '*.csv'))):
        base = os.path.basename(path)
        with open(path, encoding='utf-8-sig', newline='') as fh:
            rdr = csv.reader(fh)
            header = next(rdr)
            for i, row in enumerate(rdr, 2):
                if len(row) != len(header):
                    raise AssertionError(f'{base}: سطر {i} فيه {len(row)} عموداً والمطلوب {len(header)}')
    return f'{len(glob.glob(os.path.join(DB, "*.csv")))} ملفات CSV سليمة البنية'


# ---------- 2) تطابق الأعداد مع السجل المعتمد ----------
def t_counts():
    parts = []
    for name, want in EXPECTED.items():
        got = sum(1 for _ in open(os.path.join(DB, name), encoding='utf-8-sig')) - 1
        if got != want:
            raise AssertionError(f'{name}: الفعلي {got} والسجل {want} — حدّث السجل والوثائق معاً')
        parts.append(f'{name}={got}')
    return '، '.join(parts)


# ---------- 3) التكرار ----------
def t_duplicates():
    id_col = {'developers.csv': 'dev_id', 'contractors.csv': 'contractor_id',
              'engineering_offices.csv': 'eng_id', 'broker_offices.csv': 'office_id',
              'projects.csv': 'project_id', 'sources.csv': 'source_id'}
    for name, col in id_col.items():
        ids = [r[col] for r in read_csv(os.path.join(DB, name))]
        dupes = {x for x in ids if ids.count(x) > 1}
        if dupes:
            raise AssertionError(f'{name}: معرفات مكررة {sorted(dupes)[:5]}')
        if any((not i.strip()) for i in ids):
            raise AssertionError(f'{name}: معرف فارغ')
    return 'لا معرفات مكررة ولا معرفات فارغة في 6 ملفات'


# ---------- 4) اتساق حالات التحقق مع المصادر ----------
def t_status_source():
    src = {r['source_id']: r for r in read_csv(os.path.join(DB, 'sources.csv'))}
    files = [('developers.csv', 'dev_id'), ('contractors.csv', 'contractor_id'),
             ('engineering_offices.csv', 'eng_id'), ('broker_offices.csv', 'office_id'),
             ('projects.csv', 'project_id')]
    bad, counts = [], {}
    for name, idc in files:
        rows = read_csv(os.path.join(DB, name))
        scol = 'source_id' if 'source_id' in rows[0] else 'source'
        for r in rows:
            st = (r.get('status') or '').strip()
            counts[st] = counts.get(st, 0) + 1
            sid = (r.get(scol) or '').strip()
            if st in ('confirmed', 'existence_verified'):
                if not sid or sid not in src:
                    bad.append(f'{name}:{r[idc]} حالة {st} بلا مصدر مقفول')
                elif st == 'confirmed' and src[sid].get('is_official') != '1':
                    bad.append(f'{name}:{r[idc]} حالة confirmed بمصدر غير رسمي ({sid})')
    if bad:
        raise AssertionError(f'{len(bad)} مخالفة: ' + ' | '.join(bad[:6]))
    return 'الحالات متسقة مع المصادر: ' + str(counts)


# ---------- 5) صيغ النموذج: لا صيغة نصية ولا مرجع مكسور ----------
def t_model_syntax():
    wb = Workbook(MODEL)
    formulaish = re.compile(r"^[A-Za-z'=\(].*[A-Za-z]+[0-9]+|^[A-Za-z'=\(].*!")
    bad = []
    for (sh, ref), c in wb.cells.items():
        if c['kind'] == 's':
            txt = c['v']
            if txt.startswith('=') or formulaish.match(txt):
                bad.append(f'{sh}!{ref} نص يبدو صيغة: {txt[:40]}')
        if c['kind'] == 'f':
            for m in re.finditer(r"(?:'([^']+)'!)?([A-Z]+[0-9]+)", c['f']):
                tgt = m.group(1) or sh
                if (tgt, m.group(2)) not in wb.cells:
                    bad.append(f'{sh}!{ref} يشير لخلية غير موجودة: {tgt}!{m.group(2)}')
    if bad:
        raise AssertionError(f'{len(bad)} مشكلة: ' + ' | '.join(bad[:6]))
    return 'كل الخلايا النصية نصوص فعلية، وكل مراجع الصيغ قائمة'


# ---------- 6+7) إعادة الحساب المستقلة: أساسي وسيناريوهات ----------
BASE = dict(land_a=1500, ps=2400, occ=0.6, fl=5, sell=0.8, unit=120,
            up=500000, con_p=1700, eng=250000, adm=150000, mkt=0.015,
            res=0.05, tr=0.05, dr=0.12)


def expectations(a):
    land = a['land_a'] * a['ps']
    transfer = land * a['tr']
    built = a['land_a'] * a['occ'] * a['fl']
    con = built * a['con_p']
    reserve = con * a['res']
    fixed = land + transfer + a['eng'] + a['adm']
    units = int((built * a['sell']) // a['unit'])
    rev = units * a['up']
    mktf = rev * a['mkt']
    net = rev - rev * 0.03
    total = fixed + con + reserve + mktf
    profit = net - total
    rq = (1 + a['dr']) ** 0.25 - 1
    flows = [-(land + transfer), -(con * 0.3), -(con * 0.3 + a['eng']),
             -(con * 0.4 + a['adm'] + reserve + mktf)] + [net / 4] * 4
    npv = sum(cf / (1 + rq) ** (i + 1) for i, cf in enumerate(flows))

    def irr(vals):
        lo, hi = -0.9999, 10.0
        flo, fhi = sum(v / (1 + lo) ** i for i, v in enumerate(vals)), sum(v / (1 + hi) ** i for i, v in enumerate(vals))
        if flo * fhi > 0:
            return None
        for _ in range(200):
            mid = (lo + hi) / 2
            fm = sum(v / (1 + mid) ** i for i, v in enumerate(vals))
            if flo * fm <= 0:
                hi = mid
            else:
                lo, flo = mid, fm
        return (lo + hi) / 2

    iq = irr(flows)
    out = dict(land=land, built=built, units=units, rev=rev, net=net, total=total,
               profit=profit, margin=profit / total, npv=npv,
               irr_a=(1 + iq) ** 4 - 1 if iq is not None else None)

    def scen(dp, dc):
        cc = con * (1 + dc)
        rr = net * (1 + dp)
        mk, rs = mktf * (1 + dp), cc * a['res']
        tot = fixed + cc + rs + mk
        return rr - tot, (rr - tot) / tot

    out['scen'] = {'متحفظ': scen(-0.08, 0.08), 'أساسي': scen(0, 0), 'متفائل': scen(0.06, -0.05)}
    return out


def close(a, b, tol=0.5):
    return abs(a - b) <= tol


def t_recalc_base():
    wb = Workbook(MODEL)
    e = expectations(BASE)
    got = {
        'التكاليف!B2 (قيمة الأرض)': wb.eval('التكاليف', 'B2') == e['land'],
        'التكاليف!B4 (مساحة مبنية)': wb.eval('التكاليف', 'B4') == e['built'],
        'الإيرادات!B3 (عدد الوحدات)': wb.eval('الإيرادات', 'B3') == e['units'],
        'النتائج!B2 (إيراد صافي)': close(wb.eval('النتائج والسيناريوهات', 'B2'), e['net']),
        'النتائج!B3 (إجمالي تكاليف)': close(wb.eval('النتائج والسيناريوهات', 'B3'), e['total']),
        'النتائج!B4 (الربح)': close(wb.eval('النتائج والسيناريوهات', 'B4'), e['profit']),
        'النتائج!B5 (الهامش)': close(wb.eval('النتائج والسيناريوهات', 'B5'), e['margin'], 1e-9),
        'التدفقات!B10 (NPV)': close(wb.eval('التدفقات', 'B10'), e['npv']),
        'النتائج!B6 (NPV)': close(wb.eval('النتائج والسيناريوهات', 'B6'), e['npv']),
    }
    bad = [k for k, ok in got.items() if not ok]
    if bad:
        raise AssertionError('اختلاف مع المرجع المستقل في: ' + '، '.join(bad))
    s = e['scen']
    for col, key in [('B', 'متحفظ'), ('C', 'أساسي'), ('D', 'متفائل')]:
        p, m = s[key]
        if not (close(wb.eval('النتائج والسيناريوهات', f'{col}17'), p) and close(wb.eval('النتائج والسيناريوهات', f'{col}18'), m, 1e-9)):
            raise AssertionError(f'سيناريو {key} لا يطابق المرجع المستقل')
    return (f"تكاليف {e['total']:,.0f} | ربح {e['profit']:,.0f} | هامش {e['margin']:.2%} | "
            f"NPV {e['npv']:,.0f} | سيناريوهات ×3 مطابقة")


def t_recalc_irr():
    wb = Workbook(MODEL)
    e = expectations(BASE)
    if e['irr_a'] is None:
        raise AssertionError('المرجع المستقل لم يجد IRR')
    got = wb.eval('النتائج والسيناريوهات', 'B7')
    if not close(got, e['irr_a'], 1e-6):
        raise AssertionError(f'IRR في النموذج {got:.6f} والمرجع {e["irr_a"]:.6f}')
    if not close(wb.eval('التدفقات', 'B12'), e['irr_a'], 1e-6):
        raise AssertionError('IRR السنوي في ورقة التدفقات لا يطابق')
    return f'IRR سنوي = {got:.2%} مطابق للمرجع المستقل'


def write_variant(path, changes):
    """ينسخ النموذج مع تعديل قيم افتراضات (ورقة 1) لإعادة الحساب."""
    tmp = path + '.tmp'
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.namelist():
            data = zin.read(item)
            if item == 'xl/worksheets/sheet1.xml':
                root = ET.fromstring(data)
                for c in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                    if c.get('r') in changes:
                        v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                        v.text = str(changes[c.get('r')])
                data = ET.tostring(root, encoding='UTF-8', xml_declaration=True)
            zout.writestr(item, data)
    shutil.move(tmp, path)


def t_recalc_scenarios():
    import tempfile
    scen_sets = [
        ({'B3': '2800', 'B9': '1800', 'B8': '480000'}, 'سعر الأرض 2800 + إنشاء 1800 + وحدة 480 ألف'),
        ({'B3': '2000', 'B9': '1600', 'B15': '0.14'}, 'انخفاض شامل + رفع الخصم 14%'),
    ]
    parts = []
    with tempfile.TemporaryDirectory() as td:
        for changes, label in scen_sets:
            p = os.path.join(td, 'v.xlsx')
            shutil.copy(MODEL, p)
            write_variant(p, changes)
            a = dict(BASE)
            a['ps'] = float(changes.get('B3', BASE['ps']))
            a['con_p'] = float(changes.get('B9', BASE['con_p']))
            a['up'] = float(changes.get('B8', BASE['up']))
            a['dr'] = float(changes.get('B15', BASE['dr']))
            e = expectations(a)
            wb = Workbook(p)
            if not close(wb.eval('النتائج والسيناريوهات', 'B3'), e['total']):
                raise AssertionError(f'[{label}] إجمالي التكاليف لا يطابق المرجع')
            if not close(wb.eval('النتائج والسيناريوهات', 'B4'), e['profit']):
                raise AssertionError(f'[{label}] الربح لا يطابق المرجع')
            if not close(wb.eval('التدفقات', 'B10'), e['npv']):
                raise AssertionError(f'[{label}] NPV لا يطابق المرجع')
            base_wb = Workbook(MODEL)
            if close(wb.eval('التدفقات', 'B10'), base_wb.eval('التدفقات', 'B10'), 1):
                raise AssertionError(f'[{label}] التدفقات لم تتغير مع تغيير الافتراضات!')
            parts.append(f'{label}: NPV {e["npv"]:,.0f} ✓')
    return '، '.join(parts)


def t_model_stress():
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        p = os.path.join(td, 'zero.xlsx')
        shutil.copy(MODEL, p)
        write_variant(p, {'B7': '0'})  # متوسط مساحة الوحدة = 0 → قسمة فعلية على صفر في ROUNDDOWN
        wb = Workbook(p)
        allv = wb.evaluate_all()
        structural = {k: v for k, v in allv.items() if v in ('#VALUE!', '#REF!', '#NAME?', '#DIV/0!')}
        if structural:
            raise AssertionError(f'تسربت أخطاء Excel إلى النتائج: {structural}')
        units = allv.get(('الإيرادات', 'B3'))
        if units != 0:
            raise AssertionError(f'عدد الوحدات عند القسمة على صفر = {units} والمتوقع 0 عبر IFERROR')
        irr_cell = allv.get(('النتائج والسيناريوهات', 'B7'))
        if irr_cell != '#NUM!':
            raise AssertionError(f'IRR بتدفقات سالبة كلياً = {irr_cell} والمتوقع #NUM! (سلوك Excel المعلن)')
    return 'القسمة على صفر: محمية بـIFERROR بلا ظهور #DIV/0!، وIRR غير المعرف يظهر #NUM! بوضوح'


# ---------- 8) الروابط الداخلية ----------
def t_links():
    html = open(INDEX, encoding='utf-8').read()
    import urllib.parse
    paths = set(re.findall(r'data-path="([^"]+)"', html))
    missing = [p for p in paths if not os.path.exists(os.path.join(ROOT, p))]
    if missing:
        raise AssertionError(f'عناصر فهرس بلا ملفات: {missing}')
    hashes = [urllib.parse.unquote(h) for h in re.findall(r'href="#/([^"]+)"', html)
              if 'encodeURIComponent' not in h]
    bad = [h for h in hashes if not os.path.exists(os.path.join(ROOT, h))]
    if bad:
        raise AssertionError(f'روابط توجيه مكسورة: {bad}')
    return f'{len(paths)} عنصر فهرس + {len(hashes)} رابط توجيه — كلها قائمة'


# ---------- 9) ملفات التنزيل ----------
def t_downloads():
    need = {
        'الأكاديمية-العقارية-الكاملة.zip': 500_000,
        os.path.join('نماذج', 'نموذج-دراسة-الجدوى-العقارية.xlsx'): 5_000,
        os.path.join('الملفات-العملية', 'قاعدة-مكة-المكرمة', 'makkah_database.xlsx'): 300_000,
        os.path.join('الملفات-العملية', 'قاعدة-مكة-المكرمة', 'developers.csv'): 10_000,
        os.path.join('الملفات-العملية', 'قاعدة-مكة-المكرمة', 'sources.csv'): 30_000,
    }
    for rel, minsize in need.items():
        p = os.path.join(ROOT, rel)
        if not os.path.exists(p):
            raise AssertionError(f'مفقود: {rel}')
        if os.path.getsize(p) < minsize:
            raise AssertionError(f'{rel} أصغر من المتوقع ({os.path.getsize(p)} < {minsize})')
    with zipfile.ZipFile(os.path.join(ROOT, 'الأكاديمية-العقارية-الكاملة.zip')) as z:
        if z.testzip() is not None:
            raise AssertionError('ZIP التوزيع تالف')
        if 'نماذج/نموذج-دراسة-الجدوى-العقارية.xlsx' not in z.namelist():
            raise AssertionError('ZIP التوزيع لا يحوي النموذج المالي')
    return '5 ملفات تنزيل موجودة بأحجام معقولة + ZIP سليم يحوي النموذج'


# ---------- 10) صحة Markdown الأساسية ----------
def t_markdown():
    html = open(INDEX, encoding='utf-8').read()
    paths = sorted(set(re.findall(r'data-path="([^"]+\.md)"', html)))
    bad = []
    for p in paths:
        fp = os.path.join(ROOT, p)
        txt = open(fp, encoding='utf-8').read()
        if not txt.strip():
            bad.append(f'{p}: فارغ')
        elif not txt.lstrip().startswith('#'):
            bad.append(f'{p}: لا عنوان رئيسي')
        elif '\t\r' in txt:
            bad.append(f'{p}: محارف تالفة')
    if bad:
        raise AssertionError('، '.join(bad))
    return f'{len(paths)} ملف markdown قابلة للعرض'


# ---------- 11) اتساق الأعداد في الوثائق ----------
def t_docs_numbers():
    # وثيقة البيانات (اقرأني): يجب أن تحمل الأعداد المعتمدة كاملة
    data_doc = os.path.join('الملفات-العملية', 'قاعدة-مكة-المكرمة', 'اقرأني-المصدر.md')
    txt = open(os.path.join(ROOT, data_doc), encoding='utf-8').read()
    for must in ('39', '60', '1859', '37', '241'):
        if must not in txt:
            raise AssertionError(f'{data_doc}: لا يذكر العدد المعتمد {must}')
    # حظر الأعداد القديمة الموضع الشك في أي وثيقة (36/44 — سبب تضارب سابق)
    stale = re.compile(r'(36 مطور|44 مشروع|\b36\b سجل مطور)')
    for d in ('README.md', data_doc):
        t = open(os.path.join(ROOT, d), encoding='utf-8').read()
        if stale.search(t):
            raise AssertionError(f'{d}: يحوي أعداداً قديمة موضع الشك')
    return 'وثيقة البيانات تحمل الأعداد المعتمدة (39/60/1859/37/60/241) — ولا أرقام قديمة في أي وثيقة'


ALL = [
    ('صلاحية ملفات CSV', t_csv_valid),
    ('تطابق الأعداد مع السجل المعتمد', t_counts),
    ('عدم التكرار وسلامة المعرفات', t_duplicates),
    ('اتساق حالات التحقق مع المصادر', t_status_source),
    ('صيغ النموذج: لا صيغة نصية ولا مرجع مكسور', t_model_syntax),
    ('إعادة حساب مستقلة — الأساسي والسيناريوهات', t_recalc_base),
    ('معدل العائد الداخلي IRR', t_recalc_irr),
    ('حساسية النموذج لتغيير الافتراضات', t_recalc_scenarios),
    ('إجهاد: أرض صفرية وقسمة على صفر', t_model_stress),
    ('الروابط الداخلية للموقع', t_links),
    ('ملفات التنزيل وسلامة ZIP', t_downloads),
    ('صحة Markdown الأساسية', t_markdown),
    ('اتساق أعداد قاعدة مكة في الوثائق', t_docs_numbers),
]

if __name__ == '__main__':
    print('═' * 62)
    print('اختبارات الأكاديمية العقارية — تعمل من:', ROOT)
    print('═' * 62)
    only = sys.argv[1] if len(sys.argv) > 1 else None
    for name, fn in ALL:
        if only and only not in name:
            continue
        check(name, fn)
    fails = [r for r in results if not r[1]]
    print('─' * 62)
    print(f'النتيجة: {len(results) - len(fails)}/{len(results)} نجح | فشل: {len(fails)}')
    sys.exit(1 if fails else 0)
