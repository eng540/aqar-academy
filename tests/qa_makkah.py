#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
تمرير جودة قاعدة مكة — يعيد تصنيف حالات التحقق وفق المنطق السداسي المعتمد:
  confirmed            = موثق من مصدر رسمي (is_official=1)
  existence_verified   = تم التحقق من وجود المنشأة بمصدر ثانوي موثوق (غير رسمي)
  needs_verification   = يحتاج إلى تحقق (بلا مصدر مقفول)
  inactive / closed    = غير نشط / مغلق
  unknown              = غير معروف
يضيف أعمدة: region، verification_tier، status_ar — ويوحّد الإملاء المديني،
ويصلح الملاحظات المتعارضة (DEV-0035 وCON-0060)، ويولّد توزيع-المدن.csv.
قابل للتكرار بلا أثر مزدوج (العمليات شرطية على القيمة الحالية).
"""
import csv
import os
from collections import Counter

DB = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                  'الملفات-العملية', 'قاعدة-مكة-المكرمة')

CITY_FIX = {'القنفذه': 'القنفذة', 'الخرمه': 'الخرمة',
            'مدينة الملك عبدالله الإقتصادية': 'الاقتصادية (رابغ)'}
CITY_COLS = ('hq_city', 'city')
REGION = 'منطقة مكة المكرمة'
FILES = {
    'developers.csv': 'dev_id', 'contractors.csv': 'contractor_id',
    'engineering_offices.csv': 'eng_id', 'broker_offices.csv': 'office_id',
    'projects.csv': 'project_id',
}


def read(path):
    with open(path, encoding='utf-8-sig', newline='') as fh:
        rdr = csv.reader(fh)
        header = next(rdr)
        return header, [list(r) for r in rdr]


def write(path, header, rows):
    with open(path, 'w', encoding='utf-8', newline='') as fh:
        w = csv.writer(fh)
        w.writerow(header)
        w.writerows(rows)


def main():
    sh, sr = read(os.path.join(DB, 'sources.csv'))
    si = {h: i for i, h in enumerate(sh)}
    sources = {r[si['source_id']]: r for r in sr}

    fix_log, city_log = [], []
    tiers, statuses = Counter(), Counter()
    city_distribution = {}

    for name, idc in FILES.items():
        path = os.path.join(DB, name)
        header, rows = read(path)
        idx = {h: i for i, h in enumerate(header)}
        scol = 'source_id' if 'source_id' in idx else 'source'
        for col in ('region', 'verification_tier', 'status_ar'):
            if col not in idx:
                idx[col] = len(header)
                header.append(col)
                for r in rows:
                    r.append('')
        for r in rows:
            rid = r[idx[idc]]
            for cc in CITY_COLS:
                if cc in idx:
                    v = (r[idx[cc]] or '').strip()
                    if v in CITY_FIX:
                        city_log.append(f'{name}:{rid} «{v}» → «{CITY_FIX[v]}»')
                        v = CITY_FIX[v]
                    r[idx[cc]] = v or 'غير محدد'
                    city_distribution.setdefault(name, Counter())[r[idx[cc]]] += 1
            st = (r[idx['status']] or '').strip() if 'status' in idx else ''
            sid = (r[idx[scol]] or '').strip()
            tier = ('رسمي' if (sid and sid in sources and sources[sid][si['is_official']] == '1')
                    else 'ثانوي' if (sid and sid in sources) else 'غير مصنف')
            note = (r[idx['notes']] or '') if 'notes' in idx else ''
            if st == 'confirmed' and tier == 'رسمي':
                new_st, ar = 'confirmed', 'موثق من مصدر رسمي'
                if rid == 'CON-0060' and 'تصحيح' not in note:
                    note = ('المصدر SRC-0249 هو الموقع الرسمي للشركة — يؤكد وجودها فقط؛ '
                            'الترخيص والنشاط الحالي غير موثقين [تصحيح ملاحظة قديمة متعارضة 2026-11-05]')
                    fix_log.append(f'{name}:{rid} صحّحت الملاحظة المتعارضة (المصدر موجود: SRC-0249)')
            elif st == 'confirmed' and tier == 'ثانوي':
                new_st, ar = 'existence_verified', 'تم التحقق من وجود المنشأة (مصدر ثانوي)'
                if rid == 'DEV-0035':
                    note = ('المصدر SRC-0242 (سعوديبيديا) يثبت وجود الهيئة كمنشأة حكومية فقط — '
                            'لا يوثق التعاونات أو النشاط الحالي [تصحيح ملاحظة قديمة متعارضة 2026-11-05]')
                elif '[تصحيح' not in note:
                    note = (note + ' | ' if note else '') + \
                        '[تصحيح 2026-11-05: أُعيد التصنيف من confirmed لأن المصدر ثانوي يثبت الوجود فقط]'
                fix_log.append(f'{name}:{rid} confirmed → existence_verified (مصدر ثانوي {sid})')
            elif st == 'confirmed':
                new_st, ar = 'needs_verification', 'يحتاج إلى تحقق'
                if '[تصحيح' not in note:
                    note = (note + ' | ' if note else '') + \
                        '[تصحيح 2026-11-05: أُسقط confirmed لغياب مصدر مقفول في السجل]'
                fix_log.append(f'{name}:{rid} confirmed → needs_verification (بلا مصدر مقفول)')
            elif st in ('needs_verification', 'insufficient'):
                new_st, ar = 'needs_verification', 'يحتاج إلى تحقق'
            elif st == '':
                new_st, ar = 'unknown', 'غير معروف'
            else:
                new_st, ar = st, {'inactive': 'غير نشط', 'closed': 'مغلق'}.get(st, st)
            if 'status' in idx:
                r[idx['status']] = new_st
            r[idx['region']] = REGION
            r[idx['verification_tier']] = tier
            r[idx['status_ar']] = ar
            if 'notes' in idx:
                r[idx['notes']] = note
            tiers[tier] += 1
            statuses[new_st] += 1
        write(path, header, rows)

    # توزيع المدن
    dist_path = os.path.join(DB, 'توزيع-المدن.csv')
    with open(dist_path, 'w', encoding='utf-8', newline='') as fh:
        w = csv.writer(fh)
        w.writerow(['الملف', 'المدينة', 'عدد السجلات'])
        for fname, counter in sorted(city_distribution.items()):
            for city, n in counter.most_common():
                w.writerow([fname, city, n])

    print('=== سجل التصحيحات (%d) ===' % len(fix_log))
    for line in fix_log[:12]:
        print(' ', line)
    if len(fix_log) > 12:
        print(f'  … و{len(fix_log) - 12} أخرى')
    print('=== توحيد المدن (%d) ===' % len(city_log))
    for line in city_log[:8]:
        print(' ', line)
    print('=== توزيع الطبقات ===', dict(tiers))
    print('=== توزيع الحالات ===', dict(statuses))


if __name__ == '__main__':
    main()
