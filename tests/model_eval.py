#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
مقيّم معادلات مستقل لملفات xlsx المولدة — يعيد حساب الصيغ من XML الملف الفعلي
(لا يقرأ نية المولد). يدعم المجموعة المستخدمة في النموذج:
+ - * / ^ ، أقواس، مراجع خلايا/نطاقات مع أسماء أوراق عربية مقتبسة،
SUM، POWER، ROUNDDOWN، IFERROR، IRR (بحث ثنائي).
يبين الأخطاء بأسلوب Excel: #DIV/0! #VALUE! #REF! #NAME? #NUM!
"""
import re
import zipfile
from xml.etree import ElementTree as ET

M = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


class EvalError(Exception):
    def __init__(self, code):
        self.code = code
        super().__init__(code)


def col_to_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def num_to_col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_ref(ref):
    m = re.fullmatch(r'([A-Z]+)([0-9]+)', ref)
    if not m:
        raise EvalError('#REF!')
    return m.group(1), int(m.group(2))


class Workbook:
    """يقرأ xlsx مبسط: أوراق مرتبة sheet1..n بحسب workbook.xml."""

    def __init__(self, path):
        self.cells = {}          # (sheet, REF) -> {'kind': 'f|s|n|blank', ...}
        self.order = []
        with zipfile.ZipFile(path) as z:
            wb = ET.fromstring(z.read('xl/workbook.xml'))
            names = [s.get('name') for s in wb.iter(M + 'sheet')]
            for i, name in enumerate(names, 1):
                ws = ET.fromstring(z.read(f'xl/worksheets/sheet{i}.xml'))
                for c in ws.iter(M + 'c'):
                    ref = c.get('r')
                    f = c.find(M + 'f')
                    v = c.find(M + 'v')
                    ist = c.find(M + 'is/' + M + 't')
                    if f is not None:
                        self.cells[(name, ref)] = {'kind': 'f', 'f': f.text or ''}
                    elif ist is not None:
                        self.cells[(name, ref)] = {'kind': 's', 'v': ist.text or ''}
                    elif v is not None and (v.text or '').strip():
                        self.cells[(name, ref)] = {'kind': 'n', 'v': float(v.text)}
                    else:
                        self.cells[(name, ref)] = {'kind': 'blank', 'v': None}
                self.order.append(name)
        self._memo = {}
        self._busy = set()
        self._t = None
        self._i = 0

    # ---------- واجهة عامة ----------
    def raw_formula(self, sheet, ref):
        c = self.cells.get((sheet, ref))
        return c['f'] if c and c['kind'] == 'f' else None

    def raw_kind(self, sheet, ref):
        c = self.cells.get((sheet, ref))
        return c['kind'] if c else 'missing'

    def raw_text(self, sheet, ref):
        c = self.cells.get((sheet, ref))
        return c.get('v') if c else None

    def eval(self, sheet, ref):
        key = (sheet, ref)
        if key in self._memo:
            return self._memo[key]
        c = self.cells.get(key)
        if c is None:
            raise EvalError('#REF!')
        if c['kind'] == 'f':
            if key in self._busy:
                raise EvalError('#REF!')
            self._busy.add(key)
            try:
                val = self._formula(c['f'], sheet)
            finally:
                self._busy.discard(key)
            self._memo[key] = val
            return val
        if c['kind'] in ('n', 's'):
            return c['v']
        return 0  # خلية فارغة = صفر (سلوك Excel في العمليات الحسابية)

    def evaluate_all(self):
        out = {}
        for (sh, ref), c in self.cells.items():
            if c['kind'] == 'f':
                try:
                    out[(sh, ref)] = self.eval(sh, ref)
                except EvalError as e:
                    out[(sh, ref)] = e.code
        return out

    # ---------- محلل الصيغ ----------
    def _tokens(self, s):
        toks = []
        i, n = 0, len(s)
        while i < n:
            ch = s[i]
            if ch.isspace():
                i += 1
                continue
            if ch == '"':
                j = s.index('"', i + 1)
                toks.append(('STR', s[i + 1:j]))
                i = j + 1
                continue
            if ch == "'":                      # اسم ورقة مقتبس
                j = s.index("'", i + 1)
                j2 = s.index('!', j + 1)
                toks.append(('SHEET', s[i + 1:j]))
                i = j2 + 1
                continue
            if ch.isdigit() or (ch == '.' and i + 1 < n and s[i + 1].isdigit()):
                m = re.match(r'[0-9]*\.?[0-9]+', s[i:])
                toks.append(('NUM', float(m.group(0))))
                i += len(m.group(0))
                continue
            if ch.isalpha():
                m = re.match(r'[A-Za-z]+[0-9]+', s[i:])
                if m:
                    toks.append(('CELL', m.group(0).upper()))
                    i += len(m.group(0))
                    continue
                m = re.match(r'[A-Za-z]+', s[i:])
                toks.append(('NAME', m.group(0).upper()))
                i += len(m.group(0))
                continue
            if ch == ':':
                toks.append(('COLON', ch))
                i += 1
                continue
            if ch in '+-*/^(),':
                toks.append(('OP', ch))
                i += 1
                continue
            raise EvalError('#NAME?')
        toks.append(('END', None))
        return toks

    def _formula(self, f, sheet):
        # عزل حالة المحلل: تقييم خلية داخلية لا يدمر محلل الصيغة الأب
        saved_t, saved_i = self._t, self._i
        self._t = self._tokens(f)
        self._i = 0
        try:
            val = self._expr(sheet)
            if self._t[self._i][0] != 'END':
                raise EvalError('#NAME?')
            return val
        finally:
            self._t, self._i = saved_t, saved_i

    def _peek(self):
        return self._t[self._i]

    def _next(self):
        t = self._t[self._i]
        self._i += 1
        return t

    def _expr(self, sheet):
        v = self._term(sheet)
        while self._peek() == ('OP', '+') or self._peek() == ('OP', '-'):
            op = self._next()[1]
            r = self._term(sheet)
            v = self._num(v, '#VALUE!') + r if op == '+' else self._num(v, '#VALUE!') - self._num(r, '#VALUE!')
        return v

    def _term(self, sheet):
        v = self._pow(sheet)
        while self._peek() == ('OP', '*') or self._peek() == ('OP', '/'):
            op = self._next()[1]
            r = self._pow(sheet)
            if op == '*':
                v = self._num(v, '#VALUE!') * self._num(r, '#VALUE!')
            else:
                d = self._num(r, '#VALUE!')
                if d == 0:
                    raise EvalError('#DIV/0!')
                v = self._num(v, '#VALUE!') / d
        return v

    def _pow(self, sheet):
        v = self._unary(sheet)
        if self._peek() == ('OP', '^'):
            self._next()
            r = self._pow(sheet)
            v = self._num(v, '#VALUE!') ** self._num(r, '#VALUE!')
        return v

    def _unary(self, sheet):
        if self._peek() == ('OP', '-'):
            self._next()
            return -self._num(self._unary(sheet), '#VALUE!')
        if self._peek() == ('OP', '+'):
            self._next()
            return self._unary(sheet)
        return self._atom(sheet)

    def _atom(self, sheet):
        kind, val = self._next()
        if kind == 'NUM':
            return val
        if kind == 'STR':
            return val
        if kind == 'SHEET':
            k2, v2 = self._next()
            if k2 == 'CELL':
                if self._peek() == ('COLON', ':'):
                    self._next()
                    k3, v3 = self._next()
                    if k3 != 'CELL':
                        raise EvalError('#REF!')
                    return self._range(val, v2, v3)
                return self.eval(val, v2)
            raise EvalError('#REF!')
        if kind == 'CELL':
            if self._peek() == ('COLON', ':'):
                self._next()
                k3, v3 = self._next()
                if k3 != 'CELL':
                    raise EvalError('#REF!')
                return self._range(sheet, val, v3)
            return self.eval(sheet, val)
        if kind == 'NAME':
            if self._peek() != ('OP', '('):
                raise EvalError('#NAME?')
            self._next()
            arg_slices = []
            if self._peek() != ('OP', ')'):
                start = self._i
                depth = 0
                while True:
                    k, v = self._t[self._i]
                    if k == 'OP' and v == '(':
                        depth += 1
                    elif k == 'OP' and v == ')':
                        if depth == 0:
                            break
                        depth -= 1
                    elif k == 'OP' and v == ',' and depth == 0:
                        arg_slices.append((start, self._i))
                        start = self._i + 1
                    self._i += 1
                arg_slices.append((start, self._i))
            if self._next() != ('OP', ')'):
                raise EvalError('#NAME?')
            if val == 'IFERROR':          # تقييم كسول: الخطأ في الوسيط الأول يُلتقط هنا
                return self._iferror(sheet, arg_slices)
            args = [self._eval_slice(sheet, a, b) for a, b in arg_slices]
            return self._call(val, args)
        if kind == 'OP' and val == '(':
            v = self._expr(sheet)
            if self._next() != ('OP', ')'):
                raise EvalError('#NAME?')
            return v
        raise EvalError('#NAME?')

    def _range(self, sheet, a, b):
        c1, r1 = split_ref(a)
        c2, r2 = split_ref(b)
        vals = []
        for row in range(min(r1, r2), max(r1, r2) + 1):
            for col in range(min(col_to_num(c1), col_to_num(c2)), max(col_to_num(c1), col_to_num(c2)) + 1):
                vals.append((sheet, f'{num_to_col(col)}{row}'))
        return ('RANGE', vals)

    def _num(self, v, code='#VALUE!'):
        if isinstance(v, tuple) and v[0] == 'RANGE':
            raise EvalError(code)
        if isinstance(v, str):
            raise EvalError(code)
        return v

    def _eval_slice(self, sheet, a, b):
        saved_t, saved_i = self._t, self._i
        self._t = self._t[a:b] + [('END', None)]
        self._i = 0
        try:
            return self._expr(sheet)
        finally:
            self._t, self._i = saved_t, saved_i

    def _iferror(self, sheet, slices):
        if not slices:
            raise EvalError('#VALUE!')
        try:
            return self._eval_slice(sheet, *slices[0])
        except EvalError:
            if len(slices) < 2:
                return ''
            return self._eval_slice(sheet, *slices[1])

    def _call(self, name, args):
        if name == 'SUM':
            total = 0.0
            for a in args:
                if isinstance(a, tuple) and a[0] == 'RANGE':
                    for key in a[1]:
                        try:
                            total += self._num(self.eval(*key), '#VALUE!')
                        except EvalError as e:
                            if e.code == '#REF!':
                                raise
                else:
                    total += self._num(a, '#VALUE!')
            return total
        if name == 'POWER':
            return self._num(args[0]) ** self._num(args[1])
        if name == 'ROUNDDOWN':
            x, d = self._num(args[0]), int(self._num(args[1]))
            p = 10 ** d
            return float(int(x * p) / p) if x >= 0 else -float(int(-x * p) / p)
        if name == 'IFERROR':
            try:
                v = self._num(args[0]) if not isinstance(args[0], str) else args[0]
                return v
            except EvalError:
                return args[1]
        if name == 'IRR':
            vals = []
            for a in args:
                if isinstance(a, tuple) and a[0] == 'RANGE':
                    for key in a[1]:
                        c = self.cells.get(key)
                        vals.append(0.0 if c is None or c['kind'] == 'blank' else self._num(self.eval(*key), '#VALUE!'))
                else:
                    vals.append(self._num(a, '#VALUE!'))
            return self._irr(vals)
        raise EvalError('#NAME?')

    @staticmethod
    def _npv(vals, r):
        return sum(v / (1 + r) ** i for i, v in enumerate(vals))

    def _irr(self, vals):
        lo, hi = -0.9999, 10.0
        flo, fhi = self._npv(vals, lo), self._npv(vals, hi)
        if flo == 0:
            return lo
        if fhi == 0:
            return hi
        if flo * fhi > 0:
            raise EvalError('#NUM!')
        for _ in range(200):
            mid = (lo + hi) / 2
            fm = self._npv(vals, mid)
            if flo * fm <= 0:
                hi, fhi = mid, fm
            else:
                lo, flo = mid, fm
        return (lo + hi) / 2
