#!/usr/bin/env python3
"""Regression checks for teaching arithmetic, answer coverage, joins and shipped bundles."""
import csv,hashlib,json,re,unittest,zipfile
from pathlib import Path
from model_eval import Workbook
ROOT=Path(__file__).resolve().parents[1]

def rows(name):
 with (ROOT/'التطبيقات-المحلولة'/name).open(encoding='utf-8',newline='') as f:return list(csv.reader(f))

class LearningRegression(unittest.TestCase):
 def test_development_scenarios(self):
  for price,build,total,profit in [(520000,1600,4289200,910800),(540000,1600,4292200,1107800),(470000,1750,4508500,191500),(470000,1600,4281700,418300)]:
   revenue=10*price;construction=1440*build
   actual=1440000*1.05+construction+180000+100000+revenue*.015+construction*.05
   self.assertAlmostEqual(actual,total);self.assertAlmostEqual(revenue-actual,profit)
  # Independent shocks on the base example, including variable marketing and reserve.
  for total,profit,expected in [(4289200-260000*.015,654700,15.28),(4289200+230400*1.05,668880,14.76),(4289200+144000*1.05,759600,17.11)]:
   self.assertAlmostEqual(profit/total*100,expected,places=2)
 def test_land_price_shock_includes_acquisition_tax(self):
  total=4289200+(2700-2400)*600*1.05
  self.assertEqual(total,4478200);self.assertAlmostEqual((5200000-total)/total,0.16118083158411864)
 def test_return_timing(self):
  self.assertAlmostEqual(-4289000+5200000/1.12,353857.1428571427)
  monthly=1.12**(1/12)-1
  self.assertAlmostEqual((1+monthly)**12-1,.12)
  self.assertAlmostEqual(5200000/4289000-1,.21240382373513644)
 def test_all_answer_numbers_present(self):
  key=(ROOT/'التقييم/دليل-الإجابات-والتصحيح.md').read_text()
  for n in range(1,10):
   block=key.split(f'## المرحلة {n}\n',1)[1].split('\n## ',1)[0]
   self.assertEqual([int(x) for x in re.findall(r'^(\d+)\. ',block,re.M)],list(range(1,11)))
  blocks=key.split('## النهائي:')[1:]
  numbers=[int(x) for b in blocks for x in re.findall(r'^(\d+)\. ',b.split('\n## ',1)[0],re.M)]
  self.assertEqual(numbers,list(range(1,41)))
  self.assertIn('مفتاح التشخيصي (25 نقطة)',key)
 def test_crm_join_and_counts(self):
  properties=rows('عقارات-تدريبية.csv')[1:];clients=rows('عملاء-تدريبيون.csv')[1:];opportunities=rows('فرص-تدريبية.csv')[1:]
  prices={r[0]:int(r[3]) for r in properties};client_ids={r[0] for r in clients}
  self.assertEqual(len({r[0] for r in opportunities}),6)
  self.assertTrue(all(r[1] in client_ids and r[2] in prices and r[6]=='لا إرسال' for r in opportunities))
  self.assertEqual(sum(prices[r[2]] for r in opportunities),3215000)
  self.assertEqual({s:sum(r[3]==s for r in opportunities) for s in {r[3] for r in opportunities}}, {'جديد':2,'مؤهل':1,'معاينة':1,'عرض':1,'مغلق مرفوض':1})
 def test_cash_cumulative_is_undiscounted(self):
  wb=Workbook(str(ROOT/'نماذج/نموذج-دراسة-الجدوى-العقارية.xlsx'))
  running=0
  for r in range(2,10):
   running+=wb.eval('التدفقات',f'D{r}')
   self.assertAlmostEqual(wb.eval('التدفقات',f'G{r}'),running)
  self.assertEqual(wb.eval('التدفقات','G5'),-12437500)
  self.assertEqual(wb.eval('التدفقات','G9'),2112500)
 def test_bundles_match_manifest(self):
  for name in ['الأكاديمية-العقارية-الكاملة.zip','أكاديمية-العقار-كاملة.zip']:
   with zipfile.ZipFile(ROOT/name) as z:
    self.assertIsNone(z.testzip());manifest=json.loads(z.read('bundle-manifest.json'))
    for member,sha in manifest['files'].items():self.assertEqual(hashlib.sha256(z.read(member)).hexdigest(),sha,member)
    for required in ['التقييم/دليل-الإجابات-والتصحيح.md','التطبيقات-المحلولة/مختبر-التطبيق-العملي.md','assets/app.js','نماذج/نموذج-دراسة-الجدوى-العقارية.xlsx']:self.assertIn(required,z.namelist())
    catalog=json.loads(z.read('assets/catalog.json'))
    self.assertTrue(all(d['path'] in z.namelist() for d in catalog['documents']))
 def test_catalog_matches_current_text(self):
  catalog=json.loads((ROOT/'assets/catalog.json').read_text())
  for d in catalog['documents']:self.assertEqual((ROOT/d['path']).read_text(),d['text'],d['path'])

if __name__=='__main__':unittest.main(verbosity=2)
