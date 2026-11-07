"""Regression tests for real-world RAW quirks (v3 fixes)."""
import csv, math, sys, tempfile, unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from main import load_spec,load_raw,build_records,calculate

class RealWorldInputTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.specs=load_spec()
        with (ROOT/'data/Measurement_Raw_Data.csv').open(encoding='utf-8-sig',newline='') as f:
            cls.original=list(csv.DictReader(f))
        cls.cols=list(cls.original[0])
    def setUp(self): self.tmp=tempfile.TemporaryDirectory()
    def tearDown(self): self.tmp.cleanup()
    def write(self,rows,enc='utf-8-sig',tail=''):
        p=Path(self.tmp.name)/'in.csv'
        with p.open('w',encoding=enc,newline='') as f:
            w=csv.DictWriter(f,fieldnames=self.cols);w.writeheader();w.writerows(rows);f.write(tail)
        return p
    def test_cp949_blank_rows_spaces_lowercase(self):
        rows=[r.copy() for r in self.original]
        rows[0]['operator_id']='홍길동'; rows[1]['sample_id']=' S001 '; rows[2]['instrument_id']='device_a'
        g,i,n=load_raw(self.write(rows,'cp949',','*(len(self.cols)-1)+'\r\n'))
        self.assertEqual(n,3600)
        d=build_records(g,i,self.specs)
        self.assertEqual(len(d['uncoated_width_mm']),120)
    def test_natural_sample_order(self):
        rows=[r.copy() for r in self.original]
        for r in rows: r['sample_id']='S'+str(int(r['sample_id'][1:]))
        g,i,_=load_raw(self.write(rows)); d=build_records(g,i,self.specs)
        self.assertEqual([x['sample_id'] for x in d['uncoated_width_mm'][:3]],['S1','S2','S3'])
    def test_pooled_sigma_df_weighted(self):
        recs=[{'device_a_mean':float(k),'device_b_mean':k+.1+.01*(k%3),'device_a_sd':.1 if k else .4,'device_b_sd':.1,
               'device_a_n':6 if k==0 else 3,'device_b_n':3} for k in range(10)]
        spec=dict(self.specs[0]);spec['min_samples']=3
        r=calculate(recs,spec)
        self.assertAlmostEqual(r['a_repeatability_sigma'],math.sqrt((5*.16+9*2*.01)/(5+18)))
if __name__=='__main__':unittest.main()
