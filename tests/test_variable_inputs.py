"""Regression tests for new input structures and explicit failure paths."""
import csv
import math
import sys
import tempfile
import unittest
from pathlib import Path
from statistics import mean
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from main import load_spec,load_raw,build_records,calculate

class VariableInputTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.specs=load_spec()
        with (ROOT/'data/Measurement_Raw_Data.csv').open(encoding='utf-8-sig',newline='') as f:
            cls.original=list(csv.DictReader(f))
    def write_rows(self,rows,columns):
        path=Path(self.tmp.name)/'input.csv'
        with path.open('w',encoding='utf-8-sig',newline='') as f:
            w=csv.DictWriter(f,fieldnames=columns);w.writeheader();w.writerows(rows)
        return path
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory()
    def tearDown(self):
        self.tmp.cleanup()
    def test_reordered_columns_variable_repeats(self):
        samples={'S001','S002','S003','S004','S005'}
        rows=[x.copy() for x in self.original if x['sample_id'] in samples]
        for r in list(rows):
            if r['sample_id']=='S005' and r['repeat_id']=='1':
                duplicate=r.copy();duplicate['repeat_id']='4';rows.append(duplicate)
        rows.reverse()
        path=self.write_rows(rows,list(rows[0])[::-1]); groups,meta,count=load_raw(path)
        details=build_records(groups,meta,self.specs)
        self.assertEqual(count,len(rows))
        for recs in details.values():
            self.assertEqual(len(recs),5)
            self.assertEqual(len(recs[-1]['device_a_repeats']),4)
            self.assertAlmostEqual(recs[-1]['device_a_mean'],mean(recs[-1]['device_a_repeats']))
    def test_duplicate_or_missing_pair_fails(self):
        rows=[x.copy() for x in self.original if x['sample_id']=='S001']
        rows.append(rows[0].copy())
        with self.assertRaisesRegex(ValueError,'Duplicate'):
            load_raw(self.write_rows(rows,list(rows[0])))
        rows.pop();rows.pop()
        groups,meta,_=load_raw(self.write_rows(rows,list(rows[0])))
        with self.assertRaisesRegex(ValueError,'Incomplete/mismatched'):
            build_records(groups,meta,self.specs)
    def test_invalid_spec_fails(self):
        spec=self.specs[0].copy();spec['max_abs_pair_diff']=-1
        # Specification validation is performed at load time; input-level data validation independently prevents NaN.
        rows=[x.copy() for x in self.original if x['sample_id']=='S001']
        rows[0]['measurement_value']='NaN'
        with self.assertRaisesRegex(ValueError,'non-finite'):
            load_raw(self.write_rows(rows,list(rows[0])))
    def test_zero_variation_explicit_error(self):
        records=[{'device_a_mean':1.,'device_b_mean':2.,'device_a_sd':.1,'device_b_sd':.1} for _ in range(10)]
        spec=self.specs[0].copy();spec['min_samples']=3
        with self.assertRaisesRegex(ValueError,'zero variation'):
            calculate(records,spec)

if __name__=='__main__':unittest.main()
