import unittest
from pathlib import Path
import sys

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from main import load_spec,load_raw,build_records,calculate

class AgreementTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.specs=load_spec()
        cls.groups,cls.info,cls.raw_n=load_raw()
        cls.details=build_records(cls.groups,cls.info,cls.specs)

    def test_expected_measurement_design(self):
        self.assertEqual(self.raw_n,3600)
        self.assertEqual(len(self.specs),5)
        for value in self.details.values():
            self.assertEqual(len(value),120)
            self.assertEqual(len(value[0]['device_a_repeats']),3)
            self.assertEqual(len(value[0]['device_b_repeats']),3)

    def test_correlation_not_agreement(self):
        spec=next(x for x in self.specs if x['characteristic']=='coating_thickness_um')
        result=calculate(self.details[spec['characteristic']],spec)
        self.assertGreater(result['pearson_r'],.95)
        self.assertEqual(result['status'],'REVIEW')
        self.assertFalse(result['checks']['mean_bias'])

    def test_spec_changes_decision(self):
        spec=next(x for x in self.specs if x['characteristic']=='coating_thickness_um').copy()
        baseline=calculate(self.details[spec['characteristic']],spec)
        spec.update(max_abs_mean_bias=2,max_abs_loa_bound=2,max_abs_pair_diff=2,min_pair_pass_rate=.95)
        relaxed=calculate(self.details[spec['characteristic']],spec)
        self.assertEqual(baseline['status'],'REVIEW')
        self.assertEqual(relaxed['status'],'PASS')

    def test_repeats_not_mistaken_for_independent_samples(self):
        spec=self.specs[0]
        result=calculate(self.details[spec['characteristic']],spec)
        self.assertEqual(result['n'],120)

if __name__=='__main__':unittest.main()
