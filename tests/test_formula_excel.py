"""Tests generated Excel formula traceability without modifying workbooks."""
import csv
from pathlib import Path
import unittest
import xml.etree.ElementTree as ET
import zipfile

ROOT=Path(__file__).resolve().parents[1]
XLSX=ROOT/'outputs'/'Agreement_Analysis.xlsx'
RAW=ROOT/'data'/'Measurement_Raw_Data.csv'
NAMESPACE={'x':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


class ExcelFormulaTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with zipfile.ZipFile(XLSX) as archive:
            cls.raw_xml=ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
            cls.sample_xml=ET.fromstring(archive.read('xl/worksheets/sheet5.xml'))
            cls.result_xml=ET.fromstring(archive.read('xl/worksheets/sheet6.xml'))

    @staticmethod
    def formula(sheet,cell):
        found=sheet.find(f'.//x:c[@r="{cell}"]/x:f',NAMESPACE)
        return found.text if found is not None else None

    def test_raw_csv_still_present_and_stored(self):
        with RAW.open(encoding='utf-8-sig',newline='') as f:
            records=list(csv.DictReader(f))
        self.assertEqual(len(records),3600)
        self.assertIsNotNone(self.raw_xml.find('.//x:c[@r="I2"]',NAMESPACE))

    def test_replicate_traces_to_original_raw(self):
        self.assertEqual(self.formula(self.sample_xml,'E2'),'RAW_DATA!$I$2')
        self.assertEqual(self.formula(self.sample_xml,'K2'),'AVERAGE(E2:G2)')
        self.assertEqual(self.formula(self.sample_xml,'M2'),'L2-K2')

    def test_result_contains_auditable_formulas(self):
        self.assertIn('AVERAGE',self.formula(self.result_xml,'D2'))
        self.assertIn('CORREL',self.formula(self.result_xml,'H2'))
        self.assertIn('SPEC_USED',self.formula(self.result_xml,'L2'))
        self.assertIn('AND',self.formula(self.result_xml,'Q2'))


if __name__=='__main__':
    unittest.main()
