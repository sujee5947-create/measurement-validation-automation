"""Build the editable, synthetic instrument agreement criteria workbook."""
from pathlib import Path
from artifact_tool import Workbook, SpreadsheetFile

ROOT=Path(__file__).resolve().parents[1]
DEST=ROOT/'data'/'Measurement_Agreement_SPEC.xlsx'
wb=Workbook.create()
spec=wb.worksheets.add('AGREEMENT_SPEC')
settings=wb.worksheets.add('SETTINGS')
guide=wb.worksheets.add('GUIDE')

headers=['characteristic','display_name','unit','enabled','max_abs_mean_bias','max_abs_loa_bound','max_abs_pair_diff','min_pair_pass_rate','min_samples','min_repeats_per_device','notes']
data=[
['uncoated_width_mm','Uncoated Width','mm','Y',0.10,0.25,0.20,0.95,30,3,'Synthetic training threshold; not an approved customer or instrument specification'],
['coated_width_mm','Coated Width','mm','Y',0.18,0.45,0.40,0.95,30,3,'Synthetic training threshold; not an approved customer or instrument specification'],
['uncoated_gap_mm','Uncoated Gap','mm','Y',0.08,0.28,0.22,0.95,30,3,'Synthetic training threshold; not an approved customer or instrument specification'],
['coating_thickness_um','Coating Thickness','um','Y',0.35,1.00,0.80,0.95,30,3,'Synthetic training threshold; not an approved customer or instrument specification'],
['coating_weight_mg_cm2','Coating Weight','mg/cm2','Y',0.09,0.18,0.13,0.95,30,3,'Synthetic training threshold; not an approved customer or instrument specification']]
spec.get_range('A1:K1').values=[headers]
spec.get_range('A2:K6').values=data
spec.get_range('A1:K1').format={'fill':'#17365D','font':{'bold':True,'color':'#FFFFFF'},'horizontal_alignment':'center','vertical_alignment':'center','row_height':34}
spec.get_range('A2:K6').format.row_height=36
spec.get_range('A2:K6').format.wrap_text=True
spec.get_range('H2:H6').set_number_format('0.0%')
spec.get_range('E2:G6').set_number_format('0.000')
for col,width in {'A':27,'B':22,'C':13,'D':11,'E':23,'F':22,'G':20,'H':20,'I':15,'J':24,'K':58}.items():
    spec.get_range(f'{col}:{col}').format.column_width=width
spec.freeze_panes.freeze_rows(1)
spec.get_range('D2:D6').data_validation={'rule':{'type':'list','values':['Y','N']}}
spec.get_range('A8:B9').values=[
['DISCLAIMER','All limits in this workbook are illustrative; they are not validated production, OEM or customer acceptance limits.'],
['DECISION','All acceptance checks must pass. Pearson correlation is descriptive and NOT an agreement criterion.']]
spec.get_range('A8:A9').format={'fill':'#FEE2E2','font':{'bold':True,'color':'#991B1B'},'row_height':45}
spec.get_range('B8:K9').format.row_height=45
spec.get_range('B8:B9').format.wrap_text=True

config=[
['parameter','value','description'],
['signed_difference','B_MINUS_A','Bias is calculated as Device B minus Device A.'],
['pairing_keys','sample_id + characteristic','Compare averages of 3 repeated readings within each device and sample.'],
['repeat_pairing','sample_id + characteristic + repeat_id','Check each measurement has exactly one A and one B for the same repetition.'],
['agreement_unit','per_sample_mean','Bland–Altman sample unit = sample-level device mean, NOT all repeated measurements as independent pairs.'],
['loa_multiplier',1.96,'LoA = mean device difference +/- 1.96 * sample SD of device differences.'],
['confidence_level',0.95,'Bias confidence interval to be computed with t distribution; LoA CI is optional and must be labelled approximate if implemented.'],
['minimum_samples',30,'Minimum independent samples for initial results; note CI precision depends on n.'],
['comparison_subset_n',32,'A fixed independently sampled subset of 32 is supplied separately.'],
['reference_status','COMPARATOR_ONLY','Device A is not assumed to be a traceable true-value standard.'],
['trend_window_samples',20,'Rolling window of sample-wise B-A means for trend visualization.'],
['time_bin_minutes',60,'Hourly bins for descriptive time trend (may not contain much data).'],
['decision_logic','ALL_THRESHOLDS','PASS only if abs(mean bias), both LoA bounds, and paired-sample tolerance pass rate all meet sheet limits.'],
['data_policy','SYNTHETIC_ONLY','Scenario manifest is validation-only; do not use it to calculate results.'],
]
settings.get_range('A1:C14').values=config
settings.get_range('A1:C1').format={'fill':'#17365D','font':{'bold':True,'color':'#FFFFFF'},'row_height':34}
settings.get_range('A2:C14').format.row_height=36
settings.get_range('A2:C14').format.wrap_text=True
for col,width in {'A':27,'B':43,'C':100}.items():settings.get_range(f'{col}:{col}').format.column_width=width
settings.freeze_panes.freeze_rows(1)

guide_rows=[
['SECTION','GUIDANCE'],
['Project','Measurement Correlation & Validation — electrode manufacturing device A vs device B'],
['Inputs','Measurement_Raw_Data.csv (3,600 long-format readings) and Measurement_Agreement_SPEC.xlsx'],
['Sample count','120 independent physical sample IDs; 5 characteristics; 2 devices; 3 repeats per device and sample'],
['32-sample subset','Sample_Selection_32.csv is reproducible random selection without replacement; compare uncertainty with full 120.'],
['Baseline protocol','Assess differences using 120 per-sample means, not 360 repeated values as independent observations.'],
['Repeatability','Report within-device repeated-reading SD separately, plus intra-device range/variance.'],
['Correlation','Pearson r describes association. A strong correlation is NOT sufficient proof of agreement.'],
['Bias','Mean(Device B - Device A) across independent samples.'],
['Bland-Altman','LoA = Bias +/- 1.96 x sample SD of differences; validity depends on difference distribution and proportional bias checks.'],
['Acceptance','Limits are editable. PASS requires all bias, LoA, and sample-level difference rate criteria.'],
['No gold standard','Device A is only comparator, not ground truth. This is an agreement study, not calibration validation.'],
['Separation','Measurement acceptance limits are NOT product dimensional specification limits.'],
['Event log','Trend periods can be summarized, but cannot imply actual equipment failures or calibration events without evidence.'],
['Confidentiality','100% synthetic generated data and criteria; no company data.'],
]
guide.get_range('A1:B15').values=guide_rows
guide.get_range('A1:B1').format={'fill':'#17365D','font':{'bold':True,'color':'#FFFFFF'},'row_height':34}
guide.get_range('A2:B15').format.row_height=34
guide.get_range('A2:B15').format.wrap_text=True
guide.get_range('A:A').format.column_width=27
guide.get_range('B:B').format.column_width=105
guide.freeze_panes.freeze_rows(1)

preview=wb.inspect({'kind':'table','range':'AGREEMENT_SPEC!A1:J4','include':'values','table_max_rows':4,'table_max_cols':10})
print(preview.ndjson[:2700])
SpreadsheetFile.export_xlsx(wb).save(str(DEST))
print('Created',DEST)
