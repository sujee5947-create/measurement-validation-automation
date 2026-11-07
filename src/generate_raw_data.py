"""Reproducible synthetic electrode measurement comparison dataset.
Never pass validation-only scenario notes to the assessment engine.
"""
from pathlib import Path
import csv
from datetime import datetime, timedelta
import numpy as np

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / 'data'
DATA.mkdir(parents=True,exist_ok=True)
SEED = 20261030

METRICS = [
    dict(characteristic='uncoated_width_mm', display='Uncoated width', unit='mm', target=10., part_sd=.60, a_sd=.055, b_sd=.055, b_bias=.035),
    dict(characteristic='coated_width_mm', display='Coated width', unit='mm', target=600., part_sd=1.80, a_sd=.15, b_sd=.16, b_bias=.14),
    dict(characteristic='uncoated_gap_mm', display='Uncoated gap', unit='mm', target=5., part_sd=.30, a_sd=.038, b_sd=.043, b_bias=.020),
    dict(characteristic='coating_thickness_um', display='Coating thickness', unit='um', target=100., part_sd=2.80, a_sd=.20, b_sd=.22, b_bias=.62),
    dict(characteristic='coating_weight_mg_cm2', display='Coating weight', unit='mg/cm2', target=20., part_sd=.70, a_sd=.045, b_sd=.047, b_bias=.025),
]


def generate():
    rng = np.random.default_rng(SEED)
    records=[]
    start = datetime(2026, 1, 5, 8, 0, 0)
    seconds=0
    truth=[]
    for sample in range(1, 121):
        sample_id=f'S{sample:03d}'
        lot=f'LOT-{(sample-1)//20+1:03d}'
        position=['LEFT','CENTER','RIGHT'][(sample-1)%3]
        for metric in METRICS:
            prop=metric['characteristic']
            latent = metric['target'] + rng.normal(0,metric['part_sd'])
            for rep in range(1,4):
                for device in ('DEVICE_A','DEVICE_B'):
                    bias=0.
                    if device == 'DEVICE_B':
                        bias=metric['b_bias']
                        if prop == 'coated_width_mm':
                            # Proportional bias: high Pearson r does not imply interchangeability.
                            bias += .065*(latent-metric['target'])
                        elif prop == 'coating_weight_mg_cm2' and sample >= 76:
                            # Device-B shift late in the sequence (test scenario, not root cause).
                            bias += .15
                    sigma=metric['a_sd'] if device=='DEVICE_A' else metric['b_sd']
                    if prop=='uncoated_gap_mm' and device=='DEVICE_B' and 41<=sample<=60:
                        sigma *= 3.0
                    observed=latent + bias + rng.normal(0,sigma)
                    seconds += 20
                    records.append({
                        'timestamp': (start+timedelta(seconds=seconds)).strftime('%Y-%m-%d %H:%M:%S'),
                        'sample_id': sample_id,
                        'lot_id': lot,
                        'sample_position': position,
                        'characteristic': prop,
                        'unit': metric['unit'],
                        'instrument_id': device,
                        'repeat_id': rep,
                        'measurement_value': round(float(observed), 5),
                        'operator_id': 'OP-01',
                        'measurement_session': 'VALIDATION-2026-01',
                    })
            truth.append({'sample_id':sample_id,'characteristic':prop,'synthetic_latent_value':round(float(latent),6)})
    cols=list(records[0])
    with (DATA/'Measurement_Raw_Data.csv').open('w',encoding='utf-8-sig',newline='') as f:
        writer=csv.DictWriter(f,fieldnames=cols);writer.writeheader();writer.writerows(records)
    # 32 independent samples randomly selected with a separate fixed seed.
    chosen=sorted(np.random.default_rng(32).choice(np.arange(1,121),size=32,replace=False).tolist())
    with (DATA/'Sample_Selection_32.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['sample_id','selection_method','selection_seed'])
        for i in chosen:w.writerow([f'S{i:03d}','random_without_replacement',32])
    # Intentionally not included among analytic input files; validation labels only.
    validation=ROOT/'tests'/'Scenario_Manifest.csv'
    with validation.open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['characteristic','sample_range','injected_scenario'])
        w.writerows([
            ['uncoated_width_mm','all','Stable agreement'],
            ['coated_width_mm','all','Offset plus value-dependent bias on device B'],
            ['uncoated_gap_mm','S041-S060','Higher within-device-B variation'],
            ['coating_thickness_um','all','Constant positive device-B bias'],
            ['coating_weight_mg_cm2','S076-S120','Late positive device-B shift'],
        ])
    return len(records),len(chosen)

if __name__=='__main__':
    n,k=generate()
    print(f'Created {n} readings, {n//2} matched repeat pairs, {k} comparison subset samples')
