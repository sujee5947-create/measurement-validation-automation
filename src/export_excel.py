"""Excel traceability report, built with openpyxl (public PyPI package).

RAW_DATA is a snapshot of the input CSV (blank rows skipped, whitespace trimmed);
all derived RESULT, SUBSET_RESULT and SAMPLE_CALC cells contain native Excel
formulas. Excel recalculates every formula when the workbook is opened.
The source CSV and SPEC XLSX are never modified.
"""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as letters
from openpyxl.workbook.properties import CalcProperties
from src.raw_io import read_raw_rows

NAVY = '183557'
WHITE = 'FFFFFF'
GREEN = 'DDF4E8'
RED = 'FCE5E5'
STDEV = '_xlfn.STDEV.S'   # required prefix so Excel recognises STDEV.S written by openpyxl


def header(sheet, headers, widths=None):
    sheet.append(list(headers))
    for cell in sheet[1]:
        cell.fill = PatternFill('solid', fgColor=NAVY)
        cell.font = Font(bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    sheet.row_dimensions[1].height = 31
    sheet.freeze_panes = 'A2'
    for col, width in enumerate(widths or [], 1):
        sheet.column_dimensions[letters(col)].width = width


def fmt(sheet, ref, number_format):
    for row in sheet[ref]:
        for cell in (row if isinstance(row, tuple) else (row,)):
            cell.number_format = number_format


def export_excel(root, specs, results, details, output_dir=None, subset_size=32,
                 raw_path=None, selected_ids=None):
    root = Path(root)
    out = Path(output_dir) if output_dir else root/'outputs'
    out.mkdir(parents=True, exist_ok=True)
    path = Path(raw_path) if raw_path else root/'data'/'Measurement_Raw_Data.csv'
    cols, raw_rows, _enc = read_raw_rows(path)

    source_lookup = {}
    for excel_row, rec in enumerate(raw_rows, 2):
        key = (rec['characteristic'], rec['sample_id'], rec['instrument_id'], int(rec['repeat_id']))
        if key in source_lookup:
            raise ValueError(f'Duplicate instrument reading key: {key}')
        source_lookup[key] = excel_row

    wb = Workbook()
    wb.remove(wb.active)
    raw = wb.create_sheet('RAW_DATA')
    spec_sheet = wb.create_sheet('SPEC_USED')
    settings = wb.create_sheet('SETTINGS')
    selected = wb.create_sheet('SELECTED_SAMPLES')
    pair = wb.create_sheet('SAMPLE_CALC')
    result = wb.create_sheet('RESULT')
    subset = wb.create_sheet('SUBSET_CALC')
    subresult = wb.create_sheet('SUBSET_RESULT')
    guide = wb.create_sheet('FORMULA_GUIDE')

    header(raw, cols, [max(14, min(28, len(c)+4)) for c in cols])
    for rec in raw_rows:
        raw.append([int(rec[c]) if c == 'repeat_id' else float(rec[c]) if c == 'measurement_value' else rec[c] for c in cols])
    value_col = letters(cols.index('measurement_value')+1)
    fmt(raw, f'{value_col}2:{value_col}{len(raw_rows)+1}', '0.00000')

    spec_heads = ['characteristic', 'display_name', 'unit', 'enabled', 'max_abs_mean_bias', 'max_abs_loa_bound',
                  'max_abs_pair_diff', 'min_pair_pass_rate', 'min_samples', 'min_repeats_per_device']
    header(spec_sheet, spec_heads, [28, 25, 13, 12, 21, 20, 21, 22, 15, 22])
    for s in specs:
        spec_sheet.append([s.get(k) for k in spec_heads])
    fmt(spec_sheet, f'H2:H{len(specs)+1}', '0.0%')

    header(settings, ['Setting', 'Value', 'Purpose'], [26, 18, 82])
    for row in [
        ['LoA multiplier', 1.96, 'Bland–Altman LoA = Bias ± multiplier × SD of independent sample differences'],
        ['Signed difference', 'B_MINUS_A', 'Device B sample mean minus Device A sample mean'],
        ['Agreement unit', 'sample_mean', 'Each row is a physical sample; per-device replicate means are compared'],
        ['Comparator', 'DEVICE_A', 'Comparator is not a certified reference or true value'],
        ['Origin', 'input-dependent', 'Bundled example is synthetic; verify provenance for any supplied external raw data'],
    ]:
        settings.append(row)

    chosen_set = set(selected_ids or [])
    if not chosen_set:
        chosen_set = set(sorted({r['sample_id'] for r in raw_rows})[:subset_size])
    selected_list = sorted(chosen_set)
    header(selected, ['sample_id', 'purpose'], [18, 66])
    for x in selected_list:
        selected.append([x, 'Exploratory subset (not independent of full cohort)'])
    select_last = len(selected_list)+1

    max_reps = max(max(len(r['device_a_repeats']), len(r['device_b_repeats'])) for records in details.values() for r in records)
    a_start = 5
    b_start = a_start+max_reps
    derived_start = b_start+max_reps
    a_first, a_last = letters(a_start), letters(b_start-1)
    b_first, b_last = letters(b_start), letters(derived_start-1)
    names = ('a_mean', 'b_mean', 'bias', 'abs_diff', 'a_sd', 'b_sd', 'tolerance', 'pair_status', 'selected', 'pair_mean', 'a_n', 'b_n')
    C = {name: letters(derived_start+i) for i, name in enumerate(names)}
    pair_heads = (['characteristic', 'sample_id', 'lot_id', 'timestamp']
                  + [f'A repeat {i}' for i in range(1, max_reps+1)] + [f'B repeat {i}' for i in range(1, max_reps+1)]
                  + ['A mean', 'B mean', 'Bias (B-A)', 'Abs difference', 'A repeat sigma', 'B repeat sigma',
                     'Pair tolerance', 'Pair status', 'Selected', 'Pair mean', 'A repeat count', 'B repeat count'])
    header(pair, pair_heads, [27, 13, 14, 23]+[16]*(2*max_reps)+[17]*12)
    group_ranges = {}; sample_rows = {}
    i = 2
    for spec_row, spec in enumerate(specs, 2):
        metric = spec['characteristic']; recs = details[metric]
        begin = i
        for rec in recs:
            sample = rec['sample_id']
            readings = []
            for dev, repeat_values in (('DEVICE_A', rec['device_a_repeats']), ('DEVICE_B', rec['device_b_repeats'])):
                for rep in range(1, max_reps+1):
                    readings.append(f'=RAW_DATA!${value_col}${source_lookup[(metric, sample, dev, rep)]}' if rep <= len(repeat_values) else None)
            pair.append([metric, sample, rec['lot_id'], rec['timestamp']] + readings + [
                f'=AVERAGE({a_first}{i}:{a_last}{i})', f'=AVERAGE({b_first}{i}:{b_last}{i})',
                f'={C["b_mean"]}{i}-{C["a_mean"]}{i}', f'=ABS({C["bias"]}{i})',
                f'={STDEV}({a_first}{i}:{a_last}{i})', f'={STDEV}({b_first}{i}:{b_last}{i})',
                f'=SPEC_USED!$G${spec_row}', f'=IF({C["abs_diff"]}{i}<={C["tolerance"]}{i},"PASS","NG")',
                f'=IF(COUNTIF(SELECTED_SAMPLES!$A$2:$A${select_last},B{i})>0,1,0)',
                f'=AVERAGE({C["a_mean"]}{i}:{C["b_mean"]}{i})',
                f'=COUNT({a_first}{i}:{a_last}{i})', f'=COUNT({b_first}{i}:{b_last}{i})',
            ])
            sample_rows[(metric, sample)] = i
            i += 1
        group_ranges[metric] = (begin, i-1, spec_row)
    last = i-1
    fmt(pair, f'{a_first}2:{C["tolerance"]}{last}', '0.00000')
    fmt(pair, f'{C["pair_mean"]}2:{C["pair_mean"]}{last}', '0.00000')
    pair.conditional_formatting.add(f'{C["pair_status"]}2:{C["pair_status"]}{last}',
                                    FormulaRule(formula=[f'{C["pair_status"]}2="NG"'], fill=PatternFill('solid', fgColor=RED)))

    def pooled(sd_rng, n_rng):
        # df-weighted pooled SD: sqrt( Σ(n_i-1)s_i² / Σ(n_i-1) )
        return f'=SQRT(SUMPRODUCT({n_rng}-1,{sd_rng}^2)/(SUM({n_rng})-COUNT({n_rng})))'

    result_heads = ['characteristic', 'display_name', 'n samples', 'mean bias B-A', 'SD difference',
                    'LoA lower', 'LoA upper', 'Pearson r', 'Pair pass rate', 'A repeat sigma', 'B repeat sigma',
                    'Max |Bias|', 'Max |LoA|', 'Max pair diff', 'Min pass rate', 'Minimum n', 'Decision']
    header(result, result_heads, [27, 23, 14, 20, 19, 19, 19, 17, 18, 20, 20, 17, 17, 17, 18, 16, 17])
    for i, s in enumerate(specs, 2):
        metric = s['characteristic']; st, en, spec_r = group_ranges[metric]
        R = lambda col: f'SAMPLE_CALC!${col}${st}:${col}${en}'
        result.append([metric, s['display_name'],
            f'=COUNT({R(C["bias"])})', f'=AVERAGE({R(C["bias"])})', f'={STDEV}({R(C["bias"])})',
            f'=D{i}-SETTINGS!$B$2*E{i}', f'=D{i}+SETTINGS!$B$2*E{i}',
            f'=CORREL({R(C["a_mean"])},{R(C["b_mean"])})', f'=COUNTIF({R(C["pair_status"])},"PASS")/C{i}',
            pooled(R(C['a_sd']), R(C['a_n'])), pooled(R(C['b_sd']), R(C['b_n'])),
            f'=SPEC_USED!$E${spec_r}', f'=SPEC_USED!$F${spec_r}', f'=SPEC_USED!$G${spec_r}',
            f'=SPEC_USED!$H${spec_r}', f'=SPEC_USED!$I${spec_r}',
            f'=IF(C{i}<P{i},"INSUFFICIENT",IF(AND(ABS(D{i})<=L{i},ABS(F{i})<=M{i},ABS(G{i})<=M{i},I{i}>=O{i}),"PASS","REVIEW"))'])
    n = len(specs)+1
    fmt(result, f'D2:O{n}', '0.0000'); fmt(result, f'I2:I{n}', '0.0%'); fmt(result, f'O2:O{n}', '0.0%')
    result.conditional_formatting.add(f'Q2:Q{n}', FormulaRule(formula=['Q2="REVIEW"'], fill=PatternFill('solid', fgColor=RED)))
    result.conditional_formatting.add(f'Q2:Q{n}', FormulaRule(formula=['Q2="PASS"'], fill=PatternFill('solid', fgColor=GREEN)))

    subset_heads = ['characteristic', 'sample_id', 'A mean', 'B mean', 'Bias B-A', 'A repeat σ', 'B repeat σ',
                    'Pair status', 'A repeat count', 'B repeat count']
    header(subset, subset_heads, [27, 14, 19, 19, 18, 19, 19, 18, 16, 16])
    subset_ranges = {}; cur = 2
    for s in specs:
        metric = s['characteristic']; begin = cur
        for sample in selected_list:
            r = sample_rows[(metric, sample)]
            subset.append([metric, sample] + [f'=SAMPLE_CALC!{C[k]}{r}' for k in
                          ('a_mean', 'b_mean', 'bias', 'a_sd', 'b_sd', 'pair_status', 'a_n', 'b_n')])
            cur += 1
        subset_ranges[metric] = (begin, cur-1)
    fmt(subset, f'C2:G{cur-1}', '0.0000')

    sub_heads = ['characteristic', 'n subset', 'Bias B-A', 'SD difference', 'LoA lower', 'LoA upper',
                 'Pearson r', 'Pair pass rate', 'A repeat σ', 'B repeat σ', 'Decision']
    header(subresult, sub_heads, [27, 14, 20, 18, 18, 18, 18, 18, 18, 18, 17])
    for i, s in enumerate(specs, 2):
        metric = s['characteristic']; lo, hi = subset_ranges[metric]
        R = lambda col: f'SUBSET_CALC!${col}${lo}:${col}${hi}'
        subresult.append([metric,
            f'=COUNT({R("E")})', f'=AVERAGE({R("E")})', f'={STDEV}({R("E")})',
            f'=C{i}-SETTINGS!$B$2*D{i}', f'=C{i}+SETTINGS!$B$2*D{i}',
            f'=CORREL({R("C")},{R("D")})', f'=COUNTIF({R("H")},"PASS")/B{i}',
            pooled(R('F'), R('I')), pooled(R('G'), R('J')),
            f'=IF(B{i}<RESULT!P{i},"INSUFFICIENT",IF(AND(ABS(C{i})<=RESULT!L{i},ABS(E{i})<=RESULT!M{i},ABS(F{i})<=RESULT!M{i},H{i}>=RESULT!O{i}),"PASS","REVIEW"))'])
    fmt(subresult, f'C2:J{n}', '0.0000'); fmt(subresult, f'H2:H{n}', '0.0%')

    header(guide, ['Sheet', 'How to inspect the formula / assumptions'], [30, 112])
    for row in [
        ['RAW_DATA', 'Input CSV snapshot (blank rows skipped, surrounding spaces trimmed). The source CSV is NEVER overwritten.'],
        ['SPEC_USED', 'Values loaded from SPEC.xlsx. Change source SPEC and rerun Python to refresh all reports.'],
        ['SAMPLE_CALC', 'Click A/B repeat cells to see the original RAW_DATA reference; A/B mean = AVERAGE(available repeated readings; number of repetitions can vary).'],
        ['SAMPLE_CALC', 'Bias (B-A) = Device B mean minus Device A mean; pair pass = |Bias| <= tolerance from SPEC_USED.'],
        ['RESULT', 'All study metrics are native Excel formulas, including CORREL, Bias, SD, LoA, repeatability σ and PASS/REVIEW.'],
        ['RESULT', 'Repeatability σ = pooled within-sample SD weighted by degrees of freedom: SQRT(Σ(n_i−1)·s_i² / Σ(n_i−1)).'],
        ['SELECTED_SAMPLES', 'The selected IDs define the exploratory sample subset; this is NOT an independent new experiment.'],
        ['SUBSET_RESULT', 'Compares the selected subset with the full study using native Excel formulas.'],
        ['Interpretation', 'High Pearson r only shows linear association, not instrument agreement.'],
        ['Caution', 'LoA is for sample MEANS of repeated measurements; not for single readings.'],
        ['Caution', 'No true calibrated reference is supplied; Device A is only a comparator.'],
        ['Workflow', 'After editing Raw CSV/SPEC XLSX, rerun python main.py. Excel calculates all formula results on open.'],
    ]:
        guide.append(row)
    for row in guide.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='top')

    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    target = out/'Agreement_Analysis.xlsx'
    wb.save(str(target))
    return target
