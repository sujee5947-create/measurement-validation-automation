"""Measurement Correlation & Validation, synthetic portfolio project.

The program reads RAW and SPEC separately, assesses device agreement using
sample-level means, contrasts 32 vs 120 independent samples, and automatically
creates local HTML and PPT reports. No hidden validation scenario labels are used.
"""
from pathlib import Path
import argparse
from collections import defaultdict
import csv
import html
import json
import math
import re
from statistics import mean

import numpy as np
from scipy.stats import pearsonr, linregress, t as student_t
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import openpyxl
from src.raw_io import read_raw_rows, natural_key
from src.export_excel import export_excel
import plotly.graph_objects as go
from plotly.offline import plot as plotly_div

ROOT=Path(__file__).resolve().parent
DATA=ROOT/'data'
OUT=ROOT/'outputs'
CHARTS=OUT/'charts'
REPORTS=OUT/'reports'
for p in (OUT,CHARTS,REPORTS): p.mkdir(parents=True,exist_ok=True)

RAW_PATH=DATA/'Measurement_Raw_Data.csv'
SPEC_PATH=DATA/'Measurement_Agreement_SPEC.xlsx'
SELECT_PATH=DATA/'Sample_Selection_32.csv'


def load_spec(path=SPEC_PATH):
    if not Path(path).exists():
        raise FileNotFoundError(f'SPEC workbook not found: {path}')
    wb=openpyxl.load_workbook(str(path),data_only=True,read_only=True)
    if 'AGREEMENT_SPEC' not in wb.sheetnames:
        raise ValueError('SPEC sheet AGREEMENT_SPEC was not found')
    values=[list(row) for row in wb['AGREEMENT_SPEC'].iter_rows(min_row=1,max_row=1000,max_col=11,values_only=True)]
    wb.close()
    values=[[v.strip() if isinstance(v,str) else v for v in row] for row in values if row[0] not in (None,'')]
    if not values: raise ValueError('SPEC sheet is empty')
    fields=values[0]
    required={'characteristic','display_name','unit','enabled','max_abs_mean_bias','max_abs_loa_bound','max_abs_pair_diff','min_pair_pass_rate','min_samples','min_repeats_per_device'}
    if not required.issubset(fields):
        raise ValueError(f'SPEC missing columns: {sorted(required-set(fields))}')
    result=[]
    for row in values[1:]:
        record=dict(zip(fields,row))
        if not record.get('characteristic'):continue
        if str(record['enabled']).strip().upper()=='Y':
            result.append(record)
    if not result:raise ValueError('No enabled characteristics in SPEC workbook')
    seen=set()
    for x in result:
        name=str(x['characteristic']).strip()
        if not name or name in seen or not re.fullmatch(r'[A-Za-z0-9_\-]+',name):
            raise ValueError(f'Duplicate/invalid characteristic in SPEC: {name!r}')
        seen.add(name)
        for col in ('max_abs_mean_bias','max_abs_loa_bound','max_abs_pair_diff'):
            try: v=float(x[col])
            except (ValueError,TypeError): raise ValueError(f'{name}: invalid {col}')
            if not math.isfinite(v) or v<0: raise ValueError(f'{name}: {col} must be finite and >= 0')
        rate=float(x['min_pair_pass_rate'])
        if not math.isfinite(rate) or not 0<=rate<=1:raise ValueError(f'{name}: min_pair_pass_rate must be between 0 and 1')
        for col,minimum in (('min_samples',3),('min_repeats_per_device',2)):
            v=float(x[col])
            if not math.isfinite(v) or v!=int(v) or v<minimum:
                raise ValueError(f'{name}: {col} must be an integer >= {minimum}')
        if not str(x['unit']).strip():raise ValueError(f'{name}: unit missing')
    return result


def load_raw(path=None):
    path = Path(path) if path is not None else RAW_PATH
    if not path.exists():
        raise FileNotFoundError(f'Raw data not found: {path}')
    fieldnames,rows,_enc=read_raw_rows(path)
    required={'characteristic','sample_id','instrument_id','repeat_id','measurement_value','lot_id','sample_position','unit','timestamp'}
    if not fieldnames or not required.issubset(fieldnames):
        raise ValueError(f'RAW missing columns: {sorted(required-set(fieldnames or []))}')
    if len(fieldnames)!=len(set(fieldnames)):
        raise ValueError('RAW has duplicate column headers')
    if not rows:raise ValueError('Raw data is empty')
    groups=defaultdict(lambda: {'DEVICE_A':{},'DEVICE_B':{}})
    sample_info={}
    duplicate=[]
    for r in rows:
        row_number=r['_line']
        if not r['sample_id'] or not r['characteristic']:raise ValueError(f'RAW row {row_number}: missing sample_id or characteristic')
        key=(r['characteristic'],r['sample_id'])
        device=r['instrument_id']
        if device not in ('DEVICE_A','DEVICE_B'):raise ValueError(f'Unexpected device: {device}')
        try:repeat=int(r['repeat_id'])
        except (ValueError,TypeError):raise ValueError(f'RAW row {row_number}: invalid repeat_id')
        if repeat<1:raise ValueError(f'RAW row {row_number}: repeat_id must be >= 1')
        if repeat in groups[key][device]:duplicate.append((key,device,repeat))
        try:value=float(r['measurement_value'])
        except (ValueError,TypeError):raise ValueError(f'RAW row {row_number}: invalid measurement_value')
        if not np.isfinite(value):raise ValueError(f'RAW row {row_number}: non-finite measurement')
        groups[key][device][repeat]=value
        metadata={'lot_id':r['lot_id'],'sample_position':r['sample_position'],'unit':r['unit']}
        previous=sample_info.get(key)
        if previous and any(previous[f]!=metadata[f] for f in metadata):
            raise ValueError(f'Inconsistent sample metadata or measurement unit for: {key}')
        ts=r['timestamp'] if previous is None else min(previous['timestamp'],r['timestamp'])
        sample_info[key]={**metadata,'timestamp':ts}
    if duplicate:raise ValueError(f'Duplicate readings found: {duplicate[:3]}')
    return groups,sample_info,len(rows)


def build_records(groups,sample_info,specs):
    details=defaultdict(list)
    for s in specs:
        key_name=s['characteristic']
        repeats_needed=int(s['min_repeats_per_device'])
        for (metric,sample),dev in sorted(groups.items(),key=lambda kv:(kv[0][0],natural_key(kv[0][1]))):
            if metric!=key_name:continue
            a=dev['DEVICE_A'];b=dev['DEVICE_B']
            if set(a)!=set(b) or len(a)<repeats_needed or set(a)!=set(range(1,len(a)+1)):
                raise ValueError(f'Incomplete/mismatched repeats: {metric}/{sample}')
            av=np.array([a[i] for i in sorted(a)],dtype=float)
            bv=np.array([b[i] for i in sorted(b)],dtype=float)
            m_a=float(av.mean());m_b=float(bv.mean());difference=m_b-m_a
            info=sample_info[(metric,sample)]
            if str(info['unit']).strip()!=str(s['unit']).strip():
                raise ValueError(f'Raw unit differs from SPEC for {metric}: {info["unit"]!r} vs {s["unit"]!r}')
            details[metric].append({
                'sample_id':sample,'lot_id':info['lot_id'],'timestamp':info['timestamp'],
                'sample_position':info['sample_position'],'device_a_mean':m_a,
                'device_b_mean':m_b,'pair_mean':(m_a+m_b)/2,'difference_B_minus_A':difference,
                'device_a_sd':float(av.std(ddof=1)),
                'device_b_sd':float(bv.std(ddof=1)),
                'device_a_n':len(av),'device_b_n':len(bv),
                'device_a_repeats':av.tolist(),'device_b_repeats':bv.tolist()
            })
    missing=[x['characteristic'] for x in specs if not details[x['characteristic']]]
    if missing:raise ValueError(f'No RAW data for enabled SPEC characteristics: {missing}')
    unknown=set(k[0] for k in groups)-set(x['characteristic'] for x in specs)
    if unknown:raise ValueError(f'RAW characteristics not enabled in SPEC: {sorted(unknown)}')
    return details


def calculate(records, spec):
    n=len(records)
    min_n=int(spec['min_samples'])
    if n<min_n:
        return {'n':n,'status':'INSUFFICIENT','error':f'n={n}, min={min_n}'}
    a=np.array([x['device_a_mean'] for x in records],dtype=float)
    b=np.array([x['device_b_mean'] for x in records],dtype=float)
    d=b-a
    sd=float(np.std(d,ddof=1))
    bias=float(np.mean(d))
    loa_lo=bias-1.96*sd
    loa_hi=bias+1.96*sd
    tolerance=float(spec['max_abs_pair_diff'])
    pair_pass_rate=float(np.mean(np.abs(d)<=tolerance))
    if np.ptp(a)==0 or np.ptp(b)==0:
        raise ValueError(f'{spec["characteristic"]}: zero variation in instrument means; Pearson/regression undefined')
    if np.ptp((a+b)/2)==0:
        raise ValueError(f'{spec["characteristic"]}: zero variation in pair means; proportional-bias regression undefined')
    corr=float(pearsonr(a,b).statistic)
    regression=linregress(a,b)
    prop=linregress((a+b)/2,d)
    ci_half=float(student_t.ppf(.975,n-1)*sd/math.sqrt(n))
    # Pooled within-sample SD, weighted by degrees of freedom (n_i-1); identical to the
    # simple RMS when every sample has the same repeat count.
    def pooled(sd_key,n_key):
        return float(math.sqrt(sum((r[n_key]-1)*r[sd_key]**2 for r in records)/sum(r[n_key]-1 for r in records)))
    a_sd=pooled('device_a_sd','device_a_n')
    b_sd=pooled('device_b_sd','device_b_n')
    checks={
      'mean_bias':abs(bias)<=float(spec['max_abs_mean_bias']),
      'loa_bounds':max(abs(loa_lo),abs(loa_hi))<=float(spec['max_abs_loa_bound']),
      'pair_rate':pair_pass_rate>=float(spec['min_pair_pass_rate'])
    }
    return dict(
        n=n,bias=bias,diff_sd=sd,loa_lower=loa_lo,loa_upper=loa_hi,
        bias_ci_lower=bias-ci_half,bias_ci_upper=bias+ci_half,bias_ci_width=2*ci_half,
        pearson_r=corr,regression_slope=float(regression.slope),
        regression_intercept=float(regression.intercept),
        proportional_bias_slope=float(prop.slope),proportional_bias_p=float(prop.pvalue),
        pair_pass_rate=pair_pass_rate,a_repeatability_sigma=a_sd,b_repeatability_sigma=b_sd,
        status='PASS' if all(checks.values()) else 'REVIEW',checks=checks,
        bias_limit=float(spec['max_abs_mean_bias']),loa_bound_limit=float(spec['max_abs_loa_bound']),
        pair_diff_limit=tolerance,min_pair_pass_rate=float(spec['min_pair_pass_rate'])
    )


def write_csv(records,file_path,headers):
    with file_path.open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=headers,extrasaction='ignore');w.writeheader();w.writerows(records)


def build_charts(details,results,specs):
    plt.rcParams.update({'font.family':'DejaVu Sans','axes.spines.top':False,'axes.spines.right':False,'font.size':10})
    created={}
    for s in specs:
        metric=s['characteristic']; items=details[metric]; r=results[metric]['full']
        av=np.array([x['device_a_mean'] for x in items]);bv=np.array([x['device_b_mean'] for x in items]);
        means=(av+bv)/2;diff=bv-av
        fig,axes=plt.subplots(1,2,figsize=(11.8,4.5))
        axes[0].scatter(av,bv,c='#2761A6',alpha=.70,s=20)
        xy=[min(np.min(av),np.min(bv)),max(np.max(av),np.max(bv))]
        axes[0].plot(xy,xy,linestyle='--',c='#64748B',label='Identity: A=B')
        axes[0].set(xlabel=f"Device A ({s['unit']})",ylabel=f"Device B ({s['unit']})",title=f"Device correlation  |  r = {r['pearson_r']:.3f}")
        axes[0].legend(frameon=False)
        axes[1].scatter(means,diff,c='#2761A6',alpha=.70,s=20)
        axes[1].axhline(r['bias'],c='#2563EB',label=f"Bias {r['bias']:+.3f}")
        axes[1].axhline(r['loa_lower'],c='#B91C1C',ls='--',label='Lower LoA')
        axes[1].axhline(r['loa_upper'],c='#B91C1C',ls='--',label='Upper LoA')
        axes[1].axhline(0,c='#9CA3AF',ls=':')
        axes[1].set(xlabel=f"Pairwise mean ({s['unit']})",ylabel=f"Device B − Device A ({s['unit']})",title='Bland–Altman agreement')
        axes[1].legend(frameon=False,fontsize=8)
        fig.suptitle(s['display_name'],fontweight='bold',fontsize=15)
        fig.tight_layout()
        p=CHARTS/f'{metric}_agreement.png';fig.savefig(p,dpi=150,bbox_inches='tight');plt.close(fig)
        created[metric]=str(p)
    # Bias progression by sample for time-trend diagnostic
    trend_metrics=([k for k in ('coating_weight_mg_cm2','uncoated_gap_mm') if k in details] + [k for k in details if k not in ('coating_weight_mg_cm2','uncoated_gap_mm')])[:2]
    fig,axes=plt.subplots(len(trend_metrics),1,figsize=(12.5,2.55*len(trend_metrics)+.5),sharex=True)
    axes=np.atleast_1d(axes)
    for ax,metric in zip(axes,trend_metrics):
        xs=np.arange(1,len(details[metric])+1)
        d=np.array([x['difference_B_minus_A'] for x in details[metric]])
        ax.plot(xs,d,lw=.8,color='#94A3B8')
        window=min(15,len(d))
        rolling=np.convolve(d,np.ones(window)/window,'valid')
        ax.plot(xs[window-1:],rolling,lw=2,color='#1D4ED8',label=f'{window}-sample trailing mean')
        ax.axhline(0,color='#64748B',ls=':')
        ax.set(ylabel='B − A',title=metric)
        ax.legend(frameon=False,fontsize=8)
    axes[-1].set_xlabel('Ordered sample index')
    fig.tight_layout()
    p=CHARTS/'device_difference_trend.png';fig.savefig(p,dpi=150,bbox_inches='tight');plt.close(fig)
    return created,p


def write_html(specs,details,results):
    # Portable offline Plotly: JS injected once into the HTML file.
    content=[]
    plotly_bundle=None
    for index,s in enumerate(specs):
        k=s['characteristic']; recs=details[k]; r=results[k]['full']
        av=np.array([x['device_a_mean'] for x in recs]);bv=np.array([x['device_b_mean'] for x in recs]);
        diff=bv-av; average=(av+bv)/2
        scatter=go.Figure()
        scatter.add_trace(go.Scatter(x=av,y=bv,mode='markers',name='Sample means',marker=dict(color='#2563EB',size=7,opacity=.6),text=[x['sample_id'] for x in recs],hovertemplate='Sample %{text}<br>Device A %{x:.4f}<br>Device B %{y:.4f}<extra></extra>'))
        ends=[float(min(av.min(),bv.min())),float(max(av.max(),bv.max()))]
        scatter.add_trace(go.Scatter(x=ends,y=ends,name='A=B',mode='lines',line=dict(color='#64748B',dash='dash')))
        scatter.update_layout(title=f'Correlation (Pearson r = {r["pearson_r"]:.3f})',xaxis_title=f'Device A ({s["unit"]})',yaxis_title=f'Device B ({s["unit"]})')
        ba=go.Figure()
        ba.add_trace(go.Scatter(x=average,y=diff,mode='markers',marker=dict(color='#2563EB',size=7,opacity=.6),text=[x['sample_id'] for x in recs],hovertemplate='Sample %{text}<br>Mean %{x:.4f}<br>B−A %{y:.4f}<extra></extra>',name='B−A'))
        for y,name,color,dash in [(0,'Zero','#64748B','dot'),(r['bias'],'Bias','#2563EB','solid'),(r['loa_lower'],'Lower LoA','#B91C1C','dash'),(r['loa_upper'],'Upper LoA','#B91C1C','dash')]:
            ba.add_hline(y=y,line=dict(color=color,dash=dash),annotation_text=name,annotation_position='top left')
        ba.update_layout(title='Bland–Altman plot',xaxis_title=f'Mean of A & B ({s["unit"]})',yaxis_title=f'B−A ({s["unit"]})')
        trend=go.Figure()
        trend.add_trace(go.Scatter(x=[x['sample_id'] for x in recs],y=diff,mode='lines+markers',name='Device difference',marker=dict(size=4),line=dict(color='#3B82F6')))
        trend.add_hline(y=r['bias'],line=dict(color='#9333EA',dash='dash'),annotation_text='Overall bias')
        trend.update_layout(title='Measurement difference by ordered sample',xaxis_title='Sample ID',yaxis_title=f'B−A ({s["unit"]})')
        # Use divs that resize with the parent; Plotly include_plotlyjs only once.
        snippets=[]
        for f in (scatter,ba,trend):
            f.update_layout(height=335,margin=dict(l=50,r=15,t=55,b=50),paper_bgcolor='white',plot_bgcolor='white',font=dict(family='Arial, sans-serif',size=11),legend=dict(orientation='h',y=-.32),autosize=True)
            snippets.append(plotly_div(f,include_plotlyjs=True if plotly_bundle is None else False,output_type='div',config={'responsive':True,'displaylogo':False}))
            plotly_bundle='included'
        subset=results[k]['subset']
        status_color='ok' if r['status']=='PASS' else 'warn'
        explanations=[]
        if not r['checks']['mean_bias']:explanations.append('Mean bias exceeds the selected acceptance limit.')
        if not r['checks']['loa_bounds']:explanations.append('At least one limit of agreement exceeds the absolute agreement limit.')
        if not r['checks']['pair_rate']:explanations.append('Too many sample-level differences exceed tolerance.')
        if r['proportional_bias_p']<.05:explanations.append('Difference vs mean shows a statistically detectable slope; inspect possible proportional bias.')
        if not explanations:explanations.append('All configured agreement thresholds passed. This does not establish traceable calibration accuracy.')
        content.append(f'''<section class="metric" id="{html.escape(k)}">
<h2>{html.escape(s['display_name'])} <small>{html.escape(s['unit'])}</small> <span class="pill {status_color}">{r['status']}</span></h2>
<div class="stats">
 <div class="stat"><span>Mean bias (B−A)</span><strong>{r['bias']:+.4f}</strong></div>
 <div class="stat"><span>Pearson r</span><strong>{r['pearson_r']:.3f}</strong></div>
 <div class="stat"><span>95% LoA</span><strong>{r['loa_lower']:+.3f} ~ {r['loa_upper']:+.3f}</strong></div>
 <div class="stat"><span>Pair tolerance PASS</span><strong>{r['pair_pass_rate']:.1%}</strong></div>
 <div class="stat"><span>Device A repeatability σ</span><strong>{r['a_repeatability_sigma']:.4f}</strong></div>
 <div class="stat"><span>Device B repeatability σ</span><strong>{r['b_repeatability_sigma']:.4f}</strong></div>
</div>
<div class="note">{' '.join(html.escape(x) for x in explanations)}</div>
<div class="plots"><div class="plot">{snippets[0]}</div><div class="plot">{snippets[1]}</div><div class="plot full">{snippets[2]}</div></div>
<h3>32 vs 120 independent samples</h3>
<div class="table-scroll"><table><thead><tr><th>Study</th><th>n</th><th>Bias</th><th>Bias 95% CI</th><th>Lower LoA</th><th>Upper LoA</th><th>Decision</th></tr></thead><tbody>
<tr><td>32-sample exploratory</td><td>{subset['n']}</td><td>{subset['bias']:+.4f}</td><td>[{subset['bias_ci_lower']:+.4f}, {subset['bias_ci_upper']:+.4f}]</td><td>{subset['loa_lower']:+.4f}</td><td>{subset['loa_upper']:+.4f}</td><td>{subset['status']}</td></tr>
<tr><td>Full 120-sample</td><td>{r['n']}</td><td>{r['bias']:+.4f}</td><td>[{r['bias_ci_lower']:+.4f}, {r['bias_ci_upper']:+.4f}]</td><td>{r['loa_lower']:+.4f}</td><td>{r['loa_upper']:+.4f}</td><td>{r['status']}</td></tr>
</tbody></table></div>
</section>''')
    nav=''.join(f'<a href="#{html.escape(s["characteristic"])}">{html.escape(s["display_name"])}</a>' for s in specs)
    cards=''.join(f'<div class="hero-kpi"><span>{html.escape(s["display_name"])}</span><strong>{results[s["characteristic"]]["full"]["status"]}</strong></div>' for s in specs)
    full='''<!doctype html><html lang="en"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width, initial-scale=1"/><title>Measurement Correlation & Validation</title>
<style>
:root{--navy:#17365d;--blue:#23549a;--text:#182331;--muted:#5c6b7a;--line:#dce4ed;--bg:#f4f7fb}*{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;background:var(--bg);font-family:Arial,"Noto Sans",sans-serif;color:var(--text);overflow-x:hidden}header{background:linear-gradient(120deg,#132b4e,#285e98);padding:32px max(16px,calc((100vw - 1200px)/2));color:white}header h1{margin:0;font-size:clamp(24px,4vw,36px)}header p{line-height:1.55;max-width:860px;color:#edf4ff}main{max-width:1230px;padding:14px 16px 70px;margin:auto}.hero{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,170px),1fr));gap:10px;margin-top:20px}.hero-kpi{border:1px solid #e0e7ef;background:#fff;border-radius:12px;padding:13px;min-width:0}.hero-kpi span,.stat span{display:block;font-size:12px;color:var(--muted);line-height:1.4}.hero-kpi strong{display:block;font-size:22px;margin-top:8px}nav{display:flex;gap:9px;flex-wrap:wrap;margin:20px 0}nav a{background:white;border:1px solid var(--line);color:var(--blue);padding:9px 12px;border-radius:100px;text-decoration:none;font-size:12px}.metric{background:white;border:1px solid var(--line);border-radius:16px;padding:clamp(13px,2vw,22px);margin:20px 0;min-width:0}.metric h2{font-size:clamp(19px,2.4vw,26px);margin:0 0 15px}.metric small{color:var(--muted);font-size:14px;font-weight:normal}.pill{float:right;font-size:12px;padding:6px 11px;border-radius:90px}.pill.ok{background:#dff5e7;color:#166534}.pill.warn{background:#ffedd5;color:#9a3412}.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,155px),1fr));gap:8px}.stat{background:#f7f9fc;padding:12px;border-radius:10px;min-width:0}.stat strong{display:block;margin-top:8px;overflow-wrap:anywhere;font-size:clamp(15px,1.8vw,20px)}.note{background:#eef4fc;border-left:3px solid #2563eb;padding:12px 14px;margin:16px 0;font-size:13px;line-height:1.6}.plots{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}.plot{width:100%;max-width:100%;min-width:0;overflow:hidden}.plot.full{grid-column:1/-1}.plot .plotly-graph-div{width:100%!important;max-width:100%}h3{font-size:16px;margin:24px 0 9px}.table-scroll{max-width:100%;overflow-x:auto}table{border-collapse:collapse;min-width:700px;width:100%;font-size:12px}th,td{padding:9px;border-bottom:1px solid var(--line);text-align:left}th{background:#f0f5fb}footer{font-size:12px;color:var(--muted);margin:30px 0;line-height:1.7}@media(max-width:700px){.plots{grid-template-columns:1fr}.plot.full{grid-column:auto}.metric{border-radius:12px}.pill{float:none;display:inline-block;margin-left:7px}.stats{grid-template-columns:repeat(2,minmax(0,1fr))}.hero-kpi strong{font-size:19px}header{padding:23px 16px}}@media(max-width:390px){.stats{grid-template-columns:1fr 1fr}.stat{padding:9px}.stat strong{font-size:14px}.metric h2{font-size:19px}.metric{padding:12px}.plots{gap:2px}nav a{font-size:11px}}
</style></head><body>
<header><h1>Measurement Correlation & Validation</h1><p>Battery Electrode Quality · Instrument A vs Instrument B · 120 independent samples · 3 repeated readings per instrument per sample · 5 characteristics</p></header>
<main><p><b>Interpretation:</b> Device A is a comparator, not a certified true-value reference. Pearson correlation is not an agreement decision. Each Bland–Altman point represents the difference between the two <b>three-repeat sample means</b>; it does not represent single-reading agreement.</p>
<div class="hero">'''+cards+'''</div><nav>'''+nav+'''</nav>'''+''.join(content)+'''<footer><b>Limitations:</b> Starter dataset and acceptance thresholds are synthetic; confirm provenance of replaced inputs. Agreement limits are approximate 95% LoA assuming stable, suitably distributed differences. This report does not establish traceability, true accuracy, or causal equipment events. Bias confidence intervals are classical t intervals. The exploratory subset is drawn from the full cohort, not an independent experiment.</footer></main></body></html>'''
    n_full = next(iter(results.values()))['full']['n']
    n_sub = next(iter(results.values()))['subset']['n']
    repeat_counts=sorted({len(rec['device_a_repeats']) for items in details.values() for rec in items})
    repeat_label=str(repeat_counts[0]) if len(repeat_counts)==1 else f'{min(repeat_counts)}–{max(repeat_counts)}'
    full = full.replace('120 independent samples', f'{n_full} independent samples')
    full = full.replace('5 characteristics</p>', f'{len(specs)} characteristics</p>')
    full = full.replace('3 repeated readings per instrument per sample', f'{repeat_label} repeated readings per instrument per sample')
    full = full.replace('three-repeat sample means', 'repeated-reading sample means')
    full = full.replace('32-sample exploratory', f'{n_sub}-sample exploratory')
    full = full.replace('Full 120-sample', f'Full {n_full}-sample')
    full = full.replace('32 vs 120 independent samples', f'{n_sub} vs {n_full} independent samples').replace('32-sample subset', f'{n_sub}-sample subset')
    out=REPORTS/'Measurement_Validation_Report.html'
    out.write_text(full,encoding='utf-8')
    return out


def ppt_text(slide,text,x,y,w,h,size=16,color=(21,40,65),bold=False):
    shape=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
    tf=shape.text_frame;tf.word_wrap=True
    p=tf.paragraphs[0];p.text=str(text);p.alignment=PP_ALIGN.LEFT
    p.font.size=Pt(size);p.font.bold=bold;p.font.color.rgb=RGBColor(*color)
    return shape


def add_header(prs,title,kicker='MEASUREMENT CORRELATION & VALIDATION'):
    sl=prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg=sl.background.fill;slide_bg.solid();slide_bg.fore_color.rgb=RGBColor(247,249,252)
    ppt_text(sl,kicker,0.55,0.34,11.5,.3,10,(45,92,147),True)
    ppt_text(sl,title,.55,.73,12,.65,26,(24,49,79),True)
    return sl


def add_bottom(sl,n):
    ppt_text(sl,f'Synthetic validation dataset · Training use only   |   {n}',.55,7.15,11,.2,9,(110,125,142))


def write_ppt(specs,results,chart_paths,trend_path):
    prs=Presentation();prs.slide_width=Inches(13.333);prs.slide_height=Inches(7.5)
    n_full=next(iter(results.values()))['full']['n']
    n_subset=next(iter(results.values()))['subset']['n']
    repeat_counts=sorted({len(rec['device_a_repeats']) for items in REPORT_DETAILS.values() for rec in items})
    repeat_label=str(repeat_counts[0]) if len(repeat_counts)==1 else f'{min(repeat_counts)}–{max(repeat_counts)}'
    sl=add_header(prs,'Instrument A / B agreement validation','BATTERY ELECTRODE QUALITY  /  PROJECT 02')
    ppt_text(sl,f'{n_full} independent electrode samples  ·  {repeat_label} repeats per device  ·  {len(specs)} quality characteristics',.7,1.8,11,1.0,21,(38,72,111))
    ppt_text(sl,'What is being tested?',.7,3.1,10,.45,22,(23,49,82),True)
    ppt_text(sl,'Association (Pearson)  ≠  Interchangeability (Bias + Bland–Altman)\nRepeatability sigma separately quantified per instrument\nEditable acceptance criteria loaded from SPEC.xlsx',.72,3.66,11,2.0,18,(48,67,85))
    add_bottom(sl,1)

    sl=add_header(prs,'Study design & evaluation rules')
    ppt_text(sl,f'{n_full} independent samples  |  {len(specs)} characteristics  |  Device A vs B  |  {repeat_label} reads/device/sample',.7,1.65,12,.5,20,(42,83,125),True)
    ppt_text(sl,'Each device’s repeated measurements are averaged within each sample.\nThe difference B−A is calculated once per independent sample.\nAgreement: mean Bias, 95% LoA, pass rate within tolerance.\nRepeatability: pooled within-sample standard deviation by device.\nPearson r and regression are descriptive, not agreement acceptance.',.7,2.35,11.5,2.8,19,(50,67,84))
    ppt_text(sl,f'{n_subset}-sample sensitivity comparison uses a fixed random subset; results must not be treated as independent of the full {n_full}-sample study.',.7,5.65,11.5,.75,14,(111,67,37))
    add_bottom(sl,2)

    sl=add_header(prs,'Executive summary — agreement decisions')
    y=1.6
    for s in specs:
        r=results[s['characteristic']]['full']
        status=r['status'];color=(23,111,82) if status=='PASS' else (165,80,24)
        ppt_text(sl,s['display_name'],.72,y,3.6,.4,18,(34,56,77),True)
        ppt_text(sl,f'Pearson r  {r["pearson_r"]:.3f}   |   Bias  {r["bias"]:+.3f}   |   LoA  [{r["loa_lower"]:+.3f}, {r["loa_upper"]:+.3f}]',4.1,y,7.4,.4,14,(56,72,88))
        ppt_text(sl,status,11.5,y,1.25,.4,15,color,True)
        y+=.95
    ppt_text(sl,'Important: A high correlation does not eliminate systematic bias.',.72,6.57,11.5,.35,14,(122,67,39),True)
    add_bottom(sl,3)

    for num,s in enumerate(specs,4):
        r=results[s['characteristic']]['full'];r32=results[s['characteristic']]['subset']
        sl=add_header(prs,f'{s["display_name"]} — agreement')
        sl.shapes.add_picture(chart_paths[s['characteristic']],Inches(.63),Inches(1.55),width=Inches(12.05))
        ppt_text(sl,f"n={r['n']}  •  mean B−A {r['bias']:+.4f} {s['unit']}  •  pair pass {r['pair_pass_rate']:.1%}  •  {r['status']}",.78,6.48,11.5,.35,16,(37,72,111),True)
        ppt_text(sl,f"{n_subset}-sample bias {r32['bias']:+.4f} (95% CI {r32['bias_ci_lower']:+.4f} to {r32['bias_ci_upper']:+.4f})",.78,6.88,11.5,.3,11,(94,105,116))
        add_bottom(sl,num)

    sl=add_header(prs,'Time order diagnostic — measurement difference')
    sl.shapes.add_picture(str(trend_path),Inches(.75),Inches(1.50),width=Inches(11.8))
    ppt_text(sl,'Observed time patterns are descriptive. No machine-event log was provided; trends are not causal evidence.',.76,6.42,11.7,.4,13,(120,74,34))
    add_bottom(sl,4+len(specs))

    sl=add_header(prs,'Conclusions & validation next steps')
    ppt_text(sl,'Confirmed from the generated readings',.7,1.60,10,.45,20,(38,74,114),True)
    ppt_text(sl,f'• Agreement acceptance is characteristic-specific, using editable SPEC limits.\n• {n_full} independent sample differences support the full study.\n• Repeated measurements provide a separate device repeatability estimate.\n• {n_subset}-sample results can be compared with the full {n_full}-sample results.',.72,2.12,11.2,2.4,18,(43,62,81))
    ppt_text(sl,'Do not infer from these data alone',.7,5.07,10,.45,19,(151,83,33),True)
    ppt_text(sl,'• Traceable measurement accuracy, real production equipment faults, or causal calibration events.',.72,5.56,11,.85,17,(66,73,83))
    add_bottom(sl,5+len(specs))
    p=REPORTS/'Measurement_Validation_Report.pptx';prs.save(p);return p


def main(argv=None):
    global OUT, CHARTS, REPORTS
    parser = argparse.ArgumentParser(description='Rebuild device agreement XLSX, PPTX and offline HTML from current SPEC and Raw Data')
    parser.add_argument('--spec', type=Path, default=SPEC_PATH, help='Agreement SPEC .xlsx input')
    parser.add_argument('--raw', type=Path, default=RAW_PATH, help='Raw measurements .csv input')
    parser.add_argument('--output', type=Path, default=OUT, help='Output directory for generated reports')
    parser.add_argument('--subset-size', type=int, default=None, help='Exploratory sample count; automatic for smaller datasets (maximum 32)')
    args=parser.parse_args(argv)
    OUT=args.output.resolve(); CHARTS=OUT/'charts'; REPORTS=OUT/'reports'
    if args.subset_size is not None and args.subset_size<3:raise ValueError('--subset-size must be >= 3')
    for directory in (OUT, CHARTS, REPORTS): directory.mkdir(parents=True,exist_ok=True)

    specs=load_spec(args.spec)
    groups,info,nread=load_raw(args.raw)
    details=build_records(groups,info,specs)
    sample_sets=[set(x['sample_id'] for x in details[s['characteristic']]) for s in specs]
    if any(ids!=sample_sets[0] for ids in sample_sets[1:]):
        raise ValueError('Different sample sets across characteristics; comparison would be misleading')
    available=sample_sets[0]
    n_samples=len(available)
    min_required=max(int(s['min_samples']) for s in specs)
    if args.subset_size is None:
        args.subset_size=min(32,n_samples-1)
    if args.subset_size<min_required or args.subset_size>=n_samples:
        raise ValueError(f'Exploratory subset must be between SPEC min_samples ({min_required}) and n_samples-1 ({n_samples-1}); override --subset-size or adjust your study design')
    # Prefer the saved, deterministic 32-ID selection for the starter dataset.
    chosen=set()
    if args.subset_size==32 and SELECT_PATH.exists():
        with SELECT_PATH.open('r',encoding='utf-8-sig',newline='') as f:
            chosen={r['sample_id'] for r in csv.DictReader(f)}
    if len(chosen)!=args.subset_size or not chosen.issubset(available):
        chosen=set(np.random.default_rng(32).choice(sorted(available,key=natural_key),size=args.subset_size,replace=False).tolist())
    global REPORT_DETAILS
    REPORT_DETAILS=details
    result={};summary=[];details_out=[]
    for spec in specs:
        key=spec['characteristic'];all_records=details[key]
        sub=[x for x in all_records if x['sample_id'] in chosen]
        if len(sub)!=args.subset_size:raise ValueError('Exploratory subset did not match all samples')
        r_full=calculate(all_records,spec);r_sub=calculate(sub,spec)
        for label,res in [('full',r_full),('subset',r_sub)]:
            if res.get('status')=='INSUFFICIENT':
                raise ValueError(f"{key} {label}: not enough independent samples ({res['error']}); update min_samples or sample count")
        result[key]={'full':r_full,'subset':r_sub}
        for cohort,res in [('full',r_full),('subset',r_sub)]:
            summary.append({'characteristic':key,'study':cohort,**{k:v for k,v in res.items() if k!='checks' and k!='error'},**{f'check_{k}':v for k,v in res.get('checks',{}).items()}})
        for rec in all_records:
            details_out.append({'characteristic':key,**{k:v for k,v in rec.items() if not isinstance(v,list)}})
    write_csv(summary,OUT/'Agreement_Summary.csv',list(summary[0]))
    write_csv(details_out,OUT/'Sample_Mean_Differences.csv',list(details_out[0]))
    (OUT/'Agreement_Summary.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    charts,trend=build_charts(details,result,specs)
    html_report=write_html(specs,details,result)
    ppt_report=write_ppt(specs,result,charts,trend)
    excel_report=export_excel(ROOT,specs,result,details,output_dir=OUT,subset_size=args.subset_size,raw_path=args.raw,selected_ids=chosen)
    print(f'COMPLETED: {nread:,} raw readings | {n_samples} independent samples | {len(specs)} characteristics')
    for spec in specs:
        k=spec['characteristic'];r=result[k]['full']
        print(f"{k:28} r={r['pearson_r']:.3f} bias={r['bias']:+.4f} LoA=[{r['loa_lower']:+.4f},{r['loa_upper']:+.4f}] pair_PASS={r['pair_pass_rate']:.1%} -> {r['status']}")
    print('Generated:',html_report,ppt_report,excel_report)

if __name__=='__main__':
    main()
