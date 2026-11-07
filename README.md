# Measurement Correlation & Validation — **Automation Project**

> 두 측정 장비(Device A vs B)의 일치도(Agreement)를 raw 데이터만으로 자동 평가하고, **Excel(수식 추적) · HTML · PPT 보고서**를 생성하는 Python 도구
> Automated two-instrument agreement study (Bland–Altman, bias, repeatability) for battery electrode quality characteristics.

### 핵심 기능 (Highlights)
- **Bland–Altman 일치 분석**: 평균 편향(Bias), 95% 일치 한계(LoA), bias 95% CI, 비례 편향(Proportional Bias) 검정
- **장비별 반복성(Repeatability σ)**: 자유도 가중 pooled SD
- **SPEC 기반 자동 판정**: Bias / LoA / 시료별 허용차 통과율 → PASS · REVIEW (Pearson r은 참고용, 판정 제외)
- **Excel 수식 추적성(Traceability)**: 모든 결과 셀이 `RAW_DATA` 원본 셀까지 네이티브 Excel 수식으로 연결
- **입력 견고성**: 가변 반복 수, 열/행 순서 무관, UTF-8/CP949, 빈 행·공백 자동 처리, 결손·중복·단위 불일치 시 명시적 오류

### 검증 (Verification)
| 데이터셋 | 규모 | 검증 내용 | 결과 |
|---|---|---|---|
| 기본 데이터 (`data/`) | 120시료 × 3반복 × 5특성 (3,600행) | 독립 재계산(pandas/scipy) vs 출력, LibreOffice 수식 재계산 vs Python, 3,600개 셀 참조 전수 확인 | 오차 ≤1e-12, 판정 10/10 일치 |
| 신규 데이터 (`examples/new_raw_80samples/`) | 80시료 × 4반복 × 5특성 (3,200행), 시나리오 주입 | 정답지(`Expected_Results_AnswerKey.csv`) 대비 판정·실패 항목·수치 | 5/5 특성 일치 (음의 bias, 이상치, 점진적 drift 검출) |
| 비정형 입력 | 반복 3~5 혼재, 행 섞임, CP949, 빈 행, 공백 | Python · Excel · HTML · PPT 전 출력 | 독립 계산과 일치 |

> ⚠️ 포함된 데이터와 SPEC 허용 기준은 모두 **학습용 합성(synthetic) 값**입니다. 실제 적용 시 제품 규격 공차 또는 공정 산포 기준으로 SPEC을 재설정해야 합니다.

### Quick start
```bash
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt   # Windows
.venv\Scripts\python.exe main.py
.venv\Scripts\python.exe main.py --raw examples/new_raw_80samples/Measurement_Raw_Data_NEW.csv --output outputs_new
```
결과물은 `outputs/`에 생성됩니다 (저장소에는 포함하지 않음, 실행 시 재생성).

---

This is the runnable VS Code / Cursor Python source project, **not just a one-time report**. Every run reads the current raw measurements and SPEC workbook and rebuilds the report outputs.

## 1. 처음 실행하기 (Windows)

1. Unzip the project and open the project folder (the one containing `main.py`) in VS Code or Cursor.
2. Double-click **`INSTALL_FIRST.bat`** once to create `.venv` and install Python dependencies.
3. Double-click **`RUN_REPORT.bat`** any time you want to recalculate results, or run `python main.py` in a properly configured Python environment.
4. Open `outputs/reports/Measurement_Validation_Report.html` in a browser, or the `.pptx` in PowerPoint.

Alternatively in the VS Code terminal:

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

## 2. 자동화 입력 파일

- `data/Measurement_Raw_Data.csv`: long-format paired instrument readings. Each sample × characteristic × repeat has one DEVICE_A and one DEVICE_B measurement. Repeats must be 1..N with no gaps, at least min_repeats_per_device; N can vary across samples (both devices paired).
- `data/Measurement_Agreement_SPEC.xlsx`: edit `AGREEMENT_SPEC` to change accepted bias, 95% LoA bounds and sample-level tolerance.
- `data/Sample_Selection_32.csv`: fixed exploratory sample selection. If input data changes and IDs are incompatible, the app automatically draws a reproducible subset.

**Editing inputs and rerunning updates every output.** Neither Raw Data nor SPEC is overwritten by `main.py`.

Example: If `max_abs_mean_bias` for `coating_thickness_um` is changed in the Excel SPEC, the next run recalculates PASS/REVIEW from the new threshold. The report must use the actual measurements; it never prints fabricated PASS/REVIEW counts.

## 3. 자동 생성 출력물

```text
outputs/
  Agreement_Analysis.xlsx
  Agreement_Summary.csv
  Agreement_Summary.json
  Sample_Mean_Differences.csv
  charts/   (PNG scatter + Bland–Altman and difference trend)
  reports/
    Measurement_Validation_Report.pptx
    Measurement_Validation_Report.html
```

The HTML report bundles Plotly JavaScript and charts in the same file for offline viewing. It uses responsive layouts and scrollable tables for mobile. Layout should be checked on actual devices for production usage.

## 4. 커스텀 입력 경로

```powershell
python main.py --raw "C:\my_data\Measurement_Raw_Data.csv" --spec "C:\my_data\SPEC.xlsx"
python main.py --subset-size 40 --output "C:\reports"
python main.py --help
```

Input requirements: complete A/B paired readings per sample/metric/repeat, consistent unit, no duplicated reading key; at least the `min_samples` specified in SPEC for both full and subset analysis. This version expects two instrument IDs (`DEVICE_A`, `DEVICE_B`) and requires all configured characteristics to use the same set of sample IDs.

## 5. 통계 분석 내용

- Two-device agreement for **mean of available repeats** (not the single-reading LoA)
- Paired average **Bias = B−A**, bias 95% CI
- **95% Bland–Altman limits of agreement** (±1.96 sample SD of paired mean differences)
- Pearson r and simple regression; high correlation does not imply agreement
- Separate pooled within-sample repeatability σ for each device
- Configurable mean bias, absolute LoA bound, sample tolerance and passing fraction
- Preliminary 32-sample subset vs full cohort
- Difference trend by ordered sample, not an invented equipment event

All readings and tolerance limits are synthetic and for training. DEVICE_A is a comparator, **not** a traceable true-value reference. The 32-subset and full study overlap; they are not independent experiments. If proportional bias is significant, constant LoA may not adequately describe agreement.

## 6. 테스트

`python main.py` 실행 후 `python -m unittest discover -s tests -v` (Excel 수식 테스트는 생성된 `outputs/Agreement_Analysis.xlsx`를 사용). Tests validate paired design, high correlation vs poor agreement, live SPEC sensitivity and correct independent sample count. The report outputs are rebuilt at each successful run.

## 7. 엑셀 수식 추적(새 기능)

자동 생성되는 `outputs/Agreement_Analysis.xlsx`에 다음 시트가 들어갑니다.

| 시트 | 역할 |
| --- | --- |
| `RAW_DATA` | 입력 CSV의 **동일한 값·컬럼·행 순서**를 복사한 스냅샷. 원본 CSV 파일 자체는 변경하지 않음. |
| `SPEC_USED` | 실행 당시 입력 SPEC의 허용 기준 복사본. |
| `SELECTED_SAMPLES` | 32개(또는 사용자 지정) 예비 분석 대상 시료 ID. |
| `SAMPLE_CALC` | 각 시료의 장비별 3회 읽기, `AVERAGE`, `STDEV.S`, B−A, 절댓값 및 페어별 허용 판정. **각 반복 측정값은 헤더에서 탐색한 `RAW_DATA` 셀을 수식으로 참조하며, 실제 반복 수에 따라 수식 범위가 조정됨.** |
| `RESULT` | 전체 시료 기준 Bias, 차이 SD, 95% LoA, `CORREL`, 반복성 σ, 허용 기준 및 PASS/REVIEW를 **Excel 수식으로 계산**. |
| `SUBSET_CALC` | 선택 시료 계산 내역(전체 계산 시트 참조). |
| `SUBSET_RESULT` | 32개 시료 예비 분석 지표를 수식으로 계산해 전체 시료와 비교. |
| `SETTINGS`, `FORMULA_GUIDE` | LoA 상수 및 각 수식의 출처·주의사항. |

Excel에서 `RESULT`의 **D2(Bias)**, **H2(Pearson r)**, **Q2(PASS/REVIEW)** 등을 클릭하면 수식 입력줄에서 실제 수식을 볼 수 있습니다. `SAMPLE_CALC`의 E2 등의 셀을 클릭하면 `RAW_DATA!` 원본 셀 참조를 볼 수 있습니다. `F2`를 눌러 수식 편집도 가능합니다.

**워크플로:** 원본 `data/Measurement_Raw_Data.csv` 및 `data/Measurement_Agreement_SPEC.xlsx`만 수정 → `python main.py` 재실행 → 새 수식/보고서 생성. Excel 결과 시트에서 직접 값을 수정했다면 원본 데이터 변경을 대체하지 않습니다.

## 8. 실제 입력 적용 시 주의사항 / Supported input contract

**v3 변경 사항 (검증 후 수정)**
- `artifact_tool`(PyPI 미배포 → `INSTALL_FIRST.bat` 실패) 제거, `openpyxl`로 대체.
- 반복성 σ: 자유도 가중 pooled SD `sqrt(Σ(n_i−1)s_i² / Σ(n_i−1))`로 수정 (반복수가 같으면 기존 값과 동일, 반복수 혼재 시에만 차이).
- RAW 읽기: UTF-8 실패 시 CP949 자동 재시도, 완전히 빈 행(`,,,,`) 무시, 앞뒤 공백 제거, 장비 ID 대문자화.
- 시료 정렬: 자연 정렬(S1, S2, … S10), 시료 timestamp = 해당 시료의 가장 이른 측정 시각(행 순서 무관).


- RAW CSV는 UTF-8/UTF-8 BOM/CP949(한글 Excel 저장). 필수 컬럼: `timestamp, sample_id, lot_id, sample_position, characteristic, unit, instrument_id, repeat_id, measurement_value`. 추가 컬럼 및 컬럼 순서 변경 허용.
- 장비 ID는 `DEVICE_A`, `DEVICE_B`. 다른 장비 이름은 입력 CSV에서 매핑해 주세요.
- 각 characteristic의 sample ID 집합은 동일해야 합니다. 누락된 시료, 서로 다른 단위, 중복 measurement key, 빠진 반복, NaN/Infinity 등은 **오류로 중단**합니다. 조용히 임의 보정하지 않습니다.
- 반복수는 표본별 3/4/5회 등 가변으로 허용하며 두 장비의 반복 번호가 일치해야 합니다. SPEC의 `min_repeats_per_device` 이상 및 2 이상이어야 합니다. 원본의 반복 횟수가 다르면 `SAMPLE_CALC`가 최댓값까지 슬롯을 확장하며 비어 있는 셀은 평균/표준편차에서 제외됩니다.
- SPEC의 모든 활성 특성은 RAW에 반드시 있어야 하고, SPEC 밖의 RAW 특성도 명시적 오류로 처리합니다. 서로 다른 특성의 샘플 구성이 다르면 각각 따로 실행하거나 입력을 정리하세요.
- 최소 sample 수는 SPEC `min_samples` 이상이어야 하고, **subset도 동일한 최소 n 요구**를 적용합니다. 기본 subset은 `min(32,n-1)`입니다. 일부 데이터의 샘플 수가 매우 작으면 subset과 full 비교가 성립하지 않으므로 의도적으로 중단합니다.
- Excel 수식은 RAW의 현재 열명으로 참조합니다. CSV 행 순서 변경도 지원합니다.
- 결과가 기존보다 바뀌었다면 RAW/SPEC을 먼저 검증하십시오. 사용자 데이터가 합성 데이터인지 실제 데이터인지 프로그램이 자동으로 판별하지 않습니다.
- 특성 이름은 영문, 숫자, `_`, `-`만 허용됩니다. 단위는 CSV와 SPEC이 정확하게 일치해야 합니다.
- Pearson 및 회귀 분석에 필요한 분산이 0이면 오해를 일으키는 값을 생성하는 대신 명시적 오류가 발생합니다.
- 출력 파일은 마지막으로 성공한 실행의 결과일 수 있습니다. 콘솔에서 `COMPLETED`를 확인한 뒤 결과를 사용하세요.
- Excel 수식은 프로그램이 새 보고서를 만들 때 계산합니다. 보고서 안에서 직접 수식을 바꾼 내용은 다음 실행 때 덮어쓰여집니다.
- Excel 보고서는 공개 패키지 `openpyxl`로 생성합니다(v3부터 `artifact_tool` 의존성 제거). 수식 결과는 Excel에서 파일을 열 때 자동 계산됩니다.

### 검증된 사례

1. 번들 RAW 3,600행 × 5특성 × 120시료 × 장비별 3반복: 동일 통계값 유지, 전체 보고서 재생성 성공.
2. **열 순서를 뒤집고 행을 섞은 외부 RAW 1,400행** × 5특성 × 40시료 × 3~4반복: Excel 수식, HTML, PPT 전체 생성 성공.
3. 필수 데이터 결손은 오류로 중단 (SPEC에 설정된 최소 반복수 미달).

### 실행 명령

```powershell
.venv\Scripts\python.exe main.py --raw "C:\data\new_readings.csv" --spec "C:\data\SPEC.xlsx" --output "C:\reports"
.venv\Scripts\python.exe -m unittest discover -s tests -v
```
